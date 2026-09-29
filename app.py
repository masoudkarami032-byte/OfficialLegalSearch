from flask import Flask, render_template, request, jsonify, send_file
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import requests
import re
import threading
import uuid
import os
import time
import math

from urllib.parse import urljoin, urlencode
from html import unescape


app = Flask(__name__)

BASE = 'https://ara.jri.ac.ir'
LIST = BASE + '/Judge/Index'

JOBS = {}

SOURCE_ID = 'national_judgments'
SOURCE_NAME = 'سامانه ملی آرای قضایی پژوهشگاه قوه قضاییه'

PAGE_SIZE = 25

QAVANIN_BASE = 'https://qavanin.ir'
QAVANIN_LIST = QAVANIN_BASE + '/'
QAVANIN_SOURCE_ID = 'national_laws'
QAVANIN_SOURCE_NAME = 'سامانه ملی قوانین و مقررات جمهوری اسلامی ایران'
QAVANIN_PAGE_SIZE = 10


# =========================================================
# NORMALIZE TEXT
# =========================================================

def norm(s):
    if not s:
        return ''

    s = str(s)

    s = (
        s.replace('ي', 'ی')
         .replace('ك', 'ک')
         .replace('\u200c', ' ')
    )

    return re.sub(r'\s+', ' ', s).strip()


def fa_to_en(s):
    if s is None:
        return ''

    return str(s).translate(
        str.maketrans(
            '۰۱۲۳۴۵۶۷۸۹',
            '0123456789'
        )
    )


# =========================================================
# QUERY
# =========================================================

def parse_query(q):

    parts = re.split(
        r'\s+(AND|NOT)\s+',
        q.strip(),
        flags=re.I
    )

    include = []
    exclude = []

    mode = 'AND'

    for part in parts:

        if part.upper() in ('AND', 'NOT'):
            mode = part.upper()
            continue

        part = part.strip().strip('"“”')

        if not part:
            continue

        if mode == 'NOT':
            exclude.append(part)
        else:
            include.append(part)

        mode = 'AND'

    return include, exclude


def matches(text, query):

    text = fa_to_en(norm(text))

    include, exclude = parse_query(query)

    if not include:
        return False

    for item in include:
        if fa_to_en(norm(item)) not in text:
            return False

    for item in exclude:
        if fa_to_en(norm(item)) in text:
            return False

    return True


# =========================================================
# HTTP SESSION
# =========================================================

def make_session():

    session = requests.Session()

    session.headers.update({
        'User-Agent':
            'Mozilla/5.0 (Linux; Android 13) '
            'AppleWebKit/537.36 '
            '(KHTML, like Gecko) '
            'Chrome/120.0 Mobile Safari/537.36',

        'Accept':
            'text/html,application/xhtml+xml,'
            'application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',

        'Accept-Language':
            'fa-IR,fa;q=0.9,en-US;q=0.7,en;q=0.6',

        'Referer': LIST
    })

    return session


# =========================================================
# RESULT LINKS
# =========================================================

def get_vote_links(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    links = []
    seen = set()

    for a in soup.find_all(
        'a',
        href=True
    ):

        href = a.get(
            'href',
            ''
        )

        if '/Judge/Text/' not in href:
            continue

        url = urljoin(
            BASE,
            href.split('?')[0]
        )

        if url not in seen:

            seen.add(url)
            links.append(url)

    return links


# =========================================================
# TOTAL RESULTS / TOTAL PAGES
# =========================================================

def extract_total_results(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    text = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    patterns = [
        r'تعداد\s*یافته\s*ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*یافته‌ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*یافته\s*[:：]?\s*([0-9۰-۹,٬]+)'
    ]

    for pattern in patterns:

        m = re.search(
            pattern,
            text
        )

        if not m:
            continue

        value = fa_to_en(
            m.group(1)
        )

        value = (
            value
            .replace(',', '')
            .replace('٬', '')
        )

        try:
            return int(value)
        except ValueError:
            pass

    return None


def extract_page_numbers(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    select = soup.find(
        'select',
        attrs={
            'name': 'PageNumbers'
        }
    )

    if not select:
        return []

    pages = []

    for option in select.find_all(
        'option'
    ):

        value = fa_to_en(
            option.get(
                'value',
                ''
            )
        ).strip()

        if value.isdigit():
            pages.append(
                int(value)
            )

    return pages


def extract_total_pages(html):

    pages = extract_page_numbers(
        html
    )

    if pages:
        return max(pages)

    total_results = extract_total_results(
        html
    )

    if total_results is not None:

        return max(
            1,
            math.ceil(
                total_results / PAGE_SIZE
            )
        )

    links = get_vote_links(
        html
    )

    if links:
        return 1

    return 0


# =========================================================
# FORM HELPERS
# =========================================================

def hidden_fields(form):

    data = []

    if not form:
        return data

    for inp in form.find_all(
        'input'
    ):

        name = inp.get(
            'name'
        )

        if not name:
            continue

        input_type = (
            inp.get(
                'type',
                ''
            )
            .lower()
        )

        if input_type != 'hidden':
            continue

        value = inp.get(
            'value',
            ''
        )

        data.append(
            (name, value)
        )

    return data


def find_search_form(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    form = soup.find(
        'form',
        id='frmSearch'
    )

    if form:
        return form

    for f in soup.find_all(
        'form'
    ):

        if f.find(
            attrs={'name': 'Title'}
        ):
            return f

    return None


def find_pagination_form(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    for form in soup.find_all(
        'form'
    ):

        if form.find(
            attrs={'name': 'PageNumbers'}
        ):
            return form

    return None


# =========================================================
# INITIAL OFFICIAL SEARCH
# =========================================================

def official_search(
    session,
    query,
    search_title,
    search_abstract,
    search_text
):

    initial = session.get(
        LIST,
        timeout=30
    )

    initial.raise_for_status()

    initial.encoding = (
        initial.apparent_encoding
        or 'utf-8'
    )

    form = find_search_form(
        initial.text
    )

    payload = hidden_fields(
        form
    )

    controlled = {
        'Title',
        'IsTitleSearch',
        'IsAbstractSearch',
        'IsTextSearch',
        '_IsOfficial',
        '_IsCivil',
        '_IsPenal',
        'SeachTextType',
        'SortColumn',
        'SortDesc'
    }

    payload = [
        (k, v)
        for k, v in payload
        if k not in controlled
    ]

    payload.append(
        ('Title', query)
    )

    if search_title:
        payload.append(
            ('IsTitleSearch', 'true')
        )

    payload.append(
        ('IsTitleSearch', 'false')
    )

    if search_abstract:
        payload.append(
            ('IsAbstractSearch', 'true')
        )

    payload.append(
        ('IsAbstractSearch', 'false')
    )

    if search_text:
        payload.append(
            ('IsTextSearch', 'true')
        )

    payload.append(
        ('IsTextSearch', 'false')
    )

    payload.append(
        ('_IsOfficial', 'false')
    )

    payload.append(
        ('_IsCivil', 'false')
    )

    payload.append(
        ('_IsPenal', 'false')
    )

    payload.append(
        ('SeachTextType', '3')
    )

    payload.append(
        ('SortColumn', 'Overdate')
    )

    payload.append(
        ('SortDesc', 'True')
    )

    response = session.post(
        LIST,
        data=payload,
        timeout=30,
        allow_redirects=True
    )

    response.raise_for_status()

    response.encoding = (
        response.apparent_encoding
        or 'utf-8'
    )

    return response.text


# =========================================================
# PAGINATION
# =========================================================

def get_page(
    session,
    previous_html,
    page
):

    form = find_pagination_form(
        previous_html
    )

    payload = hidden_fields(
        form
    )

    controlled = {
        'PageNumbers',
        'PageNumber',
        'PageSize'
    }

    payload = [
        (k, v)
        for k, v in payload
        if k not in controlled
    ]

    payload.append(
        ('PageNumbers', str(page))
    )

    payload.append(
        ('PageNumber', str(page))
    )

    payload.append(
        ('PageSize', str(PAGE_SIZE))
    )

    response = session.post(
        LIST,
        data=payload,
        timeout=30,
        allow_redirects=True
    )

    response.raise_for_status()

    response.encoding = (
        response.apparent_encoding
        or 'utf-8'
    )

    return response.text


# =========================================================
# EXTRACT SECTIONS FROM EACH JUDGMENT
# =========================================================

def extract_between(
    text,
    starts,
    ends
):

    text = norm(text)

    start_position = None

    for marker in starts:

        p = text.find(marker)

        if p != -1:

            p += len(marker)

            if (
                start_position is None
                or p < start_position
            ):
                start_position = p

    if start_position is None:
        return ''

    end_position = len(text)

    for marker in ends:

        p = text.find(
            marker,
            start_position
        )

        if (
            p != -1
            and p < end_position
        ):
            end_position = p

    return norm(
        text[
            start_position:
            end_position
        ]
    )


def fetch_vote(
    url,
    session
):

    response = session.get(
        url,
        timeout=30
    )

    response.raise_for_status()

    response.encoding = (
        response.apparent_encoding
        or 'utf-8'
    )

    soup = BeautifulSoup(
        response.text,
        'html.parser'
    )

    full_text = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    title = extract_between(
        full_text,
        ['عنوان: ', 'عنوان : ', 'عنوان'],
        [
            'پیام:',
            'پیام :',
            'پیام',
            'مستندات'
        ]
    )

    if not title:

        h = (
            soup.find('h1')
            or soup.find('h2')
        )

        if h:

            title = norm(
                h.get_text(
                    ' ',
                    strip=True
                )
            )

    if not title:

        title = 'رأی قضایی'

    abstract = extract_between(
        full_text,
        [
            'پیام: ',
            'پیام : ',
            'پیام'
        ],
        [
            'مستندات',
            'شماره دادنامه قطعی',
            'گروه رأی',
            'آراء منتخب پرونده'
        ]
    )

    body_markers = [
        'رأی شعبه بدوی',
        'رای شعبه بدوی',
        'رأی دادگاه بدوی',
        'رای دادگاه بدوی',
        'رأی دادگاه تجدیدنظر',
        'رای دادگاه تجدیدنظر',
        'رأی شعبه تجدیدنظر',
        'رای شعبه تجدیدنظر',
        'رأی شعبه دیوان عالی کشور',
        'رای شعبه دیوان عالی کشور'
    ]

    starts = []

    for marker in body_markers:

        p = full_text.find(
            marker
        )

        if p != -1:
            starts.append(p)

    if starts:

        body = full_text[
            min(starts):
        ]

    else:

        body = full_text

    footer_position = len(body)

    for marker in [
        'نقد رأی',
        'نقد رای',
        'تعدادموافق',
        'تماس با ما'
    ]:

        p = body.find(
            marker
        )

        if (
            p != -1
            and p < footer_position
        ):
            footer_position = p

    body = norm(
        body[:footer_position]
    )

    return {
        'url': url,
        'title': title,
        'abstract': abstract,
        'body': body,
        'text': full_text,
        'source': SOURCE_NAME
    }


# =========================================================
# MATCH
# =========================================================

def vote_matches(
    vote,
    query,
    search_title,
    search_abstract,
    search_text
):

    selected = []

    if search_title:

        selected.append(
            vote.get(
                'title',
                ''
            )
        )

    if search_abstract:

        selected.append(
            vote.get(
                'abstract',
                ''
            )
        )

    if search_text:

        selected.append(
            vote.get(
                'body',
                ''
            )
        )

    return matches(
        ' '.join(selected),
        query
    )


def matched_in(
    vote,
    query,
    search_title,
    search_abstract,
    search_text
):

    locations = []

    if (
        search_title
        and matches(
            vote.get('title', ''),
            query
        )
    ):
        locations.append(
            'عنوان'
        )

    if (
        search_abstract
        and matches(
            vote.get('abstract', ''),
            query
        )
    ):
        locations.append(
            'پیام'
        )

    if (
        search_text
        and matches(
            vote.get('body', ''),
            query
        )
    ):
        locations.append(
            'متن رأی'
        )

    return locations


# =========================================================
# NEZAMAT.IR - LAWS AND REGULATIONS ENGINE
# The public route name /api/qavanin/search is kept unchanged so
# the existing index.html continues to work without modification.
# =========================================================

def make_qavanin_session():
    session = requests.Session()
    session.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Linux; Android 14) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/153.0.0.0 Mobile Safari/537.36'
        ),
        'Accept': (
            'text/html,application/xhtml+xml,application/xml;q=0.9,'
            'image/avif,image/webp,*/*;q=0.8'
        ),
        'Accept-Language': 'fa-IR,fa;q=0.9,en-US;q=0.7,en;q=0.6',
        'Referer': QAVANIN_LIST,
        'Connection': 'keep-alive'
    })
    return session




def qavanin_get_with_retry(session, url, *, params=None, timeout=35, attempts=7):
    """GET from Nezamat with bounded retry/backoff.

    A catalogue page is never silently skipped: after all retries fail, the
    exception is raised so the job is reported incomplete rather than done.
    """
    last_error = None
    retry_statuses = {408, 425, 429, 500, 502, 503, 504}
    for attempt in range(1, attempts + 1):
        try:
            response = session.get(
                url, params=params, timeout=timeout, allow_redirects=True
            )
            if response.status_code in retry_statuses:
                raise requests.exceptions.HTTPError(
                    f'temporary HTTP {response.status_code}', response=response
                )
            response.raise_for_status()
            return response
        except (
            requests.exceptions.ConnectionError,
            requests.exceptions.Timeout,
            requests.exceptions.ChunkedEncodingError,
            requests.exceptions.HTTPError,
        ) as exc:
            last_error = exc
            # Do not retry ordinary permanent 4xx responses.
            status = getattr(getattr(exc, 'response', None), 'status_code', None)
            if status is not None and 400 <= status < 500 and status not in retry_statuses:
                raise
            if attempt >= attempts:
                break
            # Re-open the TCP connection on the next attempt.
            try:
                session.close()
            except Exception:
                pass
            delay = min(8.0, 0.6 * (2 ** (attempt - 1)))
            time.sleep(delay)
    raise last_error or RuntimeError('ارتباط با نظامات برقرار نشد.')


def qavanin_is_challenge(html):
    """Detect qavanin.ir/Arvan transfer pages.

    IMPORTANT: a challenge is an access failure, never a valid zero-result page.
    """
    raw = norm(BeautifulSoup(html or '', 'html.parser').get_text(' ', strip=True))
    lower = (html or '').lower()
    return (
        'transferring to the website' in lower
        or 'در حال انتقال به سایت مورد نظر هستید' in raw
        or ('arvancloud' in lower and 'transfer' in lower)
    )


def _qavanin_bool_params(search_title=True, search_text=False):
    # ASP.NET MVC checkbox helpers submit checked=true followed by hidden=false.
    params = []
    if search_title:
        params.append(('IsTitleSearch', 'true'))
    params.append(('IsTitleSearch', 'false'))
    if search_text:
        params.append(('IsTextSearch', 'true'))
    params.append(('IsTextSearch', 'false'))
    return params


def _qavanin_search_params(query, page=1, size=100, search_title=True, search_text=False):
    params = [('CAPTION', query), ('Zone', '')]
    params.extend(_qavanin_bool_params(search_title, search_text))
    params.extend([
        ('_isLaw', 'false'), ('_isRegulation', 'false'),
        ('_IsVote', 'false'), ('_isOpenion', 'false'),
        # 3 = «بخشی از واژه» exactly as the official form.
        ('SeachTextType', '3'),
        ('fromApproveDate', ''), ('APPROVEDATE', ''),
        ('IsTitleSubject', 'False'), ('IsMain', ''),
        ('COMMANDNO', ''), ('fromCommandDate', ''), ('COMMANDDATE', ''),
        ('NEWSPAPERNO', ''), ('fromNewspaperDate', ''), ('NEWSPAPERDATE', ''),
        ('SortColumn', 'APPROVEDATE'), ('SortDesc', 'True'), ('Report_ID', ''),
        ('PageNumber', str(max(1, int(page)))), ('page', str(max(1, int(page)))),
        ('size', str(min(max(1, int(size)), 1000))),
        ('txtZone', ''), ('txtSubjects', ''), ('txtExecutors', ''),
        ('txtApprovers', ''), ('txtLawStatus', ''), ('txtLawTypes', '')
    ])
    return params


def _qavanin_get(session, url, *, params=None, timeout=45, attempts=3):
    last = None
    for attempt in range(1, attempts + 1):
        try:
            r = session.get(url, params=params, timeout=timeout, allow_redirects=True)
            r.raise_for_status()
            r.encoding = 'utf-8'
            if qavanin_is_challenge(r.text):
                raise RuntimeError(
                    'qavanin.ir صفحه امنیتی/انتقال برگرداند؛ این پاسخ نتیجه جست‌وجو نیست.'
                )
            return r
        except RuntimeError:
            raise
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError, requests.exceptions.HTTPError) as exc:
            last = exc
            if attempt < attempts:
                time.sleep(min(4, 0.8 * (2 ** (attempt - 1))))
    raise last or RuntimeError('ارتباط با qavanin.ir برقرار نشد.')


def _qavanin_int(text):
    raw = fa_to_en(norm(text)).replace(',', '').replace('٬', '')
    m = re.search(r'\d+', raw)
    return int(m.group()) if m else 0


def _qavanin_parse_search(html):
    soup = BeautifulSoup(html, 'html.parser')
    page_text = norm(soup.get_text(' ', strip=True))

    total = 0
    m = re.search(r'تعداد\s*یافته\s*ها\s*[:：]?\s*([0-9۰-۹٠-٩,٬]+)', page_text)
    if m:
        total = _qavanin_int(m.group(1))

    results = []
    seen = set()
    for a in soup.select('a[href*="/Law/TreeText/"]'):
        href = (a.get('href') or '').strip()
        title = norm(a.get_text(' ', strip=True))
        if not href or not title:
            continue
        url = urljoin(QAVANIN_BASE, href)
        if url in seen:
            continue
        seen.add(url)
        tr = a.find_parent('tr')
        cells = [norm(td.get_text(' ', strip=True)) for td in tr.find_all('td')] if tr else []
        approval_date = cells[2] if len(cells) >= 3 else ''
        approver = cells[3] if len(cells) >= 4 else ''
        results.append({
            'url': url, 'title': title, 'approval_date': approval_date,
            'approver': approver, 'source': QAVANIN_SOURCE_NAME
        })
    return results, total


def qavanin_search_page(session, query, page=1, search_title=True, search_text=False, size=100):
    r = _qavanin_get(
        session, QAVANIN_LIST,
        params=_qavanin_search_params(query, page, size, search_title, search_text)
    )
    return r.text, r.url


def fetch_law(url, session):
    r = _qavanin_get(session, url, timeout=45, attempts=3)
    soup = BeautifulSoup(r.text, 'html.parser')
    for tag in soup(['script', 'style', 'noscript', 'svg']):
        tag.decompose()
    title = ''
    for selector in ('h1', 'h2', '.law-title', '.title'):
        node = soup.select_one(selector)
        if node:
            candidate = norm(node.get_text(' ', strip=True))
            if candidate:
                title = candidate
                break
    body = norm((soup.body or soup).get_text(' ', strip=True))
    return {
        'url': r.url, 'title': title or 'قانون یا مقرره', 'abstract': '',
        'body': body, 'text': body, 'source': QAVANIN_SOURCE_NAME,
        'approval_date': '', 'publication_date': '', 'document_number': '', 'category': ''
    }


def law_matches(law, query, search_title, search_text):
    # qavanin.ir itself is authoritative for candidate selection. This helper is
    # retained for compatibility when inspecting a fetched document.
    selected = []
    if search_title:
        selected.append(law.get('title', ''))
    if search_text:
        selected.append(law.get('body', ''))
    return matches(' '.join(selected), query)


def law_matched_in(law, query, search_title, search_text):
    locations = []
    if search_title:
        locations.append('عنوان')
    if search_text:
        locations.append('متن قانون/مقرره')
    return locations


def qavanin_worker(jid, query, max_pages, search_title, search_text):
    job = JOBS[jid]
    session = make_qavanin_session()
    page_size = 100
    try:
        job['message'] = 'در حال ارسال جست‌وجو به سامانه ملی قوانین و مقررات...'
        first_html, official_url = qavanin_search_page(
            session, query, page=1, search_title=search_title,
            search_text=search_text, size=page_size
        )
        first_items, official_total = _qavanin_parse_search(first_html)

        # Guard against the dangerous failure mode seen before: an unfiltered
        # catalogue page must not be reported as a legitimate search result.
        if official_total > 100000 and query:
            raise RuntimeError(
                'qavanin.ir فیلتر جست‌وجو را اعمال نکرد و فهرست عمومی را برگرداند؛ '
                'برای جلوگیری از نتیجه نادرست، عملیات متوقف شد.'
            )

        total_pages = max(1, (official_total + page_size - 1) // page_size) if official_total else 1
        total_pages = min(total_pages, max_pages)
        job['official_results'] = official_total
        job['site_total_pages'] = total_pages
        job['total_pages'] = total_pages
        job['official_url'] = official_url

        if official_total and not first_items:
            raise RuntimeError('qavanin.ir تعداد نتیجه اعلام کرد اما ردیف‌های نتیجه قابل استخراج نبودند.')

        seen = set()
        for page in range(1, total_pages + 1):
            if job['cancel']:
                break
            job['current_page'] = page
            if page == 1:
                items = first_items
            else:
                html, _ = qavanin_search_page(
                    session, query, page=page, search_title=search_title,
                    search_text=search_text, size=page_size
                )
                items, _ = _qavanin_parse_search(html)

            for item in items:
                if job['cancel']:
                    break
                if item['url'] in seen:
                    continue
                seen.add(item['url'])
                job['checked'] += 1

                law = dict(item)
                law['abstract'] = ''
                law['body'] = ''
                law['text'] = ''
                law['publication_date'] = ''
                law['document_number'] = ''
                law['category'] = item.get('approver', '')
                law['matched_in'] = ([] if not search_title else ['عنوان']) + ([] if not search_text else ['متن قانون/مقرره'])

                # Fetch the document body for text searches / later reporting.
                # A detail-page failure does not erase a valid official hit.
                if search_text:
                    try:
                        detail = fetch_law(item['url'], session)
                        if detail.get('title') and detail['title'] != 'قانون یا مقرره':
                            law['title'] = detail['title']
                        law['body'] = detail.get('body', '')
                        law['text'] = detail.get('text', '')
                    except Exception:
                        job['failed_items'] += 1

                job['results'].append(law)
                job['found'] = len(job['results'])
                if official_total:
                    job['progress'] = min(99, round(job['checked'] / official_total * 100, 1))
                job['message'] = (
                    f'در حال دریافت نتایج رسمی: {job["checked"]} از {official_total or "?"}'
                )

            job['completed_pages'] = page
            if not items:
                break

        job['found'] = len(job['results'])
        if job['cancel']:
            job['status'] = 'cancelled'
            job['message'] = 'جست‌وجو به درخواست کاربر متوقف شد.'
        else:
            job['status'] = 'done'
            job['progress'] = 100
            job['message'] = (
                f'جست‌وجوی رسمی qavanin.ir تکمیل شد؛ '
                f'{official_total} یافته رسمی و {job["found"]} نتیجه دریافت شد.'
            )
    except Exception as e:
        job['status'] = 'error'
        job['message'] = 'خطا در جست‌وجوی رسمی قوانین و مقررات: ' + str(e)


# =========================================================
# WORKER
# =========================================================

def worker(
    jid,
    query,
    max_pages,
    search_title,
    search_abstract,
    search_text
):

    job = JOBS[jid]

    session = make_session()

    try:

        job['message'] = (
            'در حال جست‌وجو در سامانه رسمی...'
        )

        first_html = official_search(
            session,
            query,
            search_title,
            search_abstract,
            search_text
        )

        first_links = get_vote_links(
            first_html
        )

        official_total = extract_total_results(
            first_html
        )

        real_total_pages = extract_total_pages(
            first_html
        )

        if (
            real_total_pages >= 1000
            and query.strip()
        ):

            job['status'] = 'error'

            job['total_pages'] = 0

            job['message'] = (
                'سامانه رسمی نتیجه جست‌وجو را '
                'اعمال نکرد و فهرست کل آرا را برگرداند. '
                'برای جلوگیری از بررسی اشتباه کل سامانه، '
                'جست‌وجو متوقف شد.'
            )

            return

        if not first_links:

            job['status'] = 'done'

            job['total_pages'] = 0

            job['progress'] = 100

            job['message'] = (
                'برای این عبارت نتیجه‌ای '
                'در سامانه رسمی یافت نشد.'
            )

            return

        if real_total_pages <= 0:

            real_total_pages = 1

        pages_to_scan = min(
            real_total_pages,
            max_pages
        )

        job['total_pages'] = (
            pages_to_scan
        )

        job['site_total_pages'] = (
            real_total_pages
        )

        job['official_results'] = (
            official_total
        )

        seen = set()

        previous_links = None

        current_html = first_html

        for page in range(
            1,
            pages_to_scan + 1
        ):

            if job['cancel']:
                break

            job['current_page'] = page

            job['message'] = (
                f'در حال بررسی صفحه '
                f'{page} از {pages_to_scan}'
            )

            if page == 1:

                html = first_html

            else:

                html = get_page(
                    session,
                    current_html,
                    page
                )

                current_html = html

            links = get_vote_links(
                html
            )

            # =================================================
            # اگر صفحه آخر خالی برگشت، جست‌وجو موفق تمام شود
            # و نتایج قبلی برای Word حفظ شوند.
            # =================================================

            if not links:

                job['completed_pages'] = page
                job['current_page'] = page
                job['progress'] = 100
                job['status'] = 'done'

                job['message'] = (
                    f'جست‌وجو تکمیل شد. '
                    f'{job["checked"]} رأی بررسی شد و '
                    f'{job["found"]} نتیجه منطبق یافت شد.'
                )

                return

            current_links = set(
                links
            )

            if (
                page > 1
                and previous_links is not None
                and current_links == previous_links
            ):

                job['status'] = 'error'

                job['message'] = (
                    f'صفحه {page} همان آرای '
                    'صفحه قبلی را برگرداند؛ '
                    'صفحه‌بندی سامانه صحیح انجام نشد.'
                )

                return

            previous_links = (
                current_links
            )

            new_links = []

            for url in links:

                if url not in seen:

                    seen.add(url)
                    new_links.append(url)

            for url in new_links:

                if job['cancel']:
                    break

                job['checked'] += 1

                try:

                    vote = fetch_vote(
                        url,
                        session
                    )

                    if vote_matches(
                        vote,
                        query,
                        search_title,
                        search_abstract,
                        search_text
                    ):

                        vote['matched_in'] = (
                            matched_in(
                                vote,
                                query,
                                search_title,
                                search_abstract,
                                search_text
                            )
                        )

                        job['results'].append(
                            vote
                        )

                        job['found'] = len(
                            job['results']
                        )

                except Exception:

                    job[
                        'failed_items'
                    ] += 1

            job[
                'completed_pages'
            ] = page

            job['progress'] = min(
                99,
                round(
                    page
                    / pages_to_scan
                    * 100,
                    1
                )
            )

            job['message'] = (
                f'صفحه {page} از '
                f'{pages_to_scan} بررسی شد.'
            )

            time.sleep(0.15)

        if job['cancel']:

            job['status'] = (
                'cancelled'
            )

            job['message'] = (
                'جست‌وجو به درخواست '
                'کاربر متوقف شد.'
            )

        elif job['status'] == 'running':

            job['status'] = 'done'

            job['progress'] = 100

            job['message'] = (
                'جست‌وجو تکمیل شد.'
            )

    except Exception as e:

        job['status'] = 'error'

        job['message'] = (
            'خطا در ارتباط با سامانه رسمی: '
            + str(e)
        )


# =========================================================
# WORD
# =========================================================

def clean_xml_text(text):
    """
    حذف فقط کاراکترهایی که XML/Word اجازه ذخیره آن‌ها را نمی‌دهد.
    متن فارسی، اعداد، علائم نگارشی و محتوای رأی حفظ می‌شوند.
    """

    if text is None:
        return ''

    text = str(text)

    return ''.join(
        ch for ch in text
        if (
            ch in '\t\n\r'
            or 0x20 <= ord(ch) <= 0xD7FF
            or 0xE000 <= ord(ch) <= 0xFFFD
            or 0x10000 <= ord(ch) <= 0x10FFFF
        )
    )


def rtl(paragraph):

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT
    )

    pPr = (
        paragraph
        ._p
        .get_or_add_pPr()
    )

    pPr.append(
        OxmlElement('w:bidi')
    )

    for run in paragraph.runs:

        rPr = (
            run
            ._r
            .get_or_add_rPr()
        )

        x = OxmlElement(
            'w:rtl'
        )

        x.set(
            qn('w:val'),
            '1'
        )

        rPr.append(x)


def make_doc(jid):
    job = JOBS[jid]
    doc = Document()
    is_laws = job.get('source_id') == QAVANIN_SOURCE_ID

    rtl(doc.add_heading(
        clean_xml_text('قوانین و مقررات یافت‌شده' if is_laws else 'آرای قضایی یافت‌شده'),
        0
    ))
    rtl(doc.add_paragraph(clean_xml_text(f"منبع: {job['source_name']}")))
    rtl(doc.add_paragraph(clean_xml_text(f"عبارت جست‌وجو: {job['query']}")))

    places = []
    if job['search_title']:
        places.append('عنوان')
    if job['search_abstract']:
        places.append('پیام')
    if job['search_text']:
        places.append('متن کامل سند + تحقیق/تنقیح/مرتبط' if is_laws else 'متن رأی')

    rtl(doc.add_paragraph(clean_xml_text('محل جست‌وجو: ' + '، '.join(places))))
    rtl(doc.add_paragraph(clean_xml_text(f"تعداد نتایج: {len(job['results'])}")))
    rtl(doc.add_paragraph(clean_xml_text(
        f"{'اسناد' if is_laws else 'آرای'} بررسی‌شده: {job['checked']}"
    )))

    if job['status'] == 'cancelled':
        rtl(doc.add_paragraph(clean_xml_text(
            'توجه: جست‌وجو پیش از تکمیل توسط کاربر متوقف شده است.'
        )))

    for i, item in enumerate(job['results'], 1):
        rtl(doc.add_heading(clean_xml_text(f"{i}. {item['title']}"), 1))

        locations = item.get('matched_in', [])
        if locations:
            rtl(doc.add_paragraph(clean_xml_text(
                'عبارت موردنظر در: ' + '، '.join(locations)
            )))

        if is_laws:
            meta = []
            if item.get('approval_date'):
                meta.append('تاریخ تصویب: ' + item['approval_date'])
            if item.get('publication_date'):
                meta.append('تاریخ انتشار: ' + item['publication_date'])
            if item.get('document_number'):
                meta.append('شماره: ' + item['document_number'])
            if item.get('category'):
                meta.append('دسته: ' + item['category'])
            for line in meta:
                rtl(doc.add_paragraph(clean_xml_text(line)))
        elif item.get('abstract'):
            rtl(doc.add_paragraph(clean_xml_text('پیام رأی: ' + item['abstract'])))

        rtl(doc.add_paragraph(clean_xml_text(item.get('body', ''))))
        rtl(doc.add_paragraph(clean_xml_text(
            ('منبع سند: ' if is_laws else 'منبع رسمی: ') + item.get('url', '')
        )))
        doc.add_page_break()

    path = f'/tmp/{jid}.docx'
    doc.save(path)
    return path


# =========================================================
# ROUTES
# =========================================================

@app.get('/')
def home():

    return render_template(
        'index.html'
    )


@app.post('/api/search')
def start():

    data = request.get_json(
        force=True
    )

    query = (
        data.get('query')
        or ''
    ).strip()

    search_title = bool(
        data.get(
            'search_title',
            True
        )
    )

    search_abstract = bool(
        data.get(
            'search_abstract',
            False
        )
    )

    search_text = bool(
        data.get(
            'search_text',
            False
        )
    )

    if not query:

        return jsonify(
            error=(
                'عبارت جست‌وجو الزامی است'
            )
        ), 400

    if not (
        search_title
        or search_abstract
        or search_text
    ):

        return jsonify(
            error=(
                'حداقل یکی از گزینه‌های '
                'عنوان، پیام یا متن رأی '
                'را انتخاب کنید.'
            )
        ), 400

    try:

        max_pages = int(
            data.get(
                'max_pages',
                1100
            )
        )

    except Exception:

        max_pages = 1100

    max_pages = min(
        max(
            max_pages,
            1
        ),
        1100
    )

    jid = str(
        uuid.uuid4()
    )

    JOBS[jid] = {

        'job_id': jid,

        'source_id':
            SOURCE_ID,

        'source_name':
            SOURCE_NAME,

        'query':
            query,

        'search_title':
            search_title,

        'search_abstract':
            search_abstract,

        'search_text':
            search_text,

        'status':
            'running',

        'cancel':
            False,

        'checked':
            0,

        'found':
            0,

        'failed_items':
            0,

        'current_page':
            0,

        'completed_pages':
            0,

        'total_pages':
            0,

        'site_total_pages':
            0,

        'official_results':
            None,

        'progress':
            0,

        'results':
            [],

        'message':
            'جست‌وجو آغاز شد.'
    }

    threading.Thread(
        target=worker,
        args=(
            jid,
            query,
            max_pages,
            search_title,
            search_abstract,
            search_text
        ),
        daemon=True
    ).start()

    return jsonify(
        job_id=jid,
        source_id=SOURCE_ID
    )


def qavanin_diagnostic(session=None):
    """Inspect Nezamat's live WordPress REST registration from Render.

    This does not guess a post type. It asks /wp/v2/types which content types
    are exposed through REST, then probes each collection's first page and
    records the server-reported totals. It also records the generic search
    index totals/subtypes for comparison.
    """
    session = session or make_qavanin_session()
    report = {
        'base': QAVANIN_BASE,
        'types': [],
        'search_index': {},
        'errors': []
    }

    try:
        r = session.get(
            urljoin(QAVANIN_BASE, '/wp-json/wp/v2/types'),
            timeout=35,
            allow_redirects=True
        )
        r.raise_for_status()
        types = r.json() if isinstance(r.json(), dict) else {}
    except Exception as e:
        report['errors'].append('types: ' + repr(e))
        types = {}

    for slug, info in types.items():
        if not isinstance(info, dict):
            continue
        rest_base = (info.get('rest_base') or slug or '').strip('/')
        row = {
            'slug': slug,
            'name': info.get('name') or '',
            'rest_base': rest_base,
            'viewable': info.get('viewable'),
            'total': None,
            'total_pages': None,
            'sample_count': 0,
            'sample_titles': [],
            'error': None,
        }
        if rest_base:
            try:
                rr = session.get(
                    urljoin(QAVANIN_BASE, '/wp-json/wp/v2/' + rest_base),
                    params={'page': 1, 'per_page': 5},
                    timeout=35,
                    allow_redirects=True
                )
                rr.raise_for_status()
                data = rr.json()
                if isinstance(data, list):
                    row['sample_count'] = len(data)
                    for item in data[:5]:
                        if not isinstance(item, dict):
                            continue
                        raw = item.get('title') or item.get('name') or ''
                        if isinstance(raw, dict):
                            raw = raw.get('rendered', '')
                        title = norm(BeautifulSoup(unescape(str(raw)), 'html.parser').get_text(' ', strip=True))
                        if title:
                            row['sample_titles'].append(title)
                try:
                    row['total'] = int(rr.headers.get('X-WP-Total', '') or 0)
                except Exception:
                    row['total'] = 0
                try:
                    row['total_pages'] = int(rr.headers.get('X-WP-TotalPages', '') or 0)
                except Exception:
                    row['total_pages'] = 0
            except Exception as e:
                row['error'] = repr(e)
        report['types'].append(row)

    try:
        sr = session.get(
            urljoin(QAVANIN_BASE, '/wp-json/wp/v2/search'),
            params={'page': 1, 'per_page': 100, 'type': 'post', '_fields': 'id,title,url,subtype'},
            timeout=35,
            allow_redirects=True
        )
        sr.raise_for_status()
        data = sr.json()
        subtype_counts = {}
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    st = (item.get('subtype') or '(empty)').strip()
                    subtype_counts[st] = subtype_counts.get(st, 0) + 1
        report['search_index'] = {
            'returned': len(data) if isinstance(data, list) else 0,
            'total': int(sr.headers.get('X-WP-Total', '') or 0),
            'total_pages': int(sr.headers.get('X-WP-TotalPages', '') or 0),
            'subtypes_first_100': subtype_counts,
        }
    except Exception as e:
        report['errors'].append('search: ' + repr(e))

    report['types'].sort(key=lambda x: (-(x.get('total') or 0), x.get('slug') or ''))
    return report


@app.get('/api/qavanin/diagnostic')
def qavanin_diagnostic_route():
    try:
        return jsonify(qavanin_diagnostic())
    except Exception as e:
        return jsonify(error=repr(e)), 500


@app.post('/api/qavanin/search')
def start_qavanin():
    """Create a browser-assisted qavanin.ir search job.

    qavanin.ir blocks server-to-server requests from Render.  The search is
    therefore opened in the user's browser; the companion Chrome extension
    reads the official result pages and posts the extracted rows back here.
    """
    data = request.get_json(force=True)
    query = (data.get('query') or '').strip()
    search_title = bool(data.get('search_title', True))
    search_text = bool(data.get('search_text', False))

    if not query:
        return jsonify(error='عبارت جست‌وجو الزامی است'), 400
    if not (search_title or search_text):
        return jsonify(error='حداقل یکی از گزینه‌های عنوان یا متن قانون را انتخاب کنید.'), 400

    jid = str(uuid.uuid4())
    JOBS[jid] = {
        'job_id': jid,
        'source_id': QAVANIN_SOURCE_ID,
        'source_name': QAVANIN_SOURCE_NAME,
        'query': query,
        'search_title': search_title,
        'search_abstract': False,
        'search_text': search_text,
        'status': 'waiting_browser',
        'cancel': False,
        'checked': 0,
        'found': 0,
        'failed_items': 0,
        'current_page': 0,
        'completed_pages': 0,
        'total_pages': 0,
        'site_total_pages': 0,
        'official_results': None,
        'progress': 0,
        'results': [],
        'message': 'صفحه رسمی qavanin.ir در مرورگر باز می‌شود؛ در انتظار دریافت نتایج رسمی...'
    }

    # Keep the exact ASP.NET MVC checkbox convention used by qavanin.ir.
    params = _qavanin_search_params(
        query, page=1, size=1000,
        search_title=search_title, search_text=search_text
    )
    params.append(('bridge_job', jid))
    params.append(('bridge_target', request.host_url.rstrip('/')))
    browser_url = QAVANIN_LIST + '?' + urlencode(params, doseq=True)

    return jsonify(
        job_id=jid,
        source_id=QAVANIN_SOURCE_ID,
        browser_url=browser_url,
        browser_bridge=True
    )


@app.route('/api/qavanin/browser-import/<jid>', methods=['POST', 'OPTIONS'])
def qavanin_browser_import(jid):
    """Receive official qavanin.ir rows collected inside the user's browser."""
    if request.method == 'OPTIONS':
        r = jsonify(ok=True)
        r.headers['Access-Control-Allow-Origin'] = '*'
        r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        r.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        return r

    job = JOBS.get(jid)
    if not job or job.get('source_id') != QAVANIN_SOURCE_ID:
        return jsonify(error='شناسه جست‌وجوی قوانین معتبر نیست.'), 404

    data = request.get_json(force=True, silent=True) or {}
    if data.get('error'):
        job['status'] = 'error'
        job['message'] = 'خطا در دریافت نتایج از مرورگر: ' + str(data.get('error'))
        r = jsonify(ok=False, error=job['message'])
        r.headers['Access-Control-Allow-Origin'] = '*'
        return r, 400

    try:
        official_total = int(data.get('official_total') or 0)
    except Exception:
        official_total = 0
    try:
        total_pages = int(data.get('total_pages') or 1)
    except Exception:
        total_pages = 1

    # Safety guard: this is the unfiltered catalogue, not a legitimate search.
    if official_total > 100000 and job.get('query'):
        job['status'] = 'error'
        job['message'] = (
            'qavanin.ir فیلتر جست‌وجو را اعمال نکرد و فهرست عمومی را برگرداند؛ '
            'نتیجه برای جلوگیری از ورود داده نادرست رد شد.'
        )
        r = jsonify(ok=False, error=job['message'])
        r.headers['Access-Control-Allow-Origin'] = '*'
        return r, 400

    incoming = data.get('results') or []
    clean = []
    seen = set()
    for item in incoming:
        if not isinstance(item, dict):
            continue
        url = (item.get('url') or '').strip()
        title = norm(item.get('title') or '')
        if not url or not title or not url.startswith(QAVANIN_BASE):
            continue
        if url in seen:
            continue
        seen.add(url)
        body = norm(item.get('body') or '')
        clean.append({
            'url': url,
            'title': title,
            'abstract': '',
            'body': body,
            'text': body,
            'source': QAVANIN_SOURCE_NAME,
            'approval_date': norm(item.get('approval_date') or ''),
            'publication_date': '',
            'document_number': '',
            'category': norm(item.get('approver') or ''),
            'approver': norm(item.get('approver') or ''),
            'matched_in': ([] if not job.get('search_title') else ['عنوان']) +
                          ([] if not job.get('search_text') else ['متن قانون/مقرره'])
        })

    # If the official page says N hits, do not silently accept a partial import.
    if official_total and len(clean) < official_total:
        job['status'] = 'error'
        job['checked'] = len(clean)
        job['found'] = len(clean)
        job['official_results'] = official_total
        job['results'] = clean
        job['message'] = (
            f'qavanin.ir تعداد {official_total} یافته اعلام کرد، اما مرورگر فقط '
            f'{len(clean)} نتیجه را منتقل کرد؛ نتیجه ناقص است.'
        )
        r = jsonify(ok=False, error=job['message'])
        r.headers['Access-Control-Allow-Origin'] = '*'
        return r, 400

    job['results'] = clean
    job['checked'] = len(clean)
    job['found'] = len(clean)
    job['official_results'] = official_total or len(clean)
    job['current_page'] = total_pages
    job['completed_pages'] = total_pages
    job['total_pages'] = total_pages
    job['site_total_pages'] = total_pages
    job['progress'] = 100
    job['status'] = 'done'
    job['message'] = (
        f'جست‌وجوی رسمی qavanin.ir تکمیل شد؛ '
        f'{job["official_results"]} یافته رسمی دریافت شد.'
    )
    r = jsonify(ok=True, found=len(clean), official_total=job['official_results'])
    r.headers['Access-Control-Allow-Origin'] = '*'
    return r


@app.get('/api/status/<jid>')
def status(jid):

    job = JOBS.get(
        jid
    )

    if not job:

        return jsonify(
            error='یافت نشد'
        ), 404

    return jsonify(

        job_id=
            job['job_id'],

        source_id=
            job['source_id'],

        source_name=
            job['source_name'],

        status=
            job['status'],

        checked=
            job['checked'],

        found=
            job['found'],

        failed_items=
            job['failed_items'],

        current_page=
            job['current_page'],

        completed_pages=
            job['completed_pages'],

        total_pages=
            job['total_pages'],

        progress=
            job['progress'],

        message=
            job['message']
    )


@app.post('/api/cancel/<jid>')
def cancel(jid):

    job = JOBS.get(
        jid
    )

    if not job:

        return jsonify(
            error='یافت نشد'
        ), 404

    job['cancel'] = True

    return jsonify(
        ok=True
    )


@app.get('/api/download/<jid>')
def download(jid):

    if jid not in JOBS:

        return (
            'Not found',
            404
        )

    job = JOBS[jid]

    filename = (
        'national-laws-and-regulations.docx'
        if job.get('source_id') == QAVANIN_SOURCE_ID
        else 'national-judicial-decisions.docx'
    )

    return send_file(
        make_doc(jid),
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    app.run(
        host='0.0.0.0',
        port=int(
            os.getenv(
                'PORT',
                5000
            )
        )
    )