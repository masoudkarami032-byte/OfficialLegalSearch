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

from urllib.parse import urljoin


app = Flask(__name__)

BASE = 'https://ara.jri.ac.ir'
LIST = BASE + '/Judge/Index'

JOBS = {}

SOURCE_ID = 'national_judgments'
SOURCE_NAME = 'سامانه ملی آرای قضایی پژوهشگاه قوه قضاییه'


# =========================================================
# TEXT
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

    text = norm(text)

    include, exclude = parse_query(query)

    if not include:
        return False

    for item in include:
        if norm(item) not in text:
            return False

    for item in exclude:
        if norm(item) in text:
            return False

    return True


# =========================================================
# SEARCH RESULT PAGE
# =========================================================

def get_vote_links(html):

    soup = BeautifulSoup(html, 'html.parser')

    links = []
    seen = set()

    for a in soup.find_all('a', href=True):

        href = a.get('href', '')

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


def extract_total_results(html):

    soup = BeautifulSoup(html, 'html.parser')

    text = norm(
        soup.get_text(' ', strip=True)
    )

    patterns = [
        r'تعداد\s*یافته\s*ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*یافته‌ها\s*[:：]?\s*([0-9۰-۹,٬]+)'
    ]

    for pattern in patterns:

        match = re.search(pattern, text)

        if not match:
            continue

        value = match.group(1)

        value = value.translate(
            str.maketrans(
                '۰۱۲۳۴۵۶۷۸۹',
                '0123456789'
            )
        )

        value = (
            value.replace(',', '')
                 .replace('٬', '')
        )

        try:
            return int(value)
        except ValueError:
            pass

    return None


def extract_total_pages(html, page_size=25):

    soup = BeautifulSoup(html, 'html.parser')

    select = soup.find(
        'select',
        attrs={'name': 'PageNumbers'}
    )

    pages = []

    if select:

        for option in select.find_all('option'):

            value = option.get(
                'value',
                ''
            ).strip()

            if value.isdigit():
                pages.append(int(value))

    if pages:
        return max(pages)

    total_results = extract_total_results(html)

    if total_results is not None:
        return max(
            1,
            math.ceil(
                total_results / page_size
            )
        )

    return 1


# =========================================================
# SESSION
# =========================================================

def make_session():

    session = requests.Session()

    session.headers.update({
        'User-Agent':
            'Mozilla/5.0 (Linux; Android 10) '
            'AppleWebKit/537.36 '
            '(KHTML, like Gecko) '
            'Chrome/120.0 Safari/537.36',

        'Accept':
            'text/html,application/xhtml+xml,'
            'application/xml;q=0.9,*/*;q=0.8',

        'Accept-Language':
            'fa-IR,fa;q=0.9,en-US;q=0.7,en;q=0.6',

        'Referer': LIST
    })

    return session


# =========================================================
# OFFICIAL SEARCH
# =========================================================

def search_first_page(
    session,
    query,
    search_title,
    search_abstract,
    search_text,
    page_size=25
):

    # Establish cookies/session first
    first = session.get(
        LIST,
        timeout=30
    )

    first.raise_for_status()

    payload = [
        ('Title', query),

        (
            'IsTitleSearch',
            'true' if search_title else 'false'
        ),

        (
            'IsAbstractSearch',
            'true' if search_abstract else 'false'
        ),

        (
            'IsTextSearch',
            'true' if search_text else 'false'
        ),

        # Search all legal groups
        ('_IsOfficial', 'true'),
        ('_IsCivil', 'true'),
        ('_IsPenal', 'true'),

        # "بخشی از کلمه ها"
        ('SeachTextType', '3'),

        ('SortColumn', 'Overdate'),
        ('SortDesc', 'True'),

        ('PageNumber', '1'),
        ('PageNumbers', '1'),
        ('PageSize', str(page_size))
    ]

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


def fetch_result_page(
    session,
    page,
    page_size=25
):

    payload = {
        'PageNumbers': str(page),
        'PageNumber': str(page),
        'PageSize': str(page_size)
    }

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
# VOTE PARSER
# =========================================================

def extract_labeled_section(text, start_labels, end_labels):

    text = norm(text)

    start_pos = -1

    for label in start_labels:

        pos = text.find(label)

        if pos != -1:
            start_pos = pos + len(label)
            break

    if start_pos == -1:
        return ''

    end_pos = len(text)

    for label in end_labels:

        pos = text.find(
            label,
            start_pos
        )

        if (
            pos != -1
            and pos < end_pos
        ):
            end_pos = pos

    return norm(
        text[start_pos:end_pos]
    )


def fetch_vote(url, session):

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

    # -----------------------------------------------------
    # TITLE
    # -----------------------------------------------------

    title = ''

    # First try page heading
    heading = (
        soup.find('h1')
        or soup.find('h2')
    )

    if heading:
        title = norm(
            heading.get_text(
                ' ',
                strip=True
            )
        )

    # The JRI pages often expose "عنوان ... پیام ..."
    if not title or title == 'سامانه ملی آرا':

        extracted_title = extract_labeled_section(
            full_text,
            ['عنوان'],
            [
                'پیام',
                'مستندات',
                'شماره دادنامه'
            ]
        )

        if extracted_title:
            title = extracted_title

    if not title:

        if soup.title:
            title = norm(
                soup.title.get_text(
                    ' ',
                    strip=True
                )
            )
        else:
            title = 'رأی قضایی'

    # -----------------------------------------------------
    # ABSTRACT / پیام
    # -----------------------------------------------------

    abstract = extract_labeled_section(
        full_text,
        ['پیام'],
        [
            'مستندات',
            'شماره دادنامه',
            'گروه رأی',
            'آراء منتخب پرونده'
        ]
    )

    # -----------------------------------------------------
    # BODY / متن رأی
    # -----------------------------------------------------

    body = ''

    possible_starts = [
        'رأی دادگاه بدوی',
        'رای دادگاه بدوی',
        'رأی دادگاه تجدیدنظر',
        'رای دادگاه تجدیدنظر',
        'رأی شعبه دیوان عالی کشور',
        'رای شعبه دیوان عالی کشور',
        'رأی شعبه',
        'رای شعبه'
    ]

    body_start = -1

    for marker in possible_starts:

        pos = full_text.find(marker)

        if pos != -1:

            if (
                body_start == -1
                or pos < body_start
            ):
                body_start = pos

    if body_start != -1:
        body = full_text[body_start:]
    else:
        body = full_text

    # Remove footer when possible
    footer_markers = [
        'نقد رأی',
        'نقد رای',
        'تعدادموافق',
        'تماس با ما'
    ]

    footer_position = len(body)

    for marker in footer_markers:

        pos = body.find(marker)

        if (
            pos != -1
            and pos < footer_position
        ):
            footer_position = pos

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
# LOCAL MATCHING
# =========================================================

def vote_matches(
    vote,
    query,
    search_title,
    search_abstract,
    search_text
):

    selected_parts = []

    if search_title:
        selected_parts.append(
            vote.get('title', '')
        )

    if search_abstract:
        selected_parts.append(
            vote.get('abstract', '')
        )

    if search_text:
        selected_parts.append(
            vote.get('body', '')
        )

    combined = norm(
        ' '.join(selected_parts)
    )

    return matches(
        combined,
        query
    )


def matched_in(
    vote,
    query,
    search_title,
    search_abstract,
    search_text
):

    places = []

    if (
        search_title
        and matches(
            vote.get('title', ''),
            query
        )
    ):
        places.append('عنوان')

    if (
        search_abstract
        and matches(
            vote.get('abstract', ''),
            query
        )
    ):
        places.append('پیام')

    if (
        search_text
        and matches(
            vote.get('body', ''),
            query
        )
    ):
        places.append('متن')

    return places


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

    PAGE_SIZE = 25

    try:

        job['message'] = (
            'در حال ارسال جست‌وجو '
            'به سامانه رسمی...'
        )

        first_html = search_first_page(
            session,
            query,
            search_title,
            search_abstract,
            search_text,
            PAGE_SIZE
        )

        first_links = get_vote_links(
            first_html
        )

        total_results = extract_total_results(
            first_html
        )

        site_total_pages = extract_total_pages(
            first_html,
            PAGE_SIZE
        )

        if total_results is not None:
            job['official_results'] = total_results

        if not first_links:

            job['status'] = 'done'
            job['total_pages'] = 0
            job['progress'] = 100

            job['message'] = (
                'سامانه رسمی برای این '
                'جست‌وجو نتیجه‌ای برنگرداند.'
            )

            return

        pages_to_scan = min(
            site_total_pages,
            max_pages
        )

        job['total_pages'] = pages_to_scan
        job['site_total_pages'] = site_total_pages

        seen = set()

        previous_page_links = None

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
                links = first_links

            else:

                try:

                    html = fetch_result_page(
                        session,
                        page,
                        PAGE_SIZE
                    )

                except Exception as e:

                    job['status'] = 'error'

                    job['message'] = (
                        f'خطا در دریافت صفحه '
                        f'{page}: {e}'
                    )

                    return

                links = get_vote_links(
                    html
                )

            if not links:

                job['status'] = 'error'

                job['message'] = (
                    f'صفحه {page} دریافت شد '
                    'اما رأیی در آن پیدا نشد.'
                )

                return

            current_page_links = set(
                links
            )

            if (
                page > 1
                and previous_page_links is not None
                and current_page_links
                == previous_page_links
            ):

                job['status'] = 'error'

                job['message'] = (
                    f'صفحه {page} همان نتایج '
                    'صفحه قبلی را برگرداند.'
                )

                return

            previous_page_links = (
                current_page_links
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

                    job['failed_items'] += 1

            job['completed_pages'] = page

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

            job['status'] = 'cancelled'

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

def rtl(paragraph):

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT
    )

    pPr = paragraph._p.get_or_add_pPr()

    pPr.append(
        OxmlElement('w:bidi')
    )

    for run in paragraph.runs:

        rPr = run._r.get_or_add_rPr()

        element = OxmlElement('w:rtl')

        element.set(
            qn('w:val'),
            '1'
        )

        rPr.append(element)


def make_doc(jid):

    job = JOBS[jid]

    doc = Document()

    rtl(
        doc.add_heading(
            'آرای قضایی یافت‌شده',
            0
        )
    )

    rtl(
        doc.add_paragraph(
            f"منبع: {job['source_name']}"
        )
    )

    rtl(
        doc.add_paragraph(
            f"عبارت جست‌وجو: "
            f"{job['query']}"
        )
    )

    search_places = []

    if job['search_title']:
        search_places.append('عنوان')

    if job['search_abstract']:
        search_places.append('پیام')

    if job['search_text']:
        search_places.append('متن رأی')

    rtl(
        doc.add_paragraph(
            'محل جست‌وجو: '
            + '، '.join(search_places)
        )
    )

    rtl(
        doc.add_paragraph(
            f"تعداد نتایج یافت‌شده: "
            f"{len(job['results'])}"
        )
    )

    rtl(
        doc.add_paragraph(
            f"تعداد آرای بررسی‌شده: "
            f"{job['checked']}"
        )
    )

    if job['status'] == 'cancelled':

        rtl(
            doc.add_paragraph(
                'توجه: جست‌وجو پیش از '
                'تکمیل توسط کاربر متوقف شده است.'
            )
        )

    for i, vote in enumerate(
        job['results'],
        1
    ):

        rtl(
            doc.add_heading(
                f"{i}. {vote['title']}",
                1
            )
        )

        locations = vote.get(
            'matched_in',
            []
        )

        if locations:

            rtl(
                doc.add_paragraph(
                    'عبارت موردنظر در: '
                    + '، '.join(locations)
                )
            )

        if vote.get('abstract'):

            rtl(
                doc.add_paragraph(
                    'پیام رأی: '
                    + vote['abstract']
                )
            )

        rtl(
            doc.add_paragraph(
                vote['body']
            )
        )

        rtl(
            doc.add_paragraph(
                'منبع رسمی: '
                + vote['url']
            )
        )

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
            error='عبارت جست‌وجو الزامی است'
        ), 400

    if not (
        search_title
        or search_abstract
        or search_text
    ):

        return jsonify(
            error=(
                'حداقل یکی از گزینه‌های '
                'عنوان، پیام یا متن را انتخاب کنید.'
            )
        ), 400

    try:

        requested_max = int(
            data.get(
                'max_pages',
                1100
            )
        )

    except (TypeError, ValueError):

        requested_max = 1100

    max_pages = min(
        max(
            requested_max,
            1
        ),
        1100
    )

    jid = str(
        uuid.uuid4()
    )

    JOBS[jid] = {

        'job_id': jid,

        'source_id': SOURCE_ID,
        'source_name': SOURCE_NAME,

        'query': query,

        'search_title': search_title,
        'search_abstract': search_abstract,
        'search_text': search_text,

        'status': 'running',
        'cancel': False,

        'checked': 0,
        'found': 0,
        'failed_items': 0,

        'current_page': 0,
        'completed_pages': 0,
        'total_pages': max_pages,

        'progress': 0,

        'results': [],

        'message': 'جست‌وجو آغاز شد.'
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


@app.get('/api/status/<jid>')
def status(jid):

    job = JOBS.get(jid)

    if not job:

        return jsonify(
            error='یافت نشد'
        ), 404

    return jsonify(

        job_id=job['job_id'],

        source_id=job['source_id'],
        source_name=job['source_name'],

        status=job['status'],

        checked=job['checked'],
        found=job['found'],

        failed_items=job[
            'failed_items'
        ],

        current_page=job[
            'current_page'
        ],

        completed_pages=job[
            'completed_pages'
        ],

        total_pages=job[
            'total_pages'
        ],

        progress=job['progress'],

        message=job['message']
    )


@app.post('/api/cancel/<jid>')
def cancel(jid):

    job = JOBS.get(jid)

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
        return 'Not found', 404

    return send_file(
        make_doc(jid),
        as_attachment=True,
        download_name=(
            'national-judicial-decisions.docx'
        )
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
