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
import sqlite3

from urllib.parse import urljoin


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
QAVANIN_PAGE_SIZE = 25


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
# LAWS DATABASE
# =========================================================

DB_PATH = os.getenv(
    'LAWS_DB_PATH',
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'laws.db'
    )
)


def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_laws_db():
    conn = get_db()
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS laws (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                qavanin_id TEXT UNIQUE,
                title TEXT NOT NULL DEFAULT '',
                body TEXT NOT NULL DEFAULT '',
                law_type TEXT NOT NULL DEFAULT '',
                approver TEXT NOT NULL DEFAULT '',
                approval_date TEXT NOT NULL DEFAULT '',
                notification_number TEXT NOT NULL DEFAULT '',
                notification_date TEXT NOT NULL DEFAULT '',
                publication_date TEXT NOT NULL DEFAULT '',
                official_gazette_number TEXT NOT NULL DEFAULT '',
                execution_date TEXT NOT NULL DEFAULT '',
                law_status TEXT NOT NULL DEFAULT '',
                source_url TEXT NOT NULL DEFAULT '',
                attribute_url TEXT NOT NULL DEFAULT '',
                status_url TEXT NOT NULL DEFAULT '',
                related_url TEXT NOT NULL DEFAULT '',
                subject_url TEXT NOT NULL DEFAULT '',
                image_url TEXT NOT NULL DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            'CREATE INDEX IF NOT EXISTS idx_laws_qavanin_id '
            'ON laws(qavanin_id)'
        )
        conn.execute(
            'CREATE INDEX IF NOT EXISTS idx_laws_title '
            'ON laws(title)'
        )
        conn.execute(
            'CREATE INDEX IF NOT EXISTS idx_laws_approval_date '
            'ON laws(approval_date)'
        )
        conn.commit()
    finally:
        conn.close()


def build_qavanin_urls(qavanin_id):
    qid = str(qavanin_id or '').strip()
    return {
        'source_url':
            f'{QAVANIN_BASE}/Law/TreeText/?IDS={qid}',
        'attribute_url':
            f'{QAVANIN_BASE}/Law/Attribute/?IDS={qid}',
        'status_url':
            f'{QAVANIN_BASE}/Law/StatusIndex/?IDS={qid}',
        'related_url':
            f'{QAVANIN_BASE}/Law/RelatedIndex/?IDS={qid}',
        'subject_url':
            f'{QAVANIN_BASE}/Law/SubjectIndex/?IDS={qid}',
        'image_url':
            f'{QAVANIN_BASE}/Law/ImageText/?IDS={qid}'
    }


def save_law(
    qavanin_id,
    title='',
    body='',
    law_type='',
    approver='',
    approval_date='',
    notification_number='',
    notification_date='',
    publication_date='',
    official_gazette_number='',
    execution_date='',
    law_status=''
):
    qavanin_id = str(qavanin_id or '').strip()

    if not qavanin_id:
        raise ValueError('شناسه IDS قانون الزامی است.')

    urls = build_qavanin_urls(qavanin_id)
    conn = get_db()

    try:
        conn.execute(
            """
            INSERT INTO laws (
                qavanin_id,
                title,
                body,
                law_type,
                approver,
                approval_date,
                notification_number,
                notification_date,
                publication_date,
                official_gazette_number,
                execution_date,
                law_status,
                source_url,
                attribute_url,
                status_url,
                related_url,
                subject_url,
                image_url
            )
            VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?
            )
            ON CONFLICT(qavanin_id)
            DO UPDATE SET
                title = excluded.title,
                body = excluded.body,
                law_type = excluded.law_type,
                approver = excluded.approver,
                approval_date = excluded.approval_date,
                notification_number = excluded.notification_number,
                notification_date = excluded.notification_date,
                publication_date = excluded.publication_date,
                official_gazette_number = excluded.official_gazette_number,
                execution_date = excluded.execution_date,
                law_status = excluded.law_status,
                source_url = excluded.source_url,
                attribute_url = excluded.attribute_url,
                status_url = excluded.status_url,
                related_url = excluded.related_url,
                subject_url = excluded.subject_url,
                image_url = excluded.image_url,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                qavanin_id,
                norm(title),
                norm(body),
                norm(law_type),
                norm(approver),
                norm(approval_date),
                norm(notification_number),
                norm(notification_date),
                norm(publication_date),
                norm(official_gazette_number),
                norm(execution_date),
                norm(law_status),
                urls['source_url'],
                urls['attribute_url'],
                urls['status_url'],
                urls['related_url'],
                urls['subject_url'],
                urls['image_url']
            )
        )
        conn.commit()
    finally:
        conn.close()


def search_laws_database(
    query,
    search_title=True,
    search_text=False,
    limit=10000
):
    include, exclude = parse_query(query)

    if not include:
        return []

    conn = get_db()

    try:
        rows = conn.execute(
            """
            SELECT *
            FROM laws
            ORDER BY approval_date DESC, id DESC
            """
        ).fetchall()
    finally:
        conn.close()

    results = []

    for row in rows:
        law = dict(row)
        selected = []

        if search_title:
            selected.append(law.get('title', ''))

        if search_text:
            selected.append(law.get('body', ''))

        if not matches(' '.join(selected), query):
            continue

        locations = []

        if search_title and matches(law.get('title', ''), query):
            locations.append('عنوان')

        if search_text and matches(law.get('body', ''), query):
            locations.append('متن قانون')

        law['matched_in'] = locations
        law['url'] = law.get('source_url', '')
        law['abstract'] = ''
        law['source'] = QAVANIN_SOURCE_NAME

        results.append(law)

        if len(results) >= limit:
            break

    return results


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
# QAVANIN.IR - OFFICIAL LAWS ENGINE
# =========================================================

def make_qavanin_session():

    session = requests.Session()

    session.headers.update({
        'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/153.0.0.0 Safari/537.36',

        'Accept':
            'text/html,application/xhtml+xml,application/xml;'
            'q=0.9,image/avif,image/webp,*/*;q=0.8',

        'Accept-Language':
            'fa-IR,fa;q=0.9,en-US;q=0.8,en;q=0.7',

        'Upgrade-Insecure-Requests':
            '1',

        'Referer':
            QAVANIN_LIST
    })

    return session


def qavanin_is_challenge(html):

    lower = (
        html
        or ''
    ).lower()

    markers = (
        'transferring to the website',
        'error-section--waiting',
        'istehrantimezone',
        'arvancloud'
    )

    return any(
        x in lower
        for x in markers
    )


def qavanin_search_page(
    session,
    query,
    page=1,
    search_title=True,
    search_text=False
):

    params = [
        ('CAPTION', query),
        ('Zone', ''),
    ]

    if search_title:
        params.append(
            ('IsTitleSearch', 'true')
        )

    params.append(
        ('IsTitleSearch', 'false')
    )

    if search_text:
        params.append(
            ('IsTextSearch', 'true')
        )

    params.append(
        ('IsTextSearch', 'false')
    )

    params.extend([
        ('_isLaw', 'false'),
        ('_isRegulation', 'false'),
        ('_IsVote', 'false'),
        ('_isOpenion', 'false'),
        ('SeachTextType', '3'),
        ('fromApproveDate', ''),
        ('APPROVEDATE', ''),
        ('IsTitleSubject', 'False'),
        ('IsMain', ''),
        ('COMMANDNO', ''),
        ('fromCommandDate', ''),
        ('COMMANDDATE', ''),
        ('NEWSPAPERNO', ''),
        ('fromNewspaperDate', ''),
        ('NEWSPAPERDATE', ''),
        ('SortColumn', 'APPROVEDATE'),
        ('SortDesc', 'True'),
        ('Report_ID', ''),
        ('PageNumber', str(page)),
        ('page', str(page)),
        ('size', str(QAVANIN_PAGE_SIZE)),
        ('txtZone', ''),
        ('txtSubjects', ''),
        ('txtExecutors', ''),
        ('txtApprovers', ''),
        ('txtLawStatus', ''),
        ('txtLawTypes', ''),
    ])

    response = session.get(
        QAVANIN_LIST,
        params=params,
        timeout=30,
        allow_redirects=True
    )

    response.raise_for_status()

    response.encoding = (
        response.apparent_encoding
        or 'utf-8'
    )

    if qavanin_is_challenge(
        response.text
    ):

        raise RuntimeError(
            'سامانه قوانین به این سرور '
            'صفحه امنیتی ArvanCloud برگرداند.'
        )

    return response.text


def get_law_links(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    results = []
    seen = set()

    for a in soup.find_all(
        'a',
        href=True
    ):

        href = (
            a.get('href')
            or ''
        ).strip()

        href_lower = (
            href.lower()
        )

        if '/law/' not in href_lower:
            continue

        if any(
            x in href_lower
            for x in (
                '/law/index',
                '/law/search',
                'javascript:',
                '#'
            )
        ):
            continue

        url = urljoin(
            QAVANIN_BASE,
            href
        )

        if url in seen:
            continue

        title = norm(
            a.get_text(
                ' ',
                strip=True
            )
        )

        if not title:
            continue

        seen.add(url)

        results.append({
            'url': url,
            'title': title
        })

    return results


def extract_qavanin_total_results(html):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    page_text = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    patterns = [
        r'تعداد\s*یافته\s*ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*یافته‌ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*نتایج\s*[:：]?\s*([0-9۰-۹,٬]+)'
    ]

    for pattern in patterns:

        m = re.search(
            pattern,
            page_text
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


def fetch_law(
    url,
    session
):

    response = session.get(
        url,
        timeout=30,
        allow_redirects=True
    )

    response.raise_for_status()

    response.encoding = (
        response.apparent_encoding
        or 'utf-8'
    )

    if qavanin_is_challenge(
        response.text
    ):

        raise RuntimeError(
            'صفحه قانون توسط لایه امنیتی '
            'ArvanCloud مسدود شد.'
        )

    soup = BeautifulSoup(
        response.text,
        'html.parser'
    )

    for tag in soup([
        'script',
        'style',
        'noscript'
    ]):

        tag.decompose()

    full_text = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    title = ''

    for selector in (
        'h1',
        'h2',
        'h3'
    ):

        h = soup.find(
            selector
        )

        if h:

            candidate = norm(
                h.get_text(
                    ' ',
                    strip=True
                )
            )

            if candidate:

                title = candidate
                break

    if not title:

        title = (
            'قانون یا مقرره'
        )

    body = full_text

    starts = [
        'متن مصوبه',
        'متن قانون',
        'متن مقرره',
        'ماده 1',
        'ماده ۱'
    ]

    positions = [
        body.find(x)
        for x in starts
        if body.find(x) != -1
    ]

    if positions:

        body = body[
            min(positions):
        ]

    return {
        'url': url,
        'title': title,
        'abstract': '',
        'body': norm(body),
        'text': full_text,
        'source':
            QAVANIN_SOURCE_NAME
    }


def law_matches(
    law,
    query,
    search_title,
    search_text
):

    selected = []

    if search_title:

        selected.append(
            law.get(
                'title',
                ''
            )
        )

    if search_text:

        selected.append(
            law.get(
                'body',
                ''
            )
        )

    return matches(
        ' '.join(selected),
        query
    )


def law_matched_in(
    law,
    query,
    search_title,
    search_text
):

    locations = []

    if (
        search_title
        and matches(
            law.get(
                'title',
                ''
            ),
            query
        )
    ):

        locations.append(
            'عنوان'
        )

    if (
        search_text
        and matches(
            law.get(
                'body',
                ''
            ),
            query
        )
    ):

        locations.append(
            'متن قانون'
        )

    return locations


# =========================================================
# QAVANIN DATABASE WORKER
# =========================================================

def qavanin_worker(
    jid,
    query,
    max_pages,
    search_title,
    search_text
):

    job = JOBS[jid]

    try:

        job['message'] = (
            'در حال جست‌وجو در '
            'پایگاه داده قوانین...'
        )

        results = search_laws_database(
            query,
            search_title=search_title,
            search_text=search_text
        )

        job['checked'] = len(
            results
        )

        job['results'] = (
            results
        )

        job['found'] = len(
            results
        )

        job['total_pages'] = (
            1 if results else 0
        )

        job['site_total_pages'] = (
            1 if results else 0
        )

        job['current_page'] = (
            1 if results else 0
        )

        job['completed_pages'] = (
            1 if results else 0
        )

        job['official_results'] = (
            len(results)
        )

        job['progress'] = 100

        job['status'] = 'done'

        if results:

            job['message'] = (
                f'جست‌وجوی قوانین تکمیل شد. '
                f'{len(results)} نتیجه منطبق یافت شد.'
            )

        else:

            job['message'] = (
                'در پایگاه داده قوانین هنوز '
                'سند منطبقی وجود ندارد. '
                'زیرساخت دیتابیس آماده است و '
                'مرحله بعد واردکردن داده‌های '
                'قوانین است.'
            )

    except Exception as e:

        job['status'] = 'error'

        job['message'] = (
            'خطا در جست‌وجوی '
            'پایگاه داده قوانین: '
            + str(e)
        )


# =========================================================
# JUDGMENTS WORKER
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
                    new_links.append(
                        url
                    )

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

            time.sleep(
                0.15
            )

        if job['cancel']:

            job['status'] = (
                'cancelled'
            )

            job['message'] = (
                'جست‌وجو به درخواست '
                'کاربر متوقف شد.'
            )

        elif job['status'] == 'running':

            job['status'] = (
                'done'
            )

            job['progress'] = 100

            job['message'] = (
                'جست‌وجو تکمیل شد.'
            )

    except Exception as e:

        job['status'] = (
            'error'
        )

        job['message'] = (
            'خطا در ارتباط با '
            'سامانه رسمی: '
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

    is_laws = (
        job.get('source_id')
        == QAVANIN_SOURCE_ID
    )

    rtl(
        doc.add_heading(
            clean_xml_text(
                'قوانین و مقررات یافت‌شده'
                if is_laws
                else 'آرای قضایی یافت‌شده'
            ),
            0
        )
    )

    rtl(
        doc.add_paragraph(
            clean_xml_text(
                f"منبع: "
                f"{job['source_name']}"
            )
        )
    )

    rtl(
        doc.add_paragraph(
            clean_xml_text(
                f"عبارت جست‌وجو: "
                f"{job['query']}"
            )
        )
    )

    places = []

    if job['search_title']:
        places.append(
            'عنوان'
        )

    if job['search_abstract']:
        places.append(
            'پیام'
        )

    if job['search_text']:
        places.append(
            'متن قانون'
            if is_laws
            else 'متن رأی'
        )

    rtl(
        doc.add_paragraph(
            clean_xml_text(
                'محل جست‌وجو: '
                + '، '.join(places)
            )
        )
    )

    rtl(
        doc.add_paragraph(
            clean_xml_text(
                f"تعداد نتایج: "
                f"{len(job['results'])}"
            )
        )
    )

    rtl(
        doc.add_paragraph(
            clean_xml_text(
                (
                    'اسناد بررسی‌شده: '
                    if is_laws
                    else 'آرای بررسی‌شده: '
                )
                + f"{job['checked']}"
            )
        )
    )

    if job['status'] == 'cancelled':

        rtl(
            doc.add_paragraph(
                clean_xml_text(
                    'توجه: جست‌وجو پیش از '
                    'تکمیل توسط کاربر متوقف شده است.'
                )
            )
        )

    for i, vote in enumerate(
        job['results'],
        1
    ):

        rtl(
            doc.add_heading(
                clean_xml_text(
                    f"{i}. {vote['title']}"
                ),
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
                    clean_xml_text(
                        'عبارت موردنظر در: '
                        + '، '.join(
                            locations
                        )
                    )
                )
            )

        if vote.get(
            'abstract'
        ):

            rtl(
                doc.add_paragraph(
                    clean_xml_text(
                        'پیام رأی: '
                        + vote[
                            'abstract'
                        ]
                    )
                )
            )

        rtl(
            doc.add_paragraph(
                clean_xml_text(
                    vote.get(
                        'body',
                        ''
                    )
                )
            )
        )

        rtl(
            doc.add_paragraph(
                clean_xml_text(
                    'منبع رسمی: '
                    + vote.get(
                        'url',
                        ''
                    )
                )
            )
        )

        doc.add_page_break()

    path = (
        f'/tmp/{jid}.docx'
    )

    doc.save(
        path
    )

    return path


# =========================================================
# INITIALIZE DATABASE
# =========================================================

init_laws_db()


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


@app.post('/api/qavanin/search')
def start_qavanin():

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
        or search_text
    ):

        return jsonify(
            error=(
                'حداقل یکی از گزینه‌های '
                'عنوان یا متن قانون را انتخاب کنید.'
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

        'job_id':
            jid,

        'source_id':
            QAVANIN_SOURCE_ID,

        'source_name':
            QAVANIN_SOURCE_NAME,

        'query':
            query,

        'search_title':
            search_title,

        'search_abstract':
            False,

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
            'جست‌وجوی قوانین آغاز شد.'
    }

    threading.Thread(
        target=qavanin_worker,
        args=(
            jid,
            query,
            max_pages,
            search_title,
            search_text
        ),
        daemon=True
    ).start()

    return jsonify(
        job_id=jid,
        source_id=QAVANIN_SOURCE_ID
    )


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
        if job.get(
            'source_id'
        ) == QAVANIN_SOURCE_ID
        else
        'national-judicial-decisions.docx'
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
