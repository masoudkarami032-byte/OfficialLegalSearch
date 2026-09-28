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


# ---------------------------------------------------------
# Text helpers
# ---------------------------------------------------------

def norm(s):
    if not s:
        return ''

    s = str(s)
    s = s.replace('ي', 'ی').replace('ك', 'ک').replace('\u200c', ' ')
    return re.sub(r'\s+', ' ', s).strip()


def parse_query(q):
    parts = re.split(r'\s+(AND|NOT)\s+', q.strip(), flags=re.I)

    inc = []
    exc = []
    mode = 'AND'

    for p in parts:
        if p.upper() in ('AND', 'NOT'):
            mode = p.upper()
            continue

        p = p.strip().strip('"“”')

        if p:
            if mode == 'NOT':
                exc.append(p)
            else:
                inc.append(p)

        mode = 'AND'

    return inc, exc


def matches(text, q):
    text = norm(text)

    inc, exc = parse_query(q)

    return (
        bool(inc)
        and all(norm(x) in text for x in inc)
        and not any(norm(x) in text for x in exc)
    )


# ---------------------------------------------------------
# Site parsing
# ---------------------------------------------------------

def get_vote_links(html):
    soup = BeautifulSoup(html, 'html.parser')

    links = []
    local_seen = set()

    for a in soup.find_all('a', href=True):
        href = a.get('href', '')

        if '/Judge/Text/' not in href:
            continue

        url = urljoin(BASE, href.split('?')[0])

        if url not in local_seen:
            local_seen.add(url)
            links.append(url)

    return links


def extract_total_results(html):
    """
    Example:
    تعداد یافته ها : 157
    """

    soup = BeautifulSoup(html, 'html.parser')
    text = norm(soup.get_text(' ', strip=True))

    patterns = [
        r'تعداد\s*یافته\s*ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
        r'تعداد\s*یافته‌ها\s*[:：]?\s*([0-9۰-۹,٬]+)',
    ]

    for pattern in patterns:
        m = re.search(pattern, text)

        if m:
            value = m.group(1)

            persian_digits = '۰۱۲۳۴۵۶۷۸۹'
            english_digits = '0123456789'

            value = value.translate(
                str.maketrans(persian_digits, english_digits)
            )

            value = value.replace(',', '').replace('٬', '')

            try:
                return int(value)
            except ValueError:
                pass

    return None


def extract_total_pages(html, page_size=25):
    """
    Primary source:
    <select id="PageNumbers">
       ...
       <option value="1052">آخر</option>
    </select>

    Also falls back to total result count.
    """

    soup = BeautifulSoup(html, 'html.parser')

    select = soup.find(
        'select',
        attrs={'name': 'PageNumbers'}
    )

    page_numbers = []

    if select:
        for option in select.find_all('option'):
            value = option.get('value', '').strip()

            if value.isdigit():
                page_numbers.append(int(value))

    if page_numbers:
        return max(page_numbers)

    total_results = extract_total_results(html)

    if total_results is not None:
        return max(1, math.ceil(total_results / page_size))

    return 1


def get_current_page(html):
    soup = BeautifulSoup(html, 'html.parser')

    field = soup.find(
        'input',
        attrs={'name': 'PageNumber'}
    )

    if field:
        value = field.get('value', '').strip()

        if value.isdigit():
            return int(value)

    select = soup.find(
        'select',
        attrs={'name': 'PageNumbers'}
    )

    if select:
        selected = select.find('option', selected=True)

        if selected:
            value = selected.get('value', '').strip()

            if value.isdigit():
                return int(value)

    return None


# ---------------------------------------------------------
# HTTP
# ---------------------------------------------------------

def make_session():
    s = requests.Session()

    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Linux; Android 10) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0 Safari/537.36'
        ),
        'Accept': (
            'text/html,application/xhtml+xml,application/xml;'
            'q=0.9,image/avif,image/webp,*/*;q=0.8'
        ),
        'Accept-Language': 'fa-IR,fa;q=0.9,en-US;q=0.7,en;q=0.6',
        'Referer': LIST,
    })

    return s


def search_first_page(session, query, page_size=25):
    """
    The actual search form on ara.jri.ac.ir uses POST.

    Main field:
        Title

    We search Title + Abstract + Text so that the official
    site's initial result set is not limited to titles only.
    """

    # First GET establishes cookies/session.
    r = session.get(LIST, timeout=30)
    r.raise_for_status()

    payload = [
        ('Title', query),

        ('IsTitleSearch', 'true'),
        ('IsAbstractSearch', 'true'),
        ('IsTextSearch', 'true'),

        ('_IsOfficial', 'true'),
        ('_IsCivil', 'true'),
        ('_IsPenal', 'true'),

        ('SeachTextType', '3'),

        ('SortColumn', 'Overdate'),
        ('SortDesc', 'True'),

        ('PageNumber', '1'),
        ('PageNumbers', '1'),
        ('PageSize', str(page_size)),
    ]

    r = session.post(
        LIST,
        data=payload,
        timeout=30,
        allow_redirects=True
    )

    r.raise_for_status()

    r.encoding = r.apparent_encoding or 'utf-8'

    return r.text


def fetch_result_page(session, page, page_size=25):
    """
    Pagination form in the site's HTML:

    <form action="/Judge/Index" method="post">

        PageNumbers
        PageNumber
        PageSize

    The server-side session retains the current search.
    """

    payload = {
        'PageNumbers': str(page),
        'PageNumber': str(page),
        'PageSize': str(page_size),
    }

    r = session.post(
        LIST,
        data=payload,
        timeout=30,
        allow_redirects=True
    )

    r.raise_for_status()

    r.encoding = r.apparent_encoding or 'utf-8'

    return r.text


def fetch_vote(url, session):
    r = session.get(url, timeout=30)

    r.raise_for_status()

    r.encoding = r.apparent_encoding or 'utf-8'

    soup = BeautifulSoup(r.text, 'html.parser')

    text = norm(
        soup.get_text(' ', strip=True)
    )

    title_element = (
        soup.find('h1')
        or soup.find('h2')
        or soup.title
    )

    if title_element:
        title = norm(
            title_element.get_text(' ', strip=True)
        )
    else:
        title = 'رأی قضایی'

    return {
        'url': url,
        'title': title,
        'text': text,
        'source': SOURCE_NAME
    }


# ---------------------------------------------------------
# Worker
# ---------------------------------------------------------

def worker(jid, q, max_pages):
    j = JOBS[jid]

    session = make_session()

    PAGE_SIZE = 25

    try:
        j['message'] = 'در حال ارسال جست‌وجو به سامانه رسمی...'

        # -------------------------------------------------
        # First search
        # -------------------------------------------------

        first_html = search_first_page(
            session,
            q,
            PAGE_SIZE
        )

        first_links = get_vote_links(first_html)

        if not first_links:
            j['status'] = 'done'
            j['total_pages'] = 0
            j['progress'] = 100
            j['message'] = 'هیچ رأیی برای این جست‌وجو پیدا نشد.'
            return

        site_total_pages = extract_total_pages(
            first_html,
            PAGE_SIZE
        )

        total_results = extract_total_results(
            first_html
        )

        pages_to_scan = min(
            site_total_pages,
            max_pages
        )

        j['total_pages'] = pages_to_scan
        j['site_total_pages'] = site_total_pages

        if total_results is not None:
            j['total_results'] = total_results

        seen = set()

        previous_page_links = None

        # -------------------------------------------------
        # Pages
        # -------------------------------------------------

        for page in range(1, pages_to_scan + 1):

            if j['cancel']:
                break

            j['current_page'] = page

            j['message'] = (
                f'در حال دریافت صفحه {page} '
                f'از {pages_to_scan}'
            )

            # Page 1 already downloaded during search.
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
                    j['status'] = 'error'
                    j['message'] = (
                        f'خطا در دریافت صفحه {page}: {e}'
                    )
                    return

                links = get_vote_links(html)

            # -------------------------------------------------
            # Validate page
            # -------------------------------------------------

            if not links:
                j['status'] = 'error'
                j['message'] = (
                    f'صفحه {page} دریافت شد اما '
                    'هیچ لینک رأیی در آن پیدا نشد.'
                )
                return

            current_set = set(links)

            if (
                page > 1
                and previous_page_links is not None
                and current_set == previous_page_links
            ):
                j['status'] = 'error'
                j['message'] = (
                    f'صفحه {page} همان نتایج صفحه قبلی را '
                    'برگرداند؛ صفحه‌بندی توسط سامانه پذیرفته نشد.'
                )
                return

            previous_page_links = current_set

            # -------------------------------------------------
            # Votes
            # -------------------------------------------------

            new_links = []

            for u in links:
                if u not in seen:
                    seen.add(u)
                    new_links.append(u)

            for u in new_links:

                if j['cancel']:
                    break

                j['checked'] += 1

                try:
                    vote = fetch_vote(
                        u,
                        session
                    )

                    if matches(vote['text'], q):
                        j['results'].append(vote)
                        j['found'] = len(j['results'])

                except Exception:
                    j['failed_items'] += 1

            j['completed_pages'] = page

            j['progress'] = min(
                99,
                round(
                    page / pages_to_scan * 100,
                    1
                )
            )

            j['message'] = (
                f'صفحه {page} از {pages_to_scan} بررسی شد.'
            )

            time.sleep(0.15)

        # -------------------------------------------------
        # Finish
        # -------------------------------------------------

        if j['cancel']:

            j['status'] = 'cancelled'

            j['message'] = (
                'جست‌وجوی این منبع به درخواست کاربر متوقف شد.'
            )

        elif j['status'] == 'running':

            j['status'] = 'done'

            j['progress'] = 100

            j['message'] = (
                'جست‌وجوی این منبع تکمیل شد.'
            )

    except Exception as e:

        j['status'] = 'error'

        j['message'] = (
            'خطا در ارتباط با سامانه رسمی: '
            + str(e)
        )


# ---------------------------------------------------------
# Word
# ---------------------------------------------------------

def rtl(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    pPr = p._p.get_or_add_pPr()

    pPr.append(
        OxmlElement('w:bidi')
    )

    for run in p.runs:

        rPr = run._r.get_or_add_rPr()

        x = OxmlElement('w:rtl')

        x.set(
            qn('w:val'),
            '1'
        )

        rPr.append(x)


def make_doc(jid):
    j = JOBS[jid]

    doc = Document()

    rtl(
        doc.add_heading(
            'آرای قضایی یافت‌شده',
            0
        )
    )

    rtl(
        doc.add_paragraph(
            f"منبع: {j['source_name']}"
        )
    )

    rtl(
        doc.add_paragraph(
            f"عبارت جست‌وجو: {j['query']} | "
            f"تعداد نتایج: {len(j['results'])}"
        )
    )

    if j['status'] == 'cancelled':

        rtl(
            doc.add_paragraph(
                'توجه: جست‌وجو پیش از تکمیل '
                'توسط کاربر متوقف شده است.'
            )
        )

    for i, vote in enumerate(
        j['results'],
        1
    ):

        rtl(
            doc.add_heading(
                f"{i}. {vote['title']}",
                1
            )
        )

        rtl(
            doc.add_paragraph(
                vote['text']
            )
        )

        rtl(
            doc.add_paragraph(
                'منبع رسمی: ' + vote['url']
            )
        )

        doc.add_page_break()

    path = f'/tmp/{jid}.docx'

    doc.save(path)

    return path


# ---------------------------------------------------------
# Flask routes
# ---------------------------------------------------------

@app.get('/')
def home():
    return render_template('index.html')


@app.post('/api/search')
def start():

    d = request.get_json(force=True)

    q = (d.get('query') or '').strip()

    try:
        requested_max = int(
            d.get('max_pages', 1100)
        )
    except (TypeError, ValueError):
        requested_max = 1100

    max_pages = min(
        max(requested_max, 1),
        1100
    )

    if not q:
        return jsonify(
            error='عبارت جست‌وجو الزامی است'
        ), 400

    jid = str(uuid.uuid4())

    JOBS[jid] = {
        'job_id': jid,
        'source_id': SOURCE_ID,
        'source_name': SOURCE_NAME,
        'query': q,

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
        args=(jid, q, max_pages),
        daemon=True
    ).start()

    return jsonify(
        job_id=jid,
        source_id=SOURCE_ID
    )


@app.get('/api/status/<jid>')
def status(jid):

    j = JOBS.get(jid)

    if not j:
        return jsonify(
            error='یافت نشد'
        ), 404

    keys = (
        'job_id',
        'source_id',
        'source_name',
        'status',
        'checked',
        'found',
        'failed_items',
        'current_page',
        'completed_pages',
        'total_pages',
        'progress',
        'message'
    )

    return jsonify(
        **{
            k: j[k]
            for k in keys
        }
    )


@app.post('/api/cancel/<jid>')
def cancel(jid):

    j = JOBS.get(jid)

    if not j:
        return jsonify(
            error='یافت نشد'
        ), 404

    j['cancel'] = True

    return jsonify(ok=True)


@app.get('/api/download/<jid>')
def download(jid):

    if jid not in JOBS:
        return 'Not found', 404

    return send_file(
        make_doc(jid),
        as_attachment=True,
        download_name='national-judicial-decisions.docx'
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
