from flask import Flask, render_template, request, jsonify, send_file
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import requests, re, threading, uuid, os, time
from urllib.parse import urljoin

app=Flask(__name__)
BASE='https://ara.jri.ac.ir'; LIST=BASE+'/Judge/Index'; JOBS={}
SOURCE_ID='national_judgments'; SOURCE_NAME='سامانه ملی آرای قضایی پژوهشگاه قوه قضاییه'

def norm(s):
    if not s:return ''
    s=s.replace('ي','ی').replace('ك','ک').replace('\u200c',' ')
    return re.sub(r'\s+',' ',s).strip()

def parse_query(q):
    parts=re.split(r'\s+(AND|NOT)\s+',q.strip(),flags=re.I); inc=[]; exc=[]; mode='AND'
    for p in parts:
        if p.upper() in ('AND','NOT'): mode=p.upper(); continue
        p=p.strip().strip('"“”')
        if p: (exc if mode=='NOT' else inc).append(p)
        mode='AND'
    return inc,exc

def matches(text,q):
    text=norm(text); inc,exc=parse_query(q)
    return bool(inc) and all(norm(x) in text for x in inc) and not any(norm(x) in text for x in exc)

def fetch_vote(url,s):
    r=s.get(url,timeout=25); r.raise_for_status(); r.encoding=r.apparent_encoding or 'utf-8'
    soup=BeautifulSoup(r.text,'html.parser'); text=norm(soup.get_text(' ',strip=True))
    title=(soup.find('h1') or soup.find('h2') or soup.title)
    return {'url':url,'title':norm(title.get_text(' ',strip=True)) if title else 'رأی قضایی','text':text,'source':SOURCE_NAME}

def worker(jid,q,max_pages):
    j=JOBS[jid]; s=requests.Session(); s.headers['User-Agent']='Mozilla/5.0 JudicialResearchTool/2.0'
    try:
        seen=set(); j['total_pages']=max_pages
        for page in range(1,max_pages+1):
            if j['cancel']: break
            j['current_page']=page; j['message']=f'در حال دریافت صفحه {page} از {max_pages}'
            try:
                r=s.get(LIST,params={'search':q,'page':page},timeout=25); r.raise_for_status(); r.encoding=r.apparent_encoding or 'utf-8'
            except Exception as e:
                j['status']='error'; j['message']=f'خطا در دریافت صفحه {page}: {e}'; return
            soup=BeautifulSoup(r.text,'html.parser'); links=[]
            for a in soup.find_all('a',href=True):
                if '/Judge/Text/' in a['href']:
                    u=urljoin(BASE,a['href'].split('?')[0])
                    if u not in seen: seen.add(u); links.append(u)
            if not links:
                j['total_pages']=max(page-1,1); j['progress']=100; j['message']='صفحه دیگری برای بررسی دریافت نشد.'; break
            for u in links:
                if j['cancel']: break
                j['checked']+=1
                try:
                    v=fetch_vote(u,s)
                    if matches(v['text'],q): j['results'].append(v); j['found']=len(j['results'])
                except Exception: j['failed_items']+=1
            j['completed_pages']=page
            j['progress']=min(99,round(page/max_pages*100,1))
            j['message']=f'صفحه {page} بررسی شد.'
            time.sleep(.1)
        if j['cancel']:
            j['status']='cancelled'; j['message']='جست‌وجوی این منبع به درخواست کاربر متوقف شد.'
        elif j['status']=='running':
            j['status']='done'; j['progress']=100; j['message']='جست‌وجوی این منبع تکمیل شد.'
    except Exception as e:
        j['status']='error'; j['message']=str(e)

def rtl(p):
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; pPr=p._p.get_or_add_pPr(); pPr.append(OxmlElement('w:bidi'))
    for run in p.runs:
        rPr=run._r.get_or_add_rPr(); x=OxmlElement('w:rtl'); x.set(qn('w:val'),'1'); rPr.append(x)

def make_doc(jid):
    j=JOBS[jid]; doc=Document(); rtl(doc.add_heading('آرای قضایی یافت‌شده',0))
    rtl(doc.add_paragraph(f"منبع: {j['source_name']}")); rtl(doc.add_paragraph(f"عبارت جست‌وجو: {j['query']} | تعداد نتایج: {len(j['results'])}"))
    if j['status']=='cancelled': rtl(doc.add_paragraph('توجه: جست‌وجو پیش از تکمیل توسط کاربر متوقف شده است.'))
    for i,v in enumerate(j['results'],1):
        rtl(doc.add_heading(f"{i}. {v['title']}",1)); rtl(doc.add_paragraph(v['text'])); rtl(doc.add_paragraph('منبع رسمی: '+v['url'])); doc.add_page_break()
    path=f'/tmp/{jid}.docx'; doc.save(path); return path

@app.get('/')
def home(): return render_template('index.html')
@app.post('/api/search')
def start():
    d=request.get_json(force=True); q=(d.get('query') or '').strip(); max_pages=min(max(int(d.get('max_pages',1052)),1),1100)
    if not q:return jsonify(error='عبارت جست‌وجو الزامی است'),400
    jid=str(uuid.uuid4()); JOBS[jid]={'job_id':jid,'source_id':SOURCE_ID,'source_name':SOURCE_NAME,'query':q,'status':'running','cancel':False,'checked':0,'found':0,'failed_items':0,'current_page':0,'completed_pages':0,'total_pages':max_pages,'progress':0,'results':[],'message':'جست‌وجو آغاز شد.'}
    threading.Thread(target=worker,args=(jid,q,max_pages),daemon=True).start(); return jsonify(job_id=jid,source_id=SOURCE_ID)
@app.get('/api/status/<jid>')
def status(jid):
    j=JOBS.get(jid)
    if not j:return jsonify(error='یافت نشد'),404
    return jsonify(**{k:j[k] for k in ('job_id','source_id','source_name','status','checked','found','failed_items','current_page','completed_pages','total_pages','progress','message')})
@app.post('/api/cancel/<jid>')
def cancel(jid):
    j=JOBS.get(jid)
    if not j:return jsonify(error='یافت نشد'),404
    j['cancel']=True; return jsonify(ok=True)
@app.get('/api/download/<jid>')
def download(jid):
    if jid not in JOBS:return 'Not found',404
    return send_file(make_doc(jid),as_attachment=True,download_name='national-judicial-decisions.docx')
if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.getenv('PORT',5000)))
