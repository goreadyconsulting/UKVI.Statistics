"""Publish independent, source-labelled analytics for Home Office catalogue workbooks.
Never infer totals from summary tables. Detailed records retain every published dimension.
"""
import json, hashlib, pathlib, re, sys, tempfile, urllib.request, urllib.parse, zipfile, math
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from bs4 import BeautifulSoup
from openpyxl import load_workbook

OUT=pathlib.Path(sys.argv[1])
CAT=OUT/'catalogue.json'
PREVIOUS=pathlib.Path('public-data')
LIMIT=180*1024*1024
NS={'t':'urn:oasis:names:tc:opendocument:xmlns:table:1.0','o':'urn:oasis:names:tc:opendocument:xmlns:office:1.0','x':'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}
T='{'+NS['t']+'}'; O='{'+NS['o']+'}'
def request(url):
    return urllib.request.urlopen(urllib.request.Request(url,headers={'User-Agent':'UKImmigrationIntelligence/1.0 (public statistics reuse)'}),timeout=90)
def dump(path,obj):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(obj,ensure_ascii=False,separators=(',',':'),allow_nan=False),encoding='utf-8')
def slug(s): return re.sub(r'[^a-z0-9]+','-',s.lower()).strip('-')
def source_id(url): return hashlib.sha256(url.encode()).hexdigest()[:16]
def sheet_id(s): return hashlib.sha256(s.encode()).hexdigest()[:12]
def plain(x):
    if x is None: return ''
    if isinstance(x,(str,bool,int)): return x
    if isinstance(x,float): return x if math.isfinite(x) else ''
    return str(x)
def nonempty(row): return any(v not in ('',None) for v in row)
def trim(row):
    while row and row[-1] in ('',None): row.pop()
    return row
def download(url,path):
    with request(url) as r,path.open('wb') as f:
        while chunk:=r.read(1024*1024):
            f.write(chunk)
            if f.tell()>LIMIT: raise ValueError('Source exceeds download safety bound')
def links_page(url):
    with request(url) as r: soup=BeautifulSoup(r.read(),'html.parser')
    result=[]
    for a in soup.select('a[href]'):
        href=urllib.parse.urljoin(url,a['href'])
        u=urllib.parse.urlparse(href)
        ext=pathlib.PurePosixPath(u.path).suffix.lower().lstrip('.')
        title=a.get_text(' ',strip=True)
        if u.scheme=='https' and u.hostname and u.hostname.endswith('.publishing.service.gov.uk') and ext in ('ods','xlsx') and title:
            result.append({'url':href,'title':title,'format':ext})
    return list({r['url']:r for r in result}.values())
def select_sources(section):
    sid=section['id']
    direct=[l for l in section['links'] if l['format'] in ('xlsx','ods')]
    if sid=='entry-clearance-visas-granted-outside-the-uk':
        return [l for l in direct if l['format']=='ods' and 'entry clearance visas summary' in l['title'].lower()]
    if sid=='sponsored-work-visas-by-occupation-and-industry':
        return [l for l in direct if 'occupation and industry' in l['title'].lower()]
    if sid=='sponsored-study-visas-by-course-level':
        return [l for l in direct if 'course level' in l['title'].lower()]
    if sid=='regional-and-local-authority-data-on-immigration-groups':
        found=links_page('https://www.gov.uk/government/statistical-data-sets/immigration-system-statistics-regional-and-local-authority-data')
        return found[:1]
    if sid=='other-migration-and-borders-performance-data':
        found=links_page('https://www.gov.uk/government/statistical-data-sets/migration-transparency-data')
        # The catalogue lists the latest set first. Keep one current file per publication title.
        seen=set(); result=[]
        for l in found:
            key=re.split(r':\s*(?:January|April|July|October)|,\s*(?:January|April|July|October)|\b20\d\d\b',l['title'],maxsplit=1,flags=re.I)[0].lower()
            if key not in seen: seen.add(key);result.append(l)
        return result[:15]
    if sid=='related-content': return []
    return direct
def ods_sheets(path):
    with zipfile.ZipFile(path) as z:
        info=z.getinfo('content.xml')
        if info.file_size>250*1024*1024: raise ValueError('ODS XML exceeds safety bound')
        root=ET.fromstring(z.read('content.xml'))
        for tab in root.findall('.//o:spreadsheet/t:table',NS):
            name=tab.get(T+'name','Sheet')
            if meta_sheet(name):continue
            rows=[]; clipped=False
            for row in tab.findall('.//t:table-row',NS):
                cells=[]
                for cell in row:
                    if cell.tag not in (T+'table-cell',T+'covered-table-cell'):continue
                    repeat=int(cell.get(T+'number-columns-repeated','1'))
                    typ=cell.get(O+'value-type','')
                    s=' '.join(''.join(p.itertext()).strip() for p in cell.findall('x:p',NS)).strip()
                    v=s
                    if typ in ('float','currency','percentage') and cell.get(O+'value') is not None:
                        try:
                            n=float(cell.get(O+'value'))
                            if math.isfinite(n):
                                if typ=='percentage':v={'n':n*100,'display':s or str(round(n*100,4))+'%','unit':'%'}
                                elif typ=='currency':v={'n':n,'display':s or str(n),'unit':cell.get(O+'currency','currency')}
                                else:v=int(n) if n.is_integer() else n
                        except ValueError: pass
                    if repeat>512 and v!='':clipped=True
                    cells.extend([v]*min(repeat,max(0,512-len(cells))))
                    if len(cells)>=512:break
                trim(cells)
                count=int(row.get(T+'number-rows-repeated','1'))
                if not cells:
                    # Empty spacer rows do not have statistical values.
                    if rows and len(rows)<6000:rows.append([])
                    continue
                if len(rows)+count>6000:clipped=True
                rows.extend([cells]*min(count,max(0,6000-len(rows))))
                if len(rows)>=6000:break
            while rows and not rows[-1]:rows.pop()
            if any(nonempty(r) for r in rows):yield name,rows,clipped
def meta_sheet(name):
    s=name.lower().replace('_',' ').strip()
    return s in ('cover','cover sheet','contents','content','notes','list of fields','definitions','readme','metadata','index','guidance','glossary') or s.startswith('notes ')
def normhead(s):return re.sub(r'\s+',' ',str(s).strip()).lower()
def xlsx_tables(path):
    book=load_workbook(path,read_only=True,data_only=True)
    try:
        for ws in book:
            if meta_sheet(ws.title):continue
            it=ws.iter_rows()
            head=[]
            for i in range(15):
                try:head.append([plain(c.value) for c in next(it)])
                except StopIteration:break
            # Detailed Home Office tables use explicit Year / Quarter headers. Require
            # a complete header and preserve each original cell, no inferred measures.
            idx=next((i for i,row in enumerate(head) if any(normhead(v) in ('year','calendar year','reporting year','quarter','reporting quarter','year ending') for v in row) and sum(bool(v) for v in row)>=3),None)
            if idx is not None and (ws.title.lower().startswith('data') or (ws.max_row or 0)>6000):
                header=trim(head[idx][:])
                if not header or any(not isinstance(v,str) or not v.strip() for v in header) or len(set(header))!=len(header):
                    print('Skip ambiguous detailed header',ws.title,flush=True);continue
                def records():
                    for row in head[idx+1:]:yield row
                    for row in it:yield [plain(c.value) for c in row]
                yield ws.title,'detail',header,records(),False
            else:
                rows=[trim(r[:512]) for r in head];clipped=False
                for row in it:
                    if len(rows)>=6000:clipped=True;break
                    vals=[]
                    for cell in row[:512]:
                        v=plain(cell.value)
                        if isinstance(v,(int,float)) and not isinstance(v,bool) and '%' in (cell.number_format or ''):
                            v={'n':v*100,'display':str(round(v*100,4))+'%','unit':'%'}
                        vals.append(v)
                    rows.append(trim(vals))
                while rows and not rows[-1]:rows.pop()
                if rows:yield ws.title,'matrix',None,rows,clipped
    finally:book.close()
def numeric(v):return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)
def year_key(v):
    m=re.match(r'^\s*((?:19|20)\d{2})(?:\b|Q)',str(v))
    return m.group(1) if m else 'Other'
def detail(pathroot,name,header,records):
    # Numeric candidates are derived from actual cell types, not column names;
    # date and category codes are still selectable as dimensions.
    dates=[i for i,x in enumerate(header) if normhead(x) in ('year','calendar year','reporting year','quarter','reporting quarter','year ending')]
    primary=next((i for i in dates if 'year' in normhead(header[i])),dates[0])
    counts=[0]*len(header);nonblank=[0]*len(header);parts=defaultdict(list);total=0
    for row in records:
        row=(row+['']*len(header))[:len(header)]
        if not nonempty(row):continue
        if not str(row[primary]).strip():continue
        key=year_key(row[primary])
        if key=='Other' and not re.search(r'\b(?:19|20)\d{2}\b',str(row[primary])):continue
        for i,v in enumerate(row):
            if v not in ('',None):nonblank[i]+=1
            if numeric(v):counts[i]+=1
        parts[key].append(row);total+=1
        if total>1500000:raise ValueError('Detailed sheet exceeds safety row bound')
    if not total:return None
    candidates=[i for i in range(len(header)) if i not in dates and nonblank[i] and counts[i]/nonblank[i]>.75 and not re.search(r'\b(code|id|rank|age|year|month|quarter)\b',normhead(header[i]))]
    if not candidates:return None
    # Data tables are partitioned by year to keep downloads small.
    partition=[]
    for key,rs in sorted(parts.items()):
        file='part-'+slug(key)+'.json';dump(pathroot/file,rs)
        partition.append({'key':key,'file':file,'rows':len(rs)})
    return {'id':sheet_id(name),'name':name,'kind':'detail','columns':header,'measures':candidates,'dates':dates,'partitions':partition,'rows':total}
def matrix(pathroot,name,rows,clipped):
    # Never total or combine published cells; preserve suppression and row/column labels.
    numeric_count=sum(numeric(v) or (isinstance(v,dict) and 'n' in v) for row in rows for v in row)
    if numeric_count<2:return None
    dump(pathroot/'matrix.json',rows)
    return {'id':sheet_id(name),'name':name,'kind':'matrix','file':'matrix.json','rows':len(rows),'columns':max(map(len,rows),default=0),'clipped':clipped}
def process(src):
    sid=source_id(src['url']);root=OUT/'reports'/sid
    old=PREVIOUS/'reports'/sid/'manifest.json'
    # Files are immutable versioned GOV.UK attachment URLs. Reuse already parsed
    # files only if the manifest has the same format version.
    if old.exists():
        prev=json.loads(old.read_text())
        if prev.get('version')==1 and prev.get('tables'):
            import shutil
            shutil.copytree(old.parent,root,dirs_exist_ok=True)
            return prev
    with tempfile.TemporaryDirectory() as tmp:
        path=pathlib.Path(tmp)/('source.'+src['format']);download(src['url'],path)
        sha=hashlib.sha256(path.read_bytes()).hexdigest()
        tables=[]
        if src['format']=='ods':
            for name,rows,clipped in ods_sheets(path):
                t=matrix(root/sheet_id(name),name,rows,clipped)
                if t:tables.append(t)
        else:
            for name,kind,header,rows,clipped in xlsx_tables(path):
                t=detail(root/sheet_id(name),name,header,rows) if kind=='detail' else matrix(root/sheet_id(name),name,rows,clipped)
                if t:tables.append(t)
        if not tables:raise ValueError('No chartable published tables')
        # Put detailed sheets first, keeping all chartable summary tables available.
        tables.sort(key=lambda t:(t['kind']!='detail',t['name']))
        manifest={'version':1,'id':sid,'source':src,'sha256':sha,'tables':tables}
        dump(root/'manifest.json',manifest)
        print('DOMAIN READY',src['title'],len(tables),'tables',[(t['name'],t['kind'],t['rows']) for t in tables][:8],flush=True)
        return manifest
def main():
    cat=json.loads(CAT.read_text());sources={};selected={};errors={}
    for section in cat['sections']:
        try:
            chosen=select_sources(section)
            selected[section['id']]=[source_id(l['url']) for l in chosen]
            for src in chosen:sources[source_id(src['url'])]=src
        except Exception as e:
            selected[section['id']]=[];errors[section['id']]=str(e);print('DOMAIN DISCOVERY ERROR',section['id'],str(e),flush=True)
    results={}
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures={pool.submit(process,src):sid for sid,src in sources.items()}
        for future in as_completed(futures):
            sid=futures[future]
            try:
                m=future.result();results[sid]={'id':sid,'source':m['source'],'tables':len(m['tables']),'sha256':m['sha256'],'status':'ready'}
            except Exception as e:
                results[sid]={'id':sid,'source':sources[sid],'status':'error','error':str(e)[:300]};print('DOMAIN ERROR',sources[sid]['title'],str(e),flush=True)
    cat['domain_reports']={'version':1,'sections':selected,'sources':results,'discovery_errors':errors}
    dump(CAT,cat)
    print('DOMAIN SUMMARY',len(sources),'sources',sum(x['status']=='ready' for x in results.values()),'ready',sum(x['status']=='error' for x in results.values()),'errors',flush=True)
if __name__=='__main__':main()
