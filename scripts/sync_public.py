"""Fetch public Home Office catalogue and build a validated, year-partitioned entry clearance snapshot."""
import datetime as dt
import hashlib
import json
import pathlib
import re
import sys
import tempfile
import urllib.parse
import urllib.request
from collections import defaultdict
from bs4 import BeautifulSoup
from openpyxl import load_workbook

URL = 'https://www.gov.uk/government/statistical-data-sets/immigration-system-statistics-data-tables'
OUT = pathlib.Path(sys.argv[1] if len(sys.argv)>1 else 'public-data')
GROUPS = [["related-content","Related content"],["passenger-arrivals","Passenger arrivals"],["electronic-travel-authorisation","Electronic travel authorisation"],["entry-clearance-visas-granted-outside-the-uk","Entry clearance visas"],["sponsored-work-visas-by-occupation-and-industry","Sponsored work visas by occupation and industry"],["sponsored-study-visas-by-course-level","Sponsored study visas by course level"],["safe-and-legal-humanitarian-routes-to-the-uk","Safe and legal humanitarian routes"],["organised-immigration-crime","Organised immigration crime"],["illegal-entry-routes","Illegal entry routes"],["asylum","Asylum"],["regional-and-local-authority-data-on-immigration-groups","Regional and local authority data"],["extensions","Extensions"],["settlement","Settlement"],["eu-settlement-scheme","EU Settlement Scheme"],["citizenship","Citizenship"],["immigration-detention","Immigration detention"],["returns","Returns"],["transfers-to-the-uk-under-section-67-of-the-immigration-act-2016","Section 67 transfers"],["discontinued-releases","Discontinued releases"],["other-migration-and-borders-performance-data","Migration and borders performance"]]
def request(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent':'UKImmigrationIntelligence/1.0 (public statistics reuse)'}), timeout=120)
def write(name, data):
    target = OUT / name
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_suffix('.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, separators=(',',':')), encoding='utf-8')
    temp.replace(target)
def catalogue():
    with request(URL) as res:
        soup = BeautifulSoup(res.read(), 'html.parser')
    sections = []
    for anchor, label in GROUPS:
        heading = soup.find(id=anchor)
        if heading is None:
            # Publication headings can change; preserve unavailable sections as source links.
            terms = [anchor.replace('-',' '), label.lower()]
            heading = next((h for h in soup.find_all('h2') if any(t in h.get_text(' ',strip=True).lower() for t in terms)), None)
        if heading is None:
            sections.append({'id':anchor,'title':label,'url':URL+'#'+anchor,'links':[], 'notes':['This heading was not found in the latest catalogue; see the official source for the current classification.']})
            continue
        if heading.name not in ('h2','h3'):
            heading = heading.find_parent(['h2','h3']) or heading
        links = {}
        notes = []
        for node in heading.next_siblings:
            if getattr(node,'name',None) == 'h2': break
            if not getattr(node,'name',None): continue
            if node.name in ('p','blockquote','div'):
                note = node.get_text(' ',strip=True)
                if note and any(x in note.lower() for x in ('discontinued','revised','not the latest release','supersede','provisional')): notes.append(note)
            for a in node.find_all('a',href=True):
                href = urllib.parse.urljoin(URL,a['href'])
                title = a.get_text(' ',strip=True)
                host = urllib.parse.urlparse(href).hostname
                if href.startswith('https://') and host and (host=='www.gov.uk' or host.endswith('.publishing.service.gov.uk')) and title and not href.startswith(URL+'#'):
                    links[href]={'title':title,'url':href,'format':pathlib.PurePosixPath(urllib.parse.urlparse(href).path).suffix.lower().lstrip('.') or 'page'}
        sections.append({'id':anchor,'title':label,'url':URL+'#'+anchor,'links':list(links.values()),'notes':notes})
    return sections
def clean(value):
    return '' if value is None else str(value).strip()
HEAD = ['Year','Quarter','Nationality','Region','Visa type group','Visa type','Visa type subgroup','Applicant type']
def parse(path):
    book=load_workbook(path,read_only=True,data_only=True)
    rows=defaultdict(dict)
    counts={}
    for name,tail in [('Data_Vis_D01',['Applications']),('Data_Vis_D02',['Case outcome','Decisions'])]:
        sheet=book[name]
        it=sheet.iter_rows(values_only=True)
        for _ in range(3): next(it)
        header=list(next(it))
        expected=HEAD+tail
        if [clean(x) for x in header[:len(expected)]]!=expected:
            raise ValueError('Header mismatch: '+name+' '+str(header))
        count=0
        for row in it:
            if not any(v is not None for v in row): continue
            year=int(row[0]); quarter=clean(row[1])
            if quarter not in [str(year)+' Q'+str(i) for i in range(1,5)]: raise ValueError('Invalid quarter '+quarter)
            dimensions=tuple(clean(x) for x in row[2:8])
            if not all(dimensions): raise ValueError('Blank dimension '+name)
            value=row[len(expected)-1]
            if not isinstance(value,(int,float)) or value<0 or int(value)!=value: raise ValueError('Invalid count '+repr(value))
            key=(quarter[-2:],)+dimensions
            dest=rows[year].setdefault(key,{'applications':0,'outcomes':{}})
            if len(tail)==1: dest['applications']+=int(value)
            else:
                outcome=clean(row[8])
                if not outcome: raise ValueError('Blank case outcome')
                dest['outcomes'][outcome]=dest['outcomes'].get(outcome,0)+int(value)
            count+=1
        counts[name]=count
    book.close()
    years=sorted(rows)
    if not years or years[0]!=2005 or any(n<10000 for n in counts.values()): raise ValueError('Unexpected coverage')
    for year in years:
        data=[[*key,v['applications'],v['outcomes']] for key,v in sorted(rows[year].items())]
        write('visas/'+str(year)+'.json',data)
    return {'years':years,'latest_quarters':sorted({k[0] for k in rows[years[-1]]}),'source_rows':counts,'fields':['Quarter','Nationality','Region','Visa type group','Visa type','Visa type subgroup','Applicant type','Applications','Outcomes']}
def main():
    sections=catalogue()
    visa=next(x for x in sections if x['id']=='entry-clearance-visas-granted-outside-the-uk')
    candidates=[x for x in visa['links'] if x['format']=='xlsx' and 'entry clearance' in x['title'].lower() and 'application' in x['title'].lower() and 'outcome' in x['title'].lower()]
    if len(candidates)!=1: raise ValueError('Expected one detailed entry clearance workbook: '+str(candidates))
    src=candidates[0]
    with tempfile.TemporaryDirectory() as tmp:
        path=pathlib.Path(tmp)/'source.xlsx'
        digest=hashlib.sha256()
        with request(src['url']) as response, path.open('wb') as out:
            while chunk:=response.read(1024*1024):
                digest.update(chunk)
                out.write(chunk)
                if out.tell()>200*1024*1024: raise ValueError('Workbook exceeds safety bound')
        meta=parse(path)
        meta.update({'source':src,'sha256':digest.hexdigest()})
    write('catalogue.json',{'source':URL,'checked_at':dt.datetime.now(dt.timezone.utc).isoformat(),'sections':sections,'entry_clearance':meta})
    print('Validated',meta['source_rows'],'years',meta['years'],'latest quarters',meta['latest_quarters'])
if __name__=='__main__': main()
