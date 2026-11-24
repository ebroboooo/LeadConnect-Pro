"""Local contact database and import helpers."""
import csv, io, json, re, sqlite3, zipfile, unicodedata
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

STATUSES = ('New', 'Contacted', 'Replied', 'Interested', 'Demo booked', 'Won', 'Declined', 'Do not contact')
BLOCKED = {'Declined', 'Do not contact', 'Won'}
MAX_FILE = 10_000_000
MAX_ROWS = 10000

def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')

def ascii_digits(value):
    return ''.join(str(unicodedata.decimal(c)) if c.isdecimal() else c for c in str(value or '').strip())

def normalize(value):
    raw = ascii_digits(value)
    if re.fullmatch(r'\d+\.0+', raw): raw = raw.split('.')[0]
    if not re.fullmatch(r'[+\d\s().-]+', raw): return '', 'Invalid'
    digits = re.sub(r'\D', '', raw)
    if digits.startswith('0020'): digits = '20' + digits[4:]
    if re.fullmatch(r'01[0125]\d{8}', digits): return '20'+digits[1:], 'Egypt mobile'
    if re.fullmatch(r'1[0125]\d{8}', digits): return '20'+digits, 'Egypt mobile'
    if re.fullmatch(r'201[0125]\d{8}', digits): return digits, 'Egypt mobile'
    if re.fullmatch(r'0[2-9]\d{7,9}', digits): return '20'+digits[1:], 'Landline / other Egypt'
    if re.fullmatch(r'20[2-9]\d{7,9}', digits): return digits, 'Landline / other Egypt'
    if raw.startswith('+') and 8 <= len(digits) <= 15: return digits, 'International / review'
    return '', 'Invalid'

def read_xlsx(path):
    ns = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as z:
        if sum(x.file_size for x in z.infolist()) > 60_000_000: raise ValueError('Workbook is too large')
        def read_xml(name):
            data=z.read(name)
            if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper(): raise ValueError('Unsafe XML')
            return ET.fromstring(data)
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in read_xml('xl/sharedStrings.xml').findall('s:si',ns):
                shared.append(''.join(t.text or '' for t in si.findall('.//s:t',ns)))
        sheets=sorted(n for n in z.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n))
        if not sheets: raise ValueError('No worksheet found')
        rows=[]
        for row in read_xml(sheets[0]).findall('.//s:sheetData/s:row',ns)[:MAX_ROWS+1]:
            vals={}
            for c in row.findall('s:c',ns):
                m=re.match(r'([A-Z]+)',c.get('r',''))
                if not m: continue
                idx=0
                for letter in m.group(1): idx=idx*26+ord(letter)-64
                if idx>200: continue
                typ=c.get('t'); v=c.findtext('s:v',default='',namespaces=ns)
                if c.find('s:f',ns) is not None: v=''
                elif typ=='s': v=shared[int(v)] if v else ''
                elif typ=='inlineStr': v=''.join(t.text or '' for t in c.findall('.//s:t',ns))
                vals[idx-1]=v
            if vals: rows.append([vals.get(i,'') for i in range(max(vals)+1)])
    if not rows: return []
    return [dict(zip(rows[0],row)) for row in rows[1:]]

def read_records(path):
    path=Path(path)
    if path.stat().st_size>MAX_FILE: raise ValueError('File exceeds 10 MB')
    if path.suffix.lower()=='.xlsx': return read_xlsx(path)
    if path.suffix.lower()=='.json':
        obj=json.loads(path.read_text(encoding='utf-8-sig'))
        if isinstance(obj,dict): obj=obj.get('leads',[])
        if not isinstance(obj,list) or not all(isinstance(r,dict) for r in obj): raise ValueError('Expected a list of prospect records')
        return obj[:MAX_ROWS]
    if path.suffix.lower() not in ('.csv','.txt','.tsv'): raise ValueError('Choose CSV, TSV, TXT, XLSX or JSON (not legacy XLS)')
    data=path.read_bytes()
    try: content=data.decode('utf-8-sig')
    except UnicodeDecodeError: content=data.decode('cp1256')
    try: dialect=csv.Sniffer().sniff(content[:4096], delimiters=',;\t')
    except csv.Error: dialect=csv.excel
    return list(csv.DictReader(io.StringIO(content), dialect=dialect))[:MAX_ROWS]

def field(row,*names):
    keys={str(k).strip().lower().replace(' ','').replace('_',''):v for k,v in row.items()}
    for name in names:
        value=keys.get(name)
        if value is not None and str(value).strip(): return str(value).strip()
    return ''

class Store:
    def __init__(self,path):
        self.db=sqlite3.connect(path); self.db.row_factory=sqlite3.Row
        self.db.executescript("""CREATE TABLE IF NOT EXISTS contacts(id INTEGER PRIMARY KEY, key TEXT UNIQUE, name TEXT, raw TEXT, phone TEXT, kind TEXT, area TEXT, source TEXT, status TEXT DEFAULT 'New', verification TEXT DEFAULT 'Unknown', consent TEXT DEFAULT '', notes TEXT DEFAULT '', imported TEXT, sent TEXT DEFAULT '', replied TEXT DEFAULT ''); CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY, contact_id INTEGER NOT NULL, kind TEXT, happened TEXT, detail TEXT);""")
    def import_rows(self, rows):
        added=dupes=invalid=0
        with self.db:
            for r in rows:
                raw=field(r,'phone','phonenumber','mobile','number','telephone','tel','whatsapp','رقم','الهاتف','رقمالهاتف','رقمالموبايل')
                if not raw: continue
                name=field(r,'name','business','businessname','company','اسم','الاسم') or 'Unnamed business'
                area=field(r,'area','city','address','location','المدينة','العنوان')
                source=field(r,'source','mapsurl','website')
                for token in re.split(r'[;|\n]+',raw):
                    token=token.strip(); phone,kind=normalize(token)
                    if kind=='Invalid': invalid+=1
                    key=phone or 'invalid:'+token.casefold()
                    cur=self.db.execute('INSERT OR IGNORE INTO contacts(key,name,raw,phone,kind,area,source,imported) VALUES(?,?,?,?,?,?,?,?)',(key,name,token,phone,kind,area,source,now()))
                    added+=cur.rowcount; dupes+=1-cur.rowcount
        return added,dupes,invalid
    def all(self): return self.db.execute('SELECT * FROM contacts ORDER BY id').fetchall()
    def get(self,id): return self.db.execute('SELECT * FROM contacts WHERE id=?',(id,)).fetchone()
    def update(self,id,status,verification,consent,notes):
        if status not in STATUSES or verification not in ('Unknown','Confirmed manually','Unavailable'): raise ValueError('Invalid status')
        with self.db:
            self.db.execute('UPDATE contacts SET status=?, verification=?, consent=?, notes=? WHERE id=?',(status,verification,consent.strip(),notes.strip(),id))
    def event(self,id,kind,detail):
        row=self.get(id); t=now()
        if not row: raise ValueError('Contact not found')
        with self.db:
            if kind=='Sent':
                if not eligible(row): raise ValueError('Contact must have recorded permission and a valid eligible number')
                self.db.execute("UPDATE contacts SET sent=?,status=CASE WHEN status='New' THEN 'Contacted' ELSE status END WHERE id=?",(t,id))
            elif kind=='Reply':
                if not row['sent']: raise ValueError('No recorded outbound message')
                self.db.execute("UPDATE contacts SET replied=?, status=CASE WHEN status='Contacted' THEN 'Replied' ELSE status END WHERE id=?",(t,id))
            else: raise ValueError('Unknown event')
            self.db.execute('INSERT INTO events(contact_id,kind,happened,detail) VALUES(?,?,?,?)',(id,kind,t,detail[:3000]))
    def match(self,raw):
        phone,_=normalize(raw)
        if not phone: return None
        return self.db.execute("SELECT * FROM contacts WHERE phone=? AND sent<>''",(phone,)).fetchone()

def eligible(r):
    return bool(r and r['kind']=='Egypt mobile' and r['status'] not in BLOCKED and r['consent'].strip() and r['verification']!='Unavailable')

def safe_csv(v):
    s=str(v or '')
    return "'"+s if re.match(r'^[\s]*[=+\-@]',s) else s
