#!/usr/bin/env python3
"""Build the Prospecți Worker: parse the CRM xlsx into data.json and bundle it with index.html into worker.js.

Usage: python3 build.py path/to/export.xlsx
Outputs data.json and worker.js next to this script. Neither is committed (see .gitignore);
they are regenerated from the source spreadsheet, which stays out of git because it holds contact PII.
"""
import openpyxl, json, re, sys, os, base64

def s(v):
    if v is None: return ""
    if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
    return str(v).strip()

EMAIL = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
PHONE = re.compile(r'0\d[\d/.\-\s]{6,}\d')
URL = re.compile(r'(?:https?://|www\.)?[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+\.[A-Za-z]{2,}', re.I)
FREE = {'gmail.com', 'yahoo.com', 'yahoo.ro', 'hotmail.com', 'hotmail.ro', 'outlook.com', 'outlook.ro',
        'icloud.com', 'msn.com', 'live.com', 'aol.com', 'protonmail.com', 'mail.ru', 'yandex.ru'}

# CRM note key (parentheticals + year stripped, lowercased) -> (bucket, clean English label)
KNOWN = {
    'potential onboarding digital': ('qual', 'Digital onboarding'),
    'eligibil ovd digital': ('qual', 'OVD Digital'),
    'eligibil credit smart business': ('qual', 'Credit Smart Business'),
    'eligibil cc digital': ('qual', 'CC Digital'),
    'interactiuni o.crm in ultimele 3 luni': ('qual', 'CRM contact, last 3 months'),
    'cifra de afaceri': ('rev', None), 'nr. salariati': ('emp', None), 'nr salariati': ('emp', None),
    'localitate': ('skip', None), 'judet': ('skip', None),
    'persoana de contact': ('person', None), 'mobil': ('phone', None), 'telefon': ('phone', None),
    'email': ('email', None), 'mail': ('email', None), 'web': ('web', None), 'website': ('web', None),
}
QUAL_ORDER = ['Digital onboarding', 'OVD Digital', 'Credit Smart Business', 'CC Digital', 'CRM contact, last 3 months']

def phones_in(text):
    out = []
    for pm in PHONE.findall(text or ''):
        if 9 <= len(re.sub(r'\D', '', pm)) <= 12: out.append(re.sub(r'\s+', ' ', pm.strip().strip(',;')))
    return out

def website_from(web_hint, emails):
    site = web_hint or next((e.split('@')[-1] for e in emails if e.split('@')[-1].lower() not in FREE), '')
    site = re.sub(r'^https?://', '', site, flags=re.I).strip().strip('.,;/ ')
    return re.sub(r'^www\.', '', site, flags=re.I).lower()

def parse_notes(t):
    """Turn the free-text CRM note into fully structured, labelled fields. Nothing is discarded:
    anything unrecognised lands in `tags` (free lines) or `other` (unknown key:value)."""
    o = {'phones': [], 'emails': [], 'website': '', 'person': '',
         'prio': [], 'qual': [], 'rev': '', 'revYear': '', 'employees': '', 'empYear': '', 'tags': [], 'other': []}
    if not t: return o
    phones, emails, web_hint = [], [], ''
    qual = {}
    for line in t.splitlines():
        raw = line.strip()
        if not raw: continue
        low = raw.lower().lstrip('- ').strip()
        if low.startswith(('recomandari abordare', 'informatii companie', 'date de contact suplimentare')): continue
        m = re.match(r'^-?\s*([^:]{2,70}?):\s*(.*)$', raw)
        if not m:
            o['tags'].append(raw.lstrip('- ').strip()); continue
        key, val = m.group(1).strip(), m.group(2).strip().rstrip(',').strip()
        klow = key.lower()
        if klow.startswith('prio'):
            o['prio'] += [x.strip() for x in re.split(r';', re.sub(r'Prio\s*\d+\s*:', '', 'Prio 0:' + val)) if x.strip()]
            continue
        if klow.startswith('date de contact publice'):
            phones += phones_in(val); continue
        yr = re.search(r'\((\d{4})\)', key)
        norm = re.sub(r'\([^)]*\)', '', klow).strip().rstrip(':').strip()
        bucket = KNOWN.get(norm)
        if not bucket:
            if val: o['other'].append([key, val])   # unknown but shown, never dropped
            continue
        b, label = bucket
        if b == 'phone':   phones += phones_in(val)
        elif b == 'email': emails += EMAIL.findall(val)
        elif b == 'person': o['person'] = o['person'] or val
        elif b == 'web':   web_hint = web_hint or val
        elif b == 'rev':   o['rev'], o['revYear'] = (o['rev'] or val), (o['revYear'] or (yr.group(1) if yr else ''))
        elif b == 'emp':   o['employees'], o['empYear'] = (o['employees'] or val), (o['empYear'] or (yr.group(1) if yr else ''))
        elif b == 'qual':
            da = 'DA' if val.upper().startswith('DA') else 'NU' if val.upper().startswith('NU') else val
            lim = re.search(r'limit[ae]:?\s*(.+)$', val, re.I)
            qual[label] = {'label': label, 'val': da, 'limit': lim.group(1).strip() if lim else ''}
    o['prio'] = [p for p in dict.fromkeys(o['prio']) if p]
    o['qual'] = [qual[l] for l in QUAL_ORDER if l in qual]
    o['phones'] = list(dict.fromkeys(phones))
    o['emails'] = list(dict.fromkeys(e for e in emails if e))
    o['website'] = website_from(web_hint, o['emails'])
    return o

def build_data(xlsx):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb['BD CAMPANIE']
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c).strip() for c in rows[0]]
    I = {h: i for i, h in enumerate(hdr)}
    def gi(p):
        for h, i in I.items():
            if h.lower().startswith(p.lower()): return i
    C = dict(comp=I['Companie'], cui=I['CUI'], turn=gi('Cifr'), city=gi('Oraș'), county=gi('Județ'),
             due=gi('Data scaden'), prio=gi('Descriere'), cic=gi('CIC'), lastact=gi('Dată a celei'))
    recs = []
    for r in rows[1:]:
        try: turn = float(r[C['turn']])
        except (TypeError, ValueError): turn = 0.0
        p = parse_notes(s(r[C['prio']]))
        recs.append({"company": s(r[C['comp']]), "cui": s(r[C['cui']]), "turnover": turn,
                     "city": s(r[C['city']]).title() if s(r[C['city']]) else "",
                     "county": s(r[C['county']]).title() if s(r[C['county']]) else "",
                     "due": s(r[C['due']]), "cic": s(r[C['cic']]), "lastActivity": s(r[C['lastact']]),
                     "phones": p['phones'], "emails": p['emails'], "website": p['website'], "person": p['person'],
                     "prio": p['prio'], "qual": p['qual'], "rev": p['rev'], "revYear": p['revYear'],
                     "employees": p['employees'], "empYear": p['empYear'], "tags": p['tags'], "other": p['other']})
    meta = dict(total=len(recs), totalTurnover=sum(x['turnover'] for x in recs),
                counties=sorted({x['county'] for x in recs if x['county']}))
    return {"meta": meta, "rows": recs}

WORKER_TMPL = '''const DATA = %s;
const SHELL = "%s";
const HTML = new TextDecoder().decode(Uint8Array.from(atob(SHELL), c => c.charCodeAt(0)));
const STATUSES = new Set(["Not contacted","Contacted","No answer","Follow up","Interested","Meeting","In progress","Won","Lost"]);
function J(o, code){ return new Response(JSON.stringify(o), { status: code||200, headers: {'content-type':'application/json; charset=utf-8','cache-control':'no-store'} }); }
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/data.json') return J(DATA);
    if (url.pathname === '/api/statuses') {
      try { const { results } = await env.DB.prepare("SELECT cui,status FROM lead_status").all(); const m={}; for (const r of results) m[r.cui]=r.status; return J(m); }
      catch(e){ return J({}); }
    }
    if (url.pathname === '/api/status' && request.method === 'POST') {
      let b; try { b = await request.json(); } catch { return J({error:'bad json'}, 400); }
      const cui = String(b.cui||''), status = String(b.status||'');
      if (!cui || !STATUSES.has(status)) return J({error:'invalid'}, 400);
      try { await env.DB.prepare("INSERT INTO lead_status(cui,status,updated_at) VALUES(?,?,?) ON CONFLICT(cui) DO UPDATE SET status=excluded.status, updated_at=excluded.updated_at").bind(cui, status, new Date().toISOString()).run(); return J({ok:true, cui, status}); }
      catch(e){ return J({error:String(e)}, 500); }
    }
    if (url.pathname === '/health') return new Response('ok');
    return new Response(HTML, { headers: { 'content-type':'text/html; charset=utf-8','cache-control':'no-store' } });
  }
};
'''

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, 'doc.xlsx')
    data = build_data(xlsx)
    json.dump(data, open(os.path.join(here, 'data.json'), 'w'), ensure_ascii=False)
    html = open(os.path.join(here, 'index.html'), 'rb').read()
    worker = WORKER_TMPL % (json.dumps(data, ensure_ascii=False), base64.b64encode(html).decode())
    open(os.path.join(here, 'worker.js'), 'w').write(worker)
    print(f"rows={data['meta']['total']} worker.js={len(worker.encode())}B")

if __name__ == '__main__':
    main()
