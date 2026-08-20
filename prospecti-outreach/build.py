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
URL = re.compile(r'(?:https?://|www\.)[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', re.I)
# free mail providers: their domain says nothing about the company having a site
FREE = {'gmail.com', 'yahoo.com', 'yahoo.ro', 'hotmail.com', 'hotmail.ro', 'outlook.com', 'outlook.ro',
        'icloud.com', 'msn.com', 'live.com', 'aol.com', 'protonmail.com', 'mail.ru', 'yandex.ru'}

def site_from(text, emails):
    """Website is not a CRM column: take an explicit URL from the notes, else infer from a corporate email domain."""
    m = URL.search(text or '')
    site = m.group(0) if m else next((e.split('@')[-1] for e in emails if e.split('@')[-1].lower() not in FREE), '')
    site = re.sub(r'^https?://', '', site, flags=re.I).strip().rstrip('/.,;')
    return re.sub(r'^www\.', '', site, flags=re.I).lower()

# fact keys that duplicate other sections (contact / location) and must not repeat in "details"
_DROP = ('localitate', 'mobil', 'telefon', 'fax', 'email', 'persoana de contact', 'persoană de contact', 'date de contact')

def parse_notes(t):
    out = {"phones": [], "emails": [], "prio": [], "facts": [], "person": "", "extra": []}
    if not t: return out
    emails = list(dict.fromkeys(EMAIL.findall(t)))
    out["prio"] = [m.group(1).strip() for m in re.finditer(r'Prio\s*\d+\s*:\s*([^;\n]+)', t)]
    mc = re.search(r'[Dd]ate de contact[^:]*:\s*(.+)', t)
    phones = []
    for pm in PHONE.findall(mc.group(1) if mc else t):
        d = re.sub(r'\D', '', pm)
        if 9 <= len(d) <= 12: phones.append(re.sub(r'\s+', ' ', pm.strip().strip(',;')))
    for line in t.splitlines():
        raw = line.strip()
        if not raw: continue
        low = raw.lower()
        if re.match(r'-?\s*(recomandari abordare|informatii companie|date de contact)', low): continue
        if 'prio' in low and ':' in raw: continue
        m = re.match(r'^-?\s*([^:]{2,60}?):\s*(.+)$', raw)
        if m:
            k, v = m.group(1).strip(), m.group(2).strip().rstrip(',')
            kl = k.lower()
            if kl.startswith(('persoana de contact', 'persoană de contact')): out["person"] = v; continue
            if kl.startswith('mobil'):
                for pm in PHONE.findall(v):
                    if 9 <= len(re.sub(r'\D', '', pm)) <= 12: phones.append(pm.strip())
                continue
            if kl.startswith('email'): emails += EMAIL.findall(v); continue
            if kl.startswith(_DROP): continue
            if v and len(k) <= 48: out["facts"].append([k, v]); continue
        else:
            out["extra"].append(raw.lstrip('- ').strip())
    out["phones"] = list(dict.fromkeys(phones))
    out["emails"] = list(dict.fromkeys(emails))
    out["website"] = site_from(t, out["emails"])
    return out

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
             status=I['Status detaliat'], due=gi('Data scaden'), prio=gi('Descriere'), cic=gi('CIC'), lastact=gi('Dată a celei'))
    recs = []
    for r in rows[1:]:
        try: turn = float(r[C['turn']])
        except (TypeError, ValueError): turn = 0.0
        note = s(r[C['prio']]); p = parse_notes(note)
        recs.append({"company": s(r[C['comp']]), "cui": s(r[C['cui']]), "turnover": turn,
                     "city": s(r[C['city']]).title() if s(r[C['city']]) else "",
                     "county": s(r[C['county']]).title() if s(r[C['county']]) else "",
                     "due": s(r[C['due']]), "cic": s(r[C['cic']]), "lastActivity": s(r[C['lastact']]),
                     "phones": p["phones"], "emails": p["emails"], "prio": p["prio"], "facts": p["facts"],
                     "person": p["person"], "extra": p["extra"], "website": p["website"]})
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
    b64 = base64.b64encode(html).decode()
    worker = WORKER_TMPL % (json.dumps(data, ensure_ascii=False), b64)
    open(os.path.join(here, 'worker.js'), 'w').write(worker)
    print(f"rows={data['meta']['total']} worker.js={len(worker.encode())}B")

if __name__ == '__main__':
    main()
