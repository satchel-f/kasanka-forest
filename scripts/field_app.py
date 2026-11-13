# CCS Field App data builder: one compact JSON per project area for the CDO/SCO field app.
# Reads the docs the tracker refresh already writes, so the numbers match the hub tracker.
# Usage: KOBO_TOKEN=... python3 -I field_app.py <in_dir> <area> <out.json>
#   in_dir holds the tracker's saved docs (any layout): board/current, lf/meta, lf/<area>, settings/adjust, fieldapp/deadlines.
#   It exports the Lead Farmer Progress Report and Signing forms from Kobo itself (heaps, holes, NRC duplicates).
import os, re, sys, json, glob, math, time, difflib, urllib.request, datetime as dtm, warnings
import numpy as np, pandas as pd
warnings.filterwarnings('ignore')
IND, AREA, OUT = sys.argv[1:4]

def unwrap(j):
    for _ in range(3):
        if isinstance(j, dict) and isinstance(j.get('data'), dict) and not ({'rows', 'sites', 'items', 'removed'} & set(j)): j = j['data']
    return j
docs = []
for p in glob.glob(os.path.join(IND, '**', '*.json'), recursive=True):
    try: docs.append((os.path.basename(p)[:-5], unwrap(json.load(open(p)))))
    except Exception: pass
def pick(test):
    for n, d in docs:
        if isinstance(d, dict) and test(n, d): return d
    return {}
board = pick(lambda n, d: 'totals' in d and 'areas' in d and 'sites' in d)
meta = pick(lambda n, d: 'dropoffs' in d and 'sites' in d)
area_docs = [d for n, d in docs if isinstance(d, dict) and 'rows' in d and 'area' in d]
adj = pick(lambda n, d: 'removed' in d or 'keep' in d)
dl_doc = pick(lambda n, d: 'items' in d and n.startswith('deadline'))
if not board or not meta or not area_docs: raise RuntimeError(f'missing tracker docs: board {bool(board)}, lf meta {bool(meta)}, lf areas {len(area_docs)}')

TOKEN = os.environ['KOBO_TOKEN']; API = 'https://kf.kobotoolbox.org/api/v2'
def req(url, data=None):
    r = urllib.request.Request(url, data=json.dumps(data).encode() if data else None, headers={'Authorization': 'Token ' + TOKEN, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r, timeout=90) as f: return f.read()
def export(prefix, path, lang='_xml'):
    a = [x for x in json.loads(req(f'{API}/assets/?format=json&limit=200'))['results'] if x['name'].startswith(prefix)]
    if not a: raise RuntimeError(f'Kobo form "{prefix}" not found')
    uid = sorted(a, key=lambda x: x.get('date_modified', ''))[-1]['uid']; UIDS[prefix] = uid
    e = json.loads(req(f'{API}/assets/{uid}/exports/?format=json', {'fields_from_all_versions': True, 'group_sep': '/', 'hierarchy_in_labels': False,
        'lang': lang, 'multiple_select': 'both', 'type': 'xls', 'include_media_url': False}))
    for _ in range(60):
        s_ = json.loads(req(e['url']))
        if s_['status'] == 'complete': open(path, 'wb').write(req(s_['result'])); return path
        if s_['status'] == 'error': raise RuntimeError(f'Kobo export failed for {prefix}: {s_.get("messages")}')
        time.sleep(5)
    raise RuntimeError(f'Kobo export timed out for {prefix}')
UIDS = {}
WORK = os.path.join(os.path.dirname(os.path.abspath(OUT)), 'kobo'); os.makedirs(WORK, exist_ok=True)
PROG = export('CCS Lead Farmer Progress Report', os.path.join(WORK, 'prog_x.xlsx'))
SIGN = export('CCS Lead Farmer Signing', os.path.join(WORK, 'sign_x.xlsx'))
_needs = json.dumps(dl_doc)
CP = export('CCSZ Champion Plot Monthly Verification', os.path.join(WORK, 'cp_x.xlsx')) if '"cp"' in _needs else None
PREV_FORMS = sorted({it['prev']['form'] for it in dl_doc.get('items', []) if isinstance(it.get('prev'), dict) and it['prev'].get('form')})
PREVX = {f: export(f, os.path.join(WORK, re.sub(r'[^a-z0-9]', '', f.lower()) + '.xlsx'), '_default') for f in PREV_FORMS}
nz = lambda v: '' if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()
num = lambda v: None if nz(v) == '' else int(float(v))
pct = lambda a, b: round(a / b, 3) if b else None
removed, moved, keep, site_over = adj.get('removed', {}), adj.get('moved', {}), adj.get('keep', {}), adj.get('sites', {})
NOW = pd.Timestamp.utcnow().tz_convert('Africa/Lusaka')
TODAY = NOW.date()
WSTART = TODAY - dtm.timedelta(days=TODAY.weekday())
AREAS = ['Kasanka', 'Mpulungu', 'Zambezi', 'Flamponje', 'Luapula', 'North-Western']
HEAP_AREA = {'Kasanka': 14651, 'Mpulungu': 17069, 'Flamponje': 6010, 'Luapula': 4366, 'North-Western': 3920, 'Zambezi': 18136}   # from the Progress Report form

# ---- sites (100) with lead farmer targets, overrides applied
sites = {}
for x in meta['sites']:
    o = site_over.get(x['code'], {})
    sites[x['code']] = dict(x, min=o.get('min', x['min']), tgt=o.get('tgt', x['tgt']))
bsite = {x['code']: x for x in board['sites']}

# ---- lead farmers, with the tracker's removals, moves and keep lists applied
lfs = []
for d in area_docs:
    for r in d['rows']:
        r = dict(r, area=d['area'])
        if r['id'] in removed: continue
        if r['id'] in moved: r['site'] = moved[r['id']]
        if r['site'] in keep and r['id'] not in keep[r['site']]: continue
        if r['site'] in sites: r['area'] = sites[r['site']]['area']
        lfs.append(r)

# ---- latest progress report per lead farmer: heaps, holes, fields, farmer trainings
prog = pd.read_excel(PROG, 0)
sign = pd.read_excel(SIGN, 0)
sign['nrc'] = sign.lf_nrc.astype(str).str.replace(r'\s', '', regex=True).str.lower()
sign['id'] = 'lf_' + sign._id.astype(str)
sign = sign.sort_values('_submission_time')
alias = {}
for _, g in sign.groupby('nrc'):
    for i in g.id: alias[i] = g.id.iloc[-1]
prog['lfk'] = [alias.get(nz(v), nz(v)) for v in prog.lf]
prog = prog.sort_values(['date', '_submission_time'])
P = {}
for _, r in prog.iterrows():
    e = P.setdefault(r.lfk, {})
    for k, cols in (('hd', ['heaps_done']), ('hv', ['heaps_dev']), ('holes', ['holes_dug']), ('marked', ['n_fields_marked']),
                    ('t1', ['n_t1_farmers', 'n_t1']), ('t2', ['n_t2_farmers', 'n_t2']), ('t3', ['n_t3_farmers', 'n_t3'])):
        v = next((num(r.get(c)) for c in cols if c in prog.columns and nz(r.get(c)) != ''), None)
        if v is not None: e[k] = v
    if nz(r.date) and pd.to_datetime(r.date).date() >= WSTART: e['wk'] = e.get('wk', 0) + 1
heap_tgt = lambda trees: math.ceil(trees * 11 / 500) if trees else 0   # same as the form: 1 heap per 50 trees + 10%

for r in lfs:
    e = P.get(r['id'], {})
    r.update(hd=e.get('hd'), hv=e.get('hv'), holes=e.get('holes'), marked=e.get('marked'), pr_wk=e.get('wk', 0),
             t1=max(r.get('t1') or 0, e.get('t1') or 0), t2=max(r.get('t2') or 0, e.get('t2') or 0), t3=max(r.get('t3') or 0, e.get('t3') or 0),
             ht=heap_tgt(r.get('trees')))

# ---- one spelling per field officer (Osward Kaoma = Osward Koama, Memory Jack Tembo = Memory Tembo)
def fo_key(n):
    w = re.sub(r'[^a-z ]', ' ', nz(n).lower()).split()
    return (w[0], w[-1]) if len(w) >= 2 else (w[0], '') if w else ('', '')
names = {}
for r in lfs: names.setdefault(r['fo'], 0); names[r['fo']] += 2
for x in board['sites']: names.setdefault(x['fo'], 0); names[x['fo']] += 1
canon = {}
groups = []
for n in sorted(names, key=lambda n: -names[n]):
    f, l = fo_key(n)
    if not f: canon[n] = n or 'Not listed'; continue
    for g in groups:
        gf, gl = fo_key(g[0])
        if difflib.SequenceMatcher(None, f, gf).ratio() >= 0.8 and (not l or not gl or difflib.SequenceMatcher(None, l, gl).ratio() >= 0.7):
            g.append(n); break
    else: groups.append([n])
for g in groups:
    for n in g: canon[n] = ' '.join(w.capitalize() for w in g[0].split())
for r in lfs: r['fo'] = canon.get(r['fo'], r['fo'])

# ---- metrics per area (for the comparison)
def metrics(area):
    S = [x for x in board['sites'] if x['area'] == area]
    L = [r for r in lfs if r['area'] == area]
    st = [s for s in sites.values() if s['area'] == area]
    seed, tgt = sum(x['seedlings'] for x in S), sum(x['target'] for x in S)
    lf_tgt = sum(s['tgt'] for s in st)
    fr = sum(r['farmers'] or 0 for r in L)
    heaps = sum((r['hd'] or 0) + (r['hv'] or 0) for r in L)
    return dict(area=area, sites=len(S), seedlings=seed, seed_target=tgt, seed_pct=pct(seed, tgt),
                r8_wk=sum(1 for x in S if x['wk'] == 'Sent'), r8_wk_pct=pct(sum(1 for x in S if x['wk'] == 'Sent'), len(S)),
                lf=len(L), lf_target=lf_tgt, lf_pct=pct(len(L), lf_tgt),
                s1=sum(1 for r in L if r['s1']), s1_pct=pct(sum(1 for r in L if r['s1']), len(L)),
                s2=sum(1 for r in L if r['s2']), s2_pct=pct(sum(1 for r in L if r['s2']), len(L)),
                pr=sum(1 for r in L if r['pr']), pr_pct=pct(sum(1 for r in L if r['pr']), len(L)),
                farmers=fr, t1=sum(r['t1'] for r in L), t1_pct=pct(sum(r['t1'] for r in L), fr),
                heaps=heaps, heaps_done=sum(r['hd'] or 0 for r in L), heap_target=HEAP_AREA.get(area, 0), heap_pct=pct(heaps, HEAP_AREA.get(area, 0)))
M = {a: metrics(a) for a in AREAS}
KEYS = ['seed_pct', 'r8_wk_pct', 'lf_pct', 's1_pct', 's2_pct', 'pr_pct', 't1_pct', 'heap_pct']
def avg(vals):   # average of the project areas' percentages (each area counts once)
    v = [min(x, 1) for x in vals if x is not None]; return round(sum(v) / len(v), 3) if v else None   # capped at 100% so one area far over target does not lift the average
AVG = {k: avg([M[a][k] for a in AREAS]) for k in KEYS}

# ---- deadlines: progress so far and pace
def metric_now(key, area, since):
    S = [x for x in board['sites'] if x['area'] == area]
    if key == 'r8':   # sites with an 8.0 since the deadline was released
        n = sum(1 for x in S if x.get('r8_date') and x['r8_date'] >= since); return n, len(S)
    m = M[area]
    return {'seedlings': (m['seedlings'], m['seed_target']), 'lf': (m['lf'], m['lf_target']), 's1': (m['s1'], m['lf']), 's2': (m['s2'], m['lf']),
            'pr': (m['pr'], m['lf']), 'heaps': (m['heaps'], m['heap_target']), 't1': (m['t1'], m['farmers'])}.get(key, (None, None))
AREA_SLUG = {'kasanka': 'Kasanka', 'mpulungu': 'Mpulungu', 'zambezi': 'Zambezi', 'flamponje': 'Flamponje', 'luapula': 'Luapula', 'north_western': 'North-Western'}
CP_ROSTER = {}
if CP:
    for c in json.loads(req(f'{API}/assets/{UIDS["CCSZ Champion Plot Monthly Verification"]}/?format=json'))['content'].get('choices', []):
        if c.get('list_name') == 'farmer_list' and str(c.get('name', '')).startswith('CH'): CP_ROSTER[c['name']] = AREA_SLUG.get(c.get('project_area'), c.get('project_area'))
    cpdf = pd.read_excel(CP, 0); cpdf['t'] = pd.to_datetime(cpdf._submission_time)
def cp_counts(start, end, override):
    # champions verified in the round: one report per champion counts; write-ins counted once by name
    w = cpdf[(cpdf.t > pd.Timestamp(start)) & (cpdf.t <= pd.Timestamp(end or '2100-01-01'))]
    seen = {}
    for _, r in w.iterrows():
        code = nz(r.farmer_select)
        if code.startswith('CH'): key, area = code, CP_ROSTER.get(code, AREA_SLUG.get(nz(r.project_area), ''))
        elif nz(r.farmer_writein_name): key, area = 'w:' + nz(r.farmer_writein_name).lower(), AREA_SLUG.get(nz(r.project_area), '')
        else: continue
        nm = nz(r.farmer_writein_name) or ''
        for o, a in override.items():
            if o.lower() in nm.lower(): area = a
        seen[key] = area
    out = {}
    for a in seen.values(): out[a] = out.get(a, 0) + 1
    return out
def prev_sites(form, area):   # sites in this area with at least one report on an earlier nursery form
    df = pd.read_excel(PREVX[form], 0)
    colname = next((c for c in df.columns if str(c).startswith('Nursery site')), None)
    codes = set(df[colname].astype(str).str.extract(r'(S\d{3})')[0].dropna()) if colname else set()
    S = [c for c, x in sites.items() if x['area'] == area]
    return sum(1 for c in S if c in codes), len(S)
deadlines, ongoing = [], []
for d in dl_doc.get('items', []):
    if d.get('areas') and AREA not in d['areas'] and 'All' not in d['areas']: continue
    if d.get('metric') == 'dropoffs':   # count only: proposals have no fixed number to reach
        d = dict(d, n=sum(1 for x in meta.get('dropoffs', []) if x.get('area') == AREA))
    if not d.get('due'):   # standing guidance or a goal with no date
        x = dict(id=d.get('id', ''), title=d['title'], note=d.get('note', ''), released=d.get('released', ''), metric=d.get('metric', ''))
        if x['metric'] in ('heaps', 's1', 's2', 'pr', 'lf', 't1', 'seedlings'):
            n, of = metric_now(x['metric'], AREA, ''); x.update(n=n, of=of, pct=pct(n, of), avg=avg([pct(*metric_now(x['metric'], a, '')) for a in AREAS]))
        ongoing.append(x); continue
    due = dtm.date.fromisoformat(d['due']); rel = dtm.date.fromisoformat(d.get('released') or d['due'])
    start = dtm.date.fromisoformat(d.get('start') or str(rel))
    x = dict(id=d.get('id', ''), title=d['title'], due=d['due'], released=str(rel), note=d.get('note', ''), metric=d.get('metric', ''),
             goal=d.get('goal', 1), days_left=(due - TODAY).days)
    if 'n' in d: x['n'] = d['n']
    if d.get('metric') == 'cp' and CP is not None:
        exp_ = d.get('expected', {}); ov = d.get('area_override', {})
        cur = cp_counts(d['round_start'], None, ov)
        x.update(n=cur.get(AREA, 0), of=exp_.get(AREA), pct=pct(cur.get(AREA, 0), exp_.get(AREA)),
                 avg=avg([pct(cur.get(a, 0), exp_.get(a)) for a in AREAS if exp_.get(a)]))
        span = max((due - start).days, 1); el = min(max((TODAY - start).days, 0), span)
        x['expected'] = round(min(el / span, 1) * x['goal'], 3)
        pv = d.get('prev')
        if pv and pv.get('round_start'):
            pc = cp_counts(pv['round_start'], pv.get('round_end'), ov)
            x['prev'] = dict(label=pv['label'], n=pc.get(AREA, 0), of=exp_.get(AREA), pct=pct(pc.get(AREA, 0), exp_.get(AREA)))
    if isinstance(d.get('prev'), dict) and d['prev'].get('form') and d['prev']['form'] in PREVX:
        n_, of_ = prev_sites(d['prev']['form'], AREA)
        x['prev'] = dict(label=d['prev']['label'], n=n_, of=of_, pct=pct(n_, of_))
    if d.get('pct') is not None: x.update(pct=d['pct'], pct_note=d.get('pct_note', ''))   # figure given by the owner, no Kobo source
    if x['metric'] and x['metric'] not in ('dropoffs', 'cp'):
        n, of = metric_now(x['metric'], AREA, str(start))
        if of:
            x.update(n=n, of=of, pct=pct(n, of))
            span = max((due - start).days, 1); el = min(max((TODAY - start).days, 0), span)
            x['expected'] = round(min(el / span, 1) * x['goal'], 3)
            x['avg'] = avg([pct(*metric_now(x['metric'], a, str(start))) for a in AREAS])
    deadlines.append(x)
deadlines.sort(key=lambda x: x['due'])

# ---- this area: sites, lead farmers, field officers
A = AREA
SITES = []
for code, s in sorted(sites.items()):
    if s['area'] != A: continue
    b = bsite.get(code, {})
    L = sorted([r for r in lfs if r['site'] == code], key=lambda r: r['name'])
    SITES.append(dict(code=code, name=s['name'], cfmg=s['cfmg'], type=s['type'], cdo=s['cdo'], fo=canon.get(b.get('fo', ''), b.get('fo', '')),
        seed=b.get('seedlings', 0), target=b.get('target', s['seedlings']), germ=b.get('germ'), sent=b.get('wk') == 'Sent', r8=b.get('r8_date', ''),
        lf_tgt=s['tgt'], lf_min=s['min'],
        lfs=[[r['name'], r['fo'], r['s1'], r['s2'], r['pr'], r['pr_last'], r['farmers'], r['trees'], r['hd'], r['hv'], r['ht'], r['t1'], r['t2'], r['t3'],
              r['holes'], r['dec'], 1 if r['ret'] else 0] for r in L]))
fos = {}
for s in SITES:
    if s['fo']:
        f = fos.setdefault(s['fo'], dict(name=s['fo'], sites=[], cdos=set(), r8_due=0, r8_sent=0, lf=0, s1=0, s2=0, pr=0, pr_wk=0, heaps=0, heap_t=0))
        f['cdos'].add(s['cdo']); f['sites'].append(s['code']); f['r8_due'] += 1; f['r8_sent'] += 1 if s['sent'] else 0
for r in lfs:
    if r['area'] != A: continue
    f = fos.setdefault(r['fo'], dict(name=r['fo'], sites=[], cdos=set(), r8_due=0, r8_sent=0, lf=0, s1=0, s2=0, pr=0, pr_wk=0, heaps=0, heap_t=0))
    f['cdos'].add(sites[r['site']]['cdo'] if r['site'] in sites else r['cdo']); f['lf'] += 1; f['s1'] += 1 if r['s1'] else 0; f['s2'] += 1 if r['s2'] else 0; f['pr'] += 1 if r['pr'] else 0
    f['pr_wk'] += 1 if r['pr_wk'] else 0; f['heaps'] += (r['hd'] or 0) + (r['hv'] or 0); f['heap_t'] += r['ht']
FOS = sorted([dict(f, cdos=sorted(f['cdos'])) for f in fos.values()], key=lambda f: f['name'])

out = dict(v=1, area=A, asof=NOW.strftime('%Y-%m-%d %H:%M'), week_start=str(WSTART), cdos=sorted({s['cdo'] for s in SITES}),
           me=M[A], avg=AVG, n_areas=len(AREAS), deadlines=deadlines, ongoing=ongoing, sites=SITES, fos=FOS,
           lf_cols=['name', 'fo', 's1', 's2', 'pr', 'pr_last', 'farmers', 'trees', 'heaps_done', 'heaps_dev', 'heap_target', 't1', 't2', 't3', 'holes', 'decision', 'returning'])
os.makedirs(os.path.dirname(OUT) or '.', exist_ok=True)
json.dump(out, open(OUT, 'w'), separators=(',', ':'))
m = M[A]
print(json.dumps({'area': A, 'bytes': os.path.getsize(OUT), 'sites': len(SITES), 'lead_farmers': m['lf'], 'fos': len(FOS), 'deadlines': len(deadlines),
                  'seed_pct': m['seed_pct'], 'heaps': m['heaps']}))
