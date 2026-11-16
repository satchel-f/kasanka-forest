# CCS nursery forecasts for the hub tracker's Forecasts tab.
# Hardening off (confirmed and estimated, pace, projection against goals), the handout plan
# (per day, per week, per site, per area) and the farmer-readiness constraint.
# Usage: KOBO_TOKEN=... python3 -I forecast.py <in_dir> <out.json>
#   in_dir holds the tracker's saved docs (any layout): board/current, lf/meta, lf/<area>, settings/adjust, forecast/settings (optional).
import os, re, sys, json, glob, time, math, urllib.request, datetime as dtm, warnings
import numpy as np, pandas as pd
warnings.filterwarnings('ignore')
IND, OUT = sys.argv[1:3]

def unwrap(j):
    for _ in range(3):
        if isinstance(j, dict) and isinstance(j.get('data'), dict) and not ({'rows', 'sites', 'items', 'removed', 'goals'} & set(j)): j = j['data']
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
cfg = pick(lambda n, d: 'goals' in d and 'handout' in d)
if not board or not meta: raise RuntimeError('missing tracker docs (board/current, lf/meta)')

# ---- goals (forecast/settings in the tracker overrides these)
G = dict(goals=[{'label': '1.5 million hardened', 'value': 1500000, 'by': '2026-12-15'},
                {'label': '2 million hardened', 'value': 2000000, 'by': '2026-12-15'}],
         handout={'start': '2026-11-30', 'steps': [{'label': '1.5 million handed out', 'value': 1500000, 'by': '2026-12-31'},
                                                   {'label': 'All seedlings handed out', 'value': 'all', 'by': '2027-01-31'}]},
         workdays='Mon-Sat', holidays=['2026-12-25', '2027-01-01'])
G.update({k: v for k, v in cfg.items() if k in G})

TOKEN = os.environ['KOBO_TOKEN']; API = 'https://kf.kobotoolbox.org/api/v2'
def req(url, data=None):
    r = urllib.request.Request(url, data=json.dumps(data).encode() if data else None, headers={'Authorization': 'Token ' + TOKEN, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(r, timeout=90) as f: return f.read()
def export(prefix, path):
    a = [x for x in json.loads(req(f'{API}/assets/?format=json&limit=200'))['results'] if x['name'].startswith(prefix)]
    if not a: raise RuntimeError(f'Kobo form "{prefix}" not found')
    uid = sorted(a, key=lambda x: x.get('date_modified', ''))[-1]['uid']
    e = json.loads(req(f'{API}/assets/{uid}/exports/?format=json', {'fields_from_all_versions': True, 'group_sep': '/', 'hierarchy_in_labels': False,
        'lang': '_xml', 'multiple_select': 'both', 'type': 'xls', 'include_media_url': False}))
    for _ in range(60):
        s_ = json.loads(req(e['url']))
        if s_['status'] == 'complete': open(path, 'wb').write(req(s_['result'])); return path
        if s_['status'] == 'error': raise RuntimeError(f'Kobo export failed for {prefix}: {s_.get("messages")}')
        time.sleep(5)
    raise RuntimeError(f'Kobo export timed out for {prefix}')
WORK = os.path.join(os.path.dirname(os.path.abspath(OUT)), 'kobo_fc'); os.makedirs(WORK, exist_ok=True)
R8 = export('CCS Nursery Report 8.0', os.path.join(WORK, 'r8.xlsx'))

nz = lambda v: '' if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()
TODAY = pd.Timestamp.utcnow().tz_convert('Africa/Lusaka').date()
D = lambda s: dtm.date.fromisoformat(s)

# ---- sites and stock (latest count per site, as on the board)
S = pd.DataFrame(board['sites'])
S = S[['code', 'site', 'area', 'cfmg', 'type', 'cdo', 'seedlings', 'target']].copy()
total_stock = int(S.seedlings.sum())

# ---- hardening off from Report 8.0: latest report per site, with the date hardening started
r = pd.read_excel(R8, 0)
r['code'] = r['site'].astype(str).str.extract(r'(S\d{3})')[0]
r = r[r.code.isin(S.code)]
r['vdate'] = pd.to_datetime(r.visit_date, errors='coerce')
r = r.sort_values(['vdate', '_submission_time']).groupby('code').last().reset_index()
r['tot_hard'] = pd.to_numeric(r.tot_hard, errors='coerce').fillna(0)
r['tot_seedlings'] = pd.to_numeric(r.tot_seedlings, errors='coerce').fillna(0)
m = S.merge(r[['code', 'tot_seedlings', 'tot_hard', 'hard_date', 'vdate']], on='code', how='left')
m['seedlings'] = [int(t) if pd.notna(t) and t > 0 else int(x) for t, x in zip(m.tot_seedlings, m.seedlings)]   # the newest count: 8.0 where a site has one
S['seedlings'] = m['seedlings'].values
total_stock = int(S.seedlings.sum())
rep = m[m.tot_hard.notna()]
share = {a: (g.tot_hard.sum() / g.tot_seedlings.sum()) if g.tot_seedlings.sum() else 0 for a, g in rep.groupby('area')}
overall_share = rep.tot_hard.sum() / rep.tot_seedlings.sum() if rep.tot_seedlings.sum() else 0
m['hard_est'] = [h if pd.notna(h) else s * share.get(a, overall_share) for h, s, a in zip(m.tot_hard, m.seedlings, m.area)]
confirmed = int(rep.tot_hard.sum()); est_total = int(round(m.hard_est.sum()))

h = rep[rep.tot_hard > 0].copy(); h['hd'] = pd.to_datetime(h.hard_date, errors='coerce').fillna(h.vdate)
curve = h.groupby(h.hd.dt.date).tot_hard.sum().sort_index().cumsum()
last = max(curve.index) if len(curve) else TODAY
win_from = last - dtm.timedelta(days=13)
pace = int(h[h.hd.dt.date >= win_from].tot_hard.sum() / 14) if len(h) else 0   # seedlings starting hardening off per day, last 14 days of start dates
# projection of the estimate for all 100 sites: grows at the observed pace, slowing as it nears total stock
def project(start_val, start_day, until):
    out, v, d = [], float(start_val), start_day
    while d <= until:
        out.append((d, int(v)))
        room = total_stock - v
        v += min(pace * (room / max(total_stock - start_val, 1)) ** 0.5, room)   # slows as fewer unhardened seedlings are left
        d += dtm.timedelta(days=1)
    return out
end_day = max(D(g['by']) for g in G['goals'])
proj = project(est_total, TODAY, end_day + dtm.timedelta(days=14))
goals = []
for g in G['goals']:
    hit = next((d for d, v in proj if v >= g['value']), None)
    by = D(g['by'])
    goals.append(dict(label=g['label'], value=g['value'], by=g['by'], confirmed_pct=round(confirmed / g['value'], 3), est_pct=round(est_total / g['value'], 3),
                      proj_date=str(hit) if hit else '', status='Reached (estimate)' if est_total >= g['value'] else ('On track' if hit and hit <= by else 'At risk'),
                      need_per_day=max(0, int((g['value'] - est_total) / max((by - TODAY).days, 1)))))
hard = dict(confirmed=confirmed, est_total=est_total, total_stock=total_stock, sites_reported=int(len(rep)), sites_total=int(len(S)),
            not_started=sorted(rep[rep.tot_hard == 0].code.tolist()), seed_reported=int(rep.tot_seedlings.sum()),
            share_reported=round(overall_share, 3), pace=pace, pace_window=[str(win_from), str(last)], goals=goals,
            curve=[[str(d), int(v)] for d, v in curve.items()], proj=[[str(d), v] for d, v in proj[::2]],
            areas=[dict(area=a, sites=int(len(g)), reported=int(g.tot_hard.notna().sum()), stock=int(g.seedlings.sum()), confirmed=int(g.tot_hard.fillna(0).sum()),
                        est=int(g.hard_est.sum()), share=round(share.get(a, 0), 3)) for a, g in m.groupby('area')])

# ---- handout plan: working days, per day, per site, per area
hol = {D(x) for x in G['holidays']}
def workdays(a, b):
    n, d = 0, a
    while d <= b:
        if d.weekday() < 6 and d not in hol: n += 1
        d += dtm.timedelta(days=1)
    return n
start = D(G['handout']['start'])
steps, prev_val, prev_end = [], 0, start - dtm.timedelta(days=1)
for st in G['handout']['steps']:
    val = total_stock if st['value'] == 'all' else int(st['value'])
    by = D(st['by']); wd = workdays(prev_end + dtm.timedelta(days=1), by)
    steps.append(dict(label=st['label'], value=val, by=st['by'], from_=str(prev_end + dtm.timedelta(days=1)), workdays=wd,
                      per_day=int((val - prev_val) / max(wd, 1)), per_site_day=int((val - prev_val) / max(wd, 1) / len(S)), share_of_stock=round(val / total_stock, 3)))
    prev_val, prev_end = val, by
# weekly cumulative plan (even pace per working day within each step)
weeks, cum, d = [], 0, start
step_i = 0; rate = steps[0]['per_day']
while d <= prev_end:
    while step_i < len(steps) - 1 and d > D(steps[step_i]['by']): step_i += 1; rate = steps[step_i]['per_day']
    if d.weekday() < 6 and d not in hol: cum += rate
    if d.weekday() == 5 or d == prev_end or d == D(steps[step_i]['by']):
        cum = min(cum, steps[step_i]['value']) if d == D(steps[step_i]['by']) else cum
        if d == D(steps[step_i]['by']): cum = steps[step_i]['value']
        weeks.append([str(d), int(cum)])
    d += dtm.timedelta(days=1)
first = steps[0]
S['hard_est'] = m.hard_est.round().astype(int).values
S['hard_conf'] = m.tot_hard.values
S['by1'] = (S.seedlings * first['value'] / total_stock).round(-2).astype(int)
S['day1'] = (S.by1 / max(first['workdays'], 1)).round().astype(int)
S['rest'] = (S.seedlings - S.by1).astype(int)
S['day2'] = (S.rest / max(steps[-1]['workdays'], 1)).round().astype(int) if len(steps) > 1 else 0
sites_out = [dict(code=x.code, site=x.site, area=x.area, cdo=x.cdo, type=x.type, stock=int(x.seedlings), hard=None if pd.isna(x.hard_conf) else int(x.hard_conf),
                  by1=int(x.by1), day1=int(x.day1), rest=int(x.rest), day2=int(x.day2)) for x in S.sort_values('seedlings', ascending=False).itertuples()]
spread = dict(median=int(S.by1.median()), top10=int(S.nlargest(10, 'seedlings').by1.mean()), bottom10=int(S.nsmallest(10, 'seedlings').by1.mean()),
              cfmg=int(S[S.type == 'CFMG'].by1.mean()), en=int(S[S.type == 'EN'].by1.mean()),
              cfmg_day=int(S[S.type == 'CFMG'].day1.mean()), en_day=int(S[S.type == 'EN'].day1.mean()))
dist = dict(start=str(start), steps=steps, weeks=weeks, sites=sites_out, spread=spread,
            areas=[dict(area=a, sites=int(len(g)), stock=int(g.seedlings.sum()), by1=int(g.by1.sum()), day1=int(g.day1.sum()), rest=int(g.rest.sum()), day2=int(g.day2.sum()))
                   for a, g in S.groupby('area')])

# ---- farmer readiness: lead farmers, trainings and the trees their farmers have asked for
removed, moved, keep = adj.get('removed', {}), adj.get('moved', {}), adj.get('keep', {})
lfs = []
for d_ in area_docs:
    for x in d_['rows']:
        if x['id'] in removed: continue
        site = moved.get(x['id'], x['site'])
        if site in keep and x['id'] not in keep[site]: continue
        lfs.append(dict(x, area=d_['area'], site=site))
L = pd.DataFrame(lfs)
fr = []
for a, g in L.groupby('area'):
    rep_ = g[g.pr > 0]
    trees = pd.to_numeric(rep_.trees, errors='coerce')
    fr.append(dict(area=a, lf=int(len(g)), s1=int((g.s1 != '').sum()), s2=int((g.s2 != '').sum()), pr=int(len(rep_)),
                   farmers=int(pd.to_numeric(rep_.farmers, errors='coerce').fillna(0).sum()), trees=int(trees.fillna(0).sum()), trees_lf=int(trees.notna().sum()),
                   trees_per_lf=int(trees.mean()) if trees.notna().any() else None, stock=int(S[S.area == a].seedlings.sum())))
tot = {k: sum(x[k] for x in fr) for k in ('lf', 's1', 's2', 'pr', 'farmers', 'trees', 'trees_lf', 'stock')}
avg_trees = (tot['trees'] / tot['trees_lf']) if tot['trees_lf'] else 0
farm = dict(areas=fr, total=tot, trees_per_lf=int(avg_trees), demand_est=int(avg_trees * tot['lf']),
            note='Trees asked for comes from LF Progress Reports (farmers each lead farmer has, and trees they want). Only lead farmers with a report count, so demand is under-counted until reports are in.')

out = dict(asof=pd.Timestamp.utcnow().tz_convert('Africa/Lusaka').strftime('%Y-%m-%d %H:%M'), today=str(TODAY), hard=hard, dist=dist, farm=farm, goals_src='forecast/settings' if cfg else 'defaults')
os.makedirs(os.path.dirname(OUT) or '.', exist_ok=True)
json.dump(out, open(OUT, 'w'), separators=(',', ':'))
print(json.dumps({'confirmed': confirmed, 'est_total': est_total, 'stock': total_stock, 'pace': pace, 'goals': [(g['label'], g['status'], g['proj_date']) for g in goals],
                  'handout': [(s['label'], s['per_day'], s['workdays']) for s in steps], 'demand_est': farm['demand_est'], 'bytes': os.path.getsize(OUT)}))
