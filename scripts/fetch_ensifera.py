#!/usr/bin/env python3
"""
Genera ensifera.json para el Panel de Clientes a partir del dashboard de Ensífera
en Cloudflare (https://ensifera-dashboard.contacto-0e4.workers.dev/).

El Worker no envía cabeceras CORS, así que el panel (GitHub Pages) no puede leerlo
directo desde el navegador; esta copia la hace GitHub Actions cada hora.

Misma lógica que el Worker (buildMerged):
  salesCOP      = valor   del registro del coordinador  (/api/coord)
  ventas        = ventas  del registro del coordinador
  conversations = conv    del registro del coordinador  (Mercately, total)
  resultsAds    = conv_api de Meta                       (/api/daily)
  spendCOP      = spend   de Meta (la cuenta está en COP)
  campaigns     = /api/campaigns por día (gasto + resultados)

Run: python scripts/fetch_ensifera.py
"""
import json, os, sys, time
from datetime import date, datetime, timedelta, timezone
from urllib.parse import urlencode
from urllib.request import Request, urlopen

BASE = 'https://ensifera-dashboard.contacto-0e4.workers.dev'
METRIC = 'messaging_conversation_started_7d'  # la métrica por defecto del Worker
SINCE = '2026-06-01'
REFRESH_DAYS = 7          # días recientes que se vuelven a pedir siempre (Meta ajusta atribución)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'ensifera.json')

def api(path, **params):
    url = f'{BASE}{path}' + (('?' + urlencode(params)) if params else '')
    for attempt in range(3):
        try:
            # Cloudflare rechaza (403) el user-agent por defecto de urllib
            req = Request(url, headers={'User-Agent': 'jk-dashboard-panel/1.0 (+https://jkmaketing.github.io/jk-dashboard/)',
                                        'Accept': 'application/json'})
            with urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode('utf-8'))
        except Exception as e:
            if attempt == 2:
                raise RuntimeError(f'{path}: {e}')
            time.sleep(3 * (attempt + 1))

bogota = timezone(timedelta(hours=-5))
today = datetime.now(bogota).date()
until = today.isoformat()

daily = {d['date']: d for d in api('/api/daily', since=SINCE, until=until, metric=METRIC)}
coord = {}
for c in api('/api/coord'):
    coord.setdefault(c['fecha'], c)  # el Worker usa el primer registro de cada fecha

# Campañas por día: se reutiliza lo ya descargado salvo los últimos REFRESH_DAYS días
prev = {}
try:
    with open(OUT, encoding='utf-8') as f:
        prev = {d['date']: d.get('campaigns', []) for d in json.load(f).get('days', [])}
except FileNotFoundError:
    pass

refresh_from = (today - timedelta(days=REFRESH_DAYS)).isoformat()
dates = sorted(set(daily) | set(coord))
days = []
for d in dates:
    if d < SINCE or d > until:
        continue
    a = daily.get(d, {})
    c = coord.get(d, {})
    spend = a.get('spend') or 0
    if spend and (d >= refresh_from or d not in prev):
        camps = api('/api/campaigns', since=d, until=d, metric=METRIC)
        campaigns = [{'name': x['name'], 'spendCOP': x.get('spend') or 0,
                      'results': x.get('conv') or 0,
                      'impressions': x.get('impressions') or 0, 'reach': x.get('reach') or 0}
                     for x in camps if (x.get('spend') or 0) > 0]
    else:
        campaigns = prev.get(d, []) if spend else []
    days.append({
        'date': d,
        'spendCOP': spend,
        'impressions': a.get('impressions') or 0,
        'reach': a.get('reach') or 0,
        'resultsAds': a.get('conv_api') or 0,
        'conversations': c.get('conv') or 0,
        'ventas': c.get('ventas') or 0,
        'salesCOP': c.get('valor') or 0,
        'notas': c.get('notas') or '',
        'campaigns': campaigns,
    })

if not days:
    print('ERROR: el Worker no devolvió datos — no se sobrescribe ensifera.json', file=sys.stderr)
    sys.exit(1)

out = {'source': BASE, 'metric': METRIC,
       'generatedAt': datetime.now(timezone.utc).isoformat(timespec='seconds'),
       'days': days}
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

m = today.strftime('%Y-%m')
mm = [x for x in days if x['date'].startswith(m)]
print(f'Done: {len(days)} días ({days[0]["date"]} → {days[-1]["date"]}). '
      f'{m}: inversión {sum(x["spendCOP"] for x in mm):,} · venta {sum(x["salesCOP"] for x in mm):,}')
