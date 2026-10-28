import json, time, urllib.request, urllib.parse
from pathlib import Path
rows=[]
cursor=None
while True:
    params={'market':'KRW-BTC','count':200}
    if cursor: params['to']=cursor+'Z'
    url='https://api.upbit.com/v1/candles/days?'+urllib.parse.urlencode(params)
    with urllib.request.urlopen(url, timeout=30) as res:
        batch=json.load(res)
    if not batch: break
    rows.extend(batch)
    nxt=min(r['candle_date_time_utc'] for r in batch)
    assert cursor is None or nxt < cursor
    cursor=nxt
    time.sleep(.25)
rows.sort(key=lambda r:r['candle_date_time_utc'])
assert len({r['candle_date_time_utc'] for r in rows})==len(rows)
Path('work/upbit_raw.json').write_text(json.dumps(rows),encoding='utf-8')
print(json.dumps({'count':len(rows),'first':rows[0]['candle_date_time_utc'],'last':rows[-1]['candle_date_time_utc'],'earliest_boundary_empty':True}))
