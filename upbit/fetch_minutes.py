"""Download all available Upbit KRW-BTC one-minute candles, with resumable monthly files."""
import json, gzip, time, threading, http.client, urllib.parse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

root=Path('work/minute_cache'); root.mkdir(exist_ok=True)
lock=threading.Lock(); next_at=0.; stop=threading.Event()
counts={'requests':0,'candles':0,'months_done':0,'retries':0}; started=time.monotonic()
local=threading.local()
def request(to=None,count=200):
 global next_at
 for attempt in range(10):
  if stop.is_set(): raise RuntimeError('Cancelled after another worker failed')
  with lock:
   now=time.monotonic(); delay=max(0,next_at-now); next_at=max(now,next_at)+.118
  if delay:time.sleep(delay)
  params={'market':'KRW-BTC','count':count}
  if to:params['to']=to+'Z'
  path='/v1/candles/minutes/1?'+urllib.parse.urlencode(params)
  try:
   if not getattr(local,'conn',None):local.conn=http.client.HTTPSConnection('api.upbit.com',timeout=30)
   local.conn.request('GET',path,headers={'Accept':'application/json','User-Agent':'BTC-historical-research/1.0'})
   res=local.conn.getresponse(); body=res.read()
   with lock:counts['requests']+=1
   if res.status==200:
    data=json.loads(body);assert isinstance(data,list)
    return data
   if res.status in (429,418):
    pause=max(3,float(res.getheader('Retry-After','5')))
    with lock:next_at=max(next_at,time.monotonic()+pause)
    time.sleep(pause)
   elif res.status>=500: time.sleep(min(30,2**attempt))
   else:raise RuntimeError(f'HTTP {res.status}: {body[:300]}')
  except (OSError,http.client.HTTPException):
   if getattr(local,'conn',None):local.conn.close()
   local.conn=None;time.sleep(min(30,2**attempt))
  with lock:counts['retries']+=1
 raise RuntimeError('API failed after 10 retries')

meta_path=root/'download_meta.json'
if meta_path.exists(): meta=json.loads(meta_path.read_text())
else:
 latest=request(count=1)[0]['candle_date_time_utc']
 # Exclude the current minute; retain all earlier candles up to a fixed cutoff.
 meta={'market':'KRW-BTC','unit':1,'cutoff_exclusive_utc':latest,'created_at_utc':datetime.now(timezone.utc).isoformat(),
       'source':'https://api.upbit.com/v1/candles/minutes/1','earliest_probe_empty':request('2017-09-25T00:00:00')==[]}
 meta_path.write_text(json.dumps(meta,indent=2))
cutoff=meta['cutoff_exclusive_utc']
months=[]
y,m=2017,9
while f'{y:04}-{m:02}-01T00:00:00'<cutoff:
 start=f'{y:04}-{m:02}-01T00:00:00'
 yn,mn=(y+1,1) if m==12 else (y,m+1)
 end=min(f'{yn:04}-{mn:02}-01T00:00:00',cutoff)
 months.append((start,end)); y,m=yn,mn

def fetch_month(pair):
 start,end=pair; key=start[:7]; final=root/f'{key}.json.gz'; partial=root/f'{key}.partial.json.gz'
 if final.exists():
  with gzip.open(final,'rt',encoding='utf-8') as f: rows=json.load(f)
  with lock:counts['candles']+=len(rows);counts['months_done']+=1
  return key,len(rows)
 if partial.exists():
  with gzip.open(partial,'rt',encoding='utf-8') as f: rows=json.load(f)
  cursor=min(r['candle_date_time_utc'] for r in rows)
 else:rows=[];cursor=end
 pages=0;initial=len(rows)
 with lock:counts['candles']+=initial
 while True:
  batch=request(cursor)
  if not batch:break
  times=[r['candle_date_time_utc'] for r in batch]
  assert max(times)<cursor
  kept=[r for r in batch if start<=r['candle_date_time_utc']<end]
  rows.extend(kept)
  with lock:counts['candles']+=len(kept)
  nxt=min(times);assert nxt<cursor;cursor=nxt;pages+=1
  if pages%40==0:
   with gzip.open(partial,'wt',encoding='utf-8',compresslevel=1) as f:json.dump(rows,f,separators=(',',':'))
  if nxt<=start:break
 rows.sort(key=lambda r:r['candle_date_time_utc'])
 assert len({r['candle_date_time_utc'] for r in rows})==len(rows)
 with gzip.open(final,'wt',encoding='utf-8',compresslevel=3) as f:json.dump(rows,f,separators=(',',':'))
 if partial.exists():partial.unlink()
 with lock:counts['months_done']+=1
 return key,len(rows)

def monitor():
 while not stop.wait(30):
  with lock:msg=dict(counts)
  msg.update(total_months=len(months),elapsed_sec=round(time.monotonic()-started))
  (root/'progress.json').write_text(json.dumps(msg))
  print(json.dumps(msg),flush=True)
threading.Thread(target=monitor,daemon=True).start()
print(json.dumps({'status':'started','months':len(months),'cutoff':cutoff}),flush=True)
try:
 with ThreadPoolExecutor(max_workers=6) as pool:
  futures=[pool.submit(fetch_month,pair) for pair in reversed(months)]
  for future in as_completed(futures):
   key,n=future.result()
   print(json.dumps({'month_completed':key,'rows':n}),flush=True)
finally:stop.set()
summary={**counts,'elapsed_sec':round(time.monotonic()-started),'complete':True,'total_months':len(months)}
(root/'progress.json').write_text(json.dumps(summary,indent=2))
print(json.dumps(summary),flush=True)
