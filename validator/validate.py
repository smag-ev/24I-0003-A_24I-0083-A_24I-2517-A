#!/usr/bin/env python3
"""Independent V1-V8 validator for the LIFELINE Iteration-1 workbook."""
from __future__ import annotations
import argparse, math, itertools, string, hashlib
from collections import defaultdict, Counter, deque
from pathlib import Path
from openpyxl import load_workbook

HEADERS={
'cases':['case_id','group_seed','case_no','tier','scenario','capacity','target_0','target_1','target_2','edit_bound','tour_length','sensor_nodes','sensor_k','safety_vars','rations_target'],
'crates':['case_id','crate_id','cat','weight','value','expiry','serial'],
'locations':['case_id','loc_id','type','penalty'],'roads':['case_id','road_no','u','v','minutes','cable'],
'waves':['case_id','wave_no','road_no'],'radio_log':['case_id','part','text'],
'radio_patterns':['case_id','pattern_no','pattern','count'],'radio_reports':['case_id','report_a','report_b'],
'windows':['case_id','window_id','start','finish','value'],'incidents':['case_id','alarm_id','tick','priority'],
'sites':['case_id','site_no','x','y'],'tour':['case_id','step','site_no'],
'pipeline_dims':['case_id','idx','dim'],'drum_prices':['case_id','length','price'],
'sensor_edges':['case_id','u','v'],'sensor_cover':['case_id','node'],
'safety_rules':['case_id','rule_no','lit_1','lit_2','lit_3'],'safety_planted':['case_id','switch','value'],
'rations':['case_id','item','value','planted'],
'claims':['case_id','cargo_opt','window_count_opt','window_value_opt','drum_opt','tour_opt','cover_min']}
ALPH=set(string.ascii_uppercase+string.digits+'_-'); PATLEN=[2,4,8,12,20,40]
RANGES={'S':(8,20,6,12,200,500,8,20,10,40,5,9,4,8,5,14,6,14,6,15,8,20),
'M':(40,200,30,150,2000,10000,50,300,200,1000,10,16,15,60,40,200,20,40,20,50,24,40),
'L':(500,1500,500,2000,20000,100000,1000,3000,2000,5000,200,500,100,300,500,2000,100,400,20,50,50,100)}
def tier(k): return 'S' if k<20 else 'M' if k<40 else 'L' if k<46 else 'X'
def edit_distance(a,b):
    p=list(range(len(b)+1))
    for i,x in enumerate(a,1):
        q=[i]
        for j,y in enumerate(b,1):q.append(min(q[-1]+1,p[j]+1,p[j-1]+(x!=y)))
        p=q
    return p[-1]
def countpat(s,p):
    n=0; i=0
    while True:
        i=s.find(p,i)
        if i<0:return n
        n+=1;i+=1
def connected(n,edges,removed=()):
    adj=[[] for _ in range(n)]; rem=set(removed)
    for rid,(u,v,*_) in enumerate(edges):
        if rid not in rem and 0<=u<v<n:adj[u].append(v);adj[v].append(u)
    seen={0};q=[0]
    for x in q:
        for y in adj[x]:
            if y not in seen:seen.add(y);q.append(y)
    return len(seen)==n,seen,adj
def shortest_hops(n,edges,src=0):
    ok,_,adj=connected(n,edges); d=[10**9]*n;d[src]=0;q=deque([src])
    while q:
        u=q.popleft()
        for v in adj[u]:
            if d[v]>d[u]+1:d[v]=d[u]+1;q.append(v)
    return d
def shortest_cost(n,edges,pen,src=0):
    import heapq
    adj=[[] for _ in range(n)]
    for u,v,m,c in edges:
        if 0<=u<v<n:adj[u].append((v,m));adj[v].append((u,m))
    d=[10**30]*n;d[src]=0;h=[(0,src)]
    while h:
        du,u=heapq.heappop(h)
        if du!=d[u]:continue
        for v,m in adj[u]:
            nd=du+m+pen[v]
            if nd<d[v]:d[v]=nd;heapq.heappush(h,(nd,v))
    return d
def shortest_hop_cost(n,edges,pen,src=0):
    """Minimum cost among paths that use the minimum possible number of roads."""
    adj=[[] for _ in range(n)]
    for u,v,m,c in edges:
        if 0<=u<v<n: adj[u].append((v,m));adj[v].append((u,m))
    hops=shortest_hops(n,edges,src); inf=10**30; dp=[inf]*n; dp[src]=0
    order=sorted(range(n),key=lambda x:hops[x])
    for u in order:
        if dp[u]>=inf: continue
        for v,m in adj[u]:
            if hops[v]==hops[u]+1: dp[v]=min(dp[v],dp[u]+m+pen[v])
    return dp
def knap(items,cap):
    dp=[0]*(cap+1)
    for w,v in items:
        for c in range(cap,w-1,-1):dp[c]=max(dp[c],dp[c-w]+v)
    return max(dp)
def sched(wins,weighted=False):
    a=sorted(wins,key=lambda z:(z[2],z[1],z[0]));dp=[0]*(len(a)+1)
    for i,(wid,s,f,v) in enumerate(a,1):
        p=0
        for j in range(i-1,-1,-1):
            if j==0 or a[j-1][2]<=s:p=j;break
        dp[i]=max(dp[i-1],dp[p]+(v if weighted else 1))
    return dp[-1]
def rod(pr,L):
    d=[0]*(L+1)
    for x in range(1,L+1):d[x]=max(pr[j-1]+d[x-j] for j in range(1,x+1))
    return d[L]
def tourlen(sites,tour):
    return sum(abs(sites[tour[i]][0]-sites[tour[(i+1)%len(tour)]][0])+abs(sites[tour[i]][1]-sites[tour[(i+1)%len(tour)]][1]) for i in range(len(tour)))
def tsp(sites):
    n=len(sites);d=[[abs(sites[i][0]-sites[j][0])+abs(sites[i][1]-sites[j][1]) for j in range(n)] for i in range(n)];dp={(1,0):0}
    for mask in range(1<<n):
        if not mask&1:continue
        for j in range(n):
            if (mask,j) not in dp:continue
            for k in range(1,n):
                if not mask>>k&1:
                    nm=mask|1<<k;v=dp[(mask,j)]+d[j][k]
                    if v<dp.get((nm,k),10**30):dp[(nm,k)]=v
    full=(1<<n)-1
    return min(dp[(full,j)]+d[j][0] for j in range(1,n))
def cover_min(n,edges):
    for k in range(n+1):
        for c in itertools.combinations(range(n),k):
            s=set(c)
            if all(u in s or v in s for u,v in edges):return k
    return n
def nearest_tour(sites):
    tour=[0];left=set(range(1,len(sites)));cur=0
    while left:
        x=min(left,key=lambda j:(abs(sites[cur][0]-sites[j][0])+abs(sites[cur][1]-sites[j][1]),j));tour.append(x);left.remove(x);cur=x
    return tourlen(sites,tour)
def rows_by_case(wb):
    out={}
    for sh in HEADERS:
        d=defaultdict(list)
        for row in wb[sh].iter_rows(min_row=2,values_only=True):
            if row and row[0] is not None:d[row[0]].append(tuple(row))
        out[sh]=d
    return out

def validate(path, report=None):
    errors=[]
    def err(tag,cid,msg):errors.append(f'{tag} {cid}: {msg}')
    wb=load_workbook(path,read_only=True,data_only=True)
    # V1 exact sheet order and headers
    if wb.sheetnames!=list(HEADERS):errors.append(f'V1: workbook must contain exactly these 20 sheets in order; got {wb.sheetnames}')
    for sh,head in HEADERS.items():
        if sh not in wb.sheetnames:continue
        got=[x.value for x in next(wb[sh].iter_rows(min_row=1,max_row=1))]
        if got!=head:errors.append(f'V1: {sh} header mismatch: {got}')
    D=rows_by_case(wb); cases=list(wb['cases'].iter_rows(min_row=2,values_only=True)) if 'cases' in wb.sheetnames else []
    if len(cases)!=50:errors.append(f'V1: expected 50 cases, found {len(cases)}')
    case_map={r[0]:r for r in cases if r and r[0] is not None}
    if len(case_map)!=50:errors.append('V1: case_id values must be distinct')
    if sorted(r[2] for r in cases if r and len(r)>2 and isinstance(r[2],int))!=list(range(50)):errors.append('V1: case_no values must be exactly 0..49')
    seeds=set()
    for row in cases:
        if not row or len(row)<15:continue
        cid,seed,k,t,sc=row[:5];seeds.add(seed)
        if not isinstance(k,int) or not 0<=k<50:err('V1',cid,'case_no must be 0..49');continue
        if cid!=f'{seed}-{k:02d}':err('V1',cid,'case_id format does not match seed and case_no')
        if t!=tier(k):err('V1',cid,'incorrect tier')
        exp=(k%6)+1 if k<46 else {46:5,47:6,48:1,49:2}[k]
        if sc!=exp:err('V1',cid,'incorrect scenario')
    if len(seeds)>1:errors.append('V1: group_seed must be identical for all cases')
    # V2-V8 per case
    module_fingerprints=defaultdict(dict)
    for cid,c in case_map.items():
        seed,k,t,sc=c[1],c[2],c[3],c[4];base='L' if t=='X' else t
        b=RANGES[base];cmin,cmax,vmin,vmax,lmin,lmax,wmin,wmax,amin,amax,smin,smax,pmin,pmax,dmin,dmax,nmin,nmax,svmin,svmax,rmin,rmax=b
        tab={sh:D[sh].get(cid,[]) for sh in HEADERS if sh!='cases'}
        # V2 size and field ranges
        limits={'crates':(cmin,cmax),'locations':(vmin,vmax),'windows':(wmin,wmax),'incidents':(amin,amax),'sites':(smin,smax),'pipeline_dims':(pmin+1,pmax+1),'drum_prices':(dmin,dmax),'sensor_edges':(nmin-1,nmax*(nmax-1)//2),'safety_planted':(svmin,svmax),'rations':(rmin,rmax)}
        if k==46: limits['crates']=(1500,3000)
        if k==47: limits['crates']=(12,60)
        if k==48: limits['locations']=(1500,3000)
        for sh,(lo,hi) in limits.items():
            n=len(tab[sh])
            if not lo<=n<=hi:err('V2',cid,f'{sh} row count {n} outside {lo}..{hi}')
        if not 2<=len(tab['waves'])<=5:err('V2',cid,'closure waves must contain 2..5 distinct wave groups')
        if not max(1,len(tab['locations'])-1)<=len(tab['roads'])<=3*max(1,len(tab['locations'])):err('V2',cid,'road count outside V-1..3V')
        crates=tab['crates'];locs=tab['locations'];roads=tab['roads'];wins=tab['windows'];inc=tab['incidents'];sitesr=tab['sites'];
        for x in crates:
            if not (isinstance(x[1],int) and x[2] in range(5) and 1<=x[3]<=20 and 10<=x[4]<=100 and 0<=x[5]<=999 and 10_000_000<=x[6]<=99_999_999):err('V2',cid,f'invalid crate values {x[1:]}')
        if len({x[6] for x in crates})!=len(crates):err('V3',cid,'crate serials are not unique')
        if [x[1] for x in crates]!=list(range(1,len(crates)+1)):err('V3',cid,'crate_id must run consecutively from 1')
        if len(locs):
            if [x[1] for x in locs]!=list(range(len(locs))):err('V3',cid,'loc_id must run 0..V-1')
            if locs[0][2]!='D' or locs[0][3]!=0:err('V2',cid,'location 0 must be HQ type D penalty 0')
            if not any(x[2]=='S' for x in locs) or not any(x[2]=='H' for x in locs):err('V2',cid,'must contain at least one shelter and one hospital')
            if any(x[2] not in ('D','S','H','J') or x[3] not in (0,1,2,4,6) for x in locs):err('V2',cid,'invalid location type/penalty')
        V=len(locs);edge_list=[];pairset=set()
        for x in roads:
            _,rid,u,v,m,cb=x
            if not (isinstance(u,int) and isinstance(v,int) and 0<=u<v<V and 1<=m<=60 and 1<=cb<=100):
                err('V2',cid,f'invalid road {rid}'); continue
            if (u,v) in pairset:err('V3',cid,'duplicate road endpoint pair')
            pairset.add((u,v));edge_list.append((u,v,m,cb))
        if [x[1] for x in roads]!=list(range(len(roads))):err('V3',cid,'road_no must be consecutive from 0')
        ok,_,_=connected(V,edge_list)
        if V and not ok:err('V4',cid,'network is not connected before closures')
        # targets
        targets=c[6:9]
        if len(set(targets))!=3 or any(not isinstance(x,int) or x<=0 or x>=V or locs[x][2] not in ('S','H') for x in targets):err('V3',cid,'targets must be three distinct shelter/hospital IDs, not HQ')
        totalw=sum(x[3] for x in crates)
        if not totalw*.10<=c[5]<=totalw*.70 or c[5]<1:err('V2',cid,'capacity must be 10%-70% of total crate weight')
        # waves V3/V4
        waveids=set();closed=[];bywave=defaultdict(list)
        for x in tab['waves']:
            _,wi,rid=x
            if rid<0 or rid>=len(roads):err('V3',cid,'wave references missing road')
            if rid in closed:err('V4',cid,'a road occurs in more than one closure wave')
            closed.append(rid);bywave[wi].append(rid);waveids.add(wi)
        if waveids and sorted(waveids)!=list(range(max(waveids)+1)):err('V3',cid,'wave_no must be consecutive from 0')
        if any(not 1<=len(v)<=max(1,len(roads)//4) for v in bywave.values()):err('V4',cid,'wave size outside 1..floor(E/4)')
        # radio V5
        logrows=tab['radio_log'];
        if [x[1] for x in logrows]!=list(range(1,len(logrows)+1)):err('V5',cid,'radio parts must be consecutive and ordered')
        if any(len(x[2])>10000 or not set(x[2])<=ALPH for x in logrows):err('V5',cid,'radio part exceeds 10000 or uses invalid alphabet')
        log=''.join(x[2] for x in sorted(logrows,key=lambda z:z[1]))
        if not lmin<=len(log)<=lmax:err('V2',cid,'joined radio log length outside tier range')
        pats=sorted(tab['radio_patterns'],key=lambda z:z[1])
        if len(pats)!=6 or [x[1] for x in pats]!=list(range(6)) or [len(x[2]) for x in pats]!=PATLEN:err('V5',cid,'must have six patterns of lengths 2,4,8,12,20,40')
        elif any(not set(x[2])<=ALPH for x in pats):err('V5',cid,'radio pattern contains invalid alphabet symbols')
        else:
            for x in pats:
                if x[3]!=countpat(log,x[2]):err('V5',cid,f'pattern {x[1]} occurrence count incorrect')
            counts=[x[3] for x in pats]
            if sum(x>0 for x in counts)<2 or not any(x==0 for x in counts):err('V5',cid,'patterns require at least two present and one absent')
        if len(tab['radio_reports'])!=1:err('V5',cid,'exactly one radio_reports row required')
        else:
            ra,rb=tab['radio_reports'][0][1:]
            report_min,report_max=(10,30) if t=='S' else (30,150) if t=='M' else (200,1000)
            if not(report_min<=len(ra)<=report_max and report_min<=len(rb)<=report_max):err('V2',cid,'radio report lengths outside tier range')
            if not set(ra+rb)<=ALPH:err('V5',cid,'reports contain invalid alphabet')
            true=edit_distance(ra,rb)
            if c[9]<true or c[9]>len(ra)+len(rb):err('V5',cid,'edit_bound is not between true edit distance and total report length')
        # windows and touch-pairs
        for x in wins:
            if not (0<=x[2]<x[3]<=1440 and 5<=x[3]-x[2]<=480 and 1<=x[4]<=100):err('V2',cid,f'invalid window {x[1]}')
        if [x[1] for x in wins]!=list(range(1,len(wins)+1)):err('V3',cid,'window_id must be consecutive')
        touch=sum(1 for i,a in enumerate(wins) for b2 in wins[i+1:] if a[3]==b2[2] or b2[3]==a[2])
        if len(wins)>=8 and touch<2:err('V7',cid,'at least two touching window pairs required')
        # incidents
        if [x[1] for x in inc]!=list(range(1,len(inc)+1)):err('V3',cid,'alarm_id must be consecutive')
        if any(not 1<=x[3]<=10 for x in inc) or any(inc[i][2]>inc[i+1][2] for i in range(len(inc)-1)) or (inc and inc[0][2]!=0):err('V2',cid,'incident ticks/priorities invalid or ticks not nondecreasing from 0')
        if t in ('M','L','X') and inc:
            rate=len(inc)/(inc[-1][2]+1)
            if not 1.2<=rate<=1.8 or len(set(x[3] for x in inc))<6 or 1 not in [x[3] for x in inc]:err('V2',cid,'M/L alarms-per-tick rate or priority diversity invalid')
        # sites and tour V6
        if [x[1] for x in sitesr]!=list(range(len(sitesr))) or len({(x[2],x[3]) for x in sitesr})!=len(sitesr) or any(not(0<=x[2]<=999 and 0<=x[3]<=999) for x in sitesr):err('V3',cid,'site IDs/coordinates invalid or duplicated')
        sitecoords=[(x[2],x[3]) for x in sitesr];tr=sorted(tab['tour'],key=lambda z:z[1]);tour=[x[2] for x in tr]
        if len(tour)!=len(sitesr) or tour!=list(range(len(sitesr))) and (len(set(tour))!=len(sitesr) or set(tour)!=set(range(len(sitesr)))):err('V6',cid,'tour must be a permutation of all sites')
        if tour and tour[0]!=0:err('V6',cid,'tour must start at site 0')
        if tour:
            length=tourlen(sitecoords,tour)
            if length!=c[10]:err('V6',cid,'cases.tour_length does not equal recomputed closed Manhattan tour length')
            if t in ('L','X') and length>nearest_tour(sitecoords):err('V6',cid,'L/X planted tour is longer than nearest-unvisited tour')
        if [x[1] for x in tr]!=list(range(len(tr))):err('V3',cid,'tour step must be consecutive')
        # pipeline and drum
        pd=sorted(tab['pipeline_dims'],key=lambda z:z[1]);dr=sorted(tab['drum_prices'],key=lambda z:z[1])
        if [x[1] for x in pd]!=list(range(len(pd))) or any(not 2<=x[2]<=100 for x in pd):err('V2',cid,'pipeline dimensions invalid')
        if [x[1] for x in dr]!=list(range(1,len(dr)+1)) or any(not 1<=x[2]<=500 for x in dr):err('V2',cid,'drum prices invalid')
        if len(dr)>1 and len({x[2]/x[1] for x in dr})==1:err('V2',cid,'drum prices cannot all be proportional to length')
        # sensor cover
        sn=c[11];sk=c[12];se=tab['sensor_edges'];cover=[x[1] for x in tab['sensor_cover']];spairs=[]
        if not (nmin<=sn<=nmax and 2<=sk<=sn//2):err('V2',cid,'sensor_nodes or sensor_k outside tier/definition range')
        for x in se:
            u,v=x[1:]
            if not 0<=u<v<sn:err('V3',cid,'invalid sensor edge endpoint')
            spairs.append((u,v))
        if len(spairs)<sn-1 or len(set(spairs))!=len(spairs):err('V2',cid,'sensor graph needs n-1 edges and unique edges')
        if len(cover)!=sk or len(set(cover))!=sk or any(not 0<=x<sn for x in cover):err('V6',cid,'planted sensor cover must contain exactly sensor_k distinct valid nodes')
        elif any(u not in cover and v not in cover for u,v in spairs):err('V6',cid,'planted sensor cover leaves an edge uncovered')
        # safety assignment and clauses
        sv=c[13];pl=tab['safety_planted'];ass={x[1]:x[2] for x in pl};rules=tab['safety_rules']
        if len(ass)!=sv or set(ass)!=set(range(1,sv+1)) or any(v not in (0,1) for v in ass.values()):err('V6',cid,'safety assignment must define each switch exactly once')
        if not 2*sv<=len(rules)<=6*sv:err('V2',cid,'safety rule/switch ratio outside 2.0..6.0')
        if [x[1] for x in rules]!=list(range(len(rules))):err('V3',cid,'safety rule_no must be consecutive from 0')
        for x in rules:
            lits=x[2:5]
            if len(set(abs(z) for z in lits))!=3 or any(z==0 or abs(z)>sv for z in lits):err('V2',cid,'each safety clause needs three different valid nonzero switches')
            elif ass and not any((ass[abs(z)]==1 if z>0 else ass[abs(z)]==0) for z in lits):err('V6',cid,f'planted assignment violates safety rule {x[1]}')
        # rations
        rat=tab['rations'];
        if [x[1] for x in rat]!=list(range(len(rat))) or any(not 1<=x[2]<=1000 or x[3] not in (0,1) for x in rat):err('V2',cid,'ration item/value/planted fields invalid')
        if sum(x[2] for x in rat if x[3]==1)!=c[14] or c[14]<=0:err('V6',cid,'planted ration values do not sum to rations_target')
        # V7 signatures for ordinary cases
        if k<46:
            avg=2*len(roads)/V if V else 0; cats=[x[2] for x in crates]; penalties=sum(x[3]>=2 for x in locs)/max(1,V)
            if sc==1 and not(2.2<=avg<=3.2 and penalties>=.40 and len(waveids)>=3 and sum(x in (0,1) for x in cats)/len(cats)>=.55):err('V7',cid,'Delta Flood signature failed')
            if sc==2 and not(2.0<=avg<=2.6 and sum(x in (2,3) for x in cats)/len(cats)>=.45 and sum(x[3] for x in inc)/len(inc)>=6):err('V7',cid,'Mountain Earthquake signature failed')
            if sc==3:
                alphabet=set(log);overlap=sum(1 for i,a in enumerate(wins) for z in wins[i+1:] if max(a[2],z[2])<min(a[3],z[3]))/max(1,len(wins))
                if sum(x[2]=='S' for x in locs)/V<.25 or len(waveids)<3 or overlap<2 or len(alphabet)<30:err('V7',cid,'Coastal Cyclone signature failed')
            if sc==4:
                if avg<3.5 or sum(x[2]==4 for x in crates)/len(crates)<.10:err('V7',cid,'Urban Flash Flood density/rare crate signature failed')
                if t in ('M','L') and not(1.2<=len(inc)/(inc[-1][2]+1)<=1.8 and len(set(x[3] for x in inc))>=6 and 1 in [x[3] for x in inc]):err('V7',cid,'Urban Flash Flood alarm rate/priorities failed')
            if sc==5:
                deg=Counter(z for e in roads for z in e[2:4]);dist=shortest_hops(V,edge_list)
                rem=[x[2] for x in tab['waves']];_,reach,_=connected(V,edge_list,rem)
                if avg>2.4 or max(deg.values(),default=0)>4 or max(dist,default=0)<V//4 or len(waveids)<3 or all(x in reach for x in targets):err('V7',cid,'Landslide Corridor signature failed')
            if sc==6:
                if sum(x[2] in (1,2) for x in crates)/len(crates)<.55 or sum(x[5]<=300 for x in crates)/len(crates)<.5 or max(Counter(log).values(),default=0)/len(log)<.25:err('V7',cid,'Outbreak Camp signature failed')
        # V5 claims on S recomputed independently
        cl=tab['claims']
        if t=='S':
            if len(cl)!=1:err('V5',cid,'tier S requires exactly one claims row')
            else:
                x=cl[0];expected=[knap([(z[3],z[4]) for z in crates],c[5]),sched([[z[1],z[2],z[3],z[4]] for z in wins]),sched([[z[1],z[2],z[3],z[4]] for z in wins],True),rod([z[2] for z in dr],len(dr)),tsp(sitecoords),cover_min(sn,spairs)]
                if list(x[1:])!=expected:err('V5',cid,f'S-tier claims mismatch; expected {expected}, got {list(x[1:])}')
        elif cl:err('V5',cid,'claims must be empty outside tier S')
        # fingerprints for distinct modules of >=8 records
        for sh,rr in tab.items():
            if len(rr)>=8:
                canon=tuple(tuple(z[1:]) for z in rr)
                module_fingerprints[sh][cid]=hashlib.sha256(repr(canon).encode()).hexdigest()
    # V7 global variety
    s_cases=[r for r in cases if r[3]=='S']; greedy_bad=0; route_bad=0
    for c in s_cases:
        cid=c[0];cr=D['crates'][cid];cap=c[5];items=sorted(cr,key=lambda x:(-(x[4]/max(1,x[3])),x[1]));w=v=0
        for x in items:
            if w+x[3]<=cap:w+=x[3];v+=x[4]
        opt=D['claims'][cid][0][1]
        if v<opt:greedy_bad+=1
        loc=D['locations'][cid];roads=D['roads'][cid];V=len(loc);ed=[(x[2],x[3],x[4],x[5]) for x in roads];hop=shortest_hops(V,ed);cost=shortest_cost(V,ed,[x[3] for x in loc])
        fewcost=shortest_hop_cost(V,ed,[x[3] for x in loc])
        if any(hop[z]>0 and cost[z]<fewcost[z] for z in c[6:9]):route_bad+=1
    if greedy_bad<5:errors.append(f'V7: only {greedy_bad} tier-S cases defeat value/kg greedy; need at least 5')
    if route_bad<5:errors.append(f'V7: only {route_bad} tier-S cases have a target where fewest-road route is costlier; need at least 5')
    ml=[r for r in cases if r[3] in ('M','L')]; ratios=[len(D['safety_rules'][r[0]])/r[13] for r in ml]
    if sum(x<3.5 for x in ratios)<5 or sum(3.5<=x<=4.8 for x in ratios)<8 or sum(x>4.8 for x in ratios)<5:errors.append('V7: safety-rule ratio distribution across M/L cases does not meet 5/8/5 buckets')
    for sh,fp in module_fingerprints.items():
        vals=list(fp.values())
        if len(vals)!=len(set(vals)):errors.append(f'V7: duplicate case data detected in module {sh} (module size >=8)')
    # V8 trap checks
    def mono(a):return all(a[i]<=a[i+1] for i in range(len(a)-1)) or all(a[i]>=a[i+1] for i in range(len(a)-1))
    # 46
    c46='24-46'
    # locate by case_no instead of assuming seed
    byno={r[2]:r[0] for r in cases if r and len(r)>2}
    for kn in (46,47,48,49):
        cid=byno.get(kn)
        if not cid:continue
        c=case_map[cid];cr=D['crates'][cid];loc=D['locations'][cid];road=D['roads'][cid];log=''.join(x[2] for x in sorted(D['radio_log'][cid],key=lambda z:z[1]))
        if kn==46:
            if not 1500<=len(cr)<=3000:err('V8',cid,'trap 46 crate count outside 1500..3000')
            fields=[[x[4] for x in cr],[x[3] for x in cr],[x[5] for x in cr],[x[6] for x in cr],[x[4]/max(1,x[3]) for x in cr]]
            if sum(mono(f) for f in fields)<3:err('V8',cid,'trap 46 fewer than three monotone crate fields')
        if kn==47:
            if not 12<=len(cr)<=60:err('V8',cid,'trap 47 crate count outside 12..60')
            cap=c[5];w=v=0
            for x in sorted(cr,key=lambda z:(-z[4]/z[3],z[1])):
                if w+x[3]<=cap:w+=x[3];v+=x[4]
            opt=knap([(x[3],x[4]) for x in cr],cap)
            if opt<=0 or v>0.60*opt:err('V8',cid,f'trap 47 greedy value {v} exceeds 60% of optimum {opt}')
        if kn==48:
            if not 1500<=len(loc)<=3000:err('V8',cid,'trap 48 location count outside 1500..3000')
            ed=[(x[2],x[3],x[4],x[5]) for x in road];hop=shortest_hops(len(loc),ed);cost=shortest_cost(len(loc),ed,[x[3] for x in loc])
            if max(hop)<500:err('V8',cid,'trap 48 graph diameter/farthest shortest distance below 500 roads')
            fewcost=shortest_hop_cost(len(loc),ed,[x[3] for x in loc])
            if not any(fewcost[z]>=1.5*cost[z] and cost[z]<10**30 for z in c[6:9]):err('V8',cid,'trap 48 has no target with fewest-road route at least 1.5x cheapest route')
        if kn==49:
            if not 50000<=len(log)<=100000:err('V8',cid,'trap 49 log length outside 50000..100000')
            # longest prefix repetition block <=8; exact repetition fraction
            block=max((log[:i] for i in range(1,9) if log[:i]),key=lambda x:sum(log[j]==x[j%len(x)] for j in range(len(log))),default='A')
            rep=sum(log[j]==block[j%len(block)] for j in range(len(log)))/len(log)
            if rep<.90:err('V8',cid,'trap 49 less than 90% of log follows a repeated block <=8')
            pats=sorted(D['radio_patterns'][cid],key=lambda z:z[1]);good=0
            for x in pats:
                p=x[2]
                if len(p)>1:
                    # number of start positions where p[:-1] matches
                    hits=countpat(log,p[:-1])
                    if hits>=.5*max(1,len(log)-len(p)+2):good+=1
            if good<4:err('V8',cid,f'trap 49 only {good} patterns have a matching prefix at half of start positions')
    wb.close()
    result='PASS: all 50 cases satisfy V1-V8.' if not errors else 'FAIL\n'+'\n'.join(errors)
    if report:Path(report).write_text(result+'\n',encoding='utf-8')
    return errors,result
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('workbook',nargs='?',default='cases.xlsx');ap.add_argument('--report',default='validation_report.txt');a=ap.parse_args()
    e,r=validate(a.workbook,a.report);print(r);raise SystemExit(1 if e else 0)
