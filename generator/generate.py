#!/usr/bin/env python3
"""Deterministic LIFELINE Iteration-1 synthetic dataset generator."""
from __future__ import annotations
import argparse, random, string, math
from collections import defaultdict
from pathlib import Path
from openpyxl import Workbook

SHEETS={
'cases':['case_id','group_seed','case_no','tier','scenario','capacity','target_0','target_1','target_2','edit_bound','tour_length','sensor_nodes','sensor_k','safety_vars','rations_target'],
'crates':['case_id','crate_id','cat','weight','value','expiry','serial'],
'locations':['case_id','loc_id','type','penalty'],
'roads':['case_id','road_no','u','v','minutes','cable'],
'waves':['case_id','wave_no','road_no'],
'radio_log':['case_id','part','text'],
'radio_patterns':['case_id','pattern_no','pattern','count'],
'radio_reports':['case_id','report_a','report_b'],
'windows':['case_id','window_id','start','finish','value'],
'incidents':['case_id','alarm_id','tick','priority'],
'sites':['case_id','site_no','x','y'],
'tour':['case_id','step','site_no'],
'pipeline_dims':['case_id','idx','dim'],
'drum_prices':['case_id','length','price'],
'sensor_edges':['case_id','u','v'],
'sensor_cover':['case_id','node'],
'safety_rules':['case_id','rule_no','lit_1','lit_2','lit_3'],
'safety_planted':['case_id','switch','value'],
'rations':['case_id','item','value','planted'],
'claims':['case_id','cargo_opt','window_count_opt','window_value_opt','drum_opt','tour_opt','cover_min']}
ALPH=string.ascii_uppercase+string.digits+'_-'
LENS=[2,4,8,12,20,40]
RANGES={'S':(8,20,6,12,200,500,8,20,10,40,5,9,4,8,5,14,6,14,6,15,8,20),
'M':(40,200,30,150,2000,10000,50,300,200,1000,10,16,15,60,40,200,20,40,20,50,24,40),
'L':(500,1500,500,2000,20000,100000,1000,3000,2000,5000,200,500,100,300,500,2000,100,400,20,50,50,100)}
def tier(k): return 'S' if k<20 else 'M' if k<40 else 'L' if k<46 else 'X'
def edit_distance(a,b):
    p=list(range(len(b)+1))
    for i,x in enumerate(a,1):
        q=[i]
        for j,y in enumerate(b,1): q.append(min(q[-1]+1,p[j]+1,p[j-1]+(x!=y)))
        p=q
    return p[-1]
def countpat(s,p):
    n=0; start=0
    while True:
        i=s.find(p,start)
        if i<0:return n
        n+=1;start=i+1
def knap(items,cap):
    dp=[0]*(cap+1)
    for w,v in items:
        for c in range(cap,w-1,-1): dp[c]=max(dp[c],dp[c-w]+v)
    return max(dp)
def schedule(wins,by_value=False):
    arr=sorted(wins,key=lambda x:(x[2],x[1],x[0]))
    dp=[0]*(len(arr)+1)
    for i,(wid,s,f,val) in enumerate(arr,1):
        j=0
        for z in range(i-1,-1,-1):
            if z==0 or arr[z-1][2]<=s: j=z;break
        dp[i]=max(dp[i-1],dp[j]+(val if by_value else 1))
    return dp[-1]
def rod(prices,L):
    dp=[0]*(L+1)
    for j in range(1,L+1): dp[j]=max(prices[x-1]+dp[j-x] for x in range(1,j+1))
    return dp[L]
def tsp_exact(sites):
    n=len(sites); d=[[abs(sites[i][2]-sites[j][2])+abs(sites[i][3]-sites[j][3]) for j in range(n)] for i in range(n)]
    dp={(1,0):0}
    for mask in range(1<<n):
        if not mask&1: continue
        for j in range(n):
            if (mask,j) not in dp: continue
            for k in range(1,n):
                if not mask>>k&1:
                    nm=mask|1<<k; val=dp[(mask,j)]+d[j][k]
                    if val<dp.get((nm,k),10**30): dp[(nm,k)]=val
    full=(1<<n)-1
    return min(dp[(full,j)]+d[j][0] for j in range(1,n))
def cover_min(n,edges):
    for k in range(n+1):
        from itertools import combinations
        for c in combinations(range(n),k):
            s=set(c)
            if all(u in s or v in s for u,v in edges): return k
    return n

def make_case(seed,k):
    r=random.Random(seed*1000003+k*9176+24); t=tier(k); b=RANGES['L' if t=='X' else t]
    cmin,cmax,vmin,vmax,lmin,lmax,wmin,wmax,amin,amax,smin,smax,pmin,pmax,dmin,dmax,nmin,nmax,svmin,svmax,rmin,rmax=b
    scenario=(k%6)+1 if k<46 else {46:5,47:6,48:1,49:2}[k]; cid=f'{seed}-{k:02d}'
    nc=r.randint(cmin,cmax); nv=r.randint(vmin,vmax); nlog=r.randint(lmin,lmax); nw=r.randint(wmin,wmax); na=r.randint(amin,amax); ns=r.randint(smin,smax); np=r.randint(pmin,pmax); dl=r.randint(dmin,dmax); sn=r.randint(nmin,nmax); sv=r.randint(svmin,svmax); nr=r.randint(rmin,rmax)
    if k==46: nc=r.randint(1500,3000)
    if k==47: nc=13
    if k==48: nv=r.randint(1500,3000)
    if k==49: nlog=r.randint(50000,100000)
    # crate records
    crates=[]
    if k==46:
        for i in range(1,nc+1):
            val=10+(90*(i-1)//(nc-1)); exp=999*(i-1)//(nc-1); serial=10_000_000+i
            crates.append([cid,i,(i%5),10,val,exp,serial])
    elif k==47:
        base=[(11,100),(10,90),(10,90)]+[(10,10)]*10
        for i,(w,v) in enumerate(base,1): crates.append([cid,i,(i-1)%5,w,v,r.randint(0,999),10_000_000+i])
    else:
        cats=[0,1,2,3,4]
        if scenario==1: cats=[0,1,0,1,2,3,4]
        elif scenario==2: cats=[2,3,2,3,0,1,4]
        elif scenario==4: cats=[4,4,0,1,2,3,4,0,1]
        elif scenario==6: cats=[2,1,2,1,0,3,4]
        serials=r.sample(range(10_000_000,100_000_000),nc)
        for i in range(1,nc+1):
            cat=cats[(i-1)%len(cats)] if scenario in (1,2,4,6) else r.randrange(5)
            w=r.randint(1,20); v=r.randint(10,100); exp=r.randint(0,999)
            if scenario==6 and i<=math.ceil(nc/2): exp=r.randint(0,300)
            crates.append([cid,i,cat,w,v,exp,serials[i-1]])
        # enforce category shares by reassignment, preserving all other fields
        need={1:{0,1},2:{2,3},4:{4},6:{1,2}}.get(scenario)
        frac={1:.56,2:.46,4:.12,6:.56}.get(scenario)
        if need:
            q=math.ceil(nc*frac)
            for i in range(q): crates[i][2]=sorted(need)[i%len(need)]
            if scenario==6:
                for i in range(math.ceil(nc/2)): crates[i][5]=min(crates[i][5],300)
    totalw=sum(x[3] for x in crates)
    capacity=max(1,int(totalw*(.285 if k==47 else r.uniform(.2,.5))))
    if k==47: capacity=20
    # location types/penalties
    if k==48: nv=max(nv,1500)
    locs=[]
    for i in range(nv):
        typ='D' if i==0 else ('S' if i%3==1 else 'H' if i%3==2 else 'J')
        pen=0 if i==0 else r.choice([0,1,2,4,6])
        if scenario==1 and i>0: pen=2 if i%2 else r.choice([2,4,6])
        locs.append([cid,i,typ,pen])
    # Graph shape: path guarantees connectivity; add controlled chords.
    pairs={(i,i+1) for i in range(nv-1)}
    if k in (0,2,3,5,6,8): pairs.add((0,2))
    if k==48:
        pairs.add((0,2))
    else:
        if scenario==4: targetE=min(3*nv,max(nv-1,math.ceil(1.9*nv)))
        elif scenario==5: targetE=min(3*nv,max(nv-1,math.floor(1.15*nv)))
        elif scenario==2: targetE=min(3*nv,max(nv-1,math.floor(1.20*nv)))
        elif scenario in (1,2): targetE=min(3*nv,max(nv-1,math.floor(1.35*nv)))
        else: targetE=min(3*nv,max(nv-1,math.floor(1.45*nv)))
        # For case study 5 cap degree 4; path plus nonadjacent short chords.
        deg=[0]*nv
        for u,v in pairs: deg[u]+=1;deg[v]+=1
        tries=0
        while len(pairs)<targetE and tries<targetE*60:
            tries+=1
            if scenario==5:
                u=r.randint(1,nv-3); v=u+2
            else: u,v=sorted(r.sample(range(nv),2))
            if (u,v) in pairs: continue
            if scenario==5 and (deg[u]>=4 or deg[v]>=4): continue
            pairs.add((u,v));deg[u]+=1;deg[v]+=1
    # make target locations suitable, target 2 has deliberate cheap detour in trap 48
    locs[1][2]='S';locs[2][2]='H';locs[3][2]='S'
    if k==48 or k in (0,2,3,5,6,8): locs[1][3]=0;locs[2][3]=0
    roads=[]
    for j,(u,v) in enumerate(sorted(pairs)):
        mins= r.randint(1,60); cable=r.randint(1,100)
        if (k==48 or k in (0,2,3,5,6,8)) and (u,v)==(0,2): mins=60; cable=100
        if (k==48 or k in (0,2,3,5,6,8)) and (u,v) in ((0,1),(1,2)): mins=1
        roads.append([cid,j,u,v,mins,cable])
    # waves: disjoint, 3+ for studies 1,3,5. For study 5 close bridge path edges.
    nwaves=r.randint(3,5) if scenario in (1,3,5) else r.randint(2,5)
    maxwave=max(1,len(roads)//4); chosen=[]
    candidates=list(range(len(roads)))
    if scenario==5:
        chosen=[next(i for i,x in enumerate(roads) if x[2]==j and x[3]==j+1) for j in (0,2,4)]
    elif k==48: chosen=[next(i for i,x in enumerate(roads) if x[2]==j and x[3]==j+1) for j in (100,200,300)]
    else:
        r.shuffle(candidates); chosen=candidates[:min(nwaves,maxwave*nwaves)]
    waves=[]
    for z,rid in enumerate(chosen): waves.append([cid,z%nwaves,rid])
    # ensure each of required waves represented (3+)
    for wi in range(nwaves):
        if not any(x[1]==wi for x in waves):
            rid=next(x for x in candidates if x not in chosen); chosen.append(rid);waves.append([cid,wi,rid])
    # radio
    if k==49:
        block='A'; log=block*nlog
        patterns=[]
        for i,L in enumerate([2,4,8,12,20,40]):
            p=(block*((L+7)//8))[:L]
            if L==40: p=p[:-1]+('X' if p[-1]!='X' else 'Y')
            patterns.append([cid,i,p,countpat(log,p)])
    else:
        if scenario==3:
            log=''.join(ALPH[i%len(ALPH)] for i in range(len(ALPH))) + ''.join(r.choice(ALPH) for _ in range(nlog-len(ALPH)))
        elif scenario==6:
            dominant='A'; log=dominant*math.ceil(nlog*.30)+''.join(r.choice(ALPH) for _ in range(nlog-math.ceil(nlog*.30)))
        else: log=''.join(r.choice(ALPH) for _ in range(nlog))
        patterns=[]
        for i,L in enumerate(LENS):
            p=log[(i*37)%(len(log)-L+1):(i*37)%(len(log)-L+1)+L]
            if i==5: p='Z'*L
            patterns.append([cid,i,p,countpat(log,p)])
    # guarantee 2 present + one absent for nontrap by setting first patterns from log, last absent
    patterns[0][2]=log[:2];patterns[0][3]=countpat(log,patterns[0][2])
    patterns[1][2]=log[:4];patterns[1][3]=countpat(log,patterns[1][2])
    patterns[5][2]='Z'*40;patterns[5][3]=0
    # trap49 needs first five periodic pattern prefix near-miss
    if k==49:
        for i,L in enumerate([2,4,8,12,20]):
            p='A'*L; patterns[i][2]=p;patterns[i][3]=countpat(log,p)
        patterns[5][2]='A'*39+'B';patterns[5][3]=0
    ra=''.join(r.choice(ALPH) for _ in range(min(r.randint(10 if t=='S' else 30 if t=='M' else 200,30 if t=='S' else 150 if t=='M' else 1000),nlog)))
    rb=ra[:];
    if len(rb)>1: rb=rb[:-1]+('Z' if rb[-1]!='Z' else 'Y')
    eb=edit_distance(ra,rb)
    # windows; cyclone highly overlapping; all cases include at least two touching pairs
    wins=[]
    for i in range(nw):
        if scenario==3: start=200; finish=600
        else: start=(i*13)%1200; finish=min(1440,start+r.randint(5,180))
        wins.append([cid,i+1,start,finish,r.randint(1,100)])
    if nw>=8:
        wins[0][2:4]=[0,300];wins[1][2:4]=[300,600];wins[2][2:4]=[50,350];wins[3][2:4]=[350,650]
    # incidents, ticks nondecreasing; M/L rates 1.2-1.8
    if t in ('M','L','X'): last=max(1,round(na/1.5)-1)
    else: last=max(1,na//2)
    incidents=[]
    if t in ('M','L','X'):
        if scenario==2:
            priorities=[1,6,7,8,9,10]+[r.randint(7,10) for _ in range(na-6)]
        else:
            priorities=[1,2,3,4,5,6]+[r.randint(1,10) for _ in range(na-6)]
        r.shuffle(priorities)
    else:
        priorities=[max(6,r.randint(1,10)) if scenario==2 else r.randint(1,10) for _ in range(na)]
    for i in range(na):
        tick=(i*(last+1))//na if t in ('M','L','X') else i//2
        incidents.append([cid,i+1,tick,priorities[i]])
    # sites and nearest-neighbour tour
    coords=set(); sites=[]
    while len(sites)<ns:
        p=(r.randint(0,999),r.randint(0,999))
        if p not in coords: coords.add(p);sites.append([cid,len(sites),*p])
    tour=[0]; un=set(range(1,ns)); cur=0
    while un:
        nxt=min(un,key=lambda j:(abs(sites[cur][2]-sites[j][2])+abs(sites[cur][3]-sites[j][3]),j));tour.append(nxt);un.remove(nxt);cur=nxt
    tourlen=sum(abs(sites[tour[i]][2]-sites[tour[(i+1)%ns]][2])+abs(sites[tour[i]][3]-sites[tour[(i+1)%ns]][3]) for i in range(ns))
    # pipeline and drum
    dims=[r.randint(2,100) for _ in range(np+1)]; prices=[r.randint(1,500) for _ in range(dl)]
    prices[0]=min(500,prices[0]+1)
    # Connected sensor tree whose every edge touches the planted vertex cover.
    sk=max(2,min(sn//2,r.randint(2,max(2,sn//2))))
    cover=sorted(r.sample(range(sn),sk)); sedges=[]
    for a,bn in zip(cover,cover[1:]): sedges.append([cid,a,bn])
    for node in range(sn):
        if node not in cover: sedges.append([cid,min(node,r.choice(cover)),max(node,r.choice(cover))])
    # Rebuild outside-node edges with one consistent selected cover endpoint.
    sedges=[[cid,a,bn] for a,bn in zip(cover,cover[1:])]+[[cid,min(node,(z:=r.choice(cover))),max(node,z)] for node in range(sn) if node not in cover]
    # safety rules ratio buckets for M/L cases; generate satisfied 3-literal clauses
    if 20<=k<=45:
        ix=k-20
        ratio=3.0 if ix<6 else 4.1 if ix<16 else 5.2
    else: ratio=3.0
    nrules=max(2*sv,min(6*sv,round(sv*ratio)))
    assignment=[r.randint(0,1) for _ in range(sv)]; rules=[]
    for j in range(nrules):
        vv=r.sample(range(1,sv+1),3); lits=[v if assignment[v-1] else -v for v in vv]
        rules.append([cid,j,*lits])
    # rations planted subset exactly equals target
    rationvals=[r.randint(1,1000) for _ in range(nr)]; planted_count=max(1,nr//3); target=sum(rationvals[:planted_count])
    rationrows=[[cid,i,val,1 if i<planted_count else 0] for i,val in enumerate(rationvals)]
    # tier S claims by exact/DP computations
    claim=[]
    if t=='S':
        claim=[cid,knap([(x[3],x[4]) for x in crates],capacity),schedule([[x[1],x[2],x[3],x[4]] for x in wins]),schedule([[x[1],x[2],x[3],x[4]] for x in wins],True),rod(prices,dl),tsp_exact(sites),cover_min(sn,[(x[1],x[2]) for x in sedges])] 
    rows={
      'cases':[[cid,seed,k,t,scenario,capacity,1,2,3,eb,tourlen,sn,sk,sv,target]],
      'crates':crates,'locations':locs,'roads':roads,'waves':waves,
      'radio_log':[[cid,i//10000+1,log[i:i+10000]] for i in range(0,len(log),10000)],
      'radio_patterns':patterns,'radio_reports':[[cid,ra,rb]],'windows':wins,'incidents':incidents,
      'sites':sites,'tour':[[cid,i,x] for i,x in enumerate(tour)],
      'pipeline_dims':[[cid,i,x] for i,x in enumerate(dims)],
      'drum_prices':[[cid,i+1,x] for i,x in enumerate(prices)],
      'sensor_edges':sedges,'sensor_cover':[[cid,x] for x in cover],
      'safety_rules':rules,'safety_planted':[[cid,i+1,x] for i,x in enumerate(assignment)],
      'rations':rationrows,'claims':([claim] if claim else [])}
    return rows

def generate(seed,output):
    allrows=defaultdict(list)
    for k in range(50):
        for sh,rows in make_case(seed,k).items(): allrows[sh].extend(rows)
    wb=Workbook(); wb.remove(wb.active)
    for sh,heads in SHEETS.items():
        ws=wb.create_sheet(sh);ws.append(heads)
        allrows[sh].sort(key=lambda x:(x[0],x[1] if len(x)>1 else 0))
        for row in allrows[sh]: ws.append(row)
    Path(output).parent.mkdir(parents=True,exist_ok=True);wb.save(output)
    return output
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--seed',type=int,required=True);ap.add_argument('--output',default='cases.xlsx');a=ap.parse_args()
    generate(a.seed,a.output);print(f'Wrote {a.output} with seed {a.seed}')
