import sys, json, pickle, unicodedata, pandas as pd
blocks=pickle.load(open('repo1410/blocks.pkl','rb'))
pad=json.load(open(sys.argv[2]))['orden_paec']+['DEVOTO']
sii=pd.read_excel(sys.argv[1]); sii=sii[sii['ID ARTICULO'].notna()].copy(); sii['n']=pd.to_numeric(sii.ISBN,errors='coerce')
by=sii.set_index('n')
norm=lambda s: unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().upper().strip()
def dv(i):
    s=str(i); return len(s)==13 and (10-sum(int(d)*(1 if k%2==0 else 3) for k,d in enumerate(s[:12]))%10)%10==int(s[12])
for b in blocks:
    n=norm(b['local']); cand=[p for p in pad if p==n] or [p for p in pad if p.startswith(n)] or [p for p in pad if n in p]
    # verificar contra stock
    scores=[]
    for p in pad:
        ok=tot=0
        for (_,o,v,t,isbn,st) in b['lineas']:
            if int(isbn) in by.index and st is not None:
                tot+=1; ok+= int(0 if pd.isna(by.loc[int(isbn),p]) else by.loc[int(isbn),p])==int(st)
        scores.append((ok/tot if tot else 0,p))
    scores.sort(reverse=True)
    b['suc']=scores[0][1]
    isb=[x[4] for x in b['lineas']]; dup=len(isb)-len(set(isb))
    bad=[x for x in isb if not dv(x)]; nos=[x for x in isb if int(x) not in by.index]
    tit_dif=[(x[4],x[3],by.loc[int(x[4]),'TITULO']) for x in b['lineas'] if int(x[4]) in by.index and norm(str(x[3]))!=norm(str(by.loc[int(x[4]),'TITULO']))]
    sor=[x for x in b['lineas'] if str(x[4]).startswith('290')]
    falta=[(x[3],x[2],int(by.loc[int(x[4]),'DISPONIBLE'])) for x in b['lineas'] if not str(x[4]).startswith('290') and int(x[4]) in by.index and by.loc[int(x[4]),'DISPONIBLE']<x[2]]
    ceros=[x for x in b['lineas'] if not x[2] or x[2]<=0]
    print(f"{b['local']:<20} -> {scores[0][1]:<25} match {scores[0][0]:.0%} (2do {scores[1][1]} {scores[1][0]:.0%}) cand={cand} | dup {dup} dvmal {bad} noSII {nos} ventas<=0 {len(ceros)}")
    print(f"     sorpresa {len(sor)} lineas {sum(x[2] for x in sor)} u | no alcanza disp: {falta}")
    if tit_dif: print('     titulos distintos:', tit_dif[:5])
pickle.dump(blocks,open('repo1410/blocks.pkl','wb'))
