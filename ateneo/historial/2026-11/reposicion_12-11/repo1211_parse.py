import sys, json, pandas as pd
from openpyxl import load_workbook
f, cfg, outp = sys.argv[1:4]   # planilla, config.json, blocks.pkl
pad=json.load(open(cfg))['orden_paec']+['DEVOTO']
ws=load_workbook(f).worksheets[0]
rows=[r[:5] for r in ws.iter_rows(values_only=True)]   # solo A-E: a la derecha puede venir otra tabla
blocks=[]; cur=None
for i,r in enumerate(rows,1):
    if r[0] is None and r[2] and all(x is None for k,x in enumerate(r) if k!=2):
        cur={'local':r[2],'fila':i,'lineas':[]}; blocks.append(cur); continue
    if r[0]=='Orden': continue
    if isinstance(r[0],(int,float)) and cur is not None:
        cur['lineas'].append((i,)+tuple(r[:5]))
    elif any(x is not None for x in r) and i>3:
        print('fila rara',i,r)
import re
reg=[str(r[0]) for r in rows[:3] if r[0]]; print('REGLA:', reg)
obj=int(re.search(r'completar a (\d+)', ' '.join(reg)).group(1))
print(len(blocks),'bloques | completar a', obj)
for b in blocks:
    L=b['lineas']; print(f"{b['local']:<30} fila {b['fila']:>3} | {len(L):>3} lineas | {sum(x[2] for x in L):>4} u | orden 1..{max(x[1] for x in L)} ok={[x[1] for x in L]==list(range(1,len(L)+1))}")
import pickle; pickle.dump({'objetivo':obj,'blocks':blocks},open(outp,'wb'))
