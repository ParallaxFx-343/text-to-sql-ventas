import sys, json, pickle, re, zipfile, pandas as pd
from pathlib import Path
blocks=pickle.load(open('repo1410/blocks.pkl','rb'))
orden=json.load(open(sys.argv[2]))['orden_paec']
out=Path(sys.argv[3]); out.mkdir(parents=True, exist_ok=True)
sii=pd.read_excel(sys.argv[1]); sii=sii[sii['ID ARTICULO'].notna()].copy(); sii['n']=pd.to_numeric(sii.ISBN,errors='coerce'); by=sii.set_index('n')
blocks.sort(key=lambda b: orden.index(b['suc']))
# demanda total por titulo (sin Sorpresa) vs disponible
dem={}
for b in blocks:
    for (_,o,v,t,isbn,st) in b['lineas']:
        if not str(isbn).startswith('290'): dem[int(isbn)]=dem.get(int(isbn),0)+int(v)
parcial={i:(d,int(by.loc[i,'DISPONIBLE'])) for i,d in dem.items() if 0<by.loc[i,'DISPONIBLE']<d}
print('titulos con disponible que no cubre el total:', {by.loc[i,'TITULO']:x for i,x in parcial.items()})
T=['REPOSICION 14/10 - UN PAEC POR SUCURSAL',
   'Cantidad = columna "Ventas" de la planilla (enviar la cantidad exacta).',
   'Sin Libros Sorpresa (Franco: se sacan). Los títulos sin disponible van igual y van marcados con [SIN STOCK].',
   'Con [DISP x PARA y EN TOTAL]: hay disponible pero no alcanza para lo que piden todas las sucursales juntas.',
   f'Disponible según SII 13/10 07:02. Orden de las sucursales: el de los TXT de PAEC.','']
tot_l=tot_u=tot_sor=tot_ss=0; files=[]
for b in blocks:
    L=[x for x in b['lineas'] if not str(x[4]).startswith('290')]
    sor=[x for x in b['lineas'] if str(x[4]).startswith('290')]
    slug=re.sub(r'[^A-Z0-9]+','_',b['suc']).strip('_')
    p=out/f'reposicion_{slug}.csv'; files.append(p)
    p.write_bytes(('\r\n'.join(f'{x[4]};{int(x[2])}' for x in L)+'\r\n').encode())
    ss=[x for x in L if by.loc[int(x[4]),'DISPONIBLE']<=0]
    T+=['='*72, f'{b["suc"]:<52}{sum(x[2] for x in L):>6} u  {len(L)} títulos',
        f'   archivo {p.name} | sacados {len(sor)} Sorpresa ({sum(x[2] for x in sor)} u) | sin stock {len(ss)} ({sum(x[2] for x in ss)} u)','='*72]
    for x in L:
        d=int(by.loc[int(x[4]),'DISPONIBLE'])
        T.append(f'   {int(x[2]):>4}  {x[4]}  {str(by.loc[int(x[4]),"TITULO"])[:44]:<44}' + ('  [SIN STOCK]' if d<=0 else (f'  [DISP {d} PARA {parcial[int(x[4])][0]} EN TOTAL]' if int(x[4]) in parcial else '')))
    T.append('')
    tot_l+=len(L); tot_u+=sum(x[2] for x in L); tot_sor+=sum(x[2] for x in sor); tot_ss+=sum(x[2] for x in ss)
    print(f'{b["suc"]:<25} {len(L)} lineas {sum(x[2] for x in L)} u | sorpresa fuera {sum(x[2] for x in sor)} | sin stock {sum(x[2] for x in ss)}')
T+=['='*72, f'TOTAL {len(blocks)} sucursales | {tot_l} líneas | {tot_u} u (de esas, {tot_ss} sin stock) | Sorpresa sacados: {tot_sor} u']
(out/'DETALLE_REPOSICION_14-10.txt').write_text('\n'.join(T)+'\n',encoding='utf-8-sig')
with zipfile.ZipFile(out/'reposicion_14-10_por_sucursal.zip','w') as z:
    for p in files: z.write(p,p.name)
print('TOTAL',tot_l,tot_u,'sin stock',tot_ss,'sorpresa',tot_sor)
