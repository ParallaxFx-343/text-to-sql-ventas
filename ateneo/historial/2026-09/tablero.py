import os
import json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

AQUI=os.path.dirname(os.path.abspath(__file__))
falt=json.load(open(os.path.join(AQUI,'dash_falt.json'))); cad=json.load(open(os.path.join(AQUI,'dash_cad.json')))
NAVY='1F4E78'; NAVY2='2E6497'; INK='1F2A44'; MUT='6B7686'; CANVAS='EEF1F5'; SHADOW='D3D9E1'; WHITE='FFFFFF'
RED='C8373A'; ORANGE='E8834F'; LNAVY='D5E2F0'; LRED='F5D6D6'; GREY='A7B0BD'; LGREY='CDD3DC'
fill=lambda c: PatternFill('solid',fgColor=c)
thin=Side(style='thin',color='BFBFBF')

def txt(sz=900,color=MUT,b=False):
    cp=CharacterProperties(sz=sz,b=b,solidFill=color,latin=DFont(typeface='Calibri'))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),endParaRPr=cp)])
def transparente(ch):
    ch.graphical_properties=GraphicalProperties(noFill=True); ch.graphical_properties.line.noFill=True
    ch.plot_area.graphicalProperties=GraphicalProperties(noFill=True); ch.plot_area.graphicalProperties.line.noFill=True
    ch.roundedCorners=False
def anclar(ws,ch,c1,r1,c2,r2,pad=0):
    # c1..c2 / r1..r2 inclusivos. El borde final se fija dentro de la ultima celda (offset >= 0)
    r1,r2=int(r1),int(r2)
    wcol=ws.column_dimensions[c2].width or 8.43
    wemu=int(int(wcol*7+5)*9525); hemu=int((ws.row_dimensions[r2].height or 15)*12700)
    a=TwoCellAnchor()
    a._from=AnchorMarker(col=CI(c1)-1,colOff=pad,row=r1-1,rowOff=pad)
    a.to=AnchorMarker(col=CI(c2)-1,colOff=max(0,wemu-pad),row=r2-1,rowOff=max(0,hemu-pad))
    ch.anchor=a; ws.add_chart(ch)
def etiquetas(s,sz=900,color=INK,b=True,fmt=None):
    s.dLbls=DataLabelList(); s.dLbls.showVal=True
    for k in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(s.dLbls,k,False)
    s.dLbls.txPr=txt(sz,color,b)
    if fmt: s.dLbls.numFmt=fmt
def barras_h(cats,vals,colors=None,color=NAVY,gap=55,fmt='#,##0'):
    ch=BarChart(); ch.type='bar'; ch.legend=None; ch.title=None; ch.gapWidth=gap
    ch.add_data(vals,titles_from_data=False); ch.set_categories(cats)
    s=ch.series[0]; s.graphicalProperties.solidFill=color; s.graphicalProperties.line.noFill=True
    for i,c in enumerate(colors or []):
        p=DataPoint(idx=i); p.graphicalProperties.solidFill=c; p.graphicalProperties.line.noFill=True; s.dPt.append(p)
    etiquetas(s,fmt=fmt)
    ch.x_axis.scaling.orientation='maxMin'; ch.x_axis.delete=False; ch.y_axis.delete=True
    ch.x_axis.txPr=txt(900,INK); ch.x_axis.majorTickMark='none'
    ch.x_axis.graphicalProperties=GraphicalProperties(ln=LineProperties(solidFill='C9CFD8'))
    ch.y_axis.majorGridlines=None
    transparente(ch); return ch
def dona(ref,color,resto):
    ch=DoughnutChart(holeSize=68); ch.legend=None; ch.title=None; ch.firstSliceAng=0
    ch.add_data(ref,from_rows=True,titles_from_data=False); s=ch.series[0]
    for i,c in enumerate((color,resto)):
        p=DataPoint(idx=i); p.graphicalProperties.solidFill=c; p.graphicalProperties.line.noFill=True; s.dPt.append(p)
    transparente(ch)
    ch.plot_area.layout=Layout(manualLayout=ManualLayout(x=0.08,y=0.08,w=0.84,h=0.84,xMode='edge',yMode='edge'))
    return ch
def link(c,hoja,texto=None):
    if texto: c.value=texto
    c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{hoja}'!A1",display=str(c.value))
def encab(ws,row,vals,c0=1):
    for j,v in enumerate(vals,c0):
        c=ws.cell(row,j,v); c.font=Font(bold=True,color=WHITE); c.fill=fill(NAVY)
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
def ancho(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[CL(i)].width=x
def volver(ws,cell):
    c=ws[cell]; link(c,'Tablero','← Volver al tablero'); c.font=Font(color=NAVY,bold=True,underline='single')

wb=Workbook(); T=wb.active; T.title='Tablero'
A=wb.create_sheet('Acciones'); F=wb.create_sheet('Faltantes'); S=wb.create_sheet('Surtido en cero'); N=wb.create_sheet('Cómo se calculó')

# =================== ACCIONES ===================
A['A1']='Acciones de reposición · Septiembre 2026'; A['A1'].font=Font(bold=True,size=14,color=NAVY); volver(A,'G1')
A['A2']='Unidades según los PAEC y CSV entregados. Faltante: por título, lo pedido menos el disponible del SII del día.'; A['A2'].font=Font(italic=True,size=9,color=MUT)
encab(A,4,['Fecha','Acción','Regla','Títulos','Locales','Unidades cargadas','Sin stock en depósito'])
ACC=[(dt.date(2026,9,11),'Promo 2x1 $28.500','Completar a 8 · 20 Splendid · 3 en cinco locales chicos',11,33,255,293),
 (dt.date(2026,9,15),'Día de la Madre','Completar a 10 y 5 · Splendid a 30 y 20 · redondeo a 10',38,61,5101,128),
 (dt.date(2026,9,17),'Agenda + autores de octubre','Agenda a 10 y 40 en los Ateneo · autores a 10',12,61,2819,0),
 (dt.date(2026,9,17),'Canales · ranking','Pedido del ranking 1/8 al 17/9',56,'Canales',243,7),
 (dt.date(2026,9,21),'Reposición 21 al 25','Lo vendido · sin cobertura > 3 · solo lo servible',170,20,4004,1569),
 (dt.date(2026,9,23),'Canales · refuerzo','Lista cerrada · solo lo servible',43,'Canales',101,19)]
for i,r in enumerate(ACC,5):
    for j,v in enumerate(r,1):
        c=A.cell(i,j,v)
        if j==1: c.number_format='dd/mm'
        if j>=4: c.number_format='#,##0'; c.alignment=Alignment(horizontal='right')
        if j==3: c.font=Font(size=9,color=MUT)
A.cell(11,2,'TOTAL'); A['F11']='=SUM(F5:F10)'; A['G11']='=SUM(G5:G10)'
for j in range(1,8):
    c=A.cell(11,j); c.font=Font(bold=True); c.fill=fill('DDEBF7'); c.border=Border(top=Side(style='thin')); c.number_format='#,##0'
A['A14']='Reposición 21 al 25: de lo vendido a lo cargado'; A['A14'].font=Font(bold=True,size=12,color=NAVY)
encab(A,15,['Paso','Etiqueta','Detalle','Unidades','Base (aux.)','Barra (aux.)'])
WF=[('Vendido en el período','Vendido','Pedido original, 20 locales',6201),
    ('Cobertura mayor a 3','Cobertura > 3','Más de mes y medio de stock',-520),
    ('Sin stock en depósito','Sin stock','3 títulos',-781),
    ('Stock insuficiente','Insuficiente','9 títulos',-896),
    ('Cargado','Cargado','170 títulos',4004)]
for i,w in enumerate(WF,16):
    for j,v in enumerate(w,1): A.cell(i,j,v)
    A.cell(i,4).number_format='#,##0;−#,##0'
# base y altura por formula
A['E16']=0; A['F16']='=D16'
A['E17']='=F16+D17'; A['F17']='=-D17'
A['E18']='=E17+D18'; A['F18']='=-D18'
A['E19']='=E18+D19'; A['F19']='=-D19'
A['E20']=0; A['F20']='=D20'
for i in range(16,21):
    for j in (5,6): A.cell(i,j).number_format='#,##0'; A.cell(i,j).font=Font(size=9,color='A6A6A6')
A['G15']='Etiqueta gráfico'; A['G15'].font=Font(bold=True,color=WHITE); A['G15'].fill=fill(NAVY)
for r,t in zip(range(16,21),['Vendido   6.201','Cobertura > 3   −520','Sin stock   −781','Insuficiente   −896','Cargado   4.004']):
    A.cell(r,7,t).font=Font(size=9,color=MUT)
A['A21']='Control: vendido menos los tres cortes'; A['D21']='=D16+D17+D18+D19'; A['E21']='=IF(D21=D20,"cuadra","NO CUADRA")'
for c in ('A21','D21','E21'): A[c].font=Font(italic=True,size=9,color=MUT)
A['D21'].number_format='#,##0'
A['A24']='Auxiliar de las donas del tablero'; A['A24'].font=Font(bold=True,size=10,color=MUT)
encab(A,25,['Indicador','Porcentaje','Resto'])
A['A26']='Reposición 21-25: cargado / vendido'; A['B26']='=D20/D16'
A['A27']='Faltante: peso de los 5 primeros títulos'; A['B27']="=Faltantes!D28"
A['A28']='Locales alcanzados / activos'; A['B28']='=61/62'
for r in (26,27,28): A[f'C{r}']=f'=1-B{r}'; A[f'B{r}'].number_format='0%'; A[f'C{r}'].number_format='0%'
ancho(A,[34,30,46,10,11,15,15]); A.freeze_panes='A5'

# =================== FALTANTES ===================
CORTO={'PRINCIPITO, EL':'El Principito','GATITO, ¿POR QUE ESTAS ENOJADO?':'Gatito, ¿por qué…?','SAGA DE LOS ROMANOV, LA':'Saga Romanov',
'STICKERS MAGICOS: GRANJA':'Stickers Granja','STICKERS MAGICOS DINOSAURIOS':'Stickers Dinosaurios','BANDERAS DEL MUNDO':'Banderas del mundo',
'VENDER ES PODER':'Vender es poder','FIN DEL AUTOODIO, EL':'Fin del autoodio','HEDY':'Hedy','APRENDO A LEER LAS LETRAS':'Aprendo a leer: letras',
'AMIGOS DE LA GRANJA':'Amigos de la granja','APRENDO A LEER LAS SILABAS':'Aprendo a leer: sílabas'}
F['A1']='Dónde faltó stock'; F['A1'].font=Font(bold=True,size=14,color=NAVY); volver(F,'G1')
F['A2']='Unidades pedidas en septiembre que el depósito no llegó a cubrir. Disponible y cadena según SII del 22/09/2026.'; F['A2'].font=Font(italic=True,size=9,color=MUT)
encab(F,4,['ISBN','Título','Nombre corto','Faltante','Disp. depósito hoy','En la cadena hoy','Estado hoy','Acciones'])
TIT=falt['titulos']
for i,t in enumerate(TIT,5):
    est='Sin stock' if t['disp_hoy']==0 else 'Stock bajo'
    vals=[t['isbn'],t['titulo'],CORTO.get(t['titulo'],t['titulo'].title()),t['faltante'],t['disp_hoy'],t['cadena_hoy'],est,' · '.join(t['acciones'])]
    for j,v in enumerate(vals,1):
        c=F.cell(i,j,v)
        if j==1: c.number_format='0'
        if j in (4,5,6): c.number_format='#,##0'
    F.cell(i,7).font=Font(bold=True,color=(RED if est=='Sin stock' else 'B8612F'))
last=4+len(TIT)  # 25
F.cell(last+1,2,'TOTAL'); F.cell(last+1,4,f'=SUM(D5:D{last})')
for j in range(1,9):
    c=F.cell(last+1,j); c.font=Font(bold=True); c.fill=fill('DDEBF7'); c.border=Border(top=Side(style='thin')); c.number_format='#,##0'
F.cell(last+3,2,'Los 5 títulos de más faltante explican'); F.cell(last+3,4,f'=SUM(D5:D9)/D{last+1}')
F.cell(last+3,4).number_format='0%'; F.cell(last+3,2).font=Font(bold=True); F.cell(last+3,4).font=Font(bold=True,color=RED)
assert last+1==26 and last+3==28
F.auto_filter.ref=f'A4:H{last}'; F.freeze_panes='D5'; ancho(F,[16,44,22,10,13,13,12,40])

# =================== SURTIDO ===================
S['A1']='Surtido general en cero, por local'; S['A1'].font=Font(bold=True,size=14,color=NAVY); volver(S,'D1')
S['A2']=('De los 503 títulos que tiene al menos la mitad de la cadena y con 20 o más disponibles en depósito, '
         'cuántos están en 0 en cada local (SII del 22/09/2026). Los locales chicos trabajan con surtido reducido.')
S['A2'].font=Font(italic=True,size=9,color=MUT); S['A2'].alignment=Alignment(wrap_text=True,vertical='top'); S.merge_cells('A2:C2'); S.row_dimensions[2].height=48
encab(S,4,['Local','Títulos en 0'])
LOC=sorted(cad['quiebres'].items(),key=lambda x:-x[1])
for i,(k,v) in enumerate(LOC,5): S.cell(i,1,k); S.cell(i,2,v)
lastl=4+len(LOC)
S.cell(lastl+2,1,'Mediana de la cadena'); S.cell(lastl+2,2,f'=MEDIAN(B5:B{lastl})')
S.cell(lastl+3,1,'Locales en cero'); S.cell(lastl+3,2,f'=COUNTIF(B5:B{lastl},0)')
for r in (lastl+2,lastl+3): S.cell(r,1).font=Font(bold=True); S.cell(r,2).font=Font(bold=True)
S.auto_filter.ref=f'A4:B{lastl}'; S.freeze_panes='A5'; ancho(S,[32,13,24,22])
sc=barras_h(Reference(S,min_col=1,min_row=5,max_row=19),Reference(S,min_col=2,min_row=5,max_row=19),fmt='0')
sc.graphical_properties=GraphicalProperties(solidFill=WHITE); sc.graphical_properties.line.solidFill='D9D9D9'
anclar(S,sc,'E',4,'L',30)

# =================== NOTAS ===================
N['A1']='Cómo se calculó'; N['A1'].font=Font(bold=True,size=14,color=NAVY); volver(N,'C1')
notas=[('Cargado','Suma de los CSV y PAEC entregados. En Día de la Madre y en Agenda + autores, el PAEC coincide con la suma por local.'),
 ('Faltante','Por título, lo pedido menos el disponible del SII del día de la acción. Incluye unidades que igual se cargaron (Día de la Madre y Canales del 17/09).'),
 ('SII','Se excluye la fila de totales del pie. El depósito cuadra contra ella: 638.186 u. físicas y 634.202 disponibles al 22/09.'),
 ('Devoto','No figura: sin stock ni actividad al cierre, excluido por decisión en todas las acciones.'),
 ('Surtido en cero','Títulos presentes en al menos 30 de los 61 locales activos y con 20 o más disponibles en depósito.'),
 ('Fijos','Locales alcanzados (61) y títulos distintos (247) salen de cruzar los archivos: van como valor, no como fórmula.')]
for i,(k,v) in enumerate(notas,3):
    N.cell(i,1,k).font=Font(bold=True); c=N.cell(i,2,v); c.alignment=Alignment(wrap_text=True,vertical='top'); N.row_dimensions[i].height=32
ancho(N,[18,100,20])

# =================== TABLERO ===================
T.sheet_view.showGridLines=False; T.sheet_view.zoomScale=90
COLS={'A':13,'B':2.5,'I':0.9,'J':1.8,'Q':0.9,'R':1.8,'Y':0.9,'Z':1.8,'AJ':0.9,'AK':1.8}
for c in [CL(i) for i in range(1,38)]:
    T.column_dimensions[c].width=COLS.get(c,5.3)
ROWS={1:8,2:30,3:20,4:20,5:4,6:10,7:24,13:4,14:10,15:24,16:16,34:4,35:12}
for r in range(1,36): T.row_dimensions[r].height=ROWS.get(r,20 if 8<=r<=12 else 19)
# lienzo gris
for r in range(1,36):
    for c in range(2,38): T.cell(r,c).fill=fill(CANVAS)
def tarjeta(c1,r1,c2,r2):
    a,b=CI(c1),CI(c2)
    for r in range(r1,r2+1):
        for c in range(a,b+1): T.cell(r,c).fill=fill(WHITE)
    for r in range(r1+1,r2+2): T.cell(r,b+1).fill=fill(SHADOW)
    for c in range(a+1,b+2): T.cell(r2+1,c).fill=fill(SHADOW)
tarjeta('C',2,'X',4)                          # titulo
tarjeta('C',7,'H',12); tarjeta('K',7,'P',12); tarjeta('S',7,'X',12)   # KPI
tarjeta('C',15,'P',33); tarjeta('S',15,'X',33)                      # graficos abajo
tarjeta('AA',2,'AI',33)                                             # tarjeta alta
# barra lateral
for r in range(1,36): T.cell(r,1).fill=fill(NAVY)
def lat(r,texto,bold=False,sz=11,activo=False,hoja=None,h=None):
    c=T.cell(r,1,texto); c.font=Font(color=WHITE,bold=bold,size=sz,underline=None)
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    if activo: c.fill=fill(NAVY2)
    if hoja: link(c,hoja); c.font=Font(color=WHITE,bold=bold,size=sz)
T['A2']='SEP'; T['A2'].font=Font(color=WHITE,bold=True,size=20); T['A2'].alignment=Alignment(horizontal='center',vertical='center')
T['A3']='2026'; T['A3'].font=Font(color='BFD3EA',bold=True,size=12); T['A3'].alignment=Alignment(horizontal='center',vertical='top')
lat(8,'TABLERO',bold=True,sz=10,activo=True)
lat(10,'ACCIONES',sz=10,hoja='Acciones'); lat(12,'FALTANTES',sz=10,hoja='Faltantes')
lat(16,'LOCALES',sz=10,hoja='Surtido en cero'); lat(18,'MÉTODO',sz=10,hoja='Cómo se calculó')
T['A33']='SII 22/09'; T['A33'].font=Font(color='BFD3EA',size=8); T['A33'].alignment=Alignment(horizontal='center')

def txtcell(rng,val,**kw):
    c1=rng.split(':')[0]; T.merge_cells(rng); c=T[c1]; c.value=val
    c.font=Font(**{k:v for k,v in kw.items() if k in ('bold','size','color','italic')})
    c.alignment=Alignment(horizontal=kw.get('h','left'),vertical=kw.get('v','center'),wrap_text=kw.get('wrap',False),indent=kw.get('ind',1))
    if 'fmt' in kw: c.number_format=kw['fmt']
    return c
# titulo
txtcell('C2:S3','Reposición de la cadena · Septiembre 2026',bold=True,size=19,color=NAVY,v='bottom')
txtcell('C4:R4','Del 11 al 23/09/2026 · stock según el SII del día de cada acción',italic=True,size=10,color=MUT,v='top')
txtcell('T2:X3','=Acciones!F11',bold=True,size=26,color=NAVY,h='right',v='bottom',fmt='#,##0',ind=1)
txtcell('S4:X4','unidades cargadas · 6 acciones · 247 títulos',size=9,color=MUT,h='right',v='top',ind=1)
# KPIs
KP=[('C','H','F','Reposición 21 al 25','=Acciones!D20','cargadas de 6.201 vendidas',26,NAVY,LNAVY),
    ('K','P','N','Faltante de depósito','=Faltantes!D26','unidades pedidas sin stock; 5 títulos concentran el grueso',27,RED,LRED),
    ('S','X','V','Locales alcanzados','61','de 62 activos, más Canales Alternativos',28,NAVY,LNAVY)]
for c1,c2,d1,tit,num,cap,auxrow,col,resto in KP:
    a=CI(c1); dcol=CI(d1)
    txtcell(f'{c1}7:{c2}7',tit,bold=True,size=12,color=INK,v='bottom')
    e1=CL(dcol-1)
    n=txtcell(f'{c1}8:{e1}10',int(num) if num.isdigit() else num,bold=True,size=24,color=(RED if col==RED else NAVY),v='center',fmt='#,##0')
    txtcell(f'{c1}11:{e1}12',cap,size=9,color=MUT,v='top',wrap=True)
    # porcentaje detras del agujero de la dona
    txtcell(f'{d1}10:{c2}10',f'=Acciones!B{auxrow}',bold=True,size=13,color=col,h='center',v='center',fmt='0%',ind=0)
    d=dona(Reference(A,min_col=2,max_col=3,min_row=auxrow,max_row=auxrow),col,resto)
    anclar(T,d,d1,8,c2,12)
# grafico acciones
txtcell('C15:P15','Unidades cargadas por acción',bold=True,size=13,color=INK,v='bottom')
txtcell('C16:P16','En orden cronológico, del 11 al 23 de septiembre',size=9,color=MUT,v='top')
ch=barras_h(Reference(A,min_col=2,min_row=5,max_row=10),Reference(A,min_col=6,min_row=5,max_row=10),gap=70)
anclar(T,ch,'C',17,'P',33,pad=60000)
# cascada
txtcell('S15:X15','Reposición 21 al 25',bold=True,size=13,color=INK,v='bottom')
txtcell('S16:X16','De lo vendido a lo cargado',size=9,color=MUT,v='top')
wf=BarChart(); wf.type='bar'; wf.grouping='stacked'; wf.overlap=100; wf.gapWidth=45; wf.legend=None; wf.title=None
wf.add_data(Reference(A,min_col=5,min_row=16,max_row=20),titles_from_data=False)
wf.add_data(Reference(A,min_col=6,min_row=16,max_row=20),titles_from_data=False)
wf.set_categories(Reference(A,min_col=7,min_row=16,max_row=20))
b=wf.series[0]; b.graphicalProperties.noFill=True; b.graphicalProperties.line.noFill=True
v=wf.series[1]; v.graphicalProperties.line.noFill=True
for i,c in enumerate([GREY,LGREY,RED,ORANGE,NAVY]):
    p=DataPoint(idx=i); p.graphicalProperties.solidFill=c; p.graphicalProperties.line.noFill=True; v.dPt.append(p)
wf.x_axis.scaling.orientation='maxMin'; wf.x_axis.delete=False; wf.y_axis.delete=True; wf.y_axis.majorGridlines=None
wf.x_axis.txPr=txt(900,INK); wf.x_axis.majorTickMark='none'
wf.x_axis.graphicalProperties=GraphicalProperties(ln=LineProperties(solidFill='C9CFD8'))
wf.y_axis.scaling.max=6300; wf.y_axis.scaling.min=0
transparente(wf); anclar(T,wf,'S',17,'X',33,pad=40000)
# tarjeta alta: faltantes
txtcell('AA2:AI2','Dónde faltó stock',bold=True,size=13,color=INK,v='bottom')
txtcell('AA3:AI4','Faltante de depósito por título, top 12. Rojo: hoy en 0 en depósito; naranja: stock bajo.',size=9,color=MUT,v='top',wrap=True)
cols=[RED if t['disp_hoy']==0 else ORANGE for t in TIT[:12]]
fc=barras_h(Reference(F,min_col=3,min_row=5,max_row=16),Reference(F,min_col=4,min_row=5,max_row=16),colors=cols,gap=40)
fc.y_axis.scaling.max=460; fc.y_axis.scaling.min=0
anclar(T,fc,'AA',5,'AI',30,pad=50000)
txtcell('AA31:AI32','=TEXT(Faltantes!D28,"0%")&" del faltante está en los primeros 5 títulos, y los cinco siguen hoy con 2 disponibles o menos."',size=9,color=INK,v='top',wrap=True,bold=True)
txtcell('AA33:AI33','Detalle completo en la hoja Faltantes',size=8,color=MUT,v='center')
link(T['AA33'],'Faltantes'); T['AA33'].font=Font(size=8,color=NAVY,underline='single')
# impresion
T.page_setup.orientation='landscape'; T.page_setup.fitToWidth=1; T.page_setup.fitToHeight=1
T.sheet_properties.pageSetUpPr.fitToPage=True; T.print_area='A1:AK35'
for m in ('left','right','top','bottom'): setattr(T.page_margins,m,0.2)
wb.active=0
wb.save('Reposicion_Septiembre_2026_tablero.xlsx'); print('ok')
