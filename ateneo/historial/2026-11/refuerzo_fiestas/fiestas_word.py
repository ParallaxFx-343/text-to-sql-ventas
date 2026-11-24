"""Word del refuerzo fiestas desglosado por sucursal (primero AMBA, después interior; títulos en el orden del PAEC).
Uso: python -P fiestas_word.py DIR SII.xls SALIDA.docx"""
import sys, pandas as pd
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
d, fsii, fout = sys.argv[1:4]
rep = pd.read_csv(f'{d}/detalle_refuerzo_fiestas.csv', sep=';', encoding='utf-8-sig')
paec = [int(l.split(';')[0]) for l in open(f'{d}/PAEC_REFUERZO_FIESTAS.csv', encoding='utf-8').read().split() if l]
AMBA = set(open(f'{d}/amba.txt', encoding='utf-8').read().split('\n'))
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce'); by = s.set_index('n')
dem = rep.groupby('isbn').cant.sum()
corto = {i for i in dem.index if int(by.loc[i, 'DISPONIBLE']) < dem[i]}
pos = {i: k for k, i in enumerate(paec)}
assert set(rep.isbn) == set(paec) and int(rep.cant.sum()) == 41991
doc = Document()
st = doc.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(9)
sec = doc.sections[0]
for m in ('left_margin', 'right_margin'): setattr(sec, m, Cm(1.5))
sec.top_margin = sec.bottom_margin = Cm(1.3)
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
def campo(p, instr):
    r = p.add_run(); a = OxmlElement('w:fldChar'); a.set(qn('w:fldCharType'), 'begin'); r._r.append(a)
    t = OxmlElement('w:instrText'); t.set(qn('xml:space'), 'preserve'); t.text = instr; r._r.append(t)
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'end'); r._r.append(b)
fp.add_run('Refuerzos fiestas · página '); campo(fp, 'PAGE'); fp.add_run(' de '); campo(fp, 'NUMPAGES')
def sombra(cell, color):
    tc = cell._tc.get_or_add_tcPr(); sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), color); tc.append(sh)
def rep_header(row):
    trPr = row._tr.get_or_add_trPr(); h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true'); trPr.append(h)
h = doc.add_heading('Refuerzo fiestas · desglose por sucursal', 0)
p = doc.add_paragraph(); r = p.add_run('ROTULAR COMO: "Refuerzos fiestas"'); r.bold = True; r.font.size = Pt(12)
doc.add_paragraph(f'Un solo PAEC · {len(paec)} títulos · {rep.local.nunique()} locales · {int(rep.cant.sum()):,} unidades · stock y disponible: SII 24/11 07:02.'.replace(',', '.'))
doc.add_paragraph('No van: Grand Splendid, Unicenter, Bahía Blanca Shopping, Ezeiza, Patio Bullrich, Patio Olmos, Paseo Alcorta, '
                  'Palermo Portal, Ushuaia y Trelew. Primero AMBA y después interior, cada grupo en orden alfabético; los títulos en el orden del PAEC.')
p = doc.add_paragraph(); p.add_run(f'* {len(corto)} títulos piden más que el disponible del depósito; se cargan igual: ').bold = True
p.add_run('; '.join(f'{by.loc[i, "TITULO"]} (pide {int(dem[i])}, disp {int(by.loc[i, "DISPONIBLE"])})' for i in sorted(corto, key=lambda i: pos[i])) + '.')
# resumen
doc.add_heading('Resumen por sucursal', 1)
t = doc.add_table(rows=1, cols=4); t.style = 'Light Grid Accent 1'
for c, txt in zip(t.rows[0].cells, ('Zona', 'Sucursal', 'Títulos', 'Unidades')): c.text = txt
rep_header(t.rows[0])
orden_loc = []
for zona in ('AMBA', 'INTERIOR'):
    zl = sorted(l for l in rep.local.unique() if (l in AMBA) == (zona == 'AMBA')); orden_loc.append((zona, zl))
    for l in zl:
        sub = rep[rep.local == l]; c = t.add_row().cells
        c[0].text, c[1].text, c[2].text, c[3].text = zona, l, str(len(sub)), f'{int(sub.cant.sum()):,}'.replace(',', '.')
c = t.add_row().cells; c[1].text = 'TOTAL'; c[2].text = str(len(rep)); c[3].text = f'{int(rep.cant.sum()):,}'.replace(',', '.')
for cell in c: 
    for r in cell.paragraphs[0].runs: r.bold = True
for zona, zl in orden_loc:
    doc.add_page_break()
    doc.add_heading(f'{zona}: {len(zl)} locales · {int(rep[rep.local.isin(zl)].cant.sum()):,} u'.replace(',', '.'), 1)
    for k, l in enumerate(zl):
        if k: doc.add_page_break()
        sub = rep[rep.local == l].copy(); sub['o'] = sub.isbn.map(pos); sub = sub.sort_values('o')
        doc.add_heading(f'{l} · {int(sub.cant.sum()):,} u · {len(sub)} títulos'.replace(',', '.'), 2)
        tb = doc.add_table(rows=1, cols=6); tb.style = 'Table Grid'
        for cc, txt in zip(tb.rows[0].cells, ('Cant.', 'ISBN', 'ID', 'Título', 'Tiene', 'Completa a')):
            cc.text = txt; cc.paragraphs[0].runs[0].bold = True; sombra(cc, 'D9E2F3')
        rep_header(tb.rows[0])
        for _, x in sub.iterrows():
            i = int(x.isbn); c = tb.add_row().cells
            c[0].text = str(int(x.cant)); c[1].text = str(i); c[2].text = str(int(by.loc[i, 'ID ARTICULO']))
            c[3].text = str(by.loc[i, 'TITULO']) + (' *' if i in corto else ''); c[4].text = str(int(x.stock)); c[5].text = str(int(x.objetivo))
            for j in (0, 4, 5): c[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tb.autofit = False
        W = (1.3, 3.0, 1.6, 9.5, 1.2, 2.0)
        for j, w in enumerate(W): tb.columns[j].width = Cm(w)
        for row in tb.rows:
            for j, w in enumerate(W): row.cells[j].width = Cm(w)
doc.save(fout)
print('ok', rep.local.nunique(), 'locales', int(rep.cant.sum()), 'u', len(rep), 'líneas')
