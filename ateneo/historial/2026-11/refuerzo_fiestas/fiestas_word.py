"""Word del refuerzo fiestas desglosado por sucursal (primero AMBA, después interior; títulos en el orden del PAEC).
Dos PAEC (26/11): AMBA e interior, cada uno arranca en página nueva.
Uso: python -P fiestas_word.py DIR SII.xls SALIDA.docx 'SII dd/mm hh:mm'"""
import sys, pandas as pd
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
d, fsii, fout, etq = sys.argv[1:5]
rep = pd.read_csv(f'{d}/detalle_refuerzo_fiestas.csv', sep=';', encoding='utf-8-sig')
csv = {z: [tuple(map(int, l.split(';'))) for l in open(f'{d}/PAEC_REFUERZO_FIESTAS_{z}.csv', encoding='utf-8').read().split() if l] for z in ('AMBA', 'INTERIOR')}
paec = list(dict.fromkeys(i for z in csv for i, _ in csv[z]))
AMBA = set(open(f'{d}/amba.txt', encoding='utf-8').read().split('\n'))
s = pd.read_excel(fsii); s = s[s['ID ARTICULO'].notna()].copy(); s['n'] = pd.to_numeric(s.ISBN, errors='coerce'); by = s.set_index('n')
dem = rep.groupby('isbn').cant.sum()
corto = {i for i in dem.index if int(by.loc[i, 'DISPONIBLE']) < dem[i]}
assert set(rep.isbn) == set(paec) and int(rep.cant.sum()) == sum(c for z in csv for _, c in csv[z])
for z in csv: assert dict(csv[z]) == rep[rep.zona == z].groupby('isbn').cant.sum().astype(int).to_dict(), z
pos = {i: k for k, i in enumerate(dict.fromkeys(rep.isbn))}   # orden de la planilla (= orden de los dos CSV)
for z in csv: assert [i for i, _ in csv[z]] == sorted((i for i, _ in csv[z]), key=pos.get), z
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
doc.add_paragraph(f'Dos PAEC (AMBA e interior) · {len(paec)} títulos · {rep.local.nunique()} locales · {int(rep.cant.sum()):,} unidades · stock y disponible: {etq}.'.replace(',', '.'))
for z in csv:
    rz = rep[rep.zona == z]; p = doc.add_paragraph(style='List Bullet'); p.add_run(f'PAEC {z}: ').bold = True
    p.add_run(f'{len(csv[z])} títulos · {rz.local.nunique()} locales · {int(rz.cant.sum()):,} u (CSV PAEC_REFUERZO_FIESTAS_{z}.csv)'.replace(',', '.'))
doc.add_paragraph('No van: Grand Splendid, Unicenter, Bahía Blanca Shopping, Ezeiza, Patio Bullrich, Patio Olmos, Paseo Alcorta, '
                  'Palermo Portal, Ushuaia y Trelew. Primero el PAEC AMBA y después el del interior (cada uno empieza en página nueva), las sucursales en orden alfabético; los títulos en el orden del CSV.')
p = doc.add_paragraph(); p.add_run(f'* {len(corto)} títulos piden más que el disponible del depósito sumando los dos PAEC; se cargan igual: ').bold = True
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
        c[0].text, c[1].text, c[2].text, c[3].text = f'PAEC {zona}', l, str(len(sub)), f'{int(sub.cant.sum()):,}'.replace(',', '.')
c = t.add_row().cells; c[1].text = 'TOTAL'; c[2].text = str(len(rep)); c[3].text = f'{int(rep.cant.sum()):,}'.replace(',', '.')
for cell in c: 
    for r in cell.paragraphs[0].runs: r.bold = True
from docx.enum.section import WD_SECTION
st.paragraph_format.space_after = Pt(0); st.paragraph_format.space_before = Pt(0)
ns = doc.add_section(WD_SECTION.NEW_PAGE)                     # desglose a dos columnas
cols = ns._sectPr.xpath('./w:cols')
c0 = cols[0] if cols else OxmlElement('w:cols')
c0.set(qn('w:num'), '2'); c0.set(qn('w:space'), '340')
if not cols: ns._sectPr.append(c0)
def chico(cell, txt, bold=False, center=False):
    cell.text = ''; p = cell.paragraphs[0]; r = p.add_run(txt); r.font.size = Pt(6.5); r.bold = bold
    p.paragraph_format.space_after = Pt(0)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
W = (0.7, 2.05, 4.65, 1.3)
for k, (zona, zl) in enumerate(orden_loc):
    hz = doc.add_paragraph()
    if k: hz.paragraph_format.page_break_before = True
    r = hz.add_run(f'PAEC {zona}: {len(zl)} locales · {int(rep[rep.local.isin(zl)].cant.sum()):,} u'.replace(',', '.'))
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78); hz.paragraph_format.space_before = Pt(4)
    for l in zl:
        sub = rep[rep.local == l].copy(); sub['o'] = sub.isbn.map(pos); sub = sub.sort_values('o')
        hp = doc.add_paragraph(); hp.paragraph_format.space_before = Pt(6); hp.paragraph_format.keep_with_next = True
        r = hp.add_run(f'{l} · {int(sub.cant.sum()):,} u · {len(sub)} títulos'.replace(',', '.')); r.bold = True; r.font.size = Pt(9)
        tb = doc.add_table(rows=1, cols=4); tb.style = 'Table Grid'; tb.autofit = False
        for cc, txt in zip(tb.rows[0].cells, ('U.', 'ISBN', 'Título', 'Tiene→a')):
            chico(cc, txt, bold=True, center=True); sombra(cc, 'D9E2F3')
        rep_header(tb.rows[0])
        for _, x in sub.iterrows():
            i = int(x.isbn); c = tb.add_row().cells
            chico(c[0], str(int(x.cant)), center=True); chico(c[1], str(i))
            chico(c[2], str(by.loc[i, 'TITULO'])[:30] + (' *' if i in corto else '')); chico(c[3], f'{int(x.stock)}→{int(x.objetivo)}', center=True)
        for j, w in enumerate(W): tb.columns[j].width = Cm(w)
        for row in tb.rows:
            for j, w in enumerate(W): row.cells[j].width = Cm(w)
doc.save(fout)
print('ok', rep.local.nunique(), 'locales', int(rep.cant.sum()), 'u', len(rep), 'líneas')
