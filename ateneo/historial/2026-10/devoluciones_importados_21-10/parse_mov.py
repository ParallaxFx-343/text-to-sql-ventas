import sys, pandas as pd
def parse(f):
    d = pd.read_excel(f, header=None)
    heads, rows, cur = [], [], None
    i = 0
    while i < len(d):
        r = d.iloc[i]
        if r[0] == 'Codigo':
            nx = d.iloc[i+1]; cur = int(nx[0]); heads.append((cur, int(nx[1]), nx[2], nx[3], nx[4])); i += 2; continue
        if r[0] == 'Cliente' or pd.isna(r[0]): i += 1; continue
        rows.append(list(r[:11]) + [cur]); i += 1
    m = pd.DataFrame(rows, columns=['Cliente','Sucursal','Nombre','Mov','Comprobante','Codigo','Factura','Ctd','Restante','Fecha','Obs','Bloque'])
    m['Fecha'] = pd.to_datetime(m.Fecha, format='%d-%b-%y')
    for c in ['Cliente','Codigo','Ctd','Restante']: m[c] = m[c].astype('int64')
    return pd.DataFrame(heads, columns=['Codigo','ISBN','Titulo','Consignada','Restante']), m
