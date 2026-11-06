"""
Rakordim i Librit të Blerjeve: Financa vs Portali Self Care.
Konventa: Diferenca = Financa − Portal (pozitive = Financa ka më shumë).
Çiftimi i faturave: 1) NIVF e saktë, 2) NIVF e përafërt (Levenshtein ≤ 3) me të njëjtin NIPT,
3) NIPT + vlerë (≤ 3 lekë) për serialet e pavlefshme. Pastaj faturat pa çift klasifikohen në "Urën e blerjeve".
"""
import re
import numpy as np
import pandas as pd

import rakordim
from rakordim import (TOLERANCE, MUAJT, fmt_data, kush_ka_me_shume, fleta_default,
                      emri_muajit, muaji_kryesor, format_workbook, STATUS_FILLS)

UUID_RE = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')
LEV_MAX = 3

ST_OK = 'OK'
ST_PAS = 'PAS DATËS SË LIBRIT SC'
ST_VIT = 'VIT I MËPARSHËM, KONTROLLO PERIUDHËN'
ST_MUNGON_FIN = 'MUNGON NË FINANCA'
ST_MUNGON_PORTAL = 'MUNGON NË PORTAL'
ST_ANULOHET = 'ÇIFT QË ANULOHET'
ST_DIFERENCE = 'DIFERENCË VLERE'
ST_NETO = 'REGJISTRUAR NETO'
ST_KATEGORI = 'KATEGORI TVSH E NDRYSHME'
ST_PA_SHPJEGIM = 'PA SHPJEGIM'

for _k, _v in {ST_PAS: 'BDD7EE', ST_VIT: 'FFEB9C', ST_MUNGON_FIN: 'FFC7CE', ST_MUNGON_PORTAL: 'FFC7CE',
               ST_ANULOHET: 'C6EFCE', ST_DIFERENCE: 'FFC7CE', ST_NETO: 'C6EFCE', ST_KATEGORI: 'FFEB9C'}.items():
    STATUS_FILLS.setdefault(_k, _v)


def parse_libri_blerjes(path, sheet_name):
    """
    Lexon librin e blerjeve. Të dhënat fillojnë te rreshti 13 (ose pas rreshtit të kodeve a, b, c...)
    dhe mbarojnë para rreshtit që fillon me "Shuma totale", "Kutia" ose "Shpjegim".
    Kolonat: A nr. fature, B serial/NIVF, C data, D furnitori, F NIPT, G totali,
    H TVSH jashtë zbritjes, T baza 20%, U TVSH 20%.
    """
    if hasattr(path, 'seek'):
        path.seek(0)
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    start = 12
    kodet = raw.index[raw[0].astype(str).str.strip() == 'a']
    if len(kodet) and kodet[0] < 40:
        start = kodet[0] + 1
    first = raw.iloc[:, 0].astype(str).str.strip().str.lower()
    ndalo = raw.index[(raw.index >= start) & first.str.startswith(('shuma totale', 'kutia', 'shpjegim'))]
    stop = ndalo[0] if len(ndalo) else len(raw)
    cols = {0: 'NrFature', 1: 'NumriSerial', 2: 'Data', 3: 'Furnitori', 5: 'NIPT', 6: 'Vlera',
            7: 'TvshJashtZbritjes', 19: 'Baza20', 20: 'Tvsh20'}
    data = raw.iloc[start:stop, list(cols.keys())].copy()
    data.columns = list(cols.values())
    data = data.dropna(how='all')
    data = data[data['Vlera'].notna() | data['NumriSerial'].notna() | data['Furnitori'].notna()]
    data['Data'] = data['Data'].map(lambda x: pd.to_datetime(x, dayfirst=True, errors='coerce')).dt.normalize()
    for c in ('Vlera', 'TvshJashtZbritjes', 'Baza20', 'Tvsh20'):
        data[c] = pd.to_numeric(data[c], errors='coerce').fillna(0 if c != 'Vlera' else np.nan)
    for c in ('NIPT', 'NumriSerial', 'NrFature', 'Furnitori'):
        data[c] = data[c].astype(str).str.strip()
        data.loc[data[c].isin(['nan', 'None', 'NaN', '']), c] = np.nan
    data['Vlera'] = data['Vlera'].fillna(0)
    data['NIPT'] = data['NIPT'].str.upper()
    data = data.reset_index(drop=True)
    data.attrs['shuma_totale'] = (pd.to_numeric(raw.iloc[stop, 6], errors='coerce') if stop < len(raw) else None)
    return data


def serial_i_vlefshem(s):
    return isinstance(s, str) and bool(UUID_RE.match(s.strip().lower()))


def levenshtein(a, b, maks=LEV_MAX):
    if abs(len(a) - len(b)) > maks:
        return maks + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j-1] + 1, prev[j-1] + (ca != cb)))
        prev = cur
    return prev[-1]


def ciftimi(fin, port, tol=TOLERANCE):
    """Kthen listën e çifteve (idx_fin, idx_port, metoda). Çdo rresht çiftohet vetëm një herë."""
    fin_s = fin['NumriSerial'].str.lower()
    port_s = port['NumriSerial'].str.lower()
    perdorur_f, perdorur_p, cifte = set(), set(), []

    # 1. NIVF e saktë
    nga_serial = {}
    for j, s in port_s.items():
        if isinstance(s, str):
            nga_serial.setdefault(s, []).append(j)
    for i, s in fin_s.items():
        if not isinstance(s, str) or s not in nga_serial:
            continue
        kand = [j for j in nga_serial[s] if j not in perdorur_p]
        if kand:
            j = min(kand, key=lambda j: abs(fin.at[i, 'Vlera'] - port.at[j, 'Vlera']))
            cifte.append((i, j, 'NIVF'))
            perdorur_f.add(i)
            perdorur_p.add(j)

    # 2. NIVF e përafërt, vetëm me të njëjtin NIPT
    for i, s in fin_s.items():
        if i in perdorur_f or not isinstance(s, str) or pd.isna(fin.at[i, 'NIPT']):
            continue
        kand = [(levenshtein(s, port_s[j]), j) for j in port.index[port['NIPT'] == fin.at[i, 'NIPT']]
                if j not in perdorur_p and isinstance(port_s[j], str)]
        kand = [k for k in kand if k[0] <= LEV_MAX]
        if kand:
            _, j = min(kand)
            cifte.append((i, j, 'NIVF e përafërt'))
            perdorur_f.add(i)
            perdorur_p.add(j)

    # 3. NIPT + vlerë, për serialet e pavlefshme/mungesë
    for i in fin.index:
        if i in perdorur_f or serial_i_vlefshem(fin.at[i, 'NumriSerial']) or pd.isna(fin.at[i, 'NIPT']):
            continue
        kand = [j for j in port.index[port['NIPT'] == fin.at[i, 'NIPT']]
                if j not in perdorur_p and abs(fin.at[i, 'Vlera'] - port.at[j, 'Vlera']) <= tol]
        if kand:
            j = min(kand, key=lambda j: (abs(fin.at[i, 'Vlera'] - port.at[j, 'Vlera']),
                                         abs((fin.at[i, 'Data'] - port.at[j, 'Data']).days)
                                         if pd.notna(fin.at[i, 'Data']) and pd.notna(port.at[j, 'Data']) else 0))
            cifte.append((i, j, 'NIPT + vlerë'))
            perdorur_f.add(i)
            perdorur_p.add(j)
    return cifte


def _cift_anulues(rows_df, tol=TOLERANCE):
    """Çifte brenda së njëjtës anë: i njëjti furnitor (NIPT), shuma e kundërt. Kthen set çiftesh indekse."""
    perdorur, cifte = set(), []
    for i in rows_df.index:
        if i in perdorur:
            continue
        for j in rows_df.index:
            if j <= i or j in perdorur:
                continue
            if (rows_df.at[i, 'NIPT'] == rows_df.at[j, 'NIPT'] and pd.notna(rows_df.at[i, 'NIPT'])
                    and abs(rows_df.at[i, 'Vlera'] + rows_df.at[j, 'Vlera']) <= tol and abs(rows_df.at[i, 'Vlera']) > tol):
                cifte.append((i, j))
                perdorur.update((i, j))
                break
    return cifte


def rakordo_blerje(financa_path, financa_sheet, portal_path, portal_sheet, cutoff=None, tol=TOLERANCE,
                   eksport_financa_path=None):
    fin = parse_libri_blerjes(financa_path, financa_sheet)
    port = parse_libri_blerjes(portal_path, portal_sheet)

    paralajmerime = []
    m_fin, m_port = muaji_kryesor(fin), muaji_kryesor(port)
    if m_fin is not None and m_port is not None and m_fin != m_port:
        paralajmerime.append(f"Skedarët duket se mbulojnë periudha të ndryshme: Financa është kryesisht {emri_muajit(m_fin)}, Portali {emri_muajit(m_port)}.")
    for emri, df in (('Financa', fin), ('Portali', port)):
        shuma = df.attrs.get('shuma_totale')
        if shuma is not None and pd.notna(shuma) and abs(df['Vlera'].sum() - shuma) > tol:
            paralajmerime.append(f"{emri}: totali i rreshtave të lexuar ({df['Vlera'].sum():,.0f}) ndryshon nga rreshti "
                                 f"'Shuma totale' i librit ({shuma:,.0f}).")

    cutoff = pd.Timestamp(cutoff).normalize() if cutoff is not None else port['Data'].max()
    if pd.isna(cutoff):
        cutoff = None
    viti_periudhes = m_port.year if m_port is not None else None

    cifte = ciftimi(fin, port, tol)
    f_cift = {i: (j, m) for i, j, m in cifte}
    p_cift = {j for _, j, _ in cifte}
    vetem_fin = fin.loc[[i for i in fin.index if i not in f_cift]]
    vetem_port = port.loc[[j for j in port.index if j not in p_cift]]

    anulohen_f = _cift_anulues(vetem_fin, tol)
    anulohen_p = _cift_anulues(vetem_port, tol)
    anulohen_idx_f = {i for p in anulohen_f for i in p}
    anulohen_idx_p = {i for p in anulohen_p for i in p}

    # Kreditë vetëm-në-Portal (për zbulimin e regjistrimit neto)
    kredi_port = vetem_port[(vetem_port['Vlera'] < 0) & ~vetem_port.index.isin(anulohen_idx_p)]
    kredi_perdorur = set()

    rows = []

    def rresht(nr, nivf, data, furn, nipt, vf, vp, efekti, status, shenim='', metoda=''):
        nivf = nivf if isinstance(nivf, str) else 'pa NIVF'
        rows.append({'NrFature': nr, 'NIVF': nivf, 'Data': data, 'Furnitori': furn, 'NIPT': nipt,
                     'Vlera_Financa': vf, 'Vlera_Portal': vp, 'Efekti': efekti,
                     'Kush_ka_me_shume': kush_ka_me_shume(efekti), 'Statusi': status, 'Metoda': metoda, 'Shenim': shenim})

    rrumbullakime = 0.0
    n_rrumb = 0
    cifte_info = []
    status_fin = {}
    for i, (j, metoda) in f_cift.items():
        f, p = fin.loc[i], port.loc[j]
        d = f['Vlera'] - p['Vlera']
        cifte_info.append((i, j, metoda, d))
        status_fin[i] = ST_OK
        if abs(d) <= tol:
            rrumbullakime += d
            n_rrumb += 1
            continue
        # regjistrim neto: diferenca = vlera e një krediti vetëm-në-Portal të të njëjtit furnitor
        kredi = [k for k in kredi_port.index if k not in kredi_perdorur and kredi_port.at[k, 'NIPT'] == p['NIPT']
                 and abs(kredi_port.at[k, 'Vlera'] - d) <= tol]
        if kredi:
            k = kredi[0]
            kredi_perdorur.add(k)
            status_fin[i] = ST_NETO
            rresht(f['NrFature'], p['NumriSerial'], f['Data'], f['Furnitori'], f['NIPT'], f['Vlera'], p['Vlera'], d, ST_NETO,
                   f"Financa e ka faturën neto nga kredia {kredi_port.at[k, 'NrFature']} ({kredi_port.at[k, 'Vlera']:,.0f}) që Portali e ka të ndarë.", metoda)
            rresht(kredi_port.at[k, 'NrFature'], kredi_port.at[k, 'NumriSerial'], kredi_port.at[k, 'Data'], kredi_port.at[k, 'Furnitori'],
                   kredi_port.at[k, 'NIPT'], None, kredi_port.at[k, 'Vlera'], -kredi_port.at[k, 'Vlera'], ST_NETO,
                   'Kredi vetëm në Portal; përbën një me faturën e regjistruar neto në Financa. Efekti total 0.')
        else:
            status_fin[i] = ST_DIFERENCE
            rresht(f['NrFature'], p['NumriSerial'], f['Data'], f['Furnitori'], f['NIPT'], f['Vlera'], p['Vlera'], d, ST_DIFERENCE,
                   f"Diferencë vlere {d:,.0f} mes faturave të çiftuara.", metoda)

    for i in vetem_fin.index:
        r = fin.loc[i]
        if i in anulohen_idx_f:
            status_fin[i] = ST_ANULOHET
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], r['Vlera'], None, r['Vlera'], ST_ANULOHET,
                   'Dy rreshta me shumë të kundërt të të njëjtit furnitor; neto 0.')
        elif cutoff is not None and pd.notna(r['Data']) and r['Data'] > cutoff:
            status_fin[i] = ST_PAS
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], r['Vlera'], None, r['Vlera'], ST_PAS,
                   f"Data pas {fmt_data(cutoff)} (data e fundit në librin SC).")
        else:
            status_fin[i] = ST_MUNGON_PORTAL
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], r['Vlera'], None, r['Vlera'], ST_MUNGON_PORTAL)

    for j in vetem_port.index:
        if j in kredi_perdorur:
            continue
        r = port.loc[j]
        if j in anulohen_idx_p:
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], None, r['Vlera'], -r['Vlera'], ST_ANULOHET,
                   'Dy rreshta me shumë të kundërt të të njëjtit furnitor; neto 0.')
        elif viti_periudhes is not None and pd.notna(r['Data']) and r['Data'].year < viti_periudhes:
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], None, r['Vlera'], -r['Vlera'], ST_VIT,
                   'Faturë me datë të vitit të kaluar, vetëm në Portal.')
        else:
            rresht(r['NrFature'], r['NumriSerial'], r['Data'], r['Furnitori'], r['NIPT'], None, r['Vlera'], -r['Vlera'], ST_MUNGON_FIN)

    fat = pd.DataFrame(rows, columns=['NrFature', 'NIVF', 'Data', 'Furnitori', 'NIPT', 'Vlera_Financa', 'Vlera_Portal',
                                      'Efekti', 'Kush_ka_me_shume', 'Statusi', 'Metoda', 'Shenim'])

    # --- Ura e blerjeve ---
    total_diff = fin['Vlera'].sum() - port['Vlera'].sum()
    kategorite = [
        ('Pas datës së librit SC', ST_PAS), ('Vit i mëparshëm', ST_VIT),
        ('Vetëm në Portal (e vërtetë)', ST_MUNGON_FIN), ('Vetëm në Financa para cutoff', ST_MUNGON_PORTAL),
        ('Çifte që anulohen', ST_ANULOHET), ('Diferenca vlere mes fatura të çiftuara', ST_DIFERENCE),
        ('Regjistrim neto', ST_NETO),
    ]
    ura_rows = []
    for emri, st in kategorite:
        sub = fat[fat['Statusi'] == st]
        ura_rows.append({'Kategoria': emri, 'Numri': len(sub), 'Vlera': sub['Efekti'].sum(), 'Status': st})
    ura_rows.append({'Kategoria': f'Rrumbullakime ≤ {tol} lekë', 'Numri': n_rrumb, 'Vlera': rrumbullakime, 'Status': ST_OK})
    shpjeguar = sum(r['Vlera'] for r in ura_rows)
    e_pashpjeguar = total_diff - shpjeguar
    ura_rows += [
        {'Kategoria': 'Totali i shpjeguar', 'Numri': None, 'Vlera': shpjeguar, 'Status': ''},
        {'Kategoria': 'Diferenca totale (Financa − Portal)', 'Numri': None, 'Vlera': total_diff, 'Status': ''},
        {'Kategoria': 'E pashpjeguar (duhet 0)', 'Numri': None, 'Vlera': e_pashpjeguar,
         'Status': ST_OK if abs(e_pashpjeguar) <= tol else ST_PA_SHPJEGIM},
    ]
    ura = pd.DataFrame(ura_rows)

    kontroll = pd.DataFrame([
        {'Pershkrimi': f"Libri SC mbulon deri më {fmt_data(cutoff)}", 'Numri': None, 'Shuma': None},
        {'Pershkrimi': 'Financa - blerje', 'Numri': len(fin), 'Shuma': fin['Vlera'].sum()},
        {'Pershkrimi': 'Portali - blerje', 'Numri': len(port), 'Shuma': port['Vlera'].sum()},
        {'Pershkrimi': 'Diferenca (Financa − Portal)', 'Numri': None, 'Shuma': total_diff},
        {'Pershkrimi': 'Fatura të çiftuara', 'Numri': len(cifte), 'Shuma': None},
        {'Pershkrimi': 'Vetëm në Financa', 'Numri': len(vetem_fin), 'Shuma': vetem_fin['Vlera'].sum()},
        {'Pershkrimi': 'Vetëm në Portal', 'Numri': len(vetem_port), 'Shuma': vetem_port['Vlera'].sum()},
    ])

    kontrolle = _kontrolle(fin, port, cifte_info, f_cift, tol)
    kategori_tvsh, tvsh_totale = _kategori_tvsh(fin, port, cifte_info, tol)
    furnitoret = _furnitoret(fin, port, fat, tol)
    libri_me_sc, pa_magazine = _libri_me_sc(fin, port, f_cift, status_fin, eksport_financa_path, tol)
    if eksport_financa_path is not None and len(libri_me_sc):
        nuk = int((libri_me_sc['Magazina'] == 'NUK GJENDET NË SKEDAR').sum())
        if nuk:
            paralajmerime.append(f"{nuk} fatura të librit të Financa-s nuk u gjetën në skedarin e eksportit (Faturime blerje); "
                                 "për to nuk dihet nëse janë me magazinë.")

    rez = {
        'faturat': fat, 'ura': ura, 'kontroll': kontroll, 'kontrolle': kontrolle,
        'kategori_tvsh': kategori_tvsh, 'tvsh_totale': tvsh_totale, 'furnitoret': furnitoret, 'libri_me_sc': libri_me_sc, 'pa_magazine': pa_magazine,
        'diferenca': total_diff, 'e_pashpjeguar': e_pashpjeguar, 'cutoff': cutoff,
        'paralajmerime': paralajmerime, 'n_fin': len(fin), 'n_port': len(port),
    }
    rez['veprime'] = _veprimet(rez)
    rez['permbledhje'] = _permbledhja(rez)
    return rez


def _furnitoret(fin, port, fat, tol):
    """Një rresht për çdo furnitor (sipas NIPT): fatura dhe shuma në dy anët, diferenca, sa fatura nuk përputhen."""
    def celesi(df):
        return df['NIPT'].fillna('pa NIPT: ' + df['Furnitori'].fillna('?').astype(str))
    f, p = fin.assign(K=celesi(fin)), port.assign(K=celesi(port))
    emri = pd.concat([f[['K', 'Furnitori']], p[['K', 'Furnitori']]]).dropna().groupby('K')['Furnitori'].agg(lambda x: x.mode().iloc[0])
    gf = f.groupby('K')['Vlera'].agg(['count', 'sum']).rename(columns={'count': 'Fatura_Financa', 'sum': 'Shuma_Financa'})
    gp = p.groupby('K')['Vlera'].agg(['count', 'sum']).rename(columns={'count': 'Fatura_Portal', 'sum': 'Shuma_Portal'})
    t = pd.concat([gf, gp], axis=1).fillna(0)
    t['Diferenca'] = t['Shuma_Financa'] - t['Shuma_Portal']
    t['Kush_ka_me_shume'] = t['Diferenca'].map(kush_ka_me_shume)
    fat_k = fat.assign(K=fat['NIPT'].fillna('pa NIPT: ' + fat['Furnitori'].fillna('?').astype(str)))
    jo_ok = fat_k[~fat_k['Statusi'].isin([ST_ANULOHET, ST_NETO])].groupby('K').size()
    t['Fatura_Jo_OK'] = jo_ok.reindex(t.index).fillna(0).astype(int)
    t.insert(0, 'Furnitori', emri.reindex(t.index))
    t.insert(1, 'NIPT', [k if not str(k).startswith('pa NIPT') else None for k in t.index])
    t['Status'] = t.apply(lambda r: ST_OK if abs(r['Diferenca']) <= tol and r['Fatura_Jo_OK'] == 0 else
                          (ST_DIFERENCE if abs(r['Diferenca']) > tol else ST_PA_SHPJEGIM), axis=1)
    t['Status'] = t['Status'].replace({ST_PA_SHPJEGIM: 'KONTROLLO FATURAT'})
    t['_a'] = t['Diferenca'].abs()
    t = t.sort_values(['_a', 'Fatura_Jo_OK'], ascending=False).drop(columns='_a').reset_index(drop=True)
    return t


def _nr_norm(x):
    """Numri i faturës i normalizuar për krahasim (pa zero përpara, pa dallim shkronjash); datat e shtrembëruara japin ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    t = str(x).strip().lower()
    if re.match(r'^\d{4}-\d{2}-\d{2}', t):
        return ''
    if re.fullmatch(r'0*\d+', t):
        t = t.lstrip('0') or '0'
    return t


def _kod(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def parse_eksport_blerje(path):
    """
    Eksporti i faturave të blerjes nga Financa (një rresht për çdo artikull/llogari të faturës).
    TIPKLL = 'K' është artikull (me magazinë, KMAG); TIPKLL = 'L' është llogari kontabël (pa magazinë):
    KARTLLG = numri i llogarisë, PERSHKRIM = përshkrimi. Kthen një rresht për faturë.
    """
    if hasattr(path, 'seek'):
        path.seek(0)
    d = pd.read_excel(path)
    kerkohen = ['NRRENDOR', 'NRDOK', 'NRDSHOQ', 'NIPT', 'SHENIM1', 'VLERTOT', 'TIPKLL', 'KARTLLG', 'PERSHKRIM']
    mungon = [c for c in kerkohen if c not in d.columns]
    if mungon:
        raise ValueError(f"Eksporti i Financa-s nuk ka kolonat: {', '.join(mungon)}.")
    rows = []
    for nr, g in d.groupby('NRRENDOR', sort=False):
        r = g.iloc[0]
        llogari = g[g['TIPKLL'].astype(str).str.strip() == 'L']
        artikuj = g[g['TIPKLL'].astype(str).str.strip() == 'K']
        kursi = r.get('KURSI')
        eur = str(r.get('KMON', '')).strip().upper() == 'EUR' and pd.notna(kursi) and kursi
        vlera = pd.to_numeric(r['VLERTOT'], errors='coerce')
        porosi = {}
        for _, l in llogari.iterrows():
            k = (_kod(l['KARTLLG']), str(l['PERSHKRIM']).strip())
            porosi[k] = porosi.get(k, 0) + (pd.to_numeric(l.get('VLERABS'), errors='coerce') or 0)
        magazina = sorted({str(m).strip() for m in artikuj.get('KMAG', pd.Series(dtype=object)).dropna()})
        nipt = r['NIPT']
        rows.append({
            'NrDok': _kod(r['NRDOK']), 'NrDshoq': r['NRDSHOQ'], 'Furnitori': r['SHENIM1'],
            'NIPT': str(nipt).strip().upper() if pd.notna(nipt) else '',
            'Vlera': vlera, 'Vlera_Lek': vlera * float(kursi) if eur else vlera,
            'Magazina': 'ME MAGAZINË' if len(artikuj) else ('PA MAGAZINË' if len(llogari) else 'PA RRESHTA'),
            'Magazinat': ', '.join(magazina),
            'Llogaria': '; '.join(dict.fromkeys(k[0] for k in porosi)),
            'Pershkrimi': '; '.join(k[1] for k in porosi),
        })
    return pd.DataFrame(rows)


def _cifto_eksport(fin, inv, tol=TOLERANCE):
    """Çifton rreshtat e librit të Financa-s me faturat e eksportit: NIPT + numër fature + vlerë; NIPT + vlerë; NIPT + numër."""
    inv = inv.assign(_nr=inv['NrDshoq'].map(_nr_norm))
    perdorur, res = set(), {}
    f_nipt = fin['NIPT'].fillna('').astype(str).str.upper()
    f_nr = fin['NrFature'].map(_nr_norm)

    def diff(i, j):
        return min(abs(fin.at[i, 'Vlera'] - inv.at[j, 'Vlera']), abs(fin.at[i, 'Vlera'] - inv.at[j, 'Vlera_Lek']))

    for faza in (1, 2, 3):
        for i in fin.index:
            if i in res:
                continue
            kand = [j for j in inv.index[inv['NIPT'] == f_nipt[i]] if j not in perdorur]
            if faza in (1, 3):
                kand = [j for j in kand if f_nr[i] and inv.at[j, '_nr'] == f_nr[i]]
            if faza in (1, 2):
                kand = [j for j in kand if diff(i, j) <= tol]
            if kand:
                j = min(kand, key=lambda j: diff(i, j))
                res[i] = (j, {1: 'NIPT + nr. fature + vlerë', 2: 'NIPT + vlerë', 3: 'NIPT + nr. fature (vlerë ndryshe)'}[faza])
                perdorur.add(j)
    return res


def _libri_me_sc(fin, port, f_cift, status_fin, eksport_path, tol):
    """Fleta 'Libri_Financa_me_SC': çdo rresht i librit të Financa-s me ekuivalenten nga Self Care dhe, nga eksporti i Financa-s, magazina/llogaria."""
    inv = parse_eksport_blerje(eksport_path) if eksport_path is not None else None
    ekspo = _cifto_eksport(fin, inv, tol) if inv is not None else {}
    rows = []
    for i, f in fin.iterrows():
        j = f_cift.get(i, (None, ''))[0]
        p = port.loc[j] if j is not None else None
        r = {
            'NrFature_Financa': f['NrFature'], 'NIVF_Financa': f['NumriSerial'] if isinstance(f['NumriSerial'], str) else 'pa NIVF',
            'Data': f['Data'], 'Furnitori': f['Furnitori'], 'NIPT': f['NIPT'], 'Vlera_Financa': f['Vlera'],
            'NrFature_SC': p['NrFature'] if p is not None else None,
            'NIVF_SC': p['NumriSerial'] if p is not None else None,
            'Data_SC': p['Data'] if p is not None else None,
            'Vlera_SC': p['Vlera'] if p is not None else None,
            'Diferenca': (f['Vlera'] - p['Vlera']) if p is not None else None,
            'Metoda_Ciftimit': f_cift[i][1] if p is not None else None,
            'Status': status_fin.get(i, ''),
        }
        if inv is not None:
            e = inv.loc[ekspo[i][0]] if i in ekspo else None
            r.update({
                'Magazina': e['Magazina'] if e is not None else 'NUK GJENDET NË SKEDAR',
                'Magazinat': e['Magazinat'] if e is not None else None,
                'Nr_Dok_Financa': e['NrDok'] if e is not None else None,
                'Llogaria': (e['Llogaria'] or None) if e is not None else None,
                'Pershkrimi_Llogarise': (e['Pershkrimi'] or None) if e is not None else None,
            })
        rows.append(r)
    df = pd.DataFrame(rows)
    if inv is None or df.empty:
        return df, pd.DataFrame(columns=['Llogaria', 'Pershkrimi_Llogarise', 'Fatura', 'Vlera_Financa'])
    pa = df[df['Magazina'] == 'PA MAGAZINË']
    pa_llog = (pa.groupby(['Llogaria', 'Pershkrimi_Llogarise'], dropna=False)
               .agg(Fatura=('NrFature_Financa', 'size'), Vlera_Financa=('Vlera_Financa', 'sum')).reset_index()
               .sort_values('Vlera_Financa', key=abs, ascending=False))
    return df, pa_llog


def _kontrolle(fin, port, cifte_info, f_cift, tol):
    rows = []
    # Serial i pavlefshëm në Financa
    port_serial_nga_fin = {i: port.at[j, 'NumriSerial'] for i, (j, _) in f_cift.items()}
    for i, r in fin.iterrows():
        if serial_i_vlefshem(r['NumriSerial']):
            continue
        s = r['NumriSerial']
        lloji = ('pa NIVF' if pd.isna(s) else
                 ('pa NIVF (zero + nr. fature)' if re.fullmatch(r'0*\d+', str(s)) else 'NIVF me karaktere të gabuara'))
        sug = port_serial_nga_fin.get(i)
        rows.append({'Kontrolli': 'Serial i pavlefshëm në Financa', 'NrFature': r['NrFature'], 'NIVF': s,
                     'Furnitori': r['Furnitori'], 'Vlera': r['Vlera'], 'Data': r['Data'],
                     'Detaje': lloji + (f". NIVF e sugjeruar nga Portali: {sug}" if isinstance(sug, str) else '')})
    # Dyfishim i mundshëm (i njëjti NIPT, vlerë, datë)
    for emri, df in (('Financa', fin), ('Portali', port)):
        d = df[df['NIPT'].notna()]
        for _, g in d.groupby(['NIPT', 'Vlera', 'Data']):
            if len(g) > 1:
                pa_nivf = int((~g['NumriSerial'].map(serial_i_vlefshem)).sum())
                rows.append({'Kontrolli': f'Dyfishim i mundshëm ({emri})',
                             'NrFature': ' dhe '.join(g['NrFature'].astype(str)), 'NIVF': ' / '.join(g['NumriSerial'].fillna('pa NIVF')),
                             'Furnitori': g['Furnitori'].iloc[0], 'Vlera': g['Vlera'].iloc[0], 'Data': g['Data'].iloc[0],
                             'Detaje': f"{len(g)} rreshta me të njëjtin NIPT, vlerë dhe datë" + (f"; {pa_nivf} pa NIVF të vlefshme" if pa_nivf else '')})
    # Numër i përsëritur
    d = fin[fin['NrFature'].notna()]
    for (nr, nipt), g in d.groupby(['NrFature', 'NIPT'], dropna=False):
        if len(g) > 1:
            rows.append({'Kontrolli': 'Numër i përsëritur i faturës (Financa)', 'NrFature': nr,
                         'NIVF': f"{len(g)} NIVF të ndryshme", 'Furnitori': g['Furnitori'].iloc[0], 'Vlera': g['Vlera'].sum(),
                         'Data': g['Data'].min(),
                         'Detaje': f"Numri '{nr}' del {len(g)} herë; kërkimi sipas numrit nuk mjafton, përdor NIVF."})
    # Datë e ndryshme
    for i, j, metoda, dd in cifte_info:
        df_, dp_ = fin.at[i, 'Data'], port.at[j, 'Data']
        if pd.notna(df_) and pd.notna(dp_) and df_ != dp_:
            kufi = (df_.year, df_.month) != (dp_.year, dp_.month)
            rows.append({'Kontrolli': 'Datë e ndryshme', 'NrFature': fin.at[i, 'NrFature'], 'NIVF': port.at[j, 'NumriSerial'],
                         'Furnitori': fin.at[i, 'Furnitori'], 'Vlera': fin.at[i, 'Vlera'], 'Data': df_,
                         'Detaje': f"Financa {fmt_data(df_)}, Portali {fmt_data(dp_)} ({(df_ - dp_).days:+d} ditë)"
                                   + ("; KALON KUFIRIN E MUAJIT (periudha TVSH ndryshon)" if kufi else '')})
    return pd.DataFrame(rows, columns=['Kontrolli', 'NrFature', 'NIVF', 'Furnitori', 'Vlera', 'Data', 'Detaje'])


def _kategori_tvsh(fin, port, cifte_info, tol):
    def kat(r):
        if abs(r['Tvsh20']) > tol:
            return '20%'
        if abs(r['TvshJashtZbritjes']) > tol:
            return 'Jashtë zbritjes'
        return 'Tjetër'
    rows = []
    for i, j, metoda, _ in cifte_info:
        f, p = fin.loc[i], port.loc[j]
        d = f['Tvsh20'] - p['Tvsh20']
        if abs(d) > tol or kat(f) != kat(p):
            rows.append({'NrFature': f['NrFature'], 'NIVF': p['NumriSerial'], 'Furnitori': f['Furnitori'], 'Vlera_Financa': f['Vlera'],
                         'Kategoria_Financa': kat(f), 'Kategoria_Portal': kat(p),
                         'TVSH_Zbritshme_Financa': f['Tvsh20'], 'TVSH_Zbritshme_Portal': p['Tvsh20'], 'Diferenca_TVSH': d,
                         'Ku_zbritet_me_shume': 'Financa' if d > tol else ('Portali' if d < -tol else ''),
                         'Statusi': ST_KATEGORI})
    df = pd.DataFrame(rows, columns=['NrFature', 'NIVF', 'Furnitori', 'Vlera_Financa', 'Kategoria_Financa', 'Kategoria_Portal',
                                     'TVSH_Zbritshme_Financa', 'TVSH_Zbritshme_Portal', 'Diferenca_TVSH', 'Ku_zbritet_me_shume', 'Statusi'])
    tot = {'Financa': fin['Tvsh20'].sum(), 'Portali': port['Tvsh20'].sum()}
    tot['Diferenca'] = tot['Financa'] - tot['Portali']
    return df, tot


def _veprimet(rez):
    v = []
    f, k, kt = rez['faturat'], rez['kontrolle'], rez['kategori_tvsh']
    n_pas = int((f['Statusi'] == ST_PAS).sum())
    if n_pas:
        v.append((1, f"Rigjenero librin SC pas {fmt_data(rez['cutoff'])}", f"{n_pas} fatura Financa me datë pas librit SC.", None))
    for _, r in f[f['Statusi'] == ST_MUNGON_FIN].iterrows():
        v.append((2, f"Fut/kontrollo te Financa faturën {r['NrFature']} (NIVF {r['NIVF']})", f"{r['Furnitori']}, {fmt_data(r['Data'])}. Nuk e gjej në librin e Financa-s të ngarkuar: kontrollo nëse është regjistruar në muaj tjetër ose nuk është futur ende.", r['Efekti']))
    for _, r in f[f['Statusi'] == ST_MUNGON_PORTAL].iterrows():
        v.append((2, f"Kontrollo në Portal faturën {r['NrFature']} (NIVF {r['NIVF']})", f"{r['Furnitori']}, {fmt_data(r['Data'])}", r['Efekti']))
    for _, r in f[f['Statusi'] == ST_DIFERENCE].iterrows():
        v.append((2, f"Kontrollo vlerën e faturës {r['NrFature']} (NIVF {r['NIVF']})",
                  f"{r['Furnitori']}: Financa {r['Vlera_Financa']:,.0f} kundrejt Portalit {r['Vlera_Portal']:,.0f}", r['Efekti']))
    for _, r in k[k['Kontrolli'] == 'Serial i pavlefshëm në Financa'].iterrows():
        sug = re.search(r'NIVF e sugjeruar nga Portali: (\S+)', r['Detaje'])
        if sug and sug.group(1) != r['NIVF']:
            v.append((3, f"Korrigjo NIVF-në te Financa {r['NrFature']}: {r['NIVF']} → {sug.group(1)}", r['Furnitori'], None))
    for _, r in k[k['Kontrolli'].str.startswith('Dyfishim')].iterrows():
        v.append((3, f"Kontrollo dyfishimin e {r['NrFature']}", f"{r['Furnitori']}, {r['Vlera']:,.0f}. {r['Detaje']}", None))
    for furn, g in kt.groupby('Furnitori'):
        v.append((4, f"Vendos kategorinë TVSH të {furn}", f"{len(g)} fatura; TVSH e zbritshme: Financa {g['TVSH_Zbritshme_Financa'].sum():,.0f}, "
                                                          f"Portali {g['TVSH_Zbritshme_Portal'].sum():,.0f}", g['Diferenca_TVSH'].sum()))
    for _, r in f[f['Statusi'] == ST_VIT].iterrows():
        v.append((5, f"Kontrollo periudhën e faturës së vitit të kaluar {r['NrFature']} (NIVF {r['NIVF']})",
                  f"{r['Furnitori']}, {fmt_data(r['Data'])}", r['Efekti']))
    return rakordim.rendit_veprimet(v)


def _permbledhja(rez):
    u = rez['ura'].set_index('Kategoria')['Vlera']
    koha = u['Pas datës së librit SC'] + u['Vit i mëparshëm']
    reale = u['Vetëm në Portal (e vërtetë)'] + u['Vetëm në Financa para cutoff'] + u['Diferenca vlere mes fatura të çiftuara']
    return [
        f"Konventa: Diferenca = Financa − Portal (pozitive = Financa ka më shumë). Libri SC mbulon deri më {fmt_data(rez['cutoff'])}.",
        f"Financa ka {rez['n_fin']} fatura, Portali {rez['n_port']}. Diferenca totale në blerje është {rez['diferenca']:,.0f} lekë.",
        f"Prej saj {koha:,.0f} lekë janë çështje kohe/periudhe (pas datës së librit SC, ose faturë e vitit të kaluar).",
        f"{reale:,.0f} lekë janë fatura që mungojnë ose kanë vlerë ndryshe dhe duhen kontrolluar.",
        f"Çifte që anulohen, regjistrim neto dhe rrumbullakime kanë efekt 0 ose të papërfillshëm.",
        ("Çdo lek është i shpjeguar (e pashpjeguar = 0)." if abs(rez['e_pashpjeguar']) <= TOLERANCE
         else f"Kujdes: mbeten {rez['e_pashpjeguar']:,.0f} lekë të pashpjeguara."),
        *([f"{int((rez['libri_me_sc']['Magazina'] == 'PA MAGAZINË').sum())} fatura të librit të Financa-s janë pa magazinë (regjistruar direkt në llogari); "
           f"shih fletën 'Libri_Financa_me_SC' (rreshtat me ngjyrë portokalli, me numrin dhe përshkrimin e llogarisë)."]
          if len(rez['libri_me_sc']) and 'Magazina' in rez['libri_me_sc'].columns else []),
        f"TVSH 20% e zbritshme: Financa {rez['tvsh_totale']['Financa']:,.0f}, Portali {rez['tvsh_totale']['Portali']:,.0f}, "
        f"diferenca {rez['tvsh_totale']['Diferenca']:,.0f}.",
    ]


def shkruaj_raportin_blerje(rez, out):
    perm = [{'Përmbledhje': l} for l in rez['permbledhje']]
    perm += [{'Përmbledhje': ''}, {'Përmbledhje': 'VEPRIMET KRYESORE (lista e plotë te fleta "Veprime për t\'u bërë")'}]
    for i, (_, r) in enumerate(rez['veprime'].head(8).iterrows(), 1):
        perm.append({'Përmbledhje': f"{i}. {r['Veprimi']}" + (f" — {r['Detaje']}" if r['Detaje'] else '')})
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        pd.DataFrame(perm).to_excel(w, sheet_name='Përmbledhje', index=False)
        rez['ura'].to_excel(w, sheet_name='Ura e blerjeve', index=False)
        rez['veprime'].to_excel(w, sheet_name="Veprime për t'u bërë", index=False)
        if rez['paralajmerime']:
            pd.DataFrame({'Paralajmerime': rez['paralajmerime']}).to_excel(w, sheet_name='Paralajmerime', index=False)
        rez['kontroll'].to_excel(w, sheet_name='Kontroll', index=False)
        rez['faturat'].rename(columns={'Statusi': 'Status'}).to_excel(w, sheet_name='Faturat_Jo_OK', index=False)
        rez['kontrolle'].to_excel(w, sheet_name='Kontrolle', index=False)
        rez['kategori_tvsh'].rename(columns={'Statusi': 'Status'}).to_excel(w, sheet_name='Kategori_TVSH', index=False)
        rez['furnitoret'].to_excel(w, sheet_name='Furnitoret', index=False)
        rez['libri_me_sc'].to_excel(w, sheet_name='Libri_Financa_me_SC', index=False)
    if hasattr(out, 'seek'):
        out.seek(0)
    format_workbook(out)


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('financa')
    ap.add_argument('portal')
    ap.add_argument('out')
    ap.add_argument('--cutoff')
    a = ap.parse_args()
    rez = rakordo_blerje(a.financa, fleta_default(a.financa, ('blerje', 'Purchase book')),
                         a.portal, fleta_default(a.portal, ('Purchase book', 'Purchases book')), cutoff=a.cutoff)
    shkruaj_raportin_blerje(rez, a.out)
    for p in rez['paralajmerime']:
        print('KUJDES:', p)
    for l in rez['permbledhje']:
        print(l)
    print(rez['ura'].to_string())
