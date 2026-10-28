"""
Rakordim mujor: Libri i Shitjeve Financa vs Portali Self Care
Identifikon automatikisht:
 1. Diferencat në faturat me subjekt të identifikuar (match sipas Numrit Serial/NIVF)
 2. Diferencat në totalet ditore të klientëve të rastësishëm
 3. Operatorin (pikën e shitjes) përgjegjës për diferencën, në ditët problematike
"""
import pandas as pd
import numpy as np
from itertools import combinations
import re
import sys
import unicodedata

TOLERANCE = 3  # lek, diferenca ≤ kësaj vlere konsiderohen rrumbullakime

OPERATORET = {
    'gb758zl658': 'ASTIR 1', 'br784ks652': 'ASTIR 2',
    'yb060bq544': 'BARIKADAVE 1', 'zj268ci235': 'BARIKADAVE 2',
    'qs863kw091': 'BLLOKU 1', 'sj061qs812': 'BLLOKU 2',
}

# Statuset
ST_OK = 'OK'
ST_PRITJE = 'NË PRITJE FISKALIZIMI'
ST_PAS = 'PAS DATËS SË LIBRIT SC'
ST_MUNGON_PORTAL = 'MUNGON NË PORTAL'
ST_MUNGON_FIN = 'MUNGON NË FINANCA'
ST_STORNO = 'STORNO PA ÇIFT NË PORTAL'
ST_REALE = 'DIFERENCË REALE'
ST_PA_EKSPORT = 'PA EKSPORT'
ST_PA_SHPJEGIM = 'PA SHPJEGIM'

# Kolonat e kategorive te TVSH-se ne Librin e Shitjes (format zyrtar), sipas kodeve
# a,b,c,ç,d,dh,e,ë,f,g,gj,h,i,j,k,l,ll,m,n,nj,o,p,q -> indekset 0-22.
# 'e' (indeks 6, Vlera totale) tashme perfshihet si 'Vlera'; kategorite e meposhtme
# jane komponentet qe e perbejne ate total - nese nje shitje kalon gabimisht nga
# nje kategori ne tjetren me te njejten vlere, totali nuk e kap gabimin, prandaj
# krahasohen edhe keto ndaras.
KATEGORITE_TVSH = {
    7: 'Perjashtuar_e',
    8: 'PaTvsh_f',
    9: 'PaTvsh_g',
    10: 'PaTvsh_gj',
    11: 'Tatueshme20',
    12: 'Tvsh20',
    13: 'Tatueshme10',
    14: 'Tvsh10',
    15: 'Tatueshme6',
    16: 'Tvsh6',
    17: 'Import20_baze',
    18: 'Import20_tvsh',
    19: 'ImportTjeter_baze',
    20: 'ImportTjeter_tvsh',
    21: 'Tjeter_p',
    22: 'Tjeter_q',
}

MUAJT = ['Janar', 'Shkurt', 'Mars', 'Prill', 'Maj', 'Qershor',
         'Korrik', 'Gusht', 'Shtator', 'Tetor', 'Nëntor', 'Dhjetor']


def emri_muajit(periudha):
    return f"{MUAJT[periudha.month - 1]} {periudha.year}" if periudha is not None else '?'


def fleta_default(path, te_preferuara=('shitje', 'Sales book')):
    if hasattr(path, 'seek'):
        path.seek(0)
    fletet = pd.ExcelFile(path).sheet_names
    for f in te_preferuara:
        if f in fletet:
            return f
    return fletet[0]


def parse_libri(path, sheet_name):
    """
    Lexon librin e shitjeve (F5 ose Self Care). Libri F5 mund te kete nje kolone
    'Koment' para kolones se numrit te fatures, keshtu qe kolona 'a' zbulohet vete.
    Asnje rresht nuk hiqet sipas serialit: rreshtat pa serial duhet te mbeten,
    perndryshe totali del gabim.
    """
    if hasattr(path, 'seek'):
        path.seek(0)
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    offset, code_row_idx = None, None
    for off in (0, 1):
        match = raw[raw[off].astype(str).str.strip() == 'a'].index
        if len(match):
            offset, code_row_idx = off, match[0]
            break
    if offset is None:
        raise ValueError(f"Fleta '{sheet_name}' nuk duket si Liber Shitjesh (nuk u gjet rreshti i kodeve a, b, c...).")

    # Vetem kolona e numrit te fatures (dhe 'Koment' kur ekziston): seriali i rreshtave
    # kumulative te SC fillon me 'SHUMADAT-' dhe nuk duhet ngaterruar me 'Shuma totale'.
    first_cols = raw.iloc[:, :offset+1].astype(str)
    is_shuma = first_cols.apply(lambda c: c.str.strip().str.lower().str.startswith('shuma')).any(axis=1)
    stop_candidates = raw.index[is_shuma]
    stop_candidates = stop_candidates[stop_candidates > code_row_idx]
    stop_idx = stop_candidates[0] if len(stop_candidates) else len(raw)

    kat_cols = [offset + c for c in KATEGORITE_TVSH.keys()]
    base_cols = [offset + c for c in range(7)]
    data = raw.iloc[code_row_idx+1:stop_idx, base_cols + kat_cols].copy()
    data.columns = ['NrFature','NumriSerial','Data','Bleresi','Rrethi','NIPT','Vlera'] + list(KATEGORITE_TVSH.values())
    data = data.dropna(how='all')
    data['Data'] = pd.to_datetime(data['Data'], errors='coerce', dayfirst=True).dt.normalize()
    data['Vlera'] = pd.to_numeric(data['Vlera'], errors='coerce')
    for col in KATEGORITE_TVSH.values():
        data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)
    for col in ('NIPT', 'NumriSerial'):
        data[col] = data[col].astype(str).str.strip()
        data.loc[data[col].isin(['nan','None','NaN','']), col] = np.nan
    data = data.reset_index(drop=True)

    shuma_totale = None
    if stop_idx < len(raw):
        shuma_totale = pd.to_numeric(raw.iloc[stop_idx, offset+6], errors='coerce')
        if pd.isna(shuma_totale):
            shuma_totale = None
    data.attrs['shuma_totale'] = shuma_totale
    return data


def muaji_kryesor(df):
    periudhat = df['Data'].dropna().dt.to_period('M')
    return periudhat.mode().iloc[0] if len(periudhat) else None


def kontrollo_periudhat(fin, port, port_prev=None, fin_next=None):
    """Paralajmerime kur skedaret e ngarkuar nuk i perkasin muajve te pritur."""
    paralajmerime = []
    m_fin, m_port = muaji_kryesor(fin), muaji_kryesor(port)
    if m_fin is not None and m_port is not None and m_fin != m_port:
        paralajmerime.append(
            f"Skedarët duket se mbulojnë periudha të ndryshme: Financa është kryesisht {emri_muajit(m_fin)}, "
            f"Portali {emri_muajit(m_port)}. Rezultati do të nxjerrë si të munguara faturat e muajve që nuk përputhen."
        )
    else:
        fin_s = set(fin['NumriSerial'].dropna())
        port_s = set(port['NumriSerial'].dropna())
        if min(len(fin_s), len(port_s)) > 20:
            perputhje = len(fin_s & port_s) / min(len(fin_s), len(port_s))
            if perputhje < 0.5:
                paralajmerime.append(
                    f"Vetëm {perputhje:.0%} e serialeve përputhen mes dy librave. "
                    "Kontrollo nëse janë skedarët e duhur (i njëjti biznes, i njëjti muaj)."
                )
    if port_prev is not None and m_port is not None:
        m_prev = muaji_kryesor(port_prev)
        if m_prev is not None and m_prev != m_port - 1:
            paralajmerime.append(
                f"Self Care i muajit paraardhës është {emri_muajit(m_prev)}, por pritej {emri_muajit(m_port - 1)}."
            )
    if fin_next is not None and m_port is not None:
        m_next = muaji_kryesor(fin_next)
        if m_next is not None and m_next != m_port + 1:
            paralajmerime.append(
                f"Libri F5 i muajit pasardhës është {emri_muajit(m_next)}, por pritej {emri_muajit(m_port + 1)}."
            )
    for emri, df in (('Financa', fin), ('Portali', port)):
        shuma = df.attrs.get('shuma_totale')
        if shuma is not None and abs(df['Vlera'].sum() - shuma) > TOLERANCE:
            paralajmerime.append(
                f"{emri}: totali i rreshtave të lexuar ({df['Vlera'].sum():,.0f}) ndryshon nga rreshti "
                f"'Shuma totale' i librit ({shuma:,.0f}). Ndonjë rresht mund të mos jetë lexuar."
            )
    return paralajmerime


def krahaso_kategorite_tvsh(fin, port, tol=TOLERANCE):
    """
    Krahason totalet ditore te cdo kategorie TVSH-je (jo vetem totalin final).
    Kap rastet kur nje shitje kalon gabimisht nga nje kategori (p.sh. e perjashtuar
    nga TVSH) ne nje tjeter (p.sh. e tatueshme 20%) me te njejten vlere - gabim qe
    krahasimi i totalit te vetem nuk e sheh.
    """
    kategori_cols = list(KATEGORITE_TVSH.values())
    fin_by_date = fin.groupby('Data')[kategori_cols].sum()
    port_by_date = port.groupby('Data')[kategori_cols].sum()

    rows = []
    for data_val in sorted(set(fin_by_date.index) | set(port_by_date.index)):
        fin_row = fin_by_date.loc[data_val] if data_val in fin_by_date.index else pd.Series(0, index=kategori_cols)
        port_row = port_by_date.loc[data_val] if data_val in port_by_date.index else pd.Series(0, index=kategori_cols)
        for kat in kategori_cols:
            diff = fin_row[kat] - port_row[kat]
            if abs(diff) > tol:
                rows.append({
                    'Data': data_val, 'Kategoria': kat,
                    'Financa': fin_row[kat], 'Portal': port_row[kat], 'Diferenca': diff
                })
    return pd.DataFrame(rows, columns=['Data','Kategoria','Financa','Portal','Diferenca'])


def _norm(s):
    s = unicodedata.normalize('NFKD', str(s).lower())
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r'[^a-z0-9]', '', s)


def fmt_data(d):
    return d.strftime('%d.%m.%Y') if d is not None and not pd.isna(d) else ''


def kush_ka_me_shume(d, tol=TOLERANCE):
    if d is None or pd.isna(d):
        return ''
    if d > tol:
        return 'Financa ka më shumë'
    if d < -tol:
        return 'Portali ka më shumë'
    return ''


def parse_transaksionet(path):
    """
    Lexon eksportin e transaksioneve të fiskalizuara (fleta 'data'). Kolonat gjenden sipas
    emrit (pa dallim shkronjash/theksesh). Kthen kolona të standardizuara:
    DataFisk, DataLeshim, Operator, Vlera, Bleres_ID, NIVF, NrRendor.
    """
    if hasattr(path, 'seek'):
        path.seek(0)
    xls = pd.ExcelFile(path)
    sheet = 'data' if 'data' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    kol = {_norm(c): c for c in df.columns}

    def gjej(*opsione, kerkohet=False):
        for o in opsione:
            if o in kol:
                return kol[o]
        for o in opsione:
            for n, c in kol.items():
                if o in n:
                    return c
        if kerkohet:
            raise ValueError(f"Eksporti i transaksioneve nuk ka kolonën '{opsione[0]}'. Kolonat: {list(df.columns)}")
        return None

    c_fisk = gjej('dataefiskalizimit', 'fiskalizimit', kerkohet=True)
    c_leshim = gjej('dataeleshimit', 'leshimit')
    c_op = gjej('kodiioperatorit', 'operatorit', kerkohet=True)
    c_vlera = gjej('vleratotalemetvsh', 'vleratotale', kerkohet=True)
    c_bleres = gjej('identifikimittebleresit')
    c_nivf = gjej('nivf')
    c_rendor = gjej('numrirendor', 'rendor')

    out = pd.DataFrame({
        'DataFisk': pd.to_datetime(df[c_fisk], errors='coerce', dayfirst=True).dt.normalize(),
        'Operator': df[c_op].astype(str).str.strip(),
        'Vlera': pd.to_numeric(df[c_vlera], errors='coerce'),
    })
    out['DataLeshim'] = (pd.to_datetime(df[c_leshim], errors='coerce', dayfirst=True).dt.normalize()
                         if c_leshim else out['DataFisk'])
    if c_leshim:
        out['DataLeshim'] = out['DataLeshim'].fillna(out['DataFisk'])
    if c_bleres:
        b = df[c_bleres].astype(str).str.strip()
        out['Bleres_ID'] = b.where(~b.isin(['nan', 'None', 'NaN', '']), np.nan)
    else:
        out['Bleres_ID'] = np.nan
    out['NIVF'] = (df[c_nivf].astype(str).str.strip().str.lower().replace({'nan': np.nan, 'none': np.nan, '': np.nan})
                   if c_nivf else np.nan)
    out['NrRendor'] = df[c_rendor] if c_rendor else np.nan
    out.attrs['pa_leshim'] = c_leshim is None
    return out


def bashko_transaksionet(lista):
    """Bashkon eksportet ditore dhe heq dublikatat (sipas NIVF, ose rreshtit të plotë kur s'ka NIVF)."""
    if not lista:
        return None
    trans = pd.concat(lista, ignore_index=True)
    me_nivf = trans['NIVF'].notna()
    trans = pd.concat([
        trans[me_nivf].drop_duplicates('NIVF'),
        trans[~me_nivf].drop_duplicates(['DataFisk', 'DataLeshim', 'Operator', 'Vlera', 'NrRendor']),
    ], ignore_index=True)
    trans.attrs['pa_leshim'] = any(t.attrs.get('pa_leshim') for t in lista)
    return trans


def _pare_bleres_operator(fin_by_bleresi, fisk_by_code):
    """Çifton pikën e shitjes (Blerësi në Financa) me kodin e operatorit, sipas vlerave më të afërta."""
    pairs = sorted(
        ((abs(fv - pv), bleresi, code)
         for bleresi, fv in fin_by_bleresi.items() for code, pv in fisk_by_code.items()),
        key=lambda x: x[0])
    m_bleresi, m_code = {}, {}
    for _, bleresi, code in pairs:
        if bleresi in m_bleresi or code in m_code:
            continue
        m_bleresi[bleresi] = code
        m_code[code] = bleresi
    return m_bleresi, m_code


def gjej_kandidate(ex_op, shuma, tol=1, max_rez=5, max_fatura=2000):
    """
    Faturat e eksportit (të një operatori e dite) që shpjegojnë 'shuma' (±tol): fillimisht një faturë
    e vetme, përndryshe çifte. Kthen listë tekstesh me numër rendor dhe NIVF.
    """
    def etiketa(r):
        nr = r['NrRendor']
        nr = '?' if pd.isna(nr) else (int(nr) if float(nr).is_integer() else nr)
        nivf = r['NIVF'] if pd.notna(r['NIVF']) else 'pa NIVF'
        return f"nr. {nr}, NIVF {nivf}, {r['Vlera']:,.0f}"

    ex_op = ex_op.dropna(subset=['Vlera']).reset_index(drop=True)
    singles = ex_op[(ex_op['Vlera'] - shuma).abs() <= tol]
    if len(singles):
        return [etiketa(r) for _, r in singles.head(max_rez).iterrows()]
    if len(ex_op) > max_fatura:
        return []
    rend = ex_op.sort_values('Vlera').reset_index(drop=True)
    vals = rend['Vlera'].values
    gjetur = []
    for i in range(len(vals)):
        lo = np.searchsorted(vals, shuma - vals[i] - tol, side='left')
        hi = np.searchsorted(vals, shuma - vals[i] + tol, side='right')
        for j in range(max(lo, i + 1), hi):
            gjetur.append(f"çift: [{etiketa(rend.iloc[i])}] + [{etiketa(rend.iloc[j])}]")
            if len(gjetur) >= max_rez:
                return gjetur
    return gjetur


def analizo_ditet_operatori(cmp, fin_rand, trans, cutoff, tol=TOLERANCE):
    """
    Për çdo ditë me eksport: Financa / Fiskalizuar (eksport) / Në Portal (libri SC) për çdo operator.
    'Në Portal' për operator = Fiskalizuar − faturat e fiskalizuara pas cutoff (ato s'janë në librin SC).
    Kthen (detaje DataFrame, dict ditë -> shuma në pritje e shpjeguar).
    """
    rows, pritje_dita, ditet_me_eksport = [], {}, set()
    trans_rand = trans[trans['Bleres_ID'].isna()] if trans is not None else None
    if trans_rand is not None:
        ditet_me_eksport = set(trans_rand['DataLeshim'].dropna().unique())

    for _, cr in cmp.iterrows():
        data_val = cr['Data']
        fin_day = fin_rand[fin_rand['Data'] == data_val]
        fin_by = fin_day.groupby('Bleresi')['Vlera'].sum().to_dict()
        pas_cutoff = cutoff is not None and data_val > cutoff
        ka_eksport = data_val in ditet_me_eksport

        if not ka_eksport:
            d = cr['Diferenca']
            if abs(d) <= tol and not pas_cutoff:
                continue
            rows.append({
                'Data': data_val, 'Kodi_Operatorit': None, 'Operatori': None,
                'Pika_e_shitjes_Financa': ', '.join(fin_by) if fin_by else None,
                'Financa': cr['Financa_Total'], 'Fiskalizuar_Eksport': None, 'Ne_Pritje': None,
                'Ne_Portal_SC': cr['Portal_Total'], 'Data_Fiskalizimit': None, 'Vonesa_Dite': None,
                'Diferenca': d, 'Kush_ka_me_shume': kush_ka_me_shume(d),
                'Status': ST_PAS if pas_cutoff else ST_PA_EKSPORT,
                'Shenim': ('Shitje pas datës së fundit të librit SC.' if pas_cutoff else
                           'Nuk ka eksport të ngarkuar për këtë ditë; ngarko eksportin për të gjetur operatorin.'),
            })
            continue

        ex = trans_rand[trans_rand['DataLeshim'] == data_val]
        fisk = ex.groupby('Operator')['Vlera'].sum().to_dict()
        pritje = ({} if cutoff is None else
                  ex[ex['DataFisk'] > cutoff].groupby('Operator')['Vlera'].sum().to_dict())
        m_bleresi, m_code = _pare_bleres_operator(fin_by, fisk)

        def shto_rresht(bleresi, code):
            fin_v = fin_by.get(bleresi, 0.0) if bleresi is not None else 0.0
            fisk_v = fisk.get(code, 0.0) if code is not None else 0.0
            prit_v = pritje.get(code, 0.0) if code is not None else 0.0
            portal_v = fisk_v - prit_v
            diff = fin_v - portal_v
            rezid = fin_v - fisk_v
            ex_op = ex[ex['Operator'] == code] if code is not None else ex.iloc[0:0]
            datat = sorted(ex_op['DataFisk'].dropna().unique())
            vonesa = (max(datat) - data_val).days if datat else None
            shenime = []
            if pas_cutoff:
                status = ST_PAS
            elif prit_v > tol:
                status = ST_PRITJE
                shenime.append(f"Fiskalizuar pas {fmt_data(cutoff)} (nuk është në librin SC).")
            elif abs(diff) <= tol:
                status = ST_OK
            else:
                status = ST_REALE
            if not pas_cutoff:
                if rezid < -tol:
                    kand = gjej_kandidate(ex_op, -rezid)
                    if kand:
                        shenime.append(f"Portali/fiskalizimi ka {-rezid:,.0f} më shumë se Financa. Kandidatët: " + '; '.join(kand))
                    else:
                        shenime.append(f"Portali/fiskalizimi ka {-rezid:,.0f} më shumë se Financa; asnjë faturë e vetme ose çift nuk e shpjegon.")
                    if status == ST_PRITJE:
                        shenime.append("Diferenca e fshehur nën pritje do të dalë si diferencë reale pasi të hyjë në Portal.")
                elif rezid > tol:
                    shenime.append("Financa ka më shumë se fiskalizimi, kontrollo dyfishim ose faturë të pafiskalizuar.")
            if code is None:
                shenime.append("Pika e shitjes në Financa nuk përputhet me asnjë operator të eksportit.")
            if bleresi is None:
                shenime.append("Operator pa pikë shitjeje të identifikuar në Financa.")
            rows.append({
                'Data': data_val, 'Kodi_Operatorit': code,
                'Operatori': OPERATORET.get(code, 'I panjohur') if code else None,
                'Pika_e_shitjes_Financa': bleresi, 'Financa': fin_v, 'Fiskalizuar_Eksport': fisk_v,
                'Ne_Pritje': prit_v, 'Ne_Portal_SC': portal_v,
                'Data_Fiskalizimit': ', '.join(fmt_data(x) for x in datat) if datat else None,
                'Vonesa_Dite': vonesa, 'Diferenca': diff, 'Kush_ka_me_shume': kush_ka_me_shume(diff),
                'Status': status, 'Shenim': ' '.join(shenime),
            })
            return diff if prit_v > tol else 0.0

        pritje_bridge = 0.0
        for bleresi in fin_by:
            pritje_bridge += shto_rresht(bleresi, m_bleresi.get(bleresi))
        for code in fisk:
            if code not in m_code:
                pritje_bridge += shto_rresht(None, code)
        if not pas_cutoff:
            pritje_dita[data_val] = pritje_bridge

    kol = ['Data', 'Kodi_Operatorit', 'Operatori', 'Pika_e_shitjes_Financa', 'Financa', 'Fiskalizuar_Eksport',
           'Ne_Pritje', 'Ne_Portal_SC', 'Data_Fiskalizimit', 'Vonesa_Dite', 'Diferenca',
           'Kush_ka_me_shume', 'Status', 'Shenim']
    detaje = pd.DataFrame(rows, columns=kol)
    return detaje, pritje_dita


def zbulo_storno(fin, port, tol=TOLERANCE):
    """
    Çifte storno: faturë me shumë të kundërt për të njëjtin blerës (NIPT). Për çdo storno të Financa-s
    kontrollon nëse Portali e ka anën pozitive (por jo stornon).
    Kthen dict serial_storno -> shënim.
    """
    shenime = {}
    neg = fin[(fin['Vlera'] < 0) & fin['NIPT'].notna() & fin['NumriSerial'].notna()]
    for _, r in neg.iterrows():
        pos_port = port[(port['NIPT'] == r['NIPT']) & ((port['Vlera'] + r['Vlera']).abs() <= tol)]
        if len(pos_port) and not port['NumriSerial'].eq(r['NumriSerial']).any():
            shenime[r['NumriSerial']] = (
                f"Portali ka anën pozitive ({pos_port.iloc[0]['Vlera']:,.0f}) por jo stornon. "
                "Nëse storno u lëshua më datë tjetër, kontrollo datën e saj në Financa.")
        elif not port['NumriSerial'].eq(r['NumriSerial']).any():
            shenime[r['NumriSerial']] = ("Storno e Financa-s nuk gjendet në Portal. "
                                         "Nëse u lëshua më datë tjetër, kontrollo datën e saj në Financa.")
    return shenime


def find_operator_combo(target_diff, operator_totals, tol=TOLERANCE, max_combo=3):
    """Gjen kombinimin e operatorëve (deri në max_combo) që shpjegon diferencën."""
    items = list(operator_totals.items())
    for r in range(1, max_combo+1):
        for combo in combinations(items, r):
            s = sum(v for _, v in combo)
            if abs(s - target_diff) <= tol:
                return [name for name, _ in combo], s
    return None, None


def rakordo(financa_path, financa_sheet, portal_path, portal_sheet,
            portal_prev_path=None, fin_next_path=None, trans_paths=None, cutoff=None):
    """
    portal_prev_path: (opsionale) libri Self Care i muajit paraardhës - faturat F5 që
        gjenden aty shënohen "Deklaruar ne <muaji>" në vend të "Mungon në Portal".
    fin_next_path: (opsionale) libri F5 i muajit pasardhës - faturat SC që mungojnë në
        F5 por gjenden aty shënohen "<muaji> - del ne <muaji pasardhes> ne F5".
    trans_paths: (opsionale) lista e eksporteve të fiskalizimit (një ose më shumë ditë).
    cutoff: (opsionale) data e shkarkimit të Librit SC; nëse mungon, data më e madhe në librin SC.
    """
    fin = parse_libri(financa_path, financa_sheet)
    port = parse_libri(portal_path, portal_sheet)
    port_prev = parse_libri(portal_prev_path, fleta_default(portal_prev_path)) if portal_prev_path is not None else None
    fin_next = parse_libri(fin_next_path, fleta_default(fin_next_path)) if fin_next_path is not None else None
    trans = bashko_transaksionet([parse_transaksionet(t) for t in (trans_paths or [])])

    paralajmerime = kontrollo_periudhat(fin, port, port_prev, fin_next)
    m_port = muaji_kryesor(port)
    emri_prev = MUAJT[(m_port - 1).month - 1] if m_port is not None else 'muajin paraardhes'
    emri_tani = MUAJT[m_port.month - 1] if m_port is not None else 'Ky muaj'
    emri_next = MUAJT[(m_port + 1).month - 1] if m_port is not None else 'muajin pasardhes'

    if cutoff is not None:
        cutoff = pd.Timestamp(cutoff).normalize()
    else:
        cutoff = port['Data'].max()
        if pd.isna(cutoff):
            cutoff = None

    def pas_cutoff(d):
        return cutoff is not None and pd.notna(d) and d > cutoff

    if trans is not None and trans.attrs.get('pa_leshim'):
        paralajmerime.append("Eksporti i fiskalizimit nuk ka kolonën 'Data e lëshimit'; përdoret data e fiskalizimit "
                             "(ditët e fiskalizuara me vonesë nuk do të lidhen me ditën e shitjes).")

    nivf_fisk = {}
    if trans is not None:
        t = trans.dropna(subset=['NIVF']).drop_duplicates('NIVF')
        nivf_fisk = t.set_index('NIVF')['DataFisk'].to_dict()

    # --- 1. Subjekt i identifikuar ---
    fin_id = fin[fin['NIPT'].notna()]
    port_id = port[port['NIPT'].notna()]
    pa_serial = fin_id[fin_id['NumriSerial'].isna()].copy()
    kol_m = ['NumriSerial', 'NrFature', 'Data', 'Bleresi', 'NIPT', 'Vlera']
    merged = pd.merge(
        fin_id[fin_id['NumriSerial'].notna()][kol_m],
        port_id[port_id['NumriSerial'].notna()][kol_m],
        on='NumriSerial', how='outer', suffixes=('_Financa', '_Portal'), indicator=True
    )
    for col in ('NrFature', 'Data', 'Bleresi', 'NIPT'):
        merged[col] = merged[f'{col}_Financa'].combine_first(merged[f'{col}_Portal'])
    merged = merged[['NumriSerial', 'NrFature', 'Data', 'Bleresi', 'NIPT', 'Vlera_Financa', 'Vlera_Portal', '_merge']]
    merged['Diferenca'] = merged['Vlera_Financa'].fillna(0) - merged['Vlera_Portal'].fillna(0)

    prev_serials = set(port_prev['NumriSerial'].dropna()) if port_prev is not None else set()
    next_serials = set(fin_next['NumriSerial'].dropna()) if fin_next is not None else set()
    m_left, m_right = merged['_merge'] == 'left_only', merged['_merge'] == 'right_only'
    deklaruar_prev = m_left & merged['NumriSerial'].isin(prev_serials)
    del_next = m_right & merged['NumriSerial'].isin(next_serials)
    deklaruar_serials = set(merged.loc[deklaruar_prev, 'NumriSerial'])
    del_next_serials = set(merged.loc[del_next, 'NumriSerial'])

    next_rows = fin_next[fin_next['NumriSerial'].isin(del_next_serials)] if fin_next is not None else fin.iloc[0:0]
    next_val = next_rows.groupby('NumriSerial')['Vlera'].sum().to_dict()
    next_fill = next_rows.drop_duplicates('NumriSerial').set_index('NumriSerial')
    if len(next_fill):
        for col in ('Data', 'Bleresi', 'NIPT'):
            plotesim = merged['NumriSerial'].map(next_fill[col].to_dict())
            mask = del_next & plotesim.notna()
            merged.loc[mask, col] = plotesim[mask]

    storno_shenime = zbulo_storno(fin, port)

    statuset, problemet, shenimet, adj_diffs = [], [], [], []
    for idx, r in merged.iterrows():
        serial, tipi = r['NumriSerial'], r['_merge']
        adj_fin = 0.0 if serial in deklaruar_serials else (
            next_val.get(serial, 0.0) if serial in del_next_serials else (r['Vlera_Financa'] if pd.notna(r['Vlera_Financa']) else 0.0))
        adj_diffs.append(adj_fin - (r['Vlera_Portal'] if pd.notna(r['Vlera_Portal']) else 0.0))
        shenim = ''
        if serial in deklaruar_serials:
            problem, status = f"Deklaruar ne {emri_prev}", 'DEKLARUAR MUAJIN PARAARDHËS'
        elif serial in del_next_serials:
            problem, status = f"{emri_tani} - dalin ne {emri_next} ne F5", 'DALIN MUAJIN PASARDHËS NË F5'
        elif tipi == 'left_only':
            problem = 'Mungon në Portal'
            nk = str(serial).lower()
            if pas_cutoff(r['Data']):
                status = ST_PAS
            elif nk in nivf_fisk and pas_cutoff(nivf_fisk[nk]):
                status = ST_PRITJE
                shenim = f"Fiskalizuar më {fmt_data(nivf_fisk[nk])}, pas datës së librit SC."
            elif r['Vlera_Financa'] < 0:
                status = ST_STORNO
                shenim = storno_shenime.get(serial, "Nëse storno u lëshua më datë tjetër, kontrollo datën e saj në Financa.")
            else:
                status = ST_MUNGON_PORTAL
                if nk in nivf_fisk:
                    shenim = f"Fiskalizuar më {fmt_data(nivf_fisk[nk])} sipas eksportit; kontrollo pse nuk është në librin SC."
        elif tipi == 'right_only':
            problem, status = 'Mungon në Financa', ST_MUNGON_FIN
        else:
            problem = 'Vlerë e ndryshme'
            status = ST_REALE if abs(r['Diferenca']) > TOLERANCE else ST_OK
        statuset.append(status)
        problemet.append(problem if tipi != 'both' else 'Vlerë e ndryshme')
        shenimet.append(shenim)
    merged['Status'] = statuset
    merged['Problemi'] = problemet
    merged['Shenim'] = shenimet
    merged['AdjDiff'] = adj_diffs

    subjekt_diff = merged[(merged['Diferenca'].abs() > TOLERANCE) | (merged['_merge'] != 'both')].copy()
    subjekt_diff = subjekt_diff[~((subjekt_diff['_merge'] == 'both') & (subjekt_diff['Status'] == ST_OK))]
    subjekt_diff['NIVF'] = subjekt_diff['NumriSerial']
    subjekt_diff['Kush_ka_me_shume'] = subjekt_diff['Diferenca'].map(kush_ka_me_shume)

    if len(pa_serial):
        pa_rows = pd.DataFrame({
            'NumriSerial': 'pa NIVF', 'NIVF': 'pa NIVF', 'NrFature': pa_serial['NrFature'], 'Data': pa_serial['Data'],
            'Bleresi': pa_serial['Bleresi'], 'NIPT': pa_serial['NIPT'],
            'Vlera_Financa': pa_serial['Vlera'], 'Vlera_Portal': np.nan, 'Diferenca': pa_serial['Vlera'],
            'Problemi': 'Pa serial - nuk del ne SC', 'Status': 'PA NIVF (jashtë llogaritjes)',
            'Shenim': 'Faturë pa NIVF; nuk dalin në Portal, hiqet nga F5 e korrigjuar.',
            'Kush_ka_me_shume': '', '_merge': 'left_only', 'AdjDiff': 0.0,
        })
        subjekt_diff = pd.concat([subjekt_diff, pa_rows], ignore_index=True)

    deklaruar_rows = fin[fin['NumriSerial'].isin(deklaruar_serials)]

    # Libri F5 i korrigjuar per krahasim: - muaji paraardhes - pa serial + rreshtat qe dalin muajin pasardhes
    fin_adj = pd.concat([
        fin[~fin.index.isin(deklaruar_rows.index) & ~fin.index.isin(pa_serial.index)],
        next_rows,
    ], ignore_index=True)

    # --- 2. Klient i rastësishëm - totale ditore ---
    fin_rand = fin[fin['NIPT'].isna()]
    port_rand = port[port['NIPT'].isna()]
    fin_by_date = fin_rand.groupby('Data')['Vlera'].sum().rename('Financa_Total')
    port_by_date = port_rand.groupby('Data')['Vlera'].sum().rename('Portal_Total')
    cmp = pd.concat([fin_by_date, port_by_date], axis=1).fillna(0)
    cmp['Diferenca'] = cmp['Financa_Total'] - cmp['Portal_Total']
    cmp = cmp.reset_index()

    # --- 3. Operatorët: ditë me eksport ---
    detaje_operatori, pritje_dita = analizo_ditet_operatori(cmp, fin_rand, trans, cutoff)
    ditet_me_eksport = set(trans[trans['Bleres_ID'].isna()]['DataLeshim'].dropna().unique()) if trans is not None else set()

    # --- 4. Statusi ditor dhe ura e diferencës ---
    cmp_status, cmp_pritje, cmp_mbetur, cmp_shenim = [], [], [], []
    ura = {k: [0, 0.0] for k in ('pritje', 'pas', 'storno', 'mungon_portal', 'mungon_fin', 'pa_eksport', 'reale', 'rrumbullakime')}

    def shto(kat, vlera, n=1):
        ura[kat][0] += n
        ura[kat][1] += vlera

    for _, r in cmp.iterrows():
        d, data_val = r['Diferenca'], r['Data']
        pritje, mbetur, shenim = None, None, ''
        if pas_cutoff(data_val):
            status = ST_PAS
            shto('pas', d)
            shenim = 'Shitje pas datës së fundit të librit SC.'
        elif abs(d) <= TOLERANCE:
            status = ST_OK
            shto('rrumbullakime', d, 0)
        elif data_val in ditet_me_eksport:
            pritje = pritje_dita.get(data_val, 0.0)
            if abs(pritje) > TOLERANCE:
                shto('pritje', pritje)
            else:
                pritje = 0.0
            mbetur = d - pritje
            if abs(mbetur) <= TOLERANCE:
                status = ST_PRITJE if pritje else ST_OK
                shto('rrumbullakime', mbetur, 0)
            else:
                status = ST_REALE
                shto('reale', mbetur)
                if pritje:
                    shenim = f"Përfshin {pritje:,.0f} në pritje fiskalizimi; mbetja reale është {mbetur:,.0f}."
        else:
            status = ST_PA_EKSPORT
            shto('pa_eksport', d)
            shenim = 'Nuk ka eksport të ngarkuar për këtë ditë.'
        cmp_status.append(status)
        cmp_pritje.append(pritje)
        cmp_mbetur.append(mbetur)
        cmp_shenim.append(shenim)
    cmp['Ne_Pritje_Fiskalizimi'] = cmp_pritje
    cmp['Diferenca_e_Mbetur'] = cmp_mbetur
    cmp['Status'] = cmp_status
    cmp['Kush_ka_me_shume'] = cmp['Diferenca'].map(kush_ka_me_shume)
    cmp['Shenim'] = cmp_shenim

    # Rreshtat e subjektit të identifikuar
    for _, r in merged.iterrows():
        d, st = r['AdjDiff'], r['Status']
        if st == ST_PAS:
            shto('pas', d)
        elif st == ST_PRITJE:
            shto('pritje', d)
        elif st == ST_STORNO:
            shto('storno', d)
        elif st == ST_MUNGON_PORTAL:
            shto('mungon_portal', d)
        elif st == ST_MUNGON_FIN:
            shto('mungon_fin', d)
        elif abs(d) > TOLERANCE:
            shto('reale', d)
        else:
            ura['rrumbullakime'][1] += d

    kontroll_totali = fin_adj['Vlera'].sum() - port['Vlera'].sum()
    kategorite = [
        ('Në pritje fiskalizimi', ST_PRITJE, 'pritje', 'Fiskalizuar pas datës së shkarkimit të librit SC; hyjnë kur të rigjenerohet.'),
        ('Pas datës së librit SC', ST_PAS, 'pas', f"Shitje me datë pas {fmt_data(cutoff)}, ende jo në librin SC."),
        ('Storno pa çift në Portal', ST_STORNO, 'storno', 'Storno e Financa-s që Portali nuk e ka.'),
        ('Mungon në Portal', ST_MUNGON_PORTAL, 'mungon_portal', 'Faturë e identifikuar që nuk është në Portal.'),
        ('Mungon në Financa', ST_MUNGON_FIN, 'mungon_fin', 'Faturë e Portalit që Financa nuk e ka.'),
        ('Ditë pa eksport (e pa shpjeguar)', ST_PA_EKSPORT, 'pa_eksport', 'Ngarko eksportin e ditës për të gjetur operatorin.'),
        ('Diferenca reale', ST_REALE, 'reale', 'Mbetet pasi hiqen kohët dhe mungesat e mësipërme.'),
        ('Rrumbullakime ≤ ' + str(TOLERANCE) + ' lekë', ST_OK, 'rrumbullakime', ''),
    ]
    ura_rows = [{'Kategoria': k, 'Numri': ura[key][0], 'Vlera': ura[key][1], 'Status': st, 'Shenim': sh}
                for k, st, key, sh in kategorite]
    shpjeguar = sum(r['Vlera'] for r in ura_rows)
    e_pashpjeguar = kontroll_totali - shpjeguar
    ura_rows.append({'Kategoria': 'Totali i shpjeguar', 'Numri': None, 'Vlera': shpjeguar, 'Status': '', 'Shenim': ''})
    ura_rows.append({'Kategoria': 'Diferenca totale (Kontroll)', 'Numri': None, 'Vlera': kontroll_totali, 'Status': '', 'Shenim': ''})
    ura_rows.append({'Kategoria': 'E pashpjeguar (duhet 0)', 'Numri': None, 'Vlera': e_pashpjeguar,
                     'Status': ST_OK if abs(e_pashpjeguar) <= TOLERANCE else ST_PA_SHPJEGIM,
                     'Shenim': 'Kontroll: totali i Kontrollit minus kategoritë.'})
    ura_df = pd.DataFrame(ura_rows)

    kontroll = pd.DataFrame([
        {'Pershkrimi': f"Libri SC mbulon deri më {fmt_data(cutoff)}", 'Numri': None, 'Shuma': None, 'Shenim': ''},
        {'Pershkrimi': 'Financa (F5) - totali i librit', 'Numri': len(fin), 'Shuma': fin['Vlera'].sum(), 'Shenim': ''},
        {'Pershkrimi': f"(-) Deklaruar ne {emri_prev} (SC i muajit paraardhes)", 'Numri': len(deklaruar_rows), 'Shuma': -deklaruar_rows['Vlera'].sum(), 'Shenim': ''},
        {'Pershkrimi': '(-) Pa serial - nuk dalin ne SC', 'Numri': len(pa_serial), 'Shuma': -pa_serial['Vlera'].sum(), 'Shenim': ''},
        {'Pershkrimi': f"(+) Dalin ne {emri_next} ne F5", 'Numri': len(next_rows), 'Shuma': next_rows['Vlera'].sum(), 'Shenim': ''},
        {'Pershkrimi': 'Financa e korrigjuar', 'Numri': len(fin_adj), 'Shuma': fin_adj['Vlera'].sum(), 'Shenim': ''},
        {'Pershkrimi': 'Në Portal (libri SC) - totali', 'Numri': len(port), 'Shuma': port['Vlera'].sum(),
         'Shenim': 'Portali ka rreshta kumulativë ditorë për klientin e rastësishëm: numri i rreshtave nuk krahasohet me Financa-n.'},
        {'Pershkrimi': 'Diferenca (Financa e korrigjuar - Portali)', 'Numri': None, 'Shuma': kontroll_totali, 'Shenim': ''},
    ])

    # --- 5. Fleta Operatori_Problematik (ditët pa shpjegim të plotë) ---
    problem_rows = []
    for _, row in cmp[cmp['Status'].isin([ST_PA_EKSPORT, ST_REALE])].iterrows():
        data_val, diff = row['Data'], row['Diferenca']
        day_ops = fin_rand[fin_rand['Data'] == data_val].groupby('Bleresi')['Vlera'].sum().to_dict()
        combo, matched_sum = find_operator_combo(diff, day_ops)
        problem_rows.append({
            'Data': data_val, 'Diferenca': diff, 'Status': row['Status'],
            'Kush_ka_me_shume': kush_ka_me_shume(diff),
            'Operatori(et) i identifikuar': ', '.join(combo) if combo else
            'Asnjë pikë shitjeje e vetme nuk e shpjegon; ngarko eksportin e ditës për kandidatët',
            'Shuma e përputhur': matched_sum,
        })
    operator_report = pd.DataFrame(problem_rows)

    # --- 6. Diferenca sipas kategorive te TVSH-se ---
    kategori_diff = krahaso_kategorite_tvsh(fin_adj, port)

    rez = {
        'subjekt_diff': subjekt_diff.drop(columns=['_merge', 'AdjDiff']),
        'cmp': cmp,
        'operator_report': operator_report,
        'detaje_operatori': detaje_operatori,
        'fin_rand': fin_rand,
        'kategori_diff': kategori_diff,
        'kontroll': kontroll,
        'ura': ura_df,
        'diferenca': kontroll_totali,
        'e_pashpjeguar': e_pashpjeguar,
        'cutoff': cutoff,
        'paralajmerime': paralajmerime,
        'ka_eksport': trans is not None,
    }
    rez['subjekt_diff'] = rez['subjekt_diff'].rename(columns={'NrFature': 'NrFature_Financa'})
    rez['veprime'] = nderto_veprimet(rez, trans)
    rez['permbledhje'] = nderto_permbledhjen(rez)
    return rez


def nderto_veprimet(rez, trans):
    cutoff = rez['cutoff']
    v = []
    sd, det, cmp = rez['subjekt_diff'], rez['detaje_operatori'], rez['cmp']

    n_pritje = int((det['Status'] == ST_PRITJE).sum()) if len(det) else 0
    n_pritje += int((sd['Status'] == ST_PRITJE).sum())
    data_fisk = []
    if trans is not None and cutoff is not None:
        data_fisk = trans.loc[trans['DataFisk'] > cutoff, 'DataFisk'].tolist()
    n_pas = int((cmp['Status'] == ST_PAS).sum()) + int((sd['Status'] == ST_PAS).sum())
    if n_pritje:
        pas_data = fmt_data(max(data_fisk)) if data_fisk else 'datën e fundit të fiskalizimit'
        v.append((1, f"Rigjenero Librin SC pas {pas_data}",
                  f"{n_pritje} rreshta (operator/faturë) në pritje fiskalizimi.", None))
    if n_pas:
        v.append((1, f"Rigjenero Librin SC me datë të fundit pas {fmt_data(cutoff)}",
                  f"{n_pas} ditë/fatura me shitje pas datës së librit SC.", None))
    for _, r in sd[sd['Status'] == ST_MUNGON_PORTAL].iterrows():
        v.append((2, f"Kontrollo në Portal faturën {r['NrFature_Financa']} (NIVF {r['NIVF']})",
                  f"{r['Bleresi']}, {fmt_data(r['Data'])}. {r['Shenim']}".strip(), r['Diferenca']))
    for _, r in sd[sd['Status'] == ST_STORNO].iterrows():
        v.append((3, f"Kontrollo datën e stornos {r['NrFature_Financa']}",
                  f"{r['Bleresi']}, {fmt_data(r['Data'])}. {r['Shenim']}".strip(), r['Diferenca']))
    for _, r in sd[sd['Status'] == ST_MUNGON_FIN].iterrows():
        v.append((3, f"Fut/kontrollo te Financa faturën {r['NrFature_Financa']} (NIVF {r['NIVF']})",
                  f"{r['Bleresi']}, {fmt_data(r['Data'])}.", r['Diferenca']))
    for _, r in sd[(sd['Status'] == ST_REALE)].iterrows():
        v.append((4, f"Kontrollo vlerën e faturës {r['NrFature_Financa']} (NIVF {r['NIVF']})",
                  f"{r['Bleresi']}: Financa {r['Vlera_Financa']:,.0f} kundrejt Portalit {r['Vlera_Portal']:,.0f}.", r['Diferenca']))
    if len(det):
        for _, r in det[det['Status'] == ST_REALE].iterrows():
            v.append((4, f"Kontrollo {r['Operatori'] or r['Pika_e_shitjes_Financa']}, {fmt_data(r['Data'])}",
                      r['Shenim'], r['Diferenca']))
        for _, r in det[(det['Status'] == ST_PRITJE) & det['Shenim'].str.contains('Kandidatët', na=False)].iterrows():
            v.append((5, f"Pas hyrjes në Portal kontrollo {r['Operatori']}, {fmt_data(r['Data'])}", r['Shenim'], None))
    for _, r in cmp[cmp['Status'] == ST_PA_EKSPORT].iterrows():
        v.append((4, f"Ngarko eksportin e fiskalizimit të {fmt_data(r['Data'])}",
                  'Pa eksport nuk gjendet operatori përgjegjës.', r['Diferenca']))
    n_pa_niv = int((sd['Status'] == 'PA NIVF (jashtë llogaritjes)').sum())
    if n_pa_niv:
        v.append((5, f"{n_pa_niv} fatura pa NIVF në F5", 'Nuk dalin në Portal; kontrollo nëse duhej të kishin NIVF.', None))
    df = pd.DataFrame(v, columns=['Prioriteti', 'Veprimi', 'Detaje', 'Vlera']).sort_values('Prioriteti', kind='stable')
    return df.reset_index(drop=True)


def nderto_permbledhjen(rez):
    ura = rez['ura'].set_index('Kategoria')['Vlera']
    koha = ura['Në pritje fiskalizimi'] + ura['Pas datës së librit SC']
    reale = (ura['Diferenca reale'] + ura['Mungon në Portal'] + ura['Mungon në Financa']
             + ura['Storno pa çift në Portal'] + ura['Ditë pa eksport (e pa shpjeguar)'])
    rrum = ura['Rrumbullakime ≤ ' + str(TOLERANCE) + ' lekë']
    lines = [
        f"Konventa: Diferenca = Financa − Portal (pozitive = Financa ka më shumë). Libri SC mbulon deri më {fmt_data(rez['cutoff'])}.",
        f"Diferenca totale pas korrigjimeve është {rez['diferenca']:,.0f} lekë.",
        f"Prej saj {koha:,.0f} lekë janë çështje kohe: shitje në pritje fiskalizimi ose pas datës së librit SC. Do të zhduken kur të rigjenerohet Libri SC.",
        f"{reale:,.0f} lekë janë për t'u kontrolluar (diferenca reale, fatura që mungojnë, storno pa çift, ditë pa eksport).",
        f"{rrum:,.0f} lekë janë rrumbullakime (≤ {TOLERANCE} lekë për rresht).",
        ("Çdo lek është i shpjeguar (e pashpjeguar = 0)." if abs(rez['e_pashpjeguar']) <= TOLERANCE else
         f"Kujdes: mbeten {rez['e_pashpjeguar']:,.0f} lekë të pashpjeguara; kontrollo fletën 'Ura e diferencës'."),
    ]
    return lines


def _shkruaj_fleten(writer, df, emri, nota=None):
    if nota:
        df.to_excel(writer, sheet_name=emri, index=False, startrow=2)
        writer.sheets[emri].cell(row=1, column=1, value=nota)
    else:
        df.to_excel(writer, sheet_name=emri, index=False)


def shkruaj_raportin(rez, out):
    konventa = 'Konventa: Diferenca = Financa − Portal (pozitive = Financa ka më shumë).'
    perm = rez['permbledhje']
    perm_rows = [{'Përmbledhje': l} for l in perm] + [{'Përmbledhje': ''}, {'Përmbledhje': 'VEPRIMET KRYESORE (lista e plotë te fleta "Veprime për t\'u bërë")'}]
    for i, (_, r) in enumerate(rez['veprime'].head(8).iterrows(), 1):
        perm_rows.append({'Përmbledhje': f"{i}. {r['Veprimi']}" + (f" — {r['Detaje']}" if r['Detaje'] else '')})
    if rez['veprime'].empty:
        perm_rows.append({'Përmbledhje': 'Asnjë veprim i nevojshëm.'})

    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        pd.DataFrame(perm_rows).to_excel(writer, sheet_name='Përmbledhje', index=False)
        rez['ura'].to_excel(writer, sheet_name='Ura e diferencës', index=False)
        veprime = rez['veprime'] if len(rez['veprime']) else pd.DataFrame(columns=['Prioriteti', 'Veprimi', 'Detaje', 'Vlera'])
        veprime.to_excel(writer, sheet_name="Veprime për t'u bërë", index=False)
        if rez['paralajmerime']:
            pd.DataFrame({'Paralajmerime': rez['paralajmerime']}).to_excel(writer, sheet_name='Paralajmerime', index=False)
        rez['kontroll'].to_excel(writer, sheet_name='Kontroll', index=False)
        if not rez['detaje_operatori'].empty:
            _shkruaj_fleten(writer, rez['detaje_operatori'], 'Detaje_Operatori', konventa)
        operator_report = rez['operator_report']
        if operator_report.empty:
            operator_report = pd.DataFrame(columns=['Data', 'Diferenca', 'Status', 'Kush_ka_me_shume',
                                                    'Operatori(et) i identifikuar', 'Shuma e përputhur'])
        operator_report.to_excel(writer, sheet_name='Operatori_Problematik', index=False)
        _shkruaj_fleten(writer, rez['cmp'], 'Klient_Rastesishem_Ditor', konventa)
        _shkruaj_fleten(writer, rez['subjekt_diff'][[
            'NrFature_Financa', 'NIVF', 'Data', 'Bleresi', 'NIPT', 'Vlera_Financa', 'Vlera_Portal',
            'Diferenca', 'Kush_ka_me_shume', 'Problemi', 'Status', 'Shenim']], 'Subjekt_Identifikuar', konventa)
        rez['kategori_diff'].to_excel(writer, sheet_name='Diferenca_Kategori_TVSH', index=False)
    if hasattr(out, 'seek'):
        out.seek(0)
    format_workbook(out)


STATUS_FILLS = {
    ST_OK: 'C6EFCE',
    ST_PRITJE: 'BDD7EE', ST_PAS: 'BDD7EE',
    ST_MUNGON_PORTAL: 'FFC7CE', ST_MUNGON_FIN: 'FFC7CE', ST_STORNO: 'F8CBAD',
    ST_REALE: 'FFC7CE', ST_PA_EKSPORT: 'FFEB9C', ST_PA_SHPJEGIM: 'FFC7CE',
}


def format_workbook(path):
    from datetime import datetime
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)
    header_fill = PatternFill('solid', start_color='1F4E78', end_color='1F4E78')
    header_font = Font(bold=True, color='FFFFFF', name='Arial')
    base_font = Font(name='Arial')
    problemi_fills = {
        'Deklaruar ne': PatternFill('solid', start_color='FFFF99', end_color='FFFF99'),
        'dalin ne': PatternFill('solid', start_color='BDD7EE', end_color='BDD7EE'),
        'Pa serial': PatternFill('solid', start_color='F8CBAD', end_color='F8CBAD'),
    }
    red_fill = PatternFill('solid', start_color='FFC7CE', end_color='FFC7CE')
    green_fill = PatternFill('solid', start_color='C6EFCE', end_color='C6EFCE')
    status_fills = {k: PatternFill('solid', start_color=v, end_color=v) for k, v in STATUS_FILLS.items()}
    # kolonat që nuk janë shuma (numra rendorë, ditë, etj.)
    kol_pa_format = {'Numri', 'Prioriteti', 'Vonesa_Dite', 'NrFature_Financa', 'NrFature', 'Nr', 'NrRendor'}

    for ws in wb.worksheets:
        hdr = 3 if str(ws['A1'].value or '').startswith('Konventa:') else 1
        for cell in ws[hdr]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        if hdr == 3:
            ws['A1'].font = Font(name='Arial', italic=True, bold=True)
        headers = [c.value for c in ws[hdr]]
        status_col = headers.index('Status')+1 if 'Status' in headers else None
        oper_col = headers.index('Operatori(et) i identifikuar')+1 if 'Operatori(et) i identifikuar' in headers else None
        problemi_col = headers.index('Problemi')+1 if 'Problemi' in headers else None
        for row in ws.iter_rows(min_row=hdr+1):
            for cell in row:
                cell.font = base_font
                h = headers[cell.column-1] if cell.column-1 < len(headers) else None
                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD.MM.YYYY'
                elif isinstance(cell.value, float) and h not in kol_pa_format:
                    cell.number_format = '#,##0.00'
                elif isinstance(cell.value, int) and not isinstance(cell.value, bool) and h not in kol_pa_format:
                    cell.number_format = '#,##0'
            if problemi_col:
                vlera = str(row[problemi_col-1].value or '')
                for fjala, fill in problemi_fills.items():
                    if fjala in vlera:
                        for cell in row:
                            cell.fill = fill
                        break
            if status_col:
                fill = status_fills.get(row[status_col-1].value)
                if fill:
                    for cell in row:
                        cell.fill = fill
            if oper_col and row[oper_col-1].value:
                fill = red_fill if 'Asnjë' in str(row[oper_col-1].value) else green_fill
                for cell in row:
                    cell.fill = fill
        for i, col in enumerate(ws.columns, 1):
            max_len = max((len(str(c.value)) for c in col[hdr-1:] if c.value is not None), default=10)
            ws.column_dimensions[get_column_letter(i)].width = min(max_len+2, 45)
        if ws.title in ('Përmbledhje', 'Paralajmerime'):
            ws.column_dimensions['A'].width = 130
            for row in ws.iter_rows(min_row=2):
                row[0].alignment = Alignment(wrap_text=True, vertical='top')
        if ws.title == "Veprime për t'u bërë":
            ws.column_dimensions['B'].width = 60
            ws.column_dimensions['C'].width = 90
        ws.freeze_panes = f'A{hdr+1}'

    if hasattr(path, 'seek'):
        # path eshte nje file-like object (p.sh. io.BytesIO) i lexuar deri ne fund
        # nga load_workbook me lart; duhet zbrazur para se te rishkruhet, perndryshe
        # bajtet e reja shtohen mbi te vjetrat dhe skedari del i demtuar per Excel-in.
        path.seek(0)
        path.truncate()
    wb.save(path)


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('financa')
    ap.add_argument('portal')
    ap.add_argument('out')
    ap.add_argument('--portal-prev', help='Self Care i muajit paraardhes')
    ap.add_argument('--financa-next', help='F5 i muajit pasardhes')
    ap.add_argument('--eksport', nargs='*', default=[], help='Eksportet e fiskalizimit (nje ose me shume skedare)')
    ap.add_argument('--cutoff', help='Data e shkarkimit te Librit SC (YYYY-MM-DD); parazgjedhje: data me e madhe ne librin SC')
    args = ap.parse_args()

    rez = rakordo(args.financa, fleta_default(args.financa), args.portal, fleta_default(args.portal),
                  portal_prev_path=args.portal_prev, fin_next_path=args.financa_next,
                  trans_paths=args.eksport, cutoff=args.cutoff)
    shkruaj_raportin(rez, args.out)

    for p in rez['paralajmerime']:
        print("KUJDES:", p)
    for l in rez['permbledhje']:
        print(l)
    print(rez['ura'].to_string())
