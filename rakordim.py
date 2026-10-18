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
import sys

TOLERANCE = 5  # lek, diferenca nën këtë vlerë konsiderohen rrumbullakosje

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
    data['Data'] = pd.to_datetime(data['Data'], errors='coerce', dayfirst=True)
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


def parse_transaksionet(path):
    """Lexon skedarin e transaksioneve te detajuara te eksportuar nga Portali Self Care (fleta 'data')."""
    df = pd.read_excel(path, sheet_name='data')
    df['Data e fiskalizimit'] = pd.to_datetime(df['Data e fiskalizimit'], errors='coerce', dayfirst=True)
    return df


def identifiko_operatorin_transaksione(fin_rand, trans_df, tol=5):
    """
    Perputh cdo pike shitjeje (Bleresi ne Financa) me kodin e operatorit ne Portal,
    duke krahasuar totalet ditore te klienteve te rastesishem (pa NIVF).
    Kthen nje DataFrame me diferencat per cdo pike shitjeje te dites.
    """
    data_val = trans_df['Data e fiskalizimit'].mode().iloc[0]
    random_tx = trans_df[trans_df['Numri i identifikimit të blerësit'].isna()]
    portal_by_code = random_tx.groupby('Kodi i operatorit')['Vlera totale me TVSH'].sum().to_dict()

    fin_day = fin_rand[fin_rand['Data'] == data_val]
    fin_by_bleresi = fin_day.groupby('Bleresi')['Vlera'].sum().to_dict()

    pairs = [
        (abs(fv - pv), bleresi, code)
        for bleresi, fv in fin_by_bleresi.items()
        for code, pv in portal_by_code.items()
    ]
    pairs.sort(key=lambda x: x[0])

    matched_bleresi, matched_code = {}, {}
    for _, bleresi, code in pairs:
        if bleresi in matched_bleresi or code in matched_code:
            continue
        matched_bleresi[bleresi] = code
        matched_code[code] = bleresi

    rows = []
    for bleresi, fin_val in fin_by_bleresi.items():
        code = matched_bleresi.get(bleresi)
        portal_val = portal_by_code.get(code) if code else None
        diff = fin_val - (portal_val if portal_val is not None else 0)
        rows.append({
            'Data': data_val, 'Bleresi': bleresi, 'Financa_Vlera': fin_val,
            'Kodi_Operatorit': code, 'Portal_Vlera': portal_val, 'Diferenca': diff
        })
    for code, portal_val in portal_by_code.items():
        if code not in matched_code:
            rows.append({
                'Data': data_val, 'Bleresi': None, 'Financa_Vlera': None,
                'Kodi_Operatorit': code, 'Portal_Vlera': portal_val, 'Diferenca': -portal_val
            })

    result = pd.DataFrame(rows)
    result['Status'] = np.where(result['Diferenca'].abs() <= tol, 'OK', 'KONTROLLO')
    return result.sort_values('Diferenca', key=abs, ascending=False).reset_index(drop=True)


def find_operator_combo(target_diff, operator_totals, tol=5, max_combo=3):
    """Gjen kombinimin e operatorëve (deri në max_combo) që shpjegon diferencën."""
    items = list(operator_totals.items())
    for r in range(1, max_combo+1):
        for combo in combinations(items, r):
            s = sum(v for _, v in combo)
            if abs(s - target_diff) <= tol:
                return [name for name, _ in combo], s
    return None, None


def rakordo(financa_path, financa_sheet, portal_path, portal_sheet,
            portal_prev_path=None, fin_next_path=None):
    """
    portal_prev_path: (opsionale) libri Self Care i muajit paraardhës - faturat F5 që
        gjenden aty shënohen "Deklaruar ne <muaji>" në vend të "Mungon në Portal".
    fin_next_path: (opsionale) libri F5 i muajit pasardhës - faturat SC që mungojnë në
        F5 por gjenden aty shënohen "<muaji> - del ne <muaji pasardhes> ne F5".
    """
    fin = parse_libri(financa_path, financa_sheet)
    port = parse_libri(portal_path, portal_sheet)
    port_prev = parse_libri(portal_prev_path, fleta_default(portal_prev_path)) if portal_prev_path is not None else None
    fin_next = parse_libri(fin_next_path, fleta_default(fin_next_path)) if fin_next_path is not None else None

    paralajmerime = kontrollo_periudhat(fin, port, port_prev, fin_next)
    m_port = muaji_kryesor(port)
    emri_prev = MUAJT[(m_port - 1).month - 1] if m_port is not None else 'muajin paraardhes'
    emri_tani = MUAJT[m_port.month - 1] if m_port is not None else 'Ky muaj'
    emri_next = MUAJT[(m_port + 1).month - 1] if m_port is not None else 'muajin pasardhes'

    # --- 1. Subjekt i identifikuar ---
    fin_id = fin[fin['NIPT'].notna()]
    port_id = port[port['NIPT'].notna()]
    pa_serial = fin_id[fin_id['NumriSerial'].isna()].copy()
    merged = pd.merge(
        fin_id[fin_id['NumriSerial'].notna()][['NumriSerial','Data','Bleresi','NIPT','Vlera']],
        port_id[port_id['NumriSerial'].notna()][['NumriSerial','Data','Bleresi','NIPT','Vlera']],
        on='NumriSerial', how='outer', suffixes=('_Financa','_Portal'), indicator=True
    )
    merged['Diferenca'] = merged['Vlera_Financa'].fillna(0) - merged['Vlera_Portal'].fillna(0)
    # Rreshtat "Mungon ne Financa" marrin Data/Bleresi/NIPT nga Portali
    for col in ('Data', 'Bleresi', 'NIPT'):
        merged[col] = merged[f'{col}_Financa'].combine_first(merged[f'{col}_Portal'])
    merged = merged[['NumriSerial','Data','Bleresi','NIPT','Vlera_Financa','Vlera_Portal','Diferenca','_merge']]
    subjekt_diff = merged[(merged['Diferenca'].abs() > TOLERANCE) | (merged['_merge'] != 'both')].copy()
    subjekt_diff['Problemi'] = np.where(
        subjekt_diff['_merge'] == 'left_only', 'Mungon në Portal',
        np.where(subjekt_diff['_merge'] == 'right_only', 'Mungon në Financa', 'Vlerë e ndryshme')
    )

    prev_serials = set(port_prev['NumriSerial'].dropna()) if port_prev is not None else set()
    next_serials = set(fin_next['NumriSerial'].dropna()) if fin_next is not None else set()
    deklaruar_prev = (subjekt_diff['_merge'] == 'left_only') & subjekt_diff['NumriSerial'].isin(prev_serials)
    del_next = (subjekt_diff['_merge'] == 'right_only') & subjekt_diff['NumriSerial'].isin(next_serials)
    subjekt_diff.loc[deklaruar_prev, 'Problemi'] = f"Deklaruar ne {emri_prev}"
    subjekt_diff.loc[del_next, 'Problemi'] = f"{emri_tani} - dalin ne {emri_next} ne F5"
    deklaruar_serials = set(subjekt_diff.loc[deklaruar_prev, 'NumriSerial'])
    del_next_serials = set(subjekt_diff.loc[del_next, 'NumriSerial'])

    if len(pa_serial):
        pa_serial_rows = pa_serial[['NumriSerial','Data','Bleresi','NIPT','Vlera']].rename(columns={'Vlera': 'Vlera_Financa'})
        pa_serial_rows['NrFature'] = pa_serial['NrFature']
        pa_serial_rows['Vlera_Portal'] = np.nan
        pa_serial_rows['Diferenca'] = pa_serial_rows['Vlera_Financa']
        pa_serial_rows['Problemi'] = 'Pa serial - nuk del ne SC'
        pa_serial_rows['_merge'] = 'left_only'
        subjekt_diff = pd.concat([subjekt_diff, pa_serial_rows], ignore_index=True)

    # Rreshtat e SC qe mungojne ne F5 por dalin ne F5 te muajit pasardhes, me te dhenat e plota nga ai liber
    if fin_next is not None:
        next_rows = fin_next[fin_next['NumriSerial'].isin(del_next_serials)]
        fill = next_rows.drop_duplicates('NumriSerial').set_index('NumriSerial')
        for col in ('Data', 'Bleresi', 'NIPT'):
            mask = subjekt_diff['NumriSerial'].isin(fill.index) & fill[col].reindex(subjekt_diff['NumriSerial']).notna().values
            subjekt_diff.loc[mask, col] = subjekt_diff.loc[mask, 'NumriSerial'].map(fill[col])
    else:
        next_rows = fin.iloc[0:0]

    deklaruar_rows = fin[fin['NumriSerial'].isin(deklaruar_serials)]

    # Libri F5 i korrigjuar per krahasim: - muaji paraardhes - pa serial + rreshtat qe dalin muajin pasardhes
    fin_adj = pd.concat([
        fin[~fin.index.isin(deklaruar_rows.index) & ~fin.index.isin(pa_serial.index)],
        next_rows,
    ], ignore_index=True)

    kontroll = pd.DataFrame([
        {'Pershkrimi': 'Financa (F5) - totali i librit', 'Numri': len(fin), 'Shuma': fin['Vlera'].sum()},
        {'Pershkrimi': f"(-) Deklaruar ne {emri_prev} (SC i muajit paraardhes)", 'Numri': len(deklaruar_rows), 'Shuma': -deklaruar_rows['Vlera'].sum()},
        {'Pershkrimi': '(-) Pa serial - nuk dalin ne SC', 'Numri': len(pa_serial), 'Shuma': -pa_serial['Vlera'].sum()},
        {'Pershkrimi': f"(+) Dalin ne {emri_next} ne F5", 'Numri': len(next_rows), 'Shuma': next_rows['Vlera'].sum()},
        {'Pershkrimi': 'Financa e korrigjuar', 'Numri': len(fin_adj), 'Shuma': fin_adj['Vlera'].sum()},
        {'Pershkrimi': 'Portali (Self Care) - totali', 'Numri': len(port), 'Shuma': port['Vlera'].sum()},
        {'Pershkrimi': 'Diferenca (Financa e korrigjuar - Portali)', 'Numri': None, 'Shuma': fin_adj['Vlera'].sum() - port['Vlera'].sum()},
    ])

    # --- 2. Klient i rastësishëm - totale ditore ---
    fin_rand = fin[fin['NIPT'].isna()]
    port_rand = port[port['NIPT'].isna()]
    fin_by_date = fin_rand.groupby('Data')['Vlera'].sum().rename('Financa_Total')
    port_by_date = port_rand.groupby('Data')['Vlera'].sum().rename('Portal_Total')
    cmp = pd.concat([fin_by_date, port_by_date], axis=1).fillna(0)
    cmp['Diferenca'] = cmp['Financa_Total'] - cmp['Portal_Total']
    cmp['Status'] = np.where(cmp['Diferenca'].abs() <= TOLERANCE, 'OK', 'KONTROLLO')
    cmp = cmp.reset_index()

    # --- 3. Identifikimi i operatorit për datat problematike ---
    problem_rows = []
    for _, row in cmp[cmp['Status']=='KONTROLLO'].iterrows():
        data_val = row['Data']
        diff = row['Diferenca']
        day_ops = fin_rand[fin_rand['Data']==data_val].groupby('Bleresi')['Vlera'].sum().to_dict()
        combo, matched_sum = find_operator_combo(diff, day_ops, tol=TOLERANCE)
        problem_rows.append({
            'Data': data_val,
            'Diferenca': diff,
            'Operatori(et) i identifikuar': ', '.join(combo) if combo else 'I PAIDENTIFIKUAR (kontroll manual)',
            'Shuma e përputhur': matched_sum if matched_sum is not None else None
        })
    operator_report = pd.DataFrame(problem_rows)

    # --- 4. Diferenca sipas kategorive te TVSH-se (edhe kur totali ditor perputhet) ---
    kategori_diff = krahaso_kategorite_tvsh(fin_adj, port)

    return {
        'subjekt_diff': subjekt_diff.drop(columns=['_merge']),
        'cmp': cmp,
        'operator_report': operator_report,
        'fin_rand': fin_rand,
        'kategori_diff': kategori_diff,
        'kontroll': kontroll,
        'paralajmerime': paralajmerime,
    }


def shkruaj_raportin(rez, out, detaje_operatori=None):
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        if rez['paralajmerime']:
            pd.DataFrame({'Paralajmerime': rez['paralajmerime']}).to_excel(writer, sheet_name='Paralajmerime', index=False)
        rez['kontroll'].to_excel(writer, sheet_name='Kontroll', index=False)
        operator_report = rez['operator_report']
        if operator_report.empty:
            operator_report = pd.DataFrame(columns=['Data', 'Diferenca', 'Operatori(et) i identifikuar', 'Shuma e përputhur'])
        operator_report.to_excel(writer, sheet_name='Operatori_Problematik', index=False)
        rez['cmp'].to_excel(writer, sheet_name='Klient_Rastesishem_Ditor', index=False)
        rez['subjekt_diff'].to_excel(writer, sheet_name='Subjekt_Identifikuar', index=False)
        rez['kategori_diff'].to_excel(writer, sheet_name='Diferenca_Kategori_TVSH', index=False)
        if detaje_operatori is not None and not detaje_operatori.empty:
            detaje_operatori.to_excel(writer, sheet_name='Detaje_Operatori', index=False)
    if hasattr(out, 'seek'):
        out.seek(0)
    format_workbook(out)


def format_workbook(path):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)
    header_fill = PatternFill('solid', start_color='1F4E78', end_color='1F4E78')
    header_font = Font(bold=True, color='FFFFFF', name='Arial')
    red_fill = PatternFill('solid', start_color='FFC7CE', end_color='FFC7CE')
    green_fill = PatternFill('solid', start_color='C6EFCE', end_color='C6EFCE')
    base_font = Font(name='Arial')
    problemi_fills = {
        'Deklaruar ne': PatternFill('solid', start_color='FFFF99', end_color='FFFF99'),
        'dalin ne': PatternFill('solid', start_color='BDD7EE', end_color='BDD7EE'),
        'Pa serial': PatternFill('solid', start_color='F8CBAD', end_color='F8CBAD'),
    }

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        headers = [c.value for c in ws[1]]
        status_col = headers.index('Status')+1 if 'Status' in headers else None
        oper_col = headers.index('Operatori(et) i identifikuar')+1 if 'Operatori(et) i identifikuar' in headers else None
        problemi_col = headers.index('Problemi')+1 if 'Problemi' in headers else None
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = base_font
            if problemi_col:
                vlera = str(row[problemi_col-1].value or '')
                for fjala, fill in problemi_fills.items():
                    if fjala in vlera:
                        for cell in row:
                            cell.fill = fill
                        break
            if status_col and row[status_col-1].value == 'KONTROLLO':
                for cell in row:
                    cell.fill = red_fill
            if status_col and row[status_col-1].value == 'OK':
                for cell in row:
                    cell.fill = green_fill
            if oper_col and row[oper_col-1].value and 'PAIDENTIFIKUAR' in str(row[oper_col-1].value):
                for cell in row:
                    cell.fill = red_fill
            elif oper_col and row[oper_col-1].value:
                for cell in row:
                    cell.fill = green_fill
        for i, col in enumerate(ws.columns, 1):
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            ws.column_dimensions[get_column_letter(i)].width = min(max_len+2, 45)
        ws.freeze_panes = 'A2'

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
    args = ap.parse_args()

    rez = rakordo(args.financa, fleta_default(args.financa), args.portal, fleta_default(args.portal),
                  portal_prev_path=args.portal_prev, fin_next_path=args.financa_next)
    shkruaj_raportin(rez, args.out)

    for p in rez['paralajmerime']:
        print("KUJDES:", p)
    print(rez['kontroll'].to_string())
    print("Subjekt identifikuar - diferenca gjetur:", len(rez['subjekt_diff']))
    print(rez['subjekt_diff']['Problemi'].value_counts().to_string())
    print("Ditë me diferencë tek klienti i rastësishëm:", (rez['cmp']['Status']=='KONTROLLO').sum())
    print("Diferenca sipas kategorive TVSH:", len(rez['kategori_diff']))
