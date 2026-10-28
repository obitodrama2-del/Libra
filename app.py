"""
Ndërfaqe web për rakordimin mujor: Libri i Shitjeve/Blerjeve Financa vs Portali Self Care
Ngarko 2 excelet, zgjidh fletët, dhe merr raportin e krahasimit direkt në shfletues.
"""
import io
import streamlit as st
import pandas as pd

from rakordim import rakordo, shkruaj_raportin
from blerje import rakordo_blerje, shkruaj_raportin_blerje

st.set_page_config(page_title="Rakordim Financa vs Portal", layout="wide")
st.title("Rakordim: Financa vs Portali Self Care")

mode = st.radio("Libri", ["Shitje", "Blerje"], horizontal=True)


def get_sheet_names(uploaded_file):
    if uploaded_file is None:
        return []
    uploaded_file.seek(0)
    return pd.ExcelFile(uploaded_file).sheet_names


def zgjidh_fleten(etiketa, fletet, te_preferuara):
    if not fletet:
        return None
    index = next((fletet.index(f) for f in te_preferuara if f in fletet), 0)
    return st.selectbox(etiketa, fletet, index=index)


def shfaq_statuse(df):
    """Ngjyros rreshtat sipas kolonës 'Status'/'Statusi' në pamjen në ekran."""
    kol = 'Status' if 'Status' in df.columns else ('Statusi' if 'Statusi' in df.columns else None)
    if kol is None:
        return df
    ngjyra = {'OK': '#c6efce', 'NË PRITJE FISKALIZIMI': '#bdd7ee', 'PAS DATËS SË LIBRIT SC': '#bdd7ee'}
    return df.style.apply(lambda r: [f"background-color: {ngjyra.get(r[kol], '#ffc7ce' if r[kol] else 'white')}"] * len(r), axis=1)


def shfaq_paralajmerimet(rez):
    for p in rez['paralajmerime']:
        st.warning(p)
    if not rez['paralajmerime']:
        st.success("Rakordimi u krye me sukses.")
    for l in rez['permbledhje']:
        st.write(l)


col1, col2 = st.columns(2)
with col1:
    financa_file = st.file_uploader(f"Libri i {mode.lower()}ve - Financa", type=["xlsx"], key=f"financa_{mode}")
with col2:
    portal_file = st.file_uploader(f"Libri i {mode.lower()}ve - Portal Self Care", type=["xlsx"], key=f"portal_{mode}")

financa_sheets, portal_sheets = get_sheet_names(financa_file), get_sheet_names(portal_file)
col3, col4 = st.columns(2)
with col3:
    financa_sheet = zgjidh_fleten("Fleta - Financa", financa_sheets, ('shitje', 'blerje') if mode == 'Shitje' else ('blerje', 'Purchase book'))
with col4:
    portal_sheet = zgjidh_fleten("Fleta - Portal", portal_sheets, ('Sales book',) if mode == 'Shitje' else ('Purchase book', 'Purchases book'))

cutoff = None
with st.expander("Data e shkarkimit të Librit SC (opsionale)"):
    if st.checkbox("Vendos datën manualisht", key=f"cutoff_on_{mode}"):
        cutoff = pd.Timestamp(st.date_input("Libri SC mbulon deri më", key=f"cutoff_{mode}"))
    st.caption("Nëse nuk jepet, përdoret data më e madhe e lëshimit në librin SC. Faturat e fiskalizuara pas kësaj date "
               "shënohen NË PRITJE FISKALIZIMI, dhe shitjet pas saj PAS DATËS SË LIBRIT SC.")

if mode == "Shitje":
    with st.expander("Opsionale: muaji paraardhës dhe pasardhës (për faturat që kalojnë nga një muaj te tjetri)"):
        col5, col6 = st.columns(2)
        with col5:
            portal_prev_file = st.file_uploader(
                "Self Care i muajit paraardhës (faturat F5 që gjenden aty: \"Deklaruar ne ...\")", type=["xlsx"], key="portal_prev")
        with col6:
            financa_next_file = st.file_uploader(
                "F5 i muajit pasardhës (faturat SC që mungojnë në F5 por dalin aty)", type=["xlsx"], key="financa_next")

    trans_files = st.file_uploader(
        "Eksportet e fiskalizimit (opsionale, një ose më shumë ditë - identifikon pikën e shitjes dhe faturat në pritje fiskalizimi)",
        type=["xlsx"], key="transaksionet", accept_multiple_files=True
    )

run = st.button("Krahaso", type="primary", disabled=not (financa_file and portal_file))

if run and mode == "Shitje":
    try:
        with st.spinner("Duke rakorduar..."):
            for tf in trans_files or []:
                tf.seek(0)
            rez = rakordo(financa_file, financa_sheet, portal_file, portal_sheet,
                          portal_prev_path=portal_prev_file, fin_next_path=financa_next_file,
                          trans_paths=trans_files or None, cutoff=cutoff)
            buf = io.BytesIO()
            shkruaj_raportin(rez, buf)
            buf.seek(0)
    except Exception as e:
        st.error(f"Gabim gjatë përpunimit: {e}")
    else:
        shfaq_paralajmerimet(rez)
        subjekt_diff, cmp = rez['subjekt_diff'], rez['cmp']

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Diferenca finale (Financa − Portal)", f"{rez['diferenca']:,.0f}")
        m2.metric("E pashpjeguar", f"{rez['e_pashpjeguar']:,.0f}")
        m3.metric("Subjekt identifikuar - rreshta", len(subjekt_diff))
        m4.metric("Ditë me diferencë reale/pa eksport", int(cmp['Status'].isin(['DIFERENCË REALE', 'PA EKSPORT']).sum()))

        st.download_button("Shkarko raportin (Excel)", data=buf, file_name="Rakordim_Rezultat.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        tab_names = ["Ura e diferencës", "Veprime për t'u bërë", "Kontroll", "Detaje Operatori", "Operatori Problematik",
                     "Klient Rastësishëm Ditor", "Subjekt Identifikuar", "Diferenca Kategori TVSH"]
        tabs = st.tabs(tab_names)
        with tabs[0]:
            st.caption("Diferenca = Financa − Portal. Kategoritë mblidhen saktë në totalin e Kontrollit.")
            st.dataframe(shfaq_statuse(rez['ura']), use_container_width=True)
        with tabs[1]:
            st.dataframe(rez['veprime'], use_container_width=True)
        with tabs[2]:
            st.caption("Financa e korrigjuar = F5 − faturat e deklaruara muajin paraardhës − pa serial + faturat që dalin në F5 muajin pasardhës.")
            st.dataframe(rez['kontroll'], use_container_width=True)
        with tabs[3]:
            if rez['detaje_operatori'].empty:
                st.info("Nuk ka ditë me diferencë ose eksport për operatorët.")
            else:
                st.caption("Diferenca = Financa − Portal (pozitive = Financa ka më shumë).")
                st.dataframe(shfaq_statuse(rez['detaje_operatori']), use_container_width=True)
        with tabs[4]:
            st.dataframe(rez['operator_report'], use_container_width=True)
        with tabs[5]:
            st.dataframe(shfaq_statuse(cmp), use_container_width=True)
        with tabs[6]:
            if len(subjekt_diff):
                st.write(subjekt_diff['Status'].value_counts())
            st.dataframe(shfaq_statuse(subjekt_diff), use_container_width=True)
        with tabs[7]:
            st.caption("Diferenca ndaras për çdo kategori TVSH-je; kap edhe rastet kur totali ditor përputhet por një shitje ka kaluar gabimisht nga një kategori në tjetrën.")
            st.dataframe(rez['kategori_diff'], use_container_width=True)

if run and mode == "Blerje":
    try:
        with st.spinner("Duke rakorduar blerjet..."):
            rez = rakordo_blerje(financa_file, financa_sheet, portal_file, portal_sheet, cutoff=cutoff)
            buf = io.BytesIO()
            shkruaj_raportin_blerje(rez, buf)
            buf.seek(0)
    except Exception as e:
        st.error(f"Gabim gjatë përpunimit: {e}")
    else:
        shfaq_paralajmerimet(rez)
        m1, m2, m3 = st.columns(3)
        m1.metric("Diferenca (Financa − Portal)", f"{rez['diferenca']:,.0f}")
        m2.metric("E pashpjeguar", f"{rez['e_pashpjeguar']:,.0f}")
        m3.metric("Fatura Financa / Portal", f"{rez['n_fin']} / {rez['n_port']}")
        st.download_button("Shkarko raportin (Excel)", data=buf, file_name="Rakordim_Blerje.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tabs = st.tabs(["Ura e blerjeve", "Veprime për t'u bërë", "Faturat jo OK", "Kontrolle", "Kategori TVSH", "Kontroll"])
        with tabs[0]:
            st.dataframe(shfaq_statuse(rez['ura']), use_container_width=True)
        with tabs[1]:
            st.dataframe(rez['veprime'], use_container_width=True)
        with tabs[2]:
            st.dataframe(shfaq_statuse(rez['faturat']), use_container_width=True)
        with tabs[3]:
            st.dataframe(rez['kontrolle'], use_container_width=True)
        with tabs[4]:
            st.dataframe(shfaq_statuse(rez['kategori_tvsh']), use_container_width=True)
        with tabs[5]:
            st.dataframe(rez['kontroll'], use_container_width=True)
