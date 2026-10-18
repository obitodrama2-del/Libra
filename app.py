"""
Ndërfaqe web për rakordimin mujor: Libri i Shitjeve Financa vs Portali Self Care
Ngarko 2 excelet, zgjidh fletët, dhe merr raportin e krahasimit direkt në shfletues.
"""
import io
import streamlit as st
import pandas as pd

from rakordim import (rakordo, shkruaj_raportin, parse_transaksionet,
                      identifiko_operatorin_transaksione)

st.set_page_config(page_title="Rakordim Financa vs Portal", layout="wide")
st.title("Rakordim: Libri i Shitjeve Financa vs Portali Self Care")

col1, col2 = st.columns(2)
with col1:
    financa_file = st.file_uploader("Libri i Shitjeve - Financa", type=["xlsx"], key="financa")
with col2:
    portal_file = st.file_uploader("Libri i Shitjeve - Portal Self Care", type=["xlsx"], key="portal")

with st.expander("Opsionale: muaji paraardhës dhe pasardhës (për faturat që kalojnë nga një muaj te tjetri)"):
    col5, col6 = st.columns(2)
    with col5:
        portal_prev_file = st.file_uploader(
            "Self Care i muajit paraardhës (faturat F5 që gjenden aty: \"Deklaruar ne ...\")",
            type=["xlsx"], key="portal_prev"
        )
    with col6:
        financa_next_file = st.file_uploader(
            "F5 i muajit pasardhës (faturat SC që mungojnë në F5 por dalin aty)",
            type=["xlsx"], key="financa_next"
        )

trans_files = st.file_uploader(
    "Transaksionet e detajuara (opsionale, për ditët problematike - identifikon automatikisht pikën e shitjes)",
    type=["xlsx"], key="transaksionet", accept_multiple_files=True
)

def get_sheet_names(uploaded_file):
    if uploaded_file is None:
        return []
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)
    return xls.sheet_names

financa_sheets = get_sheet_names(financa_file)
portal_sheets = get_sheet_names(portal_file)

col3, col4 = st.columns(2)
with col3:
    financa_sheet = st.selectbox(
        "Fleta - Financa", financa_sheets,
        index=financa_sheets.index('shitje') if 'shitje' in financa_sheets else 0
    ) if financa_sheets else None
with col4:
    portal_sheet = st.selectbox(
        "Fleta - Portal", portal_sheets,
        index=portal_sheets.index('Sales book') if 'Sales book' in portal_sheets else 0
    ) if portal_sheets else None

run = st.button("Krahaso", type="primary", disabled=not (financa_file and portal_file))

if run:
    try:
        with st.spinner("Duke rakorduar..."):
            rez = rakordo(financa_file, financa_sheet, portal_file, portal_sheet,
                          portal_prev_path=portal_prev_file, fin_next_path=financa_next_file)

            detaje_operatori = pd.DataFrame()
            if trans_files:
                detaje_rows = []
                for tf in trans_files:
                    tf.seek(0)
                    trans_df = parse_transaksionet(tf)
                    detaje_rows.append(identifiko_operatorin_transaksione(rez['fin_rand'], trans_df))
                detaje_operatori = pd.concat(detaje_rows, ignore_index=True)

            buf = io.BytesIO()
            shkruaj_raportin(rez, buf, detaje_operatori)
            buf.seek(0)
    except Exception as e:
        st.error(f"Gabim gjatë përpunimit: {e}")
    else:
        for p in rez['paralajmerime']:
            st.warning(p)
        if not rez['paralajmerime']:
            st.success("Rakordimi u krye me sukses.")

        subjekt_diff, cmp, kategori_diff = rez['subjekt_diff'], rez['cmp'], rez['kategori_diff']
        diferenca_finale = rez['kontroll']['Shuma'].iloc[-1]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Diferenca finale (pas korrigjimeve)", f"{diferenca_finale:,.0f}")
        m2.metric("Subjekt identifikuar - rreshta", len(subjekt_diff))
        m3.metric("Ditë me diferencë (klient i rastësishëm)", int((cmp['Status'] == 'KONTROLLO').sum()))
        m4.metric("Diferenca sipas kategorive TVSH", len(kategori_diff))

        st.download_button(
            "Shkarko raportin (Excel)",
            data=buf,
            file_name="Rakordim_Rezultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        tab_names = ["Kontroll", "Subjekt Identifikuar", "Operatori Problematik",
                     "Klient Rastësishëm Ditor", "Diferenca Kategori TVSH"]
        if not detaje_operatori.empty:
            tab_names.append("Detaje Operatori (nga Transaksionet)")
        tabs = st.tabs(tab_names)
        with tabs[0]:
            st.caption("Financa e korrigjuar = F5 − faturat e deklaruara muajin paraardhës − pa serial + faturat që dalin në F5 muajin pasardhës. Duhet të bjerë me Portalin, me diferencë vetëm rrumbullakime.")
            st.dataframe(rez['kontroll'], use_container_width=True)
        with tabs[1]:
            if len(subjekt_diff):
                st.write(subjekt_diff['Problemi'].value_counts())
            st.dataframe(subjekt_diff, use_container_width=True)
        with tabs[2]:
            st.dataframe(rez['operator_report'], use_container_width=True)
        with tabs[3]:
            st.dataframe(cmp, use_container_width=True)
        with tabs[4]:
            st.caption("Diferenca ndaras per çdo kategori TVSH-je (Perjashtuar, Tatueshme 20%/10%/6%, etj.) — kap edhe rastet kur totali ditor përputhet por një shitje ka kaluar gabimisht nga një kategori në tjetrën.")
            st.dataframe(kategori_diff, use_container_width=True)
        if not detaje_operatori.empty:
            with tabs[5]:
                st.dataframe(
                    detaje_operatori.style.apply(
                        lambda r: ['background-color: #ffc7ce' if r['Status'] == 'KONTROLLO' else 'background-color: #c6efce'] * len(r),
                        axis=1
                    ),
                    use_container_width=True
                )
