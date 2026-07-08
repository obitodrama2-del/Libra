"""
Ndërfaqe web për rakordimin mujor: Libri i Shitjeve Financa vs Portali Self Care
Ngarko 2 excelet, zgjidh fletët, dhe merr raportin e krahasimit direkt në shfletues.
"""
import io
import streamlit as st
import pandas as pd

from rakordim import rakordo, format_workbook

st.set_page_config(page_title="Rakordim Financa vs Portal", layout="wide")
st.title("Rakordim: Libri i Shitjeve Financa vs Portali Self Care")

col1, col2 = st.columns(2)
with col1:
    financa_file = st.file_uploader("Libri i Shitjeve - Financa", type=["xlsx"], key="financa")
with col2:
    portal_file = st.file_uploader("Libri i Shitjeve - Portal Self Care", type=["xlsx"], key="portal")

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
    financa_file.seek(0)
    portal_file.seek(0)
    try:
        with st.spinner("Duke rakorduar..."):
            subjekt_diff, cmp, operator_report = rakordo(financa_file, financa_sheet, portal_file, portal_sheet)
            subjekt_diff = subjekt_diff.drop(columns=['_merge'])

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                operator_report.to_excel(writer, sheet_name='Operatori_Problematik', index=False)
                cmp.to_excel(writer, sheet_name='Klient_Rastesishem_Ditor', index=False)
                subjekt_diff.to_excel(writer, sheet_name='Subjekt_Identifikuar', index=False)
            buf.seek(0)
            format_workbook(buf)
            buf.seek(0)
    except Exception as e:
        st.error(f"Gabim gjatë përpunimit: {e}")
    else:
        st.success("Rakordimi u krye me sukses.")

        m1, m2 = st.columns(2)
        m1.metric("Subjekt identifikuar - diferenca gjetur", len(subjekt_diff))
        m2.metric("Ditë me diferencë (klient i rastësishëm)", int((cmp['Status'] == 'KONTROLLO').sum()))

        st.download_button(
            "Shkarko raportin (Excel)",
            data=buf,
            file_name="Rakordim_Rezultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        tab1, tab2, tab3 = st.tabs(["Operatori Problematik", "Klient Rastësishëm Ditor", "Subjekt Identifikuar"])
        with tab1:
            st.dataframe(operator_report, use_container_width=True)
        with tab2:
            st.dataframe(cmp, use_container_width=True)
        with tab3:
            st.dataframe(subjekt_diff, use_container_width=True)
