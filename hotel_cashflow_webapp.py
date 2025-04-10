
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import os

st.set_page_config(page_title="Hotel Cashflow", layout="wide")

st.title("Hotel Cashflow - Web App")

# === Caricamento file ===
uploaded_spese = st.file_uploader("Carica file Spese (.xlsx)", type=["xlsx"], key="spese")
uploaded_incassi = st.file_uploader("Carica file Incassi (.xlsx)", type=["xlsx"], key="incassi")

df_spese, df_incassi = None, None

if uploaded_spese:
    df_spese = pd.read_excel(uploaded_spese)
    st.success("File Spese caricato correttamente.")
    st.dataframe(df_spese.head())

if uploaded_incassi:
    df_incassi = pd.read_excel(uploaded_incassi)
    st.success("File Incassi caricato correttamente.")
    st.dataframe(df_incassi.head())

def esporta_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_spese.to_excel(writer, sheet_name="Dettaglio Spese", index=False)
        df_incassi.to_excel(writer, sheet_name="Dettaglio Incassi", index=False)

        # === Cashflow Mensile ===
        df_spese["Categoria"] = df_spese["Categoria"].astype(str).str.strip().str.title()
        df_spese["Mese"] = pd.to_datetime(df_spese["Data"], errors="coerce").dt.month_name()
        mesi_tradotti = {
            "January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile",
            "May": "Maggio", "June": "Giugno", "July": "Luglio", "August": "Agosto",
            "September": "Settembre", "October": "Ottobre", "November": "Novembre", "December": "Dicembre"
        }
        df_spese["Mese"] = df_spese["Mese"].map(mesi_tradotti)

        mesi_ordine = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
                       "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

        spese_mese = df_spese.groupby(["Mese", "Categoria"])["Importo"].sum().unstack(fill_value=0)
        spese_mese = spese_mese.reindex(mesi_ordine).fillna(0)
        spese_mese["Totale Spese"] = spese_mese.sum(axis=1)

        incassi_mese = df_incassi[["Mese", "Totale"]].groupby("Mese").sum()
        incassi_mese = incassi_mese.reindex(mesi_ordine).fillna(0)

        cashflow = spese_mese.copy()
        cashflow["Totale Incassi"] = incassi_mese["Totale"]
        cashflow["Risultato Netto"] = cashflow["Totale Incassi"] - cashflow["Totale Spese"]
        cashflow["Cumulato"] = cashflow["Risultato Netto"].cumsum()

        cashflow = cashflow.reset_index()
        cashflow.to_excel(writer, sheet_name="Cashflow Mensile", index=False)

    output.seek(0)
    return output

# === Bottone per generare ed esportare ===
if uploaded_spese and uploaded_incassi:
    if st.button("Genera ed Esporta Excel"):
        file_excel = esporta_excel()
        st.success("File Excel generato correttamente!")
        st.download_button(label="Scarica Excel", data=file_excel, file_name="cashflow_riepilogo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
