
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os

st.set_page_config(page_title="Hotel Cashflow", layout="wide")
st.title("Hotel Cashflow - Web App v3")

uploaded_spese = st.file_uploader("Carica file Spese (.xlsx)", type=["xlsx"], key="spese")
uploaded_incassi = st.file_uploader("Carica file Prenotazioni (.xlsx)", type=["xlsx"], key="incassi")

df_spese, df_incassi = None, None

if uploaded_spese:
    df_spese = pd.read_excel(uploaded_spese)
    st.success("File Spese caricato correttamente.")
    st.dataframe(df_spese.head())

if uploaded_incassi:
    raw_incassi = pd.read_excel(uploaded_incassi)
    st.success("File Prenotazioni caricato correttamente.")
    st.dataframe(raw_incassi.head())

def esporta_excel():
    output = BytesIO()

    if df_spese is None or raw_incassi is None:
        st.error("Carica entrambi i file per procedere.")
        return None

    # Prepara spese
    if "Categoria" not in df_spese.columns and "CatShort" in df_spese.columns:
        df_spese["Categoria"] = df_spese["CatShort"].map({"F": "Costi Fissi", "V": "Costi Variabili"})

    if "Importo" not in df_spese.columns and "Imponibile" in df_spese.columns and "IVA" in df_spese.columns:
        df_spese["Importo"] = pd.to_numeric(df_spese["Imponibile"], errors="coerce").fillna(0) + pd.to_numeric(df_spese["IVA"], errors="coerce").fillna(0)

    df_spese["Categoria"] = df_spese["Categoria"].astype(str).str.strip().str.title()
    df_spese["Mese"] = pd.to_datetime(df_spese["Data"], errors="coerce").dt.month_name()

    mesi_tradotti = {
        "January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile",
        "May": "Maggio", "June": "Giugno", "July": "Luglio", "August": "Agosto",
        "September": "Settembre", "October": "Ottobre", "November": "Novembre", "December": "Dicembre"
    }
    df_spese["Mese"] = df_spese["Mese"].map(mesi_tradotti)

    # === Prepara incassi da prenotazioni ===
    df_incassi = raw_incassi.copy()
    df_incassi["Mese"] = pd.to_datetime(df_incassi["Arrivo"], errors="coerce").dt.month_name()
    df_incassi["Mese"] = df_incassi["Mese"].map(mesi_tradotti)
    df_incassi["Prezzo(€)"] = pd.to_numeric(df_incassi["Prezzo(€)"], errors="coerce").fillna(0)

    def mappa_tipologia(alloggio):
        if pd.isna(alloggio):
            return "Altro"
        alloggio = str(alloggio).lower()
        if "base" in alloggio:
            return "STD-AD"
        elif "standard" in alloggio:
            return "STD-CON"
        elif "superior" in alloggio:
            return "SUP-CON"
        elif "lungo termine" in alloggio:
            return "Lungo Termine"
        else:
            return "Altro"

    df_incassi["Tipologia"] = df_incassi["Alloggio"].map(mappa_tipologia)

    pivot_incassi = df_incassi.pivot_table(index="Mese", columns="Tipologia", values="Prezzo(€)", aggfunc="sum", fill_value=0)
    for col in ["STD-AD", "STD-CON", "SUP-CON", "Lungo Termine"]:
        if col not in pivot_incassi.columns:
            pivot_incassi[col] = 0
    pivot_incassi["Totale"] = pivot_incassi[["STD-AD", "STD-CON", "SUP-CON", "Lungo Termine"]].sum(axis=1)
    pivot_incassi = pivot_incassi.reset_index()

    mesi_ordine = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
                   "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
    pivot_incassi["Mese"] = pd.Categorical(pivot_incassi["Mese"], categories=mesi_ordine, ordered=True)
    pivot_incassi = pivot_incassi.sort_values("Mese")

    # === Aggregazione spese ===
    spese_mese = df_spese.groupby(["Mese", "Categoria"])["Importo"].sum().unstack(fill_value=0)
    spese_mese = spese_mese.reindex(mesi_ordine).fillna(0)
    spese_mese["Totale Spese"] = spese_mese.sum(axis=1)

    incassi_mese = pivot_incassi.set_index("Mese")[["Totale"]]
    cashflow = spese_mese.copy()
    cashflow["Totale Incassi"] = incassi_mese["Totale"]
    cashflow["Risultato Netto"] = cashflow["Totale Incassi"] - cashflow["Totale Spese"]
    cashflow["Cumulato"] = cashflow["Risultato Netto"].cumsum()
    cashflow = cashflow.reset_index()

    # === Esportazione
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_spese.to_excel(writer, sheet_name="Dettaglio Spese", index=False)
        pivot_incassi.to_excel(writer, sheet_name="Dettaglio Incassi", index=False)
        cashflow.to_excel(writer, sheet_name="Cashflow Mensile", index=False)

    output.seek(0)
    return output

# === Bottone per generare ed esportare ===
if uploaded_spese is not None and uploaded_incassi is not None:
    if st.button("Genera ed Esporta Excel"):
        file_excel = esporta_excel()
        if file_excel:
            st.success("File Excel generato correttamente!")
            st.download_button(label="Scarica Excel", data=file_excel, file_name="cashflow_riepilogo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

