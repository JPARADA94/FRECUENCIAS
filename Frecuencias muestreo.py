import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Nota: asegúrate de tener instalado 'openpyxl' para leer .xlsx:
# pip install openpyxl

# ======================
# CACHE PARA OPTIMIZAR
# ======================
@st.cache
def load_data(uploaded_file):
    """
    Carga el archivo (CSV o XLSX) y retorna un DataFrame con las columnas necesarias.
    """
    filename = uploaded_file.name.lower()
    if filename.endswith('.csv'):
        df = pd.read_csv(uploaded_file, parse_dates=['Date Sampled'])
    else:
        # Para .xlsx, pandas usará openpyxl por defecto
        df = pd.read_excel(uploaded_file, parse_dates=['Date Sampled'])
    # Filtrar solo las columnas que necesitamos
    cols = ['Unit ID', 'Asset ID', 'Account Name', 'Sample Bottle ID', 'Date Sampled', 'Asset Class']
    return df[cols]

@st.cache
def analyze_df(df: pd.DataFrame, freq_unit: str):
    """
    Retorna dos DataFrames:
      1) Pivot de muestreo anual por equipo
      2) Fechas futuras (una por fila)
    """
    df = df.dropna(subset=['Date Sampled']).copy()
    df['Year'] = df['Date Sampled'].dt.year

    # Años de 2021 al actual
    current_year = datetime.today().year
    years = list(range(2021, current_year + 1))

    # 1) Conteo por año
    cnt = (
        df.groupby(['Unit ID','Asset ID','Asset Class','Account Name','Year'])['Sample Bottle ID']
          .nunique()
          .reset_index(name='Count')
    )
    pivot = (
        cnt.pivot_table(
            index=['Unit ID','Asset ID','Asset Class','Account Name'],
            columns='Year',
            values='Count',
            fill_value=0
        )
        .reindex(columns=years, fill_value=0)
        .reset_index()
    )

    # 2) Intervalos promedio
    df_sorted = df.sort_values(['Unit ID','Asset ID','Date Sampled'])
    df_sorted['Prev'] = df_sorted.groupby(['Unit ID','Asset ID'])['Date Sampled'].shift(1)
    df_sorted['Interval Days'] = (df_sorted['Date Sampled'] - df_sorted['Prev']).dt.days

    avg_int = (
        df_sorted.groupby(['Unit ID','Asset ID','Asset Class','Account Name'])['Interval Days']
                 .mean()
                 .reset_index(name='Avg Interval (Days)')
    )

    # 3) Formato de frecuencia
    divisor = 7 if freq_unit=='Semanas' else 30
    label   = 'semanas' if freq_unit=='Semanas' else 'meses'
    avg_int['Recommended Frequency'] = (
        (avg_int['Avg Interval (Days)']/divisor)
        .round(1)
        .astype(str) + f' {label}'
    )

    # 4) Fechas futuras
    future = []
    today = pd.Timestamp.today().normalize()
    limit = pd.Timestamp(year=today.year+1, month=12, day=31)
    for _, r in avg_int.iterrows():
        last = df_sorted[
            (df_sorted['Unit ID']==r['Unit ID']) &
            (df_sorted['Asset ID']==r['Asset ID'])
        ]['Date Sampled'].max().normalize()
        start = max(today, last)
        next_date = start
        while next_date <= limit:
            next_date += pd.Timedelta(days=r['Avg Interval (Days)'])
            # ajustar fin de semana
            if next_date.weekday()>=5:
                next_date += pd.Timedelta(days=(7-next_date.weekday()))
            future.append({
                'Unit ID': r['Unit ID'],
                'Asset ID': r['Asset ID'],
                'Asset Class': r['Asset Class'],
                'Account Name': r['Account Name'],
                'Future Sample Date': next_date.date()
            })
    future_df = pd.DataFrame(future)

    # 5) Unir pivot con frecuencia
    result = pivot.merge(
        avg_int[['Unit ID','Asset ID','Recommended Frequency']],
        on=['Unit ID','Asset ID'],
        how='left'
    )
    return result, future_df

def to_excel(df1: pd.DataFrame, df2: pd.DataFrame) -> bytes:
    """
    Genera un Excel con dos hojas: 'Muestreo Anual' y 'Fechas Futuras'.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df1.to_excel(writer, index=False, sheet_name='Muestreo Anual')
        df2.to_excel(writer, index=False, sheet_name='Fechas Futuras')
    return output.getvalue()

# ================
# UI de Streamlit
# ================
st.title("Análisis de Frecuencia de Muestreo por Operaciones")
st.markdown(
    "- Formato MobilServ\n"
    "- Columnas obligatorias: Unit ID, Asset ID, Account Name, Sample Bottle ID, Date Sampled, Asset Class\n"
    "- Puedes subir CSV o XLSX"
)

uploaded = st.file_uploader("Sube tu archivo", type=["xlsx","csv"])
if uploaded:
    df = load_data(uploaded)

    ops = sorted(df['Account Name'].dropna().unique())
    selected = st.multiselect(
        "Selecciona operaciones (Account Name)",
        options=ops,
        default=ops
    )

    freq = st.selectbox("Frecuencia en:", ['Semanas','Meses'])

    if selected:
        df_sel = df[df['Account Name'].isin(selected)]
        annual_df, future_df = analyze_df(df_sel, freq)

        st.subheader("Muestreo Anual")
        st.dataframe(annual_df)

        st.subheader("Fechas Futuras (una por fila)")
        st.dataframe(future_df)

        data = to_excel(annual_df, future_df)
        st.download_button(
            "Descargar Excel completo",
            data=data,
            file_name="sampling_analysis_full.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
