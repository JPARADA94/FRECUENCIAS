import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ======================
# CACHES PARA OPTIMIZAR
# ======================
@st.cache_data
def load_data(uploaded_file):
    # Carga solo las columnas necesarias para reducir memoria
    usecols = [
        'Unit ID',
        'Asset ID',
        'Account Name',
        'Sample Bottle ID',
        'Date Sampled',
        'Asset Class'
    ]
    df = pd.read_excel(
        uploaded_file,
        usecols=usecols,
        parse_dates=['Date Sampled']
    )
    return df

@st.cache_data
def analyze_df(df: pd.DataFrame, freq_unit: str):
    df = df.dropna(subset=['Date Sampled']).copy()
    df['Year'] = df['Date Sampled'].dt.year
    current_year = datetime.today().year
    years = list(range(2021, current_year + 1))

    # Conteo anual y pivote
    count = (
        df
        .groupby(['Unit ID', 'Asset ID', 'Asset Class', 'Account Name', 'Year'])['Sample Bottle ID']
        .nunique()
        .reset_index(name='Count')
    )
    pivot = (
        count
        .pivot_table(
            index=['Unit ID','Asset ID','Asset Class','Account Name'],
            columns='Year',
            values='Count',
            fill_value=0
        )
        .reindex(columns=years, fill_value=0)
        .reset_index()
    )

    # Calcular intervalos
    df_sorted = df.sort_values(['Unit ID','Asset ID','Date Sampled'])
    df_sorted['Prev'] = df_sorted.groupby(['Unit ID','Asset ID'])['Date Sampled'].shift(1)
    df_sorted['Interval Days'] = (df_sorted['Date Sampled'] - df_sorted['Prev']).dt.days
    avg_interval = (
        df_sorted
        .groupby(['Unit ID','Asset ID','Asset Class','Account Name'])['Interval Days']
        .mean()
        .reset_index(name='Avg Interval (Days)')
    )

    # Frecuencia recomendada
    divisor = 7 if freq_unit == 'Semanas' else 30
    label = 'semanas' if freq_unit == 'Semanas' else 'meses'
    avg_interval['Recommended Frequency'] = (
        (avg_interval['Avg Interval (Days)'] / divisor)
        .round(1)
        .astype(str) + f' {label}'
    )

    # Generar fechas futuras
    future = []
    today = pd.Timestamp.today().normalize()
    limit = pd.Timestamp(year=today.year+1, month=12, day=31)
    for _, row in avg_interval.iterrows():
        last = df_sorted[
            (df_sorted['Unit ID']==row['Unit ID']) &
            (df_sorted['Asset ID']==row['Asset ID'])
        ]['Date Sampled'].max().normalize()
        start = max(today, last)
        next_date = start
        while next_date <= limit:
            next_date += pd.Timedelta(days=row['Avg Interval (Days)'])
            if next_date.weekday() >= 5:
                next_date += pd.Timedelta(days=(7 - next_date.weekday()))
            future.append({
                'Unit ID': row['Unit ID'],
                'Asset ID': row['Asset ID'],
                'Asset Class': row['Asset Class'],
                'Account Name': row['Account Name'],
                'Future Sample Date': next_date.date()
            })

    future_df = pd.DataFrame(future)
    result = pivot.merge(
        avg_interval[['Unit ID','Asset ID','Recommended Frequency']],
        on=['Unit ID','Asset ID'], how='left'
    )
    return result, future_df

def to_excel(df1, df2):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df1.to_excel(writer, index=False, sheet_name='Muestreo Anual')
        df2.to_excel(writer, index=False, sheet_name='Fechas Futuras')
    return output.getvalue()

# ================
# INTERFAZ STREAMLIT
# ================
st.title("Análisis de Frecuencia de Muestreo")
st.markdown(
    "- Formato MobilServ\n"
    "- Columnas requeridas: Unit ID, Asset ID, Account Name, Sample Bottle ID, Date Sampled, Asset Class"
)

uploaded = st.file_uploader("Sube tu Excel MobilServ", type=["xlsx"])
if uploaded:
    df = load_data(uploaded)

    ops = df['Account Name'].dropna().unique().tolist()
    selected = st.multiselect("Operaciones (Account Name)", ops, default=[])
    run = st.button("Ejecutar Análisis")

    if run and selected:
        df_sel = df[df['Account Name'].isin(selected)]
        freq_unit = st.selectbox("Frecuencia en:", ['Semanas','Meses'])
        annual_df, future_df = analyze_df(df_sel, freq_unit)

        st.subheader("Muestreo Anual")
        st.dataframe(annual_df)
        st.subheader("Fechas Futuras (una por fila)")
        st.dataframe(future_df)

        data = to_excel(annual_df, future_df)
        st.download_button(
            "Descargar Excel",
            data=data,
            file_name="sampling_analysis_full.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
