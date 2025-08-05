import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Función principal de análisis
def analyze_df(df: pd.DataFrame, freq_unit: str) -> (pd.DataFrame, pd.DataFrame):
    df = df.copy()
    df['Date Sampled'] = pd.to_datetime(df['Date Sampled'], errors='coerce')
    df = df.dropna(subset=['Date Sampled'])

    # Rango de años de interés: desde 2021 hasta año actual
    current_year = datetime.today().year
    years = list(range(2021, current_year + 1))
    df['Year'] = df['Date Sampled'].dt.year

    # 1) Conteo de muestras únicas por año y pivote
    count = (
        df
        .groupby(['Unit ID', 'Asset ID', 'Asset Class', 'Account Name', 'Year'])
        ['Sample Bottle ID']
        .nunique()
        .reset_index(name='Count')
    )
    pivot = (
        count
        .pivot_table(
            index=['Unit ID', 'Asset ID', 'Asset Class', 'Account Name'],
            columns='Year',
            values='Count',
            fill_value=0
        )
        .reindex(columns=years, fill_value=0)
        .reset_index()
    )

    # 2) Cálculo de intervalos de muestreo (días)
    df_sorted = df.sort_values(['Unit ID', 'Asset ID', 'Date Sampled'])
    df_sorted['Prev Date'] = (
        df_sorted.groupby(['Unit ID', 'Asset ID'])['Date Sampled'].shift(1)
    )
    df_sorted['Interval Days'] = (
        df_sorted['Date Sampled'] - df_sorted['Prev Date']
    ).dt.days

    avg_interval = (
        df_sorted
        .groupby(['Unit ID', 'Asset ID', 'Asset Class', 'Account Name'])['Interval Days']
        .mean()
        .reset_index(name='Avg Interval (Days)')
    )

    # 3) Recomendar frecuencia según unidad
    if freq_unit == 'Semanas':
        avg_interval['Recommended Frequency'] = (
            (avg_interval['Avg Interval (Days)'] / 7)
            .round(1)
            .astype(str) + ' semanas'
        )
    else:
        avg_interval['Recommended Frequency'] = (
            (avg_interval['Avg Interval (Days)'] / 30)
            .round(1)
            .astype(str) + ' meses'
        )

    # 4) Fechas futuras una por fila
    future_list = []
    for _, row in avg_interval.iterrows():
        unit = row['Unit ID']
        asset = row['Asset ID']
        op = row['Account Name']
        asset_class = row['Asset Class']
        interval_days = row['Avg Interval (Days)']
        today = pd.Timestamp.today().normalize()
        last_date = df_sorted[
            (df_sorted['Unit ID'] == unit) &
            (df_sorted['Asset ID'] == asset)
        ]['Date Sampled'].max().normalize()
        start = max(today, last_date)
        end = pd.Timestamp(year=today.year + 1, month=12, day=31)
        next_date = start
        while next_date <= end:
            next_date += pd.Timedelta(days=interval_days)
            if next_date.weekday() >= 5:  # ajustar fin de semana
                next_date += pd.Timedelta(days=(7 - next_date.weekday()))
            future_list.append({
                'Unit ID': unit,
                'Asset ID': asset,
                'Asset Class': asset_class,
                'Account Name': op,
                'Future Sample Date': next_date.date()
            })
    future_df = pd.DataFrame(future_list)

    # 5) Unir análisis anual con frecuencia
    result = pivot.merge(
        avg_interval[['Unit ID', 'Asset ID', 'Recommended Frequency']],
        on=['Unit ID', 'Asset ID'],
        how='left'
    )
    return result, future_df

# Función para convertir DataFrames a Excel bytes con dos hojas
def to_excel(df1: pd.DataFrame, df2: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df1.to_excel(writer, index=False, sheet_name='Muestreo Anual')
        df2.to_excel(writer, index=False, sheet_name='Fechas Futuras')
    return output.getvalue()

# Interfaz Streamlit
st.title("Análisis de Frecuencia de Muestreo por Operaciones")
st.markdown(
    "Sube el archivo **Excel en formato MobilServ** con las columnas:\n"
    "- Unit ID\n"
    "- Asset ID\n"
    "- Account Name\n"
    "- Sample Bottle ID\n"
    "- Date Sampled\n"
    "- Asset Class"
)

uploaded_file = st.file_uploader(
    "Sube el archivo Excel en formato MobilServ",
    type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file, parse_dates=['Date Sampled'])

    # Selector multiselección de operaciones
    ops = sorted(df['Account Name'].dropna().unique())
    selected_ops = st.multiselect(
        "Selecciona las operaciones (Account Name)",
        options=ops,
        default=ops
    )
    df_sel = df[df['Account Name'].isin(selected_ops)]
    st.success(f"Operaciones seleccionadas: {', '.join(selected_ops)} (registros: {len(df_sel)})")

    # Elegir unidad de frecuencia
    freq_unit = st.selectbox(
        "Unidad de frecuencia recomendada",
        ['Semanas', 'Meses']
    )

    # Ejecutar análisis
    annual_df, future_df = analyze_df(df_sel, freq_unit)

    # Mostrar tablas
    st.subheader("Muestreo Anual (por año)")
    st.dataframe(annual_df)

    st.subheader("Fechas de Toma de Muestras Futuras")
    st.dataframe(future_df)

    # Botón de descarga
    excel_data = to_excel(annual_df, future_df)
    st.download_button(
        label="Descargar resultados en Excel",
        data=excel_data,
        file_name="sampling_analysis_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
