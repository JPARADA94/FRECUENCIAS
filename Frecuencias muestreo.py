import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Función principal de análisis
def analyze_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['Date Sampled'] = pd.to_datetime(df['Date Sampled'], errors='coerce')
    df = df.dropna(subset=['Date Sampled'])
    df['Year'] = df['Date Sampled'].dt.year

    # 1) Conteo de muestras únicas por año
    count_per_year = (
        df
        .groupby(['Unit ID', 'Asset ID', 'Account Name', 'Asset Class', 'Year'])
        ['Sample Bottle ID']
        .nunique()
        .reset_index(name='Samples per Year')
    )

    # 2) Cálculo de intervalos de muestreo
    df_sorted = df.sort_values(['Unit ID', 'Asset ID', 'Date Sampled'])
    df_sorted['Prev Date'] = (
        df_sorted
        .groupby(['Unit ID', 'Asset ID'])['Date Sampled']
        .shift(1)
    )
    df_sorted['Interval Days'] = (
        df_sorted['Date Sampled'] - df_sorted['Prev Date']
    ).dt.days

    # 3) Intervalo promedio por equipo
    avg_interval = (
        df_sorted
        .groupby(['Unit ID', 'Asset ID', 'Account Name'])['Interval Days']
        .mean()
        .reset_index(name='Avg Interval (Days)')
    )

    # Asociar Asset Class
    asset_classes = df[['Unit ID', 'Asset ID', 'Asset Class']].drop_duplicates()
    avg_interval = avg_interval.merge(
        asset_classes, on=['Unit ID', 'Asset ID'], how='left'
    )

    # 4) Z-score manual del intervalo por Asset Class
    avg_interval['Interval Z-Score'] = (
        avg_interval
        .groupby('Asset Class')['Avg Interval (Days)']
        .transform(lambda x: (x - x.mean()) / x.std(ddof=0))
    )

    # 5) Unir resultados finales
    result = count_per_year.merge(
        avg_interval[['Unit ID', 'Asset ID', 'Avg Interval (Days)', 'Interval Z-Score']],
        on=['Unit ID', 'Asset ID'], how='left'
    )
    return result

# Función para convertir DataFrame a Excel bytes
def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Analysis')
    return output.getvalue()

# Interfaz Streamlit
st.title("Análisis de Frecuencia de Muestreo por Equipo")
st.markdown(
    "Sube el archivo **Excel en formato MobilServ** con las columnas:"
    " Unit ID, Asset ID, Account Name, Sample Bottle ID, Date Sampled y Asset Class."
)

uploaded_file = st.file_uploader(
    "Sube el archivo Excel en formato MobilServ",
    type=["xlsx"]
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, parse_dates=['Date Sampled'])
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    # Selector de operación
    operations = sorted(df['Account Name'].dropna().unique())
    selected_op = st.selectbox("Selecciona la operación (Account Name)", operations)
    df_op = df[df['Account Name'] == selected_op]

    st.success(f"Operación '{selected_op}' cargada con {len(df_op)} registros.")

    # Análisis filtrado
    result = analyze_df(df_op)

    # Cálculo de fechas futuras
    op_interval = int(round(result['Avg Interval (Days)'].mean())) if not result.empty else 30
    today = pd.Timestamp.today().normalize()
    last_date = df_op['Date Sampled'].max().normalize()
    start_date = max(today, last_date)
    end_date = pd.Timestamp(year=today.year + 1, month=12, day=31)

    future_dates = []
    current = start_date
    while current <= end_date:
        current += pd.Timedelta(days=op_interval)
        if current.weekday() >= 5:  # Si es fin de semana
            current += pd.Timedelta(days=(7 - current.weekday()))
        future_dates.append(current.date())

    dates_str = ", ".join(str(d) for d in future_dates)
    result['Future Sample Dates'] = dates_str

    # Mostrar y descargar
    st.subheader("Resultado del Análisis")
    st.dataframe(result)

    excel_data = to_excel(result)
    st.download_button(
        label="Descargar resultado en Excel",
        data=excel_data,
        file_name="sampling_analysis_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
