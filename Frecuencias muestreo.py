import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ======================
# CACHES PARA OPTIMIZAR
# ======================
@st.cache_data
def load_data(uploaded_file):
    """
    Carga el archivo CSV o XLSX y retorna solo las columnas necesarias.
    """
    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        df = pd.read_csv(uploaded_file, parse_dates=["Date Sampled"])
    else:
        df = pd.read_excel(uploaded_file, parse_dates=["Date Sampled"])
    cols = [
        "Unit ID",
        "Asset ID",
        "Account Name",
        "Sample Bottle ID",
        "Date Sampled",
        "Asset Class",
    ]
    return df[cols]

@st.cache_data
def analyze_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula, para cada equipo y cada año:
      - número de muestras
      - meses entre muestras (12 / count) -> frecuencia en meses
    Añade además una recomendación promedio (meses) por equipo.
    """
    df = df.dropna(subset=["Date Sampled"]).copy()
    df["Year"] = df["Date Sampled"].dt.year

    # Años desde 2021 hasta el actual
    current_year = datetime.today().year
    years = list(range(2021, current_year + 1))

    # 1) Conteo de muestras únicas por año
    cnt = (
        df
        .groupby(["Unit ID", "Asset ID", "Asset Class", "Account Name", "Year"])["Sample Bottle ID"]
        .nunique()
        .reset_index(name="Samples")
    )

    # 2) Pivot: muestras por año
    pivot = (
        cnt
        .pivot_table(
            index=["Unit ID", "Asset ID", "Asset Class", "Account Name"],
            columns="Year",
            values="Samples",
            fill_value=0
        )
        .reindex(columns=years, fill_value=0)
        .reset_index()
    )

    # 3) Calcular frecuencia en meses: 12 / muestras
    freq_df = pivot[["Unit ID", "Asset ID", "Asset Class", "Account Name"]].copy()
    for y in years:
        freq_df[f"{y} (meses)"] = pivot[y].apply(
            lambda c: round(12 / c, 1) if c > 0 else None
        )

    # 4) Recomendación general: promedio de frecuencias de año
    freq_cols = [f"{y} (meses)" for y in years]
    freq_df["Recommended Frequency (meses)"] = (
        freq_df[freq_cols]
        .mean(axis=1, skipna=True)
        .round(1)
    )

    # 5) Unir conteos y frecuencias
    result = pivot.merge(
        freq_df.drop(columns=["Unit ID", "Asset ID", "Asset Class", "Account Name"]),
        left_index=True,
        right_index=True
    )

    return result

def to_excel(df: pd.DataFrame) -> bytes:
    """
    Genera un Excel con la hoja 'Frecuencia Anual'.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Frecuencia Anual")
    return output.getvalue()

# ====================
# INTERFAZ STREAMLIT
# ====================
st.title("Análisis de Frecuencia de Muestreo (meses)")
st.markdown(
    "- Formato MobilServ\n"
    "- Columnas obligatorias: Unit ID, Asset ID, Account Name, Sample Bottle ID, Date Sampled, Asset Class\n"
    "- Sube CSV o XLSX"
)

uploaded = st.file_uploader("1) Sube tu archivo MobilServ", type=["csv", "xlsx"])
if not uploaded:
    st.info("**2) Selecciona operaciones** (primero sube el archivo).")
    st.stop()

df = load_data(uploaded)

# 2) Selector de operaciones (vacío al inicio)
ops = sorted(df["Account Name"].dropna().unique())
selected_ops = st.multiselect(
    "2) Selecciona las operaciones (Account Name)",
    options=ops,
    default=[],
)

if not selected_ops:
    st.info("Por favor selecciona al menos una operación para continuar.")
    st.stop()

# Filtrar y analizar
df_sel = df[df["Account Name"].isin(selected_ops)]
result_df = analyze_df(df_sel)

# Mostrar resultados
st.subheader("Frecuencia de muestreo por año (en meses)")
st.dataframe(result_df, use_container_width=True)

# Descargar
excel_data = to_excel(result_df)
st.download_button(
    label="📥 Descargar resultados en Excel",
    data=excel_data,
    file_name="sampling_frequency_months.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
