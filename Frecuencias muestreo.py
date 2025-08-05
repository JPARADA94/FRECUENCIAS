import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ======================
# CACHE PARA OPTIMIZAR
# ======================
@st.cache_data
def load_data(uploaded_file):
    """
    Carga CSV o XLSX y retorna solo las columnas necesarias.
    """
    fname = uploaded_file.name.lower()
    if fname.endswith(".csv"):
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
    1) Cuenta muestras por año (pivot desde 2021).
    2) Calcula intervalo entre muestras y su mediana.
    3) Recomienda frecuencia en meses = mediana_días / 30.
    """
    df = df.dropna(subset=["Date Sampled"]).copy()
    df["Year"] = df["Date Sampled"].dt.year

    # Definir años de 2021 hasta actual
    current_year = datetime.today().year
    years = list(range(2021, current_year + 1))

    # 1) Conteo de muestras únicas por año
    cnt = (
        df.groupby(
            ["Unit ID", "Asset ID", "Asset Class", "Account Name", "Year"]
        )["Sample Bottle ID"]
        .nunique()
        .reset_index(name="Samples")
    )
    pivot = (
        cnt.pivot_table(
            index=["Unit ID", "Asset ID", "Asset Class", "Account Name"],
            columns="Year",
            values="Samples",
            fill_value=0,
        )
        .reindex(columns=years, fill_value=0)
        .reset_index()
    )

    # 2) Calcular todos los intervalos en días
    df_sorted = df.sort_values(
        ["Unit ID", "Asset ID", "Date Sampled"]
    )
    df_sorted["Prev"] = (
        df_sorted.groupby(
            ["Unit ID", "Asset ID"]
        )["Date Sampled"]
        .shift(1)
    )
    df_sorted["Interval Days"] = (
        df_sorted["Date Sampled"] - df_sorted["Prev"]
    ).dt.days

    # 3) Mediana de intervalos
    med = (
        df_sorted.groupby(
            ["Unit ID", "Asset ID", "Asset Class", "Account Name"]
        )["Interval Days"]
        .median()
        .reset_index(name="Median Interval (Days)")
    )
    # 4) Frecuencia recomendada en meses
    med["Recommended Frequency (Months)"] = (
        med["Median Interval (Days)"] / 30
    ).round(1)

    # 5) Unir conteos anuales con mediana y recomendación
    result = pivot.merge(
        med[
            [
                "Unit ID",
                "Asset ID",
                "Median Interval (Days)",
                "Recommended Frequency (Months)",
            ]
        ],
        on=["Unit ID", "Asset ID"],
        how="left",
    )

    return result

def to_excel(df: pd.DataFrame) -> bytes:
    """
    Genera un Excel con hoja 'Frecuencia Mensual Recomendada'.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Recommended Sampling",
        )
    return output.getvalue()

# ====================
# INTERFAZ STREAMLIT
# ====================
st.title("Análisis de Frecuencia de Muestreo (Mediana)")
st.markdown(
    "- Formato MobilServ\n"
    "- Columnas obligatorias: Unit ID, Asset ID, Account Name,\n"
    "  Sample Bottle ID, Date Sampled, Asset Class\n"
    "- Sube CSV o XLSX"
)

# Paso 1: subir archivo
uploaded = st.file_uploader("1) Sube tu archivo MobilServ", type=["csv", "xlsx"])
if not uploaded:
    st.info("Por favor sube primero el archivo.")
    st.stop()

df = load_data(uploaded)

# Paso 2: selección de operaciones (vacío al inicio)
ops = sorted(df["Account Name"].dropna().unique())
selected_ops = st.multiselect(
    "2) Selecciona operaciones (Account Name)",
    options=ops,
    default=[],
)
if not selected_ops:
    st.info("Selecciona al menos una operación para continuar.")
    st.stop()

# Filtrar y analizar
df_sel = df[df["Account Name"].isin(selected_ops)]
result_df = analyze_df(df_sel)

# Mostrar resultados
st.subheader("Frecuencia recomendada basada en mediana de intervalos")
st.dataframe(result_df, use_container_width=True)

# Descargar Excel
excel_bytes = to_excel(result_df)
st.download_button(
    "📥 Descargar recomendación en Excel",
    data=excel_bytes,
    file_name="sampling_median_recommendation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
