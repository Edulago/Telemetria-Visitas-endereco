import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO  # para usar no hash_funcs do cache

def _normaliza(txt: str) -> str:
    if isinstance(txt, str):
        txt = unicodedata.normalize("NFKD", txt)
        txt = txt.encode("ascii", "ignore").decode("ascii")
        return txt.lower().strip()
    return ""

@st.cache_data   
def ler_telemetria(path: str) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        usecols=["Data Comunicação", "Endereços"], 
        engine="openpyxl"
    )
    df["Data"] = pd.to_datetime(
        df["Data Comunicação"], dayfirst=True, errors="coerce"
    ).dt.date
    df.dropna(subset=["Data"], inplace=True)
    return df[["Data", "Endereços"]]

@st.cache_data(           
    hash_funcs={BytesIO: lambda _: None} 
)
def ler_visitas(file_bytes: BytesIO) -> pd.DataFrame:
    df = pd.read_excel(
        file_bytes,
        usecols=["Data de Início", "Referente a",
                 "Status da Atividade", "Proprietário"],
        engine="openpyxl"
    )
    df["status_norm"] = df["Status da Atividade"].map(_normaliza)
    df = df[df["status_norm"].str.contains("conclu", na=False)]
    df["Data"] = pd.to_datetime(
        df["Data de Início"], errors="coerce"
    ).dt.date
    df.dropna(subset=["Data"], inplace=True)
    return df[["Data", "Proprietário", "Referente a"]]

@st.cache_data 
def cruza(visitas_df: pd.DataFrame, tele_df: pd.DataFrame) -> pd.DataFrame:
    cruz = pd.merge(visitas_df, tele_df, on="Data", how="inner")
    return cruz.drop_duplicates()

st.title("Telemetria × Visitas")

upload = st.file_uploader(
    "Envie o Excel de visitas (Atividades PJ)", type=["xlsx", "xls"]
)

try:
    tele = pd.concat(
        [ler_telemetria("DataFrame.xlsx"),
         ler_telemetria("DataFrame2.xlsx")],
        ignore_index=True
    )
except Exception as e:
    st.error(f"Erro ao ler telemetria: {e}")
    st.stop()

if upload:
    try:
        vis = ler_visitas(upload)       
    except Exception as e:
        st.error(f"Erro ao ler visitas: {e}")
        st.stop()

    res = cruza(vis, tele)              

    if res.empty:
        st.warning("Nenhuma visita concluída coincide com datas da telemetria.")
        st.stop()

    datas_opc = sorted(res["Data"].unique())
    props_opc = sorted(res["Proprietário"].unique())

    data_sel = st.selectbox(
        "Escolha a data:",
        options=datas_opc,
        format_func=lambda d: d.strftime("%d/%m/%Y")
    )
    prop_sel = st.selectbox("Escolha o proprietário:", options=props_opc)

    filtrado = res
    if data_sel:
        filtrado = filtrado[filtrado["Data"] == data_sel]
    if prop_sel:
        filtrado = filtrado[filtrado["Proprietário"] == prop_sel]

    st.subheader("Resultado")
    st.dataframe(filtrado, use_container_width=True)
else:
    st.info("Envie o arquivo de visitas para começar.")
