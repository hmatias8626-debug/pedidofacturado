import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control OC vs Facturas", layout="wide")

st.title("Control de OC vs Facturas / Ingresos")
st.write("Subí el Excel exportado y la app calcula lo pedido real por OC/producto, lo facturado, lo ingresado y los pendientes.")

archivo = st.file_uploader("Subir archivo Excel", type=["xlsx", "xls", "csv", "txt"])

def limpiar_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)

def normalizar_texto(serie):
    return serie.astype(str).str.strip()

def leer_archivo(archivo):
    nombre = archivo.name.lower()

    if nombre.endswith(".csv") or nombre.endswith(".txt"):
        contenido = archivo.getvalue()
        texto = contenido.decode("utf-8", errors="ignore")
        # intenta detectar si viene tabulado o separado por ;
        sep = "\t" if "\t" in texto.splitlines()[0] else ";"
        return pd.read_csv(io.StringIO(texto), sep=sep)

    return pd.read_excel(archivo)

def preparar(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    requeridas = ["OC", "prd_id", "CODIGO", "Cantidad Pedida", "FACTURADO_TOTAL", "NROFAC", "INGRESADO_TOTAL", "NROING"]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        st.error(f"Faltan columnas necesarias: {', '.join(faltan)}")
        st.stop()

    for col in ["OC", "prd_id", "CODIGO", "PROVEEDOR", "FLIA", "NROFAC", "NROING"]:
        if col in df.columns:
            df[col] = normalizar_texto(df[col])

    for col in ["Cantidad Pedida", "FACTURADO_TOTAL", "INGRESADO_TOTAL", "$ PEDIDO", "$ FACTURADO", "$ INGRESADO"]:
        if col in df.columns:
            df[col] = limpiar_numero(df[col])

    # valores tipo 0, vacío o nan no cuentan como documento real
    for col in ["NROFAC", "NROING"]:
        df[col] = df[col].replace(["0", "0.0", "nan", "None", ""], pd.NA)

    return df

def generar_resumen(df):
    claves = ["OC", "prd_id", "CODIGO"]
    info_extra = [c for c in ["PROVEEDOR", "FLIA", "Fecha Pedido", "COSTO", "VENTAS_TOTAL", "STOCK", "StockSuc"] if c in df.columns]

    # Pedido: una sola vez por OC + producto.
    pedido_cols = claves + info_extra + ["Cantidad Pedida"]
    pedido = (
        df[pedido_cols]
        .drop_duplicates(claves)
        .rename(columns={"Cantidad Pedida": "CANT_PEDIDA_REAL"})
    )

    # Facturas: una sola vez por OC + producto + NROFAC.
    facturas_base = df[df["NROFAC"].notna()].copy()
    facturas = (
        facturas_base
        .drop_duplicates(claves + ["NROFAC"])
        .groupby(claves, as_index=False)
        .agg(
            FACTURADO_REAL=("FACTURADO_TOTAL", "sum"),
            FACTURAS=("NROFAC", lambda x: " | ".join(sorted(set(map(str, x)))))
        )
    )

    # Ingresos: una sola vez por OC + producto + NROING.
    ingresos_base = df[df["NROING"].notna()].copy()
    ingresos = (
        ingresos_base
        .drop_duplicates(claves + ["NROING"])
        .groupby(claves, as_index=False)
        .agg(
            INGRESADO_REAL=("INGRESADO_TOTAL", "sum"),
            INGRESOS=("NROING", lambda x: " | ".join(sorted(set(map(str, x)))))
        )
    )

    resumen = pedido.merge(facturas, on=claves, how="left").merge(ingresos, on=claves, how="left")

    resumen["FACTURADO_REAL"] = resumen["FACTURADO_REAL"].fillna(0)
    resumen["INGRESADO_REAL"] = resumen["INGRESADO_REAL"].fillna(0)
    resumen["FACTURAS"] = resumen["FACTURAS"].fillna("")
    resumen["INGRESOS"] = resumen["INGRESOS"].fillna("")

    resumen["PEND_FACTURA"] = resumen["CANT_PEDIDA_REAL"] - resumen["FACTURADO_REAL"]
    resumen["PEND_INGRESO"] = resumen["FACTURADO_REAL"] - resumen["INGRESADO_REAL"]

    def estado(row):
        if row["FACTURADO_REAL"] == 0:
            return "SIN FACTURAR"
        if row["PEND_FACTURA"] > 0:
            return "FACTURADO PARCIAL"
        if row["PEND_FACTURA"] < 0:
            return "FACTURADO DE MÁS"
        if row["PEND_INGRESO"] > 0:
            return "FACTURADO SIN INGRESAR COMPLETO"
        if row["PEND_INGRESO"] < 0:
            return "INGRESADO DE MÁS"
        return "OK"

    resumen["ESTADO"] = resumen.apply(estado, axis=1)

    columnas_finales = claves + info_extra + [
        "CANT_PEDIDA_REAL",
        "FACTURADO_REAL",
        "INGRESADO_REAL",
        "PEND_FACTURA",
        "PEND_INGRESO",
        "ESTADO",
        "FACTURAS",
        "INGRESOS",
    ]

    return resumen[columnas_finales].sort_values(["OC", "CODIGO", "prd_id"])

if archivo:
    df = leer_archivo(archivo)
    df = preparar(df)
    resumen = generar_resumen(df)

    st.subheader("Resumen")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Líneas originales", len(df))
    c2.metric("OC/productos reales", len(resumen))
    c3.metric("Con pendiente factura", int((resumen["PEND_FACTURA"] > 0).sum()))
    c4.metric("Con pendiente ingreso", int((resumen["PEND_INGRESO"] > 0).sum()))

    estados = ["TODOS"] + sorted(resumen["ESTADO"].unique().tolist())
    estado_sel = st.selectbox("Filtrar por estado", estados)

    vista = resumen.copy()
    if estado_sel != "TODOS":
        vista = vista[vista["ESTADO"] == estado_sel]

    buscar = st.text_input("Buscar por OC, código, proveedor o familia")
    if buscar:
        b = buscar.lower()
        vista = vista[
            vista.astype(str).apply(lambda fila: fila.str.lower().str.contains(b, regex=False).any(), axis=1)
        ]

    st.dataframe(vista, use_container_width=True, height=550)

    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen OC")
        df.to_excel(writer, index=False, sheet_name="Datos originales")

    st.download_button(
        "Descargar Excel procesado",
        data=salida.getvalue(),
        file_name="control_oc_facturas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Esperando archivo...")