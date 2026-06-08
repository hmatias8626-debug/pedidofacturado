import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control OC + Resumen MAHLE", layout="wide")

st.title("Control OC vs Facturas / Resumen MAHLE")
st.write(
    "Versión MAHLE interna. Subí el Excel exportado y la app calcula pedido real, "
    "facturado real, resumen anual y cuadro tipo MAHLE por año."
)

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
        primera_linea = texto.splitlines()[0] if texto.splitlines() else ""

        if "\t" in primera_linea:
            sep = "\t"
        elif ";" in primera_linea:
            sep = ";"
        else:
            sep = ","

        return pd.read_csv(io.StringIO(texto), sep=sep)

    return pd.read_excel(archivo)


def preparar(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    requeridas = [
        "OC",
        "prd_id",
        "CODIGO",
        "FLIA",
        "Cantidad Pedida",
        "FACTURADO_TOTAL",
        "NROFAC",
        "INGRESADO_TOTAL",
        "NROING",
    ]

    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        st.error(f"Faltan columnas necesarias: {', '.join(faltan)}")
        st.stop()

    for col in ["OC", "prd_id", "CODIGO", "PROVEEDOR", "FLIA", "NROFAC", "NROING"]:
        if col in df.columns:
            df[col] = normalizar_texto(df[col])

    for col in [
        "Cantidad Pedida",
        "FACTURADO_TOTAL",
        "INGRESADO_TOTAL",
        "$ PEDIDO",
        "$ FACTURADO",
        "$ INGRESADO",
        "COSTO",
        "VENTAS_TOTAL",
        "STOCK",
        "StockSuc",
    ]:
        if col in df.columns:
            df[col] = limpiar_numero(df[col])

    for col in ["NROFAC", "NROING"]:
        df[col] = df[col].replace(["0", "0.0", "nan", "NaN", "None", "", " "], pd.NA)

    for col in ["Fecha Pedido", "Cpbprov_Fecha", "FECHA INGRESO"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    return df


def generar_resumen_oc(df):
    claves = ["OC", "prd_id", "CODIGO"]
    info_extra = [
        c
        for c in ["PROVEEDOR", "FLIA", "Fecha Pedido", "COSTO", "VENTAS_TOTAL", "STOCK", "StockSuc"]
        if c in df.columns
    ]

    pedido = (
        df[claves + info_extra + ["Cantidad Pedida"]]
        .drop_duplicates(claves)
        .rename(columns={"Cantidad Pedida": "CANT_PEDIDA_REAL"})
    )

    facturas_base = df[df["NROFAC"].notna()].copy()
    if len(facturas_base):
        facturas = (
            facturas_base.drop_duplicates(claves + ["NROFAC"])
            .groupby(claves, as_index=False)
            .agg(
                FACTURADO_REAL=("FACTURADO_TOTAL", "sum"),
                FACTURAS=("NROFAC", lambda x: " | ".join(sorted(set(map(str, x)))))
            )
        )
    else:
        facturas = pd.DataFrame(columns=claves + ["FACTURADO_REAL", "FACTURAS"])

    ingresos_base = df[df["NROING"].notna()].copy()
    if len(ingresos_base):
        ingresos = (
            ingresos_base.drop_duplicates(claves + ["NROING"])
            .groupby(claves, as_index=False)
            .agg(
                INGRESADO_REAL=("INGRESADO_TOTAL", "sum"),
                INGRESOS=("NROING", lambda x: " | ".join(sorted(set(map(str, x)))))
            )
        )
    else:
        ingresos = pd.DataFrame(columns=claves + ["INGRESADO_REAL", "INGRESOS"])

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


def base_pedidos(df):
    claves = ["OC", "prd_id", "CODIGO"]
    cols = claves + ["FLIA", "Fecha Pedido", "Cantidad Pedida"]
    if "PROVEEDOR" in df.columns:
        cols.insert(3, "PROVEEDOR")

    b = df[[c for c in cols if c in df.columns]].drop_duplicates(claves).copy()
    b = b.rename(columns={"Cantidad Pedida": "pedido"})
    return b


def base_facturas(df):
    claves = ["OC", "prd_id", "CODIGO", "NROFAC"]
    cols = claves + ["FLIA", "Cpbprov_Fecha", "FACTURADO_TOTAL"]
    if "PROVEEDOR" in df.columns:
        cols.insert(4, "PROVEEDOR")

    b = df[df["NROFAC"].notna()][[c for c in cols if c in df.columns]].copy()
    b = b.drop_duplicates(claves)
    b = b.rename(columns={"FACTURADO_TOTAL": "facturado"})
    return b


def resumen_anual(df):
    pedidos = base_pedidos(df)
    facturas = base_facturas(df)

    p = pd.DataFrame(columns=["Año", "pedido"])
    f = pd.DataFrame(columns=["Año", "facturado"])

    if "Fecha Pedido" in pedidos.columns:
        aux = pedidos.dropna(subset=["Fecha Pedido"]).copy()
        aux["Año"] = aux["Fecha Pedido"].dt.year
        p = aux.groupby("Año", as_index=False).agg(pedido=("pedido", "sum"))

    if "Cpbprov_Fecha" in facturas.columns:
        aux = facturas.dropna(subset=["Cpbprov_Fecha"]).copy()
        aux["Año"] = aux["Cpbprov_Fecha"].dt.year
        f = aux.groupby("Año", as_index=False).agg(facturado=("facturado", "sum"))

    out = p.merge(f, on="Año", how="outer").fillna(0).sort_values("Año")
    out["diferencia"] = out["pedido"] - out["facturado"]
    out["% facturado"] = out.apply(lambda r: 0 if r["pedido"] == 0 else r["facturado"] / r["pedido"], axis=1)

    total = pd.DataFrame([{
        "Año": "TOTAL",
        "pedido": out["pedido"].sum(),
        "facturado": out["facturado"].sum(),
        "diferencia": out["diferencia"].sum(),
        "% facturado": 0 if out["pedido"].sum() == 0 else out["facturado"].sum() / out["pedido"].sum(),
    }])

    return pd.concat([out, total], ignore_index=True)


def resumen_familias_total(df):
    pedidos = base_pedidos(df).groupby("FLIA", as_index=False).agg(pedido=("pedido", "sum"))
    facturas = base_facturas(df).groupby("FLIA", as_index=False).agg(facturado=("facturado", "sum"))
    return pedidos.merge(facturas, on="FLIA", how="outer").fillna(0).sort_values("FLIA")


def resumen_familias_por_anio(df):
    pedidos = base_pedidos(df)
    facturas = base_facturas(df)

    p = pd.DataFrame(columns=["Año", "FLIA", "pedido"])
    f = pd.DataFrame(columns=["Año", "FLIA", "facturado"])

    if "Fecha Pedido" in pedidos.columns:
        aux = pedidos.dropna(subset=["Fecha Pedido"]).copy()
        aux["Año"] = aux["Fecha Pedido"].dt.year
        p = aux.groupby(["Año", "FLIA"], as_index=False).agg(pedido=("pedido", "sum"))

    if "Cpbprov_Fecha" in facturas.columns:
        aux = facturas.dropna(subset=["Cpbprov_Fecha"]).copy()
        aux["Año"] = aux["Cpbprov_Fecha"].dt.year
        f = aux.groupby(["Año", "FLIA"], as_index=False).agg(facturado=("facturado", "sum"))

    return p.merge(f, on=["Año", "FLIA"], how="outer").fillna(0).sort_values(["Año", "FLIA"])


def suma_si(familias, textos, columna):
    if familias.empty:
        return 0

    flia = familias["FLIA"].astype(str).str.upper()
    mask = pd.Series(False, index=familias.index)

    for texto in textos:
        mask = mask | flia.str.contains(texto.upper(), regex=False, na=False)

    return familias.loc[mask, columna].sum()


def cuadro_mahle(familias):
    reglas = [
        ("Filter", ["FILTRO"]),
        ("Mechatronics", ["BATERIA", "CABLE"]),
        ("Batteries", ["BATERIA"]),
        ("Spark Plug Wire", ["CABLE"]),
        ("Powertrain", ["COJINETE", "JUNTA", "GUIA", "ASIENTO", "BOBINA", "CAMISA", "CONJUNTO", "ARO", "TURBO", "VALVULA"]),
        ("Bearings", ["COJINETE"]),
        ("Gasket", ["JUNTA"]),
        ("Guides & Seats", ["GUIA", "ASIENTO"]),
        ("Ignition Coil", ["BOBINA"]),
        ("Liners", ["CAMISA"]),
        ("Piston", ["CONJUNTO"]),
        ("Rings", ["ARO"]),
        ("Turbocharger", ["TURBO"]),
        ("Valve", ["VALVULA"]),
        ("Thermal", ["COMPRESOR", "TERMOS", "TERMOSTATO"]),
        ("Compressor A/C", ["COMPRESOR"]),
        ("Thermostat", ["TERMOS", "TERMOSTATO"]),
    ]

    filas = []
    for rubro, textos in reglas:
        filas.append({
            "Rubro": rubro,
            "pedido": suma_si(familias, textos, "pedido"),
            "facturado": suma_si(familias, textos, "facturado"),
        })

    res = pd.DataFrame(filas)

    total_pedido = res.loc[res["Rubro"].isin(["Filter", "Mechatronics", "Powertrain", "Thermal"]), "pedido"].sum()
    total_facturado = res.loc[res["Rubro"].isin(["Filter", "Mechatronics", "Powertrain", "Thermal"]), "facturado"].sum()

    res.loc[len(res)] = ["TOTAL", total_pedido, total_facturado]

    return res


def cuadro_mahle_por_anio(familias_anio):
    resultados = []

    for anio in sorted(familias_anio["Año"].dropna().unique()):
        fam = familias_anio[familias_anio["Año"] == anio][["FLIA", "pedido", "facturado"]].copy()
        tmp = cuadro_mahle(fam)
        tmp.insert(0, "Año", anio)
        resultados.append(tmp)

    if not resultados:
        return pd.DataFrame(columns=["Año", "Rubro", "pedido", "facturado"])

    return pd.concat(resultados, ignore_index=True)


def escribir_excel(resumen_oc, resumen_anual_df, familias_total, familias_anio, mahle_total, mahle_anio, df_original):
    salida = io.BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen_anual_df.to_excel(writer, index=False, sheet_name="Resumen anual")
        mahle_anio.to_excel(writer, index=False, sheet_name="MAHLE por año")
        mahle_total.to_excel(writer, index=False, sheet_name="MAHLE total")
        familias_anio.to_excel(writer, index=False, sheet_name="Familias por año")
        familias_total.to_excel(writer, index=False, sheet_name="Familias total")
        resumen_oc.to_excel(writer, index=False, sheet_name="Detalle OC")
        df_original.to_excel(writer, index=False, sheet_name="Datos originales")

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"

            for cell in ws[1]:
                cell.style = "Headline 3"

            for col in ws.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 38)

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        header = str(ws.cell(row=1, column=cell.column).value)
                        if "%" in header:
                            cell.number_format = "0.00%"
                        else:
                            cell.number_format = '#,##0.00'

    return salida.getvalue()


if archivo:
    df = leer_archivo(archivo)
    df = preparar(df)

    resumen_oc = generar_resumen_oc(df)
    resumen_anual_df = resumen_anual(df)
    familias_total = resumen_familias_total(df)
    familias_anio = resumen_familias_por_anio(df)
    mahle_total = cuadro_mahle(familias_total)
    mahle_anio = cuadro_mahle_por_anio(familias_anio)

    st.subheader("Resumen")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Líneas originales", len(df))
    c2.metric("OC/productos reales", len(resumen_oc))
    c3.metric("Con pendiente factura", int((resumen_oc["PEND_FACTURA"] > 0).sum()))
    c4.metric("Con pendiente ingreso", int((resumen_oc["PEND_INGRESO"] > 0).sum()))

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "MAHLE por año",
        "MAHLE total",
        "Resumen anual",
        "Familias por año",
        "Familias total",
        "Detalle OC",
    ])

    with tab1:
        anios = ["TODOS"] + [str(x) for x in sorted(mahle_anio["Año"].dropna().unique())]
        anio_sel = st.selectbox("Año", anios)

        vista = mahle_anio.copy()
        if anio_sel != "TODOS":
            vista = vista[vista["Año"].astype(str) == anio_sel]

        st.dataframe(vista, use_container_width=True, height=600)

    with tab2:
        st.dataframe(mahle_total, use_container_width=True, height=600)

    with tab3:
        st.dataframe(resumen_anual_df, use_container_width=True, height=400)

    with tab4:
        st.dataframe(familias_anio, use_container_width=True, height=600)

    with tab5:
        st.dataframe(familias_total, use_container_width=True, height=600)

    with tab6:
        estados = ["TODOS"] + sorted(resumen_oc["ESTADO"].unique().tolist())
        estado_sel = st.selectbox("Filtrar por estado", estados)

        vista = resumen_oc.copy()
        if estado_sel != "TODOS":
            vista = vista[vista["ESTADO"] == estado_sel]

        buscar = st.text_input("Buscar por OC, código, proveedor o familia")
        if buscar:
            b = buscar.lower()
            vista = vista[
                vista.astype(str).apply(
                    lambda fila: fila.str.lower().str.contains(b, regex=False).any(),
                    axis=1,
                )
            ]

        st.dataframe(vista, use_container_width=True, height=600)

    excel = escribir_excel(
        resumen_oc=resumen_oc,
        resumen_anual_df=resumen_anual_df,
        familias_total=familias_total,
        familias_anio=familias_anio,
        mahle_total=mahle_total,
        mahle_anio=mahle_anio,
        df_original=df,
    )

    st.download_button(
        "Descargar Excel procesado",
        data=excel,
        file_name="control_oc_facturas_mahle_interno.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Esperando archivo...")
