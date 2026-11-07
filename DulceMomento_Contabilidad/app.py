
import io
import json
import math
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# TEMA 1: CONFIGURACION
# ==========================================

st.set_page_config(
    page_title="Dulce Momento | Contabilidad II",
    page_icon="📊",
    layout="wide"
)

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "productos_estandar_v2.json"
ARCHIVO_ANTERIOR = CARPETA / "productos_estandar.json"

# ==========================================
# TEMA 2: DISEÑO GRIS PERLA
# ==========================================

st.markdown("""
<style>
.stApp,
[data-testid="stAppViewContainer"] {
    background-color: #E7E8EC !important;
    color: #292B32 !important;
}

[data-testid="stHeader"] {
    background-color: #E7E8EC !important;
}

[data-testid="stSidebar"],
[data-testid="stSidebarContent"] {
    background-color: #D9DBE1 !important;
}

/* Solo texto: no modificar fuentes de iconos */

.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4,
.stApp p,
.stApp label,
.stApp li,
.stApp input,
.stApp textarea {
    font-family: "Segoe UI", Arial, sans-serif;
}

.stApp h1,
.stApp h2,
.stApp h3 {
    color: #30323A !important;
    font-weight: 700;
}

.stApp h1 {
    font-size: 32px !important;
}

.stApp h2 {
    font-size: 25px !important;
}

.stApp h3 {
    font-size: 21px !important;
}

.stApp p,
.stApp li,
.stApp label {
    color: #383A43;
}

[data-testid="stWidgetLabel"] p {
    color: #30323A !important;
    font-size: 16px !important;
    font-weight: 600;
}

[data-testid="stCaptionContainer"] p {
    color: #555862 !important;
    font-size: 14px !important;
}

.stApp input,
.stApp textarea {
    background-color: #F7F7F9 !important;
    color: #242630 !important;
    font-size: 16px !important;
}

.stButton button,
.stDownloadButton button,
[data-testid="stFormSubmitButton"] button {
    background-color: #F4F4F6 !important;
    color: #30323A !important;
    border: 1px solid #A8AAB4 !important;
    border-radius: 10px !important;
    min-height: 43px;
}

.stButton button p,
.stDownloadButton button p,
[data-testid="stFormSubmitButton"] button p {
    color: #30323A !important;
    font-size: 15px !important;
    font-weight: 600 !important;
}

.stButton button:hover,
.stDownloadButton button:hover {
    background-color: #E5E0EB !important;
    border-color: #786989 !important;
}

.stButton button[kind="primary"] {
    background-color: #71627E !important;
    border-color: #71627E !important;
}

.stButton button[kind="primary"] p {
    color: #FFFFFF !important;
}

[data-testid="stMetric"] {
    background-color: #F4F4F6;
    border: 1px solid #D0D1D8;
    border-radius: 12px;
    padding: 16px;
}

[data-testid="stMetricLabel"] p {
    color: #50525B !important;
}

[data-testid="stMetricValue"] {
    color: #292B32 !important;
    font-size: 24px !important;
}

.titulo-campo {
    color: #30323A;
    font-size: 16px;
    font-weight: 650;
    margin-bottom: 8px;
}

.ayuda {
    background-color: #F4F4F6;
    border: 1px solid #D0D1D8;
    border-left: 4px solid #786989;
    border-radius: 10px;
    padding: 18px;
    color: #30323A;
    line-height: 1.7;
}

.ayuda p {
    color: #30323A !important;
}

</style>
""", unsafe_allow_html=True)

# ==========================================
# TEMA 3: FUNCIONES GENERALES
# ==========================================

def numero(valor):
    try:
        valor = float(valor)
        if math.isfinite(valor) and valor >= 0:
            return valor
    except (ValueError, TypeError):
        pass
    return 0.0


def dinero(valor):
    return f"Q{valor:,.2f}"


def clasificar(valor):
    if abs(valor) < 0.000001:
        return "Sin variación"
    return "Desfavorable" if valor > 0 else "Favorable"


def ingrediente(nombre="", cantidad=0, unidad="g"):
    return {
        "Ingrediente": nombre,
        "Unidad": unidad,
        "Cantidad estándar": cantidad,
        "Precio estándar Q": 0.0,
        "Cantidad real": 0.0,
        "Precio real Q": 0.0,
        "Proveedor": ""
    }


# ==========================================
# TEMA 4: CREACION DE PRODUCTOS
# ==========================================

def crear_producto(nombre, ejemplo=False):

    receta = []
    descripcion = ""

    if ejemplo:
        descripcion = (
            "Pastel artesanal de vainilla, relleno "
            "de crema y fresas frescas. Decorado "
            "con crema batida y fresas naturales. "
            "Rinde aproximadamente 10 a 12 porciones."
        )

        lista = [
            ("Harina", 300, "g"),
            ("Azúcar", 250, "g"),
            ("Huevos", 5, "unidad"),
            ("Leche", 200, "ml"),
            ("Aceite vegetal", 100, "ml"),
            ("Polvo de hornear", 12, "g"),
            ("Vainilla", 10, "ml"),
            ("Fresas frescas", 500, "g"),
            ("Crema para batir", 750, "ml"),
            ("Azúcar glass", 100, "g"),
            ("Base para pastel", 1, "unidad"),
            ("Caja para pastel", 1, "unidad")
        ]

        receta = [
            ingrediente(n, c, u)
            for n, c, u in lista
        ]

    return {
        "nombre": nombre,
        "descripcion": descripcion,
        "cantidad": 20,
        "cantidad_real": 20,
        "precio_venta": 0.0,
        "crecimiento": 10.0,
        "ingredientes": receta,
        "mano_obra": [
            {
                "Actividad": "Preparación y horneado",
                "Horas estándar por unidad": 0.0,
                "Tarifa estándar Q": 0.0,
                "Horas reales totales": 0.0,
                "Tarifa real Q": 0.0
            },
            {
                "Actividad": "Decoración y empaque",
                "Horas estándar por unidad": 0.0,
                "Tarifa estándar Q": 0.0,
                "Horas reales totales": 0.0,
                "Tarifa real Q": 0.0
            }
        ],
        "cif": {
            "Variable estándar por unidad": 0.0,
            "Fijo estándar del período": 0.0,
            "Variable real total": 0.0,
            "Fijo real total": 0.0
        }
    }


# ==========================================
# TEMA 5: GUARDAR Y CARGAR JSON
# ==========================================

def cargar():

    for ruta in [ARCHIVO, ARCHIVO_ANTERIOR]:

        if ruta.exists():
            try:
                datos = json.loads(
                    ruta.read_text(encoding="utf-8")
                )

                if (
                    isinstance(datos.get("productos"), dict)
                    and datos["productos"]
                ):
                    return datos

            except (OSError, ValueError, TypeError):
                pass

    nombre = "Pastel de Fresas con Crema"

    return {
        "productos": {
            nombre: crear_producto(nombre, True)
        }
    }


def guardar():

    contenido = json.dumps(
        st.session_state.catalogo,
        ensure_ascii=False,
        indent=2
    )

    temporal = ARCHIVO.with_suffix(".tmp")

    temporal.write_text(
        contenido,
        encoding="utf-8"
    )

    temporal.replace(ARCHIVO)


def guardar_automaticamente():
    try:
        guardar()
    except OSError as error:
        st.error(
            f"No se pudo guardar el archivo JSON: {error}"
        )


# ==========================================
# TEMA 6: CALCULOS CONTABLES
# ==========================================

def calcular(p):

    q = max(1, int(numero(p["cantidad"])))
    qr = max(1, int(numero(p["cantidad_real"])))
    precio = numero(p["precio_venta"])

    mp_unit = 0.0
    mp_est = 0.0
    mp_real = 0.0
    var_mp_precio = 0.0
    var_mp_cantidad = 0.0

    for f in p["ingredientes"]:

        qe = numero(f.get("Cantidad estándar"))
        pe = numero(f.get("Precio estándar Q"))
        qa = numero(f.get("Cantidad real"))
        pa = numero(f.get("Precio real Q"))

        permitido = qe * qr

        mp_unit += qe * pe
        mp_est += permitido * pe
        mp_real += qa * pa

        var_mp_precio += qa * (pa - pe)
        var_mp_cantidad += (qa - permitido) * pe

    mo_unit = 0.0
    mo_est = 0.0
    mo_real = 0.0
    var_mo_tarifa = 0.0
    var_mo_eficiencia = 0.0

    for f in p["mano_obra"]:

        he = numero(
            f.get("Horas estándar por unidad")
        )
        te = numero(f.get("Tarifa estándar Q"))
        hr = numero(f.get("Horas reales totales"))
        tr = numero(f.get("Tarifa real Q"))

        horas_permitidas = he * qr

        mo_unit += he * te
        mo_est += horas_permitidas * te
        mo_real += hr * tr

        var_mo_tarifa += hr * (tr - te)
        var_mo_eficiencia += (
            hr - horas_permitidas
        ) * te

    cif = p["cif"]

    cv_est = numero(
        cif["Variable estándar por unidad"]
    )
    cf_est = numero(
        cif["Fijo estándar del período"]
    )
    cv_real = numero(
        cif["Variable real total"]
    )
    cf_real = numero(
        cif["Fijo real total"]
    )

    cif_est = qr * cv_est + cf_est
    cif_real = cv_real + cf_real

    total_est = mp_est + mo_est + cif_est
    total_real = mp_real + mo_real + cif_real

    variable_unit = mp_unit + mo_unit + cv_est
    costo_est_unit = variable_unit + cf_est / q
    costo_real_unit = total_real / qr

    ventas = qr * precio
    utilidad = ventas - total_real

    margen = (
        utilidad / ventas * 100
        if ventas > 0 else None
    )

    contribucion = precio - variable_unit

    equilibrio = (
        math.ceil(cf_est / contribucion)
        if contribucion > 0 else None
    )

    q_proyectada = math.ceil(
        q * (1 + numero(p["crecimiento"]) / 100)
    )

    costo_proyectado = (
        q_proyectada * variable_unit + cf_est
    )

    return {
        "q": q,
        "qr": qr,
        "precio": precio,
        "mp_unit": mp_unit,
        "mo_unit": mo_unit,
        "mp_est": mp_est,
        "mp_real": mp_real,
        "mo_est": mo_est,
        "mo_real": mo_real,
        "cif_est": cif_est,
        "cif_real": cif_real,
        "cv_est": cv_est,
        "cf_est": cf_est,
        "total_est": total_est,
        "total_real": total_real,
        "costo_est_unit": costo_est_unit,
        "costo_real_unit": costo_real_unit,
        "variable_unit": variable_unit,
        "var_mp_precio": var_mp_precio,
        "var_mp_cantidad": var_mp_cantidad,
        "var_mo_tarifa": var_mo_tarifa,
        "var_mo_eficiencia": var_mo_eficiencia,
        "var_cif": cif_real - cif_est,
        "variacion_total": total_real - total_est,
        "ventas": ventas,
        "utilidad": utilidad,
        "margen": margen,
        "equilibrio": equilibrio,
        "q_proyectada": q_proyectada,
        "costo_proyectado": costo_proyectado
    }


# ==========================================
# TEMA 7: EXCEL CON FORMULAS
# ==========================================

def crear_excel(catalogo):

    libro = Workbook()
    general = libro.active
    general.title = "Resumen General"

    general.append([
        "Producto",
        "Producción real",
        "Costo estándar permitido",
        "Costo real",
        "Variación total",
        "Ventas",
        "Utilidad",
        "Margen",
        "Punto equilibrio",
        "Producción proyectada"
    ])

    for indice, (nombre, p) in enumerate(
        catalogo["productos"].items(), 1
    ):

        base = f"P{indice:02d}"

        # DATOS

        datos = libro.create_sheet(base + " Datos")

        for fila in [
            ["Concepto", "Valor"],
            ["Producto", nombre],
            ["Descripción", p.get("descripcion", "")],
            ["Producción presupuestada", p["cantidad"]],
            ["Producción real", p["cantidad_real"]],
            ["Precio de venta", p["precio_venta"]],
            ["Crecimiento", p["crecimiento"] / 100],
            [
                "CIF variable estándar unitario",
                numero(
                    p["cif"]["Variable estándar por unidad"]
                )
            ],
            [
                "CIF fijo estándar",
                numero(
                    p["cif"]["Fijo estándar del período"]
                )
            ],
            [
                "CIF variable real",
                numero(p["cif"]["Variable real total"])
            ],
            [
                "CIF fijo real",
                numero(p["cif"]["Fijo real total"])
            ]
        ]:
            datos.append(fila)

        datos["B7"].number_format = "0.00%"

        # MATERIA PRIMA

        mp = libro.create_sheet(base + " MP")

        mp.append([
            "Ingrediente",
            "Unidad",
            "Cantidad estándar unit.",
            "Precio estándar Q",
            "Cantidad real total",
            "Precio real Q",
            "Cantidad permitida",
            "Costo estándar",
            "Costo real",
            "Variación precio",
            "Variación cantidad"
        ])

        for i, f in enumerate(p["ingredientes"], 2):

            mp.append([
                f.get("Ingrediente", ""),
                f.get("Unidad", ""),
                numero(f.get("Cantidad estándar")),
                numero(f.get("Precio estándar Q")),
                numero(f.get("Cantidad real")),
                numero(f.get("Precio real Q")),
                f"=C{i}*'{datos.title}'!B5",
                f"=G{i}*D{i}",
                f"=E{i}*F{i}",
                f"=E{i}*(F{i}-D{i})",
                f"=(E{i}-G{i})*D{i}"
            ])

        fmp = mp.max_row + 1
        mp.cell(fmp, 1, "TOTALES")

        for col in range(8, 12):
            letra = get_column_letter(col)
            mp.cell(
                fmp,
                col,
                f"=SUM({letra}2:{letra}{fmp-1})"
            )

        # MANO DE OBRA

        mo = libro.create_sheet(base + " MO")

        mo.append([
            "Actividad",
            "Horas estándar unit.",
            "Tarifa estándar Q",
            "Horas reales totales",
            "Tarifa real Q",
            "Horas permitidas",
            "Costo estándar",
            "Costo real",
            "Variación tarifa",
            "Variación eficiencia"
        ])

        for i, f in enumerate(p["mano_obra"], 2):

            mo.append([
                f.get("Actividad", ""),
                numero(
                    f.get("Horas estándar por unidad")
                ),
                numero(f.get("Tarifa estándar Q")),
                numero(f.get("Horas reales totales")),
                numero(f.get("Tarifa real Q")),
                f"=B{i}*'{datos.title}'!B5",
                f"=F{i}*C{i}",
                f"=D{i}*E{i}",
                f"=D{i}*(E{i}-C{i})",
                f"=(D{i}-F{i})*C{i}"
            ])

        fmo = mo.max_row + 1
        mo.cell(fmo, 1, "TOTALES")

        for col in range(7, 11):
            letra = get_column_letter(col)
            mo.cell(
                fmo,
                col,
                f"=SUM({letra}2:{letra}{fmo-1})"
            )

        # CIF

        cif = libro.create_sheet(base + " CIF")

        for fila in [
            ["Concepto", "Valor Q"],
            [
                "CIF variable estándar",
                f"='{datos.title}'!B5*'{datos.title}'!B8"
            ],
            [
                "CIF fijo estándar",
                f"='{datos.title}'!B9"
            ],
            [
                "CIF estándar permitido",
                "=SUM(B2:B3)"
            ],
            [
                "CIF variable real",
                f"='{datos.title}'!B10"
            ],
            [
                "CIF fijo real",
                f"='{datos.title}'!B11"
            ],
            [
                "CIF real total",
                "=SUM(B5:B6)"
            ],
            [
                "Variación CIF",
                "=B7-B4"
            ]
        ]:
            cif.append(fila)

        # RESULTADOS

        res = libro.create_sheet(base + " Resultados")

        filas = [
            ["Concepto", "Valor"],
            [
                "Producción presupuestada",
                f"='{datos.title}'!B4"
            ],
            [
                "Producción real",
                f"='{datos.title}'!B5"
            ],
            [
                "Precio venta",
                f"='{datos.title}'!B6"
            ],
            [
                "MP estándar permitido",
                f"='{mp.title}'!H{fmp}"
            ],
            [
                "MP real",
                f"='{mp.title}'!I{fmp}"
            ],
            [
                "MO estándar permitido",
                f"='{mo.title}'!G{fmo}"
            ],
            [
                "MO real",
                f"='{mo.title}'!H{fmo}"
            ],
            [
                "CIF estándar permitido",
                f"='{cif.title}'!B4"
            ],
            [
                "CIF real",
                f"='{cif.title}'!B7"
            ],
            [
                "Costo estándar permitido",
                "=B5+B7+B9"
            ],
            [
                "Costo real",
                "=B6+B8+B10"
            ],
            [
                "Variación MP precio",
                f"='{mp.title}'!J{fmp}"
            ],
            [
                "Variación MP cantidad",
                f"='{mp.title}'!K{fmp}"
            ],
            [
                "Variación MO tarifa",
                f"='{mo.title}'!I{fmo}"
            ],
            [
                "Variación MO eficiencia",
                f"='{mo.title}'!J{fmo}"
            ],
            [
                "Variación CIF",
                f"='{cif.title}'!B8"
            ],
            [
                "Variación total",
                "=SUM(B13:B17)"
            ],
            [
                "Comprobación",
                "=B12-B11-B18"
            ],
            [
                "Costo real unitario",
                "=IF(B3>0,B12/B3,0)"
            ],
            [
                "MP estándar unitario",
                "=IF(B3>0,B5/B3,0)"
            ],
            [
                "MO estándar unitario",
                "=IF(B3>0,B7/B3,0)"
            ],
            [
                "Variable estándar unitario",
                f"=B21+B22+'{datos.title}'!B8"
            ],
            [
                "Costo estándar unitario",
                f"=B23+'{datos.title}'!B9/B2"
            ],
            [
                "Ventas reales",
                "=B3*B4"
            ],
            [
                "Utilidad real",
                "=B25-B12"
            ],
            [
                "Margen",
                "=IF(B25>0,B26/B25,0)"
            ],
            [
                "Contribución estándar",
                "=B4-B23"
            ],
            [
                "Punto equilibrio",
                (
                    f'=IF(B28>0,ROUNDUP('
                    f"'{datos.title}'!B9/B28,0),"
                    '"No calculable")'
                )
            ]
        ]

        for fila in filas:
            res.append(fila)

        res["B27"].number_format = "0.00%"

        # PRESUPUESTO

        pres = libro.create_sheet(
            base + " Presupuesto"
        )

        for fila in [
            ["Concepto", "Inicial", "Proyectado"],
            [
                "Unidades",
                f"='{datos.title}'!B4",
                f"=ROUNDUP(B2*(1+'{datos.title}'!B7),0)"
            ],
            [
                "Materia prima",
                f"=B2*'{res.title}'!B21",
                f"=C2*'{res.title}'!B21"
            ],
            [
                "Mano de obra",
                f"=B2*'{res.title}'!B22",
                f"=C2*'{res.title}'!B22"
            ],
            [
                "CIF variable",
                f"=B2*'{datos.title}'!B8",
                f"=C2*'{datos.title}'!B8"
            ],
            [
                "CIF fijo",
                f"='{datos.title}'!B9",
                f"='{datos.title}'!B9"
            ],
            [
                "Costo total",
                "=SUM(B3:B6)",
                "=SUM(C3:C6)"
            ],
            [
                "Ventas",
                f"=B2*'{res.title}'!B4",
                f"=C2*'{res.title}'!B4"
            ],
            [
                "Utilidad",
                "=B8-B7",
                "=C8-C7"
            ]
        ]:
            pres.append(fila)

        general.append([
            f"='{datos.title}'!B2",
            f"='{res.title}'!B3",
            f"='{res.title}'!B11",
            f"='{res.title}'!B12",
            f"='{res.title}'!B18",
            f"='{res.title}'!B25",
            f"='{res.title}'!B26",
            f"='{res.title}'!B27",
            f"='{res.title}'!B29",
            f"='{pres.title}'!C2"
        ])

        general.cell(
            general.max_row, 8
        ).number_format = "0.00%"

    # FORMATO

    for hoja in libro:

        hoja.freeze_panes = "A2"
        hoja.sheet_view.showGridLines = False
        hoja.row_dimensions[1].height = 34

        for celda in hoja[1]:
            celda.fill = PatternFill(
                "solid",
                fgColor="71627E"
            )
            celda.font = Font(
                color="FFFFFF",
                bold=True
            )
            celda.alignment = Alignment(
                wrap_text=True,
                vertical="center"
            )

        for col in range(
            1, hoja.max_column + 1
        ):
            hoja.column_dimensions[
                get_column_letter(col)
            ].width = 38 if col == 1 else 24

        for fila in hoja.iter_rows(min_row=2):
            for celda in fila:
                if celda.row % 2 == 0:
                    celda.fill = PatternFill(
                        "solid",
                        fgColor="F0F0F3"
                    )

    libro.calculation.fullCalcOnLoad = True
    libro.calculation.forceFullCalc = True

    salida = io.BytesIO()
    libro.save(salida)

    return salida.getvalue()


# ==========================================
# TEMA 8: INICIAR EL SISTEMA
# ==========================================

if "catalogo" not in st.session_state:
    st.session_state.catalogo = cargar()

productos = st.session_state.catalogo["productos"]

if (
    "seleccion" not in st.session_state
    or st.session_state.seleccion not in productos
):
    st.session_state.seleccion = next(
        iter(productos)
    )

# ==========================================
# TEMA 9: MENU DE PRODUCTOS
# ==========================================

with st.sidebar:

    st.header("Catálogo de productos")

    nombres = list(productos)

    elegido = st.selectbox(
        "Seleccionar producto",
        nombres,
        index=nombres.index(
            st.session_state.seleccion
        )
    )

    if elegido != st.session_state.seleccion:
        st.session_state.seleccion = elegido
        st.rerun()

    st.divider()

    with st.form(
        "nuevo_producto",
        clear_on_submit=True
    ):

        nuevo = st.text_input(
            "Nombre del nuevo producto"
        )

        crear = st.form_submit_button(
            "Agregar producto"
        )

    if crear:

        nuevo = nuevo.strip()

        if not nuevo:
            st.error("Escribe un nombre.")

        elif nuevo in productos:
            st.warning("El producto ya existe.")

        else:
            productos[nuevo] = crear_producto(nuevo)
            st.session_state.seleccion = nuevo
            guardar_automaticamente()
            st.rerun()

    st.divider()

    # ELIMINAR SIN EXPANDER

    st.markdown("**Administrar productos**")

    confirmar = st.checkbox(
        "Confirmar eliminación del producto",
        key="confirmar_eliminacion"
    )

    if st.button(
        "Eliminar producto seleccionado",
        disabled=(
            not confirmar or len(productos) == 1
        ),
        use_container_width=True
    ):

        del productos[st.session_state.seleccion]

        st.session_state.seleccion = next(
            iter(productos)
        )

        guardar_automaticamente()
        st.rerun()

    st.divider()

    st.caption(
        "Guardado automático activado."
    )

# ==========================================
# TEMA 10: PANTALLA PRINCIPAL
# ==========================================

nombre = st.session_state.seleccion
p = productos[nombre]

st.title("Dulce Momento, Gran Sentimiento")

st.caption(
    "Sistema de Gestión de Costos Estándar "
    "y Control de Producción"
)

st.divider()

st.header(nombre)

p["descripcion"] = st.text_area(
    "Descripción del producto",
    value=p.get("descripcion", ""),
    height=100,
    key=f"descripcion_{nombre}"
)

# ==========================================
# TEMA 11: CONFIGURACION DE PRODUCCION
# ==========================================

st.subheader("Configuración de producción")

izquierda, derecha = st.columns(2, gap="large")

with izquierda:

    st.markdown("**Producción presupuestada**")

    p["cantidad"] = st.number_input(
        "Producción presupuestada",
        min_value=1,
        value=int(p["cantidad"]),
        label_visibility="collapsed",
        key=f"cantidad_{nombre}"
    )

    st.caption("Unidades que se planea fabricar.")

with derecha:

    st.markdown("**Producción real terminada**")

    p["cantidad_real"] = st.number_input(
        "Producción real terminada",
        min_value=1,
        value=int(p["cantidad_real"]),
        label_visibility="collapsed",
        key=f"real_{nombre}"
    )

    st.caption("Unidades realmente fabricadas.")

izquierda, derecha = st.columns(2, gap="large")

with izquierda:

    st.markdown("**Precio de venta unitario (Q)**")

    p["precio_venta"] = st.number_input(
        "Precio de venta",
        min_value=0.0,
        value=float(p["precio_venta"]),
        label_visibility="collapsed",
        key=f"precio_{nombre}"
    )

    st.caption("Precio por producto en quetzales.")

with derecha:

    st.markdown("**Crecimiento proyectado (%)**")

    p["crecimiento"] = st.number_input(
        "Crecimiento proyectado",
        min_value=0.0,
        value=float(p["crecimiento"]),
        label_visibility="collapsed",
        key=f"crecimiento_{nombre}"
    )

    st.caption("Porcentaje para el presupuesto.")

# ==========================================
# TEMA 12: INSTRUCCIONES SIN EXPANDER
# ==========================================

mostrar_ayuda = st.checkbox(
    "Mostrar instrucciones para ingresar los datos",
    value=False,
    key="mostrar_ayuda"
)

if mostrar_ayuda:

    st.markdown("""
<div class="ayuda">

<p><strong>Producción presupuestada:</strong>
unidades que se planea fabricar.</p>

<p><strong>Producción real:</strong>
unidades realmente terminadas.</p>

<p><strong>Materia prima estándar:</strong>
ingredientes por unidad y precio estándar.</p>

<p><strong>Materia prima real:</strong>
cantidades totales consumidas y precios reales.</p>

<p><strong>Mano de obra estándar:</strong>
horas por producto y tarifa estándar.</p>

<p><strong>Mano de obra real:</strong>
horas totales trabajadas y tarifa real.</p>

<p><strong>CIF:</strong>
costos indirectos variables y fijos.</p>

<p><strong>Ejemplo:</strong>
500 g de fresas por pastel multiplicados
por 20 pasteles equivalen a 10,000 g.</p>

</div>
""", unsafe_allow_html=True)

# ==========================================
# TEMA 13: MATERIA PRIMA
# ==========================================

st.header("1. Materia prima directa")

st.write(
    "Registra los ingredientes, cantidades "
    "estándar, precios y consumos reales."
)

columnas_mp = [
    "Ingrediente",
    "Unidad",
    "Cantidad estándar",
    "Precio estándar Q",
    "Cantidad real",
    "Precio real Q",
    "Proveedor"
]

tabla_mp = st.data_editor(
    pd.DataFrame(
        p["ingredientes"],
        columns=columnas_mp
    ),
    num_rows="dynamic",
    hide_index=True,
    use_container_width=True,
    key=f"mp_{nombre}",
    column_config={
        "Precio estándar Q":
            st.column_config.NumberColumn(
                min_value=0.0,
                format="%.4f"
            ),
        "Precio real Q":
            st.column_config.NumberColumn(
                min_value=0.0,
                format="%.4f"
            )
    }
)

p["ingredientes"] = tabla_mp.where(
    pd.notna(tabla_mp),
    None
).to_dict("records")

st.caption(
    "Ejemplo: 1,000 g de harina por Q10 "
    "equivalen a Q0.01 por gramo."
)

# ==========================================
# TEMA 14: MANO DE OBRA
# ==========================================

st.header("2. Mano de obra directa")

columnas_mo = [
    "Actividad",
    "Horas estándar por unidad",
    "Tarifa estándar Q",
    "Horas reales totales",
    "Tarifa real Q"
]

tabla_mo = st.data_editor(
    pd.DataFrame(
        p["mano_obra"],
        columns=columnas_mo
    ),
    num_rows="dynamic",
    hide_index=True,
    use_container_width=True,
    key=f"mo_{nombre}"
)

p["mano_obra"] = tabla_mo.where(
    pd.notna(tabla_mo),
    None
).to_dict("records")

# ==========================================
# TEMA 15: COSTOS INDIRECTOS
# ==========================================

st.header("3. Costos indirectos de fabricación")

cif = p["cif"]

izquierda, derecha = st.columns(2, gap="large")

with izquierda:

    cif["Variable estándar por unidad"] = (
        st.number_input(
            "CIF variable estándar por unidad (Q)",
            min_value=0.0,
            value=float(
                cif["Variable estándar por unidad"]
            ),
            key=f"cve_{nombre}"
        )
    )

    cif["Variable real total"] = st.number_input(
        "CIF variable real total (Q)",
        min_value=0.0,
        value=float(cif["Variable real total"]),
        key=f"cvr_{nombre}"
    )

with derecha:

    cif["Fijo estándar del período"] = (
        st.number_input(
            "CIF fijo estándar del período (Q)",
            min_value=0.0,
            value=float(
                cif["Fijo estándar del período"]
            ),
            key=f"cfe_{nombre}"
        )
    )

    cif["Fijo real total"] = st.number_input(
        "CIF fijo real total (Q)",
        min_value=0.0,
        value=float(cif["Fijo real total"]),
        key=f"cfr_{nombre}"
    )

# ==========================================
# TEMA 16: RESPALDO AUTOMATICO
# ==========================================

guardar_automaticamente()

# ==========================================
# TEMA 17: RESULTADOS CONTABLES
# ==========================================

r = calcular(p)

st.header("4. Resultados contables")

a, b, c, d = st.columns(4)

a.metric(
    "Costo estándar unitario",
    dinero(r["costo_est_unit"])
)

b.metric(
    "Costo real unitario",
    dinero(r["costo_real_unit"])
)

c.metric(
    "Costo estándar permitido",
    dinero(r["total_est"])
)

d.metric(
    "Costo real total",
    dinero(r["total_real"])
)

resumen = pd.DataFrame({
    "Concepto": [
        "Materia prima estándar permitida",
        "Materia prima real",
        "Mano de obra estándar permitida",
        "Mano de obra real",
        "CIF estándar permitido",
        "CIF real",
        "Costo estándar permitido",
        "Costo real total"
    ],
    "Monto Q": [
        r["mp_est"],
        r["mp_real"],
        r["mo_est"],
        r["mo_real"],
        r["cif_est"],
        r["cif_real"],
        r["total_est"],
        r["total_real"]
    ]
})

st.dataframe(
    resumen,
    hide_index=True,
    use_container_width=True
)

# ==========================================
# TEMA 18: VARIACIONES
# ==========================================

st.header("5. Análisis de variaciones")

variaciones = pd.DataFrame([
    ["Materia prima - Precio", r["var_mp_precio"]],
    ["Materia prima - Cantidad", r["var_mp_cantidad"]],
    ["Mano de obra - Tarifa", r["var_mo_tarifa"]],
    ["Mano de obra - Eficiencia", r["var_mo_eficiencia"]],
    ["CIF - Variación global", r["var_cif"]],
    ["VARIACIÓN TOTAL", r["variacion_total"]]
], columns=["Concepto", "Variación Q"])

variaciones["Clasificación"] = (
    variaciones["Variación Q"].apply(clasificar)
)

st.dataframe(
    variaciones,
    hide_index=True,
    use_container_width=True
)

if r["variacion_total"] > 0:
    st.warning(
        "Variación desfavorable: "
        "el costo real supera al estándar."
    )

elif r["variacion_total"] < 0:
    st.success(
        "Variación favorable: "
        "el costo real es menor al estándar."
    )

else:
    st.info("No existe variación total.")

# ==========================================
# TEMA 19: RENTABILIDAD
# ==========================================

st.header("6. Rentabilidad")

a, b, c = st.columns(3)

a.metric("Ventas", dinero(r["ventas"]))
b.metric("Utilidad", dinero(r["utilidad"]))

c.metric(
    "Margen sobre ventas",
    (
        f'{r["margen"]:.2f}%'
        if r["margen"] is not None
        else "No calculable"
    )
)

# ==========================================
# TEMA 20: PUNTO DE EQUILIBRIO
# ==========================================

st.header("7. Punto de equilibrio")

if r["equilibrio"] is not None:
    st.success(
        f'Se necesitan vender {r["equilibrio"]} '
        "unidades para cubrir los costos estándar."
    )
else:
    st.warning(
        "El precio de venta debe superar "
        "el costo variable estándar unitario."
    )

# ==========================================
# TEMA 21: PRESUPUESTO
# ==========================================

st.header("8. Presupuesto de producción")

costo_inicial = (
    r["q"] * r["variable_unit"] + r["cf_est"]
)

ventas_iniciales = r["q"] * r["precio"]

ventas_proyectadas = (
    r["q_proyectada"] * r["precio"]
)

presupuesto = pd.DataFrame({
    "Concepto": [
        "Unidades",
        "Costo estándar total Q",
        "Ventas Q",
        "Utilidad estándar Q"
    ],
    "Inicial": [
        r["q"],
        costo_inicial,
        ventas_iniciales,
        ventas_iniciales - costo_inicial
    ],
    "Proyectado": [
        r["q_proyectada"],
        r["costo_proyectado"],
        ventas_proyectadas,
        ventas_proyectadas - r["costo_proyectado"]
    ]
})

st.dataframe(
    presupuesto,
    hide_index=True,
    use_container_width=True
)

st.subheader("Comparación de costos")

st.bar_chart(
    pd.DataFrame({
        "Costo Q": {
            "Estándar permitido": r["total_est"],
            "Real": r["total_real"]
        }
    })
)

# ==========================================
# TEMA 22: EXPORTACION
# ==========================================

st.header("9. Exportar información")

st.download_button(
    "Descargar Excel de todos los productos",
    data=crear_excel(st.session_state.catalogo),
    file_name="Dulce_Momento_Costos_Estandar.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)

st.caption(
    "Los datos de los productos se guardan "
    "automáticamente en el sistema."
)

st.caption(
    "Dulce Momento, Gran Sentimiento | Contabilidad II"
)
