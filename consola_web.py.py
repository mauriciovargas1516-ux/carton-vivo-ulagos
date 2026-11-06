# -*- coding: utf-8 -*-
"""
consola_web.py — CONSOLA WEB INTERACTIVA (BLINDADA PARA STREAMLIT CLOUD / LINUX)
Conecta el motor financiero y generador de Excels/Word (Script 2) con el integrador
de Árboles, Croquis, CANVAS y Marco Lógico (Script 3).
"""

import os
import sys
import types
import copy
import shutil
import zipfile
import tempfile
import subprocess
import importlib.util

# Forzar backend sin pantalla para que Matplotlib funcione en cualquier servidor web Linux
import matplotlib
matplotlib.use("Agg")

import streamlit as st
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==============================================================================
# 1. COMPATIBILIDAD TOTAL CON SERVIDORES LINUX EN LA NUBE (SIN TKINTER)
# ==============================================================================

def _asegurar_tkinter_en_nube():
    """Emula todos los submódulos de tkinter en servidores Linux para que ningún script falle al importar."""
    try:
        import tkinter  # noqa: F401
        import tkinter.font  # noqa: F401
    except ImportError:
        tk_mock = types.ModuleType("tkinter")
        tk_mock.Tk = object
        tk_mock.Toplevel = object
        tk_mock.Frame = object
        tk_mock.Label = object
        tk_mock.Button = object
        tk_mock.Canvas = object
        tk_mock.StringVar = lambda *a, **k: None
        tk_mock.DoubleVar = lambda *a, **k: None
        tk_mock.IntVar = lambda *a, **k: None
        tk_mock.ttk = types.ModuleType("tkinter.ttk")
        tk_mock.messagebox = types.ModuleType("tkinter.messagebox")
        tk_mock.filedialog = types.ModuleType("tkinter.filedialog")
        tk_mock.font = types.ModuleType("tkinter.font")
        tk_mock.constants = types.ModuleType("tkinter.constants")
        sys.modules["tkinter"] = tk_mock
        sys.modules["tkinter.ttk"] = tk_mock.ttk
        sys.modules["tkinter.messagebox"] = tk_mock.messagebox
        sys.modules["tkinter.filedialog"] = tk_mock.filedialog
        sys.modules["tkinter.font"] = tk_mock.font
        sys.modules["tkinter.constants"] = tk_mock.constants


_asegurar_tkinter_en_nube()


def cargar_modulo_por_ruta(nombre_interno, nombre_archivo):
    carpeta = os.path.dirname(os.path.abspath(__file__))
    ruta = os.path.join(carpeta, nombre_archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró '{nombre_archivo}' en: {carpeta}")
    spec = importlib.util.spec_from_file_location(nombre_interno, ruta)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[nombre_interno] = mod
    spec.loader.exec_module(mod)
    return mod


# Cargamos únicamente el Script 2 (Motor + 3 Excels + Word base) y el Script 3 (Árboles + CANVAS + MML)
mod2 = cargar_modulo_por_ruta("script2_corrector", "2_corrector_formato_y_excels.py")
mod3 = cargar_modulo_por_ruta("script3_josue", "3_integrador_arboles_y_pauta_josue.py")


# ==============================================================================
# 2. REPARADOR DEL ÍNDICE (PÁGINA 2) Y PIPELINE UNIFICADO (SCRIPTS 2 + 3)
# ==============================================================================

def _es_linea_toc(p):
    return len(p._p.findall(qn("w:hyperlink"))) > 0 or len(p._p.findall(qn("w:fldSimple"))) > 0


def _crear_entrada_toc_con_puntos(doc, texto_sec, anchor_name, pag_str="20", nivel=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3.5 if nivel == 1 else 1.5)
    p.paragraph_format.space_after = Pt(3.5 if nivel == 1 else 1.5)
    p.paragraph_format.left_indent = Inches(0.0 if nivel == 1 else 0.25)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(6.65), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor_name)
    hyperlink.set(qn("w:history"), "1")

    def _run_xml(txt, bold=True, pt=10):
        r_el = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), "Arial")
        rFonts.set(qn("w:hAnsi"), "Arial")
        rFonts.set(qn("w:cs"), "Arial")
        rPr.append(rFonts)
        rPr.append(OxmlElement("w:noProof"))
        if bold:
            rPr.append(OxmlElement("w:b"))
            rPr.append(OxmlElement("w:bCs"))
        c_el = OxmlElement("w:color")
        c_el.set(qn("w:val"), "000000")
        rPr.append(c_el)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(pt * 2)))
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(pt * 2)))
        rPr.append(sz)
        rPr.append(szCs)
        r_el.append(rPr)
        t_el = OxmlElement("w:t")
        t_el.text = txt
        r_el.append(t_el)
        return r_el

    hyperlink.append(_run_xml(texto_sec, bold=(nivel == 1), pt=10 if nivel == 1 else 9.5))

    r_tab = OxmlElement("w:r")
    rPr_t = OxmlElement("w:rPr")
    rFonts_t = OxmlElement("w:rFonts")
    rFonts_t.set(qn("w:ascii"), "Arial")
    rFonts_t.set(qn("w:hAnsi"), "Arial")
    rPr_t.append(rFonts_t)
    r_tab.append(rPr_t)
    r_tab.append(OxmlElement("w:tab"))
    hyperlink.append(r_tab)

    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" PAGEREF {anchor_name} \\h ")
    fld.append(_run_xml(str(pag_str), bold=True, pt=10 if nivel == 1 else 9.5))
    hyperlink.append(fld)

    p._p.append(hyperlink)
    return p


def _reparar_toc_y_ubicar_matriz_en_cuerpo(ruta_word):
    doc = mod2.Document(ruta_word)
    body = doc.element.body

    toc_resumen_el = None
    toc_sec1_el = None
    for p in doc.paragraphs:
        if _es_linea_toc(p):
            t_up = p.text.strip().upper()
            if t_up.startswith("RESUMEN EJECUTIVO") and toc_resumen_el is None:
                toc_resumen_el = p._p
            elif t_up.startswith("1. EL PROBLEMA") and toc_sec1_el is None:
                toc_sec1_el = p._p
                break

    p_sec1_real = None
    for p in doc.paragraphs:
        if not _es_linea_toc(p) and p.text.strip().upper().startswith("1. EL PROBLEMA EN OSORNO"):
            p_sec1_real = p
            break

    if toc_resumen_el is not None and toc_sec1_el is not None:
        hijos = list(body)
        idx_ini = hijos.index(toc_resumen_el)
        idx_fin = hijos.index(toc_sec1_el)
        intrusos = hijos[idx_ini + 1:idx_fin]
        for el in intrusos:
            body.remove(el)
            if p_sec1_real is not None:
                p_sec1_real._p.addprevious(el)

    for p in list(doc.paragraphs):
        t = p.text.strip().upper()
        if not _es_linea_toc(p) and (
            t.startswith("10. MODELO DE GESTIÓN ESTRATÉGICA")
            or t.startswith("11. MODELO DE SEGUIMIENTO")
            or t.startswith("12. CONTROL DE CONSISTENCIA, MEJORAS Y REFERENCIAS")
        ):
            if list(body).index(p._p) < 60:
                p._element.getparent().remove(p._element)

    p_toc_anx1 = None
    for p in doc.paragraphs:
        if _es_linea_toc(p) and p.text.strip().upper().startswith("ANEXO 1:"):
            p_toc_anx1 = p
            break

    texto_toc_actual = "\n".join(p.text for p in doc.paragraphs[:60])
    if p_toc_anx1 is not None and "10. MODELO DE GESTIÓN ESTRATÉGICA" not in texto_toc_actual:
        for txt_cap, anchor, pag_def in [
            ("10. MODELO DE GESTIÓN ESTRATÉGICA - BUSINESS MODEL CANVAS (PUNTO 9 PAUTA)", "_Toc_Canvas", "20"),
            ("11. MODELO DE SEGUIMIENTO - MATRIZ DE MARCO LÓGICO (PUNTO 10 PAUTA)", "_Toc_MML", "21"),
            ("12. CONTROL DE CONSISTENCIA, MEJORAS Y REFERENCIAS OFICIALES SNI 2026", "_Toc_Consistencia", "21"),
        ]:
            p_nuevo = _crear_entrada_toc_con_puntos(doc, txt_cap, anchor, pag_def, nivel=1)
            p_toc_anx1._p.addprevious(p_nuevo._p)

    doc.save(ruta_word)


def calcular_escenario_sensibilidad(cfg_base, factor_ingresos=1.0, factor_costos=1.0, crecimiento_override=None):
    cfg_esc = copy.deepcopy(cfg_base)
    if crecimiento_override is not None:
        cfg_esc["TASA_CRECIMIENTO"] = float(crecimiento_override)

    cfg_esc["P1_PRECIO_BRUTO"] = float(cfg_base["P1_PRECIO_BRUTO"]) * factor_ingresos
    cfg_esc["P2_PRECIO_BRUTO"] = float(cfg_base["P2_PRECIO_BRUTO"]) * factor_ingresos

    for cv in cfg_esc["COSTOS_VAR_LISTA"]:
        cv["costo"] = float(cv["costo"]) * factor_costos
    for cf in cfg_esc["COSTOS_FIJOS_LISTA"]:
        cf["mensual"] = float(cf["mensual"]) * factor_costos
    for p in cfg_esc["PERSONAL_LISTA"]:
        p["sueldo_bruto_mensual"] = float(p["sueldo_bruto_mensual"]) * factor_costos

    res = mod2.ejecutar_motor_financiero(cfg_esc)
    m = res["metricas"]
    return {
        "van_p": m["van_privado"],
        "tir_p": m["tir_privada"],
        "van_s": m["van_social"],
        "tir_s": m["tir_social"],
    }


def generar_entrega_completa_en_servidor(cfg_vivo, directorio_destino):
    for k, v in cfg_vivo.items():
        if k not in ["INVERSIONES_LISTA", "COSTOS_VAR_LISTA", "COSTOS_FIJOS_LISTA", "PERSONAL_LISTA"]:
            mod2.DEFAULTS[k] = copy.deepcopy(v)

    mod2.INVERSIONES_DEFAULT = copy.deepcopy(cfg_vivo["INVERSIONES_LISTA"])
    mod2.COSTOS_VAR_DEFAULT = copy.deepcopy(cfg_vivo["COSTOS_VAR_LISTA"])
    mod2.COSTOS_FIJOS_DEFAULT = copy.deepcopy(cfg_vivo["COSTOS_FIJOS_LISTA"])
    mod2.PERSONAL_DEFAULT = copy.deepcopy(cfg_vivo["PERSONAL_LISTA"])

    resultados_vivos = mod2.ejecutar_motor_financiero(cfg_vivo)
    apellidos = cfg_vivo.get("APELLIDOS_ENTREGA", "Vargas,Muñoz,Jofre").strip()
    carpeta_entrega = os.path.join(directorio_destino, apellidos)
    os.makedirs(carpeta_entrega, exist_ok=True)

    # 1. Generar los 3 Excels
    r_exc1, r_exc2, r_exc3 = mod2.exportar_tres_excels_limpios(resultados_vivos, carpeta_entrega)

    # 2. Generar el Word base (Script 2)
    ruta_word = os.path.join(carpeta_entrega, f"{apellidos}.docx")
    mod2.generar_informe_word_limpio(resultados_vivos, ruta_word, generar_pdf=False)

    for img_temp in ["grafico_flujos_cartonvivo.png", "grafico_carton_ramirez.png"]:
        p_img = os.path.join(carpeta_entrega, img_temp)
        if os.path.exists(p_img):
            try:
                os.remove(p_img)
            except Exception:
                pass

    # 3. Inyectar Árboles, Croquis, CANVAS y Marco Lógico (Script 3) y reparar Índice
    mod3.integrar_trabajo_josue_en_word(ruta_word, carpeta_entrega)
    _reparar_toc_y_ubicar_matriz_en_cuerpo(ruta_word)

    # 4. Actualizar tablas de resumen y sensibilidad con las cifras vivas
    m = resultados_vivos["metricas"]
    esc_pes = calcular_escenario_sensibilidad(cfg_vivo, factor_ingresos=0.85, factor_costos=1.00)
    esc_cos = calcular_escenario_sensibilidad(cfg_vivo, factor_ingresos=1.00, factor_costos=1.10)
    esc_bas = calcular_escenario_sensibilidad(cfg_vivo, factor_ingresos=1.00, factor_costos=1.00)
    esc_cer = calcular_escenario_sensibilidad(cfg_vivo, factor_ingresos=1.00, factor_costos=1.00, crecimiento_override=0.0)

    doc_final = mod2.Document(ruta_word)
    for tbl in doc_final.tables:
        c0 = tbl.rows[0].cells[0].text.strip().upper()
        if c0 == "INDICADOR" and len(tbl.rows) >= 5:
            tbl.rows[1].cells[1].text = f"{mod2.fmt_dec_cl(cfg_vivo['TASA_DESCUENTO_PRIVADA']*100, 1)}%"
            tbl.rows[1].cells[2].text = f"{mod2.fmt_dec_cl(cfg_vivo['TASA_DESCUENTO_SOCIAL']*100, 1)}%"
            tbl.rows[2].cells[1].text = mod2.fmt_clp(m["van_privado"])
            tbl.rows[2].cells[2].text = mod2.fmt_clp(m["van_social"])
            tbl.rows[3].cells[1].text = f"{mod2.fmt_dec_cl(m['tir_privada'], 2)}%"
            tbl.rows[3].cells[2].text = f"{mod2.fmt_dec_cl(m['tir_social'], 2)}%"
            tbl.rows[4].cells[1].text = f"{mod2.fmt_dec_cl(m['payback_privado'], 2)} años"
            tbl.rows[4].cells[2].text = f"{mod2.fmt_dec_cl(m['payback_social'], 2)} años"
        elif c0 == "ESCENARIO EVALUADO" and len(tbl.rows) >= 5:
            for idx_fila, esc in enumerate([esc_pes, esc_cos, esc_bas, esc_cer], start=1):
                tbl.rows[idx_fila].cells[1].text = mod2.fmt_clp(esc["van_p"])
                tbl.rows[idx_fila].cells[2].text = f"{mod2.fmt_dec_cl(esc['tir_p'], 2)}%"
                tbl.rows[idx_fila].cells[3].text = mod2.fmt_clp(esc["van_s"])
                tbl.rows[idx_fila].cells[4].text = f"{mod2.fmt_dec_cl(esc['tir_s'], 2)}%"

    doc_final.save(ruta_word)

    # 5. Exportar PDF (en Windows usa Word; en Linux/Cloud usa LibreOffice)
    ruta_pdf, ruta_zip, pdf_ok = mod3.actualizar_pdf_y_zip_final(ruta_word)
    if not pdf_ok and shutil.which("libreoffice"):
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", carpeta_entrega, ruta_word],
                check=True, timeout=60
            )
            with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for arch in os.listdir(carpeta_entrega):
                    r_full = os.path.join(carpeta_entrega, arch)
                    if os.path.isfile(r_full):
                        zf.write(r_full, arcname=os.path.join(apellidos, arch))
        except Exception:
            pass

    return {
        "ruta_zip": ruta_zip,
        "ruta_word": ruta_word,
        "ruta_pdf": ruta_pdf,
        "excels": [r_exc1, r_exc2, r_exc3],
    }


# ==============================================================================
# 3. INTERFAZ VISUAL EN EL NAVEGADOR (3 COLUMNAS ALINEADAS IGUAL A TU CONSOLA)
# ==============================================================================

st.set_page_config(page_title="Consola Interactiva ULagos - Cartón Vivo Osorno", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1220px; }
    .sec-header {
        background-color: #DBEAFE;
        color: #1E3A8A;
        padding: 7px 14px;
        font-weight: 700;
        font-size: 14px;
        border-radius: 5px;
        margin-top: 14px;
        margin-bottom: 8px;
        border-left: 4px solid #1D4ED8;
    }
    .lbl-col { font-weight: 700; color: #0F172A; font-size: 13.5px; padding-top: 6px; }
    .exp-col { color: #475569; font-size: 12.5px; padding-top: 6px; line-height: 1.35; }
    div[data-testid="stTextInput"] { margin-bottom: -10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

if "cfg" not in st.session_state:
    c_init = copy.deepcopy(mod2.DEFAULTS)
    c_init["INVERSIONES_LISTA"] = copy.deepcopy(mod2.INVERSIONES_DEFAULT)
    c_init["COSTOS_VAR_LISTA"] = copy.deepcopy(mod2.COSTOS_VAR_DEFAULT)
    c_init["COSTOS_FIJOS_LISTA"] = copy.deepcopy(mod2.COSTOS_FIJOS_DEFAULT)
    c_init["PERSONAL_LISTA"] = copy.deepcopy(mod2.PERSONAL_DEFAULT)
    st.session_state.cfg = c_init

cfg = st.session_state.cfg

col_p1, col_p2, col_p3 = st.columns([1.2, 1.2, 1.6])
if col_p1.button("🔶 Cargar Precios Sombra SNI 2026 (TSD 5,5%)", use_container_width=True):
    cfg["TASA_DESCUENTO_SOCIAL"] = 0.055
    cfg["FC_MOC"], cfg["FC_MOSC"], cfg["FC_MONC"], cfg["FC_DIVISA"] = 0.97, 0.95, 0.91, 1.00
    cfg["OTROS_BEN_SOC_ANUAL"] = 0.0
    st.rerun()

if col_p2.button("🔷 Cargar Caso Base Original (TSD 6,0%)", use_container_width=True):
    cfg["TASA_DESCUENTO_SOCIAL"] = 0.06
    cfg["FC_MOC"], cfg["FC_MOSC"], cfg["FC_MONC"], cfg["FC_DIVISA"] = 1.00, 0.85, 0.68, 1.01
    cfg["OTROS_BEN_SOC_ANUAL"] = 2500000.0
    st.rerun()

banner_placeholder = st.empty()


def fila_3_columnas(clave_unica, etiqueta, valor_actual, explicacion, es_texto=False, es_entero=False):
    c_lbl, c_inp, c_exp = st.columns([2.3, 1.6, 4.1])
    with c_lbl:
        st.markdown(f"<div class='lbl-col'>{etiqueta}</div>", unsafe_allow_html=True)
    with c_inp:
        nuevo_txt = st.text_input(
            etiqueta, value=str(valor_actual), key=f"in_{clave_unica}", label_visibility="collapsed"
        )
    with c_exp:
        st.markdown(f"<div class='exp-col'>{explicacion}</div>", unsafe_allow_html=True)

    if es_texto:
        return nuevo_txt.strip()
    try:
        val_num = float(nuevo_txt.strip().replace(",", "."))
        return int(val_num) if es_entero else val_num
    except Exception:
        return valor_actual


# SECCIÓN 1
st.markdown("<div class='sec-header'>1. PARÁMETROS GENERALES Y TASAS DE DESCUENTO</div>", unsafe_allow_html=True)
cfg["NOMBRE_PROYECTO"] = fila_3_columnas("NOM", "Nombre del Proyecto:", cfg["NOMBRE_PROYECTO"], "(Nombre oficial que saldrá en la portada de la ULagos y en las conclusiones automáticas del Word)", es_texto=True)
cfg["APELLIDOS_ENTREGA"] = fila_3_columnas("APE", "Apellidos Carpeta y ZIP:", cfg["APELLIDOS_ENTREGA"], "(Nombre exacto para la carpeta, el Word, el PDF y el archivo comprimido ZIP de entrega)", es_texto=True)
cfg["INTEGRANTES"] = fila_3_columnas("INT", "Integrantes del Grupo:", cfg["INTEGRANTES"], "(Nombres completos en mayúsculas separados por coma para el bloque PROPUESTO POR de la portada)", es_texto=True)
cfg["DOCENTE"] = fila_3_columnas("DOC", "Nombre Docente:", cfg["DOCENTE"], "(Nombre de la profesora guía del módulo Evaluación Integrada de Proyectos)", es_texto=True)
cfg["HORIZONTE_ANIOS"] = fila_3_columnas("HOR", "Horizonte de Evaluación:", cfg["HORIZONTE_ANIOS"], "(Cantidad de años a proyectar en las tablas. La pauta exige evaluar a 10 años -> deja 10)", es_entero=True)
cfg["IVA"] = fila_3_columnas("IVA", "Tasa de IVA (Impuesto al Valor Agregado):", cfg["IVA"], "(En Chile el IVA legal es 19% -> escribe 0.19. El programa lo usa para descontar el IVA automáticamente)")
cfg["IMPUESTO_RENTA"] = fila_3_columnas("IMP", "Impuesto a la Renta (1ª Categoría SII):", cfg["IMPUESTO_RENTA"], "(Impuesto sobre las ganancias de la empresa: usa 0.25 para Pyme [25%]. Solo se cobra en el Flujo Privado)")
cfg["TASA_DESCUENTO_PRIVADA"] = fila_3_columnas("TDP", "Tasa de Descuento Privada (Costo de Capital):", cfg["TASA_DESCUENTO_PRIVADA"], "(Rentabilidad mínima anual que exige el dueño/inversionista para calcular el VAN Privado: ej. 0.12 = 12%)")
cfg["TASA_DESCUENTO_SOCIAL"] = fila_3_columnas("TDS", "Tasa Social de Descuento (MIDESO / SNI):", cfg["TASA_DESCUENTO_SOCIAL"], "(Tasa exigida por el Estado para calcular el VANS: usa 0.06 [6%] o 0.055 [5,5% oficial vigente SNI 2026])")
cfg["DIAS_DESFASE_CT"] = fila_3_columnas("DCT", "Días de Desfase para Capital de Trabajo:", cfg["DIAS_DESFASE_CT"], "(Días que la planta debe operar pagando sueldos e insumos antes de cobrar las primeras facturas. Ej: 45 días)", es_entero=True)

# SECCIÓN 2
st.markdown("<div class='sec-header'>2. FACTORES DE CORRECCIÓN SOCIAL (PRECIOS SOMBRA SNI / MIDESO - PUNTO 8)</div>", unsafe_allow_html=True)
cfg["FC_MOC"] = fila_3_columnas("MOC", "Factor Mano de Obra Calificada (MOC):", cfg["FC_MOC"], "(Para profesionales/jefes. MIDESO usa 1.00 en caso base o 0.97 según Precios Sociales SNI 2026)")
cfg["FC_MOSC"] = fila_3_columnas("MOSC", "Factor Mano de Obra Semicalificada (MOSC):", cfg["FC_MOSC"], "(Para técnicos/operadores de máquinas. MIDESO usa 0.85 en caso base o 0.95 según Precios Sociales SNI 2026)")
cfg["FC_MONC"] = fila_3_columnas("MONC", "Factor Mano de Obra No Calificada (MONC):", cfg["FC_MONC"], "(Para recolectores/ayudantes. MIDESO usa 0.68 en caso base o 0.91 según Precios Sociales SNI 2026)")
cfg["FC_DIVISA"] = fila_3_columnas("DIV", "Factor Social de la Divisa (Dólar/Importados):", cfg["FC_DIVISA"], "(Premio social a la divisa de MIDESO = 1.01 [o 1.00 en SNI 2026]. Ajusta el valor de máquinas importadas)")
cfg["ARANCEL_IMP"] = fila_3_columnas("ARA", "Arancel Aduanero de Importación:", cfg["ARANCEL_IMP"], "(Impuesto de aduana en Chile = 6% -> escribe 0.06. En el Flujo Social el programa lo quita porque es impuesto interno)")

# SECCIÓN 3
st.markdown("<div class='sec-header'>3. DEMANDA, PRODUCTOS E INGRESOS (PUNTOS 4, 5.a y 5.d)</div>", unsafe_allow_html=True)
cfg["RESIDUO_TON_ANIO_1"] = fila_3_columnas("TON", "Cartón Recolectado el Año 1 (Toneladas):", cfg["RESIDUO_TON_ANIO_1"], "(Toneladas de cartón limpio captadas en Osorno el primer año. Ej: 15.0 ton/año = 50 kg/día)")
cfg["TASA_CRECIMIENTO"] = fila_3_columnas("CRE", "Tasa de Crecimiento Anual (Años 2 al 10):", cfg["TASA_CRECIMIENTO"], "(Ritmo al que crecen cada año las ventas y la recolección de cartón. Ej: escribe 0.05 para subir 5% anual)")
cfg["P1_NOMBRE"] = fila_3_columnas("P1N", "Nombre Producto Línea 1:", cfg["P1_NOMBRE"], "(Nombre comercial de tu primera línea de productos para colegios y jardines infantiles)", es_texto=True)
cfg["P1_UNIDADES"] = fila_3_columnas("P1U", "Unidades Vendidas Año 1 - Línea 1:", cfg["P1_UNIDADES"], "(Cantidad de unidades de la Línea 1 vendidas en todo el primer año. Ej: 3000 unidades/año)", es_entero=True)
cfg["P1_PRECIO_BRUTO"] = fila_3_columnas("P1P", "Precio de Venta CON IVA - Línea 1 ($):", cfg["P1_PRECIO_BRUTO"], "(Precio al que vendes cada unidad con IVA incluido. Ej: 16000 pesos. El programa lo divide por 1.19 solo)")
cfg["P1_KG_CARTON"] = fila_3_columnas("P1K", "Kg de Cartón por Unidad - Línea 1:", cfg["P1_KG_CARTON"], "(Kilos de cartón que se necesitan para fabricar 1 unidad de la Línea 1. Ej: 2.5 kg por set)")
cfg["P2_NOMBRE"] = fila_3_columnas("P2N", "Nombre Producto Línea 2:", cfg["P2_NOMBRE"], "(Nombre comercial de tu segunda línea enfocada en centros de adulto mayor y estimulación cognitiva)", es_texto=True)
cfg["P2_UNIDADES"] = fila_3_columnas("P2U", "Unidades Vendidas Año 1 - Línea 2:", cfg["P2_UNIDADES"], "(Cantidad de unidades de la Línea 2 vendidas en todo el primer año. Ej: 1500 unidades/año)", es_entero=True)
cfg["P2_PRECIO_BRUTO"] = fila_3_columnas("P2P", "Precio de Venta CON IVA - Línea 2 ($):", cfg["P2_PRECIO_BRUTO"], "(Precio al que vendes cada unidad con IVA incluido. Ej: 19000 pesos. El programa le quita el 19% de IVA solo)")
cfg["P2_KG_CARTON"] = fila_3_columnas("P2K", "Kg de Cartón por Unidad - Línea 2:", cfg["P2_KG_CARTON"], "(Kilos de cartón necesarios para fabricar 1 unidad de la Línea 2. Ej: 2.0 kg por juego)")

# SECCIÓN 4
st.markdown("<div class='sec-header'>4. BENEFICIOS SOCIALES EXTERNOS (LIBERACIÓN DE RECURSOS - PUNTO 8)</div>", unsafe_allow_html=True)
cfg["AHORRO_MUNI_TON"] = fila_3_columnas("AHM", "Ahorro Municipal por Tonelada Desviada ($):", cfg["AHORRO_MUNI_TON"], "(Dinero que ahorra la Municipalidad de Osorno en camiones y Vertedero de Curaco por cada tonelada reciclada: ej. 48000)")
cfg["OTROS_BEN_SOC_ANUAL"] = fila_3_columnas("OTB", "Otros Beneficios Sociales Anuales ($):", cfg["OTROS_BEN_SOC_ANUAL"], "(Valor anual en pesos de beneficios extra para la comuna: pon 2500000 o 0 si aplicas criterio conservador SNI 2026)")

# SECCIÓN 5
st.markdown("<div class='sec-header'>5. INVERSIONES EN ACTIVOS FIJOS CON IVA (PUNTO 5.b)</div>", unsafe_allow_html=True)
etiquetas_inv = [
    ("INV0", "Troqueladora Industrial ($ con IVA):", "(Costo de compra en Año 0: ej. 8500000. Es máquina importada, dura 10 años y al final recupera el 10% residual)"),
    ("INV1", "Guillotina, Mesones y Moldes ($ con IVA):", "(Costo de compra en Año 0: ej. 3500000. Es nacional y tiene vida útil de 5 años, se renueva en el Año 5)"),
    ("INV2", "Equipo Serigrafía y Secado ($ con IVA):", "(Costo de compra en Año 0: ej. 2000000. Es importado, vida útil de 5 años con renovación en el Año 5)"),
    ("INV3", "Habilitación de Galpón y Permisos ($ con IVA):", "(Gasto inicial en Año 0 en adecuar el taller en Ovejería, electricidad y permisos SEREMI: ej. 2000000)"),
]
for idx, (k_inv, lbl_i, exp_i) in enumerate(etiquetas_inv):
    cfg["INVERSIONES_LISTA"][idx]["monto_bruto"] = fila_3_columnas(
        k_inv, lbl_i, cfg["INVERSIONES_LISTA"][idx]["monto_bruto"], exp_i
    )

# SECCIÓN 6
st.markdown("<div class='sec-header'>6. COSTOS OPERATIVOS VARIABLES Y FIJOS (PUNTO 5.c)</div>", unsafe_allow_html=True)
etiquetas_cv = [
    ("CV0", "Costo Unitario Tintas y Adhesivos ($ c/IVA):", "(Gasto variable en tintas ecológicas al agua EN-71, pegamento y laca por cada juego fabricado: ej. 1800)"),
    ("CV1", "Costo Unitario Empaque y Etiqueta ($ c/IVA):", "(Gasto variable en caja protectora, instructivo y etiquetado por cada juego fabricado: ej. 600)"),
    ("CV2", "Costo Flete Furgón por Tonelada ($ c/IVA):", "(Tarifa pagada al furgón contratado por cada tonelada retirada en calle Eleuterio Ramírez: ej. 30000)"),
]
for idx, (k_cv, lbl_cv, exp_cv) in enumerate(etiquetas_cv):
    cfg["COSTOS_VAR_LISTA"][idx]["costo"] = fila_3_columnas(
        k_cv, lbl_cv, cfg["COSTOS_VAR_LISTA"][idx]["costo"], exp_cv
    )

etiquetas_cf = [
    ("CF0", "Arriendo Mensual Galpón Ovejería ($ Exento):", "(Arriendo mensual del galpón de 140 m2 en Sector Industrial Ovejería: ej. 450000 pesos al mes)"),
    ("CF1", "Luz (Saesa), Agua (Suralis) e Internet ($/mes):", "(Gasto fijo mensual en electricidad trifásica Saesa, agua potable Suralis e internet: ej. 180000 al mes)"),
    ("CF2", "Contador, Patente y Seguros ($/mes c/IVA):", "(Gasto fijo mensual en contabilidad Pro Pyme, patente municipal en Osorno y seguros: ej. 150000 al mes)"),
]
for idx, (k_cf, lbl_cf, exp_cf) in enumerate(etiquetas_cf):
    cfg["COSTOS_FIJOS_LISTA"][idx]["mensual"] = fila_3_columnas(
        k_cf, lbl_cf, cfg["COSTOS_FIJOS_LISTA"][idx]["mensual"], exp_cf
    )

# SECCIÓN 7
st.markdown("<div class='sec-header'>7. DOTACIÓN DE PERSONAL Y SUELDOS MENSUALES (PUNTO 4.3)</div>", unsafe_allow_html=True)
etiquetas_pers = [
    ("PE0", "Sueldo Bruto Jefe de Planta (1 persona - MOC):", "(Sueldo mensual bruto del Jefe de Planta y Ventas [1 trabajador MOC]: ej. 950000 pesos/mes)"),
    ("PE1", "Sueldo Bruto Operador Técnico (2 personas - MOSC):", "(Sueldo mensual bruto unitario de cada Operador de Corte y Pintura [son 2 trabajadores MOSC]: ej. 600000 c/u)"),
    ("PE2", "Sueldo Bruto Operario Recolección (2 pers. - MONC):", "(Sueldo mensual bruto unitario de cada Operario de Recolección y Armado [son 2 trabajadores MONC]: ej. 520000 c/u)"),
]
for idx, (k_pe, lbl_pe, exp_pe) in enumerate(etiquetas_pers):
    cfg["PERSONAL_LISTA"][idx]["sueldo_bruto_mensual"] = fila_3_columnas(
        k_pe, lbl_pe, cfg["PERSONAL_LISTA"][idx]["sueldo_bruto_mensual"], exp_pe
    )

res_vivo = mod2.ejecutar_motor_financiero(cfg)
m = res_vivo["metricas"]
tp = cfg["TASA_DESCUENTO_PRIVADA"] * 100
ts = cfg["TASA_DESCUENTO_SOCIAL"] * 100

banner_placeholder.markdown(
    f"""
    <div style="background-color:#0F172A;padding:14px 22px;border-radius:8px;margin-bottom:10px;">
        <div style="color:#FFFFFF;font-weight:700;font-size:15px;margin-bottom:6px;">
            PANEL INTERACTIVO DE SUPUESTOS Y SIMULACIÓN EN VIVO (PAUTA 10 PUNTOS ULAGOS)
        </div>
        <div style="color:#38BDF8;font-family:Consolas,monospace;font-weight:700;font-size:14px;">
            &gt; PRIVADO ({tp:.1f}%): &nbsp;VAN = {mod2.fmt_clp(m['van_privado'])} &nbsp;&nbsp;|&nbsp;&nbsp; TIR = {m['tir_privada']:.2f}% &nbsp;&nbsp;|&nbsp;&nbsp; Payback = {m['payback_privado']} años
        </div>
        <div style="color:#4ADE80;font-family:Consolas,monospace;font-weight:700;font-size:14px;margin-top:3px;">
            &gt; SOCIAL &nbsp;({ts:.1f}%): &nbsp;VANS = {mod2.fmt_clp(m['van_social'])} &nbsp;&nbsp;|&nbsp;&nbsp; TIRS = {m['tir_social']:.2f}% &nbsp;&nbsp;|&nbsp;&nbsp; Payback = {m['payback_social']} años
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()

b_izq, b_der = st.columns(2)
if b_izq.button("🔄 RECALCULAR VAN / TIR EN VIVO", use_container_width=True):
    st.toast("VAN y TIR recalculados en el panel superior.", icon="✅")

if b_der.button("🚀 GENERAR ENTREGA COMPLETA (WORD + PDF + 3 EXCELS EN .ZIP)", type="primary", use_container_width=True):
    with st.spinner("Ejecutando Scripts 2 + 3, reparando Índice y empaquetando ZIP..."):
        dir_temp = tempfile.mkdtemp()
        info = generar_entrega_completa_en_servidor(cfg, dir_temp)
        with open(info["ruta_zip"], "rb") as fz:
            st.session_state["ultimo_zip_bytes"] = fz.read()
            st.session_state["ultimo_zip_nombre"] = os.path.basename(info["ruta_zip"])

if "ultimo_zip_bytes" in st.session_state:
    st.success(f"¡Paquete '{st.session_state['ultimo_zip_nombre']}' generado con tus cifras actuales!")
    st.download_button(
        label=f"📥 HAZ CLIC AQUÍ PARA DESCARGAR {st.session_state['ultimo_zip_nombre']}",
        data=st.session_state["ultimo_zip_bytes"],
        file_name=st.session_state["ultimo_zip_nombre"],
        mime="application/zip",
        use_container_width=True,
    )
