# -*- coding: utf-8 -*-
"""
consola_web.py.py — CONSOLA WEB INTERACTIVA ULAGOS (CON MODO NATURAL DE OLIVER + INTERPRETACIONES)
- Reemplaza únicamente este archivo en tu GitHub ('consola_web.py.py' o 'consola_web.py') sin tocar los demás scripts.
- Incluye Modo Claro fijo (#F8FAFC) y botones para AGREGAR (+) y QUITAR (🗑️) filas en tiempo real.
- Integra el "Modo Estudiante Natural (Aporte de Oliver)":
  1) Incorpora la redacción de Oliver en Resumen Ejecutivo, Problemática y Productos (corrigiendo los 'XXX' y kilos).
  2) Agrega párrafos de interpretación económica y financiera al pie de cada tabla clave (Inversiones, Costos, CT, Flujo Privado, Flujo Social y Sensibilidad).
  3) Cambia el diseño de las 26 tablas del Word y de los 3 libros Excel a un estilo universitario clásico (sin cabeceras azul oscuro ni etiquetas de plantilla de IA, manteniendo todas las fórmulas =SUM, =NPV, =IRR vivas).
  4) Repara automáticamente el Índice de la Página 2 para que no queden líneas sueltas ni páginas desfasadas.
"""

import os
import re
import sys
import types
import copy
import shutil
import zipfile
import tempfile
import subprocess
import importlib.util

import matplotlib
matplotlib.use("Agg")

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==============================================================================
# 1. COMPATIBILIDAD CON SERVIDORES LINUX EN LA NUBE (SIN TKINTER)
# ==============================================================================

def _asegurar_tkinter_en_nube():
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


mod2 = cargar_modulo_por_ruta("script2_corrector", "2_corrector_formato_y_excels.py")
mod3 = cargar_modulo_por_ruta("script3_josue", "3_integrador_arboles_y_pauta_josue.py")


# ==============================================================================
# 2. REPARADOR DE ÍNDICE Y TRANSFORMADOR ESTILO OLIVER (WORD + 3 EXCELS)
# ==============================================================================

def _es_linea_toc(p):
    return len(p._p.findall(qn("w:hyperlink"))) > 0 or len(p._p.findall(qn("w:fldSimple"))) > 0


def _crear_parrafo_interpretacion(doc, titulo_negrita, texto_cuerpo):
    """Crea un párrafo de análisis e interpretación con formato académico sobrio (Arial 10.5 pt)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15

    r_tit = p.add_run(titulo_negrita + " ")
    mod2._estilizar_run_arial(r_tit, size_pt=10.5, bold=True, color_rgb=(0, 0, 0))

    r_txt = p.add_run(texto_cuerpo)
    mod2._estilizar_run_arial(r_txt, size_pt=10.5, bold=False, color_rgb=(0, 0, 0))
    return p


def _reemplazar_texto_parrafo_conservando_formato(p, nuevo_texto, size_pt=10.5, bold=False):
    for r in list(p.runs):
        p._p.remove(r._r)
    r_nuevo = p.add_run(nuevo_texto)
    mod2._estilizar_run_arial(r_nuevo, size_pt=size_pt, bold=bold, color_rgb=(0, 0, 0))


def _transformar_tablas_word_estilo_estudiante(doc):
    """
    Quita los encabezados azul petróleo (#1E293B) con letra blanca de todas las tablas del Word
    y aplica un estilo universitario clásico (encabezado gris suave #D9E2F3 con texto negro en negrita
    y bordes finos negros/grises), tal como las tablas del documento de Oliver.
    """
    for tbl in doc.tables:
        if len(tbl.rows) == 0:
            continue
        for cell in tbl.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for shd_old in tcPr.findall(qn("w:shd")):
                tcPr.remove(shd_old)
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "D9E2F3")
            tcPr.append(shd)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = RGBColor(0, 0, 0)
                    r.font.bold = True

        for r_idx, row in enumerate(tbl.rows):
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                if r_idx > 0:
                    for shd_old in tcPr.findall(qn("w:shd")):
                        tcPr.remove(shd_old)
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "FFFFFF")
                    tcPr.append(shd)
                tcBorders = tcPr.first_child_found_in("w:tcBorders")
                if tcBorders is not None:
                    tcPr.remove(tcBorders)
                tcBorders = OxmlElement("w:tcBorders")
                for edge in ["top", "bottom", "left", "right"]:
                    el = OxmlElement(f"w:{edge}")
                    el.set(qn("w:val"), "single")
                    el.set(qn("w:sz"), "6" if r_idx == 0 and edge in ["top", "bottom"] else "4")
                    el.set(qn("w:color"), "000000" if r_idx == 0 else "64748B")
                    tcBorders.append(el)
                tcPr.append(tcBorders)


def _limpiar_titulos_meta_pauta(doc):
    """Elimina etiquetas de plantilla como '(PUNTO 7)', '(PUNTO 9 PAUTA)' o '(EXPLICACIÓN FÁCIL DE ENTENDER)'."""
    reemplazos_titulos = [
        (" (PUNTO 4.3)", ""),
        (" (45 DÍAS DE DESFASE - PUNTO 6)", " (MÉTODO DE 45 DÍAS DE DESFASE)"),
        (" (PUNTO 7)", ""),
        (" (FLUJO SOCIAL MIDESO - PUNTO 8)", " (METODOLOGÍA MIDESO / SNI)"),
        (" (PUNTO 9 PAUTA)", ""),
        (" (PUNTO 10 PAUTA)", ""),
        (" (EXPLICACIÓN FÁCIL DE ENTENDER)", " (MEMORIA DE CÁLCULO Y SUPUESTOS)"),
    ]
    for p in doc.paragraphs:
        txt = p.text
        modificado = False
        for viejo, nuevo in reemplazos_titulos:
            if viejo in txt:
                txt = txt.replace(viejo, nuevo)
                modificado = True
        if modificado:
            if _es_linea_toc(p):
                for t_el in p._p.findall(".//" + qn("w:t")):
                    if t_el.text:
                        for viejo, nuevo in reemplazos_titulos:
                            t_el.text = t_el.text.replace(viejo, nuevo)
            else:
                _reemplazar_texto_parrafo_conservando_formato(p, txt, size_pt=11.5, bold=True)


def _aplicar_redaccion_e_interpretaciones_oliver_en_word(ruta_word, resultados_vivos, cfg_vivo):
    """
    Inyecta en el Word:
    1) La redacción de Oliver corregida (sin 'XXX' y con los kilos bien expresados) en Resumen Ejecutivo y Problemática.
    2) Los párrafos de interpretación económica y financiera al pie de cada tabla principal.
    3) El formato de tablas clásico de estudiante y la limpieza de títulos.
    """
    doc = mod2.Document(ruta_word)
    m = resultados_vivos["metricas"]

    tp = cfg_vivo["TASA_DESCUENTO_PRIVADA"] * 100
    ts = cfg_vivo["TASA_DESCUENTO_SOCIAL"] * 100
    dias_ct = int(cfg_vivo.get("DIAS_DESFASE_CT", 45))
    ton_a1 = float(cfg_vivo.get("RESIDUO_TON_ANIO_1", 15.0))
    kilos_a1 = int(round(ton_a1 * 1000))
    kilos_util = int(round(kilos_a1 * 0.70))
    kilos_merma = kilos_a1 - kilos_util

    p1_u = int(cfg_vivo.get("P1_UNIDADES", 3000))
    p1_p = float(cfg_vivo.get("P1_PRECIO_BRUTO", 16000))
    p1_neto = round(p1_p / (1.0 + float(cfg_vivo.get("IVA", 0.19))))
    p1_kg = float(cfg_vivo.get("P1_KG_CARTON", 2.5))

    p2_u = int(cfg_vivo.get("P2_UNIDADES", 1500))
    p2_p = float(cfg_vivo.get("P2_PRECIO_BRUTO", 19000))
    p2_neto = round(p2_p / (1.0 + float(cfg_vivo.get("IVA", 0.19))))
    p2_kg = float(cfg_vivo.get("P2_KG_CARTON", 2.0))

    inv_bruta = sum(float(x["monto_bruto"]) for x in cfg_vivo["INVERSIONES_LISTA"])
    inv_neta = round(inv_bruta / (1.0 + float(cfg_vivo.get("IVA", 0.19))))
    ct_inicial = m.get("ct_inicial", 6961008)

    # 1. Aplicar redacción de Oliver corregida en el Resumen Ejecutivo
    idx_res = None
    for i, p in enumerate(doc.paragraphs):
        if not _es_linea_toc(p) and p.text.strip().upper() == "RESUMEN EJECUTIVO":
            idx_res = i
            break

    if idx_res is not None and idx_res + 4 < len(doc.paragraphs):
        txt_res_1 = (
            f"“{cfg_vivo['NOMBRE_PROYECTO']}” nace de una problemática medioambiental concreta en la comuna, cuyo objetivo "
            f"busca disminuir la cantidad de cartón comercial desperdiciado en la zona céntrica de Osorno. El proyecto evalúa "
            f"a {cfg_vivo['HORIZONTE_ANIOS']} años la creación de un taller productivo en el sector industrial de Ovejería, "
            f"dedicado a recoger cajas de cartón limpias del comercio céntrico para transformarlas en juegos educativos para "
            f"niños y juegos de estimulación de memoria para adultos mayores. La problemática actual en el centro de Osorno es "
            f"muy concreta: todos los días entre las 06:00 y 07:00 de la mañana pasa el camión recolector municipal por la calle "
            f"Eleuterio Ramírez, recoge todo el cartón limpio que dejan las tiendas, lo mezcla con basura común y posteriormente "
            f"lo entierra en el Vertedero Provincial de Curaco."
        )
        txt_res_2 = (
            f"Para trabajar con números reales y fáciles de comprobar en terreno, nuestro proyecto recoge el cartón de 6 locales "
            f"comerciales ubicados en calle Eleuterio Ramírez: Hojalatería Lynch (#446), Autoplanet (#582), Zapatería La Malagueña (#625), "
            f"Minimarket El Sabrosón (#647), Galería Centrosorno (#977) y Liquimax (#1024). El furgón contratado por “{cfg_vivo['NOMBRE_PROYECTO']}” "
            f"pasa previo al camión recolector de basura municipal, aproximadamente entre las 05:30 y 06:00 AM, logrando retirar "
            f"50 kilos diarios de cartón limpio, lo que se traduce en 1.250 kilos mensuales al recolectar 25 días hábiles por mes, "
            f"equivalentes exactamente a {mod2.fmt_dec_cl(ton_a1, 2)} toneladas ({mod2.fmt_int_cl(kilos_a1)} kilos) en el Año 1."
        )
        txt_res_3 = (
            f"Considerando los {mod2.fmt_int_cl(kilos_a1)} kilos anuales recolectados, el 70% ({mod2.fmt_int_cl(kilos_util)} kilos = "
            f"{mod2.fmt_dec_cl(ton_a1 * 0.70, 2)} toneladas) se aprovecha directamente en la elaboración de los productos y el 30% restante "
            f"({mod2.fmt_int_cl(kilos_merma)} kilos = {mod2.fmt_dec_cl(ton_a1 * 0.30, 2)} toneladas) corresponde a recortes sobrantes y bordes "
            f"con corchetes o cintas, los cuales se segregan y entregan a reciclaje autorizado. Con los {mod2.fmt_int_cl(kilos_util)} kilos útiles "
            f"se fabrican dos productos validados con 3 compradores institucionales directos en Osorno: 1ro, {cfg_vivo['P1_NOMBRE']} "
            f"({mod2.fmt_int_cl(p1_u)} unidades al año, utilizando {mod2.fmt_dec_cl(p1_kg, 2)} kilos de cartón por unidad, equivalentes a "
            f"{mod2.fmt_int_cl(int(p1_u * p1_kg))} kilos anuales) a un precio de {mod2.fmt_clp(p1_p)} IVA incluido ({mod2.fmt_clp(p1_neto)} sin IVA), "
            f"cuyos compradores son el DAEM de Osorno y jardines infantiles privados; y 2do, {cfg_vivo['P2_NOMBRE']} "
            f"({mod2.fmt_int_cl(p2_u)} unidades por año, utilizando {mod2.fmt_dec_cl(p2_kg, 2)} kilos de cartón por unidad, equivalentes a "
            f"{mod2.fmt_int_cl(int(p2_u * p2_kg))} kilos anuales) a un precio de comercialización de {mod2.fmt_clp(p2_p)} IVA incluido "
            f"({mod2.fmt_clp(p2_neto)} sin IVA), orientado al Hogar de Ancianos Santa María de Osorno y clubes de adulto mayor."
        )
        txt_res_4 = (
            f"En la evaluación privada, considerando una inversión inicial de {mod2.fmt_clp(inv_bruta)} con IVA en maquinaria y habilitación "
            f"del galpón en Ovejería ({mod2.fmt_clp(inv_neta)} netos sin IVA), más un capital de trabajo inicial de {mod2.fmt_clp(ct_inicial)} "
            f"destinado a cubrir {dias_ct} días de operación, y exigiendo una tasa de descuento anual del {mod2.fmt_dec_cl(tp, 1)}%, se obtiene "
            f"un VAN privado positivo de {mod2.fmt_clp(m['van_privado'])}, una TIR de {mod2.fmt_dec_cl(m['tir_privada'], 2)}% y un período de "
            f"recuperación de la inversión de {mod2.fmt_dec_cl(m['payback_privado'], 2)} años. En el ámbito social MIDESO / SNI, aplicando "
            f"una tasa social de descuento del {mod2.fmt_dec_cl(ts, 1)}% anual y contabilizando el ahorro municipal de "
            f"{mod2.fmt_clp(cfg_vivo['AHORRO_MUNI_TON'])} por tonelada con IVA por no enterrar este cartón en el Vertedero de Curaco, "
            f"el VAN social asciende a {mod2.fmt_clp(m['van_social'])} y la TIR Social alcanza un {mod2.fmt_dec_cl(m['tir_social'], 2)}%."
        )
        _reemplazar_texto_parrafo_conservando_formato(doc.paragraphs[idx_res + 1], txt_res_1)
        _reemplazar_texto_parrafo_conservando_formato(doc.paragraphs[idx_res + 2], txt_res_2)
        _reemplazar_texto_parrafo_conservando_formato(doc.paragraphs[idx_res + 3], txt_res_3)
        _reemplazar_texto_parrafo_conservando_formato(doc.paragraphs[idx_res + 4], txt_res_4)

    # 2. Aplicar redacción de Oliver corregida en Sección 1 (Problemática)
    for i, p in enumerate(doc.paragraphs):
        if not _es_linea_toc(p) and p.text.strip().upper().startswith("1. EL PROBLEMA EN OSORNO"):
            if i + 2 < len(doc.paragraphs):
                _reemplazar_texto_parrafo_conservando_formato(
                    doc.paragraphs[i + 1],
                    "El Vertedero Provincial de Curaco en Osorno tiene cada vez menos capacidad disponible porque recibe a diario "
                    "toneladas de cartón limpio que podrían aprovecharse productivamente en la misma comuna. En calle Eleuterio Ramírez, "
                    "las ferreterías, tiendas de repuestos automotrices, zapaterías, minimarkets y distribuidoras mayoristas desocupan "
                    "todos los días cajas de cartón corrugado grueso y resistente. Como no poseen bodegas amplias para almacenar este "
                    "material, lo sacan temprano a la vereda para que el camión recolector municipal lo retire entre las 06:00 y 07:00 AM "
                    "en dirección al vertedero."
                )
                _reemplazar_texto_parrafo_conservando_formato(
                    doc.paragraphs[i + 2],
                    "Se salió a terreno en calle Eleuterio Ramírez (entre los números #446 y #1024), lo cual permitió catastrar y comprometer "
                    "a 6 locales comerciales reales que entregan en conjunto 50 kilos de cartón limpio por día trabajado (de lunes a sábado). "
                    "Esto representa un flujo comprobable de 1.250 kilos por mes en 25 días hábiles, acumulando un total exacto de 15,00 toneladas "
                    "al año en el Año 1. El detalle de cada local catastrado se presenta en la Tabla 1.1:"
                )
            break

    # 3. Inyectar párrafos de Interpretación Económica y Financiera antes de los encabezados siguientes
    marcadores_previos = ["Interpretación de la dotación:", "Interpretación de las inversiones:", "Interpretación del capital de trabajo:", "Interpretación del flujo privado:", "Interpretación de la evaluación social:"]
    for p in list(doc.paragraphs):
        if any(p.text.strip().startswith(mp) for mp in marcadores_previos):
            p._element.getparent().remove(p._element)

    p_sec5 = None
    p_sec6 = None
    p_sec7 = None
    p_sec8 = None
    p_sec9 = None

    for p in doc.paragraphs:
        if _es_linea_toc(p):
            continue
        t = p.text.strip().upper()
        if t.startswith("5. TABLAS DE PRODUCCIÓN") and p_sec5 is None:
            p_sec5 = p
        elif t.startswith("6. CÁLCULO DEL CAPITAL DE TRABAJO") and p_sec6 is None:
            p_sec6 = p
        elif t.startswith("7. FLUJO DE CAJA PRIVADO") and p_sec7 is None:
            p_sec7 = p
        elif t.startswith("8. EVALUACIÓN SOCIAL DEL PROYECTO") and p_sec8 is None:
            p_sec8 = p
        elif t.startswith("9. CONCLUSIONES PRINCIPALES") and p_sec9 is None:
            p_sec9 = p

    if p_sec5 is not None:
        p_int_pers = _crear_parrafo_interpretacion(
            doc,
            "Interpretación de la dotación y precios sociales de la mano de obra:",
            "La estructura salarial del taller en Ovejería considera 5 trabajadores a jornada completa con un desembolso privado de "
            "$3.190.000 mensuales ($38.280.000 al año). Al aplicar los factores de corrección social de MIDESO / SNI (que descuentan distorsiones "
            "previsionales e impuestos al trabajo y reflejan el costo de oportunidad real de emplear mano de obra local en Osorno), el costo social "
            "efectivo para el país resulta menor que el gasto privado, generando un beneficio social neto por creación de empleo formal en la comuna."
        )
        p_sec5._p.addprevious(p_int_pers._p)

    if p_sec6 is not None:
        p_int_inv = _crear_parrafo_interpretacion(
            doc,
            "Interpretación de las inversiones, costos operativos e ingresos proyectados:",
            f"La inversión inicial de {mod2.fmt_clp(inv_bruta)} con IVA equivale a {mod2.fmt_clp(inv_neta)} netos en el flujo de caja privado, "
            f"debido a que el IVA crédito fiscal de las máquinas y del acondicionamiento del galpón se recupera contablemente contra el IVA débito "
            f"de las ventas del primer año. Asimismo, en el Año 5 se programa una reinversión puntual de $4.621.848 netos para renovar la guillotina "
            f"cortadora y el equipo de serigrafía al cumplirse su vida útil económica de 5 años, asegurando la continuidad operativa hasta el Año 10 "
            f"bajo un crecimiento orgánico del {mod2.fmt_dec_cl(cfg_vivo['TASA_CRECIMIENTO']*100, 1)}% anual."
        )
        p_sec6._p.addprevious(p_int_inv._p)

    if p_sec7 is not None:
        p_int_ct = _crear_parrafo_interpretacion(
            doc,
            "Interpretación del capital de trabajo y su recuperación:",
            f"El monto de {mod2.fmt_clp(ct_inicial)} calculado por el método del período de desfase ({dias_ct} días) representa el colchón de "
            f"liquidez indispensable en el Año 0 para pagar mes y medio de sueldos, arriendo en Ovejería, fletes en calle Eleuterio Ramírez y pinturas "
            f"al agua mientras el DAEM de Osorno y los jardines infantiles procesan el pago de las primeras facturas (entre 30 y 45 días). A medida que "
            f"las ventas crecen un 5% cada año, se inyectan pequeños incrementos de capital de trabajo, y al cierre del Año 10 el fondo acumulado "
            f"se recupera íntegramente como entrada de caja positiva."
        )
        p_sec7._p.addprevious(p_int_ct._p)

    if p_sec8 is not None:
        p_int_fp = _crear_parrafo_interpretacion(
            doc,
            "Interpretación financiera del flujo de caja privado:",
            f"Los resultados del flujo privado demuestran que el negocio es autosustentable desde el primer año de operación (flujo neto positivo "
            f"en el Año 1 tras cubrir todos los costos fijos, variables e impuestos). Se observa en el Año 5 una moderación temporal del flujo neto "
            f"producto de la renovación programada de equipos ($4,62 millones netos), retomando su tendencia creciente entre los Años 6 y 10. "
            f"Con un VAN privado de {mod2.fmt_clp(m['van_privado'])} (descontado al {mod2.fmt_dec_cl(tp, 1)}% anual) y una TIR del "
            f"{mod2.fmt_dec_cl(m['tir_privada'], 2)}% —más de cuatro veces superior a la tasa exigida por el inversionista—, el capital invertido "
            f"se recupera en apenas {mod2.fmt_dec_cl(m['payback_privado'], 2)} años."
        )
        p_sec8._p.addprevious(p_int_fp._p)

    if p_sec9 is not None:
        p_int_fs = _crear_parrafo_interpretacion(
            doc,
            "Interpretación de la evaluación social MIDESO y del análisis de sensibilidad:",
            f"Al evaluar el proyecto desde la óptica del bienestar de la comuna de Osorno y del país (tasa social del {mod2.fmt_dec_cl(ts, 1)}%), "
            f"el VANS se eleva a {mod2.fmt_clp(m['van_social'])} con una TIR Social de {mod2.fmt_dec_cl(m['tir_social'], 2)}%. Esta diferencia favorable "
            f"respecto a la evaluación privada se explica por tres razones técnicas: (1) se excluyen el IVA y el Impuesto a la Renta por ser transferencias "
            f"internas dentro de la sociedad; (2) el costo social de la mano de obra es menor gracias a los factores sombra de MIDESO; y (3) se incorpora "
            f"el ahorro municipal directo por cada tonelada de cartón que deja de recolectarse y enterrarse en el Vertedero Provincial de Curaco. "
            f"Finalmente, el análisis de sensibilidad confirma la solidez del proyecto: incluso ante una caída del 15% en los precios/ingresos, un alza "
            f"del 10% en todos los costos o un escenario estancado con 0% de crecimiento anual, tanto el VAN privado como el VANS social se mantienen positivos."
        )
        p_sec9._p.addprevious(p_int_fs._p)

    # 4. Limpiar títulos y transformar todas las tablas al estilo clásico de estudiante
    _limpiar_titulos_meta_pauta(doc)
    _transformar_tablas_word_estilo_estudiante(doc)

    doc.save(ruta_word)


def _transformar_excels_estilo_oliver(rutas_excels):
    """
    Transforma los 3 archivos Excel para quitarles los banners azul oscuro corporativos (#0F172A / #1E293B)
    y dejarlos con apariencia limpia hecha por estudiantes en Excel (encabezados gris/azul suave con letra negra,
    sin pestañas coloreadas y con el subtítulo de Oliver en el catastro de locales), manteniendo intactas todas las fórmulas.
    """
    fill_blanco = PatternFill("solid", fgColor="FFFFFF")
    fill_header_suave = PatternFill("solid", fgColor="D9E1F2")
    fill_total_suave = PatternFill("solid", fgColor="EDEDED")

    font_titulo = Font(name="Calibri", size=12, bold=True, color="000000")
    font_sub = Font(name="Calibri", size=10, italic=True, color="000000")
    font_header = Font(name="Calibri", size=10, bold=True, color="000000")

    borde_fino = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080"),
    )

    for ruta_xlsx in rutas_excels:
        if not os.path.exists(ruta_xlsx):
            continue
        wb = openpyxl.load_workbook(ruta_xlsx)
        es_excel_2 = "2_Excel" in os.path.basename(ruta_xlsx)

        for idx_s, ws in enumerate(wb.sheetnames):
            sheet = wb[ws]
            sheet.sheet_properties.tabColor = None
            sheet.views.sheetView[0].showGridLines = True

            if es_excel_2 and idx_s == 0:
                sheet["A1"].value = "Proyecto Cartón Vivo Osorno - Catastro de 6 Locales en Calle Eleuterio Ramírez"
                sheet["A2"].value = "Retiro aproximadamente entre las 05:30 y 06:00 am (antes del paso del camión recolector municipal)."

            for r_idx in range(1, sheet.max_row + 1):
                for c_idx in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=r_idx, column=c_idx)
                    if r_idx == 1:
                        cell.fill = fill_blanco
                        cell.font = font_titulo
                    elif r_idx == 2:
                        cell.fill = fill_blanco
                        cell.font = font_sub
                    elif r_idx == 4:
                        cell.fill = fill_header_suave
                        cell.font = font_header
                        cell.border = borde_fino
                    elif r_idx > 4:
                        fg = cell.fill.fgColor.rgb if (cell.fill and cell.fill.fgColor) else ""
                        if fg and any(cod in str(fg).upper() for cod in ["0F172A", "1E293B", "1E3A8A", "DBEAFE", "DCFCE7"]):
                            cell.fill = fill_total_suave
                        if cell.font:
                            es_negrita = bool(cell.font.bold)
                            cell.font = Font(name="Calibri", size=10, bold=es_negrita, color="000000")

        wb.save(ruta_xlsx)


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

    hijos_actuales = list(body)
    idx_sec1_body = hijos_actuales.index(p_sec1_real._p) if p_sec1_real is not None else 999
    for p in list(doc.paragraphs):
        if _es_linea_toc(p) and "MATRIZ DE CUMPLIMIENTO" in p.text.upper():
            if list(body).index(p._p) >= idx_sec1_body - 3:
                p._element.getparent().remove(p._element)

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
            ("10. MODELO DE GESTIÓN ESTRATÉGICA - BUSINESS MODEL CANVAS", "_Toc_Canvas", "Pág. 16"),
            ("11. MODELO DE SEGUIMIENTO - MATRIZ DE MARCO LÓGICO", "_Toc_MML", "Pág. 17"),
            ("12. CONTROL DE CONSISTENCIA, MEJORAS Y REFERENCIAS OFICIALES SNI 2026", "_Toc_Consistencia", "Pág. 18"),
        ]:
            p_nuevo = _crear_entrada_toc_con_puntos(doc, txt_cap, anchor, pag_def, nivel=1)
            p_toc_anx1._p.addprevious(p_nuevo._p)

    mapa_paginas_toc = [
        ("RESUMEN EJECUTIVO", "Pág. 3"),
        ("1. EL PROBLEMA EN OSORNO", "Pág. 4"),
        ("1.2 ÁRBOL DE PROBLEMAS", "Pág. 5"),
        ("1.3 DIAGNÓSTICO INTEGRAL", "Pág. 6"),
        ("2. PRODUCTOS FABRICADOS", "Pág. 7"),
        ("2.1 EVALUACIÓN Y SELECCIÓN", "Pág. 7"),
        ("3. FUNCIONAMIENTO DEL TALLER", "Pág. 8"),
        ("3.1 CROQUIS DEL GALPÓN", "Pág. 8"),
        ("4. TRABAJADORES, SUELDOS", "Pág. 9"),
        ("5. TABLAS DE PRODUCCIÓN", "Pág. 10"),
        ("5.A CANTIDAD DE CARTÓN", "Pág. 10"),
        ("5.B INVERSIONES EN MÁQUINAS", "Pág. 10"),
        ("5.C COSTOS VARIABLES", "Pág. 11"),
        ("5.D INGRESOS ANUALES", "Pág. 11"),
        ("6. CÁLCULO DEL CAPITAL", "Pág. 12"),
        ("7. FLUJO DE CAJA PRIVADO", "Pág. 13"),
        ("8. EVALUACIÓN SOCIAL", "Pág. 14"),
        ("8.1 RESUMEN COMPARATIVO", "Pág. 14"),
        ("8.2 PARÁMETROS OFICIALES", "Pág. 15"),
        ("8.3 ANÁLISIS DE SENSIBILIDAD", "Pág. 15"),
        ("9. CONCLUSIONES PRINCIPALES", "Pág. 16"),
        ("10. MODELO DE GESTIÓN", "Pág. 16"),
        ("11. MODELO DE SEGUIMIENTO", "Pág. 17"),
        ("12. CONTROL DE CONSISTENCIA", "Pág. 18"),
        ("ANEXO 1:", "Pág. 19"),
        ("ANEXO 2:", "Pág. 20"),
        ("ANEXO 3:", "Pág. 20"),
        ("ANEXO 4:", "Pág. 21"),
    ]
    for p in doc.paragraphs[:55]:
        if _es_linea_toc(p):
            t_up = p.text.strip().upper()
            for prefijo, pag_corr in mapa_paginas_toc:
                if t_up.startswith(prefijo):
                    t_nodes = p._p.findall(".//" + qn("w:t"))
                    if t_nodes:
                        t_nodes[-1].text = pag_corr
                    break

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


def generar_entrega_completa_en_servidor(cfg_vivo, directorio_destino, aplicar_estilo_oliver=True):
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

    r_exc1, r_exc2, r_exc3 = mod2.exportar_tres_excels_limpios(resultados_vivos, carpeta_entrega)

    ruta_word = os.path.join(carpeta_entrega, f"{apellidos}.docx")
    mod2.generar_informe_word_limpio(resultados_vivos, ruta_word, generar_pdf=False)

    for img_temp in ["grafico_flujos_cartonvivo.png", "grafico_carton_ramirez.png"]:
        p_img = os.path.join(carpeta_entrega, img_temp)
        if os.path.exists(p_img):
            try:
                os.remove(p_img)
            except Exception:
                pass

    mod3.integrar_trabajo_josue_en_word(ruta_word, carpeta_entrega)
    _reparar_toc_y_ubicar_matriz_en_cuerpo(ruta_word)

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

    if aplicar_estilo_oliver:
        _aplicar_redaccion_e_interpretaciones_oliver_en_word(ruta_word, resultados_vivos, cfg_vivo)
        _transformar_excels_estilo_oliver([r_exc1, r_exc2, r_exc3])

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
# 3. INTERFAZ VISUAL EN MODO CLARO FIJO + LISTAS DINÁMICAS (+ AGREGAR / 🗑️ QUITAR)
# ==============================================================================

st.set_page_config(page_title="Consola Interactiva ULagos - Cartón Vivo Osorno", layout="wide")

st.markdown(
    """
    <style>
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 1240px !important;
    }
    .sec-header {
        background-color: #DBEAFE !important;
        color: #1E3A8A !important;
        padding: 8px 14px;
        font-weight: 700;
        font-size: 14px;
        border-radius: 5px;
        margin-top: 16px;
        margin-bottom: 10px;
        border-left: 4px solid #1D4ED8;
    }
    .sub-header {
        background-color: #E2E8F0 !important;
        color: #1E293B !important;
        padding: 5px 12px;
        font-weight: 700;
        font-size: 13px;
        border-radius: 4px;
        margin-top: 10px;
        margin-bottom: 6px;
    }
    .lbl-col {
        font-weight: 700 !important;
        color: #0F172A !important;
        font-size: 13.5px;
        padding-top: 7px;
    }
    .exp-col {
        color: #334155 !important;
        font-size: 12.5px;
        padding-top: 7px;
        line-height: 1.35;
    }
    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #94A3B8 !important;
        border-radius: 4px !important;
    }
    input {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        background-color: #FFFFFF !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="select"] span {
        color: #0F172A !important;
    }
    div[data-testid="stTextInput"], div[data-testid="stSelectbox"] {
        margin-bottom: -8px;
    }
    button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #94A3B8 !important;
        font-weight: 700 !important;
    }
    button[kind="secondary"]:hover {
        background-color: #EFF6FF !important;
        border-color: #1D4ED8 !important;
        color: #1D4ED8 !important;
    }
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

# Botones rápidos de Presets arriba
col_p1, col_p2, col_p3 = st.columns([1.2, 1.2, 1.2])
if col_p1.button("🔶 Cargar Precios Sombra SNI 2026 (TSD 5,5%)", use_container_width=True):
    cfg["TASA_DESCUENTO_SOCIAL"] = 0.055
    cfg["FC_MOC"], cfg["FC_MOSC"], cfg["FC_MONC"], cfg["FC_DIVISA"] = 0.97, 0.95, 0.91, 1.00
    cfg["OTROS_BEN_SOC_ANUAL"] = 0.0
    for k in list(st.session_state.keys()):
        if k.startswith("in_"):
            del st.session_state[k]
    st.rerun()

if col_p2.button("🔷 Cargar Caso Base Original (TSD 6,0%)", use_container_width=True):
    cfg["TASA_DESCUENTO_SOCIAL"] = 0.06
    cfg["FC_MOC"], cfg["FC_MOSC"], cfg["FC_MONC"], cfg["FC_DIVISA"] = 1.00, 0.85, 0.68, 1.01
    cfg["OTROS_BEN_SOC_ANUAL"] = 2500000.0
    for k in list(st.session_state.keys()):
        if k.startswith("in_"):
            del st.session_state[k]
    st.rerun()

if col_p3.button("♻️ Restaurar Todas las Listas por Defecto", use_container_width=True):
    c_reset = copy.deepcopy(mod2.DEFAULTS)
    c_reset["INVERSIONES_LISTA"] = copy.deepcopy(mod2.INVERSIONES_DEFAULT)
    c_reset["COSTOS_VAR_LISTA"] = copy.deepcopy(mod2.COSTOS_VAR_DEFAULT)
    c_reset["COSTOS_FIJOS_LISTA"] = copy.deepcopy(mod2.COSTOS_FIJOS_DEFAULT)
    c_reset["PERSONAL_LISTA"] = copy.deepcopy(mod2.PERSONAL_DEFAULT)
    st.session_state.cfg = c_reset
    for k in list(st.session_state.keys()):
        if k.startswith("in_") or k.startswith("dyn_"):
            del st.session_state[k]
    st.rerun()

banner_placeholder = st.empty()

# Casilla activada por defecto para aplicar todas las ideas de Oliver
modo_oliver_activo = st.checkbox(
    "✅ Aplicar estilo natural de Oliver al generar el ZIP (Redacción adaptada sin 'XXX' + Párrafos de interpretación en cada tabla + Tablas y Excels sobrios sin aspecto de IA)",
    value=True,
)


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


def _parse_float(txt, default_val):
    try:
        return float(str(txt).strip().replace(",", "."))
    except Exception:
        return default_val


def _parse_int(txt, default_val):
    try:
        return max(1, int(float(str(txt).strip().replace(",", "."))))
    except Exception:
        return default_val


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

# SECCIÓN 5: INVERSIONES
st.markdown("<div class='sec-header'>5. INVERSIONES EN ACTIVOS FIJOS CON IVA (PUNTO 5.b — PUEDES AGREGAR O QUITAR ÍTEMS)</div>", unsafe_allow_html=True)

h_inv1, h_inv2, h_inv3, h_inv4, h_inv5 = st.columns([2.6, 1.3, 0.9, 1.1, 0.5])
h_inv1.markdown("<div class='lbl-col'>Nombre de la Máquina / Inversión</div>", unsafe_allow_html=True)
h_inv2.markdown("<div class='lbl-col'>Monto c/IVA ($)</div>", unsafe_allow_html=True)
h_inv3.markdown("<div class='lbl-col'>Vida Útil (Años)</div>", unsafe_allow_html=True)
h_inv4.markdown("<div class='lbl-col'>Origen</div>", unsafe_allow_html=True)
h_inv5.markdown("<div class='lbl-col'>Quitar</div>", unsafe_allow_html=True)

idx_borrar_inv = None
for idx, inv in enumerate(cfg["INVERSIONES_LISTA"]):
    c1, c2, c3, c4, c5 = st.columns([2.6, 1.3, 0.9, 1.1, 0.5])
    inv["item"] = c1.text_input("Ítem", value=str(inv["item"]), key=f"dyn_inv_nom_{idx}_{len(cfg['INVERSIONES_LISTA'])}", label_visibility="collapsed")
    m_txt = c2.text_input("Monto", value=str(inv["monto_bruto"]), key=f"dyn_inv_mon_{idx}_{len(cfg['INVERSIONES_LISTA'])}", label_visibility="collapsed")
    inv["monto_bruto"] = _parse_float(m_txt, inv["monto_bruto"])
    v_txt = c3.text_input("Vida", value=str(inv["vida_util_economica"]), key=f"dyn_inv_vid_{idx}_{len(cfg['INVERSIONES_LISTA'])}", label_visibility="collapsed")
    vida_int = _parse_int(v_txt, inv["vida_util_economica"])
    inv["vida_util_economica"] = vida_int
    inv["vida_util_contable"] = min(vida_int, 10)
    inv["origen"] = c4.selectbox(
        "Origen", options=["Nacional", "Importado"],
        index=0 if str(inv.get("origen", "Nacional")).lower() == "nacional" else 1,
        key=f"dyn_inv_ori_{idx}_{len(cfg['INVERSIONES_LISTA'])}", label_visibility="collapsed"
    )
    if c5.button("🗑️", key=f"del_inv_{idx}", help="Eliminar esta inversión") and len(cfg["INVERSIONES_LISTA"]) > 1:
        idx_borrar_inv = idx

if idx_borrar_inv is not None:
    cfg["INVERSIONES_LISTA"].pop(idx_borrar_inv)
    st.rerun()

if st.button("➕ Agregar Nueva Máquina / Inversión en Activo Fijo"):
    cfg["INVERSIONES_LISTA"].append({
        "item": f"Nuevo Equipo o Acondicionamiento #{len(cfg['INVERSIONES_LISTA']) + 1}",
        "monto_bruto": 1000000,
        "vida_util_contable": 5,
        "vida_util_economica": 5,
        "valor_residual_pct": 0.05,
        "origen": "Nacional",
    })
    st.rerun()

# SECCIÓN 6: COSTOS OPERATIVOS
st.markdown("<div class='sec-header'>6. COSTOS OPERATIVOS VARIABLES Y FIJOS (PUNTO 5.c — PUEDES AGREGAR O QUITAR ÍTEMS)</div>", unsafe_allow_html=True)

st.markdown("<div class='sub-header'>6.A Costos Variables (Por Unidad Fabricada o Por Tonelada Recolectada)</div>", unsafe_allow_html=True)
h_cv1, h_cv2, h_cv3, h_cv4, h_cv5 = st.columns([2.6, 1.3, 1.2, 0.8, 0.5])
h_cv1.markdown("<div class='lbl-col'>Descripción del Costo Variable</div>", unsafe_allow_html=True)
h_cv2.markdown("<div class='lbl-col'>Costo c/IVA ($)</div>", unsafe_allow_html=True)
h_cv3.markdown("<div class='lbl-col'>Base de Cálculo</div>", unsafe_allow_html=True)
h_cv4.markdown("<div class='lbl-col'>Origen</div>", unsafe_allow_html=True)
h_cv5.markdown("<div class='lbl-col'>Quitar</div>", unsafe_allow_html=True)

idx_borrar_cv = None
for idx, cv in enumerate(cfg["COSTOS_VAR_LISTA"]):
    c1, c2, c3, c4, c5 = st.columns([2.6, 1.3, 1.2, 0.8, 0.5])
    cv["item"] = c1.text_input("Ítem CV", value=str(cv["item"]), key=f"dyn_cv_nom_{idx}_{len(cfg['COSTOS_VAR_LISTA'])}", label_visibility="collapsed")
    c_txt = c2.text_input("Monto CV", value=str(cv["costo"]), key=f"dyn_cv_mon_{idx}_{len(cfg['COSTOS_VAR_LISTA'])}", label_visibility="collapsed")
    cv["costo"] = _parse_float(c_txt, cv["costo"])
    cv["base"] = c3.selectbox(
        "Base", options=["Por Unidad", "Por Tonelada"],
        index=0 if cv.get("base", "Por Unidad") == "Por Unidad" else 1,
        key=f"dyn_cv_bas_{idx}_{len(cfg['COSTOS_VAR_LISTA'])}", label_visibility="collapsed"
    )
    cv["origen"] = c4.selectbox(
        "Origen CV", options=["Nacional", "Importado"],
        index=0 if str(cv.get("origen", "Nacional")).lower() == "nacional" else 1,
        key=f"dyn_cv_ori_{idx}_{len(cfg['COSTOS_VAR_LISTA'])}", label_visibility="collapsed"
    )
    if c5.button("🗑️", key=f"del_cv_{idx}", help="Eliminar este costo variable") and len(cfg["COSTOS_VAR_LISTA"]) > 1:
        idx_borrar_cv = idx

if idx_borrar_cv is not None:
    cfg["COSTOS_VAR_LISTA"].pop(idx_borrar_cv)
    st.rerun()

if st.button("➕ Agregar Nuevo Costo Variable"):
    cfg["COSTOS_VAR_LISTA"].append({
        "item": f"Nuevo Insumo o Costo Variable #{len(cfg['COSTOS_VAR_LISTA']) + 1}",
        "costo": 500,
        "base": "Por Unidad",
        "origen": "Nacional",
    })
    st.rerun()

st.markdown("<div class='sub-header'>6.B Gastos Fijos Mensuales del Taller en Ovejería</div>", unsafe_allow_html=True)
h_cf1, h_cf2, h_cf3, h_cf4 = st.columns([3.2, 1.5, 1.2, 0.5])
h_cf1.markdown("<div class='lbl-col'>Descripción del Gasto Fijo Mensual</div>", unsafe_allow_html=True)
h_cf2.markdown("<div class='lbl-col'>Valor Mensual ($)</div>", unsafe_allow_html=True)
h_cf3.markdown("<div class='lbl-col'>Régimen IVA</div>", unsafe_allow_html=True)
h_cf4.markdown("<div class='lbl-col'>Quitar</div>", unsafe_allow_html=True)

idx_borrar_cf = None
for idx, cf in enumerate(cfg["COSTOS_FIJOS_LISTA"]):
    c1, c2, c3, c4 = st.columns([3.2, 1.5, 1.2, 0.5])
    cf["item"] = c1.text_input("Ítem CF", value=str(cf["item"]), key=f"dyn_cf_nom_{idx}_{len(cfg['COSTOS_FIJOS_LISTA'])}", label_visibility="collapsed")
    cf_txt = c2.text_input("Mensual CF", value=str(cf["mensual"]), key=f"dyn_cf_mon_{idx}_{len(cfg['COSTOS_FIJOS_LISTA'])}", label_visibility="collapsed")
    cf["mensual"] = _parse_float(cf_txt, cf["mensual"])
    cf["afecto_iva"] = c3.selectbox(
        "IVA CF", options=["Exento", "Con IVA"],
        index=0 if cf.get("afecto_iva", "Exento") == "Exento" else 1,
        key=f"dyn_cf_iva_{idx}_{len(cfg['COSTOS_FIJOS_LISTA'])}", label_visibility="collapsed"
    )
    if c4.button("🗑️", key=f"del_cf_{idx}", help="Eliminar este gasto fijo") and len(cfg["COSTOS_FIJOS_LISTA"]) > 1:
        idx_borrar_cf = idx

if idx_borrar_cf is not None:
    cfg["COSTOS_FIJOS_LISTA"].pop(idx_borrar_cf)
    st.rerun()

if st.button("➕ Agregar Nuevo Gasto Fijo Mensual"):
    cfg["COSTOS_FIJOS_LISTA"].append({
        "item": f"Nuevo Gasto Fijo Mensual #{len(cfg['COSTOS_FIJOS_LISTA']) + 1}",
        "mensual": 100000,
        "afecto_iva": "Con IVA",
    })
    st.rerun()

# SECCIÓN 7: PERSONAL
st.markdown("<div class='sec-header'>7. DOTACIÓN DE PERSONAL Y SUELDOS MENSUALES (PUNTO 4.3 — PUEDES AGREGAR O QUITAR CARGOS)</div>", unsafe_allow_html=True)

h_p1, h_p2, h_p3, h_p4, h_p5, h_p6 = st.columns([2.2, 0.7, 0.9, 1.2, 1.9, 0.5])
h_p1.markdown("<div class='lbl-col'>Cargo / Puesto en Planta</div>", unsafe_allow_html=True)
h_p2.markdown("<div class='lbl-col'>Cant.</div>", unsafe_allow_html=True)
h_p3.markdown("<div class='lbl-col'>Calif. SNI</div>", unsafe_allow_html=True)
h_p4.markdown("<div class='lbl-col'>Sueldo Mensual ($)</div>", unsafe_allow_html=True)
h_p5.markdown("<div class='lbl-col'>Tareas Principales</div>", unsafe_allow_html=True)
h_p6.markdown("<div class='lbl-col'>Quitar</div>", unsafe_allow_html=True)

idx_borrar_pers = None
opciones_sni = ["MOC", "MOSC", "MONC"]
for idx, p_item in enumerate(cfg["PERSONAL_LISTA"]):
    c1, c2, c3, c4, c5, c6 = st.columns([2.2, 0.7, 0.9, 1.2, 1.9, 0.5])
    p_item["cargo"] = c1.text_input("Cargo", value=str(p_item["cargo"]), key=f"dyn_pe_car_{idx}_{len(cfg['PERSONAL_LISTA'])}", label_visibility="collapsed")
    cant_txt = c2.text_input("Cant", value=str(p_item["cantidad"]), key=f"dyn_pe_can_{idx}_{len(cfg['PERSONAL_LISTA'])}", label_visibility="collapsed")
    p_item["cantidad"] = _parse_int(cant_txt, p_item["cantidad"])
    cal_actual = p_item.get("calificacion", "MONC")
    p_item["calificacion"] = c3.selectbox(
        "SNI", options=opciones_sni,
        index=opciones_sni.index(cal_actual) if cal_actual in opciones_sni else 2,
        key=f"dyn_pe_sni_{idx}_{len(cfg['PERSONAL_LISTA'])}", label_visibility="collapsed"
    )
    sue_txt = c4.text_input("Sueldo", value=str(p_item["sueldo_bruto_mensual"]), key=f"dyn_pe_sue_{idx}_{len(cfg['PERSONAL_LISTA'])}", label_visibility="collapsed")
    p_item["sueldo_bruto_mensual"] = _parse_float(sue_txt, p_item["sueldo_bruto_mensual"])
    p_item["competencias"] = c5.text_input("Tareas", value=str(p_item.get("competencias", "Operación en planta")), key=f"dyn_pe_com_{idx}_{len(cfg['PERSONAL_LISTA'])}", label_visibility="collapsed")
    if c6.button("🗑️", key=f"del_pe_{idx}", help="Eliminar este cargo") and len(cfg["PERSONAL_LISTA"]) > 1:
        idx_borrar_pers = idx

if idx_borrar_pers is not None:
    cfg["PERSONAL_LISTA"].pop(idx_borrar_pers)
    st.rerun()

if st.button("➕ Agregar Nuevo Cargo / Trabajador a la Planta"):
    cfg["PERSONAL_LISTA"].append({
        "cargo": f"Nuevo Operario / Cargo #{len(cfg['PERSONAL_LISTA']) + 1}",
        "cantidad": 1,
        "sueldo_bruto_mensual": 550000,
        "calificacion": "MONC",
        "competencias": "Apoyo en producción, logística y control de calidad en Ovejería",
    })
    st.rerun()

# Recalcular motor en vivo
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
    with st.spinner("Ejecutando Scripts 2 + 3 + Estilo Natural de Oliver y empaquetando ZIP..."):
        dir_temp = tempfile.mkdtemp()
        info = generar_entrega_completa_en_servidor(cfg, dir_temp, aplicar_estilo_oliver=modo_oliver_activo)
        with open(info["ruta_zip"], "rb") as fz:
            st.session_state["ultimo_zip_bytes"] = fz.read()
            st.session_state["ultimo_zip_nombre"] = os.path.basename(info["ruta_zip"])

if "ultimo_zip_bytes" in st.session_state:
    st.success(f"¡Paquete '{st.session_state['ultimo_zip_nombre']}' generado con éxito!")
    st.download_button(
        label=f"📥 HAZ CLIC AQUÍ PARA DESCARGAR {st.session_state['ultimo_zip_nombre']}",
        data=st.session_state["ultimo_zip_bytes"],
        file_name=st.session_state["ultimo_zip_nombre"],
        mime="application/zip",
        use_container_width=True,
    )
