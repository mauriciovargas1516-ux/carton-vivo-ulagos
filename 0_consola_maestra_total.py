# -*- coding: utf-8 -*-
"""
0_consola_maestra_total.py — PANEL INTERACTIVO DE SUPUESTOS Y SIMULACIÓN EN VIVO (CORREGIDO)
Diseño vertical alineado en 3 columnas (Etiqueta | Valor Editable | Explicación Guía)
que conecta automáticamente tus 3 scripts:
  1) 1_simulador_interactivo_cartonvivo.py
  2) 2_corrector_formato_y_excels.py
  3) 3_integrador_arboles_y_pauta_josue.py
"""

import os
import sys
import copy
import importlib.util
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==============================================================================
# 1. CARGA DINÁMICA DE LOS 3 SCRIPTS EN MEMORIA
# ==============================================================================

def cargar_modulo_por_ruta(nombre_interno, nombre_archivo):
    carpeta = os.path.dirname(os.path.abspath(__file__))
    ruta = os.path.join(carpeta, nombre_archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el archivo '{nombre_archivo}' en:\n{carpeta}")
    spec = importlib.util.spec_from_file_location(nombre_interno, ruta)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[nombre_interno] = mod
    spec.loader.exec_module(mod)
    return mod


mod1 = cargar_modulo_por_ruta("script1_simulador", "1_simulador_interactivo_cartonvivo.py")
mod2 = cargar_modulo_por_ruta("script2_corrector", "2_corrector_formato_y_excels.py")
mod3 = cargar_modulo_por_ruta("script3_josue", "3_integrador_arboles_y_pauta_josue.py")


# ==============================================================================
# 2. REPARADOR AUTOMÁTICO DE LA TABLA DE CONTENIDOS (PÁGINA 2) Y SINCRONIZADOR
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


def ejecutar_pipeline_completo_1_2_3(cfg_vivo, directorio_destino):
    for k, v in cfg_vivo.items():
        if k not in ["INVERSIONES_LISTA", "COSTOS_VAR_LISTA", "COSTOS_FIJOS_LISTA", "PERSONAL_LISTA"]:
            mod2.DEFAULTS[k] = copy.deepcopy(v)

    mod2.INVERSIONES_DEFAULT = copy.deepcopy(cfg_vivo["INVERSIONES_LISTA"])
    mod2.COSTOS_VAR_DEFAULT = copy.deepcopy(cfg_vivo["COSTOS_VAR_LISTA"])
    mod2.COSTOS_FIJOS_DEFAULT = copy.deepcopy(cfg_vivo["COSTOS_FIJOS_LISTA"])
    mod2.PERSONAL_DEFAULT = copy.deepcopy(cfg_vivo["PERSONAL_LISTA"])

    resultados_vivos = mod2.ejecutar_motor_financiero(cfg_vivo)
    apellidos = cfg_vivo.get("APELLIDOS_ENTREGA", "Vargas,Muñoz,Jofre").strip()

    if os.path.basename(os.path.normpath(directorio_destino)).lower() == apellidos.lower():
        carpeta_entrega = os.path.normpath(directorio_destino)
    else:
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
    ruta_pdf, ruta_zip, pdf_ok = mod3.actualizar_pdf_y_zip_final(ruta_word)

    return {
        "carpeta_entrega": carpeta_entrega,
        "ruta_zip": ruta_zip,
        "ruta_word": ruta_word,
        "ruta_pdf": ruta_pdf,
        "pdf_ok": pdf_ok,
        "excels": [r_exc1, r_exc2, r_exc3],
    }


# ==============================================================================
# 3. INTERFAZ VERTICAL ALINEADA DE 3 COLUMNAS (TAMAÑOS DE FUENTE ENTEROS)
# ==============================================================================

class ConsolaMaestraSupuestos:
    def __init__(self, root):
        self.root = root
        self.root.title("Simulador y Generador Automático de Proyectos - Cartón Vivo Osorno (Scripts 1 + 2 + 3 Conectados)")
        self.root.geometry("1120x880")
        self.root.configure(bg="#F8FAFC")

        self.carpeta_salida = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))

        self.cfg = copy.deepcopy(mod2.DEFAULTS)
        self.cfg["INVERSIONES_LISTA"] = copy.deepcopy(mod2.INVERSIONES_DEFAULT)
        self.cfg["COSTOS_VAR_LISTA"] = copy.deepcopy(mod2.COSTOS_VAR_DEFAULT)
        self.cfg["COSTOS_FIJOS_LISTA"] = copy.deepcopy(mod2.COSTOS_FIJOS_DEFAULT)
        self.cfg["PERSONAL_LISTA"] = copy.deepcopy(mod2.PERSONAL_DEFAULT)

        self.vars_simples = {}
        self.vars_inv = []
        self.vars_cv = []
        self.vars_cf = []
        self.vars_pers = []

        self._construir_ui()
        self.recalcular_en_vivo(mostrar_aviso=False)

    def _construir_ui(self):
        frame_banner = tk.Frame(self.root, bg="#0F172A", padx=20, pady=12)
        frame_banner.pack(fill="x", side="top")

        tk.Label(
            frame_banner,
            text="PANEL INTERACTIVO DE SUPUESTOS Y SIMULACIÓN EN VIVO (PAUTA 10 PUNTOS ULAGOS)",
            font=("Arial", 12, "bold"), fg="#FFFFFF", bg="#0F172A"
        ).pack(anchor="w", pady=(0, 6))

        self.lbl_kpi_priv = tk.Label(
            frame_banner,
            text="> PRIVADO (12%):  VAN = $0   |   TIR = 0.00%   |   Payback = 0.00 años",
            font=("Consolas", 10, "bold"), fg="#38BDF8", bg="#0F172A"
        )
        self.lbl_kpi_priv.pack(anchor="w")

        self.lbl_kpi_soc = tk.Label(
            frame_banner,
            text="> SOCIAL  (6%):   VANS = $0   |   TIRS = 0.00%   |   Payback = 0.00 años",
            font=("Consolas", 10, "bold"), fg="#4ADE80", bg="#0F172A"
        )
        self.lbl_kpi_soc.pack(anchor="w", pady=(2, 0))

        frame_tools = tk.Frame(self.root, bg="#E2E8F0", padx=18, pady=6)
        frame_tools.pack(fill="x", side="top")

        tk.Label(frame_tools, text="Carpeta Destino:", font=("Arial", 9, "bold"), bg="#E2E8F0", fg="#1E293B").pack(side="left")
        ttk.Entry(frame_tools, textvariable=self.carpeta_salida, width=48).pack(side="left", padx=6)

        tk.Button(
            frame_tools, text="Cambiar...", command=self._elegir_carpeta,
            bg="#334155", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=8, pady=2
        ).pack(side="left", padx=2)

        tk.Button(
            frame_tools, text="Abrir Carpeta", command=lambda: mod1.abrir_carpeta_en_pantalla(self.carpeta_salida.get()),
            bg="#0D9488", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=8, pady=2
        ).pack(side="left", padx=4)

        tk.Button(
            frame_tools, text="Cargar Precios Sombra SNI 2026 (TSD 5,5%)", command=self._cargar_preset_sni_2026,
            bg="#EA580C", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=8, pady=2
        ).pack(side="right", padx=4)

        tk.Button(
            frame_tools, text="Cargar Caso Base Original (TSD 6,0%)", command=self._cargar_preset_base,
            bg="#475569", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=8, pady=2
        ).pack(side="right", padx=4)

        frame_bottom = tk.Frame(self.root, bg="#CBD5E1", padx=16, pady=10)
        frame_bottom.pack(fill="x", side="bottom")

        btn_recalc = tk.Button(
            frame_bottom,
            text="🔄  RECALCULAR VAN / TIR EN VIVO",
            command=lambda: self.recalcular_en_vivo(mostrar_aviso=True),
            bg="#0284C7", fg="white", font=("Arial", 10, "bold"),
            relief="flat", pady=8, cursor="hand2"
        )
        btn_recalc.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_generar = tk.Button(
            frame_bottom,
            text="🚀  ACTUALIZAR EXCEL + WORD + PDF + ZIP (PAUTA 10 PUNTOS)",
            command=self.accion_actualizar_todo,
            bg="#15803D", fg="white", font=("Arial", 10, "bold"),
            relief="flat", pady=8, cursor="hand2"
        )
        btn_generar.pack(side="left", fill="x", expand=True, padx=(8, 0))

        contenedor_scroll = tk.Frame(self.root, bg="#F8FAFC")
        contenedor_scroll.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(contenedor_scroll, bg="#F8FAFC", highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor_scroll, orient="vertical", command=self.canvas.yview)
        self.frame_form = tk.Frame(self.canvas, bg="#F8FAFC", padx=22, pady=10)

        self.frame_form.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_win = self.canvas.create_window((0, 0), window=self.frame_form, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_win, width=e.width)
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.frame_form.grid_columnconfigure(0, minsize=290, weight=0)
        self.frame_form.grid_columnconfigure(1, minsize=210, weight=0)
        self.frame_form.grid_columnconfigure(2, weight=1)

        self.fila_actual = 0
        self._poblar_formulario_completo()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _agregar_encabezado_seccion(self, titulo):
        f_hdr = tk.Frame(self.frame_form, bg="#DBEAFE", padx=10, pady=5)
        f_hdr.grid(row=self.fila_actual, column=0, columnspan=3, sticky="ew", pady=(14, 6))
        tk.Label(
            f_hdr, text=titulo, font=("Arial", 10, "bold"),
            fg="#1E3A8A", bg="#DBEAFE"
        ).pack(anchor="w")
        self.fila_actual += 1

    def _agregar_fila_parametro(self, clave, etiqueta, explicacion):
        tk.Label(
            self.frame_form, text=etiqueta, font=("Arial", 9, "bold"),
            fg="#0F172A", bg="#F8FAFC", anchor="w"
        ).grid(row=self.fila_actual, column=0, sticky="w", padx=(4, 10), pady=4)

        var = tk.StringVar(value=str(self.cfg[clave]))
        self.vars_simples[clave] = var
        ent = ttk.Entry(self.frame_form, textvariable=var, width=28, font=("Arial", 9))
        ent.grid(row=self.fila_actual, column=1, sticky="w", padx=(0, 14), pady=4)

        tk.Label(
            self.frame_form, text=explicacion, font=("Arial", 9),
            fg="#475569", bg="#F8FAFC", anchor="w", justify="left", wraplength=540
        ).grid(row=self.fila_actual, column=2, sticky="w", pady=4)

        self.fila_actual += 1

    def _poblar_formulario_completo(self):
        self._agregar_encabezado_seccion("1. PARÁMETROS GENERALES Y TASAS DE DESCUENTO")
        self._agregar_fila_parametro("NOMBRE_PROYECTO", "Nombre del Proyecto:", "(Nombre oficial que saldrá en la portada de la ULagos y en las conclusiones automáticas del Word)")
        self._agregar_fila_parametro("APELLIDOS_ENTREGA", "Apellidos Carpeta y ZIP:", "(Nombre exacto para la carpeta, el Word, el PDF y el archivo comprimido ZIP de entrega)")
        self._agregar_fila_parametro("INTEGRANTES", "Integrantes del Grupo:", "(Nombres completos en mayúsculas separados por coma para el bloque de la portada)")
        self._agregar_fila_parametro("DOCENTE", "Nombre Docente:", "(Nombre de la profesora guía del módulo Formulación y Evaluación de Proyectos)")
        self._agregar_fila_parametro("HORIZONTE_ANIOS", "Horizonte de Evaluación:", "(Cantidad de años a proyectar en las tablas. La pauta exige evaluar a 10 años -> deja 10)")
        self._agregar_fila_parametro("IVA", "Tasa de IVA (Impuesto al Valor Agregado):", "(En Chile el IVA legal es 19% -> escribe 0.19. El programa lo usa para descontar el IVA automáticamente)")
        self._agregar_fila_parametro("IMPUESTO_RENTA", "Impuesto a la Renta (1ª Categoría SII):", "(Impuesto sobre las ganancias de la empresa: usa 0.25 para Pyme [25%]. Solo se cobra en el Flujo Privado)")
        self._agregar_fila_parametro("TASA_DESCUENTO_PRIVADA", "Tasa de Descuento Privada (Costo de Capital):", "(Rentabilidad mínima anual que exige el dueño/inversionista para calcular el VAN Privado: ej. 0.12 = 12%)")
        self._agregar_fila_parametro("TASA_DESCUENTO_SOCIAL", "Tasa Social de Descuento (MIDESO / SNI):", "(Tasa exigida por el Estado para calcular el VANS: usa 0.06 [6%] o 0.055 [5,5% oficial vigente SNI 2026])")
        self._agregar_fila_parametro("DIAS_DESFASE_CT", "Días de Desfase para Capital de Trabajo:", "(Días que la planta debe operar pagando sueldos e insumos antes de cobrar las primeras facturas. Ej: 45 días)")

        self._agregar_encabezado_seccion("2. FACTORES DE CORRECCIÓN SOCIAL (PRECIOS SOMBRA SNI / MIDESO - PUNTO 8)")
        self._agregar_fila_parametro("FC_MOC", "Factor Mano de Obra Calificada (MOC):", "(Para profesionales/jefes. Usa 1.00 en caso base o 0.97 según Precios Sociales SNI 2026)")
        self._agregar_fila_parametro("FC_MOSC", "Factor Mano de Obra Semicalificada (MOSC):", "(Para técnicos/operadores de máquinas. Usa 0.85 en caso base o 0.95 según Precios Sociales SNI 2026)")
        self._agregar_fila_parametro("FC_MONC", "Factor Mano de Obra No Calificada (MONC):", "(Para recolectores/operarios. Usa 0.68 en caso base o 0.91 según Precios Sociales SNI 2026)")
        self._agregar_fila_parametro("FC_DIVISA", "Factor Social de la Divisa (Dólar/Importados):", "(Premio social a la divisa de MIDESO = 1.01 [o 1.00 en SNI 2026]. Ajusta máquinas compradas en el extranjero)")
        self._agregar_fila_parametro("ARANCEL_IMP", "Arancel Aduanero de Importación:", "(Impuesto de aduana en Chile = 6% -> escribe 0.06. En el Flujo Social el programa lo descuenta)")

        self._agregar_encabezado_seccion("3. DEMANDA, PRODUCTOS E INGRESOS (PUNTOS 4, 5.a y 5.d)")
        self._agregar_fila_parametro("RESIDUO_TON_ANIO_1", "Cartón Recolectado el Año 1 (Toneladas):", "(Toneladas de cartón limpio captadas en calle E. Ramírez el primer año. Ej: 15.0 ton/año = 50 kg/día)")
        self._agregar_fila_parametro("TASA_CRECIMIENTO", "Tasa de Crecimiento Anual (Años 2 al 10):", "(Ritmo al que crecen cada año las ventas y la recolección de cartón. Ej: escribe 0.05 para subir 5% anual)")
        self._agregar_fila_parametro("P1_NOMBRE", "Nombre Producto Línea 1:", "(Nombre comercial de tu primera línea de productos para escuelas DAEM y jardines infantiles)")
        self._agregar_fila_parametro("P1_UNIDADES", "Unidades Vendidas Año 1 - Línea 1:", "(Cantidad de unidades de la Línea 1 vendidas en todo el primer año. Ej: 3000 unidades/año)")
        self._agregar_fila_parametro("P1_PRECIO_BRUTO", "Precio de Venta CON IVA - Línea 1 ($):", "(Precio al que vendes cada unidad con IVA incluido. Ej: 16000 pesos. El programa lo divide por 1.19 solo)")
        self._agregar_fila_parametro("P1_KG_CARTON", "Kg de Cartón por Unidad - Línea 1:", "(Kilos de cartón que se necesitan para fabricar 1 unidad de la Línea 1. Ej: 2.5 kg por set)")
        self._agregar_fila_parametro("P2_NOMBRE", "Nombre Producto Línea 2:", "(Nombre comercial de tu segunda línea enfocada en centros y clubes de adulto mayor)")
        self._agregar_fila_parametro("P2_UNIDADES", "Unidades Vendidas Año 1 - Línea 2:", "(Cantidad de unidades de la Línea 2 vendidas en todo el primer año. Ej: 1500 unidades/año)")
        self._agregar_fila_parametro("P2_PRECIO_BRUTO", "Precio de Venta CON IVA - Línea 2 ($):", "(Precio al que vendes cada unidad con IVA incluido. Ej: 19000 pesos. El programa le quita el 19% de IVA solo)")
        self._agregar_fila_parametro("P2_KG_CARTON", "Kg de Cartón por Unidad - Línea 2:", "(Kilos de cartón necesarios para fabricar 1 unidad de la Línea 2. Ej: 2.0 kg por juego)")

        self._agregar_encabezado_seccion("4. BENEFICIOS SOCIALES EXTERNOS (LIBERACIÓN DE RECURSOS - PUNTO 8)")
        self._agregar_fila_parametro("AHORRO_MUNI_TON", "Ahorro Municipal por Tonelada Desviada ($):", "(Dinero con IVA que ahorra la Municipalidad de Osorno en camiones y Vertedero de Curaco por tonelada: ej. 48000)")
        self._agregar_fila_parametro("OTROS_BEN_SOC_ANUAL", "Otros Beneficios Sociales Anuales ($):", "(Valor anual en pesos de beneficios extra para la comuna. Pon 2500000 o 0 si usas criterio conservador SNI)")

        self._agregar_encabezado_seccion("5. INVERSIONES EN ACTIVOS FIJOS CON IVA (PUNTO 5.b)")
        explicaciones_inv = [
            "(Costo de compra en Año 0 con IVA: ej. 8500000. Es máquina importada, dura 10 años y recupera 10% residual)",
            "(Costo en Año 0 con IVA: ej. 3500000. Es nacional y tiene vida útil de 5 años, por lo que se renueva en el Año 5)",
            "(Costo en Año 0 con IVA: ej. 2000000. Es importado, vida útil de 5 años con renovación en el Año 5)",
            "(Gasto inicial en Año 0 en adecuar el galpón en Ovejería, electricidad trifásica y permiso SEREMI: ej. 2000000)",
        ]
        etiquetas_inv = [
            "Troqueladora Industrial ($ con IVA):",
            "Guillotina, Mesones y Moldes ($ con IVA):",
            "Equipo Serigrafía y Secado ($ con IVA):",
            "Habilitación de Galpón y Permisos ($ con IVA):",
        ]
        for idx, inv in enumerate(self.cfg["INVERSIONES_LISTA"]):
            tk.Label(
                self.frame_form, text=etiquetas_inv[idx], font=("Arial", 9, "bold"),
                fg="#0F172A", bg="#F8FAFC", anchor="w"
            ).grid(row=self.fila_actual, column=0, sticky="w", padx=(4, 10), pady=4)

            v_monto = tk.StringVar(value=str(inv["monto_bruto"]))
            self.vars_inv.append(v_monto)
            ttk.Entry(self.frame_form, textvariable=v_monto, width=28, font=("Arial", 9)).grid(
                row=self.fila_actual, column=1, sticky="w", padx=(0, 14), pady=4
            )
            tk.Label(
                self.frame_form, text=explicaciones_inv[idx], font=("Arial", 9),
                fg="#475569", bg="#F8FAFC", anchor="w", justify="left", wraplength=540
            ).grid(row=self.fila_actual, column=2, sticky="w", pady=4)
            self.fila_actual += 1

        self._agregar_encabezado_seccion("6. COSTOS OPERATIVOS VARIABLES Y FIJOS (PUNTO 5.c)")
        etiquetas_cv = [
            "Costo Unitario Tintas y Barniz ($ c/IVA):",
            "Costo Unitario Caja y Etiqueta ($ c/IVA):",
            "Costo Flete Furgón por Tonelada ($ c/IVA):",
        ]
        explicaciones_cv = [
            "(Gasto variable en pinturas ecológicas al agua EN-71, pegamento y barniz por cada juego fabricado: ej. 1800)",
            "(Gasto variable en caja de presentación, instructivo y etiqueta por cada juego fabricado: ej. 600)",
            "(Tarifa pagada al furgón contratado por cada tonelada retirada en calle Eleuterio Ramírez: ej. 30000)",
        ]
        for idx, cv in enumerate(self.cfg["COSTOS_VAR_LISTA"]):
            tk.Label(
                self.frame_form, text=etiquetas_cv[idx], font=("Arial", 9, "bold"),
                fg="#0F172A", bg="#F8FAFC", anchor="w"
            ).grid(row=self.fila_actual, column=0, sticky="w", padx=(4, 10), pady=4)

            v_cv = tk.StringVar(value=str(cv["costo"]))
            self.vars_cv.append(v_cv)
            ttk.Entry(self.frame_form, textvariable=v_cv, width=28, font=("Arial", 9)).grid(
                row=self.fila_actual, column=1, sticky="w", padx=(0, 14), pady=4
            )
            tk.Label(
                self.frame_form, text=explicaciones_cv[idx], font=("Arial", 9),
                fg="#475569", bg="#F8FAFC", anchor="w", justify="left", wraplength=540
            ).grid(row=self.fila_actual, column=2, sticky="w", pady=4)
            self.fila_actual += 1

        etiquetas_cf = [
            "Arriendo Mensual Galpón Ovejería ($ Exento):",
            "Luz (Saesa), Agua (Suralis) e Internet ($/mes):",
            "Contador, Patente y Seguros ($/mes c/IVA):",
        ]
        explicaciones_cf = [
            "(Arriendo mensual del galpón de 140 m2 en Sector Industrial Ovejería: ej. 450000 pesos al mes)",
            "(Gasto fijo mensual en electricidad trifásica Saesa, agua potable Suralis e internet: ej. 180000 al mes)",
            "(Gasto fijo mensual en contabilidad Pro Pyme, patente municipal en Osorno y seguros: ej. 150000 al mes)",
        ]
        for idx, cf in enumerate(self.cfg["COSTOS_FIJOS_LISTA"]):
            tk.Label(
                self.frame_form, text=etiquetas_cf[idx], font=("Arial", 9, "bold"),
                fg="#0F172A", bg="#F8FAFC", anchor="w"
            ).grid(row=self.fila_actual, column=0, sticky="w", padx=(4, 10), pady=4)

            v_cf = tk.StringVar(value=str(cf["mensual"]))
            self.vars_cf.append(v_cf)
            ttk.Entry(self.frame_form, textvariable=v_cf, width=28, font=("Arial", 9)).grid(
                row=self.fila_actual, column=1, sticky="w", padx=(0, 14), pady=4
            )
            tk.Label(
                self.frame_form, text=explicaciones_cf[idx], font=("Arial", 9),
                fg="#475569", bg="#F8FAFC", anchor="w", justify="left", wraplength=540
            ).grid(row=self.fila_actual, column=2, sticky="w", pady=4)
            self.fila_actual += 1

        self._agregar_encabezado_seccion("7. DOTACIÓN DE PERSONAL Y SUELDOS MENSUALES (PUNTO 4.3)")
        etiquetas_pers = [
            "Sueldo Bruto Jefe de Planta (1 persona - MOC):",
            "Sueldo Bruto Operador Técnico (2 personas - MOSC):",
            "Sueldo Bruto Operario Recolección (2 pers. - MONC):",
        ]
        explicaciones_pers = [
            "(Sueldo mensual bruto del Jefe de Planta y Ventas [1 trabajador MOC]: ej. 950000 pesos/mes)",
            "(Sueldo mensual bruto unitario de cada Operador de Corte y Pintura [son 2 trabajadores MOSC]: ej. 600000 c/u)",
            "(Sueldo mensual bruto unitario de cada Operario de Recolección y Armado [son 2 trabajadores MONC]: ej. 520000 c/u)",
        ]
        for idx, p_item in enumerate(self.cfg["PERSONAL_LISTA"]):
            tk.Label(
                self.frame_form, text=etiquetas_pers[idx], font=("Arial", 9, "bold"),
                fg="#0F172A", bg="#F8FAFC", anchor="w"
            ).grid(row=self.fila_actual, column=0, sticky="w", padx=(4, 10), pady=4)

            v_p = tk.StringVar(value=str(p_item["sueldo_bruto_mensual"]))
            self.vars_pers.append(v_p)
            ttk.Entry(self.frame_form, textvariable=v_p, width=28, font=("Arial", 9)).grid(
                row=self.fila_actual, column=1, sticky="w", padx=(0, 14), pady=4
            )
            tk.Label(
                self.frame_form, text=explicaciones_pers[idx], font=("Arial", 9),
                fg="#475569", bg="#F8FAFC", anchor="w", justify="left", wraplength=540
            ).grid(row=self.fila_actual, column=2, sticky="w", pady=4)
            self.fila_actual += 1

    def _cargar_preset_sni_2026(self):
        self.vars_simples["TASA_DESCUENTO_SOCIAL"].set("0.055")
        self.vars_simples["FC_MOC"].set("0.97")
        self.vars_simples["FC_MOSC"].set("0.95")
        self.vars_simples["FC_MONC"].set("0.91")
        self.vars_simples["FC_DIVISA"].set("1.00")
        self.vars_simples["OTROS_BEN_SOC_ANUAL"].set("0")
        self.recalcular_en_vivo(mostrar_aviso=True)

    def _cargar_preset_base(self):
        self.vars_simples["TASA_DESCUENTO_SOCIAL"].set("0.06")
        self.vars_simples["FC_MOC"].set("1.00")
        self.vars_simples["FC_MOSC"].set("0.85")
        self.vars_simples["FC_MONC"].set("0.68")
        self.vars_simples["FC_DIVISA"].set("1.01")
        self.vars_simples["OTROS_BEN_SOC_ANUAL"].set("2500000")
        self.recalcular_en_vivo(mostrar_aviso=True)

    def _elegir_carpeta(self):
        c = filedialog.askdirectory(initialdir=self.carpeta_salida.get(), title="Seleccionar carpeta de entrega")
        if c:
            self.carpeta_salida.set(os.path.normpath(c))

    def _leer_todo_el_formulario(self):
        campos_texto = {"NOMBRE_PROYECTO", "APELLIDOS_ENTREGA", "INTEGRANTES", "DOCENTE", "P1_NOMBRE", "P2_NOMBRE"}
        campos_enteros = {"HORIZONTE_ANIOS", "DIAS_DESFASE_CT", "P1_UNIDADES", "P2_UNIDADES"}

        for k, var in self.vars_simples.items():
            val_str = var.get().strip()
            if k in campos_texto:
                self.cfg[k] = val_str
            elif k in campos_enteros:
                self.cfg[k] = int(float(val_str.replace(",", ".")))
            else:
                self.cfg[k] = float(val_str.replace(",", "."))

        for idx, v_inv in enumerate(self.vars_inv):
            self.cfg["INVERSIONES_LISTA"][idx]["monto_bruto"] = float(v_inv.get().strip().replace(",", "."))

        for idx, v_cv in enumerate(self.vars_cv):
            self.cfg["COSTOS_VAR_LISTA"][idx]["costo"] = float(v_cv.get().strip().replace(",", "."))

        for idx, v_cf in enumerate(self.vars_cf):
            self.cfg["COSTOS_FIJOS_LISTA"][idx]["mensual"] = float(v_cf.get().strip().replace(",", "."))

        for idx, v_p in enumerate(self.vars_pers):
            self.cfg["PERSONAL_LISTA"][idx]["sueldo_bruto_mensual"] = float(v_p.get().strip().replace(",", "."))

    def recalcular_en_vivo(self, mostrar_aviso=False):
        try:
            self._leer_todo_el_formulario()
            res = mod2.ejecutar_motor_financiero(self.cfg)
            m = res["metricas"]
            tp = self.cfg["TASA_DESCUENTO_PRIVADA"] * 100
            ts = self.cfg["TASA_DESCUENTO_SOCIAL"] * 100

            self.lbl_kpi_priv.config(
                text=f"> PRIVADO ({tp:.1f}%):  VAN = {mod2.fmt_clp(m['van_privado'])}   |   TIR = {m['tir_privada']:.2f}%   |   Payback = {m['payback_privado']} años"
            )
            self.lbl_kpi_soc.config(
                text=f"> SOCIAL  ({ts:.1f}%):  VANS = {mod2.fmt_clp(m['van_social'])}   |   TIRS = {m['tir_social']:.2f}%   |   Payback = {m['payback_social']} años"
            )
            if mostrar_aviso:
                messagebox.showinfo(
                    "Simulación en Vivo Actualizada",
                    f"Indicadores recalculados correctamente:\n\n"
                    f"• VAN Privado ({tp:.1f}%): {mod2.fmt_clp(m['van_privado'])} (TIR: {m['tir_privada']:.2f}%)\n"
                    f"• VANS Social ({ts:.1f}%): {mod2.fmt_clp(m['van_social'])} (TIRS: {m['tir_social']:.2f}%)"
                )
        except Exception as e:
            messagebox.showerror("Error en Formato Numérico", f"Revisa que las casillas numéricas no tengan letras:\n\n{e}")

    def accion_actualizar_todo(self):
        try:
            self.recalcular_en_vivo(mostrar_aviso=False)
            carpeta_destino = self.carpeta_salida.get().strip()
            info = ejecutar_pipeline_completo_1_2_3(self.cfg, carpeta_destino)
            mod1.abrir_carpeta_en_pantalla(info["carpeta_entrega"])

            messagebox.showinfo(
                "¡Entrega Completa Actualizada (Scripts 1 + 2 + 3)!",
                f"Se actualizaron todos los archivos con tus nuevos supuestos:\n\n"
                f"📦 ZIP: {info['ruta_zip']}\n"
                f"📄 Word: {os.path.basename(info['ruta_word'])} (Índice en Pág. 2 reparado + Árboles + CANVAS + MML)\n"
                f"📕 PDF: {os.path.basename(info['ruta_pdf'])}\n"
                f"📊 3 Libros Excel con fórmulas vivas actualizados."
            )
        except Exception as e:
            messagebox.showerror("Error al Generar Archivos", f"Cierra Word y Excel antes de actualizar:\n\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ConsolaMaestraSupuestos(root)
    root.lift()
    root.focus_force()
    root.mainloop()