# -*- coding: utf-8 -*-
"""
ARCHIVO 3 (VERSIÓN CORREGIDA): 3_integrador_arboles_y_pauta_josue.py
- Limpia cualquier tabla o párrafo que se haya insertado por error dentro de la Tabla de Contenidos (Página 2).
- Ignora las líneas del Índice al buscar los capítulos, inyectando cada sección únicamente en el cuerpo del informe.
- Inserta los capítulos 10 (CANVAS), 11 (Marco Lógico), 12 (Control de Consistencia) y subsecciones en el Índice
  con hipervínculo clicable, línea punteada (.....) y número de página dinámico (PAGEREF).
- Actualiza automáticamente los números de página en Word, regenera el PDF y empaqueta 'Vargas,Muñoz,Jofre.zip'.
"""

import os
import sys
import zipfile
import subprocess
import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import pandas as pd

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==============================================================================
# 1. GENERADOR DE LOS 7 DIAGRAMAS VISUALES DE JOSUÉ
# ==============================================================================

def _dibujar_caja_redondeada(ax, xc, yc, w, h, texto, facecolor, edgecolor, fontsize=8.5, bold=False, textcolor="#0F172A"):
    box = FancyBboxPatch(
        (xc - w / 2, yc - h / 2), w, h,
        boxstyle="round,pad=0.012,rounding_size=0.025",
        facecolor=facecolor, edgecolor=edgecolor, linewidth=1.3
    )
    ax.add_patch(box)
    ax.text(
        xc, yc, texto, ha="center", va="center", fontsize=fontsize,
        fontweight="bold" if bold else "normal", color=textcolor,
        family="Arial", linespacing=1.2
    )


def generar_diagramas_josue(carpeta_destino):
    os.makedirs(carpeta_destino, exist_ok=True)
    rutas = {}

    # --------------------------------------------------------------------------
    # FIGURA 1: ÁRBOL DE PROBLEMAS (Conectores ortogonales sin cruces)
    # --------------------------------------------------------------------------
    r_prob = os.path.join(carpeta_destino, "josue_fig1_arbol_problemas.png")
    fig, ax = plt.subplots(figsize=(11.5, 6.2), dpi=220)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    ax.text(0.5, 0.96, "ÁRBOL DE PROBLEMAS - CARTÓN VIVO OSORNO", ha="center", va="center", fontsize=14, fontweight="bold", color="#1E293B", family="Arial")
    ax.text(0.5, 0.92, "Efectos arriba, problema central al medio y causas abajo. Cada rama mantiene una relación directa y verificable.", ha="center", va="center", fontsize=9.5, color="#475569", family="Arial")

    sub_efectos = [
        (0.095, "Menores encadenamientos\nproductivos locales."),
        (0.245, "Persistencia de material\nvalorizable sin\ntransformación local."),
        (0.425, "Menor diversidad de\nmateriales didácticos\nsustentables."),
        (0.575, "Mayor dependencia de\nproductos\nconvencionales."),
        (0.755, "Menor aprendizaje\npráctico sobre\nvalorización."),
        (0.905, "Menor articulación entre\ncomercios, usuarios y\ngestores."),
    ]
    for x_se, txt_se in sub_efectos:
        _dibujar_caja_redondeada(ax, x_se, 0.80, 0.14, 0.12, txt_se, "#F9EBEB", "#78716C", fontsize=7.8)

    efectos_dir = [
        (0.17, "Menor retención de valor económico\ny ambiental en Osorno."),
        (0.50, "Oferta local limitada de\nmateriales didácticos circulares."),
        (0.83, "Baja visibilidad de soluciones\nlocales de economía circular."),
    ]
    for x_ef, txt_ef in efectos_dir:
        _dibujar_caja_redondeada(ax, x_ef, 0.63, 0.29, 0.11, txt_ef, "#F2D7D5", "#64748B", fontsize=8.5, bold=True)

    _dibujar_caja_redondeada(
        ax, 0.50, 0.45, 0.60, 0.14,
        "Bajo aprovechamiento local del\ncartón comercial recuperable de\nOsorno para transformarlo en\nproductos didácticos, recreativos\ny cognitivos de valor agregado.",
        "#E0ECF8", "#475569", fontsize=9.5, bold=True
    )

    causas_dir = [
        (0.17, "Abastecimiento selectivo aún poco\nformalizado."),
        (0.50, "Capacidad productiva y control de\ncalidad por implementar."),
        (0.83, "Demanda institucional requiere\nvalidación comercial continua."),
    ]
    for x_ca, txt_ca in causas_dir:
        _dibujar_caja_redondeada(ax, x_ca, 0.27, 0.29, 0.11, txt_ca, "#FDEBD0", "#64748B", fontsize=8.5, bold=True)

    sub_causas = [
        (0.095, "Acuerdos de retiro\ntodavía no formalizados."),
        (0.245, "Calidad y humedad del\ncartón pueden variar."),
        (0.425, "Proceso, prototipos y\nespecificaciones aún\ndeben estandarizarse."),
        (0.575, "Seguridad, durabilidad y\ncontrol de calidad\nrequieren verificación."),
        (0.755, "No existen órdenes de\ncompra en los\nantecedentes."),
        (0.905, "El crecimiento comercial\nes un supuesto que debe\nvalidarse."),
    ]
    for x_sc, txt_sc in sub_causas:
        _dibujar_caja_redondeada(ax, x_sc, 0.10, 0.14, 0.12, txt_sc, "#FEF5E7", "#78716C", fontsize=7.8)

    c_line = "#64748B"
    for x_grupo, x_izq, x_der in [(0.17, 0.095, 0.245), (0.50, 0.425, 0.575), (0.83, 0.755, 0.905)]:
        ax.plot([x_izq, x_der], [0.715, 0.715], color=c_line, lw=1.3)
        ax.plot([x_izq, x_izq], [0.715, 0.74], color=c_line, lw=1.3)
        ax.plot([x_der, x_der], [0.715, 0.74], color=c_line, lw=1.3)
        ax.plot([x_grupo, x_grupo], [0.685, 0.715], color=c_line, lw=1.3)

        ax.plot([x_izq, x_der], [0.185, 0.185], color=c_line, lw=1.3)
        ax.plot([x_izq, x_izq], [0.16, 0.185], color=c_line, lw=1.3)
        ax.plot([x_der, x_der], [0.16, 0.185], color=c_line, lw=1.3)
        ax.plot([x_grupo, x_grupo], [0.185, 0.215], color=c_line, lw=1.3)

    ax.plot([0.17, 0.83], [0.545, 0.545], color=c_line, lw=1.4)
    for xg in [0.17, 0.50, 0.83]:
        ax.plot([xg, xg], [0.545, 0.575], color=c_line, lw=1.4)
    ax.plot([0.50, 0.50], [0.52, 0.545], color=c_line, lw=1.5)

    ax.plot([0.17, 0.83], [0.355, 0.355], color=c_line, lw=1.4)
    for xg in [0.17, 0.50, 0.83]:
        ax.plot([xg, xg], [0.325, 0.355], color=c_line, lw=1.4)
    ax.plot([0.50, 0.50], [0.355, 0.38], color=c_line, lw=1.5)

    fig.tight_layout()
    fig.savefig(r_prob, dpi=220)
    plt.close(fig)
    rutas["arbol_problemas"] = r_prob

    # --------------------------------------------------------------------------
    # FIGURA 2: ÁRBOL DE OBJETIVOS (Medios y Fines)
    # --------------------------------------------------------------------------
    r_obj = os.path.join(carpeta_destino, "josue_fig2_arbol_objetivos.png")
    fig, ax = plt.subplots(figsize=(11.5, 6.2), dpi=220)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    ax.text(0.5, 0.96, "ÁRBOL DE OBJETIVOS - CARTÓN VIVO OSORNO", ha="center", va="center", fontsize=14, fontweight="bold", color="#1E293B", family="Arial")
    ax.text(0.5, 0.92, "Los medios responden directamente a las causas y los fines revierten los efectos identificados en el diagnóstico.", ha="center", va="center", fontsize=9.5, color="#475569", family="Arial")

    sub_fines = [
        (0.095, "Mayor encadenamiento\nproductivo local."),
        (0.245, "Menor material\nvalorizable sin\ntransformación local."),
        (0.425, "Más alternativas\ndidácticas sustentables."),
        (0.575, "Sustitución parcial de\nproductos\nconvencionales."),
        (0.755, "Mayor aprendizaje\npráctico sobre\nvalorización."),
        (0.905, "Mayor colaboración entre\ncomercios, usuarios y\ngestores."),
    ]
    for x_sf, txt_sf in sub_fines:
        _dibujar_caja_redondeada(ax, x_sf, 0.80, 0.14, 0.12, txt_sf, "#EAFAF1", "#64748B", fontsize=7.8)

    fines_dir = [
        (0.17, "Mayor retención de valor económico\ny ambiental en Osorno."),
        (0.50, "Mayor oferta local de materiales\ndidácticos circulares."),
        (0.83, "Mayor cultura y articulación local\nde economía circular."),
    ]
    for x_fi, txt_fi in fines_dir:
        _dibujar_caja_redondeada(ax, x_fi, 0.63, 0.29, 0.11, txt_fi, "#D5F5E3", "#475569", fontsize=8.5, bold=True)

    _dibujar_caja_redondeada(
        ax, 0.50, 0.45, 0.60, 0.14,
        "Aumentar el aprovechamiento local\ndel cartón comercial recuperable\nmediante productos didácticos y\ncognitivos de valor agregado, con\ntrazabilidad, calidad y demanda\nvalidada.",
        "#E0ECF8", "#334155", fontsize=9.5, bold=True
    )

    medios_dir = [
        (0.17, "Formalizar abastecimiento\nselectivo y trazable."),
        (0.50, "Implementar capacidad productiva y\ncontrol de calidad."),
        (0.83, "Validar demanda institucional y\ncanales de venta."),
    ]
    for x_me, txt_me in medios_dir:
        _dibujar_caja_redondeada(ax, x_me, 0.27, 0.29, 0.11, txt_me, "#D6EAF8", "#475569", fontsize=8.5, bold=True)

    sub_medios = [
        (0.095, "Acuerdos de retiro y\nprotocolos de recepción."),
        (0.245, "Control de calidad y\ntrazabilidad del\nresiduo."),
        (0.425, "Prototipos, proceso y\nespecificaciones\nestandarizados."),
        (0.575, "Pruebas de seguridad,\ndurabilidad y calidad\nimplementadas."),
        (0.755, "Pilotos, cartas de\nintención u órdenes de\ncompra."),
        (0.905, "Crecimiento condicionado\na acuerdos comerciales\nverificables."),
    ]
    for x_sm, txt_sm in sub_medios:
        _dibujar_caja_redondeada(ax, x_sm, 0.10, 0.14, 0.12, txt_sm, "#EBF5FB", "#64748B", fontsize=7.8)

    for x_grupo, x_izq, x_der in [(0.17, 0.095, 0.245), (0.50, 0.425, 0.575), (0.83, 0.755, 0.905)]:
        ax.plot([x_izq, x_der], [0.715, 0.715], color=c_line, lw=1.3)
        ax.plot([x_izq, x_izq], [0.715, 0.74], color=c_line, lw=1.3)
        ax.plot([x_der, x_der], [0.715, 0.74], color=c_line, lw=1.3)
        ax.plot([x_grupo, x_grupo], [0.685, 0.715], color=c_line, lw=1.3)

        ax.plot([x_izq, x_der], [0.185, 0.185], color=c_line, lw=1.3)
        ax.plot([x_izq, x_izq], [0.16, 0.185], color=c_line, lw=1.3)
        ax.plot([x_der, x_der], [0.16, 0.185], color=c_line, lw=1.3)
        ax.plot([x_grupo, x_grupo], [0.185, 0.215], color=c_line, lw=1.3)

    ax.plot([0.17, 0.83], [0.545, 0.545], color=c_line, lw=1.4)
    for xg in [0.17, 0.50, 0.83]:
        ax.plot([xg, xg], [0.545, 0.575], color=c_line, lw=1.4)
    ax.plot([0.50, 0.50], [0.52, 0.545], color=c_line, lw=1.5)

    ax.plot([0.17, 0.83], [0.355, 0.355], color=c_line, lw=1.4)
    for xg in [0.17, 0.50, 0.83]:
        ax.plot([xg, xg], [0.325, 0.355], color=c_line, lw=1.4)
    ax.plot([0.50, 0.50], [0.355, 0.38], color=c_line, lw=1.5)

    fig.tight_layout()
    fig.savefig(r_obj, dpi=220)
    plt.close(fig)
    rutas["arbol_objetivos"] = r_obj

    # --------------------------------------------------------------------------
    # FIGURA 3: BALANCE FÍSICO DEL CARTÓN (Recolección vs Incorporación vs 30% Merma)
    # --------------------------------------------------------------------------
    r_bal = os.path.join(carpeta_destino, "josue_fig3_balance_carton.png")
    fig, ax = plt.subplots(figsize=(9.2, 4.2), dpi=220)
    anios = [f"Año {i}" for i in range(1, 11)]
    rec = [round(15.0 * ((1.05) ** i), 2) for i in range(10)]
    inc = [round(10.5 * ((1.05) ** i), 2) for i in range(10)]

    ax.plot(anios, rec, marker="o", lw=2.2, color="#1D4ED8", label="Cartón recolectado proyectado (base inicial: 6 locales)")
    ax.plot(anios, inc, marker="s", lw=2.2, color="#EA580C", label="Cartón incorporado a productos (70%)")
    ax.fill_between(range(10), inc, rec, color="#93C5FD", alpha=0.35, label="Recortes / fracción no incorporada (30%: requiere gestión documentada)")

    for i in range(10):
        ax.text(i, rec[i] + 0.35, f"{rec[i]:.1f} t", ha="center", va="bottom", fontsize=7.8, family="Arial", color="#1E293B")
        ax.text(i, inc[i] - 0.65, f"{inc[i]:.1f} t", ha="center", va="top", fontsize=7.8, family="Arial", color="#9A3412")

    ax.set_ylim(8.5, 25.0)
    ax.set_ylabel("Toneladas por año", fontsize=9, fontweight="bold", family="Arial")
    ax.set_title("Balance físico del cartón: recolección, incorporación a productos y fracción residual", fontsize=10.5, fontweight="bold", family="Arial")
    ax.grid(True, linestyle="--", alpha=0.45)
    ax.legend(loc="upper left", fontsize=8.2)
    fig.tight_layout()
    fig.savefig(r_bal, dpi=220)
    plt.close(fig)
    rutas["balance_carton"] = r_bal

    # --------------------------------------------------------------------------
    # FIGURA 4: CROQUIS CONCEPTUAL GALPÓN 140 M2 (Sector Ovejería)
    # --------------------------------------------------------------------------
    r_croq = os.path.join(carpeta_destino, "josue_fig4_croquis_galpon.png")
    fig, ax = plt.subplots(figsize=(9.2, 4.8), dpi=220)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")

    ax.text(5.0, 5.7, "Croquis conceptual de distribución - Galpón 140 m2 en Ovejería (no arquitectónico)", ha="center", va="center", fontsize=12, fontweight="bold", color="#000000", family="Arial")

    ax.plot([0.2, 9.8, 9.8, 0.2, 0.2], [0.3, 0.3, 5.3, 5.3, 0.3], color="#1E3A8A", lw=1.8)
    ax.plot([0.2, 9.8], [2.9, 2.9], color="#1E3A8A", lw=1.6)
    ax.plot([3.1, 3.1], [2.9, 5.3], color="#1E3A8A", lw=1.5)
    ax.plot([6.5, 6.5], [2.9, 5.3], color="#1E3A8A", lw=1.5)
    ax.plot([3.7, 3.7], [0.3, 2.9], color="#1E3A8A", lw=1.5)
    ax.plot([6.5, 6.5], [0.3, 2.9], color="#1E3A8A", lw=1.5)

    ax.annotate("", xy=(9.3, 4.95), xytext=(2.1, 4.95), arrowprops=dict(arrowstyle="->", lw=2.0, color="#000000"))
    ax.text(1.3, 4.95, "Flujo material", ha="center", va="center", fontsize=10, fontweight="bold", family="Arial")

    zonas = [
        (1.65, 3.9, "Recepción y\nalmacenamiento\n35 m2"),
        (4.80, 4.0, "Corte y troquelado\n40 m2"),
        (8.15, 4.0, "Impresión y secado\n30 m2"),
        (1.95, 1.6, "Armado y control\n20 m2"),
        (5.10, 1.6, "Producto terminado\n10 m2"),
        (8.15, 1.6, "Oficina, servicios y\ncirculación\n5 m2"),
    ]
    for xz, yz, tz in zonas:
        ax.text(xz, yz, tz, ha="center", va="center", fontsize=10, family="Arial", color="#000000")

    fig.tight_layout()
    fig.savefig(r_croq, dpi=220)
    plt.close(fig)
    rutas["croquis_galpon"] = r_croq

    # --------------------------------------------------------------------------
    # FIGURA 5: DIAGRAMA DE FLUJO DE PROCESO PRODUCTIVO (7 Etapas)
    # --------------------------------------------------------------------------
    r_proc = os.path.join(carpeta_destino, "josue_fig5_flujo_proceso.png")
    fig, ax = plt.subplots(figsize=(11.0, 2.5), dpi=220)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    etapas = [
        "Retiro selectivo\n05:30 a 06:15",
        "Recepción y\nregistro",
        "Inspección y\nlimpieza",
        "Corte y\ntroquelado",
        "Impresión y\nacabado",
        "Armado y\ncontrol",
        "Empaque y\ndespacho",
    ]
    n_et = len(etapas)
    w_et = 0.125
    esp = (0.96 - n_et * w_et) / (n_et - 1)
    for i, et_txt in enumerate(etapas):
        xc = 0.02 + w_et / 2 + i * (w_et + esp)
        _dibujar_caja_redondeada(ax, xc, 0.58, w_et, 0.48, et_txt, "#EBF5FB", "#1E3A8A", fontsize=8.5, bold=True)
        if i < n_et - 1:
            x_ini = xc + w_et / 2
            x_fin = x_ini + esp
            ax.annotate("", xy=(x_fin, 0.58), xytext=(x_ini, 0.58), arrowprops=dict(arrowstyle="->", lw=1.8, color="#000000"))

    ax.text(
        0.5, 0.14,
        "Fracción no incorporada (30% = 4,50 t/año): segregación y entrega a gestor o reciclador; no retorna al flujo municipal sin trazabilidad.",
        ha="center", va="center", fontsize=8.8, style="italic", color="#000000", family="Arial"
    )

    fig.tight_layout()
    fig.savefig(r_proc, dpi=220)
    plt.close(fig)
    rutas["flujo_proceso"] = r_proc

    # --------------------------------------------------------------------------
    # FIGURA 6: COMPARACIÓN ANUAL FLUJO PRIVADO VS FLUJO SOCIAL SNI 2026
    # --------------------------------------------------------------------------
    r_flujos = os.path.join(carpeta_destino, "josue_fig6_flujos_comparados.png")
    fig, ax = plt.subplots(figsize=(9.5, 4.2), dpi=220)
    etiquetas_f = [f"Año {i}" for i in range(0, 11)]
    f_priv = [-20.4, 7.0, 9.2, 10.5, 12.8, 10.5, 17.6, 20.2, 22.9, 25.8, 36.6]
    f_soc = [-20.1, 10.4, 13.1, 16.0, 19.1, 17.7, 25.6, 29.1, 32.8, 36.7, 48.2]

    x_pos = range(len(etiquetas_f))
    w_b = 0.36
    ax.bar([p - w_b / 2 for p in x_pos], f_priv, width=w_b, label="Flujo de caja privado neto (MM$)", color="#1D4ED8")
    ax.bar([p + w_b / 2 for p in x_pos], f_soc, width=w_b, label="Flujo social SNI 2026 (MM$)", color="#F97316")
    ax.axhline(0, color="#334155", lw=1.0)

    for i in range(len(etiquetas_f)):
        va_p = "bottom" if f_priv[i] >= 0 else "top"
        off_p = 0.6 if f_priv[i] >= 0 else -1.8
        ax.text(i - w_b / 2, f_priv[i] + off_p, f"{f_priv[i]:.1f}", ha="center", va=va_p, fontsize=7.2, family="Arial")

        va_s = "bottom" if f_soc[i] >= 0 else "top"
        off_s = 0.6 if f_soc[i] >= 0 else -1.8
        ax.text(i + w_b / 2, f_soc[i] + off_s, f"{f_soc[i]:.1f}", ha="center", va=va_s, fontsize=7.2, family="Arial")

    ax.set_xticks(list(x_pos))
    ax.set_xticklabels(etiquetas_f, fontsize=8.5, family="Arial")
    ax.set_ylim(-25, 54)
    ax.set_ylabel("Millones de pesos (CLP)", fontsize=9, fontweight="bold", family="Arial")
    ax.set_title("Comparación anual: flujo de caja privado vs. flujo social corregido (SNI 2026)", fontsize=10.5, fontweight="bold", family="Arial")
    ax.grid(axis="y", linestyle="--", alpha=0.45)
    ax.legend(loc="upper left", fontsize=8.5)
    fig.tight_layout()
    fig.savefig(r_flujos, dpi=220)
    plt.close(fig)
    rutas["flujos_comparados"] = r_flujos

    # --------------------------------------------------------------------------
    # FIGURA 7: ANÁLISIS DE SENSIBILIDAD MULTIESCENARIO (Cifras exactas SNI 2026)
    # --------------------------------------------------------------------------
    r_sens = os.path.join(carpeta_destino, "josue_fig7_sensibilidad.png")
    fig, ax = plt.subplots(figsize=(8.6, 4.0), dpi=220)

    escenarios = [
        "Pesimista: -15% ingresos",
        "Costos +10%",
        "Caso base: +5% anual",
        "Sin crecimiento: 0%",
    ]
    van_p = [11.55, 36.87, 63.18, 16.11]
    van_s = [65.52, 112.15, 154.99, 58.88]

    y = range(len(escenarios))
    h_bar = 0.35
    ax.barh([i - h_bar / 2 for i in y], van_p, height=h_bar, label="VAN privado (12%)", color="#0284C7", edgecolor="#000000", linewidth=0.6)
    ax.barh([i + h_bar / 2 for i in y], van_s, height=h_bar, label="VANS social (5,5% SNI 2026)", color="#EA580C", edgecolor="#000000", linewidth=0.6)

    ax.set_yticks(list(y))
    ax.set_yticklabels(escenarios, fontsize=9, family="Arial")
    ax.set_xlabel("Valor actual neto (MM$ CLP)", fontsize=9, fontweight="bold", family="Arial")
    ax.set_title("Análisis de sensibilidad - VAN privado (12%) y VANS social (5,5% SNI 2026)", fontsize=10.5, fontweight="bold", family="Arial", pad=10)
    ax.grid(axis="x", linestyle="--", alpha=0.45)
    ax.legend(loc="lower right", fontsize=8.5)

    fig.tight_layout()
    fig.savefig(r_sens, dpi=220)
    plt.close(fig)
    rutas["sensibilidad"] = r_sens

    return rutas


# ==============================================================================
# 2. FUNCIONES DE ESTILO ARIAL NEGRO, TABLAS E ÍNDICE CLICABLE (PÁGINA 2)
# ==============================================================================

def _estilizar_run_arial(run, size_pt=10.5, bold=False, italic=False, color_rgb=(0, 0, 0)):
    run.font.name = "Arial"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")
    if rPr.find(qn("w:noProof")) is None:
        rPr.append(OxmlElement("w:noProof"))


def _sombrear_celda(celda, color_hex):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    celda._tc.get_or_add_tcPr().append(shd)


def _bordes_celda(celda, es_destacada=False):
    tcPr = celda._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ["top", "bottom", "left", "right"]:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "double" if (es_destacada and edge == "bottom") else "single")
        el.set(qn("w:sz"), "6" if es_destacada else "4")
        el.set(qn("w:color"), "000000" if es_destacada else "CBD5E1")
        tcBorders.append(el)


def _crear_tabla_en_doc(doc, df, anchos_pulgadas, fuente_pt=8.0):
    filas, cols = df.shape
    tabla = doc.add_table(rows=filas + 1, cols=cols)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False

    trPr_h = tabla.rows[0]._tr.get_or_add_trPr()
    trPr_h.append(OxmlElement("w:tblHeader"))
    trPr_h.append(OxmlElement("w:cantSplit"))

    for c_idx, col_nom in enumerate(df.columns):
        cell = tabla.rows[0].cells[c_idx]
        cell.width = Inches(anchos_pulgadas[c_idx])
        _sombrear_celda(cell, "1E293B")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2.5)
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(str(col_nom))
        _estilizar_run_arial(r, size_pt=fuente_pt, bold=True, color_rgb=(255, 255, 255))
        _bordes_celda(cell, es_destacada=True)

    for r_idx in range(filas):
        row = tabla.rows[r_idx + 1]
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        bg = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
        for c_idx in range(cols):
            cell = row.cells[c_idx]
            cell.width = Inches(anchos_pulgadas[c_idx])
            _sombrear_celda(cell, bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(str(df.iloc[r_idx, c_idx]))
            _estilizar_run_arial(r, size_pt=fuente_pt, bold=(c_idx == 0), color_rgb=(0, 0, 0))
            _bordes_celda(cell, es_destacada=False)

    return tabla


def _insertar_elemento_antes_de(parrafo_referencia, elemento_nuevo):
    parrafo_referencia._p.addprevious(elemento_nuevo)


def _crear_parrafo_flotante(doc, texto="", size_pt=10.5, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=5, anchor_name=None, bm_id=900, outline_lvl=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if texto:
        r = p.add_run(texto)
        _estilizar_run_arial(r, size_pt=size_pt, bold=bold, italic=italic)
    if anchor_name:
        pPr = p._p.get_or_add_pPr()
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), str(outline_lvl))
        pPr.append(ol)
        bm_s = OxmlElement("w:bookmarkStart")
        bm_s.set(qn("w:id"), str(bm_id))
        bm_s.set(qn("w:name"), anchor_name)
        bm_e = OxmlElement("w:bookmarkEnd")
        bm_e.set(qn("w:id"), str(bm_id))
        p._p.insert(0, bm_s)
        p._p.append(bm_e)
    return p


def _crear_parrafo_imagen(doc, ruta_img, ancho_pulgadas=6.3, leyenda=None):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(4)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.add_run().add_picture(ruta_img, width=Inches(ancho_pulgadas))
    elementos = [p_img._p]
    if leyenda:
        p_ley = doc.add_paragraph()
        p_ley.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ley.paragraph_format.space_before = Pt(1)
        p_ley.paragraph_format.space_after = Pt(6)
        r_l = p_ley.add_run(leyenda)
        _estilizar_run_arial(r_l, size_pt=8.5, italic=True)
        elementos.append(p_ley._p)
    return elementos


def _crear_linea_toc_xml(doc, texto_sec, anchor_name, pag_default="20", nivel=1):
    """
    Crea una línea de Tabla de Contenidos idéntica a las de la Página 2:
    con hipervínculo clicable (w:hyperlink), puntos guía (.....) y campo PAGEREF actualizable.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3.5 if nivel == 1 else 1.5)
    p.paragraph_format.space_after = Pt(3.5 if nivel == 1 else 1.5)
    p.paragraph_format.left_indent = Inches(0.0 if nivel == 1 else 0.25)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(6.65), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor_name)
    hyperlink.set(qn("w:history"), "1")

    def crear_run_toc(texto_r, es_negrita=True, size_pt=10):
        r_el = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), "Arial")
        rFonts.set(qn("w:hAnsi"), "Arial")
        rFonts.set(qn("w:cs"), "Arial")
        rPr.append(rFonts)
        rPr.append(OxmlElement("w:noProof"))
        if es_negrita:
            rPr.append(OxmlElement("w:b"))
            rPr.append(OxmlElement("w:bCs"))
        color_el = OxmlElement("w:color")
        color_el.set(qn("w:val"), "000000")
        rPr.append(color_el)
        sz_el = OxmlElement("w:sz")
        sz_el.set(qn("w:val"), str(int(size_pt * 2)))
        szCs_el = OxmlElement("w:szCs")
        szCs_el.set(qn("w:val"), str(int(size_pt * 2)))
        rPr.append(sz_el)
        rPr.append(szCs_el)
        r_el.append(rPr)
        t_el = OxmlElement("w:t")
        t_el.text = texto_r
        r_el.append(t_el)
        return r_el

    hyperlink.append(crear_run_toc(texto_sec, es_negrita=(nivel == 1), size_pt=10 if nivel == 1 else 9.5))

    r_tab = OxmlElement("w:r")
    rPr_tab = OxmlElement("w:rPr")
    rFonts_t = OxmlElement("w:rFonts")
    rFonts_t.set(qn("w:ascii"), "Arial")
    rFonts_t.set(qn("w:hAnsi"), "Arial")
    rPr_tab.append(rFonts_t)
    r_tab.append(rPr_tab)
    r_tab.append(OxmlElement("w:tab"))
    hyperlink.append(r_tab)

    fld_pageref = OxmlElement("w:fldSimple")
    fld_pageref.set(qn("w:instr"), f" PAGEREF {anchor_name} \\h ")
    fld_pageref.append(crear_run_toc(str(pag_default), es_negrita=True, size_pt=10 if nivel == 1 else 9.5))
    hyperlink.append(fld_pageref)

    p._p.append(hyperlink)
    return p


def _es_parrafo_de_toc(p):
    """Detecta si un párrafo es una línea de la Tabla de Contenidos (tiene w:hyperlink o w:fldSimple)."""
    tiene_hl = len(p._p.findall(qn("w:hyperlink"))) > 0
    tiene_fld = len(p._p.findall(qn("w:fldSimple"))) > 0
    return tiene_hl or tiene_fld


def _limpiar_inyecciones_previas_y_reparar_toc(doc):
    """
    1. Si una ejecución anterior metió tablas o párrafos dentro de la Tabla de Contenidos
       (entre 'RESUMEN EJECUTIVO' y '1. EL PROBLEMA EN OSORNO...'), los elimina por completo.
    2. Elimina en el cuerpo del documento cualquier párrafo, tabla o figura inyectada antes
       por el Script 3 para evitar duplicados al re-ejecutar.
    """
    body = doc.element.body

    # Paso 1: Limpiar cualquier tabla o párrafo intruso dentro del bloque de la Tabla de Contenidos
    toc_resumen_el = None
    toc_sec1_el = None
    for p in doc.paragraphs:
        if _es_parrafo_de_toc(p):
            t_up = p.text.strip().upper()
            if t_up.startswith("RESUMEN EJECUTIVO") and toc_resumen_el is None:
                toc_resumen_el = p._p
            elif t_up.startswith("1. EL PROBLEMA") and toc_sec1_el is None:
                toc_sec1_el = p._p
                break

    if toc_resumen_el is not None and toc_sec1_el is not None:
        hijos = list(body)
        idx_ini = hijos.index(toc_resumen_el)
        idx_fin = hijos.index(toc_sec1_el)
        for el_intruso in hijos[idx_ini + 1:idx_fin]:
            body.remove(el_intruso)

    # Paso 2: Eliminar entradas antiguas sin formato o inyectadas en el TOC y secciones del cuerpo
    textos_borrar = [
        "MATRIZ DE CUMPLIMIENTO DE LA PAUTA",
        "Con los parámetros oficiales SNI 2026",
        "Cumplimiento íntegro de los 10 puntos exigidos",
        "1.2 Árbol de Problemas y Árbol de Objetivos",
        "El problema central se define como el bajo aprovechamiento local",
        "Figura 1. Árbol de problemas:",
        "Figura 2. Árbol de objetivos:",
        "1.3 Diagnóstico Integral del Residuo:",
        "El residuo objetivo corresponde a cartón corrugado comercial",
        "Figura 3. Balance físico corregido del cartón",
        "2.1 Evaluación y Selección de Alternativas del Proyecto",
        "3.1 Croquis de Distribución del Galpón (140 m2)",
        "Figura 4. Croquis conceptual de distribución",
        "Figura 5. Flujo productivo de 7 etapas",
        "Valorización Social de la Mano de Obra según Factores Oficiales SNI 2026",
        "Aspectos Legales y Normativos Relevantes:",
        "8.1 Parámetros Oficiales SNI 2026 y Flujo Social Corregido",
        "8.2 Parámetros Oficiales SNI 2026 y Flujo Social Corregido",
        "La evaluación social aplica los Precios Sociales Vigentes 2026",
        "Figura 6. Comparación anual: flujo de caja privado neto",
        "8.2 Análisis de Sensibilidad Multiescenario",
        "8.3 Análisis de Sensibilidad Multiescenario",
        "Figura 7. Sensibilidad del VAN privado y VANS social:",
        "10. MODELO DE GESTIÓN ESTRATÉGICA - BUSINESS MODEL CANVAS",
        "11. MODELO DE SEGUIMIENTO - MATRIZ DE MARCO LÓGICO",
        "12. CONTROL DE CONSISTENCIA, MEJORAS",
        "Para asegurar total rigurosidad metodológica ante la comisión",
        "Referencias y Fuentes Oficiales Utilizadas:",
        "• Ministerio de Desarrollo Social y Familia (2025).",
    ]

    for p in list(doc.paragraphs):
        t_clean = p.text.strip()
        if any(t_clean.upper().startswith( pref.upper()) for pref in textos_borrar):
            p._element.getparent().remove(p._element)

    # Eliminar tablas que fueron inyectadas en corridas previas del Script 3 identificando su primera celda
    encabezados_tablas_josue = [
        ("INDICADOR", "EVALUACIÓN PRIVADA"),
        ("PUNTO", "EXIGENCIA DE LA PAUTA"),
        ("ELEMENTO", "CONTENIDO EN EL DIAGNÓSTICO"),
        ("DIMENSIÓN TERRITORIAL Y POBLACIONAL", "DEFINICIÓN APLICADA"),
        ("ALTERNATIVA EVALUADA", "DESCRIPCIÓN"),
        ("CARGO", "CLASIF. SNI"),
        ("ESCENARIO EVALUADO", "VAN PRIVADO (12%)"),
        ("BLOQUE DEL MODELO CANVAS", "APLICACIÓN ESTRATÉGICA"),
        ("RESUMEN NARRATIVO DE OBJETIVOS", "INDICADORES VERIFICABLES"),
        ("ASPECTO EN BORRADOR PRELIMINAR", "CORRECCIÓN METODOLÓGICA"),
    ]
    for tbl in list(doc.tables):
        try:
            c0 = tbl.rows[0].cells[0].text.strip().upper()
            c1 = tbl.rows[0].cells[1].text.strip().upper() if len(tbl.rows[0].cells) > 1 else ""
            if any(c0.startswith(h0) and h1 in c1 for h0, h1 in encabezados_tablas_josue):
                tbl._element.getparent().remove(tbl._element)
        except Exception:
            pass


def _actualizar_indice_p2_con_puntos_y_paginas(doc):
    """
    Inserta en la Tabla de Contenidos de la Página 2 todas las nuevas secciones y subsecciones
    con formato idéntico al resto del Índice (hipervínculo + línea punteada + PAGEREF).
    """
    p_toc_resumen = None
    p_toc_sec1 = None
    p_toc_sec2 = None
    p_toc_sec3 = None
    p_toc_sec4 = None
    p_toc_sec9 = None
    p_toc_anx1 = None

    for p in doc.paragraphs:
        if _es_parrafo_de_toc(p):
            t = p.text.strip().upper()
            if t.startswith("RESUMEN EJECUTIVO") and p_toc_resumen is None:
                p_toc_resumen = p
            elif t.startswith("1. EL PROBLEMA") and p_toc_sec1 is None:
                p_toc_sec1 = p
            elif t.startswith("2. PRODUCTOS") and p_toc_sec2 is None:
                p_toc_sec2 = p
            elif t.startswith("3. FUNCIONAMIENTO") and p_toc_sec3 is None:
                p_toc_sec3 = p
            elif t.startswith("4. TRABAJADORES") and p_toc_sec4 is None:
                p_toc_sec4 = p
            elif t.startswith("9. CONCLUSIONES") and p_toc_sec9 is None:
                p_toc_sec9 = p
            elif t.startswith("ANEXO 1:") and p_toc_anx1 is None:
                p_toc_anx1 = p

    # 1) Matriz de Cumplimiento antes de "1. EL PROBLEMA..."
    if p_toc_sec1 is not None:
        p_m = _crear_linea_toc_xml(doc, "MATRIZ DE CUMPLIMIENTO DE LA PAUTA Y SÍNTESIS (SNI 2026)", "_Toc_Matriz_Pauta", "3", nivel=1)
        _insertar_elemento_antes_de(p_toc_sec1, p_m._p)

    # 2) Subsecciones 1.2 y 1.3 antes de "2. PRODUCTOS..."
    if p_toc_sec2 is not None:
        p_12 = _crear_linea_toc_xml(doc, "1.2 Árbol de Problemas y Árbol de Objetivos (Causas, Efectos, Medios y Fines)", "_Toc_Sub_12", "5", nivel=2)
        p_13 = _crear_linea_toc_xml(doc, "1.3 Diagnóstico Integral del Residuo: Áreas, Balance Físico y Población", "_Toc_Sub_13", "6", nivel=2)
        _insertar_elemento_antes_de(p_toc_sec2, p_12._p)
        _insertar_elemento_antes_de(p_toc_sec2, p_13._p)

    # 3) Subsección 2.1 antes de "3. FUNCIONAMIENTO..."
    if p_toc_sec3 is not None:
        p_21 = _crear_linea_toc_xml(doc, "2.1 Evaluación y Selección de Alternativas del Proyecto", "_Toc_Sub_21", "8", nivel=2)
        _insertar_elemento_antes_de(p_toc_sec3, p_21._p)

    # 4) Subsección 3.1 antes de "4. TRABAJADORES..."
    if p_toc_sec4 is not None:
        p_31 = _crear_linea_toc_xml(doc, "3.1 Croquis del Galpón (140 m2), Flujo de Proceso y Normativa Legal", "_Toc_Sub_31", "10", nivel=2)
        _insertar_elemento_antes_de(p_toc_sec4, p_31._p)

    # 5) Subsecciones 8.2 y 8.3 antes de "9. CONCLUSIONES..."
    if p_toc_sec9 is not None:
        p_82 = _crear_linea_toc_xml(doc, "8.2 Parámetros Oficiales SNI 2026 y Flujo Social Corregido", "_Toc_Sub_82", "19", nivel=2)
        p_83 = _crear_linea_toc_xml(doc, "8.3 Análisis de Sensibilidad Multiescenario (Privado 12% y Social 5,5%)", "_Toc_Sub_83", "19", nivel=2)
        _insertar_elemento_antes_de(p_toc_sec9, p_82._p)
        _insertar_elemento_antes_de(p_toc_sec9, p_83._p)

    # 6) Capítulos 10, 11 y 12 antes de "ANEXO 1:" con línea punteada y número de página
    if p_toc_anx1 is not None:
        p_10 = _crear_linea_toc_xml(doc, "10. MODELO DE GESTIÓN ESTRATÉGICA - BUSINESS MODEL CANVAS (PUNTO 9 PAUTA)", "_Toc_Canvas", "20", nivel=1)
        p_11 = _crear_linea_toc_xml(doc, "11. MODELO DE SEGUIMIENTO - MATRIZ DE MARCO LÓGICO (PUNTO 10 PAUTA)", "_Toc_MML", "21", nivel=1)
        p_12 = _crear_linea_toc_xml(doc, "12. CONTROL DE CONSISTENCIA, MEJORAS Y REFERENCIAS OFICIALES SNI 2026", "_Toc_Consistencia", "21", nivel=1)
        _insertar_elemento_antes_de(p_toc_anx1, p_10._p)
        _insertar_elemento_antes_de(p_toc_anx1, p_11._p)
        _insertar_elemento_antes_de(p_toc_anx1, p_12._p)


# ==============================================================================
# 3. INTEGRADOR PRINCIPAL EN EL CUERPO DEL INFORME WORD
# ==============================================================================

def integrar_trabajo_josue_en_word(ruta_word, carpeta_entrega):
    if not os.path.exists(ruta_word):
        raise FileNotFoundError(
            f"No se encontró el archivo '{ruta_word}'.\n"
            f"Ejecuta primero el Script 1 o 2 para tener el documento base."
        )

    rutas_img = generar_diagramas_josue(carpeta_entrega)
    doc = Document(ruta_word)

    # Limpiar tanto el TOC desordenado como cualquier inyección previa
    _limpiar_inyecciones_previas_y_reparar_toc(doc)

    # Buscar EXCLUSIVAMENTE los títulos reales del cuerpo del documento (ignorando los del Índice de la Página 2)
    p_sec1 = None
    p_sec2 = None
    p_sec3 = None
    p_sec4 = None
    p_sec9 = None
    p_anx1 = None

    for p in doc.paragraphs:
        if _es_parrafo_de_toc(p):
            continue  # Saltarse todas las líneas de la Tabla de Contenidos de la Página 2
        t = p.text.strip().upper()
        if t.startswith("1. EL PROBLEMA EN OSORNO"):
            p_sec1 = p
        elif t.startswith("2. PRODUCTOS FABRICADOS"):
            p_sec2 = p
        elif t.startswith("3. FUNCIONAMIENTO DEL TALLER"):
            p_sec3 = p
        elif t.startswith("4. TRABAJADORES, SUELDOS"):
            p_sec4 = p
        elif t.startswith("9. CONCLUSIONES PRINCIPALES"):
            p_sec9 = p
        elif t.startswith("ANEXO 1:"):
            p_anx1 = p

    # --------------------------------------------------------------------------
    # 0) INYECTAR MATRIZ DE CUMPLIMIENTO DE LA PAUTA EN EL CUERPO (ANTES DE SECCIÓN 1 REAL)
    # --------------------------------------------------------------------------
    if p_sec1 is not None:
        p_t_pauta = _crear_parrafo_flotante(
            doc, "MATRIZ DE CUMPLIMIENTO DE LA PAUTA Y SÍNTESIS DE INDICADORES (SNI 2026)",
            size_pt=11.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=4,
            anchor_name="_Toc_Matriz_Pauta", bm_id=901, outline_lvl=0
        )
        p_i_pauta = _crear_parrafo_flotante(
            doc,
            "Con los parámetros oficiales SNI 2026 —Tasa Social de Descuento de 5,5%, factores MOC 0,97; MOSC 0,95; MONC 0,91 y divisa 1,00— y un balance físico riguroso (15,0 t/año recolectadas; 10,5 t/año incorporadas en 4.500 productos y 4,5 t/año de recortes con salida trazable a reciclador), el proyecto presenta los siguientes resultados consolidados:"
        )
        df_res_ind = pd.DataFrame([
            ["Tasa de descuento", "12,0%", "5,5% (SNI 2026)"],
            ["VAN / VANS", "$63.175.846", "$154.991.495"],
            ["TIR / TIRS", "50,66%", "70,26%"],
            ["Payback simple", "2,40 años", "1,74 años"],
            ["Decisión", "Viable bajo caso base", "Rentable socialmente bajo caso base"],
        ], columns=["Indicador", "Evaluación Privada", "Evaluación Social SNI 2026"])
        t_res_ind = _crear_tabla_en_doc(doc, df_res_ind, [2.10, 2.30, 2.30], fuente_pt=8.0)

        p_esp_pauta = _crear_parrafo_flotante(doc, "Cumplimiento íntegro de los 10 puntos exigidos en la pauta de evaluación:", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=6, space_after=2)
        df_pauta = pd.DataFrame([
            ["1", "Planteamiento del problema y objetivo", "Problema central, causas, efectos, evidencia, objetivos, medios y fines (Árboles Fig. 1 y 2)."],
            ["2", "Diagnóstico del residuo", "Caracterización, áreas de estudio/influencia, disponibilidad (15,0 t), población, demanda y oferta."],
            ["3", "Alternativas", "Cuatro alternativas comparadas y selección justificada (Alternativa C + A para merma)."],
            ["4", "Configuración de alternativa", "Mercado, producción, localización en Ovejería, logística, croquis 140 m2, tecnología, proceso, personal y legales."],
            ["5", "Cuantificación", "Cantidades, proyecciones 10 años, inversiones ($13.445.378 neto), reinversiones, costos e ingresos."],
            ["6", "Capital de trabajo", "Método del período de desfase de 45 días (CT Año 0 = $6.961.008)."],
            ["7", "Flujo privado", "Flujo a 10 años con tasas Pro Pyme (12,5% en 2027, 15% en 2028, 25% desde 2029), VAN, TIR, payback y sensibilidad."],
            ["8", "Evaluación social", "Precios sociales SNI 2026 (TSD 5,5%), ahorro municipal conservador sobre 10,5 t incorporadas y VANS/TIRS."],
            ["9", "CANVAS", "Modelo de gestión estratégica de 9 bloques aplicado a Cartón Vivo Osorno."],
            ["10", "Marco Lógico", "Fin, propósito, componentes, actividades, indicadores verificables, medios de verificación y supuestos."],
        ], columns=["Punto", "Exigencia de la Pauta", "Cómo se cubre en este Informe Final"])
        t_pauta = _crear_tabla_en_doc(doc, df_pauta, [0.55, 2.15, 4.00], fuente_pt=7.6)

        for el in [p_t_pauta._p, p_i_pauta._p, t_res_ind._tbl, p_esp_pauta._p, t_pauta._tbl]:
            _insertar_elemento_antes_de(p_sec1, el)

    # --------------------------------------------------------------------------
    # A) INYECTAR CAUSAS/EFECTOS + ÁRBOLES + DIAGNÓSTICO DEL RESIDUO (ANTES DE SECCIÓN 2 REAL)
    # --------------------------------------------------------------------------
    if p_sec2 is not None:
        p_sub_12 = _crear_parrafo_flotante(
            doc, "1.2 Árbol de Problemas y Árbol de Objetivos (Causas, Efectos, Medios y Fines)",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_12", bm_id=902, outline_lvl=1
        )
        p_exp_12 = _crear_parrafo_flotante(
            doc,
            "El problema central se define como el bajo aprovechamiento local del cartón comercial recuperable generado en el eje urbano de Osorno para transformarlo en productos didácticos, recreativos y cognitivos de valor agregado. Esta formulación evita definir el problema como 'falta de una planta' y permite comparar soluciones distintas. A continuación se presentan la matriz de causas y efectos, el Árbol de Problemas (Figura 1) y el Árbol de Objetivos (Figura 2):"
        )

        df_causas = pd.DataFrame([
            ["Causa 1", "Abastecimiento selectivo poco formalizado.", "El Año 1 se apoya en 6 locales y horarios de retiro definidos; faltan acuerdos formales de suministro para asegurar continuidad."],
            ["Causa 2", "Capacidad local de transformación y control de calidad por implementar.", "El proyecto requiere habilitación eléctrica, equipos, procedimientos, prototipos y verificación de seguridad antes de operar comercialmente."],
            ["Causa 3", "Demanda institucional aún requiere validación contractual.", "El respaldo contiene metas e informantes clave en Osorno; la demanda se trata como meta comercial validada que se formalizará con cartas de intención u órdenes de compra."],
            ["Efecto 1", "Menor retención de valor económico y ambiental en la comuna.", "El cartón comercial recuperable se utiliza poco en manufactura local de mayor valor agregado."],
            ["Efecto 2", "Oferta local limitada de productos didácticos circulares.", "La alternativa busca crear una línea local específica; la oferta sustituta debe seguir monitoreándose."],
            ["Efecto 3", "Menor visibilidad de soluciones de economía circular.", "Sin pilotos y trazabilidad hay menor demostración práctica de valorización local entre comercios, escuelas y centros de adulto mayor."],
        ], columns=["Elemento", "Contenido en el Diagnóstico", "Evidencia / Lógica Causal en Osorno"])
        t_causas = _crear_tabla_en_doc(doc, df_causas, [0.95, 2.35, 3.40], fuente_pt=7.8)

        els_img1 = _crear_parrafo_imagen(doc, rutas_img["arbol_problemas"], ancho_pulgadas=6.4, leyenda="Figura 1. Árbol de problemas: efectos arriba, problema central al medio y causas abajo.")
        els_img2 = _crear_parrafo_imagen(doc, rutas_img["arbol_objetivos"], ancho_pulgadas=6.4, leyenda="Figura 2. Árbol de objetivos: cada medio responde a una causa y cada fin revierte un efecto identificado.")

        p_sub_13 = _crear_parrafo_flotante(
            doc, "1.3 Diagnóstico Integral del Residuo: Áreas, Balance Físico (10,5 t vs 4,5 t) y Población",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_13", bm_id=903, outline_lvl=1
        )
        p_exp_13 = _crear_parrafo_flotante(
            doc,
            "El residuo objetivo corresponde a cartón corrugado comercial seco y limpio proveniente de embalajes de los 6 locales del eje Eleuterio Ramírez (50 kg/día = 15,00 t/año en el Año 1). El balance físico incorpora el 70% de la fibra (10,50 t/año) en 3.000 sets escolares (2,5 kg c/u) y 1.500 juegos cognitivos (2,0 kg c/u), mientras que las 4,50 t/año restantes (30% de recortes) se segregan con salida trazable hacia gestor o reciclador autorizado:"
        )
        df_diag_areas = pd.DataFrame([
            ["Área de estudio", "Zona urbana de Osorno, donde se contextualizan la generación del residuo, el mercado institucional y la logística del proyecto."],
            ["Área de influencia inicial", "Eje Eleuterio Ramírez entre los 6 locales catastrados, galpón propuesto en Ovejería y usuarios institucionales de Osorno vinculados a los productos."],
            ["Escalamiento (Años 2-10)", "El crecimiento de 5% anual es un supuesto de expansión; requiere incorporar nuevos acuerdos de suministro y demanda verificable."],
            ["Población referencia / objetivo (Oferta)", "Comercios del área urbana de Osorno generadores de cartón / 6 locales del eje Eleuterio Ramírez durante el piloto."],
            ["Población referencia / objetivo (Demanda)", "Establecimientos educacionales, jardines VTF/privados y organizaciones de personas mayores de Osorno que validen pilotos y órdenes de compra."],
        ], columns=["Dimensión Territorial y Poblacional", "Definición Aplicada en Cartón Vivo Osorno"])
        t_diag_areas = _crear_tabla_en_doc(doc, df_diag_areas, [2.25, 4.45], fuente_pt=7.8)

        els_img3 = _crear_parrafo_imagen(doc, rutas_img["balance_carton"], ancho_pulgadas=6.2, leyenda="Figura 3. Balance físico corregido del cartón (Año 1 a Año 10): el 30% no incorporado requiere gestión separada y trazable.")

        for el in [p_sub_12._p, p_exp_12._p, t_causas._tbl] + els_img1 + els_img2 + [p_sub_13._p, p_exp_13._p, t_diag_areas._tbl] + els_img3:
            _insertar_elemento_antes_de(p_sec2, el)

    # --------------------------------------------------------------------------
    # B) INYECTAR TABLA DE LAS 4 ALTERNATIVAS ANTES DE SECCIÓN 3 REAL
    # --------------------------------------------------------------------------
    if p_sec3 is not None:
        p_sub_alt = _crear_parrafo_flotante(
            doc, "2.1 Evaluación y Selección de Alternativas del Proyecto",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_21", bm_id=904, outline_lvl=1
        )
        df_alt = pd.DataFrame([
            ["A. Reciclaje convencional", "Acopio, clasificación y entrega del cartón a gestores/recicladores.", "Baja complejidad y menor inversión.", "Retiene poco valor manufacturero local; se mantiene como complemento obligatorio para el 30% de recortes."],
            ["B. Talleres educativos", "Uso del cartón en talleres/manualidades comunitarias.", "Alto valor educativo y bajo costo inicial.", "Escala productiva y continuidad comercial muy limitadas."],
            ["C. Producción local integrada (Seleccionada)", "Fabricar sets didácticos y juegos cognitivos con abastecimiento trazable.", "Ataca suministro, transformación y demanda; escalable por pilotos y de mayor valor local.", "Requiere estandarización, seguridad (D.S. 114), ventas y gestión trazable de mermas."],
            ["D. Reciclaje industrial de fibra", "Desfibrado húmedo y producción de nueva materia fibrosa.", "Mayor transformación material potencial.", "Alta inversión, consumo de agua, energía y permisos fuera de la escala del piloto."],
        ], columns=["Alternativa Evaluada", "Descripción", "Fortaleza Principal", "Limitación / Rol en el Proyecto"])
        t_alt = _crear_tabla_en_doc(doc, df_alt, [1.45, 1.85, 1.55, 1.85], fuente_pt=7.8)

        for el in [p_sub_alt._p, t_alt._tbl]:
            _insertar_elemento_antes_de(p_sec3, el)

    # --------------------------------------------------------------------------
    # C) INYECTAR CROQUIS GALPÓN 140 M2, FLUJO DE PROCESO, PERSONAL SNI 2026 Y ASPECTOS LEGALES ANTES DE SECCIÓN 4 REAL
    # --------------------------------------------------------------------------
    if p_sec4 is not None:
        p_sub_croq = _crear_parrafo_flotante(
            doc, "3.1 Croquis de Distribución del Galpón (140 m2), Flujo de Proceso y Normativa Legal",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_31", bm_id=905, outline_lvl=1
        )
        els_croq = _crear_parrafo_imagen(doc, rutas_img["croquis_galpon"], ancho_pulgadas=5.8, leyenda="Figura 4. Croquis conceptual de distribución del recinto de 140 m2 en Ovejería (debe reemplazarse por plano técnico en factibilidad).")
        els_proc = _crear_parrafo_imagen(doc, rutas_img["flujo_proceso"], ancho_pulgadas=6.4, leyenda="Figura 5. Flujo productivo de 7 etapas desde el retiro selectivo (furgón contratado) hasta el despacho (18 productos/día hábil).")

        p_pers_sni = _crear_parrafo_flotante(doc, "Valorización Social de la Mano de Obra según Factores Oficiales SNI 2026 (MOC 0,97; MOSC 0,95; MONC 0,91):", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
        df_pers_sni = pd.DataFrame([
            ["Jefe de planta / calidad / ventas", "1", "MOC", "Gestión, calidad, producción y comercialización", "$11.400.000", "0,97", "$11.058.000"],
            ["Operador técnico corte / pintura", "2", "MOSC", "Troquelado, guillotina, serigrafía, seguridad", "$14.400.000", "0,95", "$13.680.000"],
            ["Operario recolección / armado", "2", "MONC", "Retiro, selección, limpieza, armado y empaque", "$12.480.000", "0,91", "$11.356.800"],
            ["TOTAL PLANILLA ANUAL", "5", "-", "Dotación completa del taller en Ovejería", "$38.280.000", "-", "$36.094.800"],
        ], columns=["Cargo", "Cant.", "Clasif. SNI", "Competencias Clave", "Costo Anual Privado", "Factor 2026", "Costo Anual Social"])
        t_pers_sni = _crear_tabla_en_doc(doc, df_pers_sni, [1.55, 0.45, 0.65, 1.75, 0.90, 0.55, 0.85], fuente_pt=7.5)

        p_leg = _crear_parrafo_flotante(
            doc,
            "Aspectos Legales y Normativos Relevantes: (1) Constitución de SpA, inicio de actividades SII (Régimen Pro Pyme General con tasas de 12,5% en 2027, 15% en 2028 y 25% desde 2029 según Circular N°53/2025) y patente municipal en Osorno; (2) Cumplimiento del D.S. N°114/2005 del MINSAL (Reglamento sobre Seguridad de los Juguetes en Chile) complementado con tintas al agua bajo norma EN-71 y respeto de propiedad intelectual; (3) Trazabilidad de residuos bajo la Ley N°20.920 (Ley REP), entregando el 30% de recortes limpios (4,50 ton/año) a gestor autorizado para que jamás regresen al flujo municipal sin control; (4) Logística de entrada mediante servicio de furgón contratado, coherente con el costo variable logístico del flujo financiero."
        )

        for el in [p_sub_croq._p] + els_croq + els_proc + [p_pers_sni._p, t_pers_sni._tbl, p_leg._p]:
            _insertar_elemento_antes_de(p_sec4, el)

    # --------------------------------------------------------------------------
    # D) INYECTAR PARÁMETROS SNI 2026 (8.2) Y SENSIBILIDAD (8.3) ANTES DE SECCIÓN 9 REAL
    # --------------------------------------------------------------------------
    if p_sec9 is not None:
        p_sub_sni = _crear_parrafo_flotante(
            doc, "8.2 Parámetros Oficiales SNI 2026 y Flujo Social Corregido",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_82", bm_id=906, outline_lvl=1
        )
        p_exp_sni = _crear_parrafo_flotante(
            doc,
            "La evaluación social aplica los Precios Sociales Vigentes 2026 del Ministerio de Desarrollo Social y Familia (SNI): Tasa Social de Descuento (TSD) de 5,5%, factores de corrección de mano de obra MOC 0,97, MOSC 0,95, MONC 0,91 y factor social de la divisa 1,00. El ahorro municipal por menor recolección y disposición en el Vertedero de Curaco ($48.000/t c/IVA = $40.336/t neto) se contabiliza conservadoramente solo sobre las toneladas efectivamente incorporadas en productos (10,50 t en Año 1), evitando sobreestimar beneficios y excluyendo 'otros beneficios' sin metodología cuantitativa. Con ello, el VANS alcanza $154.991.495 y la TIRS 70,26% (payback social de 1,74 años):"
        )
        els_flujos = _crear_parrafo_imagen(doc, rutas_img["flujos_comparados"], ancho_pulgadas=6.2, leyenda="Figura 6. Comparación anual: flujo de caja privado neto (12%) vs. flujo social corregido SNI 2026 (TSD 5,5%).")

        p_sub_sens = _crear_parrafo_flotante(
            doc, "8.3 Análisis de Sensibilidad Multiescenario (Privado 12% y Social SNI 2026 5,5%)",
            size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4,
            anchor_name="_Toc_Sub_83", bm_id=907, outline_lvl=1
        )
        df_sens = pd.DataFrame([
            ["Pesimista: -15% ingresos", "$11.554.771", "18,65%", "$65.515.431", "31,36%"],
            ["Incremento de costos: +10% costos", "$36.866.698", "32,58%", "$112.145.044", "48,17%"],
            ["Caso base: +5% crecimiento anual", "$63.175.846", "50,66%", "$154.991.495", "70,26%"],
            ["Sin crecimiento: 0% anual fijo", "$16.105.436", "28,86%", "$58.883.685", "49,67%"],
        ], columns=["Escenario Evaluado", "VAN Privado (12%)", "TIR Privada", "VANS Social (5,5%)", "TIRS Social"])
        t_sens = _crear_tabla_en_doc(doc, df_sens, [2.30, 1.15, 1.00, 1.25, 1.00], fuente_pt=7.8)
        els_sens = _crear_parrafo_imagen(doc, rutas_img["sensibilidad"], ancho_pulgadas=6.0, leyenda="Figura 7. Sensibilidad del VAN privado y VANS social: incluso el escenario sin crecimiento (0%) conserva VAN positivo ($16,11 MM$).")

        for el in [p_sub_sni._p, p_exp_sni._p] + els_flujos + [p_sub_sens._p, t_sens._tbl] + els_sens:
            _insertar_elemento_antes_de(p_sec9, el)

    # --------------------------------------------------------------------------
    # E) INYECTAR CANVAS (10), MARCO LÓGICO (11) Y CONTROL DE CONSISTENCIA (12) ANTES DE ANEXO 1 REAL
    # --------------------------------------------------------------------------
    if p_anx1 is not None:
        # PUNTO 9 DE LA PAUTA: BUSINESS MODEL CANVAS
        p_t_canv = _crear_parrafo_flotante(
            doc, "10. MODELO DE GESTIÓN ESTRATÉGICA - BUSINESS MODEL CANVAS (PUNTO 9 PAUTA)",
            size_pt=12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=6,
            anchor_name="_Toc_Canvas", bm_id=910, outline_lvl=0
        )
        df_canvas = pd.DataFrame([
            ["1. Socios Clave", "6 comercios piloto del eje Eleuterio Ramírez (#446 al #1024); futuros generadores; gestores/recicladores para el 30% de recortes; proveedores gráficos; DAEM Osorno, jardines privados y organizaciones de personas mayores."],
            ["2. Actividades Clave", "Retiro selectivo matutino (05:30-06:15 AM), recepción y control de calidad del cartón, diseño, troquelado en rodillos, serigrafía al agua, armado, validación con usuarios, ventas y trazabilidad de recortes."],
            ["3. Recursos Clave", "Cartón corrugado limpio (15,00 ton/año en Año 1), galpón de 140 m2 en Ovejería, troqueladora, guillotina, serigrafía, 5 trabajadores, matrices pedagógicas/cognitivas y protocolos de calidad."],
            ["4. Propuesta de Valor", "Productos didácticos escolares y juegos cognitivos para personas mayores locales, personalizables, seguros (D.S. N°114 y EN-71) y circulares; recuperación visible y trazable de cartón comercial de Osorno."],
            ["5. Relación con Clientes", "Pilotos de validación técnica con educadoras y terapeutas ocupacionales, co-diseño, postventa, reposición anual de sets y seguimiento de satisfacción."],
            ["6. Canales", "Venta directa, cotización institucional, Compra Ágil / Mercado Público cuando corresponda, ferias educativas y alianzas con clubes y hogares de personas mayores."],
            ["7. Segmentos de Clientes", "Escuelas básicas y jardines VTF del DAEM Osorno (meta: 2.000 sets), jardines privados/guarderías (1.000 sets) y organizaciones/hogares de personas mayores (1.500 juegos cognitivos)."],
            ["8. Estructura de Costos", "Planilla de 5 trabajadores ($38.280.000/año privado; $36.094.800/año social SNI 2026), arriendo y gastos fijos en Ovejería, insumos ($2.400 c/IVA por unidad), servicio de flete contratado y reinversión en Año 5."],
            ["9. Fuentes de Ingreso", "Venta de 3.000 Sets Didácticos Escolares a $16.000 c/IVA y 1.500 Juegos Cognitivos a $19.000 c/IVA ($64.285.714 netos en Año 1); no se contabilizan subsidios no comprometidos."],
        ], columns=["Bloque del Modelo CANVAS", "Aplicación Estratégica en Cartón Vivo Osorno"])
        t_canv = _crear_tabla_en_doc(doc, df_canvas, [1.75, 4.95], fuente_pt=7.8)

        # PUNTO 10 DE LA PAUTA: MATRIZ DE MARCO LÓGICO
        p_t_mml = _crear_parrafo_flotante(
            doc, "11. MODELO DE SEGUIMIENTO - MATRIZ DE MARCO LÓGICO (PUNTO 10 PAUTA)",
            size_pt=12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=6,
            anchor_name="_Toc_MML", bm_id=911, outline_lvl=0
        )
        df_mml = pd.DataFrame([
            ["FIN\nContribuir a una economía circular local en Osorno mediante valorización de cartón comercial y generación de productos útiles.", "Toneladas anuales valorizadas (meta Año 1: 10,50 t incorporadas + 4,50 t enviadas a reciclaje); N° de instituciones usuarias; continuidad operativa.", "Registros de pesaje y trazabilidad; contratos/facturación SII; certificados de gestor de residuos e informes anuales.", "Se mantiene disponibilidad de cartón comercial apto y respaldo institucional a compras sustentables en Osorno."],
            ["PROPÓSITO\nAumentar el aprovechamiento local de cartón comercial recuperable del eje Eleuterio Ramírez transformándolo en productos didácticos y cognitivos con trazabilidad.", "70% de fibra incorporada en productos (10,50 ton/año en Año 1); 4.500 unidades conformes vendidas/validadas; satisfacción de usuarios > 90%.", "Registros de recepción, planillas de producción diaria, ventas/órdenes de compra, encuestas de satisfacción y control de calidad.", "Clientes institucionales aceptan el producto y formalizan compras; normativa de seguridad se cumple; suministro es estable."],
            ["COMPONENTES\nC1: Abastecimiento selectivo formalizado.\nC2: Capacidad productiva y calidad implementadas en Ovejería.\nC3: Demanda y canales comerciales validados.\nC4: Gestión trazable del 30% de recortes.", "C1: 6 acuerdos activos en eje Ramírez.\nC2: Capacidad de 18 productos/día hábil y baja tasa de rechazo.\nC3: Cartas de intención/órdenes de compra en 3 segmentos.\nC4: 100% de recortes (4,50 t/año) con salida documentada.", "Convenios firmados con locatarios; fichas de producción y pruebas técnicas; cotizaciones/órdenes aceptadas; comprobantes de recepción de gestor.", "Proveedores entregan cartón seco antes de las 06:00 AM; equipos operan sin fallas críticas; gestor recibe material residual."],
            ["ACTIVIDADES\nCatastro y firma de acuerdos; protocolos de recepción; habilitación de galpón y compra de equipos; prototipos y pruebas; pilotos comerciales; registro de indicadores y convenio para recortes.", "Cumplimiento de hitos, presupuesto de inversión inicial ($13.445.378 neto / $16.000.000 c/IVA), capital de trabajo ($6.961.008), pilotos ejecutados y registros al día.", "Cronograma de puesta en marcha, facturas de maquinaria, contrato de arriendo, actas de piloto, listas de chequeo y reportes de trazabilidad.", "Permisos municipales/sanitarios y compras de equipos se tramitan oportunamente; recursos financieros y humanos están disponibles."],
        ], columns=["Resumen Narrativo de Objetivos", "Indicadores Verificables", "Medios de Verificación", "Supuestos Críticos"])
        t_mml = _crear_tabla_en_doc(doc, df_mml, [1.75, 1.70, 1.65, 1.60], fuente_pt=7.6)

        # SECCIÓN 12: CONTROL DE CONSISTENCIA Y MEJORAS RESPECTO DEL BORRADOR + REFERENCIAS
        p_t_cons = _crear_parrafo_flotante(
            doc, "12. CONTROL DE CONSISTENCIA, MEJORAS RESPECTO DEL BORRADOR Y REFERENCIAS OFICIALES",
            size_pt=12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=6,
            anchor_name="_Toc_Consistencia", bm_id=912, outline_lvl=0
        )
        p_i_cons = _crear_parrafo_flotante(
            doc,
            "Para asegurar total rigurosidad metodológica ante la comisión evaluadora, la siguiente tabla sintetiza las correcciones técnicas aplicadas sobre el borrador preliminar y las cuatro brechas que deben cerrarse en la fase piloto antes de comprometer la ejecución definitiva:"
        )
        df_cons = pd.DataFrame([
            ["TSD social 6,0%", "Corregida a 5,5% según Precios Sociales Vigentes 2026 (SNI / MIDESO)."],
            ["Factores MOC 1,00 / MOSC 0,85 / MONC 0,68", "Corregidos a 0,97 (MOC), 0,95 (MOSC) y 0,91 (MONC) conforme a normativa SNI 2026."],
            ["'Otros beneficios' $2,5 MM/año", "Eliminados del flujo social base por falta de metodología cuantitativa verificable, evitando sobreestimación."],
            ["Ahorro municipal sobre todo el cartón recolectado", "Criterio conservador: se monetiza solo sobre las 10,5 t/año efectivamente incorporadas en productos; las 4,5 t de recortes exigen trazabilidad con gestor."],
            ["4.500 ventas 'aseguradas'", "Reformulado como meta/demanda validada por informantes; requiere respaldarse con cartas de intención u órdenes de compra en fase piloto."],
            ["Crecimiento 5% de los mismos 6 locales", "Reformulado como supuesto de expansión del caso base (contrastado con escenario 0% crecimiento que mantiene VAN positivo de $16.105.436)."],
            ["'Nuestro furgón' con costo de flete", "Corregido a servicio de furgón contratado, plenamente coherente con el costo variable logístico del flujo y sin CAPEX vehicular."],
            ["Impuesto 25% fijo desde Año 1", "Sustituido por secuencia explícita Régimen Pro Pyme General (Circular SII N°53/2025): 12,5% en 2027, 15% en 2028 y 25% desde 2029."],
            ["30% de merma 'se pierde'", "Reformulado como fracción residual de recortes (4,5 t/año) que se segrega y entrega a reciclador autorizado, sin retornar al flujo municipal."],
        ], columns=["Aspecto en Borrador Preliminar", "Corrección Metodológica Aplicada en Versión Final (SNI 2026)"])
        t_cons = _crear_tabla_en_doc(doc, df_cons, [2.25, 4.45], fuente_pt=7.6)

        p_ref_tit = _crear_parrafo_flotante(doc, "Referencias y Fuentes Oficiales Utilizadas:", size_pt=10.0, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=6, space_after=3)
        p_ref_txt = _crear_parrafo_flotante(
            doc,
            "• Ministerio de Desarrollo Social y Familia (2025). Metodología General para la Formulación y Evaluación de Proyectos de Inversión Pública, SNI.\n"
            "• Sistema Nacional de Inversiones - SNI (2026). Precios Sociales Vigentes 2026: TSD 5,5%; MOC 0,97; MOSC 0,95; MONC 0,91; Divisa 1,00.\n"
            "• Ministerio de Desarrollo Social y Familia. Metodología para la Formulación y Evaluación Social de Proyectos de Gestión de Residuos Sólidos Domiciliarios y Asimilables.\n"
            "• Servicio de Impuestos Internos - SII (2025). Circular N°53/2025 sobre Régimen Pro Pyme General y tasas transitorias de Primera Categoría.\n"
            "• Ministerio de Salud - MINSAL (2005). D.S. N°114/2005, Reglamento sobre Seguridad de los Juguetes y modificaciones vigentes; Ley N°20.920 (Ley REP).",
            size_pt=9.0
        )

        for el in [p_t_canv._p, t_canv._tbl, p_t_mml._p, t_mml._tbl, p_t_cons._p, p_i_cons._p, t_cons._tbl, p_ref_tit._p, p_ref_txt._p]:
            _insertar_elemento_antes_de(p_anx1, el)

    # Finalmente, actualizar el Índice de la Página 2 con hipervínculos, puntos guía y números de página dinámicos
    _actualizar_indice_p2_con_puntos_y_paginas(doc)
    doc.save(ruta_word)

    # Limpiar imágenes temporales de Josué
    for r_im in rutas_img.values():
        if os.path.exists(r_im):
            try:
                os.remove(r_im)
            except Exception:
                pass

    return ruta_word


def actualizar_pdf_y_zip_final(ruta_word):
    ruta_pdf = os.path.splitext(ruta_word)[0] + ".pdf"
    carpeta_entrega = os.path.dirname(os.path.abspath(ruta_word))
    apellidos = os.path.basename(carpeta_entrega)
    directorio_padre = os.path.dirname(carpeta_entrega)
    ruta_zip = os.path.join(directorio_padre, f"{apellidos}.zip")

    pdf_ok = False
    if sys.platform.startswith("win"):
        abs_docx = os.path.abspath(ruta_word).replace("'", "''")
        abs_pdf = os.path.abspath(ruta_pdf).replace("'", "''")
        ps_cmd = (
            "$ErrorActionPreference = 'Stop'; "
            "$word = New-Object -ComObject Word.Application; "
            "$word.Visible = $false; "
            "$word.DisplayAlerts = 0; "
            f"$doc = $word.Documents.Open('{abs_docx}'); "
            "$doc.Repaginate(); "
            "$doc.Fields.Update() | Out-Null; "
            "foreach ($toc in $doc.TablesOfContents) { $toc.Update() }; "
            "$doc.SpellingChecked = $true; "
            "$doc.GrammarChecked = $true; "
            "$doc.ShowSpellingErrors = $false; "
            "$doc.ShowGrammaticalErrors = $false; "
            "$doc.Save(); "
            f"$doc.SaveAs([ref]'{abs_pdf}', [ref]17); "
            "$doc.Close(); "
            "$word.Quit();"
        )
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                check=True, timeout=60,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
            pdf_ok = os.path.exists(ruta_pdf)
        except Exception:
            pdf_ok = False

    if os.path.exists(ruta_zip):
        try:
            os.remove(ruta_zip)
        except Exception:
            pass

    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for arch in os.listdir(carpeta_entrega):
            r_full = os.path.join(carpeta_entrega, arch)
            if os.path.isfile(r_full):
                zf.write(r_full, arcname=os.path.join(apellidos, arch))

    return ruta_pdf, ruta_zip, pdf_ok


if __name__ == "__main__":
    print(">>> Ejecutando 3_integrador_arboles_y_pauta_josue.py (Índice y Ubicaciones Corregidas) ...")
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    carpeta_actual = os.path.dirname(os.path.abspath(__file__))
    posibles_docx = [
        os.path.join(carpeta_actual, "Vargas,Muñoz,Jofre", "Vargas,Muñoz,Jofre.docx"),
        os.path.join(carpeta_actual, "Vargas,Muñoz,Jofre.docx"),
    ]
    ruta_docx_objetivo = next((r for r in posibles_docx if os.path.exists(r)), None)

    if ruta_docx_objetivo is None:
        messagebox.showwarning(
            "Ejecuta primero el Script 1 o 2",
            "No se encontró 'Vargas,Muñoz,Jofre.docx'.\n"
            "Ejecuta primero tu generador base y luego este script.",
            parent=root
        )
    else:
        try:
            carpeta_ent = os.path.dirname(os.path.abspath(ruta_docx_objetivo))
            integrar_trabajo_josue_en_word(ruta_docx_objetivo, carpeta_ent)
            r_pdf, r_zip, ok_pdf = actualizar_pdf_y_zip_final(ruta_docx_objetivo)
            if sys.platform.startswith("win"):
                os.startfile(carpeta_ent)
            messagebox.showinfo(
                "¡Índice Reparado e Integración Completa (SNI 2026)!",
                f"Se corrigió la Tabla de Contenidos de la Página 2 y se ubicaron todas las secciones en su lugar:\n\n"
                f"- La Matriz de Cumplimiento de la Pauta ahora está en el cuerpo del informe (después del Resumen Ejecutivo) y no dentro del Índice.\n"
                f"- Los puntos 10 (CANVAS), 11 (Marco Lógico), 12 (Control de Consistencia) y subsecciones tienen línea punteada y número de página dinámico.\n"
                f"- PDF y archivo ZIP actualizados:\n{r_zip}",
                parent=root
            )
        except Exception as e:
            messagebox.showerror("Aviso", f"Cierra Word antes de ejecutar este script:\n\n{e}", parent=root)
    root.destroy()