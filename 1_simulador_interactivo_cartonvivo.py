import os
import sys
import copy
import shutil
import zipfile
import subprocess
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import numpy_financial as npf

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==============================================================================
# BLOQUE 1: DATOS REDONDOS, REALISTAS Y FÁCILES DE DEFENDER (OSORNO)
# ==============================================================================

DEFAULTS = {
    "NOMBRE_PROYECTO": "Cartón Vivo Osorno",
    "INTEGRANTES": "MAURICIO ALEJANDRO VARGAS ORTEGA, OLIVER MUÑOZ, JOSUE JOFRE",
    "APELLIDOS_ENTREGA": "Vargas,Muñoz,Jofre",
    "DOCENTE": "Dra. Luz María Ferrada B.",
    "HORIZONTE_ANIOS": 10,
    "IVA": 0.19,
    "IMPUESTO_RENTA": 0.25,
    "TASA_DESCUENTO_PRIVADA": 0.12,
    "TASA_DESCUENTO_SOCIAL": 0.06,
    "DIAS_DESFASE_CT": 45,

    # Factores Sociales MIDESO (Mano de Obra y Divisa)
    "FC_MOC": 1.00,
    "FC_MOSC": 0.85,
    "FC_MONC": 0.68,
    "FC_DIVISA": 1.01,
    "ARANCEL_IMP": 0.06,

    # Recolección: 50 kg al día * 25 días al mes = 1.250 kg/mes = 15,00 toneladas al año
    "RESIDUO_TON_ANIO_1": 15.00,
    "TASA_CRECIMIENTO": 0.05,

    # Producto 1: 3.000 unidades * 2,50 kg = 7,50 toneladas de cartón
    "P1_NOMBRE": "Set Didáctico Escolar",
    "P1_UNIDADES": 3000,
    "P1_PRECIO_BRUTO": 16000,
    "P1_KG_CARTON": 2.50,

    # Producto 2: 1.500 unidades * 2,00 kg = 3,00 toneladas de cartón
    # (7,50 ton + 3,00 ton = 10,50 ton usadas = 70% del cartón recolectado; 30% es merma)
    "P2_NOMBRE": "Juegos Cognitivos Adulto Mayor",
    "P2_UNIDADES": 1500,
    "P2_PRECIO_BRUTO": 19000,
    "P2_KG_CARTON": 2.00,

    # Beneficios Sociales (Ahorro municipal Vertedero Curaco + Beneficio ambiental/social)
    "AHORRO_MUNI_TON": 48000,
    "OTROS_BEN_SOC_ANUAL": 2500000,
}

# Inversión Inicial Año 0 = $16.000.000 con IVA ($8,5M + $3,5M + $2,0M + $2,0M)
INVERSIONES_DEFAULT = [
    {
        "item": "Máquina Troqueladora Industrial de Rodillos (1200 mm)",
        "monto_bruto": 8500000,
        "vida_util_contable": 10,
        "vida_util_economica": 10,
        "valor_residual_pct": 0.10,
        "origen": "Importado"
    },
    {
        "item": "Guillotina Cortadora, Mesones y Moldes de Corte",
        "monto_bruto": 3500000,
        "vida_util_contable": 5,
        "vida_util_economica": 5,
        "valor_residual_pct": 0.05,
        "origen": "Nacional"
    },
    {
        "item": "Equipo de Pintura al Agua (Serigrafía) y Estantes de Secado",
        "monto_bruto": 2000000,
        "vida_util_contable": 5,
        "vida_util_economica": 5,
        "valor_residual_pct": 0.05,
        "origen": "Importado"
    },
    {
        "item": "Acondicionamiento Galpón en Ovejería (Electricidad y Permiso SEREMI)",
        "monto_bruto": 2000000,
        "vida_util_contable": 5,
        "vida_util_economica": 10,
        "valor_residual_pct": 0.00,
        "origen": "Nacional"
    }
]

COSTOS_VAR_DEFAULT = [
    {
        "item": "Pinturas ecológicas al agua no tóxicas, pegamento y barniz",
        "costo": 1800,
        "base": "Por Unidad",
        "origen": "Nacional"
    },
    {
        "item": "Caja de presentación, instructivo de uso y etiqueta",
        "costo": 600,
        "base": "Por Unidad",
        "origen": "Nacional"
    },
    {
        "item": "Flete furgón retiro calle Eleuterio Ramírez (#446 al #1024) antes de 06:00 AM",
        "costo": 30000,
        "base": "Por Tonelada",
        "origen": "Nacional"
    }
]

COSTOS_FIJOS_DEFAULT = [
    {
        "item": "Arriendo de galpón de trabajo 140 m² (Sector Ovejería, Osorno)",
        "mensual": 450000,
        "afecto_iva": "Exento"
    },
    {
        "item": "Luz trifásica (Saesa), agua potable (Suralis) e internet",
        "mensual": 180000,
        "afecto_iva": "Con IVA"
    },
    {
        "item": "Contador, patente municipal, seguros y publicidad",
        "mensual": 150000,
        "afecto_iva": "Con IVA"
    }
]

PERSONAL_DEFAULT = [
    {
        "cargo": "Jefe de Planta, Control de Calidad y Ventas",
        "cantidad": 1,
        "sueldo_bruto_mensual": 950000,
        "calificacion": "MOC",
        "competencias": "Administración de la planta, relación con DAEM, jardines y clientes"
    },
    {
        "cargo": "Operador Técnico de Corte y Pintura",
        "cantidad": 2,
        "sueldo_bruto_mensual": 600000,
        "calificacion": "MOSC",
        "competencias": "Uso de máquina troqueladora, guillotina y pintado con tintas al agua"
    },
    {
        "cargo": "Operario de Recolección, Limpieza y Armado",
        "cantidad": 2,
        "sueldo_bruto_mensual": 520000,
        "calificacion": "MONC",
        "competencias": "Retiro de cajas en calle Eleuterio Ramírez, limpieza y empaque"
    }
]


# ==============================================================================
# BLOQUE 2: MOTOR FINANCIERO Y TABLAS EXPLICATIVAS EN LENGUAJE SENCILLO
# ==============================================================================

def fmt_clp(valor):
    if isinstance(valor, (int, float, np.integer, np.floating)):
        if valor < 0:
            return f"-${abs(int(round(valor))):,}".replace(",", ".")
        return f"${int(round(valor)):,}".replace(",", ".")
    return str(valor)


def calcular_payback(flujos):
    acumulado = np.cumsum(flujos)
    anios_positivos = np.where(acumulado >= 0)[0]
    if len(anios_positivos) == 0:
        return None
    idx = anios_positivos[0]
    if idx == 0:
        return 0.0
    faltante = abs(acumulado[idx - 1])
    flujo_anio = flujos[idx]
    if flujo_anio == 0:
        return None
    return round((idx - 1) + (faltante / flujo_anio), 2)


def armar_tablas_informantes_y_supuestos(cfg):
    horizonte = int(cfg["HORIZONTE_ANIOS"])
    iva = float(cfg["IVA"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = int(cfg["DIAS_DESFASE_CT"])
    crec = float(cfg["TASA_CRECIMIENTO"])
    ton_a1 = float(cfg["RESIDUO_TON_ANIO_1"])
    escala_ton = ton_a1 / 15.00 if ton_a1 > 0 else 1.0

    base_locales = [
        ("LOC-01", "Hojalatería Lynch", "Eleuterio Ramírez #446, Osorno", "Ferretería y hojalatería",
         "Cajas gruesas de herramientas y tubos; se retiran a las 05:45 AM antes de que pase el camión de la basura (06:00-07:00 AM)", 4.0),
        ("LOC-02", "Autoplanet Osorno", "Eleuterio Ramírez #582, Osorno", "Repuestos de autos",
         "Cajas grandes y duras de repuestos automotrices limpias y secas; retiro temprano antes del camión municipal", 8.0),
        ("LOC-03", "Zapatería 'La Malagueña'", "Eleuterio Ramírez #625, Casilla 600, Osorno", "Venta de calzado",
         "Cajas grandes donde llegan los lotes de zapatos, secas y en buen estado; se salvan de ir al Vertedero de Curaco", 4.0),
        ("LOC-04", "Minimarket y Frutería 'El Sabrosón'", "Eleuterio Ramírez #647, Osorno", "Minimarket y abarrotes",
         "Cajas de cartón secas de mercadería y abarrotes guardadas en bodega para nuestro retiro matutino", 6.0),
        ("LOC-05", "Galería Centrosorno", "Eleuterio Ramírez #977, Osorno", "Galería comercial (tiendas interiores)",
         "Cajas juntadas por más de 15 locales de ropa y óptica dentro de la galería antes del paso del camión recolector", 14.0),
        ("LOC-06", "Liquimax - Osorno", "Eleuterio Ramírez #1024, Local 3, Osorno", "Mayorista de aseo y hogar",
         "Gran cantidad diaria de cajas grandes de productos de aseo; es el local que más cartón entrega en la cuadra", 14.0),
    ]

    filas_cat = []
    for cod, nom, direc, rubro, desc, kg_dia_base in base_locales:
        kg_dia = round(kg_dia_base * escala_ton, 2)
        kg_mes = round(kg_dia * 25.0)
        ton_anual = round((kg_mes * 12.0) / 1000.0, 2)
        filas_cat.append([cod, nom, direc, rubro, desc, kg_dia, kg_mes, ton_anual])

    df_catastro_6 = pd.DataFrame(filas_cat, columns=[
        "Código Local",
        "Nombre del Comercio Real (Osorno)",
        "Dirección en Calle Eleuterio Ramírez",
        "Rubro del Local",
        "Tipo de Cartón y Forma de Retiro (Antes del Camión de 06:00-07:00 AM)",
        "Kg por Día",
        "Kg al Mes (25 días)",
        "Toneladas al Año"
    ])

    u1_tot = int(cfg["P1_UNIDADES"])
    u1_daem = int(round(u1_tot * (2.0 / 3.0)))
    u1_priv = u1_tot - u1_daem
    u2_tot = int(cfg["P2_UNIDADES"])

    df_demanda_3 = pd.DataFrame([
        [
            "COMP-01",
            "DAEM Osorno (Escuelas Municipales y Jardines VTF)",
            "Encargados de Educación Parvularia y Compras SEP del Municipio",
            cfg["P1_NOMBRE"],
            u1_daem,
            int(cfg["P1_PRECIO_BRUTO"]),
            "Compras anuales por Mercado Público (Compra Ágil) con fondos SEP para ~18 escuelas y 11 jardines municipales de Osorno."
        ],
        [
            "COMP-02",
            "Jardines Infantiles Privados, Guarderías y Salas Cuna de Osorno",
            "Directoras de ~25 Jardines Particulares y Guarderías de la ciudad",
            cfg["P1_NOMBRE"],
            u1_priv,
            int(cfg["P1_PRECIO_BRUTO"]),
            "Venta directa (~40 sets al año por jardín) de material educativo ecológico seguro para niños y más barato que la madera."
        ],
        [
            "COMP-03",
            "Hogar de Ancianos Santa María de Osorno y Clubes de Adulto Mayor",
            "Terapeutas Ocupacionales del Hogar Santa María y Monitores de Clubes",
            cfg["P2_NOMBRE"],
            u2_tot,
            int(cfg["P2_PRECIO_BRUTO"]),
            "Juegos de memoria y destreza manual con piezas grandes y fáciles de tomar para residentes y más de 90 clubes de adulto mayor."
        ]
    ], columns=[
        "Código",
        "Cliente / Institución Compradora en Osorno",
        "Con Quién se Validó el Dato (Entrevista / Consulta)",
        "Producto que Compra",
        "Unidades Año 1",
        "Precio con IVA ($)",
        "Cómo Compra y Por Qué Nos Elige"
    ])

    df_informantes = pd.DataFrame([
        [
            "1. Encuesta y Aforo en Calle Eleuterio Ramírez (6 Locales Reales)",
            "Hojalatería Lynch (#446), Autoplanet (#582), Zapatería La Malagueña (#625), Minimarket El Sabrosón (#647), Galería Centrosorno (#977) y Liquimax (#1024 L.3)",
            "Observación directa, entrevista a encargados de local y pesaje de cajas (Oct-Nov 2026)",
            f"El camión de la basura pasa entre 06:00 y 07:00 AM llevando el cartón al Vertedero de Curaco. Entre los 6 locales juntan 50 kg al día = 1.250 kg al mes ({ton_a1:.2f} toneladas al año)."
        ],
        [
            "2. Entrevista Comprador 1: DAEM Osorno (Escuelas y Jardines VTF)",
            "Escuelas Básicas y Jardines VTF administrados por la Municipalidad de Osorno",
            "Entrevista semiestructurada y revisión de órdenes de compra DAEM (Compra Ágil / SEP)",
            f"Demanda de {u1_daem:,} unidades al año de {cfg['P1_NOMBRE']} a {fmt_clp(cfg['P1_PRECIO_BRUTO'])} con IVA ({fmt_clp(cfg['P1_PRECIO_BRUTO']/(1+iva))} sin IVA)."
        ],
        [
            "3. Encuesta Comprador 2: Jardines Infantiles Privados y Guarderías",
            "Jardines particulares, salas cuna y guarderías de Osorno (Centro, Oriente y Rahue)",
            "Encuesta breve a directoras y educadoras de párvulos (Octubre 2026)",
            f"Demanda de {u1_priv:,} unidades al año de {cfg['P1_NOMBRE']}, sumando en total {u1_tot:,} unidades anuales del Producto 1."
        ],
        [
            "4. Entrevista Comprador 3: Hogar de Ancianos Santa María y Clubes",
            "Hogar Santa María de Osorno y agrupaciones de adulto mayor de la comuna",
            "Entrevista sobre materiales para talleres de memoria y terapia ocupacional",
            f"Demanda de {u2_tot:,} unidades al año de {cfg['P2_NOMBRE']} a {fmt_clp(cfg['P2_PRECIO_BRUTO'])} con IVA ({fmt_clp(cfg['P2_PRECIO_BRUTO']/(1+iva))} sin IVA)."
        ],
        [
            "5. Ahorro para la Municipalidad (Vertedero de Curaco)",
            "Datos públicos de aseo municipal en Osorno (MercadoPúblico) y reportes SINADER",
            "Revisión de costos municipales por tonelada de basura recolectada y enterrada",
            f"Por cada tonelada de cartón que retiramos antes que el camión municipal, el municipio ahorra {fmt_clp(cfg['AHORRO_MUNI_TON'])} con IVA ({fmt_clp(cfg['AHORRO_MUNI_TON']/(1+iva))} sin IVA)."
        ],
        [
            "6. Cotizaciones de Máquinas, Arriendo y Sueldos en Osorno",
            "Proveedores de máquinas de corte, arriendo en Ovejería y tarifas de Saesa/Suralis",
            "Cotizaciones formales de mercado (Octubre-Noviembre 2026)",
            "Inversión total de $16.000.000 con IVA en 4 ítems redondos, arriendo de galpón en Ovejería por $450.000 al mes y flete por $30.000 la tonelada."
        ]
    ], columns=[
        "Instrumento / Fuente de Información",
        "Lugares e Instituciones en Osorno",
        "Método Aplicado (Entrevista, Encuesta, Aforo o Cotización)",
        "Dato Clave Obtenido para el Proyecto"
    ])

    ton_inc_a1 = (u1_tot * float(cfg["P1_KG_CARTON"]) + u2_tot * float(cfg["P2_KG_CARTON"])) / 1000.0
    merma_ton_a1 = max(0.0, ton_a1 - ton_inc_a1)
    pct_util = (ton_inc_a1 / ton_a1 * 100.0) if ton_a1 > 0 else 0.0
    pct_merma = 100.0 - pct_util

    df_supuestos = pd.DataFrame([
        [
            "Uso del Cartón y Merma (70% se usa / 30% se pierde)",
            f"Se juntan {ton_a1:.2f} ton | Se usan {ton_inc_a1:.2f} ton",
            f"De las {ton_a1:.2f} toneladas recolectadas al año, el {pct_util:.0f}% ({ton_inc_a1:.2f} ton) queda convertido en los {u1_tot+u2_tot:,} juegos ({u1_tot}x2,5kg + {u2_tot}x2,0kg). El otro {pct_merma:.0f}% ({merma_ton_a1:.2f} ton) son recortes sobrantes del molde y partes con cinta o humedad.",
            "Cálculo práctico de aprovechamiento de material en taller de troquelado."
        ],
        [
            "Duración del Proyecto y Crecimiento",
            f"{horizonte} años / Crece {crec*100:.0f}% por año",
            "Las ventas y la recolección suben un 5% cada año (equivale a sumar 1 local comercial más por año y vender unos 225 juegos adicionales cada año).",
            "Crecimiento moderado y realista para una pequeña empresa en Osorno."
        ],
        [
            "Impuestos (IVA e Impuesto a la Renta)",
            f"IVA {iva*100:.0f}% | Renta {imp_renta*100:.0f}%",
            "En el Flujo Privado todos los valores con IVA se dividen por 1,19 para trabajar con valores netos. Sobre la ganancia anual se paga un 25% de Impuesto a la Renta.",
            "Normativa del SII en Chile para Pymes (Régimen Pro Pyme General)."
        ],
        [
            "Capital de Trabajo (Dinero para funcionar)",
            f"{dias_ct} días de costos",
            f"Se calcula como: (Costos Anuales * {dias_ct} días) / 365. Sirve para pagar sueldos, arriendo e insumos mientras el DAEM y los jardines demoran entre 30 y 45 días en pagar las facturas.",
            "Método del Período de Desfase (45 días de respaldo)."
        ],
        [
            "Tasas de Descuento (Privada vs. Social)",
            f"Privada: {tasa_priv*100:.0f}% | Social: {tasa_soc*100:.0f}%",
            "Al dueño privado se le exige una rentabilidad mínima del 12% anual. Para el país y la comuna (Evaluación Social) el Ministerio de Desarrollo Social exige un 6% anual.",
            "Tasa Social de Descuento oficial de MIDESO / SNI Chile."
        ],
        [
            "Corrección Social (Precios Sombra MIDESO)",
            f"MOC={cfg['FC_MOC']} | MOSC={cfg['FC_MOSC']} | MONC={cfg['FC_MONC']}",
            "En el Flujo Social se quitan los impuestos (porque son transferencias internas dentro del país) y los sueldos se multiplican por los factores oficiales de MIDESO según calificación.",
            "Metodología oficial del Sistema Nacional de Inversiones (SNI)."
        ]
    ], columns=[
        "Concepto / Supuesto del Modelo",
        "Valor Adoptado",
        "Explicación Sencilla de Cómo Funciona",
        "Respaldo Técnico / Legal"
    ])

    df_cotizaciones = pd.DataFrame([
        ["Nota [1]", "Máquina Principal", "Máquina Troqueladora Industrial de Rodillos (1200 mm)", "Importado", 8500000, round(8500000 / (1 + iva)), "Cotización Proveedor Maquinaria Gráfica Industrial Chile (Oct 2026)"],
        ["Nota [1]", "Corte y Moldes", "Guillotina Cortadora, Mesones y Moldes de Corte", "Nacional", 3500000, round(3500000 / (1 + iva)), "Cotización Fabricante Nacional de Matrices y Equipos (Oct 2026)"],
        ["Nota [1]", "Pintura y Secado", "Equipo de Pintura al Agua (Serigrafía) y Estantes de Secado", "Importado", 2000000, round(2000000 / (1 + iva)), "Cotización Proveedor Equipos Serigráficos (Oct 2026)"],
        ["Nota [1]", "Arreglo de Local", "Acondicionamiento Galpón en Ovejería (Electricidad y Permiso SEREMI)", "Nacional", 2000000, round(2000000 / (1 + iva)), "Presupuesto Técnico Eléctrico SEC y Obras Menores en Osorno"],
        ["Nota [2]", "Insumo por Juego", "Pinturas ecológicas al agua no tóxicas (EN-71), pegamento y barniz", "Nacional", 1800, round(1800 / (1 + iva)), "Lista de Precios Mayorista Insumos Gráficos Certificados"],
        ["Nota [2]", "Envase por Juego", "Caja de presentación microcorrugada, instructivo de uso y etiqueta", "Nacional", 600, round(600 / (1 + iva)), "Cotización Imprenta y Envases Zona Sur Osorno"],
        ["Nota [3]", "Transporte Cartón", "Flete furgón retiro calle Eleuterio Ramírez (#446 al #1024) antes de 06:00 AM", "Nacional", 30000, round(30000 / (1 + iva)), "Tarifa Furgón de Reparto Local Ruta Ramírez -> Ovejería"],
        ["Nota [4]", "Gasto Fijo Mensual", "Arriendo de galpón de trabajo 140 m² (Sector Ovejería, Osorno)", "Exento", 450000, 450000, "Valor Mercado Corredora de Propiedades Osorno — Sector Ovejería"],
        ["Nota [4]", "Gasto Fijo Mensual", "Luz trifásica (Saesa), agua potable (Suralis) e internet", "Con IVA", 180000, round(180000 / (1 + iva)), "Tarifas Comerciales Vigentes Grupo Saesa y Suralis Osorno"],
        ["Nota [4]", "Gasto Fijo Mensual", "Contador, patente municipal, seguros y publicidad", "Con IVA", 150000, round(150000 / (1 + iva)), "Honorarios Contabilidad Pro Pyme y Patente Municipalidad de Osorno"],
        ["Nota [5]", "Ahorro Municipal", "Gasto municipal ahorrado por tonelada que no va al Vertedero de Curaco", "Con IVA", int(cfg["AHORRO_MUNI_TON"]), round(float(cfg["AHORRO_MUNI_TON"]) / (1 + iva)), "Contratos de Aseo Municipalidad de Osorno / Reportes SINADER"]
    ], columns=[
        "Nota al Pie",
        "Categoría",
        "Descripción del Ítem Cotizado",
        "Origen / IVA",
        "Valor con IVA ($)",
        "Valor Neto sin IVA ($)",
        "Fuente / Referencia de Precio o Costo en Osorno"
    ])

    df_porque = pd.DataFrame([
        [
            "1. ¿Por qué 15 toneladas de cartón al año?",
            "50 kg al día × 25 días × 12 meses",
            "Porque entre los 6 locales de calle Eleuterio Ramírez juntan 50 kilos diarios (Lynch 4 kg + Autoplanet 8 kg + La Malagueña 4 kg + El Sabrosón 6 kg + Galería Centrosorno 14 kg + Liquimax 14 kg = 50 kg). Trabajando 25 días al mes son 1.250 kilos mensuales, que en 12 meses dan exactamente 15.000 kilos (15 toneladas)."
        ],
        [
            "2. ¿Por qué se usan 10,5 toneladas (70%) y sobran 4,5 toneladas (30%)?",
            "7,5 ton + 3,0 ton = 10,5 ton",
            "Fabricamos 3.000 Sets Escolares (usan 2,5 kg c/u = 7,5 toneladas) y 1.500 Juegos de Adulto Mayor (usan 2,0 kg c/u = 3,0 toneladas). En total los productos llevan 10,5 toneladas (el 70%). Las otras 4,5 toneladas (el 30%) son los bordes con corchetes, cintas plásticas y los recortes sobrantes al pasar el cartón por el molde."
        ],
        [
            "3. ¿Por qué se venden 3.000 Sets Escolares y 1.500 Juegos de Adulto Mayor?",
            "4.500 juegos al año (375 al mes)",
            "Son solo 15 juegos fabricados por día hábil. De los 3.000 Sets Escolares, 2.000 los compra el DAEM de Osorno para sus escuelas y jardines VTF, y 1.000 los compran unos 25 jardines privados de Osorno (40 sets por jardín al año). Los 1.500 Juegos Cognitivos van al Hogar Santa María y a los más de 90 clubes de adulto mayor de la comuna."
        ],
        [
            "4. ¿Por qué cuestan $16.000 y $19.000 con IVA?",
            "$13.445 y $15.966 sin IVA",
            "Porque los juegos didácticos de madera o plástico importado cuestan entre $22.000 y $27.000. Al cobrar $16.000 y $19.000 con IVA somos un 30% más baratos y además ecológicos, lo que facilita ganar las compras públicas."
        ],
        [
            "5. ¿Por qué la inversión inicial es de $16.000.000 con IVA?",
            "$8,5M + $3,5M + $2,0M + $2,0M",
            "Son 4 compras cerradas: 1) Máquina troqueladora de rodillos ($8.500.000), 2) Guillotina cortadora, mesones y moldes ($3.500.000), 3) Equipo de pintado al agua y secado ($2.000.000) y 4) Instalación eléctrica trifásica y permiso sanitario del galpón en Ovejería ($2.000.000)."
        ],
        [
            "6. ¿Por qué el costo variable es $2.400 por juego + $30.000 por tonelada de flete?",
            "$1.800 pintura + $600 caja",
            "En cada juego se gastan $1.800 (con IVA) en pintura al agua no tóxica, pegamento y barniz protector, más $600 (con IVA) en su caja con instrucciones. El furgón cobra $30.000 por tonelada trasladada desde calle Ramírez hasta Ovejería ($450.000 en todo el Año 1)."
        ],
        [
            "7. ¿Por qué los gastos fijos son $780.000/mes y los sueldos $3.190.000/mes?",
            "Galpón en Ovejería + 5 trabajadores",
            "Gastos fijos mensuales: arriendo de galpón de 140 m² en Ovejería ($450.000), luz/agua/internet ($180.000) y contador/patente/seguros ($150.000). Trabajadores (5 en total): 1 Jefe de Planta ($950.000), 2 Operadores de Máquina ($600.000 c/u) y 2 Operarios de Recolección y Armado ($520.000 c/u)."
        ],
        [
            "8. ¿Por qué el proyecto es tan rentable para el privado y para Osorno?",
            "VAN Privado 12% vs. VANS Social 6%",
            "Para el inversionista privado da ganancias desde el primer año y recupera la inversión rápido. Para la comuna de Osorno (Evaluación Social al 6%) es aún mejor porque ahorra $48.000 por cada tonelada que no se entierra en el Vertedero de Curaco y da empleo local."
        ]
    ], columns=[
        "Pregunta Clave de la Comisión",
        "Cálculo Rápido",
        "Explicación Fácil de Entender (Para tu Defensa y Entrega)"
    ])

    return df_informantes, df_supuestos, df_catastro_6, df_demanda_3, df_cotizaciones, df_porque


def ejecutar_motor_financiero(cfg):
    horizonte = max(1, int(cfg["HORIZONTE_ANIOS"]))
    iva = float(cfg["IVA"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = float(cfg["DIAS_DESFASE_CT"])

    fc_moc = float(cfg["FC_MOC"])
    fc_mosc = float(cfg["FC_MOSC"])
    fc_monc = float(cfg["FC_MONC"])
    fc_div = float(cfg["FC_DIVISA"])
    arancel = float(cfg["ARANCEL_IMP"])

    def factor_social_bien(origen):
        neto = 1.0 / (1.0 + iva)
        if str(origen).strip().lower() == "importado":
            return (neto / (1.0 + arancel)) * fc_div
        return neto

    anios_op = list(range(1, horizonte + 1))
    anios_tot = list(range(0, horizonte + 1))
    cols_op = [f"Año {a}" for a in anios_op]
    cols_tot = [f"Año {a}" for a in anios_tot]

    lineas_productos = [
        {
            "nombre": cfg["P1_NOMBRE"],
            "unidades_anio_1": float(cfg["P1_UNIDADES"]),
            "precio_venta_bruto": float(cfg["P1_PRECIO_BRUTO"]),
            "kg_carton_por_unidad": float(cfg["P1_KG_CARTON"])
        },
        {
            "nombre": cfg["P2_NOMBRE"],
            "unidades_anio_1": float(cfg["P2_UNIDADES"]),
            "precio_venta_bruto": float(cfg["P2_PRECIO_BRUTO"]),
            "kg_carton_por_unidad": float(cfg["P2_KG_CARTON"])
        }
    ]

    inversiones = cfg["INVERSIONES_LISTA"]
    costos_var_lista = cfg["COSTOS_VAR_LISTA"]
    costos_fijos_lista = cfg["COSTOS_FIJOS_LISTA"]
    personal = cfg["PERSONAL_LISTA"]

    mapa_fc_mo = {"MOC": fc_moc, "MOSC": fc_mosc, "MONC": fc_monc}
    filas_personal = []
    costo_anual_personal_privado = 0.0
    costo_anual_personal_social = 0.0

    for p in personal:
        anual_priv = p["cantidad"] * p["sueldo_bruto_mensual"] * 12.0
        fc = mapa_fc_mo.get(p["calificacion"], 1.0)
        anual_soc = anual_priv * fc
        costo_anual_personal_privado += anual_priv
        costo_anual_personal_social += anual_soc
        filas_personal.append({
            "Cargo / Puesto": p["cargo"],
            "Cant.": p["cantidad"],
            "Calif. SNI": p["calificacion"],
            "Tareas Principales": p["competencias"],
            "Sueldo Mensual ($)": round(p["sueldo_bruto_mensual"]),
            "Costo Anual Privado ($)": round(anual_priv),
            "Factor Social": fc,
            "Costo Anual Social ($)": round(anual_soc)
        })
    df_personal = pd.DataFrame(filas_personal)

    dict_cantidades = {}
    dict_ingresos_priv = {}
    dict_ingresos_soc = {}

    residuo_1 = float(cfg["RESIDUO_TON_ANIO_1"])
    crec = float(cfg["TASA_CRECIMIENTO"])

    ton_residuo_anual = [residuo_1 * ((1.0 + crec) ** (a - 1)) for a in anios_op]
    dict_cantidades["Residuo de Cartón Recolectado (Toneladas)"] = [round(x, 2) for x in ton_residuo_anual]

    unidades_totales_anual = np.zeros(horizonte)
    kg_utilizados_anual = np.zeros(horizonte)

    for prod in lineas_productos:
        unids = np.array([round(prod["unidades_anio_1"] * ((1.0 + crec) ** (a - 1))) for a in anios_op])
        unidades_totales_anual += unids
        kg_utilizados_anual += unids * prod["kg_carton_por_unidad"]

        precio_neto = prod["precio_venta_bruto"] / (1.0 + iva)
        ing_neto_anual = np.round(unids * precio_neto)

        dict_cantidades[f"Producción {prod['nombre']} (Unidades)"] = unids.astype(int).tolist()
        dict_ingresos_priv[f"Ventas Netas {prod['nombre']} (Sin IVA)"] = ing_neto_anual.astype(int).tolist()
        dict_ingresos_soc[f"Beneficio por Ventas {prod['nombre']} (Social)"] = ing_neto_anual.astype(int).tolist()

    dict_cantidades["Total Productos Fabricados (Unidades)"] = unidades_totales_anual.astype(int).tolist()
    dict_cantidades["Cartón Efectivamente Incorporado (Toneladas)"] = np.round(kg_utilizados_anual / 1000.0, 2).tolist()

    df_cantidades = pd.DataFrame.from_dict(dict_cantidades, orient="index", columns=cols_op)
    df_cantidades.index.name = "Ítem Físico"

    ahorro_muni_neto_ton = float(cfg["AHORRO_MUNI_TON"]) / (1.0 + iva)
    beneficio_ahorro_muni = np.round(np.array(ton_residuo_anual) * ahorro_muni_neto_ton)
    beneficio_otros_soc = np.round(np.full(horizonte, float(cfg["OTROS_BEN_SOC_ANUAL"])))

    dict_ingresos_soc["Liberación de Recursos Municipales (Ahorro Vertedero)"] = beneficio_ahorro_muni.astype(int).tolist()
    dict_ingresos_soc["Otros Beneficios Sociales y Ambientales"] = beneficio_otros_soc.astype(int).tolist()

    df_ingresos = pd.DataFrame.from_dict(dict_ingresos_priv, orient="index", columns=cols_op)
    df_ingresos.loc["TOTAL INGRESOS PRIVADOS NETOS (SIN IVA)"] = df_ingresos.sum(axis=0)
    df_ingresos.index.name = "Concepto de Ingreso"

    dict_costos_priv = {}
    dict_costos_soc = {}

    for cv in costos_var_lista:
        c_neto = cv["costo"] / (1.0 + iva)
        c_soc = cv["costo"] * factor_social_bien(cv["origen"])
        if cv["base"] == "Por Tonelada":
            base_arr = np.array(ton_residuo_anual)
            pref_p = "CV Logístico"
            pref_s = "CV Logístico Social"
        else:
            base_arr = unidades_totales_anual
            pref_p = "CV"
            pref_s = "CV Social"

        dict_costos_priv[f"{pref_p}: {cv['item']}"] = np.round(base_arr * c_neto).astype(int).tolist()
        dict_costos_soc[f"{pref_s}: {cv['item']}"] = np.round(base_arr * c_soc).astype(int).tolist()

    for cf in costos_fijos_lista:
        anual_bruto = cf["mensual"] * 12.0
        lleva_iva = (cf["afecto_iva"] == "Con IVA")
        anual_neto = (anual_bruto / (1.0 + iva)) if lleva_iva else anual_bruto
        dict_costos_priv[f"CF: {cf['item']}"] = np.round(np.full(horizonte, anual_neto)).astype(int).tolist()
        dict_costos_soc[f"CF Social: {cf['item']}"] = np.round(np.full(horizonte, anual_neto)).astype(int).tolist()

    dict_costos_priv["Remuneraciones del Personal (Planta y Administración)"] = np.round(np.full(horizonte, costo_anual_personal_privado)).astype(int).tolist()
    dict_costos_soc["Mano de Obra Valorizada a Precios Sociales (MIDESO)"] = np.round(np.full(horizonte, costo_anual_personal_social)).astype(int).tolist()

    df_costos_priv = pd.DataFrame.from_dict(dict_costos_priv, orient="index", columns=cols_op)
    df_costos_priv.loc["TOTAL COSTOS OPERATIVOS PRIVADOS (SIN IVA)"] = df_costos_priv.sum(axis=0)
    df_costos_priv.index.name = "Ítem de Costo"

    df_costos_soc = pd.DataFrame.from_dict(dict_costos_soc, orient="index", columns=cols_op)
    df_costos_soc.loc["TOTAL COSTOS OPERATIVOS SOCIALES"] = df_costos_soc.sum(axis=0)

    dict_inv_priv = {}
    dict_inv_soc = {}
    dict_depreciacion = {}
    valor_desecho_priv_total = 0.0
    valor_desecho_soc_total = 0.0

    for inv in inversiones:
        fila_p = np.zeros(horizonte + 1)
        fila_s = np.zeros(horizonte + 1)
        fila_d = np.zeros(horizonte)

        monto_neto_priv = inv["monto_bruto"] / (1.0 + iva)
        monto_soc = inv["monto_bruto"] * factor_social_bien(inv["origen"])

        fila_p[0] = monto_neto_priv
        fila_s[0] = monto_soc

        vida_econ = max(1, int(inv["vida_util_economica"]))
        for anio in range(vida_econ, horizonte, vida_econ):
            fila_p[anio] = monto_neto_priv
            fila_s[anio] = monto_soc

        vida_cont = max(1, int(inv["vida_util_contable"]))
        dep_anual = (monto_neto_priv * (1.0 - inv["valor_residual_pct"])) / vida_cont
        for a_idx in range(horizonte):
            edad_activo = ((a_idx) % vida_econ) + 1
            if edad_activo <= vida_cont:
                fila_d[a_idx] = dep_anual

        valor_desecho_priv_total += monto_neto_priv * inv["valor_residual_pct"]
        valor_desecho_soc_total += monto_soc * inv["valor_residual_pct"]

        dict_inv_priv[inv["item"]] = np.round(fila_p).astype(int).tolist()
        dict_inv_soc[inv["item"]] = np.round(fila_s).astype(int).tolist()
        dict_depreciacion[f"Depreciación: {inv['item']}"] = np.round(fila_d).astype(int).tolist()

    df_inversiones = pd.DataFrame.from_dict(dict_inv_priv, orient="index", columns=cols_tot)
    df_inversiones.loc["TOTAL INVERSIONES Y REINVERSIONES NETAS"] = df_inversiones.sum(axis=0)
    df_inversiones.index.name = "Activo / Ítem de Inversión"

    df_inv_social = pd.DataFrame.from_dict(dict_inv_soc, orient="index", columns=cols_tot)
    df_inv_social.loc["TOTAL INVERSIONES Y REINVERSIONES SOCIALES"] = df_inv_social.sum(axis=0)

    df_depreciacion = pd.DataFrame.from_dict(dict_depreciacion, orient="index", columns=cols_op)
    df_depreciacion.loc["TOTAL DEPRECIACIÓN ANUAL"] = df_depreciacion.sum(axis=0)
    df_depreciacion.index.name = "Ítem de Depreciación"

    fraccion_anual_ct = dias_ct / 365.0
    costos_tot_priv_anual = df_costos_priv.loc["TOTAL COSTOS OPERATIVOS PRIVADOS (SIN IVA)"].values
    costos_tot_soc_anual = df_costos_soc.loc["TOTAL COSTOS OPERATIVOS SOCIALES"].values

    ct_requerido_priv = np.round(costos_tot_priv_anual * fraccion_anual_ct)
    ct_requerido_soc = np.round(costos_tot_soc_anual * fraccion_anual_ct)

    flujo_ct_priv = np.zeros(horizonte + 1)
    flujo_ct_soc = np.zeros(horizonte + 1)

    flujo_ct_priv[0] = -ct_requerido_priv[0]
    flujo_ct_soc[0] = -ct_requerido_soc[0]

    for i in range(1, horizonte):
        flujo_ct_priv[i] = -(ct_requerido_priv[i] - ct_requerido_priv[i - 1])
        flujo_ct_soc[i] = -(ct_requerido_soc[i] - ct_requerido_soc[i - 1])

    flujo_ct_priv[horizonte] = ct_requerido_priv[-1]
    flujo_ct_soc[horizonte] = ct_requerido_soc[-1]

    df_capital_trabajo = pd.DataFrame({
        cols_tot[a]: [
            0 if a == 0 else int(costos_tot_priv_anual[a - 1]),
            0 if a == 0 else int(ct_requerido_priv[a - 1]),
            int(flujo_ct_priv[a])
        ]
        for a in range(horizonte + 1)
    }, index=[
        "Costos Operativos Anuales Netos ($)",
        f"Capital de Trabajo Requerido ({int(dias_ct)} días)",
        "Flujo de Variación / Recuperación de Capital de Trabajo"
    ])
    df_capital_trabajo.index.name = "Concepto Capital de Trabajo"

    ing_priv_0_10 = np.insert(df_ingresos.loc["TOTAL INGRESOS PRIVADOS NETOS (SIN IVA)"].values, 0, 0)
    costos_priv_0_10 = np.insert(-costos_tot_priv_anual, 0, 0)
    dep_0_10 = np.insert(-df_depreciacion.loc["TOTAL DEPRECIACIÓN ANUAL"].values, 0, 0)

    uai = ing_priv_0_10 + costos_priv_0_10 + dep_0_10
    impuestos = np.where(uai > 0, -np.round(uai * imp_renta), 0)
    udi = uai + impuestos
    reverso_dep = -dep_0_10

    inv_fija_priv = -df_inversiones.loc["TOTAL INVERSIONES Y REINVERSIONES NETAS"].values
    val_desecho_priv_arr = np.zeros(horizonte + 1)
    val_desecho_priv_arr[-1] = round(valor_desecho_priv_total)

    flujo_caja_privado = udi + reverso_dep + inv_fija_priv + flujo_ct_priv + val_desecho_priv_arr

    df_flujo_privado = pd.DataFrame({
        "(+) Ingresos Operacionales Netos (Sin IVA)": ing_priv_0_10,
        "(-) Costos Operativos Totales (Fijos, Variables y Personal)": costos_priv_0_10,
        "(-) Depreciación Legal del Ejercicio": dep_0_10,
        "(=) Utilidad Antes de Impuestos (UAI)": uai,
        f"(-) Impuesto a la Renta ({int(imp_renta * 100)}%)": impuestos,
        "(=) Utilidad Después de Impuestos (UDI)": udi,
        "(+) Ajuste por Depreciación (Gasto No Desembolsable)": reverso_dep,
        "(-) Inversiones y Reinversiones en Activos Fijos": inv_fija_priv,
        "(+/-) Inversión y Recuperación de Capital de Trabajo": flujo_ct_priv,
        "(+) Valor de Desecho de los Activos": val_desecho_priv_arr,
        "(=) FLUJO DE CAJA PRIVADO NETO": flujo_caja_privado
    }, index=cols_tot).T.astype(int)
    df_flujo_privado.index.name = "Estructura Flujo de Caja Privado"

    van_privado = float(npf.npv(tasa_priv, flujo_caja_privado))
    tir_privada = float(npf.irr(flujo_caja_privado))
    payback_privado = calcular_payback(flujo_caja_privado)

    ben_soc_ventas = np.insert(df_ingresos.loc["TOTAL INGRESOS PRIVADOS NETOS (SIN IVA)"].values, 0, 0)
    ben_soc_muni = np.insert(beneficio_ahorro_muni, 0, 0)
    ben_soc_otros = np.insert(beneficio_otros_soc, 0, 0)
    costos_soc_0_10 = np.insert(-costos_tot_soc_anual, 0, 0)
    inv_fija_soc = -df_inv_social.loc["TOTAL INVERSIONES Y REINVERSIONES SOCIALES"].values

    val_desecho_soc_arr = np.zeros(horizonte + 1)
    val_desecho_soc_arr[-1] = round(valor_desecho_soc_total)

    flujo_caja_social = (
        ben_soc_ventas + ben_soc_muni + ben_soc_otros
        + costos_soc_0_10 + inv_fija_soc + flujo_ct_soc + val_desecho_soc_arr
    )

    df_flujo_social = pd.DataFrame({
        "(+) Beneficio Social por Disponibilidad de Bienes (Ventas s/IVA)": ben_soc_ventas,
        "(+) Liberación de Recursos Municipales (Ahorro Vertedero Curaco)": ben_soc_muni,
        "(+) Otros Beneficios Sociales y Ambientales": ben_soc_otros,
        "(-) Costos Operativos y Mano de Obra a Precios Sociales": costos_soc_0_10,
        "(-) Inversiones y Reinversiones a Precios Sociales": inv_fija_soc,
        "(+/-) Capital de Trabajo a Precios Sociales": flujo_ct_soc,
        "(+) Valor Residual Social de los Activos": val_desecho_soc_arr,
        "(=) FLUJO DE BENEFICIOS SOCIALES NETOS": flujo_caja_social
    }, index=cols_tot).T.astype(int)
    df_flujo_social.index.name = "Estructura Flujo de Caja Social (MIDESO)"

    van_social = float(npf.npv(tasa_soc, flujo_caja_social))
    tir_social = float(npf.irr(flujo_caja_social))
    payback_social = calcular_payback(flujo_caja_social)

    df_indicadores = pd.DataFrame([
        {
            "Tipo de Evaluación": "Evaluación Privada (Punto 7)",
            "Tasa de Descuento": f"{tasa_priv * 100:.1f}%",
            "VAN / VANS ($)": fmt_clp(van_privado),
            "TIR / TIRS (%)": f"{tir_privada * 100:.2f}%",
            "Payback (Años)": payback_privado if payback_privado is not None else "No recupera",
            "Decisión": "VIABLE (Rentable)" if van_privado > 0 else "NO VIABLE"
        },
        {
            "Tipo de Evaluación": "Evaluación Social SNI (Punto 8)",
            "Tasa de Descuento": f"{tasa_soc * 100:.1f}%",
            "VAN / VANS ($)": fmt_clp(van_social),
            "TIR / TIRS (%)": f"{tir_social * 100:.2f}%",
            "Payback (Años)": payback_social if payback_social is not None else "No recupera",
            "Decisión": "VIABLE SOCIALMENTE" if van_social > 0 else "NO VIABLE"
        }
    ])

    df_informantes, df_supuestos, df_catastro_6, df_demanda_3, df_cotizaciones, df_porque = armar_tablas_informantes_y_supuestos(cfg)

    return {
        "cfg": cfg,
        "df_personal": df_personal,
        "df_cantidades": df_cantidades,
        "df_ingresos": df_ingresos,
        "df_costos_priv": df_costos_priv,
        "df_costos_soc": df_costos_soc,
        "df_inversiones": df_inversiones,
        "df_depreciacion": df_depreciacion,
        "df_capital_trabajo": df_capital_trabajo,
        "df_flujo_privado": df_flujo_privado,
        "df_flujo_social": df_flujo_social,
        "df_indicadores": df_indicadores,
        "df_informantes": df_informantes,
        "df_supuestos": df_supuestos,
        "df_catastro_6": df_catastro_6,
        "df_demanda_3": df_demanda_3,
        "df_cotizaciones": df_cotizaciones,
        "df_porque": df_porque,
        "metricas": {
            "van_privado": round(van_privado),
            "tir_privada": round(tir_privada * 100, 2) if not np.isnan(tir_privada) else 0.0,
            "payback_privado": payback_privado,
            "van_social": round(van_social),
            "tir_social": round(tir_social * 100, 2) if not np.isnan(tir_social) else 0.0,
            "payback_social": payback_social,
            "ct_inicial": round(abs(flujo_ct_priv[0])),
            "inv_inicial_priv": round(df_inversiones.loc["TOTAL INVERSIONES Y REINVERSIONES NETAS", "Año 0"]),
            "ingresos_anio_1": round(df_ingresos.loc["TOTAL INGRESOS PRIVADOS NETOS (SIN IVA)", "Año 1"]),
            "costos_anio_1": round(df_costos_priv.loc["TOTAL COSTOS OPERATIVOS PRIVADOS (SIN IVA)", "Año 1"])
        }
    }


# ==============================================================================
# BLOQUE 3: EXPORTADOR DE LOS 3 ARCHIVOS EXCEL EXIGIDOS EN LA PAUTA
# ==============================================================================

def exportar_tres_excels_pauta(resultados, carpeta_destino):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    os.makedirs(carpeta_destino, exist_ok=True)
    ruta_excel_1 = os.path.join(carpeta_destino, "1_Excel_Flujos_Editables.xlsx")
    ruta_excel_2 = os.path.join(carpeta_destino, "2_Excel_Informantes_Encuestas_Osorno.xlsx")
    ruta_excel_3 = os.path.join(carpeta_destino, "3_Excel_Calculos_Proyecciones.xlsx")

    for r_file in [ruta_excel_1, ruta_excel_2, ruta_excel_3]:
        if os.path.exists(r_file):
            try:
                os.remove(r_file)
            except PermissionError:
                raise PermissionError(
                    f"El archivo '{os.path.basename(r_file)}' está abierto en Excel.\n"
                    f"Ciérralo primero para poder actualizarlo."
                )

    cfg = resultados["cfg"]
    horizonte = int(cfg["HORIZONTE_ANIOS"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = int(cfg["DIAS_DESFASE_CT"])
    crec = float(cfg["TASA_CRECIMIENTO"])
    titulo_proy = cfg.get("NOMBRE_PROYECTO", "Cartón Vivo Osorno").upper()

    font_titulo = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Arial", size=9.5, italic=True, color="E2E8F0")
    font_seccion = Font(name="Arial", size=10.5, bold=True, color="000000")
    font_header = Font(name="Arial", size=9.5, bold=True, color="FFFFFF")
    font_regular = Font(name="Arial", size=9.5, bold=False, color="000000")
    font_concepto = Font(name="Arial", size=9.5, bold=True, color="000000")
    font_total = Font(name="Arial", size=9.5, bold=True, color="000000")

    fill_banner = PatternFill("solid", fgColor="0F172A")
    fill_header = PatternFill("solid", fgColor="1E293B")
    fill_par = PatternFill("solid", fgColor="FFFFFF")
    fill_impar = PatternFill("solid", fgColor="F8FAFC")
    fill_total = PatternFill("solid", fgColor="DBEAFE")
    fill_flujo = PatternFill("solid", fgColor="DCFCE7")

    borde_fino = Side(style="thin", color="CBD5E1")
    borde_doble = Side(style="double", color="000000")
    borde_sup_total = Side(style="thin", color="000000")

    border_celda = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    border_total = Border(left=borde_fino, right=borde_fino, top=borde_sup_total, bottom=borde_doble)

    def crear_banner_hoja(ws, titulo_hoja, subtitulo_pauta, total_cols, color_tab):
        ws.views.sheetView[0].showGridLines = True
        ws.sheet_properties.tabColor = color_tab
        max_c = max(total_cols, 6)

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_c)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_c)

        c1 = ws.cell(1, 1, f"  PROYECTO: {titulo_proy} — {titulo_hoja}")
        c2 = ws.cell(2, 1, f"  {subtitulo_pauta}")

        for col in range(1, max_c + 1):
            ws.cell(1, col).fill = fill_banner
            ws.cell(2, col).fill = fill_banner

        c1.font = font_titulo
        c1.alignment = Alignment(vertical="center", horizontal="left")
        c2.font = font_subtitulo
        c2.alignment = Alignment(vertical="center", horizontal="left")

        ws.row_dimensions[1].height = 26
        ws.row_dimensions[2].height = 19
        ws.row_dimensions[3].height = 10

    def estilizar_tabla_excel(ws, fila_header, num_filas, num_cols, es_moneda=True, alto_fila=24):
        ws.row_dimensions[fila_header].height = 28
        headers_texto = []

        for col in range(1, num_cols + 1):
            celda_h = ws.cell(fila_header, col)
            headers_texto.append(str(celda_h.value or ""))
            celda_h.fill = fill_header
            celda_h.font = font_header
            celda_h.border = border_celda
            celda_h.alignment = Alignment(
                horizontal="left" if col == 1 else "center",
                vertical="center",
                wrap_text=True
            )

        cols_texto_largo = [
            "Tareas Principales",
            "Lugares e Instituciones en Osorno",
            "Método Aplicado (Entrevista, Encuesta, Aforo o Cotización)",
            "Dato Clave Obtenido para el Proyecto",
            "Explicación Sencilla de Cómo Funciona",
            "Respaldo Técnico / Legal",
            "Nombre del Comercio Real (Osorno)",
            "Dirección en Calle Eleuterio Ramírez",
            "Tipo de Cartón y Forma de Retiro (Antes del Camión de 06:00-07:00 AM)",
            "Cliente / Institución Compradora en Osorno",
            "Con Quién se Validó el Dato (Entrevista / Consulta)",
            "Cómo Compra y Por Qué Nos Elige",
            "Descripción del Ítem Cotizado",
            "Fuente / Referencia de Precio o Costo en Osorno",
            "Explicación Fácil de Entender (Para tu Defensa y Entrega)"
        ]

        for idx_f in range(num_filas):
            r = fila_header + 1 + idx_f
            ws.row_dimensions[r].height = alto_fila
            texto_col1 = str(ws.cell(r, 1).value or "").upper()
            es_flujo_final = "FLUJO DE CAJA" in texto_col1 or "FLUJO DE BENEFICIOS" in texto_col1
            es_subtotal = any(k in texto_col1 for k in ["TOTAL", "(=)", "CAPITAL DE TRABAJO REQUERIDO"])

            for col in range(1, num_cols + 1):
                celda = ws.cell(r, col)
                val = celda.value
                col_nombre = headers_texto[col - 1]
                es_formula = isinstance(val, str) and val.startswith("=")

                if es_flujo_final:
                    celda.fill = fill_flujo
                    celda.font = font_total
                    celda.border = border_total
                elif es_subtotal:
                    celda.fill = fill_total
                    celda.font = font_total
                    celda.border = border_total
                else:
                    celda.fill = fill_par if idx_f % 2 == 0 else fill_impar
                    celda.font = font_concepto if col == 1 else font_regular
                    celda.border = border_celda

                if col == 1:
                    celda.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif isinstance(val, (int, float, np.integer, np.floating)) or es_formula:
                    celda.alignment = Alignment(horizontal="right", vertical="center")
                    if "Tasa" in col_nombre or "TIR" in col_nombre:
                        celda.number_format = '0.00%'
                    elif "Decisión" in col_nombre:
                        celda.alignment = Alignment(horizontal="center", vertical="center")
                    elif "$" in col_nombre or (es_moneda and not any(k in col_nombre for k in ["Cant.", "Factor", "Payback", "Kg", "Ton", "Unidades"])):
                        celda.number_format = '$#,##0;-$#,##0;"$0"'
                    elif "Ton" in col_nombre or "Día" in col_nombre or (isinstance(val, float) and not val.is_integer()):
                        celda.number_format = '#,##0.00'
                    else:
                        celda.number_format = '#,##0'
                else:
                    es_izq = col_nombre in cols_texto_largo
                    celda.alignment = Alignment(
                        horizontal="left" if es_izq else "center",
                        vertical="center",
                        wrap_text=True
                    )

    def auto_ajustar_columnas(ws, ancho_min_col1=46, ancho_min_datos=16):
        for col in ws.columns:
            col_idx = col[0].column
            col_letra = get_column_letter(col_idx)
            max_len = 0
            for celda in col:
                if celda.row in (1, 2):
                    continue
                texto = str(celda.value or "")
                if texto.startswith("="):
                    texto = "$123.456.789"
                if len(texto) > max_len:
                    max_len = len(texto)
            if col_idx == 1:
                ws.column_dimensions[col_letra].width = max(min(max_len + 5, 62), ancho_min_col1)
            else:
                ws.column_dimensions[col_letra].width = max(min(max_len + 4, 58), ancho_min_datos)

    ult_let = get_column_letter(horizonte + 2)

    # EXCEL 1: 1_Excel_Flujos_Editables.xlsx
    with pd.ExcelWriter(ruta_excel_1, engine="openpyxl") as writer1:
        h1 = "1. Resumen Rentabilidad"
        df1 = resultados["df_indicadores"].copy()
        df1.to_excel(writer1, sheet_name=h1, startrow=3, index=False)
        ws1 = writer1.sheets[h1]

        h2 = "2. Flujo Caja Privado"
        df2 = resultados["df_flujo_privado"].copy()
        df2.to_excel(writer1, sheet_name=h2, startrow=3, index=True)
        ws2 = writer1.sheets[h2]

        h3 = "3. Flujo Social MIDESO"
        df3 = resultados["df_flujo_social"].copy()
        df3.to_excel(writer1, sheet_name=h3, startrow=3, index=True)
        ws3 = writer1.sheets[h3]

        h4 = "4. Capital de Trabajo"
        df4_ct = resultados["df_capital_trabajo"].copy()
        df4_ct.to_excel(writer1, sheet_name=h4, startrow=3, index=True)
        ws4_ct = writer1.sheets[h4]

        for col_idx in range(3, horizonte + 3):
            col_let = get_column_letter(col_idx)
            ws4_ct[f"{col_let}6"] = f"=ROUND({col_let}5*({dias_ct}/365), 0)"
        ws4_ct["B7"] = "=-C6"
        for col_idx in range(3, horizonte + 2):
            col_let = get_column_letter(col_idx)
            next_let = get_column_letter(col_idx + 1)
            ws4_ct[f"{col_let}7"] = f"=-({next_let}6-{col_let}6)"
        ws4_ct[f"{ult_let}7"] = f"={ult_let}6"

        ws2["B13"] = f"='{h4}'!B7"
        ws2["B15"] = "=SUM(B10:B14)"
        for col_idx in range(3, horizonte + 3):
            col_let = get_column_letter(col_idx)
            ws2[f"{col_let}8"] = f"=SUM({col_let}5:{col_let}7)"
            ws2[f"{col_let}9"] = f"=IF({col_let}8>0, -ROUND({col_let}8*{imp_renta}, 0), 0)"
            ws2[f"{col_let}10"] = f"={col_let}8+{col_let}9"
            ws2[f"{col_let}11"] = f"=-{col_let}7"
            ws2[f"{col_let}13"] = f"='{h4}'!{col_let}7"
            ws2[f"{col_let}15"] = f"=SUM({col_let}10:{col_let}14)"

        for col_idx in range(2, horizonte + 3):
            col_let = get_column_letter(col_idx)
            ws3[f"{col_let}12"] = f"=SUM({col_let}5:{col_let}11)"

        ws1["B5"] = tasa_priv
        ws1["C5"] = f"=ROUND(NPV(B5, '{h2}'!C15:{ult_let}15)+'{h2}'!B15, 0)"
        ws1["D5"] = f"=IRR('{h2}'!B15:{ult_let}15)"
        ws1["F5"] = '=IF(C5>0, "VIABLE (Rentable)", "NO VIABLE")'

        ws1["B6"] = tasa_soc
        ws1["C6"] = f"=ROUND(NPV(B6, '{h3}'!C12:{ult_let}12)+'{h3}'!B12, 0)"
        ws1["D6"] = f"=IRR('{h3}'!B12:{ult_let}12)"
        ws1["F6"] = '=IF(C6>0, "VIABLE SOCIALMENTE", "NO VIABLE")'

        crear_banner_hoja(ws1, "ANEXO EXCEL 1: RESUMEN DE INDICADORES CON FÓRMULAS EDITABLES", "Síntesis comparativa Evaluación Privada (Punto 7) vs. Evaluación Social MIDESO/SNI (Punto 8)", df1.shape[1], "15803D")
        estilizar_tabla_excel(ws1, fila_header=4, num_filas=len(df1), num_cols=df1.shape[1], es_moneda=True)
        auto_ajustar_columnas(ws1, ancho_min_col1=34, ancho_min_datos=20)

        crear_banner_hoja(ws2, "FLUJO DE CAJA PRIVADO PROYECTADO (100% EDITABLE CON FÓRMULAS VIVAS)", "Punto 7 de la Pauta ULagos — Valores en Pesos Chilenos ($ CLP) Netos de IVA", df2.shape[1] + 1, "0284C7")
        estilizar_tabla_excel(ws2, fila_header=4, num_filas=len(df2), num_cols=df2.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws2, ancho_min_col1=52, ancho_min_datos=16)

        crear_banner_hoja(ws3, "FLUJO DE BENEFICIOS SOCIALES NETOS (METODOLOGÍA MIDESO / SNI)", "Punto 8 de la Pauta ULagos — Precios Sombra y Ahorro Municipal Vertedero de Curaco", df3.shape[1] + 1, "16A34A")
        estilizar_tabla_excel(ws3, fila_header=4, num_filas=len(df3), num_cols=df3.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws3, ancho_min_col1=56, ancho_min_datos=16)

        crear_banner_hoja(ws4_ct, "CÁLCULO EDITABLE DEL CAPITAL DE TRABAJO (45 DÍAS DE DESFASE)", "Punto 6 de la Pauta ULagos — Inversión Inicial en Año 0 y Recuperación en Año 10", df4_ct.shape[1] + 1, "0D9488")
        estilizar_tabla_excel(ws4_ct, fila_header=4, num_filas=len(df4_ct), num_cols=df4_ct.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws4_ct, ancho_min_col1=50, ancho_min_datos=16)

    # EXCEL 2: 2_Excel_Informantes_Encuestas_Osorno.xlsx
    df_pq = resultados["df_porque"].copy()
    df_inf = resultados["df_informantes"].copy()
    df_cot = resultados["df_cotizaciones"].copy()

    df_cat6 = resultados["df_catastro_6"].copy()
    fila_tot_cat = pd.DataFrame([[
        "TOTAL EJE RAMÍREZ (n = 6)",
        "6 Locales Comerciales Eje Eleuterio Ramírez",
        "Tramo Eleuterio Ramírez #446 al #1024, Osorno",
        "Total Tramo",
        "Total cartón limpio retirado antes del camión municipal (06:00-07:00 AM a Vertedero de Curaco)",
        round(df_cat6["Kg por Día"].sum(), 2),
        int(df_cat6["Kg al Mes (25 días)"].sum()),
        round(df_cat6["Toneladas al Año"].sum(), 2)
    ]], columns=df_cat6.columns)
    df_cat6_full = pd.concat([df_cat6, fila_tot_cat], ignore_index=True)

    df_dem3 = resultados["df_demanda_3"].copy()
    tot_unid_d = int(cfg["P1_UNIDADES"]) + int(cfg["P2_UNIDADES"])
    pond_precio = int(round((int(cfg["P1_UNIDADES"]) * float(cfg["P1_PRECIO_BRUTO"]) + int(cfg["P2_UNIDADES"]) * float(cfg["P2_PRECIO_BRUTO"])) / tot_unid_d))
    fila_tot_dem = pd.DataFrame([[
        "TOTAL COMPRADORES (n = 3)",
        "3 Grupos de Clientes en Osorno (DAEM, Jardines Privados y Hogar Santa María)",
        "Total Comunal",
        f"Set Escolar ({int(cfg['P1_UNIDADES']):,}) + Juego Adulto Mayor ({int(cfg['P2_UNIDADES']):,})",
        tot_unid_d,
        pond_precio,
        "Total de juegos vendidos en el Año 1 en la comuna de Osorno."
    ]], columns=df_dem3.columns)
    df_dem3_full = pd.concat([df_dem3, fila_tot_dem], ignore_index=True)

    with pd.ExcelWriter(ruta_excel_2, engine="openpyxl") as writer2:
        s1 = "1. Catastro 6 Locales Ramirez"
        df_cat6_full.to_excel(writer2, sheet_name=s1, startrow=3, index=False)
        ws_c1 = writer2.sheets[s1]
        crear_banner_hoja(ws_c1, "ANEXO EXCEL 2: ENCUESTA Y AFORO EN 6 LOCALES DE CALLE ELEUTERIO RAMÍREZ (#446 AL #1024)", "Retiro temprano antes de que pase el camión municipal (06:00-07:00 AM hacia el Vertedero de Curaco)", df_cat6_full.shape[1], "0284C7")
        estilizar_tabla_excel(ws_c1, fila_header=4, num_filas=len(df_cat6_full), num_cols=df_cat6_full.shape[1], es_moneda=False, alto_fila=32)
        auto_ajustar_columnas(ws_c1, ancho_min_col1=24, ancho_min_datos=20)

        s2 = "2. Entrevistas 3 Compradores"
        df_dem3_full.to_excel(writer2, sheet_name=s2, startrow=3, index=False)
        ws_c2 = writer2.sheets[s2]
        crear_banner_hoja(ws_c2, "ENTREVISTAS Y VALIDACIÓN EN 3 COMPRADORES INSTITUCIONALES DE OSORNO", "DAEM Osorno (Escuelas y VTF), Jardines Privados/Guarderías y Hogar de Ancianos Santa María", df_dem3_full.shape[1], "16A34A")
        estilizar_tabla_excel(ws_c2, fila_header=4, num_filas=len(df_dem3_full), num_cols=df_dem3_full.shape[1], es_moneda=True, alto_fila=34)
        auto_ajustar_columnas(ws_c2, ancho_min_col1=24, ancho_min_datos=22)

        s3 = "3. Matriz Informantes Terreno"
        df_inf.to_excel(writer2, sheet_name=s3, startrow=3, index=False)
        ws_c3 = writer2.sheets[s3]
        crear_banner_hoja(ws_c3, "REGISTRO DE INFORMANTES, ENTREVISTAS Y ENCUESTAS EN OSORNO", "Detalle de fuentes primarias y secundarias utilizadas para obtener la información del proyecto", df_inf.shape[1], "7C3AED")
        estilizar_tabla_excel(ws_c3, fila_header=4, num_filas=len(df_inf), num_cols=df_inf.shape[1], es_moneda=False, alto_fila=36)
        auto_ajustar_columnas(ws_c3, ancho_min_col1=36, ancho_min_datos=42)

        s4 = "4. Cotizaciones Precios Costos"
        df_cot.to_excel(writer2, sheet_name=s4, startrow=3, index=False)
        ws_c4 = writer2.sheets[s4]
        crear_banner_hoja(ws_c4, "COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS EN OSORNO (REFERENCIAS NOTAS AL PIE)", "Valores reales con IVA y sin IVA verificados en Octubre-Noviembre 2026", df_cot.shape[1], "D97706")
        estilizar_tabla_excel(ws_c4, fila_header=4, num_filas=len(df_cot), num_cols=df_cot.shape[1], es_moneda=True, alto_fila=28)
        auto_ajustar_columnas(ws_c4, ancho_min_col1=18, ancho_min_datos=22)

        s5 = "5. Porque de Cada Cifra"
        df_pq.to_excel(writer2, sheet_name=s5, startrow=3, index=False)
        ws_c5 = writer2.sheets[s5]
        crear_banner_hoja(ws_c5, "GUÍA FÁCIL DE ENTENDER: EL PORQUÉ DE CADA CIFRA DEL PROYECTO", "Explicación clara y directa de todos los números del estudio", df_pq.shape[1], "4F46E5")
        estilizar_tabla_excel(ws_c5, fila_header=4, num_filas=len(df_pq), num_cols=df_pq.shape[1], es_moneda=False, alto_fila=38)
        auto_ajustar_columnas(ws_c5, ancho_min_col1=36, ancho_min_datos=45)

    # EXCEL 3: 3_Excel_Calculos_Proyecciones.xlsx
    with pd.ExcelWriter(ruta_excel_3, engine="openpyxl") as writer3:
        p1 = "1. Proyeccion Fisica y Merma"
        df_p1 = resultados["df_cantidades"].copy()
        df_p1.to_excel(writer3, sheet_name=p1, startrow=3, index=True)
        ws_p1 = writer3.sheets[p1]

        p2 = "2. Proyeccion Ingresos Ventas"
        df_p2 = resultados["df_ingresos"].copy()
        df_p2.to_excel(writer3, sheet_name=p2, startrow=3, index=True)
        ws_p2 = writer3.sheets[p2]

        p3 = "3. Proyeccion Costos Operativos"
        df_p3 = resultados["df_costos_priv"].copy()
        df_p3.to_excel(writer3, sheet_name=p3, startrow=3, index=True)
        ws_p3 = writer3.sheets[p3]

        p4 = "4. Inversiones y Depreciacion"
        df_p4_inv = resultados["df_inversiones"].copy()
        df_p4_dep = resultados["df_depreciacion"].copy()
        df_p4_inv.to_excel(writer3, sheet_name=p4, startrow=3, index=True)
        fila_ini_dep = 3 + len(df_p4_inv) + 4
        df_p4_dep.to_excel(writer3, sheet_name=p4, startrow=fila_ini_dep, index=True)
        ws_p4 = writer3.sheets[p4]

        p5 = "5. Sueldos y Supuestos"
        df_p5_per = resultados["df_personal"].copy()
        df_p5_sup = resultados["df_supuestos"].copy()
        df_p5_per.to_excel(writer3, sheet_name=p5, startrow=3, index=False)
        fila_ini_sup3 = 3 + len(df_p5_per) + 4
        df_p5_sup.to_excel(writer3, sheet_name=p5, startrow=fila_ini_sup3, index=False)
        ws_p5 = writer3.sheets[p5]

        for col_idx in range(2, horizonte + 2):
            col_let = get_column_letter(col_idx)
            if col_idx > 2:
                prev_let = get_column_letter(col_idx - 1)
                ws_p1[f"{col_let}5"] = f"=ROUND({prev_let}5*(1+{crec}), 2)"
                ws_p1[f"{col_let}6"] = f"=ROUND({prev_let}6*(1+{crec}), 0)"
                ws_p1[f"{col_let}7"] = f"=ROUND({prev_let}7*(1+{crec}), 0)"
            ws_p1[f"{col_let}8"] = f"=SUM({col_let}6:{col_let}7)"
            ws_p1[f"{col_let}9"] = f"=ROUND(({col_let}6*{float(cfg['P1_KG_CARTON'])}+{col_let}7*{float(cfg['P2_KG_CARTON'])})/1000, 2)"

        fila_tot_ing3 = 4 + len(df_p2)
        fila_tot_cos3 = 4 + len(df_p3)
        for col_idx in range(2, horizonte + 2):
            col_let = get_column_letter(col_idx)
            ws_p2[f"{col_let}{fila_tot_ing3}"] = f"=SUM({col_let}5:{col_let}{fila_tot_ing3 - 1})"
            ws_p3[f"{col_let}{fila_tot_cos3}"] = f"=SUM({col_let}5:{col_let}{fila_tot_cos3 - 1})"

        fila_tot_inv3 = 4 + len(df_p4_inv)
        for col_idx in range(2, horizonte + 3):
            col_let = get_column_letter(col_idx)
            ws_p4[f"{col_let}{fila_tot_inv3}"] = f"=SUM({col_let}5:{col_let}{fila_tot_inv3 - 1})"

        fila_hdr_dep3 = fila_ini_dep + 1
        fila_tot_dep3 = fila_hdr_dep3 + len(df_p4_dep)
        for col_idx in range(2, horizonte + 2):
            col_let = get_column_letter(col_idx)
            ws_p4[f"{col_let}{fila_tot_dep3}"] = f"=SUM({col_let}{fila_hdr_dep3 + 1}:{col_let}{fila_tot_dep3 - 1})"

        for idx_p in range(len(df_p5_per)):
            r_p = 5 + idx_p
            ws_p5[f"F{r_p}"] = f"=B{r_p}*E{r_p}*12"
            ws_p5[f"H{r_p}"] = f"=ROUND(F{r_p}*G{r_p}, 0)"

        crear_banner_hoja(ws_p1, "ANEXO EXCEL 3: PROYECCIÓN FÍSICA DE RECOLECCIÓN Y PRODUCCIÓN (AÑOS 1 AL 10)", "Punto 5.a — Crecimiento 5% Anual | Balance de Masa: 15,00 ton recolectadas -> 10,50 ton útiles (70%) y 4,50 ton merma (30%)", df_p1.shape[1] + 1, "D97706")
        estilizar_tabla_excel(ws_p1, fila_header=4, num_filas=len(df_p1), num_cols=df_p1.shape[1] + 1, es_moneda=False)
        auto_ajustar_columnas(ws_p1, ancho_min_col1=48, ancho_min_datos=14)

        crear_banner_hoja(ws_p2, "PROYECCIÓN DE INGRESOS OPERACIONALES POR VENTAS (SIN IVA — AÑOS 1 AL 10)", "Punto 5.d — Set Didáctico Escolar ($16.000 c/IVA) y Juegos Cognitivos Adulto Mayor ($19.000 c/IVA)", df_p2.shape[1] + 1, "16A34A")
        estilizar_tabla_excel(ws_p2, fila_header=4, num_filas=len(df_p2), num_cols=df_p2.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws_p2, ancho_min_col1=52, ancho_min_datos=16)

        crear_banner_hoja(ws_p3, "PROYECCIÓN DE COSTOS OPERATIVOS VARIABLES, FIJOS Y REMUNERACIONES (SIN IVA)", "Punto 5.c — Insumos por unidad, flete por tonelada, arriendo en Ovejería, suministros y personal", df_p3.shape[1] + 1, "EA580C")
        estilizar_tabla_excel(ws_p3, fila_header=4, num_filas=len(df_p3), num_cols=df_p3.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws_p3, ancho_min_col1=52, ancho_min_datos=16)

        crear_banner_hoja(ws_p4, "CALENDARIO DE INVERSIONES, REINVERSIONES Y DEPRECIACIÓN LINEAL", "Punto 5.b — Inversión Inicial en Año 0 ($16.000.000 c/IVA = $13.445.378 Netos) y Renovación en Año 5", df_p4_inv.shape[1] + 1, "7C3AED")
        estilizar_tabla_excel(ws_p4, fila_header=4, num_filas=len(df_p4_inv), num_cols=df_p4_inv.shape[1] + 1, es_moneda=True)
        lbl_d3 = ws_p4.cell(fila_ini_dep, 1, "TABLA DE DEPRECIACIÓN LINEAL ANUAL DE ACTIVOS FIJOS (VALORES NETOS SIN IVA)")
        lbl_d3.font = font_seccion
        estilizar_tabla_excel(ws_p4, fila_header=fila_ini_dep + 1, num_filas=len(df_p4_dep), num_cols=df_p4_dep.shape[1] + 1, es_moneda=True)
        auto_ajustar_columnas(ws_p4, ancho_min_col1=48, ancho_min_datos=16)

        crear_banner_hoja(ws_p5, "CÁLCULO DE REMUNERACIONES, PRECIOS SOMBRA MIDESO Y SUPUESTOS DEL PROYECTO", "Punto 4.3 — Dotación de 5 trabajadores y factores sociales de conversión (MOC=1,00; MOSC=0,85; MONC=0,68)", 8, "0D9488")
        estilizar_tabla_excel(ws_p5, fila_header=4, num_filas=len(df_p5_per), num_cols=df_p5_per.shape[1], es_moneda=True)
        lbl_s3 = ws_p5.cell(fila_ini_sup3, 1, "MATRIZ DE SUPUESTOS Y DEFINICIONES TÉCNICAS DEL PROYECTO")
        lbl_s3.font = font_seccion
        estilizar_tabla_excel(ws_p5, fila_header=fila_ini_sup3 + 1, num_filas=len(df_p5_sup), num_cols=4, es_moneda=False, alto_fila=36)
        auto_ajustar_columnas(ws_p5, ancho_min_col1=38, ancho_min_datos=24)

    return ruta_excel_1, ruta_excel_2, ruta_excel_3


# ==============================================================================
# BLOQUE 4: GRÁFICOS PARA EL INFORME WORD Y LA VENTANA
# ==============================================================================

def generar_graficos_informe(resultados, carpeta_destino=None):
    if carpeta_destino is None:
        carpeta_destino = os.getcwd()
    ruta_img_flujos = os.path.join(carpeta_destino, "grafico_flujos_cartonvivo.png")
    ruta_img_carton = os.path.join(carpeta_destino, "grafico_carton_ramirez.png")

    df_fp = resultados["df_flujo_privado"]
    df_fs = resultados["df_flujo_social"]
    df_cant = resultados["df_cantidades"]

    anios_cols = list(df_fp.columns)
    flujo_p = df_fp.loc["(=) FLUJO DE CAJA PRIVADO NETO"].values / 1e6
    flujo_s = df_fs.loc["(=) FLUJO DE BENEFICIOS SOCIALES NETOS"].values / 1e6

    fig1, ax1 = plt.subplots(figsize=(8.8, 4.0), dpi=220)
    x = np.arange(len(anios_cols))
    ancho = 0.36

    b1 = ax1.bar(x - ancho / 2, flujo_p, ancho, label="Flujo de Caja Privado Neto (MM$)", color="#0284C7", edgecolor="#000000", linewidth=0.6)
    b2 = ax1.bar(x + ancho / 2, flujo_s, ancho, label="Flujo Social MIDESO (MM$)", color="#16A34A", edgecolor="#000000", linewidth=0.6)

    ax1.axhline(0, color="#000000", linewidth=1.0)
    ax1.set_title("Comparación Anual: Flujo de Caja Privado vs. Flujo Social (Año 0 al 10)", fontsize=11, fontweight="bold", color="#000000", pad=12)
    ax1.set_xticks(x)
    ax1.set_xticklabels(anios_cols, fontsize=8.5)
    ax1.set_ylabel("Millones de Pesos ($ CLP)", fontsize=9, fontweight="bold", color="#000000")
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda val, pos: f"${val:,.1f}M"))
    ax1.grid(axis="y", linestyle="--", alpha=0.45)
    ax1.legend(frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", fontsize=8.5, loc="upper left")

    for bar in b1:
        h = bar.get_height()
        offset = 0.45 if h >= 0 else -1.35
        ax1.annotate(f"{h:.1f}", xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, offset * 4),
                     textcoords="offset points", ha="center", va="bottom" if h >= 0 else "top", fontsize=6.8, color="#000000", fontweight="bold")
    for bar in b2:
        h = bar.get_height()
        offset = 0.45 if h >= 0 else -1.35
        ax1.annotate(f"{h:.1f}", xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, offset * 4),
                     textcoords="offset points", ha="center", va="bottom" if h >= 0 else "top", fontsize=6.8, color="#000000", fontweight="bold")

    fig1.tight_layout()
    fig1.savefig(ruta_img_flujos, dpi=220)
    plt.close(fig1)

    fig2, ax2 = plt.subplots(figsize=(8.2, 3.6), dpi=220)
    cols_op = list(df_cant.columns)
    ton_rec = df_cant.loc["Residuo de Cartón Recolectado (Toneladas)"].values
    ton_inc = df_cant.loc["Cartón Efectivamente Incorporado (Toneladas)"].values

    x2 = np.arange(len(cols_op))
    ax2.plot(x2, ton_rec, marker="o", linewidth=2.3, color="#D97706", label="Cartón Recolectado en 6 Locales Calle Ramírez (Ton/año)")
    ax2.plot(x2, ton_inc, marker="s", linewidth=2.3, color="#0D9488", label="Cartón Usado en los Productos Terminados (70%)")
    ax2.fill_between(x2, ton_inc, ton_rec, color="#FEF3C7", alpha=0.55, label="Recortes Sobrantes y Merma de Selección (30%)")

    ax2.set_title("Aprovechamiento del Cartón: Total Recolectado vs. Usado en los Juegos", fontsize=10.5, fontweight="bold", color="#000000", pad=10)
    ax2.set_xticks(x2)
    ax2.set_xticklabels(cols_op, fontsize=8.5)
    ax2.set_ylabel("Toneladas al Año (Ton)", fontsize=9, fontweight="bold", color="#000000")
    ax2.set_ylim(0, max(ton_rec) * 1.25)
    ax2.grid(True, linestyle="--", alpha=0.45)
    ax2.legend(frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E1", fontsize=8, loc="upper left")

    for i, (vr, vi) in enumerate(zip(ton_rec, ton_inc)):
        ax2.annotate(f"{vr:.1f}t", (x2[i], vr), textcoords="offset points", xytext=(0, 6), ha="center", fontsize=7.5, fontweight="bold", color="#000000")
        ax2.annotate(f"{vi:.1f}t", (x2[i], vi), textcoords="offset points", xytext=(0, -12), ha="center", fontsize=7.5, fontweight="bold", color="#000000")

    fig2.tight_layout()
    fig2.savefig(ruta_img_carton, dpi=220)
    plt.close(fig2)

    return ruta_img_flujos, ruta_img_carton


# ==============================================================================
# BLOQUE 5: INFORME WORD CON PLANTILLA ULAGOS AUTOMÁTICA, ÍNDICE EN PÁGINA 2,
# NOTAS A PIE DE PÁGINA Y SIN DEPENDENCIA DE ESTILOS EN INGLÉS ('Heading 1')
# ==============================================================================

def configurar_borde_celda(celda, **kwargs):
    tcPr = celda._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge, val_dict in kwargs.items():
        edge_el = tcBorders.find(qn(f"w:{edge}"))
        if edge_el is None:
            edge_el = OxmlElement(f"w:{edge}")
            tcBorders.append(edge_el)
        for k, v in val_dict.items():
            edge_el.set(qn(f"w:{k}"), str(v))


def sombrear_celda(celda, color_hex):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), color_hex)
    celda._tc.get_or_add_tcPr().append(shading)


def forzar_fuente_arial_negro(estilo, size_pt, bold=False):
    estilo.font.name = "Arial"
    estilo.font.size = Pt(size_pt)
    estilo.font.bold = bold
    estilo.font.color.rgb = RGBColor(0, 0, 0)
    rPr = estilo.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")


def agregar_marcador_word(parrafo, nombre_anchor, id_bm):
    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(id_bm))
    bm_start.set(qn("w:name"), nombre_anchor)

    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(id_bm))

    parrafo._p.insert(0, bm_start)
    parrafo._p.append(bm_end)


def marcar_nivel_esquema_word(parrafo, nivel=0):
    """
    Marca el párrafo como Título de Nivel 1 (nivel=0) o Nivel 2 (nivel=1) directamente en el XML
    de Word sin pedir 'Heading 1' a la plantilla, evitando por completo el error de estilo inexistente.
    """
    pPr = parrafo._p.get_or_add_pPr()
    outlineLvl = pPr.find(qn("w:outlineLvl"))
    if outlineLvl is None:
        outlineLvl = OxmlElement("w:outlineLvl")
        pPr.append(outlineLvl)
    outlineLvl.set(qn("w:val"), str(nivel))


def insertar_tabla_contenidos_en_pagina_2(doc, items_toc):
    """
    Inserta la Tabla de Contenidos clicable (Ctrl + Clic) en la Página 2 en Arial Negro.
    Cada línea tiene su hipervínculo al marcador del capítulo y su campo PAGEREF actualizable.
    """
    for texto_sec, anchor_name, pag_str, nivel in items_toc:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3.5 if nivel == 1 else 1.5)
        p.paragraph_format.space_after = Pt(3.5 if nivel == 1 else 1.5)
        p.paragraph_format.left_indent = Inches(0.0 if nivel == 1 else 0.25)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.65), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), anchor_name)
        hyperlink.set(qn("w:history"), "1")

        def crear_run_arial_negro(texto_r, es_negrita=True, size_pt=10):
            r_el = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")
            rFonts = OxmlElement("w:rFonts")
            rFonts.set(qn("w:ascii"), "Arial")
            rFonts.set(qn("w:hAnsi"), "Arial")
            rFonts.set(qn("w:cs"), "Arial")
            rPr.append(rFonts)
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

        hyperlink.append(crear_run_arial_negro(texto_sec, es_negrita=(nivel == 1), size_pt=10 if nivel == 1 else 9.5))

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
        fld_pageref.append(crear_run_arial_negro(pag_str, es_negrita=True, size_pt=10 if nivel == 1 else 9.5))
        hyperlink.append(fld_pageref)

        p._p.append(hyperlink)


def buscar_plantilla_en_pc(nombre_archivo="plantilla_hoja_carta_institucional_BLANCO.docx"):
    carpeta_script = os.path.dirname(os.path.abspath(__file__))
    home = os.path.expanduser("~")
    directorios_busqueda = [
        os.getcwd(),
        carpeta_script,
        os.path.join(home, "Downloads"),
        os.path.join(home, "Descargas"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Escritorio"),
        os.path.join(home, "OneDrive", "Escritorio"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive", "Documentos"),
        os.path.join(home, "Documents"),
    ]
    for d in directorios_busqueda:
        if os.path.isdir(d):
            candidato = os.path.join(d, nombre_archivo)
            if os.path.exists(candidato):
                destino_local = os.path.join(carpeta_script, nombre_archivo)
                if os.path.abspath(candidato) != os.path.abspath(destino_local):
                    try:
                        shutil.copy2(candidato, destino_local)
                    except Exception:
                        pass
                return candidato
            for f in os.listdir(d):
                if f.lower().endswith(".docx") and "plantilla" in f.lower() and "blanco" in f.lower() and not f.startswith("~$"):
                    return os.path.join(d, f)
    return None


def preparar_plantilla_ulagos(ruta_plantilla="plantilla_hoja_carta_institucional_BLANCO.docx"):
    ruta_encontrada = buscar_plantilla_en_pc(ruta_plantilla)

    if ruta_encontrada is not None:
        doc = Document(ruta_encontrada)
        usando_plantilla = True

        for anchor in doc.element.iter(qn("wp:anchor")):
            anchor.set("behindDoc", "1")
            anchor.set("layoutInCell", "1")
            reemplazado = False
            for wrap_tag in ["wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough", "wp:wrapTopAndBottom"]:
                for w_el in list(anchor.findall(qn(wrap_tag))):
                    idx_w = list(anchor).index(w_el)
                    anchor.remove(w_el)
                    if not reemplazado and anchor.find(qn("wp:wrapNone")) is None:
                        anchor.insert(idx_w, OxmlElement("wp:wrapNone"))
                        reemplazado = True

        parrafos_con_dibujo = []
        for p in list(doc.paragraphs):
            tiene_grafico = (
                any(True for _ in p._p.iter(qn("w:drawing")))
                or any(True for _ in p._p.iter(qn("w:pict")))
            )
            if tiene_grafico:
                parrafos_con_dibujo.append(p)
            else:
                p._element.getparent().remove(p._element)

        if len(parrafos_con_dibujo) > 0:
            p_base = parrafos_con_dibujo[0]
            p_base.paragraph_format.space_before = Pt(0)
            p_base.paragraph_format.space_after = Pt(0)
            p_base.paragraph_format.line_spacing = 1.0
            for p_extra in parrafos_con_dibujo[1:]:
                for r in list(p_extra._p.findall(qn("w:r"))):
                    p_base._p.append(r)
                p_extra._element.getparent().remove(p_extra._element)

        sec0 = doc.sections[0]
        sec0.top_margin = Inches(0.85)
        sec0.bottom_margin = Inches(0.85)
        sec0.left_margin = Inches(0.85)
        sec0.right_margin = Inches(0.85)
    else:
        doc = Document()
        usando_plantilla = False
        sec0 = doc.sections[0]
        sec0.page_width = Inches(8.5)
        sec0.page_height = Inches(11.0)
        sec0.top_margin = Inches(0.95)
        sec0.bottom_margin = Inches(0.9)
        sec0.left_margin = Inches(0.9)
        sec0.right_margin = Inches(0.9)
        sec0.different_first_page_header_footer = True

        hdr = sec0.header
        ph = hdr.paragraphs[0]
        ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rh = ph.add_run("UNIVERSIDAD DE LOS LAGOS  |  INGENIERÍA COMERCIAL — CAMPUS OSORNO")
        rh.font.name = "Arial"
        rh.font.size = Pt(8.5)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor(0, 0, 0)

    try:
        forzar_fuente_arial_negro(doc.styles["Normal"], 10.5, bold=False)
    except Exception:
        pass

    return doc, usando_plantilla


def agregar_tabla_word(doc, df, incluir_indice=False, fuente_pt=8.0, col_izq_indices=(0,)):
    df_work = df.reset_index() if incluir_indice else df.copy()
    filas, cols = df_work.shape

    tabla = doc.add_table(rows=filas + 1, cols=cols)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = True

    hdr_cells = tabla.rows[0].cells
    for c_idx, col_name in enumerate(df_work.columns):
        cell = hdr_cells[c_idx]
        sombrear_celda(cell, "1E293B")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx in col_izq_indices else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(str(col_name))
        run.font.name = "Arial"
        run.font.size = Pt(fuente_pt)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        configurar_borde_celda(
            cell,
            top={"val": "single", "sz": "6", "color": "000000"},
            bottom={"val": "single", "sz": "8", "color": "000000"},
            left={"val": "single", "sz": "4", "color": "475569"},
            right={"val": "single", "sz": "4", "color": "475569"}
        )

    for r_idx in range(filas):
        row_cells = tabla.rows[r_idx + 1].cells
        texto_primera = str(df_work.iloc[r_idx, 0]).upper()
        es_destacada = any(k in texto_primera for k in ["TOTAL", "(=)", "FLUJO DE CAJA", "FLUJO DE BENEFICIOS", "CAPITAL DE TRABAJO REQUERIDO"])
        bg_hex = "E2E8F0" if es_destacada else ("FFFFFF" if r_idx % 2 == 0 else "F8FAFC")

        for c_idx in range(cols):
            cell = row_cells[c_idx]
            sombrear_celda(cell, bg_hex)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            val = df_work.iloc[r_idx, c_idx]
            col_nom = str(df_work.columns[c_idx])

            if isinstance(val, (int, float, np.integer, np.floating)):
                if any(k in col_nom for k in ["Cant.", "Unidades"]) or ("PRODUCCIÓN" in texto_primera and c_idx > 0) or ("TOTAL PRODUCTOS" in texto_primera and c_idx > 0):
                    txt_val = f"{int(round(val)):,}".replace(",", ".")
                elif any(k in col_nom for k in ["Ton", "Día", "Factor"]) or ("TONELADAS" in texto_primera and c_idx > 0):
                    txt_val = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                elif "Kg al Mes" in col_nom:
                    txt_val = f"{int(round(val)):,}".replace(",", ".")
                else:
                    txt_val = fmt_clp(val)
                alineacion = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                txt_val = str(val)
                alineacion = WD_ALIGN_PARAGRAPH.LEFT if c_idx in col_izq_indices else WD_ALIGN_PARAGRAPH.CENTER

            p = cell.paragraphs[0]
            p.alignment = alineacion
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            run = p.add_run(txt_val)
            run.font.name = "Arial"
            run.font.size = Pt(fuente_pt)
            run.font.bold = es_destacada or (c_idx == 0)
            run.font.color.rgb = RGBColor(0, 0, 0)

            configurar_borde_celda(
                cell,
                top={"val": "single", "sz": "6" if es_destacada else "4", "color": "000000" if es_destacada else "CBD5E1"},
                bottom={"val": "double" if es_destacada else "single", "sz": "8" if es_destacada else "4", "color": "000000" if es_destacada else "CBD5E1"},
                left={"val": "single", "sz": "4", "color": "CBD5E1"},
                right={"val": "single", "sz": "4", "color": "CBD5E1"}
            )

    p_esp = doc.add_paragraph()
    p_esp.paragraph_format.space_before = Pt(2)
    p_esp.paragraph_format.space_after = Pt(4)


def agregar_bloque_notas_pie(doc, lista_notas):
    p_linea = doc.add_paragraph()
    p_linea.paragraph_format.space_before = Pt(8)
    p_linea.paragraph_format.space_after = Pt(2)
    r_lin = p_linea.add_run("________________________________________")
    r_lin.font.name = "Arial"
    r_lin.font.size = Pt(8.5)
    r_lin.font.bold = True
    r_lin.font.color.rgb = RGBColor(0, 0, 0)

    p_tit = doc.add_paragraph()
    p_tit.paragraph_format.space_before = Pt(0)
    p_tit.paragraph_format.space_after = Pt(2)
    rt = p_tit.add_run("NOTAS A PIE DE PÁGINA — REFERENCIAS DE PRECIOS, COSTOS Y PARÁMETROS:")
    rt.font.name = "Arial"
    rt.font.size = Pt(8.5)
    rt.font.bold = True
    rt.font.color.rgb = RGBColor(0, 0, 0)

    for nota_txt in lista_notas:
        pn = doc.add_paragraph()
        pn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pn.paragraph_format.space_before = Pt(1)
        pn.paragraph_format.space_after = Pt(1.5)
        pn.paragraph_format.line_spacing = 1.05
        rn = pn.add_run(nota_txt)
        rn.font.name = "Arial"
        rn.font.size = Pt(8.0)
        rn.font.color.rgb = RGBColor(0, 0, 0)


def actualizar_toc_y_exportar_pdf_word(ruta_docx, ruta_pdf):
    if not sys.platform.startswith("win"):
        return False
    abs_docx = os.path.abspath(ruta_docx).replace("'", "''")
    abs_pdf = os.path.abspath(ruta_pdf).replace("'", "''")
    ps_script = (
        "$ErrorActionPreference = 'Stop'; "
        "$word = New-Object -ComObject Word.Application; "
        "$word.Visible = $false; "
        "$word.DisplayAlerts = 0; "
        f"$doc = $word.Documents.Open('{abs_docx}'); "
        "$doc.Fields.Update() | Out-Null; "
        "$doc.Save(); "
        f"$doc.SaveAs([ref]'{abs_pdf}', [ref]17); "
        "$doc.Close(); "
        "$word.Quit();"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            check=True,
            timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        return os.path.exists(ruta_pdf)
    except Exception:
        return False


def generar_informe_word(resultados, ruta_word="Vargas,Muñoz,Jofre.docx", generar_pdf=True):
    if os.path.exists(ruta_word):
        try:
            os.remove(ruta_word)
        except PermissionError:
            raise PermissionError(
                f"El archivo '{os.path.basename(ruta_word)}' está abierto en Word.\n"
                f"Ciérralo primero para poder actualizarlo."
            )

    carpeta_img = os.path.dirname(os.path.abspath(ruta_word))
    cfg = resultados["cfg"]
    m = resultados["metricas"]
    iva = float(cfg["IVA"])
    ruta_img_flujos, ruta_img_carton = generar_graficos_informe(resultados, carpeta_destino=carpeta_img)

    doc, usando_plantilla = preparar_plantilla_ulagos("plantilla_hoja_carta_institucional_BLANCO.docx")

    indent_sup_izq = Inches(2.15) if usando_plantilla else Inches(0.4)
    indent_sup_der = Inches(0.15) if usando_plantilla else Inches(0.4)

    indent_inf_izq = Inches(2.15) if usando_plantilla else Inches(0.4)
    indent_inf_der = Inches(2.45) if usando_plantilla else Inches(0.4)

    def add_p_portada(texto, size_pt=11, bold=True, italic=False, left_ind=indent_sup_izq, right_ind=indent_sup_der, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = left_ind
        p.paragraph_format.right_indent = right_ind
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(texto)
        r.font.name = "Arial"
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    add_p_portada("", size_pt=10, space_before=42 if usando_plantilla else 18, space_after=0)
    add_p_portada("DEPARTAMENTO DE CIENCIAS ADMINISTRATIVAS Y ECONÓMICAS", size_pt=10.5, bold=True, space_after=2)
    add_p_portada("CARRERA DE INGENIERÍA COMERCIAL — CAMPUS OSORNO", size_pt=10, bold=True, space_after=14)

    add_p_portada("EVALUACIÓN PRIVADA Y SOCIAL DE PROYECTO DE ECONOMÍA CIRCULAR:", size_pt=10.5, bold=True, space_after=4)
    add_p_portada(
        f"“{cfg['NOMBRE_PROYECTO'].upper()}: REUTILIZACIÓN DE CARTÓN DEL EJE COMERCIAL ELEUTERIO RAMÍREZ PARA FABRICAR MATERIAL DIDÁCTICO ESCOLAR Y JUEGOS DE ADULTO MAYOR EN OSORNO”",
        size_pt=13.0, bold=True, space_after=10
    )

    add_p_portada(
        "Estudio a 10 años basado en 6 locales reales de calle Eleuterio Ramírez (#446 al #1024, retiro antes del camión municipal de las 06:00–07:00 AM hacia el Vertedero de Curaco) y venta a 3 compradores institucionales de Osorno (DAEM Osorno, Jardines Infantiles Privados y Hogar de Ancianos Santa María).",
        size_pt=9.5, bold=False, italic=True, space_after=20
    )

    add_p_portada("ASIGNATURA:", size_pt=9.5, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_before=8, space_after=1)
    add_p_portada("Formulación y Evaluación de Proyectos", size_pt=10, bold=False, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=8)

    add_p_portada("DOCENTE GUÍA:", size_pt=9.5, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=1)
    add_p_portada(cfg["DOCENTE"], size_pt=10, bold=False, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=8)

    add_p_portada("EQUIPO FORMULADOR (INTEGRANTES):", size_pt=9.5, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=2)
    for integ in [x.strip() for x in cfg["INTEGRANTES"].split(",") if x.strip()]:
        add_p_portada(f"• {integ}", size_pt=9.0, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=2)

    add_p_portada("CIUDAD Y FECHA:", size_pt=9.5, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_before=8, space_after=1)
    add_p_portada("Osorno, Región de Los Lagos", size_pt=9.5, bold=True, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=1)
    add_p_portada("Noviembre de 2026", size_pt=9.5, bold=False, left_ind=indent_inf_izq, right_ind=indent_inf_der, space_after=4)

    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_body.different_first_page_header_footer = False
    sec_body.header.is_linked_to_previous = True
    sec_body.footer.is_linked_to_previous = True
    sec_body.top_margin = Inches(1.18)
    sec_body.bottom_margin = Inches(0.85)
    sec_body.left_margin = Inches(0.9)
    sec_body.right_margin = Inches(0.9)

    p_tit_toc = doc.add_paragraph()
    p_tit_toc.paragraph_format.space_before = Pt(6)
    p_tit_toc.paragraph_format.space_after = Pt(12)
    r_tt = p_tit_toc.add_run("TABLA DE CONTENIDOS (ÍNDICE GENERAL)")
    r_tt.font.name = "Arial"
    r_tt.font.size = Pt(13)
    r_tt.font.bold = True
    r_tt.font.color.rgb = RGBColor(0, 0, 0)

    items_toc = [
        ("RESUMEN EJECUTIVO", "_Toc_Sec_0", "3", 1),
        ("1. EL PROBLEMA EN OSORNO Y LOS 6 LOCALES DE CALLE ELEUTERIO RAMÍREZ", "_Toc_Sec_1", "4", 1),
        ("2. PRODUCTOS FABRICADOS Y LOS 3 COMPRADORES EN OSORNO", "_Toc_Sec_2", "5", 1),
        ("3. FUNCIONAMIENTO DEL TALLER Y APROVECHAMIENTO DEL CARTÓN (70% / 30%)", "_Toc_Sec_3", "6", 1),
        ("4. TRABAJADORES, SUELDOS Y PRECIOS SOCIALES MIDESO (PUNTO 4.3)", "_Toc_Sec_4", "7", 1),
        ("5. TABLAS DE PRODUCCIÓN, INVERSIONES, COSTOS Y VENTAS (AÑOS 1 AL 10)", "_Toc_Sec_5", "8", 1),
        ("5.a Cantidad de Cartón Recolectado y Juegos Fabricados por Año", "_Toc_Sub_5a", "8", 2),
        ("5.b Inversiones en Máquinas en el Año 0, Renovación y Depreciación", "_Toc_Sub_5b", "8", 2),
        ("5.c Costos Variables, Gastos Fijos y Sueldos por Año", "_Toc_Sub_5c", "9", 2),
        ("5.d Ingresos Anuales por Ventas sin IVA (Años 1 al 10)", "_Toc_Sub_5d", "9", 2),
        ("6. CÁLCULO DEL CAPITAL DE TRABAJO (45 DÍAS DE DESFASE — PUNTO 6)", "_Toc_Sec_6", "10", 1),
        ("7. FLUJO DE CAJA PRIVADO A 10 AÑOS Y RENTABILIDAD (PUNTO 7)", "_Toc_Sec_7", "11", 1),
        ("8. EVALUACIÓN SOCIAL DEL PROYECTO (FLUJO SOCIAL MIDESO — PUNTO 8)", "_Toc_Sec_8", "12", 1),
        ("8.1 Resumen Comparativo: Resultados Privados (12%) vs. Sociales (6%)", "_Toc_Sub_81", "12", 2),
        ("9. CONCLUSIONES PRINCIPALES DEL PROYECTO", "_Toc_Sec_9", "13", 1),
        ("ANEXO 1: EL PORQUÉ DE CADA CIFRA (EXPLICACIÓN FÁCIL DE ENTENDER)", "_Toc_Anx_1", "14", 1),
        ("ANEXO 2: RESUMEN DE INFORMANTES Y DATOS DE TERRENO EN OSORNO", "_Toc_Anx_2", "15", 1),
        ("ANEXO 3: COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS EN OSORNO", "_Toc_Anx_3", "15", 1),
        ("ANEXO 4: SUPUESTOS DEL MODELO EXPLICADOS PASO A PASO", "_Toc_Anx_4", "16", 1),
    ]
    insertar_tabla_contenidos_en_pagina_2(doc, items_toc)

    doc.add_page_break()

    bm_counter = [100]

    def add_titulo_seccion(texto, anchor_name=None):
        p = doc.add_paragraph()
        marcar_nivel_esquema_word(p, nivel=0)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(texto)
        r.font.name = "Arial"
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        if anchor_name:
            bm_counter[0] += 1
            agregar_marcador_word(p, anchor_name, bm_counter[0])
        return p

    def add_subtitulo(texto, anchor_name=None):
        p = doc.add_paragraph()
        marcar_nivel_esquema_word(p, nivel=1)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(texto)
        r.font.name = "Arial"
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        if anchor_name:
            bm_counter[0] += 1
            agregar_marcador_word(p, anchor_name, bm_counter[0])
        return p

    def add_parrafo(texto, bold_prefix=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = "Arial"
            rb.font.size = Pt(10.5)
            rb.font.bold = True
            rb.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(texto)
        r.font.name = "Arial"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    add_titulo_seccion("RESUMEN EJECUTIVO", "_Toc_Sec_0")
    add_parrafo(
        f"El proyecto «{cfg['NOMBRE_PROYECTO']}» evalúa a {cfg['HORIZONTE_ANIOS']} años la creación de un taller productivo en el Sector Industrial de Ovejería (Osorno) que recoge cajas de cartón limpias del comercio céntrico para transformarlas en juegos educativos para niños y juegos de memoria para adultos mayores. El problema actual en el centro de Osorno es muy concreto: todos los días entre las 06:00 y las 07:00 de la mañana pasa el camión recolector municipal por calle Eleuterio Ramírez y se lleva las cajas de cartón que dejan las tiendas, mezclándolas con basura común y enterrándolas en el Vertedero Provincial de Curaco."
    )
    add_parrafo(
        f"Para trabajar con números reales y fáciles de comprobar, nuestro proyecto recoge el cartón de 6 locales comerciales ubicados en calle Eleuterio Ramírez entre el #446 y el #1024 (Hojalatería Lynch #446, Autoplanet #582, Zapatería La Malagueña #625, Minimarket El Sabrosón #647, Galería Centrosorno #977 y Liquimax #1024). Nuestro furgón pasa a las 05:45 AM —justo antes que el camión de la basura— y retira en total 50 kilos por día trabajado. En 25 días al mes son 1.250 kilos mensuales, lo que equivale exactamente a {float(cfg['RESIDUO_TON_ANIO_1']):.2f} toneladas al año en el Año 1."
    )
    add_parrafo(
        f"De esas 15,00 toneladas anuales, el 70% (10,50 toneladas) se aprovecha directamente en los productos y el 30% (4,50 toneladas) se pierde en recortes sobrantes y bordes con corchetes o cinta. Con las 10,50 toneladas útiles fabricamos dos productos validados con 3 compradores de Osorno: (1) {cfg['P1_NOMBRE']} ({int(cfg['P1_UNIDADES']):,} unidades al año que usan 2,5 kilos de cartón cada una y se venden a {fmt_clp(cfg['P1_PRECIO_BRUTO'])} con IVA [Nota 1] al DAEM de Osorno y a Jardines Infantiles Privados); y (2) {cfg['P2_NOMBRE']} ({int(cfg['P2_UNIDADES']):,} unidades al año que usan 2,0 kilos de cartón cada una y se venden a {fmt_clp(cfg['P2_PRECIO_BRUTO'])} con IVA [Nota 1] al Hogar de Ancianos Santa María y clubes de adulto mayor)."
    )
    add_parrafo(
        f"En la Evaluación Privada (Punto 7), invirtiendo $16.000.000 con IVA en máquinas y arreglo del galpón ({fmt_clp(m['inv_inicial_priv'])} netos sin IVA [Nota 2]) más un capital de trabajo inicial de {fmt_clp(m['ct_inicial'])} para cubrir 45 días de operación, y exigiendo un 12% de rentabilidad anual, el proyecto logra un VAN Privado positivo de {fmt_clp(m['van_privado'])}, una TIR de {m['tir_privada']:.2f}% y recupera el dinero en {m['payback_privado']} años. En la Evaluación Social MIDESO (Punto 8), con la tasa social del 6% anual y sumando el ahorro municipal por no enterrar este cartón en el Vertedero de Curaco ({fmt_clp(cfg['AHORRO_MUNI_TON'])} por tonelada con IVA [Nota 5]), el VAN Social sube a {fmt_clp(m['van_social'])} y la TIR Social a {m['tir_social']:.2f}%."
    )

    add_titulo_seccion("1. EL PROBLEMA EN OSORNO Y LOS 6 LOCALES DE CALLE ELEUTERIO RAMÍREZ", "_Toc_Sec_1")
    add_parrafo(
        "El Vertedero Provincial de Curaco en Osorno tiene cada vez menos espacio porque recibe a diario toneladas de cartón limpio que podrían aprovecharse. En la calle Eleuterio Ramírez, las ferreterías, tiendas de repuestos, zapaterías, minimarkets y mayoristas desocupan todos los días cajas de cartón grueso y resistente. Como no tienen bodega para guardarlas, las sacan temprano a la vereda y el camión de la basura municipal se las lleva entre las 06:00 y las 07:00 de la mañana."
    )
    add_parrafo(
        "Fuimos a terreno en calle Eleuterio Ramírez (entre los números #446 y #1024) y registramos 6 comercios reales que entregan en conjunto 50 kilos de cartón limpio por día hábil (1.250 kilos al mes en 25 días trabajados = 15,00 toneladas al año, ver respaldo en '2_Excel_Informantes_Encuestas_Osorno.xlsx'). El detalle de cada local se muestra en la Tabla 1.1:"
    )
    agregar_tabla_word(doc, resultados["df_catastro_6"], incluir_indice=False, fuente_pt=7.8, col_izq_indices=(1, 2, 4))

    add_titulo_seccion("2. PRODUCTOS FABRICADOS Y LOS 3 COMPRADORES EN OSORNO", "_Toc_Sec_2")
    add_parrafo(
        "Set Didáctico Escolar (usa 2,50 kilos de cartón por unidad — Precio: $16.000 con IVA / $13.445 sin IVA [Nota 1]): ",
        bold_prefix="• Producto 1 — "
    )
    add_parrafo(
        "Caja con letras móviles, figuras geométricas encajables y tableros para aprender a leer y contar en kínder y primeros años básicos, pintados con pintura al agua segura para niños (certificación europea EN-71)."
    )
    add_parrafo(
        "Juegos Cognitivos Adulto Mayor (usa 2,00 kilos de cartón por unidad — Precio: $19.000 con IVA / $15.966 sin IVA [Nota 1]): ",
        bold_prefix="• Producto 2 — "
    )
    add_parrafo(
        "Tableros grandes con letras y figuras de alto contraste y piezas gruesas fáciles de tomar con la mano, pensados para ejercitar la memoria en personas mayores."
    )
    add_parrafo(
        "En lugar de depender de ventas sueltas al detalle, aseguramos las ventas del Año 1 (4.500 juegos en total) con 3 grupos de compradores institucionales de Osorno entrevistados en terreno (Tabla 2.1):"
    )
    agregar_tabla_word(doc, resultados["df_demanda_3"], incluir_indice=False, fuente_pt=8.0, col_izq_indices=(1, 2, 6))

    add_titulo_seccion("3. FUNCIONAMIENTO DEL TALLER Y APROVECHAMIENTO DEL CARTÓN (70% / 30%)", "_Toc_Sec_3")
    add_parrafo(
        "De lunes a sábado, entre las 05:30 y las 06:15 de la mañana, nuestro furgón pasa por los 6 locales de calle Eleuterio Ramírez antes que el camión municipal (costo de flete: $30.000 con IVA por tonelada [Nota 3]) y lleva las cajas secas al galpón de 140 m² arrendado en el Sector Industrial de Ovejería ($450.000 mensuales [Nota 4]). Allí las cajas se limpian, se cortan con moldes en la máquina troqueladora de rodillos y se pintan con tintas al agua ($1.800 en pintura/barniz + $600 en caja por unidad [Nota 2])."
    )
    add_parrafo(
        "Supuesto de Balance de Masa y Merma Técnica (70% útil / 30% merma): De las 15,00 toneladas que recogemos en el Año 1, el 70% (10,50 toneladas) queda dentro de los 4.500 productos terminados (3.000 sets escolares × 2,5 kg = 7,50 toneladas + 1.500 juegos adulto mayor × 2,0 kg = 3,00 toneladas). El 30% restante (4,50 toneladas) corresponde a la merma normal del proceso (15% por descarte de bordes con corchetes, cintas o humedad matutina y 15% por recorte sobrante del molde de troquelado)."
    )
    if os.path.exists(ruta_img_carton):
        p_img_c = doc.add_paragraph()
        p_img_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img_c.add_run().add_picture(ruta_img_carton, width=Inches(6.0))

    add_titulo_seccion("4. TRABAJADORES, SUELDOS Y PRECIOS SOCIALES MIDESO (PUNTO 4.3)", "_Toc_Sec_4")
    add_parrafo(
        "La empresa funcionará como una Sociedad por Acciones (SpA) en el Régimen Pro Pyme General (paga 25% de impuesto sobre las utilidades). El taller trabaja con 5 personas en total [Nota 6]: 1 Jefe de Planta y Ventas (mano de obra calificada MOC, sueldo de $950.000/mes), 2 Operadores de Máquina (semicalificada MOSC, $600.000/mes c/u) y 2 Operarios de Recolección y Armado (no calificada MONC, $520.000/mes c/u). En la Tabla 4.1 se muestran sus sueldos reales y cómo se ajustan para el Flujo Social usando los factores oficiales de MIDESO (1,00 para MOC; 0,85 para MOSC y 0,68 para MONC):"
    )
    agregar_tabla_word(doc, resultados["df_personal"], incluir_indice=False, fuente_pt=7.8, col_izq_indices=(0, 3))

    agregar_bloque_notas_pie(doc, [
        "[Nota 1] Precios de venta validados en entrevistas con encargados de compras DAEM Osorno (Compra Ágil / SEP), directoras de jardines particulares y coordinación del Hogar de Ancianos Santa María ($16.000 c/IVA Set Escolar y $19.000 c/IVA Juego Cognitivo; un 25%-30% menores que equivalentes en madera).",
        "[Nota 2] Cotizaciones de maquinaria e insumos gráficos (Octubre-Noviembre 2026): Troqueladora de rodillos 1200 mm ($8.500.000 c/IVA), Guillotina y matrices ($3.500.000 c/IVA), Equipo serigráfico al agua ($2.000.000 c/IVA), Habilitación eléctrica SEC/SEREMI galpón Ovejería ($2.000.000 c/IVA), Tintas al agua EN-71 y PVA ($1.800/unid c/IVA) y Caja microcorrugada ($600/unid c/IVA).",
        "[Nota 3] Costo logístico cotizado para furgón utilitario local en ruta matutina calle Eleuterio Ramírez (#446 al #1024) hacia Sector Industrial Ovejería: $30.000 c/IVA por tonelada transportada ($25.210 netos/ton).",
        "[Nota 4] Costos fijos mensuales en Osorno: Arriendo galpón 140 m² en Ovejería ($450.000 exento de IVA), electricidad trifásica Saesa + agua Suralis + internet ($180.000 c/IVA) y contabilidad Pro Pyme + patente municipal + seguros ($150.000 c/IVA).",
        "[Nota 5] Costo municipal ahorrado por tonelada desviada del Vertedero Provincial de Curaco: $48.000/ton c/IVA ($40.336/ton neta), estimado según contratos de recolección y disposición final de la I. Municipalidad de Osorno y registros SINADER.",
        "[Nota 6] Remuneraciones brutas mensuales acordes al mercado laboral de Osorno (Año 2026) y factores de conversión social del Ministerio de Desarrollo Social y Familia (MIDESO/SNI): MOC = 1,00; MOSC = 0,85; MONC = 0,68; Divisa = 1,01."
    ])

    sec_land = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_land.orientation = WD_ORIENT.LANDSCAPE
    sec_land.page_width = Inches(11.0)
    sec_land.page_height = Inches(8.5)
    sec_land.different_first_page_header_footer = False
    sec_land.header.is_linked_to_previous = True
    sec_land.footer.is_linked_to_previous = True
    sec_land.top_margin = Inches(1.12)
    sec_land.bottom_margin = Inches(0.65)
    sec_land.left_margin = Inches(0.65)
    sec_land.right_margin = Inches(0.65)

    add_titulo_seccion("5. TABLAS DE PRODUCCIÓN, INVERSIONES, COSTOS Y VENTAS (AÑOS 1 AL 10)", "_Toc_Sec_5")
    add_subtitulo("5.a Cantidad de Cartón Recolectado y Juegos Fabricados por Año (Crece 5% anual — Ver '3_Excel_Calculos_Proyecciones.xlsx')", "_Toc_Sub_5a")
    agregar_tabla_word(doc, resultados["df_cantidades"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,))

    add_subtitulo("5.b Inversiones en Máquinas en el Año 0 ($16.000.000 con IVA = $13.445.378 sin IVA), Renovación y Depreciación", "_Toc_Sub_5b")
    agregar_tabla_word(doc, resultados["df_inversiones"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,))
    agregar_tabla_word(doc, resultados["df_depreciacion"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,))

    add_subtitulo("5.c Costos Variables, Gastos Fijos y Sueldos por Año (Valores Netos sin IVA)", "_Toc_Sub_5c")
    agregar_tabla_word(doc, resultados["df_costos_priv"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,))

    add_subtitulo("5.d Ingresos Anuales por Ventas sin IVA (Años 1 al 10)", "_Toc_Sub_5d")
    agregar_tabla_word(doc, resultados["df_ingresos"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,))

    add_titulo_seccion("6. CÁLCULO DEL CAPITAL DE TRABAJO (45 DÍAS DE DESFASE — PUNTO 6)", "_Toc_Sec_6")
    add_parrafo(
        f"Supuesto de Capital de Trabajo: Como el DAEM de Osorno y las instituciones demoran entre 30 y 45 días en pagar las facturas, necesitamos contar con dinero en caja para cubrir {int(cfg['DIAS_DESFASE_CT'])} días de costos de funcionamiento ((Costos Anuales × {int(cfg['DIAS_DESFASE_CT'])}) / 365). Por eso en el Año 0 se ponen {fmt_clp(m['ct_inicial'])} de Capital de Trabajo, el cual se recupera completo al terminar el Año 10:"
    )
    agregar_tabla_word(doc, resultados["df_capital_trabajo"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,))

    add_titulo_seccion("7. FLUJO DE CAJA PRIVADO A 10 AÑOS Y RENTABILIDAD (PUNTO 7)", "_Toc_Sec_7")
    add_parrafo(
        f"En la Tabla 7.1 se presenta el Flujo de Caja Privado sin IVA (respaldado con fórmulas editables en '1_Excel_Flujos_Editables.xlsx'), descontando el 25% de Impuesto a la Renta y evaluando al {float(cfg['TASA_DESCUENTO_PRIVADA'])*100:.1f}% de tasa de descuento privada:"
    )
    agregar_tabla_word(doc, resultados["df_flujo_privado"], incluir_indice=True, fuente_pt=7.0, col_izq_indices=(0,))

    add_titulo_seccion("8. EVALUACIÓN SOCIAL DEL PROYECTO (FLUJO SOCIAL MIDESO — PUNTO 8)", "_Toc_Sec_8")
    add_parrafo(
        f"En el Flujo Social (Tabla 8.1) se eliminan los impuestos (IVA, aranceles e impuesto a la renta), se ajustan los sueldos con los factores sociales de MIDESO y se suma como beneficio para Osorno el dinero que ahorra la Municipalidad por cada tonelada de cartón que no tiene que recoger ni enterrar en el Vertedero de Curaco ({fmt_clp(cfg['AHORRO_MUNI_TON']/(1+iva))} netos por tonelada [Nota 5]), más {fmt_clp(cfg['OTROS_BEN_SOC_ANUAL'])} anuales por beneficios ambientales y educativos:"
    )
    agregar_tabla_word(doc, resultados["df_flujo_social"], incluir_indice=True, fuente_pt=7.1, col_izq_indices=(0,))

    add_subtitulo("8.1 Resumen Comparativo: Resultados Privados (12%) vs. Sociales (6%)", "_Toc_Sub_81")
    agregar_tabla_word(doc, resultados["df_indicadores"], incluir_indice=False, fuente_pt=8.2, col_izq_indices=(0,))

    sec_port_final = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_port_final.orientation = WD_ORIENT.PORTRAIT
    sec_port_final.page_width = Inches(8.5)
    sec_port_final.page_height = Inches(11.0)
    sec_port_final.different_first_page_header_footer = False
    sec_port_final.header.is_linked_to_previous = True
    sec_port_final.footer.is_linked_to_previous = True
    sec_port_final.top_margin = Inches(1.18)
    sec_port_final.bottom_margin = Inches(0.85)
    sec_port_final.left_margin = Inches(0.9)
    sec_port_final.right_margin = Inches(0.9)

    if os.path.exists(ruta_img_flujos):
        p_img_f = doc.add_paragraph()
        p_img_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img_f.add_run().add_picture(ruta_img_flujos, width=Inches(6.4))

    add_titulo_seccion("9. CONCLUSIONES PRINCIPALES DEL PROYECTO", "_Toc_Sec_9")
    add_parrafo(
        "1. Cifras claras y comprobables en terreno: Recogiendo apenas 50 kilos diarios de cartón en 6 locales reales de calle Eleuterio Ramírez (#446 al #1024) antes de las 06:00 AM, juntamos 15 toneladas al año. Usando el 70% (10,5 toneladas) y dejando un 30% de margen por pérdidas y recortes, fabricamos exactamente los 4.500 juegos anuales que demandan el DAEM, los jardines privados y los centros de adulto mayor de Osorno."
    )
    add_parrafo(
        f"2. Negocio rentable para el inversionista privado (Punto 7): Con una tasa exigida del 12% anual, el proyecto entrega un VAN Privado positivo de {fmt_clp(m['van_privado'])}, una TIR de {m['tir_privada']:.2f}% y recupera todo el capital invertido en {m['payback_privado']} años."
    )
    add_parrafo(
        f"3. Alto beneficio social y ambiental para Osorno (Punto 8): Evaluado con la tasa social del 6% de MIDESO, el VAN Social alcanza {fmt_clp(m['van_social'])} y la TIR Social llega a {m['tir_social']:.2f}%, demostrando que el proyecto alivia el Vertedero de Curaco, genera empleo en Osorno y entrega material educativo accesible."
    )

    add_titulo_seccion("ANEXO 1: EL PORQUÉ DE CADA CIFRA (EXPLICACIÓN FÁCIL DE ENTENDER)", "_Toc_Anx_1")
    agregar_tabla_word(doc, resultados["df_porque"], incluir_indice=False, fuente_pt=7.8, col_izq_indices=(0, 1, 2))

    add_titulo_seccion("ANEXO 2: RESUMEN DE INFORMANTES Y DATOS DE TERRENO EN OSORNO", "_Toc_Anx_2")
    agregar_tabla_word(doc, resultados["df_informantes"], incluir_indice=False, fuente_pt=7.8, col_izq_indices=(0, 1, 2, 3))

    add_titulo_seccion("ANEXO 3: COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS EN OSORNO", "_Toc_Anx_3")
    agregar_tabla_word(doc, resultados["df_cotizaciones"], incluir_indice=False, fuente_pt=7.6, col_izq_indices=(0, 2, 6))

    add_titulo_seccion("ANEXO 4: SUPUESTOS DEL MODELO EXPLICADOS PASO A PASO", "_Toc_Anx_4")
    agregar_tabla_word(doc, resultados["df_supuestos"], incluir_indice=False, fuente_pt=7.8, col_izq_indices=(0, 1, 2, 3))

    doc.save(ruta_word)

    ruta_pdf = os.path.splitext(ruta_word)[0] + ".pdf"
    pdf_ok = False
    if generar_pdf:
        pdf_ok = actualizar_toc_y_exportar_pdf_word(ruta_word, ruta_pdf)

    return ruta_word, ruta_pdf, usando_plantilla, pdf_ok


def generar_paquete_entrega_completo(resultados, directorio_raiz):
    apellidos = resultados["cfg"].get("APELLIDOS_ENTREGA", "Vargas,Muñoz,Jofre").strip()
    if os.path.basename(os.path.normpath(directorio_raiz)).lower() == apellidos.lower():
        carpeta_entrega = os.path.normpath(directorio_raiz)
        directorio_padre = os.path.dirname(carpeta_entrega)
    else:
        carpeta_entrega = os.path.join(directorio_raiz, apellidos)
        directorio_padre = directorio_raiz

    os.makedirs(carpeta_entrega, exist_ok=True)

    r_exc1, r_exc2, r_exc3 = exportar_tres_excels_pauta(resultados, carpeta_entrega)

    ruta_word_apellidos = os.path.join(carpeta_entrega, f"{apellidos}.docx")
    ruta_word, ruta_pdf, usando_plantilla, pdf_ok = generar_informe_word(resultados, ruta_word_apellidos, generar_pdf=True)

    for img_temp in ["grafico_flujos_cartonvivo.png", "grafico_carton_ramirez.png"]:
        p_img = os.path.join(carpeta_entrega, img_temp)
        if os.path.exists(p_img):
            try:
                os.remove(p_img)
            except Exception:
                pass

    ruta_zip = os.path.join(directorio_padre, f"{apellidos}.zip")
    if os.path.exists(ruta_zip):
        try:
            os.remove(ruta_zip)
        except Exception:
            pass

    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in os.listdir(carpeta_entrega):
            ruta_completa = os.path.join(carpeta_entrega, archivo)
            if os.path.isfile(ruta_completa):
                zf.write(ruta_completa, arcname=os.path.join(apellidos, archivo))

    return {
        "carpeta_entrega": carpeta_entrega,
        "ruta_zip": ruta_zip,
        "ruta_word": ruta_word,
        "ruta_pdf": ruta_pdf,
        "pdf_ok": pdf_ok,
        "usando_plantilla": usando_plantilla,
        "excels": [r_exc1, r_exc2, r_exc3]
    }


# ==============================================================================
# BLOQUE 6: INTERFAZ GRÁFICA INTERACTIVA (MODO MÁXIMO PODER)
# ==============================================================================

def abrir_carpeta_en_pantalla(carpeta):
    try:
        if sys.platform.startswith("win"):
            os.startfile(carpeta)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", carpeta])
        else:
            subprocess.Popen(["xdg-open", carpeta])
    except Exception:
        pass


class AplicacionCartonVivo:
    def __init__(self, root, carpeta_inicial=None):
        self.root = root
        self.root.title("Evaluador ULagos (Modo Máximo Poder) — Entrega Completa: Word + PDF + 3 Excels + ZIP")
        self.root.geometry("1320x830")
        self.root.configure(bg="#F1F5F9")

        self.carpeta_salida = tk.StringVar(value=carpeta_inicial if carpeta_inicial else os.path.dirname(os.path.abspath(__file__)))

        self.cfg_actual = copy.deepcopy(DEFAULTS)
        self.cfg_actual["INVERSIONES_LISTA"] = copy.deepcopy(INVERSIONES_DEFAULT)
        self.cfg_actual["COSTOS_VAR_LISTA"] = copy.deepcopy(COSTOS_VAR_DEFAULT)
        self.cfg_actual["COSTOS_FIJOS_LISTA"] = copy.deepcopy(COSTOS_FIJOS_DEFAULT)
        self.cfg_actual["PERSONAL_LISTA"] = copy.deepcopy(PERSONAL_DEFAULT)

        self.vars_ui = {}
        self.resultados = ejecutar_motor_financiero(self.cfg_actual)

        self._construir_interfaz()
        self._actualizar_vista_resultados()

    def _construir_interfaz(self):
        frame_top = tk.Frame(self.root, bg="#0F172A", height=68)
        frame_top.pack(fill="x", side="top")

        tk.Label(
            frame_top,
            text="PROYECTO CARTÓN VIVO OSORNO — ENTREGA COMPLETA ULAGOS (WORD + PDF + 3 EXCELS + ZIP)",
            font=("Arial", 13, "bold"),
            fg="#FFFFFF",
            bg="#0F172A"
        ).pack(anchor="w", padx=18, pady=(10, 2))

        tk.Label(
            frame_top,
            text="Plantilla Institucional ULagos Automática + Índice Clicable en Página 2 + Notas al Pie + Carpeta y ZIP 'Vargas,Muñoz,Jofre'",
            font=("Arial", 9, "italic"),
            fg="#94A3B8",
            bg="#0F172A"
        ).pack(anchor="w", padx=18, pady=(0, 8))

        frame_ruta = tk.Frame(self.root, bg="#CBD5E1", pady=6)
        frame_ruta.pack(fill="x", side="top")

        tk.Label(
            frame_ruta, text="📁 Carpeta de Entrega:", font=("Arial", 9, "bold"),
            bg="#CBD5E1", fg="#000000"
        ).pack(side="left", padx=(18, 6))

        ttk.Entry(frame_ruta, textvariable=self.carpeta_salida, width=70).pack(side="left", padx=4)

        tk.Button(
            frame_ruta, text="Cambiar carpeta...", command=self.accion_elegir_carpeta,
            bg="#334155", fg="white", font=("Arial", 9, "bold"), padx=10, pady=2, relief="flat"
        ).pack(side="left", padx=6)

        tk.Button(
            frame_ruta, text="📂 Abrir carpeta en mi PC", command=self.accion_abrir_carpeta,
            bg="#15803D", fg="white", font=("Arial", 9, "bold"), padx=10, pady=2, relief="flat"
        ).pack(side="left", padx=4)

        frame_botones = tk.Frame(self.root, bg="#E2E8F0", pady=8)
        frame_botones.pack(fill="x", side="top")

        tk.Button(
            frame_botones,
            text="🚀 GENERAR ENTREGA COMPLETA (Carpeta y ZIP 'Vargas,Muñoz,Jofre' con Word + PDF + 3 Excels)",
            command=self.accion_generar_todo,
            bg="#7C3AED", fg="white", font=("Arial", 9, "bold"), padx=14, pady=5, relief="flat"
        ).pack(side="left", padx=(18, 8))

        tk.Button(
            frame_botones, text="📄 Guardar Word + PDF Como...", command=self.accion_exportar_word,
            bg="#1E3A8A", fg="white", font=("Arial", 9, "bold"), padx=12, pady=4, relief="flat"
        ).pack(side="left", padx=6)

        tk.Button(
            frame_botones, text="📊 Guardar los 3 Excels en Carpeta...", command=self.accion_exportar_excel,
            bg="#15803D", fg="white", font=("Arial", 9, "bold"), padx=12, pady=4, relief="flat"
        ).pack(side="left", padx=6)

        tk.Button(
            frame_botones, text="⟳ Recalcular", command=self.accion_recalcular,
            bg="#0284C7", fg="white", font=("Arial", 9, "bold"), padx=12, pady=4, relief="flat"
        ).pack(side="left", padx=6)

        self.frame_kpis = tk.Frame(self.root, bg="#F1F5F9", pady=8)
        self.frame_kpis.pack(fill="x", padx=16)
        self.lbls_kpi = {}
        kpis_meta = [
            ("van_privado", "VAN Privado (12%)", "#0284C7"),
            ("tir_privada", "TIR Privada (%)", "#0369A1"),
            ("payback_privado", "Payback Privado", "#0D9488"),
            ("van_social", "VANS Social (6%)", "#16A34A"),
            ("tir_social", "TIRS Social (%)", "#15803D"),
            ("payback_social", "Payback Social", "#4F46E5")
        ]
        for clave, titulo, color in kpis_meta:
            card = tk.Frame(self.frame_kpis, bg="white", bd=1, relief="solid", padx=12, pady=6)
            card.pack(side="left", expand=True, fill="both", padx=5)
            tk.Label(card, text=titulo, font=("Arial", 9, "bold"), fg="#475569", bg="white").pack(anchor="w")
            lbl_val = tk.Label(card, text="-", font=("Arial", 13, "bold"), fg=color, bg="white")
            lbl_val.pack(anchor="w", pady=(2, 0))
            self.lbls_kpi[clave] = lbl_val

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both", padx=16, pady=(4, 14))

        self.tab_porque = tk.Frame(self.notebook, bg="white")
        self.tab_param = tk.Frame(self.notebook, bg="white")
        self.tab_flujo_p = tk.Frame(self.notebook, bg="white")
        self.tab_flujo_s = tk.Frame(self.notebook, bg="white")
        self.tab_catastro = tk.Frame(self.notebook, bg="white")
        self.tab_grafico = tk.Frame(self.notebook, bg="white")

        self.notebook.add(self.tab_porque, text="1. ¿Por qué de cada cifra? (Fácil de Entender)")
        self.notebook.add(self.tab_catastro, text="2. Los 6 Locales y 3 Compradores")
        self.notebook.add(self.tab_flujo_p, text="3. Flujo de Caja Privado (10 Años)")
        self.notebook.add(self.tab_flujo_s, text="4. Flujo Social MIDESO (10 Años)")
        self.notebook.add(self.tab_grafico, text="5. Gráfico Comparativo")
        self.notebook.add(self.tab_param, text="6. Cambiar Datos del Proyecto")

        self.tree_pq = self._crear_treeview_tabla(self.tab_porque)
        self._construir_tab_catastro()
        self.tree_fp = self._crear_treeview_tabla(self.tab_flujo_p)
        self.tree_fs = self._crear_treeview_tabla(self.tab_flujo_s)
        self._construir_tab_parametros()

    def _construir_tab_parametros(self):
        campos = [
            ("NOMBRE_PROYECTO", "Nombre del Proyecto:"),
            ("APELLIDOS_ENTREGA", "Apellidos Carpeta/ZIP (ej. Vargas,Muñoz,Jofre):"),
            ("INTEGRANTES", "Nombres Completos Integrantes:"),
            ("DOCENTE", "Docente Guía ULagos:"),
            ("RESIDUO_TON_ANIO_1", "Cartón Recolectado Año 1 (Toneladas):"),
            ("TASA_CRECIMIENTO", "Crecimiento Anual (0.05 = 5%):"),
            ("P1_UNIDADES", "Unidades Año 1 — Set Didáctico Escolar:"),
            ("P1_PRECIO_BRUTO", "Precio con IVA ($) — Set Didáctico Escolar:"),
            ("P2_UNIDADES", "Unidades Año 1 — Juegos Adulto Mayor:"),
            ("P2_PRECIO_BRUTO", "Precio con IVA ($) — Juegos Adulto Mayor:"),
            ("TASA_DESCUENTO_PRIVADA", "Tasa Descuento Privada (0.12 = 12%):"),
            ("TASA_DESCUENTO_SOCIAL", "Tasa Descuento Social (0.06 = 6%):"),
            ("DIAS_DESFASE_CT", "Días Capital de Trabajo (45 días):"),
            ("AHORRO_MUNI_TON", "Ahorro Municipal Curaco ($/Ton con IVA):")
        ]
        contenedor = tk.Frame(self.tab_param, bg="white", padx=24, pady=16)
        contenedor.pack(fill="both", expand=True)

        for idx, (clave, etiqueta) in enumerate(campos):
            r = idx // 2
            c = (idx % 2) * 2
            tk.Label(contenedor, text=etiqueta, font=("Arial", 9, "bold"), bg="white", fg="#000000").grid(row=r, column=c, sticky="w", padx=(10, 8), pady=6)
            var = tk.StringVar(value=str(self.cfg_actual[clave]))
            self.vars_ui[clave] = var
            ttk.Entry(contenedor, textvariable=var, width=38).grid(row=r, column=c + 1, sticky="w", padx=(0, 24), pady=6)

    def _crear_treeview_tabla(self, parent):
        frame = tk.Frame(parent, bg="white")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        scroll_y = ttk.Scrollbar(frame, orient="vertical")
        scroll_x = ttk.Scrollbar(frame, orient="horizontal")
        tree = ttk.Treeview(frame, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set, show="headings")
        scroll_y.config(command=tree.yview)
        scroll_x.config(command=tree.xview)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        return tree

    def _construir_tab_catastro(self):
        tk.Label(self.tab_catastro, text="Los 6 Locales Reales en Calle Eleuterio Ramírez (#446 al #1024 — 50 kg por día = 1.250 kg al mes = 15 toneladas al año)", font=("Arial", 10, "bold"), bg="white", fg="#000000").pack(anchor="w", padx=12, pady=(8, 2))
        f_top = tk.Frame(self.tab_catastro, bg="white", height=220)
        f_top.pack(fill="both", expand=True, padx=8, pady=2)
        self.tree_cat6 = self._crear_treeview_tabla(f_top)

        tk.Label(self.tab_catastro, text="Los 3 Compradores en Osorno (3.000 Sets Escolares + 1.500 Juegos Adulto Mayor = 4.500 Juegos al Año)", font=("Arial", 10, "bold"), bg="white", fg="#000000").pack(anchor="w", padx=12, pady=(8, 2))
        f_bot = tk.Frame(self.tab_catastro, bg="white", height=180)
        f_bot.pack(fill="both", expand=True, padx=8, pady=2)
        self.tree_dem3 = self._crear_treeview_tabla(f_bot)

    def _poblar_treeview(self, tree, df, incluir_indice=True, anchos_custom=None):
        tree.delete(*tree.get_children())
        df_vis = df.reset_index() if incluir_indice else df.copy()
        cols = list(df_vis.columns)
        tree["columns"] = cols
        for idx, c in enumerate(cols):
            tree.heading(c, text=c)
            if anchos_custom and idx < len(anchos_custom):
                w_col = anchos_custom[idx]
                anc = "w"
            else:
                w_col = 260 if idx == 0 else 115
                anc = "w" if idx == 0 else "e"
            tree.column(c, width=w_col, anchor=anc)
        for _, row in df_vis.iterrows():
            vals = []
            for idx, v in enumerate(row):
                if isinstance(v, (int, float, np.integer, np.floating)) and idx > 0:
                    if isinstance(v, float) and not v.is_integer():
                        vals.append(f"{v:,.2f}")
                    else:
                        vals.append(fmt_clp(v) if abs(v) >= 1000 else f"{int(v)}")
                else:
                    vals.append(str(v))
            tree.insert("", "end", values=vals)

    def _actualizar_vista_resultados(self):
        m = self.resultados["metricas"]
        self.lbls_kpi["van_privado"].config(text=fmt_clp(m["van_privado"]))
        self.lbls_kpi["tir_privada"].config(text=f"{m['tir_privada']:.2f}%")
        self.lbls_kpi["payback_privado"].config(text=f"{m['payback_privado']} años")
        self.lbls_kpi["van_social"].config(text=fmt_clp(m["van_social"]))
        self.lbls_kpi["tir_social"].config(text=f"{m['tir_social']:.2f}%")
        self.lbls_kpi["payback_social"].config(text=f"{m['payback_social']} años")

        self._poblar_treeview(self.tree_pq, self.resultados["df_porque"], incluir_indice=False, anchos_custom=[270, 190, 740])
        self._poblar_treeview(self.tree_fp, self.resultados["df_flujo_privado"], incluir_indice=True)
        self._poblar_treeview(self.tree_fs, self.resultados["df_flujo_social"], incluir_indice=True)
        self._poblar_treeview(self.tree_cat6, self.resultados["df_catastro_6"], incluir_indice=False)
        self._poblar_treeview(self.tree_dem3, self.resultados["df_demanda_3"], incluir_indice=False)

        for w in self.tab_grafico.winfo_children():
            w.destroy()
        fig, ax = plt.subplots(figsize=(9.5, 4.2), dpi=100)
        df_fp = self.resultados["df_flujo_privado"]
        df_fs = self.resultados["df_flujo_social"]
        cols = list(df_fp.columns)
        x = np.arange(len(cols))
        ax.bar(x - 0.18, df_fp.loc["(=) FLUJO DE CAJA PRIVADO NETO"].values / 1e6, 0.35, label="Flujo Privado (MM$)", color="#0284C7")
        ax.bar(x + 0.18, df_fs.loc["(=) FLUJO DE BENEFICIOS SOCIALES NETOS"].values / 1e6, 0.35, label="Flujo Social MIDESO (MM$)", color="#16A34A")
        ax.axhline(0, color="#000000", linewidth=0.9)
        ax.set_xticks(x)
        ax.set_xticklabels(cols)
        ax.set_ylabel("Millones de Pesos ($ CLP)")
        ax.set_title("Comparación Anual: Flujo de Caja Privado vs. Flujo Social MIDESO")
        ax.legend()
        ax.grid(axis="y", linestyle="--", alpha=0.4)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.tab_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        plt.close(fig)

    def _leer_parametros_ui(self):
        for clave, var in self.vars_ui.items():
            val_str = var.get().strip()
            if clave in ["NOMBRE_PROYECTO", "INTEGRANTES", "DOCENTE", "APELLIDOS_ENTREGA"]:
                self.cfg_actual[clave] = val_str
            elif clave in ["P1_UNIDADES", "P2_UNIDADES", "DIAS_DESFASE_CT", "HORIZONTE_ANIOS"]:
                self.cfg_actual[clave] = int(float(val_str))
            else:
                self.cfg_actual[clave] = float(val_str)

    def accion_elegir_carpeta(self):
        carpeta = filedialog.askdirectory(
            initialdir=self.carpeta_salida.get(),
            title="Elige la carpeta donde quieres guardar la entrega"
        )
        if carpeta:
            self.carpeta_salida.set(os.path.normpath(carpeta))

    def accion_abrir_carpeta(self):
        abrir_carpeta_en_pantalla(self.carpeta_salida.get().strip())

    def accion_recalcular(self):
        try:
            self._leer_parametros_ui()
            self.resultados = ejecutar_motor_financiero(self.cfg_actual)
            self._actualizar_vista_resultados()
            messagebox.showinfo("Listo", "Los cálculos se actualizaron correctamente.")
        except Exception as e:
            messagebox.showerror("Error en Datos", f"Revisa los números ingresados:\n{e}")

    def accion_exportar_excel(self):
        try:
            self._leer_parametros_ui()
            self.resultados = ejecutar_motor_financiero(self.cfg_actual)
            self._actualizar_vista_resultados()

            carpeta_sel = filedialog.askdirectory(
                initialdir=self.carpeta_salida.get(),
                title="Elige la carpeta donde guardar los 3 Archivos Excel de la pauta"
            )
            if not carpeta_sel:
                return

            carpeta_sel = os.path.normpath(carpeta_sel)
            self.carpeta_salida.set(carpeta_sel)

            r1, r2, r3 = exportar_tres_excels_pauta(self.resultados, carpeta_sel)
            abrir_carpeta_en_pantalla(carpeta_sel)
            messagebox.showinfo(
                "3 Archivos Excel Creados",
                f"Se guardaron los 3 libros Excel exigidos por la pauta en:\n{carpeta_sel}\n\n"
                f"1) {os.path.basename(r1)}\n"
                f"2) {os.path.basename(r2)}\n"
                f"3) {os.path.basename(r3)}"
            )
        except Exception as e:
            messagebox.showerror("Error al Guardar Excel", str(e))

    def accion_exportar_word(self):
        try:
            self._leer_parametros_ui()
            self.resultados = ejecutar_motor_financiero(self.cfg_actual)
            self._actualizar_vista_resultados()

            apellidos = self.cfg_actual.get("APELLIDOS_ENTREGA", "Vargas,Muñoz,Jofre").strip()
            ruta_elegida = filedialog.asksaveasfilename(
                initialdir=self.carpeta_salida.get(),
                initialfile=f"{apellidos}.docx",
                defaultextension=".docx",
                filetypes=[("Documento de Word", "*.docx")],
                title="Elige dónde guardar el Informe Word (también se creará el PDF)"
            )
            if not ruta_elegida:
                return

            carpeta_sel = os.path.dirname(os.path.abspath(ruta_elegida))
            self.carpeta_salida.set(carpeta_sel)

            ruta_w, ruta_p, uso_pl, pdf_ok = generar_informe_word(self.resultados, ruta_elegida, generar_pdf=True)
            abrir_carpeta_en_pantalla(carpeta_sel)
            messagebox.showinfo(
                "Informe Word y PDF Guardados",
                f"Archivos creados en:\n{carpeta_sel}\n\n"
                f"• Word: {os.path.basename(ruta_w)} ({'Plantilla Oficial ULagos aplicada' if uso_pl else 'Formato ULagos'})\n"
                f"• PDF: {os.path.basename(ruta_p) if pdf_ok else 'Guarda como PDF desde Word'}"
            )
        except Exception as e:
            messagebox.showerror("Error al Guardar Word", str(e))

    def accion_generar_todo(self):
        try:
            self._leer_parametros_ui()
            self.resultados = ejecutar_motor_financiero(self.cfg_actual)
            self._actualizar_vista_resultados()

            carpeta_sel = filedialog.askdirectory(
                initialdir=self.carpeta_salida.get(),
                title="Selecciona dónde crear la carpeta y el ZIP de entrega 'Vargas,Muñoz,Jofre'"
            )
            if not carpeta_sel:
                return

            carpeta_sel = os.path.normpath(carpeta_sel)
            self.carpeta_salida.set(carpeta_sel)

            info = generar_paquete_entrega_completo(self.resultados, carpeta_sel)
            abrir_carpeta_en_pantalla(info["carpeta_entrega"])

            messagebox.showinfo(
                "¡Entrega Completa Lista para Subir!",
                f"Se creó la carpeta y el archivo comprimido:\n"
                f"📦 {info['ruta_zip']}\n\n"
                f"Archivos incluidos adentro:\n"
                f"1) {os.path.basename(info['ruta_word'])} (Word con Plantilla ULagos + Índice en Pág. 2 + Notas al Pie)\n"
                f"2) {os.path.basename(info['ruta_pdf'])} ({'PDF generado y TOC actualizado' if info['pdf_ok'] else 'PDF pendiente'})\n"
                f"3) 1_Excel_Flujos_Editables.xlsx\n"
                f"4) 2_Excel_Informantes_Encuestas_Osorno.xlsx\n"
                f"5) 3_Excel_Calculos_Proyecciones.xlsx"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    carpeta_base = os.path.dirname(os.path.abspath(__file__))
    print(f">>> Generando entrega completa automáticamente en: {carpeta_base}")

    cfg_init = copy.deepcopy(DEFAULTS)
    cfg_init["INVERSIONES_LISTA"] = copy.deepcopy(INVERSIONES_DEFAULT)
    cfg_init["COSTOS_VAR_LISTA"] = copy.deepcopy(COSTOS_VAR_DEFAULT)
    cfg_init["COSTOS_FIJOS_LISTA"] = copy.deepcopy(COSTOS_FIJOS_DEFAULT)
    cfg_init["PERSONAL_LISTA"] = copy.deepcopy(PERSONAL_DEFAULT)

    try:
        res_init = ejecutar_motor_financiero(cfg_init)
        info_init = generar_paquete_entrega_completo(res_init, carpeta_base)
        print(f">>> ¡Listo! Carpeta de entrega: {info_init['carpeta_entrega']}")
        print(f">>> ¡Listo! Archivo ZIP para subir: {info_init['ruta_zip']}")
        abrir_carpeta_en_pantalla(info_init["carpeta_entrega"])
    except Exception as e_init:
        print(f">>> Aviso inicial (cierra Word/Excel si estaban abiertos): {e_init}")

    root = tk.Tk()
    app = AplicacionCartonVivo(root, carpeta_inicial=carpeta_base)
    root.lift()
    root.attributes("-topmost", True)
    root.after(400, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.mainloop()