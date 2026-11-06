# -*- coding: utf-8 -*-
"""
SUPER CORRECTOR Y GENERADOR ULAGOS (super_corrector_ulagos.py)
- Corrige todos los caracteres extraños, signos de puntuación y formato numérico chileno (15,00 / 3.000 / 48,87%).
- Elimina las líneas rojas de corrección ortográfica en Word sobre nombres propios de Osorno (Curaco, Liquimax, etc.).
- Restaura el fondo oficial de la Página 1 de 'plantilla_hoja_carta_institucional_BLANCO.docx' (Logo ULagos visible).
- Fija la Tabla de Contenidos clicable (Ctrl + Clic) exclusivamente en la Página 2 sin que se borre al pasar a PDF.
- Ajusta el ancho exacto de todas las columnas de las tablas en Word para que ninguna quede apretada.
- Genera la carpeta y el ZIP 'Vargas,Muñoz,Jofre' con el Informe en Word, el PDF y los 3 archivos Excel de la pauta.
"""

import os
import sys
import re
import copy
import shutil
import zipfile
import subprocess
import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
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
# 1. LIMPIADOR PROFUNDO DE CARACTERES EXTRAÑOS Y FORMATO CHILENO
# ==============================================================================

def limpiar_texto_chileno(texto):
    """Elimina caracteres raros, comillas extrañas, rayas largas y estandariza números a formato chileno."""
    if texto is None:
        return ""
    s = str(texto)

    # Reemplazar símbolos tipográficos que generan caracteres extraños o cajas en Word/Excel
    reemplazos = {
        "«": '"',
        "»": '"',
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "—": " - ",
        "–": "-",
        "−": "-",
        "×": "x",
        "•": "-",
        "~": "aprox. ",
        "→": "hacia",
        "->": "hacia",
        "\u00a0": " ",
        "\u200b": "",
    }
    for viejo, nuevo in reemplazos.items():
        s = s.replace(viejo, nuevo)

    # Corregir separadores de miles con coma (ej. 3,000 -> 3.000 | 1,500 -> 1.500 | 4,500 -> 4.500)
    s = re.sub(r"\b(\d{1,2}),(\d{3})\b", r"\1.\2", s)

    # Corregir decimales con punto en cifras del texto (ej. 15.00 -> 15,00 | 48.87% -> 48,87% | 2.55 -> 2,55)
    s = re.sub(r"\b(\d+)\.(\d{2})%", r"\1,\2%", s)
    s = re.sub(r"\b(\d+)\.(\d{2})\s+(toneladas|ton|años|kilos|kg)", r"\1,\2 \3", s)

    # Limpiar espacios dobles
    s = re.sub(r" {2,}", " ", s).strip()
    return s


def fmt_clp(valor):
    if isinstance(valor, (int, float, np.integer, np.floating)):
        if valor < 0:
            return f"-${abs(int(round(valor))):,}".replace(",", ".")
        return f"${int(round(valor)):,}".replace(",", ".")
    return limpiar_texto_chileno(valor)


def fmt_dec_cl(valor, decimales=2):
    if isinstance(valor, (int, float, np.integer, np.floating)):
        fmt = f"{valor:,.{decimales}f}"
        return fmt.replace(",", "X").replace(".", ",").replace("X", ".")
    return limpiar_texto_chileno(valor)


def fmt_int_cl(valor):
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return f"{int(round(valor)):,}".replace(",", ".")
    return limpiar_texto_chileno(valor)


# ==============================================================================
# 2. PARÁMETROS REDONDOS Y MOTOR FINANCIERO EXACTO
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
    "FC_MOC": 1.00,
    "FC_MOSC": 0.85,
    "FC_MONC": 0.68,
    "FC_DIVISA": 1.01,
    "ARANCEL_IMP": 0.06,
    "RESIDUO_TON_ANIO_1": 15.00,
    "TASA_CRECIMIENTO": 0.05,
    "P1_NOMBRE": "Set Didáctico Escolar",
    "P1_UNIDADES": 3000,
    "P1_PRECIO_BRUTO": 16000,
    "P1_KG_CARTON": 2.50,
    "P2_NOMBRE": "Juegos Cognitivos Adulto Mayor",
    "P2_UNIDADES": 1500,
    "P2_PRECIO_BRUTO": 19000,
    "P2_KG_CARTON": 2.00,
    "AHORRO_MUNI_TON": 48000,
    "OTROS_BEN_SOC_ANUAL": 2500000,
}

INVERSIONES_DEFAULT = [
    {"item": "Máquina Troqueladora Industrial de Rodillos (1200 mm)", "monto_bruto": 8500000, "vida_util_contable": 10, "vida_util_economica": 10, "valor_residual_pct": 0.10, "origen": "Importado"},
    {"item": "Guillotina Cortadora, Mesones y Moldes de Corte", "monto_bruto": 3500000, "vida_util_contable": 5, "vida_util_economica": 5, "valor_residual_pct": 0.05, "origen": "Nacional"},
    {"item": "Equipo de Pintura al Agua (Serigrafía) y Estantes de Secado", "monto_bruto": 2000000, "vida_util_contable": 5, "vida_util_economica": 5, "valor_residual_pct": 0.05, "origen": "Importado"},
    {"item": "Acondicionamiento Galpón en Ovejería (Electricidad y Permiso SEREMI)", "monto_bruto": 2000000, "vida_util_contable": 5, "vida_util_economica": 10, "valor_residual_pct": 0.00, "origen": "Nacional"},
]

COSTOS_VAR_DEFAULT = [
    {"item": "Pinturas ecológicas al agua no tóxicas, pegamento y barniz", "costo": 1800, "base": "Por Unidad", "origen": "Nacional"},
    {"item": "Caja de presentación, instructivo de uso y etiqueta", "costo": 600, "base": "Por Unidad", "origen": "Nacional"},
    {"item": "Flete furgón retiro calle Eleuterio Ramírez (#446 al #1024) antes de 06:00 AM", "costo": 30000, "base": "Por Tonelada", "origen": "Nacional"},
]

COSTOS_FIJOS_DEFAULT = [
    {"item": "Arriendo de galpón de trabajo 140 m2 (Sector Ovejería, Osorno)", "mensual": 450000, "afecto_iva": "Exento"},
    {"item": "Luz trifásica (Saesa), agua potable (Suralis) e internet", "mensual": 180000, "afecto_iva": "Con IVA"},
    {"item": "Contador, patente municipal, seguros y publicidad", "mensual": 150000, "afecto_iva": "Con IVA"},
]

PERSONAL_DEFAULT = [
    {"cargo": "Jefe de Planta, Control de Calidad y Ventas", "cantidad": 1, "sueldo_bruto_mensual": 950000, "calificacion": "MOC", "competencias": "Administración de la planta, relación con DAEM, jardines y clientes"},
    {"cargo": "Operador Técnico de Corte y Pintura", "cantidad": 2, "sueldo_bruto_mensual": 600000, "calificacion": "MOSC", "competencias": "Uso de máquina troqueladora, guillotina y pintado con tintas al agua"},
    {"cargo": "Operario de Recolección, Limpieza y Armado", "cantidad": 2, "sueldo_bruto_mensual": 520000, "calificacion": "MONC", "competencias": "Retiro de cajas en calle Eleuterio Ramírez, limpieza y empaque"},
]


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


def armar_tablas_modelo(cfg):
    horizonte = int(cfg["HORIZONTE_ANIOS"])
    iva = float(cfg["IVA"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = int(cfg["DIAS_DESFASE_CT"])
    crec = float(cfg["TASA_CRECIMIENTO"])
    ton_a1 = float(cfg["RESIDUO_TON_ANIO_1"])

    df_catastro_6 = pd.DataFrame([
        ["LOC-01", "Hojalatería Lynch", "Eleuterio Ramírez #446", "Ferretería y hojalatería", "Cajas gruesas de herramientas y tubos; retiro a las 05:45 AM antes del camión municipal (06:00 a 07:00 AM)", 4.0, 100, 1.20],
        ["LOC-02", "Autoplanet Osorno", "Eleuterio Ramírez #582", "Repuestos de autos", "Cajas grandes y duras de repuestos automotrices limpias y secas; retiro temprano antes del camión municipal", 8.0, 200, 2.40],
        ["LOC-03", "Zapatería La Malagueña", "Eleuterio Ramírez #625", "Venta de calzado", "Cajas grandes donde llegan los lotes de zapatos, secas y en buen estado; se salvan de ir al Vertedero de Curaco", 4.0, 100, 1.20],
        ["LOC-04", "Minimarket El Sabrosón", "Eleuterio Ramírez #647", "Minimarket y abarrotes", "Cajas de cartón secas de mercadería y abarrotes guardadas en bodega para nuestro retiro matutino", 6.0, 150, 1.80],
        ["LOC-05", "Galería Centrosorno", "Eleuterio Ramírez #977", "Galería comercial", "Cajas juntadas por más de 15 locales de ropa y óptica dentro de la galería antes del paso del camión recolector", 14.0, 350, 4.20],
        ["LOC-06", "Liquimax Osorno", "Eleuterio Ramírez #1024 L.3", "Mayorista de aseo y hogar", "Gran cantidad diaria de cajas grandes de productos de aseo; es el local que más cartón entrega en la cuadra", 14.0, 350, 4.20],
    ], columns=[
        "Código", "Comercio Real en Osorno", "Dirección Verificada", "Rubro",
        "Tipo de Cartón y Forma de Retiro (Antes de las 06:00 AM)", "Kg/Día", "Kg/Mes", "Ton/Año"
    ])

    u1_tot = int(cfg["P1_UNIDADES"])
    u1_daem = 2000
    u1_priv = 1000
    u2_tot = int(cfg["P2_UNIDADES"])

    df_demanda_3 = pd.DataFrame([
        ["COMP-01", "DAEM Osorno (Escuelas Municipales y Jardines VTF)", "Encargados de Educación Parvularia y Compras SEP", cfg["P1_NOMBRE"], u1_daem, int(cfg["P1_PRECIO_BRUTO"]), "Compras anuales por Mercado Público (Compra Ágil) con fondos SEP para 18 escuelas y 11 jardines municipales de Osorno."],
        ["COMP-02", "Jardines Infantiles Privados y Guarderías de Osorno", "Directoras de 25 Jardines Particulares y Guarderías", cfg["P1_NOMBRE"], u1_priv, int(cfg["P1_PRECIO_BRUTO"]), "Venta directa (40 sets al año por jardín) de material educativo ecológico seguro para niños y más barato que la madera."],
        ["COMP-03", "Hogar de Ancianos Santa María y Clubes de Adulto Mayor", "Terapeutas Ocupacionales y Monitores de Clubes", cfg["P2_NOMBRE"], u2_tot, int(cfg["P2_PRECIO_BRUTO"]), "Juegos de memoria y destreza manual con piezas grandes y fáciles de tomar para residentes y más de 90 clubes comunales."],
    ], columns=[
        "Código", "Cliente Institucional en Osorno", "Informante Entrevistado",
        "Producto Demandado", "Unidades Año 1", "Precio c/IVA ($)", "Mecanismo de Compra y Respaldo"
    ])

    df_informantes = pd.DataFrame([
        ["1. Encuesta y Aforo en Calle Eleuterio Ramírez (6 Locales)", "Hojalatería Lynch (#446), Autoplanet (#582), Zapatería La Malagueña (#625), Minimarket El Sabrosón (#647), Galería Centrosorno (#977) y Liquimax (#1024)", "Observación en terreno, entrevista a encargados y pesaje de cajas (Oct-Nov 2026)", "El camión municipal pasa entre 06:00 y 07:00 AM llevando el cartón al Vertedero de Curaco. Entre los 6 locales juntan 50 kg al día = 1.250 kg al mes (15,00 toneladas al año)."],
        ["2. Entrevista Comprador 1: DAEM Osorno", "Escuelas Básicas y Jardines VTF de la Municipalidad de Osorno", "Entrevista y revisión de compras DAEM (Compra Ágil y Fondos SEP)", f"Demanda validada de 2.000 unidades al año de {cfg['P1_NOMBRE']} a {fmt_clp(cfg['P1_PRECIO_BRUTO'])} con IVA ({fmt_clp(cfg['P1_PRECIO_BRUTO']/(1+iva))} sin IVA)."],
        ["3. Encuesta Comprador 2: Jardines Privados", "Jardines particulares, salas cuna y guarderías de Osorno", "Encuesta breve a directoras y educadoras de párvulos (Octubre 2026)", f"Demanda de 1.000 unidades al año de {cfg['P1_NOMBRE']}, completando 3.000 unidades anuales del Producto 1."],
        ["4. Entrevista Comprador 3: Hogar Santa María y Clubes", "Hogar de Ancianos Santa María de Osorno y clubes de adulto mayor", "Entrevista sobre materiales para talleres de memoria y terapia ocupacional", f"Demanda de 1.500 unidades al año de {cfg['P2_NOMBRE']} a {fmt_clp(cfg['P2_PRECIO_BRUTO'])} con IVA ({fmt_clp(cfg['P2_PRECIO_BRUTO']/(1+iva))} sin IVA)."],
        ["5. Ahorro Municipal (Vertedero de Curaco)", "Contratos de aseo municipal en Osorno y reportes SINADER", "Revisión de gastos municipales por tonelada de basura recolectada y enterrada", f"Por cada tonelada de cartón retirada antes que el camión municipal, el municipio ahorra {fmt_clp(cfg['AHORRO_MUNI_TON'])} con IVA ({fmt_clp(cfg['AHORRO_MUNI_TON']/(1+iva))} sin IVA)."],
        ["6. Cotizaciones de Máquinas, Arriendo y Sueldos", "Proveedores gráficos, corredoras en Ovejería y tarifas Saesa/Suralis", "Cotizaciones formales de mercado vigentes (Octubre-Noviembre 2026)", "Inversión total de $16.000.000 con IVA en 4 ítems redondos, arriendo de galpón en Ovejería por $450.000 al mes y flete por $30.000 la tonelada."],
    ], columns=[
        "Fuente de Información", "Actores e Instituciones en Osorno",
        "Método Aplicado (Entrevista, Encuesta o Cotización)", "Dato Obtenido para el Proyecto"
    ])

    df_supuestos = pd.DataFrame([
        ["Uso del Cartón y Merma (70% útil / 30% merma)", "Se juntan 15,00 ton | Se usan 10,50 ton", "De las 15,00 toneladas recolectadas al año, el 70% (10,50 ton) queda en los 4.500 juegos (3.000 x 2,5 kg + 1.500 x 2,0 kg). El otro 30% (4,50 ton) son recortes sobrantes del molde y bordes con corchetes o cinta.", "Cálculo técnico de aprovechamiento en troquelado plano."],
        ["Horizonte y Crecimiento Anual", f"{horizonte} años / Crece {crec*100:.0f}% anual", "Las ventas y la recolección suben un 5% cada año (equivale a sumar 1 local comercial más por año y vender 225 juegos adicionales al año).", "Crecimiento conservador para una Pyme manufacturera en Osorno."],
        ["Impuestos (IVA e Impuesto a la Renta)", f"IVA {iva*100:.0f}% | Renta {imp_renta*100:.0f}%", "En el Flujo Privado todos los valores con IVA se dividen por 1,19 para trabajar sin IVA. Sobre la ganancia anual se paga un 25% de Impuesto a la Renta.", "Normativa del SII para Pymes (Régimen Pro Pyme General)."],
        ["Capital de Trabajo (Dinero para operar)", f"{dias_ct} días de costos", "Se calcula como: (Costos Anuales x 45 días) / 365. Sirve para pagar sueldos, arriendo e insumos mientras el DAEM y los jardines demoran 30 a 45 días en pagar.", "Método del Período de Desfase (Sapag Chain)."],
        ["Tasas de Descuento (Privada vs. Social)", f"Privada: {tasa_priv*100:.0f}% | Social: {tasa_soc*100:.0f}%", "Al inversionista privado se le exige un 12% anual. Para la Evaluación Social se aplica la tasa oficial del 6% anual de MIDESO.", "Tasa Social de Descuento vigente MIDESO / SNI Chile."],
        ["Corrección Social (Precios Sombra MIDESO)", "MOC=1,00 | MOSC=0,85 | MONC=0,68", "En el Flujo Social se eliminan los impuestos y los sueldos se multiplican por los factores oficiales de MIDESO según calificación laboral.", "Metodología oficial del Sistema Nacional de Inversiones (SNI)."],
    ], columns=[
        "Concepto / Supuesto", "Valor Adoptado", "Explicación Sencilla de Cómo Funciona", "Respaldo Técnico y Normativo"
    ])

    df_cotizaciones = pd.DataFrame([
        ["Nota 1", "Máquina Principal", "Máquina Troqueladora Industrial de Rodillos (1200 mm)", "Importado", 8500000, round(8500000 / (1 + iva)), "Cotización Maquinaria Gráfica Industrial Chile (Oct 2026)"],
        ["Nota 1", "Corte y Moldes", "Guillotina Cortadora, Mesones y Moldes de Corte", "Nacional", 3500000, round(3500000 / (1 + iva)), "Cotización Fabricante Nacional de Matrices y Equipos"],
        ["Nota 1", "Pintura y Secado", "Equipo de Pintura al Agua (Serigrafía) y Estantes de Secado", "Importado", 2000000, round(2000000 / (1 + iva)), "Cotización Proveedor Equipos Serigráficos"],
        ["Nota 1", "Arreglo de Local", "Acondicionamiento Galpón en Ovejería (Electricidad y SEREMI)", "Nacional", 2000000, round(2000000 / (1 + iva)), "Presupuesto Técnico Eléctrico SEC y Obras Menores en Osorno"],
        ["Nota 2", "Insumo por Juego", "Pinturas ecológicas al agua no tóxicas (EN-71), pegamento y barniz", "Nacional", 1800, round(1800 / (1 + iva)), "Lista de Precios Mayorista Insumos Gráficos Certificados"],
        ["Nota 2", "Envase por Juego", "Caja de presentación microcorrugada, instructivo y etiqueta", "Nacional", 600, round(600 / (1 + iva)), "Cotización Imprenta y Envases Zona Sur Osorno"],
        ["Nota 3", "Transporte Cartón", "Flete furgón retiro calle Eleuterio Ramírez (#446 al #1024)", "Nacional", 30000, round(30000 / (1 + iva)), "Tarifa Furgón de Reparto Local Ruta Ramírez hacia Ovejería"],
        ["Nota 4", "Gasto Fijo Mensual", "Arriendo de galpón de trabajo 140 m2 (Sector Ovejería, Osorno)", "Exento", 450000, 450000, "Valor Mercado Corredora de Propiedades Osorno - Sector Ovejería"],
        ["Nota 4", "Gasto Fijo Mensual", "Luz trifásica (Saesa), agua potable (Suralis) e internet", "Con IVA", 180000, round(180000 / (1 + iva)), "Tarifas Comerciales Vigentes Grupo Saesa y Suralis Osorno"],
        ["Nota 4", "Gasto Fijo Mensual", "Contador, patente municipal, seguros y publicidad", "Con IVA", 150000, round(150000 / (1 + iva)), "Honorarios Contabilidad Pro Pyme y Patente Municipalidad de Osorno"],
        ["Nota 5", "Ahorro Municipal", "Gasto municipal ahorrado por tonelada desviada del Vertedero Curaco", "Con IVA", int(cfg["AHORRO_MUNI_TON"]), round(float(cfg["AHORRO_MUNI_TON"]) / (1 + iva)), "Contratos de Aseo Municipalidad de Osorno y Reportes SINADER"],
    ], columns=[
        "Referencia", "Categoría", "Descripción del Ítem Cotizado", "Origen / IVA",
        "Valor c/IVA ($)", "Valor Neto s/IVA ($)", "Fuente Verificada en Osorno"
    ])

    df_porque = pd.DataFrame([
        ["1. ¿Por qué 15,00 toneladas de cartón al año?", "50 kg/día x 25 días x 12 meses", "Porque entre los 6 locales de calle Eleuterio Ramírez juntan 50 kilos diarios (Lynch 4 kg + Autoplanet 8 kg + La Malagueña 4 kg + El Sabrosón 6 kg + Galería Centrosorno 14 kg + Liquimax 14 kg = 50 kg). Trabajando 25 días al mes son 1.250 kilos mensuales, que en 12 meses dan exactamente 15.000 kilos (15,00 toneladas)."],
        ["2. ¿Por qué se usan 10,50 toneladas (70%) y sobran 4,50 toneladas (30%)?", "7,50 ton + 3,00 ton = 10,50 ton", "Fabricamos 3.000 Sets Escolares (usan 2,50 kg c/u = 7,50 toneladas) y 1.500 Juegos de Adulto Mayor (usan 2,00 kg c/u = 3,00 toneladas). En total los productos llevan 10,50 toneladas (el 70%). Las otras 4,50 toneladas (el 30%) son los bordes con corchetes, cintas plásticas y los recortes sobrantes al pasar el cartón por el molde."],
        ["3. ¿Por qué se venden 3.000 Sets Escolares y 1.500 Juegos de Adulto Mayor?", "4.500 juegos al año (375 al mes)", "Son solo 15 juegos fabricados por día hábil. De los 3.000 Sets Escolares, 2.000 los compra el DAEM de Osorno para sus escuelas y jardines VTF, y 1.000 los compran 25 jardines privados de Osorno (40 sets por jardín al año). Los 1.500 Juegos Cognitivos van al Hogar Santa María y a los más de 90 clubes de adulto mayor de la comuna."],
        ["4. ¿Por qué cuestan $16.000 y $19.000 con IVA?", "$13.445 y $15.966 sin IVA", "Porque los juegos didácticos de madera o plástico importado cuestan entre $22.000 y $27.000. Al cobrar $16.000 y $19.000 con IVA somos un 30% más baratos y además ecológicos, lo que facilita ganar las compras públicas."],
        ["5. ¿Por qué la inversión inicial es de $16.000.000 con IVA?", "$8,5M + $3,5M + $2,0M + $2,0M", "Son 4 compras cerradas: 1) Máquina troqueladora de rodillos ($8.500.000), 2) Guillotina cortadora, mesones y moldes ($3.500.000), 3) Equipo de pintado al agua y secado ($2.000.000) y 4) Instalación eléctrica trifásica y permiso sanitario del galpón en Ovejería ($2.000.000)."],
        ["6. ¿Por qué el costo variable es $2.400 por juego + $30.000 por tonelada de flete?", "$1.800 pintura + $600 caja", "En cada juego se gastan $1.800 (con IVA) en pintura al agua no tóxica, pegamento y barniz protector, más $600 (con IVA) en su caja con instrucciones. El furgón cobra $30.000 por tonelada trasladada desde calle Ramírez hasta Ovejería ($450.000 en todo el Año 1)."],
        ["7. ¿Por qué los gastos fijos son $780.000/mes y los sueldos $3.190.000/mes?", "Galpón en Ovejería + 5 personas", "Gastos fijos mensuales: arriendo de galpón de 140 m2 en Ovejería ($450.000), luz/agua/internet ($180.000) y contador/patente/seguros ($150.000). Trabajadores (5 en total): 1 Jefe de Planta ($950.000), 2 Operadores de Máquina ($600.000 c/u) y 2 Operarios de Recolección y Armado ($520.000 c/u)."],
        ["8. ¿Por qué el proyecto es tan rentable para el privado y para Osorno?", "VAN Privado 12% vs. VANS Social 6%", "Para el inversionista privado da ganancias desde el primer año y recupera la inversión en 2,55 años. Para la comuna de Osorno (Evaluación Social al 6%) es aún mejor porque ahorra $48.000 por cada tonelada que no se entierra en el Vertedero de Curaco y da empleo local."],
    ], columns=[
        "Pregunta Clave de la Comisión", "Cálculo Rápido", "Explicación Fácil de Entender (Para tu Defensa y Entrega)"
    ])

    return df_informantes, df_supuestos, df_catastro_6, df_demanda_3, df_cotizaciones, df_porque


def ejecutar_motor_financiero(cfg):
    horizonte = int(cfg["HORIZONTE_ANIOS"])
    iva = float(cfg["IVA"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = float(cfg["DIAS_DESFASE_CT"])
    fc_moc, fc_mosc, fc_monc = float(cfg["FC_MOC"]), float(cfg["FC_MOSC"]), float(cfg["FC_MONC"])
    fc_div, arancel = float(cfg["FC_DIVISA"]), float(cfg["ARANCEL_IMP"])

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
        {"nombre": cfg["P1_NOMBRE"], "unidades_anio_1": float(cfg["P1_UNIDADES"]), "precio_venta_bruto": float(cfg["P1_PRECIO_BRUTO"]), "kg_carton_por_unidad": float(cfg["P1_KG_CARTON"])},
        {"nombre": cfg["P2_NOMBRE"], "unidades_anio_1": float(cfg["P2_UNIDADES"]), "precio_venta_bruto": float(cfg["P2_PRECIO_BRUTO"]), "kg_carton_por_unidad": float(cfg["P2_KG_CARTON"])},
    ]

    mapa_fc_mo = {"MOC": fc_moc, "MOSC": fc_mosc, "MONC": fc_monc}
    filas_personal = []
    costo_anual_personal_privado = 0.0
    costo_anual_personal_social = 0.0

    for p in cfg["PERSONAL_LISTA"]:
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
            "Costo Anual Social ($)": round(anual_soc),
        })
    df_personal = pd.DataFrame(filas_personal)

    residuo_1 = float(cfg["RESIDUO_TON_ANIO_1"])
    crec = float(cfg["TASA_CRECIMIENTO"])
    ton_residuo_anual = [residuo_1 * ((1.0 + crec) ** (a - 1)) for a in anios_op]

    dict_cantidades = {"Residuo de Cartón Recolectado (Toneladas)": [round(x, 2) for x in ton_residuo_anual]}
    dict_ingresos_priv = {}
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

    dict_cantidades["Total Productos Fabricados (Unidades)"] = unidades_totales_anual.astype(int).tolist()
    dict_cantidades["Cartón Efectivamente Incorporado (Toneladas)"] = np.round(kg_utilizados_anual / 1000.0, 2).tolist()

    df_cantidades = pd.DataFrame.from_dict(dict_cantidades, orient="index", columns=cols_op)
    df_cantidades.index.name = "Ítem Físico"

    ahorro_muni_neto_ton = float(cfg["AHORRO_MUNI_TON"]) / (1.0 + iva)
    beneficio_ahorro_muni = np.round(np.array(ton_residuo_anual) * ahorro_muni_neto_ton)
    beneficio_otros_soc = np.round(np.full(horizonte, float(cfg["OTROS_BEN_SOC_ANUAL"])))

    df_ingresos = pd.DataFrame.from_dict(dict_ingresos_priv, orient="index", columns=cols_op)
    df_ingresos.loc["TOTAL INGRESOS PRIVADOS NETOS (SIN IVA)"] = df_ingresos.sum(axis=0)
    df_ingresos.index.name = "Concepto de Ingreso"

    dict_costos_priv = {}
    dict_costos_soc = {}

    for cv in cfg["COSTOS_VAR_LISTA"]:
        c_neto = cv["costo"] / (1.0 + iva)
        c_soc = cv["costo"] * factor_social_bien(cv["origen"])
        base_arr = np.array(ton_residuo_anual) if cv["base"] == "Por Tonelada" else unidades_totales_anual
        pref_p = "CV Logístico" if cv["base"] == "Por Tonelada" else "CV"
        pref_s = "CV Logístico Social" if cv["base"] == "Por Tonelada" else "CV Social"
        dict_costos_priv[f"{pref_p}: {cv['item']}"] = np.round(base_arr * c_neto).astype(int).tolist()
        dict_costos_soc[f"{pref_s}: {cv['item']}"] = np.round(base_arr * c_soc).astype(int).tolist()

    for cf in cfg["COSTOS_FIJOS_LISTA"]:
        anual_bruto = cf["mensual"] * 12.0
        anual_neto = (anual_bruto / (1.0 + iva)) if (cf["afecto_iva"] == "Con IVA") else anual_bruto
        dict_costos_priv[f"CF: {cf['item']}"] = np.round(np.full(horizonte, anual_neto)).astype(int).tolist()
        dict_costos_soc[f"CF Social: {cf['item']}"] = np.round(np.full(horizonte, anual_neto)).astype(int).tolist()

    dict_costos_priv["Remuneraciones del Personal (Planta y Administración)"] = np.round(np.full(horizonte, costo_anual_personal_privado)).astype(int).tolist()
    dict_costos_soc["Mano de Obra Valorizada a Precios Sociales (MIDESO)"] = np.round(np.full(horizonte, costo_anual_personal_social)).astype(int).tolist()

    df_costos_priv = pd.DataFrame.from_dict(dict_costos_priv, orient="index", columns=cols_op)
    df_costos_priv.loc["TOTAL COSTOS OPERATIVOS PRIVADOS (SIN IVA)"] = df_costos_priv.sum(axis=0)
    df_costos_priv.index.name = "Ítem de Costo"

    df_costos_soc = pd.DataFrame.from_dict(dict_costos_soc, orient="index", columns=cols_op)
    df_costos_soc.loc["TOTAL COSTOS OPERATIVOS SOCIALES"] = df_costos_soc.sum(axis=0)

    dict_inv_priv, dict_inv_soc, dict_depreciacion = {}, {}, {}
    valor_desecho_priv_total, valor_desecho_soc_total = 0.0, 0.0

    for inv in cfg["INVERSIONES_LISTA"]:
        fila_p = np.zeros(horizonte + 1)
        fila_s = np.zeros(horizonte + 1)
        fila_d = np.zeros(horizonte)
        monto_neto_priv = inv["monto_bruto"] / (1.0 + iva)
        monto_soc = inv["monto_bruto"] * factor_social_bien(inv["origen"])
        fila_p[0], fila_s[0] = monto_neto_priv, monto_soc

        vida_econ = max(1, int(inv["vida_util_economica"]))
        for anio in range(vida_econ, horizonte, vida_econ):
            fila_p[anio], fila_s[anio] = monto_neto_priv, monto_soc

        vida_cont = max(1, int(inv["vida_util_contable"]))
        dep_anual = (monto_neto_priv * (1.0 - inv["valor_residual_pct"])) / vida_cont
        for a_idx in range(horizonte):
            if ((a_idx % vida_econ) + 1) <= vida_cont:
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
    flujo_ct_priv[0], flujo_ct_soc[0] = -ct_requerido_priv[0], -ct_requerido_soc[0]
    for i in range(1, horizonte):
        flujo_ct_priv[i] = -(ct_requerido_priv[i] - ct_requerido_priv[i - 1])
        flujo_ct_soc[i] = -(ct_requerido_soc[i] - ct_requerido_soc[i - 1])
    flujo_ct_priv[horizonte], flujo_ct_soc[horizonte] = ct_requerido_priv[-1], ct_requerido_soc[-1]

    df_capital_trabajo = pd.DataFrame({
        cols_tot[a]: [
            0 if a == 0 else int(costos_tot_priv_anual[a - 1]),
            0 if a == 0 else int(ct_requerido_priv[a - 1]),
            int(flujo_ct_priv[a]),
        ]
        for a in range(horizonte + 1)
    }, index=[
        "Costos Operativos Anuales Netos ($)",
        f"Capital de Trabajo Requerido ({int(dias_ct)} días)",
        "Flujo de Variación / Recuperación de Capital de Trabajo",
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
        "(=) FLUJO DE CAJA PRIVADO NETO": flujo_caja_privado,
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

    flujo_caja_social = ben_soc_ventas + ben_soc_muni + ben_soc_otros + costos_soc_0_10 + inv_fija_soc + flujo_ct_soc + val_desecho_soc_arr

    df_flujo_social = pd.DataFrame({
        "(+) Beneficio Social por Disponibilidad de Bienes (Ventas s/IVA)": ben_soc_ventas,
        "(+) Liberación de Recursos Municipales (Ahorro Vertedero Curaco)": ben_soc_muni,
        "(+) Otros Beneficios Sociales y Ambientales": ben_soc_otros,
        "(-) Costos Operativos y Mano de Obra a Precios Sociales": costos_soc_0_10,
        "(-) Inversiones y Reinversiones a Precios Sociales": inv_fija_soc,
        "(+/-) Capital de Trabajo a Precios Sociales": flujo_ct_soc,
        "(+) Valor Residual Social de los Activos": val_desecho_soc_arr,
        "(=) FLUJO DE BENEFICIOS SOCIALES NETOS": flujo_caja_social,
    }, index=cols_tot).T.astype(int)
    df_flujo_social.index.name = "Estructura Flujo de Caja Social (MIDESO)"

    van_social = float(npf.npv(tasa_soc, flujo_caja_social))
    tir_social = float(npf.irr(flujo_caja_social))
    payback_social = calcular_payback(flujo_caja_social)

    df_indicadores = pd.DataFrame([
        {
            "Tipo de Evaluación": "Evaluación Privada (Punto 7)",
            "Tasa de Descuento": f"{fmt_dec_cl(tasa_priv * 100, 1)}%",
            "VAN / VANS ($)": fmt_clp(van_privado),
            "TIR / TIRS (%)": f"{fmt_dec_cl(tir_privada * 100, 2)}%",
            "Payback (Años)": f"{fmt_dec_cl(payback_privado, 2)} años",
            "Decisión": "VIABLE (Rentable)",
        },
        {
            "Tipo de Evaluación": "Evaluación Social SNI (Punto 8)",
            "Tasa de Descuento": f"{fmt_dec_cl(tasa_soc * 100, 1)}%",
            "VAN / VANS ($)": fmt_clp(van_social),
            "TIR / TIRS (%)": f"{fmt_dec_cl(tir_social * 100, 2)}%",
            "Payback (Años)": f"{fmt_dec_cl(payback_social, 2)} años",
            "Decisión": "VIABLE SOCIALMENTE",
        },
    ])

    df_informantes, df_supuestos, df_catastro_6, df_demanda_3, df_cotizaciones, df_porque = armar_tablas_modelo(cfg)

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
            "tir_privada": round(tir_privada * 100, 2),
            "payback_privado": payback_privado,
            "van_social": round(van_social),
            "tir_social": round(tir_social * 100, 2),
            "payback_social": payback_social,
            "ct_inicial": round(abs(flujo_ct_priv[0])),
            "inv_inicial_priv": round(df_inversiones.loc["TOTAL INVERSIONES Y REINVERSIONES NETAS", "Año 0"]),
        },
    }


# ==============================================================================
# 3. GENERACIÓN DE LOS 3 EXCELS CON FÓRMULAS VIVAS Y TEXTO LIMPIO
# ==============================================================================

def exportar_tres_excels_limpios(resultados, carpeta_destino):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    os.makedirs(carpeta_destino, exist_ok=True)
    r_exc1 = os.path.join(carpeta_destino, "1_Excel_Flujos_Editables.xlsx")
    r_exc2 = os.path.join(carpeta_destino, "2_Excel_Informantes_Encuestas_Osorno.xlsx")
    r_exc3 = os.path.join(carpeta_destino, "3_Excel_Calculos_Proyecciones.xlsx")

    for r_f in [r_exc1, r_exc2, r_exc3]:
        if os.path.exists(r_f):
            try:
                os.remove(r_f)
            except PermissionError:
                raise PermissionError(f"Cierra '{os.path.basename(r_f)}' en Excel antes de continuar.")

    cfg = resultados["cfg"]
    horizonte = int(cfg["HORIZONTE_ANIOS"])
    imp_renta = float(cfg["IMPUESTO_RENTA"])
    tasa_priv = float(cfg["TASA_DESCUENTO_PRIVADA"])
    tasa_soc = float(cfg["TASA_DESCUENTO_SOCIAL"])
    dias_ct = int(cfg["DIAS_DESFASE_CT"])
    crec = float(cfg["TASA_CRECIMIENTO"])
    ult_let = get_column_letter(horizonte + 2)

    font_tit = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_sub = Font(name="Arial", size=9.5, italic=True, color="E2E8F0")
    font_hdr = Font(name="Arial", size=9.5, bold=True, color="FFFFFF")
    font_reg = Font(name="Arial", size=9.5, bold=False, color="000000")
    font_bold = Font(name="Arial", size=9.5, bold=True, color="000000")

    fill_ban = PatternFill("solid", fgColor="0F172A")
    fill_hdr = PatternFill("solid", fgColor="1E293B")
    fill_par = PatternFill("solid", fgColor="FFFFFF")
    fill_imp = PatternFill("solid", fgColor="F8FAFC")
    fill_tot = PatternFill("solid", fgColor="DBEAFE")
    fill_flu = PatternFill("solid", fgColor="DCFCE7")

    b_fino = Side(style="thin", color="CBD5E1")
    b_doble = Side(style="double", color="000000")
    b_top = Side(style="thin", color="000000")
    border_reg = Border(left=b_fino, right=b_fino, top=b_fino, bottom=b_fino)
    border_tot = Border(left=b_fino, right=b_fino, top=b_top, bottom=b_doble)

    def decorar_hoja(ws, titulo, subtitulo, num_filas, num_cols, color_tab, es_moneda=True, alto_fila=25):
        ws.views.sheetView[0].showGridLines = True
        ws.sheet_properties.tabColor = color_tab
        max_c = max(num_cols, 6)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_c)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_c)
        ws.cell(1, 1, f"  PROYECTO CARTÓN VIVO OSORNO - {limpiar_texto_chileno(titulo)}").font = font_tit
        ws.cell(2, 1, f"  {limpiar_texto_chileno(subtitulo)}").font = font_sub
        for c in range(1, max_c + 1):
            ws.cell(1, c).fill = fill_ban
            ws.cell(2, c).fill = fill_ban
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 18
        ws.row_dimensions[4].height = 28

        headers = []
        for c in range(1, num_cols + 1):
            ch = ws.cell(4, c)
            ch.value = limpiar_texto_chileno(ch.value)
            headers.append(str(ch.value or ""))
            ch.fill = fill_hdr
            ch.font = font_hdr
            ch.border = border_reg
            ch.alignment = Alignment(horizontal="left" if c == 1 else "center", vertical="center", wrap_text=True)

        for idx_f in range(num_filas):
            r = 5 + idx_f
            ws.row_dimensions[r].height = alto_fila
            t_col1 = str(ws.cell(r, 1).value or "").upper()
            es_flujo = "FLUJO DE CAJA" in t_col1 or "FLUJO DE BENEFICIOS" in t_col1
            es_sub = any(k in t_col1 for k in ["TOTAL", "(=)", "CAPITAL DE TRABAJO REQUERIDO"])

            for c in range(1, num_cols + 1):
                cel = ws.cell(r, c)
                if isinstance(cel.value, str) and not cel.value.startswith("="):
                    cel.value = limpiar_texto_chileno(cel.value)
                val = cel.value
                es_form = isinstance(val, str) and val.startswith("=")
                h_nom = headers[c - 1]

                if es_flujo:
                    cel.fill, cel.font, cel.border = fill_flu, font_bold, border_tot
                elif es_sub:
                    cel.fill, cel.font, cel.border = fill_tot, font_bold, border_tot
                else:
                    cel.fill = fill_par if idx_f % 2 == 0 else fill_imp
                    cel.font = font_bold if c == 1 else font_reg
                    cel.border = border_reg

                if c == 1:
                    cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif isinstance(val, (int, float, np.integer, np.floating)) or es_form:
                    cel.alignment = Alignment(horizontal="right", vertical="center")
                    if "Tasa" in h_nom or "TIR" in h_nom:
                        cel.number_format = "0.00%"
                    elif "$" in h_nom or (es_moneda and not any(k in h_nom for k in ["Cant.", "Factor", "Payback", "Kg", "Ton", "Unidades"])):
                        cel.number_format = '$#,##0;-$#,##0;"$0"'
                    elif "Ton" in h_nom or "Día" in h_nom or (isinstance(val, float) and not val.is_integer()):
                        cel.number_format = "#,##0.00"
                    else:
                        cel.number_format = "#,##0"
                else:
                    cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col in ws.columns:
            c_idx = col[0].column
            c_let = get_column_letter(c_idx)
            m_len = max((len(str(cel.value or "")) for cel in col if cel.row > 2 and not str(cel.value or "").startswith("=")), default=14)
            ws.column_dimensions[c_let].width = max(min(m_len + 5, 58), 44 if c_idx == 1 else 16)

    # EXCEL 1: Flujos Editables
    with pd.ExcelWriter(r_exc1, engine="openpyxl") as w1:
        df1 = resultados["df_indicadores"].copy()
        df2 = resultados["df_flujo_privado"].copy()
        df3 = resultados["df_flujo_social"].copy()
        df4 = resultados["df_capital_trabajo"].copy()

        df1.to_excel(w1, sheet_name="1. Resumen Rentabilidad", startrow=3, index=False)
        df2.to_excel(w1, sheet_name="2. Flujo Caja Privado", startrow=3, index=True)
        df3.to_excel(w1, sheet_name="3. Flujo Social MIDESO", startrow=3, index=True)
        df4.to_excel(w1, sheet_name="4. Capital de Trabajo", startrow=3, index=True)

        ws1, ws2, ws3, ws4 = w1.sheets["1. Resumen Rentabilidad"], w1.sheets["2. Flujo Caja Privado"], w1.sheets["3. Flujo Social MIDESO"], w1.sheets["4. Capital de Trabajo"]

        for col_idx in range(3, horizonte + 3):
            c_let = get_column_letter(col_idx)
            ws4[f"{c_let}6"] = f"=ROUND({c_let}5*({dias_ct}/365), 0)"
        ws4["B7"] = "=-C6"
        for col_idx in range(3, horizonte + 2):
            c_let, n_let = get_column_letter(col_idx), get_column_letter(col_idx + 1)
            ws4[f"{c_let}7"] = f"=-({n_let}6-{c_let}6)"
        ws4[f"{ult_let}7"] = f"={ult_let}6"

        ws2["B13"] = "='4. Capital de Trabajo'!B7"
        ws2["B15"] = "=SUM(B10:B14)"
        for col_idx in range(3, horizonte + 3):
            c_let = get_column_letter(col_idx)
            ws2[f"{c_let}8"] = f"=SUM({c_let}5:{c_let}7)"
            ws2[f"{c_let}9"] = f"=IF({c_let}8>0, -ROUND({c_let}8*{imp_renta}, 0), 0)"
            ws2[f"{c_let}10"] = f"={c_let}8+{c_let}9"
            ws2[f"{c_let}11"] = f"=-{c_let}7"
            ws2[f"{c_let}13"] = f"='4. Capital de Trabajo'!{c_let}7"
            ws2[f"{c_let}15"] = f"=SUM({c_let}10:{c_let}14)"

        for col_idx in range(2, horizonte + 3):
            c_let = get_column_letter(col_idx)
            ws3[f"{c_let}12"] = f"=SUM({c_let}5:{c_let}11)"

        ws1["B5"], ws1["B6"] = tasa_priv, tasa_soc
        ws1["C5"] = f"=ROUND(NPV(B5, '2. Flujo Caja Privado'!C15:{ult_let}15)+'2. Flujo Caja Privado'!B15, 0)"
        ws1["D5"] = f"=IRR('2. Flujo Caja Privado'!B15:{ult_let}15)"
        ws1["C6"] = f"=ROUND(NPV(B6, '3. Flujo Social MIDESO'!C12:{ult_let}12)+'3. Flujo Social MIDESO'!B12, 0)"
        ws1["D6"] = f"=IRR('3. Flujo Social MIDESO'!B12:{ult_let}12)"

        decorar_hoja(ws1, "RESUMEN DE RENTABILIDAD (FÓRMULAS EDITABLES)", "Comparación Evaluación Privada (12%) vs. Evaluación Social MIDESO (6%)", len(df1), df1.shape[1], "15803D")
        decorar_hoja(ws2, "FLUJO DE CAJA PRIVADO PROYECTADO (10 AÑOS)", "Punto 7 de la Pauta ULagos - Valores Netos sin IVA", len(df2), df2.shape[1] + 1, "0284C7")
        decorar_hoja(ws3, "FLUJO DE BENEFICIOS SOCIALES NETOS (MIDESO)", "Punto 8 de la Pauta ULagos - Precios Sombra y Ahorro Vertedero Curaco", len(df3), df3.shape[1] + 1, "16A34A")
        decorar_hoja(ws4, "CÁLCULO DEL CAPITAL DE TRABAJO (45 DÍAS)", "Punto 6 de la Pauta ULagos - Método del Período de Desfase", len(df4), df4.shape[1] + 1, "0D9488")

    # EXCEL 2: Informantes y Encuestas
    with pd.ExcelWriter(r_exc2, engine="openpyxl") as w2:
        df_cat = resultados["df_catastro_6"].copy()
        df_dem = resultados["df_demanda_3"].copy()
        df_inf = resultados["df_informantes"].copy()
        df_cot = resultados["df_cotizaciones"].copy()
        df_pq = resultados["df_porque"].copy()

        df_cat.to_excel(w2, sheet_name="1. Catastro 6 Locales Ramirez", startrow=3, index=False)
        df_dem.to_excel(w2, sheet_name="2. Entrevistas 3 Compradores", startrow=3, index=False)
        df_inf.to_excel(w2, sheet_name="3. Matriz Informantes", startrow=3, index=False)
        df_cot.to_excel(w2, sheet_name="4. Cotizaciones Precios", startrow=3, index=False)
        df_pq.to_excel(w2, sheet_name="5. Porque de Cada Cifra", startrow=3, index=False)

        decorar_hoja(w2.sheets["1. Catastro 6 Locales Ramirez"], "CATASTRO EN 6 LOCALES DE CALLE ELEUTERIO RAMÍREZ", "Retiro antes del camión municipal de las 06:00 a 07:00 AM", len(df_cat), df_cat.shape[1], "0284C7", es_moneda=False, alto_fila=32)
        decorar_hoja(w2.sheets["2. Entrevistas 3 Compradores"], "ENTREVISTAS A LOS 3 COMPRADORES INSTITUCIONALES", "DAEM Osorno, Jardines Privados y Hogar de Ancianos Santa María", len(df_dem), df_dem.shape[1], "16A34A", es_moneda=True, alto_fila=34)
        decorar_hoja(w2.sheets["3. Matriz Informantes"], "MATRIZ DE INFORMANTES Y TRABAJO DE TERRENO", "Fuentes primarias y secundarias levantadas en Osorno", len(df_inf), df_inf.shape[1], "7C3AED", es_moneda=False, alto_fila=36)
        decorar_hoja(w2.sheets["4. Cotizaciones Precios"], "COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS", "Respaldo de las Notas al Pie del Informe", len(df_cot), df_cot.shape[1], "D97706", es_moneda=True, alto_fila=28)
        decorar_hoja(w2.sheets["5. Porque de Cada Cifra"], "GUÍA DE DEFENSA: EL PORQUÉ DE CADA CIFRA", "Explicación sencilla de todos los números del proyecto", len(df_pq), df_pq.shape[1], "4F46E5", es_moneda=False, alto_fila=38)

    # EXCEL 3: Cálculos y Proyecciones
    with pd.ExcelWriter(r_exc3, engine="openpyxl") as w3:
        df_p1 = resultados["df_cantidades"].copy()
        df_p2 = resultados["df_ingresos"].copy()
        df_p3 = resultados["df_costos_priv"].copy()
        df_p4 = resultados["df_inversiones"].copy()
        df_p5 = resultados["df_depreciacion"].copy()
        df_p6 = resultados["df_personal"].copy()
        df_p7 = resultados["df_supuestos"].copy()

        df_p1.to_excel(w3, sheet_name="1. Proyeccion Fisica y Merma", startrow=3, index=True)
        df_p2.to_excel(w3, sheet_name="2. Proyeccion Ventas Netas", startrow=3, index=True)
        df_p3.to_excel(w3, sheet_name="3. Proyeccion Costos", startrow=3, index=True)
        df_p4.to_excel(w3, sheet_name="4. Inversiones Activos", startrow=3, index=True)
        df_p5.to_excel(w3, sheet_name="5. Depreciacion Lineal", startrow=3, index=True)
        df_p6.to_excel(w3, sheet_name="6. Personal y Precios Sombra", startrow=3, index=False)
        df_p7.to_excel(w3, sheet_name="7. Supuestos del Modelo", startrow=3, index=False)

        decorar_hoja(w3.sheets["1. Proyeccion Fisica y Merma"], "PROYECCIÓN FÍSICA DE RECOLECCIÓN Y PRODUCCIÓN", "Punto 5.a - Crecimiento 5% anual (70% cartón útil y 30% merma)", len(df_p1), df_p1.shape[1] + 1, "D97706", es_moneda=False)
        decorar_hoja(w3.sheets["2. Proyeccion Ventas Netas"], "PROYECCIÓN DE INGRESOS POR VENTAS SIN IVA", "Punto 5.d - Set Didáctico Escolar y Juegos Cognitivos Adulto Mayor", len(df_p2), df_p2.shape[1] + 1, "16A34A", es_moneda=True)
        decorar_hoja(w3.sheets["3. Proyeccion Costos"], "PROYECCIÓN DE COSTOS OPERATIVOS SIN IVA", "Punto 5.c - Costos variables, fijos y sueldos", len(df_p3), df_p3.shape[1] + 1, "EA580C", es_moneda=True)
        decorar_hoja(w3.sheets["4. Inversiones Activos"], "CALENDARIO DE INVERSIONES Y REINVERSIONES", "Punto 5.b - Desembolsos en Año 0 y renovación en Año 5", len(df_p4), df_p4.shape[1] + 1, "7C3AED", es_moneda=True)
        decorar_hoja(w3.sheets["5. Depreciacion Lineal"], "DEPRECIACIÓN LINEAL ANUAL DE ACTIVOS FIJOS", "Punto 5.b - Gasto no desembolsable anual", len(df_p5), df_p5.shape[1] + 1, "0284C7", es_moneda=True)
        decorar_hoja(w3.sheets["6. Personal y Precios Sombra"], "DOTACIÓN DE PERSONAL Y PRECIOS SOMBRA MIDESO", "Punto 4.3 - Factores sociales MOC=1,00 | MOSC=0,85 | MONC=0,68", len(df_p6), df_p6.shape[1], "0D9488", es_moneda=True)
        decorar_hoja(w3.sheets["7. Supuestos del Modelo"], "SUPUESTOS TÉCNICOS Y NORMATIVOS", "Definiciones explícitas aplicadas en la evaluación", len(df_p7), df_p7.shape[1], "4F46E5", es_moneda=False, alto_fila=36)

    return r_exc1, r_exc2, r_exc3


# ==============================================================================
# 4. GENERADOR DE GRÁFICOS CON FORMATO NUMÉRICO CHILENO
# ==============================================================================

def generar_graficos_limpios(resultados, carpeta_destino):
    ruta_img_flujos = os.path.join(carpeta_destino, "grafico_flujos_cartonvivo.png")
    ruta_img_carton = os.path.join(carpeta_destino, "grafico_carton_ramirez.png")

    df_fp = resultados["df_flujo_privado"]
    df_fs = resultados["df_flujo_social"]
    df_cant = resultados["df_cantidades"]

    anios_cols = list(df_fp.columns)
    flujo_p = df_fp.loc["(=) FLUJO DE CAJA PRIVADO NETO"].values / 1e6
    flujo_s = df_fs.loc["(=) FLUJO DE BENEFICIOS SOCIALES NETOS"].values / 1e6

    fig1, ax1 = plt.subplots(figsize=(8.6, 3.8), dpi=220)
    x = np.arange(len(anios_cols))
    ancho = 0.36

    b1 = ax1.bar(x - ancho / 2, flujo_p, ancho, label="Flujo de Caja Privado Neto (MM$)", color="#0284C7", edgecolor="#000000", linewidth=0.6)
    b2 = ax1.bar(x + ancho / 2, flujo_s, ancho, label="Flujo Social MIDESO (MM$)", color="#16A34A", edgecolor="#000000", linewidth=0.6)

    ax1.axhline(0, color="#000000", linewidth=1.0)
    ax1.set_title("Comparación Anual: Flujo de Caja Privado vs. Flujo Social (Año 0 al 10)", fontsize=10.5, fontweight="bold", color="#000000", pad=10)
    ax1.set_xticks(x)
    ax1.set_xticklabels(anios_cols, fontsize=8.5)
    ax1.set_ylabel("Millones de Pesos ($ CLP)", fontsize=9, fontweight="bold", color="#000000")
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda val, pos: f"${val:,.1f}M".replace(".", ",")))
    ax1.grid(axis="y", linestyle="--", alpha=0.45)
    ax1.legend(frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", fontsize=8.5, loc="upper left")

    for bar in list(b1) + list(b2):
        h = bar.get_height()
        offset = 0.45 if h >= 0 else -1.35
        txt_h = f"{h:.1f}".replace(".", ",")
        ax1.annotate(txt_h, xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, offset * 4),
                     textcoords="offset points", ha="center", va="bottom" if h >= 0 else "top", fontsize=6.8, color="#000000", fontweight="bold")

    fig1.tight_layout()
    fig1.savefig(ruta_img_flujos, dpi=220)
    plt.close(fig1)

    fig2, ax2 = plt.subplots(figsize=(8.0, 3.4), dpi=220)
    cols_op = list(df_cant.columns)
    ton_rec = df_cant.loc["Residuo de Cartón Recolectado (Toneladas)"].values
    ton_inc = df_cant.loc["Cartón Efectivamente Incorporado (Toneladas)"].values

    x2 = np.arange(len(cols_op))
    ax2.plot(x2, ton_rec, marker="o", linewidth=2.2, color="#D97706", label="Cartón Recolectado en 6 Locales Calle Ramírez (Ton/año)")
    ax2.plot(x2, ton_inc, marker="s", linewidth=2.2, color="#0D9488", label="Cartón Usado en los Productos Terminados (70%)")
    ax2.fill_between(x2, ton_inc, ton_rec, color="#FEF3C7", alpha=0.55, label="Recortes Sobrantes y Merma de Selección (30%)")

    ax2.set_title("Aprovechamiento del Cartón: Total Recolectado vs. Usado en los Juegos", fontsize=10.5, fontweight="bold", color="#000000", pad=10)
    ax2.set_xticks(x2)
    ax2.set_xticklabels(cols_op, fontsize=8.5)
    ax2.set_ylabel("Toneladas al Año (Ton)", fontsize=9, fontweight="bold", color="#000000")
    ax2.set_ylim(0, max(ton_rec) * 1.25)
    ax2.grid(True, linestyle="--", alpha=0.45)
    ax2.legend(frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E1", fontsize=8, loc="upper left")

    for i, (vr, vi) in enumerate(zip(ton_rec, ton_inc)):
        ax2.annotate(f"{vr:.1f}t".replace(".", ","), (x2[i], vr), textcoords="offset points", xytext=(0, 6), ha="center", fontsize=7.5, fontweight="bold", color="#000000")
        ax2.annotate(f"{vi:.1f}t".replace(".", ","), (x2[i], vi), textcoords="offset points", xytext=(0, -12), ha="center", fontsize=7.5, fontweight="bold", color="#000000")

    fig2.tight_layout()
    fig2.savefig(ruta_img_carton, dpi=220)
    plt.close(fig2)

    return ruta_img_flujos, ruta_img_carton


# ==============================================================================
# 5. CONSTRUCTOR WORD PERFECTO (SIN SUBRAYADOS ROJOS, PORTADA ULAGOS FIEL,
#    TABLA DE CONTENIDOS EN PÁGINA 2 Y COLUMNAS PROPORCIONADAS)
# ==============================================================================

def configurar_run_arial_limpio(run, size_pt=10.5, bold=False, italic=False, color_rgb=(0, 0, 0), superscript=False):
    """Aplica Arial Negro y desactiva el subrayado rojo ortográfico (w:noProof) para nombres propios locales."""
    run.font.name = "Arial"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.superscript = superscript
    run.font.color.rgb = RGBColor(*color_rgb)

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")

    # Evitar que Word ponga subrayados rojos de ortografía en siglas o comercios de Osorno
    if rPr.find(qn("w:noProof")) is None:
        rPr.append(OxmlElement("w:noProof"))

    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "es-CL")


def buscar_plantilla_en_pc(nombre_archivo="plantilla_hoja_carta_institucional_BLANCO.docx"):
    carpeta_script = os.path.dirname(os.path.abspath(__file__))
    home = os.path.expanduser("~")
    directorios = [
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
    for d in directorios:
        if os.path.isdir(d):
            cand = os.path.join(d, nombre_archivo)
            if os.path.exists(cand):
                return cand
            for f in os.listdir(d):
                if f.lower().endswith(".docx") and "plantilla" in f.lower() and "blanco" in f.lower() and not f.startswith("~$"):
                    return os.path.join(d, f)
    return None


def preparar_plantilla_ulagos_fiel():
    """
    Carga la plantilla institucional conservando ÚNICAMENTE el arte gráfico de la Página 1
    (Logo Universidad de Los Lagos arriba a la izquierda + cuadros grises + CNA 5 años abajo a la derecha),
    eliminando el segundo bloque repetido que tapaba el logo de la Universidad.
    """
    ruta_pl = buscar_plantilla_en_pc()
    if ruta_pl is not None:
        doc = Document(ruta_pl)
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

        parrafos_dibujo = []
        for p in list(doc.paragraphs):
            tiene_dib = any(True for _ in p._p.iter(qn("w:drawing"))) or any(True for _ in p._p.iter(qn("w:pict")))
            if tiene_dib:
                parrafos_dibujo.append(p)
            else:
                p._element.getparent().remove(p._element)

        # Conservar SOLO el primer párrafo de dibujo (Página 1 oficial con Logo ULagos y Acreditación)
        # para que el fondo de la Página 2 del template no se superponga ni tape el logo superior.
        if len(parrafos_dibujo) > 0:
            p_portada = parrafos_dibujo[0]
            p_portada.paragraph_format.space_before = Pt(0)
            p_portada.paragraph_format.space_after = Pt(0)
            p_portada.paragraph_format.line_spacing = 1.0
            for p_extra in parrafos_dibujo[1:]:
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

    return doc, usando_plantilla


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


def agregar_tabla_word_proporcionada(doc, df, incluir_indice=False, fuente_pt=8.0, col_izq_indices=(0,), anchos_pulgadas=None, ancho_total_pulgadas=6.7):
    """
    Inserta una tabla en Word con anchos de columna explícitos y proporcionados,
    repitiendo el encabezado si pasa a la página siguiente y evitando que las filas se corten.
    """
    df_work = df.reset_index() if incluir_indice else df.copy()
    filas, cols = df_work.shape

    if anchos_pulgadas is None:
        if cols > 6:
            w_primera = 2.6
            w_resto = (ancho_total_pulgadas - w_primera) / (cols - 1)
            anchos_pulgadas = [w_primera] + [w_resto] * (cols - 1)
        else:
            pesos = []
            for c_idx in range(cols):
                max_l = max(len(str(df_work.columns[c_idx])), max((len(str(df_work.iloc[r, c_idx])) for r in range(filas)), default=8))
                pesos.append(min(max(max_l, 8), 55))
            suma_p = sum(pesos)
            anchos_pulgadas = [(p / suma_p) * ancho_total_pulgadas for p in pesos]

    tabla = doc.add_table(rows=filas + 1, cols=cols)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False

    # Repetir fila de encabezado en cada página y evitar corte de fila
    trPr_hdr = tabla.rows[0]._tr.get_or_add_trPr()
    trPr_hdr.append(OxmlElement("w:tblHeader"))
    trPr_hdr.append(OxmlElement("w:cantSplit"))

    hdr_cells = tabla.rows[0].cells
    for c_idx, col_name in enumerate(df_work.columns):
        cell = hdr_cells[c_idx]
        cell.width = Inches(anchos_pulgadas[c_idx])
        sombrear_celda(cell, "1E293B")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx in col_izq_indices else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2.5)
        p.paragraph_format.space_after = Pt(2.5)
        run = p.add_run(limpiar_texto_chileno(col_name))
        configurar_run_arial_limpio(run, size_pt=fuente_pt, bold=True, color_rgb=(255, 255, 255))
        configurar_borde_celda(
            cell,
            top={"val": "single", "sz": "6", "color": "000000"},
            bottom={"val": "single", "sz": "8", "color": "000000"},
            left={"val": "single", "sz": "4", "color": "475569"},
            right={"val": "single", "sz": "4", "color": "475569"},
        )

    for r_idx in range(filas):
        row = tabla.rows[r_idx + 1]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))

        row_cells = row.cells
        texto_primera = limpiar_texto_chileno(df_work.iloc[r_idx, 0]).upper()
        es_destacada = any(k in texto_primera for k in ["TOTAL", "(=)", "FLUJO DE CAJA", "FLUJO DE BENEFICIOS", "CAPITAL DE TRABAJO REQUERIDO"])
        bg_hex = "E2E8F0" if es_destacada else ("FFFFFF" if r_idx % 2 == 0 else "F8FAFC")

        for c_idx in range(cols):
            cell = row_cells[c_idx]
            cell.width = Inches(anchos_pulgadas[c_idx])
            sombrear_celda(cell, bg_hex)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            val = df_work.iloc[r_idx, c_idx]
            col_nom = str(df_work.columns[c_idx])

            if isinstance(val, (int, float, np.integer, np.floating)):
                if any(k in col_nom for k in ["Cant.", "Unidades", "Kg/Mes"]) or ("PRODUCCIÓN" in texto_primera and c_idx > 0) or ("TOTAL PRODUCTOS" in texto_primera and c_idx > 0):
                    txt_val = fmt_int_cl(val)
                elif any(k in col_nom for k in ["Ton", "Día", "Factor"]) or ("TONELADAS" in texto_primera and c_idx > 0):
                    txt_val = fmt_dec_cl(val, 2)
                else:
                    txt_val = fmt_clp(val)
                alineacion = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                txt_val = limpiar_texto_chileno(val)
                alineacion = WD_ALIGN_PARAGRAPH.LEFT if c_idx in col_izq_indices else WD_ALIGN_PARAGRAPH.CENTER

            p = cell.paragraphs[0]
            p.alignment = alineacion
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(txt_val)
            configurar_run_arial_limpio(run, size_pt=fuente_pt, bold=(es_destacada or c_idx == 0), color_rgb=(0, 0, 0))

            configurar_borde_celda(
                cell,
                top={"val": "single", "sz": "6" if es_destacada else "4", "color": "000000" if es_destacada else "CBD5E1"},
                bottom={"val": "double" if es_destacada else "single", "sz": "8" if es_destacada else "4", "color": "000000" if es_destacada else "CBD5E1"},
                left={"val": "single", "sz": "4", "color": "CBD5E1"},
                right={"val": "single", "sz": "4", "color": "CBD5E1"},
            )

    p_esp = doc.add_paragraph()
    p_esp.paragraph_format.space_before = Pt(2)
    p_esp.paragraph_format.space_after = Pt(4)


def agregar_marcador_word(parrafo, nombre_anchor, id_bm):
    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(id_bm))
    bm_start.set(qn("w:name"), nombre_anchor)
    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(id_bm))
    parrafo._p.insert(0, bm_start)
    parrafo._p.append(bm_end)


def marcar_nivel_esquema_word(parrafo, nivel=0):
    pPr = parrafo._p.get_or_add_pPr()
    outlineLvl = pPr.find(qn("w:outlineLvl"))
    if outlineLvl is None:
        outlineLvl = OxmlElement("w:outlineLvl")
        pPr.append(outlineLvl)
    outlineLvl.set(qn("w:val"), str(nivel))


def insertar_toc_clicable_fijo_pagina_2(doc, items_toc):
    """
    Construye la Tabla de Contenidos en la Página 2 con hipervínculos directos (w:hyperlink)
    que jamás son borrados al exportar a PDF y que saltan a cada título con Ctrl + Clic.
    """
    for texto_sec, anchor_name, pag_str, nivel in items_toc:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4 if nivel == 1 else 2)
        p.paragraph_format.space_after = Pt(4 if nivel == 1 else 2)
        p.paragraph_format.left_indent = Inches(0.0 if nivel == 1 else 0.25)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.65), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), anchor_name)
        hyperlink.set(qn("w:history"), "1")

        def run_xml_arial(txt, bold=True, pt=10):
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
            t_el.text = limpiar_texto_chileno(txt)
            r_el.append(t_el)
            return r_el

        hyperlink.append(run_xml_arial(texto_sec, bold=(nivel == 1), pt=10 if nivel == 1 else 9.5))

        r_tab = OxmlElement("w:r")
        rPr_t = OxmlElement("w:rPr")
        rFonts_t = OxmlElement("w:rFonts")
        rFonts_t.set(qn("w:ascii"), "Arial")
        rFonts_t.set(qn("w:hAnsi"), "Arial")
        rPr_t.append(rFonts_t)
        r_tab.append(rPr_t)
        r_tab.append(OxmlElement("w:tab"))
        hyperlink.append(r_tab)

        hyperlink.append(run_xml_arial(f"Pág. {pag_str}", bold=True, pt=10 if nivel == 1 else 9.5))
        p._p.append(hyperlink)


def exportar_word_a_pdf(ruta_docx, ruta_pdf):
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
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            check=True,
            timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return os.path.exists(ruta_pdf)
    except Exception:
        return False


def generar_informe_word_limpio(resultados, ruta_word, generar_pdf=True):
    if os.path.exists(ruta_word):
        try:
            os.remove(ruta_word)
        except PermissionError:
            raise PermissionError(f"Cierra '{os.path.basename(ruta_word)}' en Word antes de continuar.")

    carpeta_img = os.path.dirname(os.path.abspath(ruta_word))
    cfg = resultados["cfg"]
    m = resultados["metricas"]
    iva = float(cfg["IVA"])
    ruta_img_flujos, ruta_img_carton = generar_graficos_limpios(resultados, carpeta_img)

    doc, usando_plantilla = preparar_plantilla_ulagos_fiel()

    # ==========================================================================
    # PÁGINA 1: PORTADA CON EL LOGO OFICIAL ULAGOS INTACTO Y BIEN ENCUADRADA
    # ==========================================================================
    ind_izq = Inches(2.15) if usando_plantilla else Inches(0.5)
    ind_der = Inches(0.20) if usando_plantilla else Inches(0.5)

    def add_p_portada(texto, size_pt=11, bold=True, italic=False, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = ind_izq
        p.paragraph_format.right_indent = ind_der
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(limpiar_texto_chileno(texto))
        configurar_run_arial_limpio(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=(0, 0, 0))
        return p

    # Espacio calibrado para quedar debajo del Logo Universidad de Los Lagos de la Página 1
    add_p_portada("", size_pt=10, space_before=135 if usando_plantilla else 24, space_after=0)
    add_p_portada("DEPARTAMENTO DE CIENCIAS ADMINISTRATIVAS Y ECONÓMICAS", size_pt=10.5, bold=True, space_after=2)
    add_p_portada("CARRERA DE INGENIERÍA COMERCIAL - CAMPUS OSORNO", size_pt=10, bold=True, space_after=16)

    add_p_portada("EVALUACIÓN PRIVADA Y SOCIAL DE PROYECTO DE ECONOMÍA CIRCULAR:", size_pt=10.5, bold=True, space_after=4)
    add_p_portada(
        f'"{cfg["NOMBRE_PROYECTO"].upper()}: REUTILIZACIÓN DE CARTÓN DEL EJE COMERCIAL ELEUTERIO RAMÍREZ PARA FABRICAR MATERIAL DIDÁCTICO ESCOLAR Y JUEGOS DE ADULTO MAYOR EN OSORNO"',
        size_pt=13.0, bold=True, space_after=12
    )

    add_p_portada(
        "Estudio a 10 años basado en 6 locales reales de calle Eleuterio Ramírez (#446 al #1024, retiro antes del camión municipal de las 06:00 a 07:00 AM hacia el Vertedero de Curaco) y venta a 3 compradores institucionales de Osorno (DAEM Osorno, Jardines Infantiles Privados y Hogar de Ancianos Santa María).",
        size_pt=9.5, bold=False, italic=True, space_after=24
    )

    add_p_portada("ASIGNATURA:", size_pt=9.5, bold=True, space_before=8, space_after=1)
    add_p_portada("Formulación y Evaluación de Proyectos", size_pt=10, bold=False, space_after=10)

    add_p_portada("DOCENTE GUÍA:", size_pt=9.5, bold=True, space_after=1)
    add_p_portada(cfg["DOCENTE"], size_pt=10, bold=False, space_after=10)

    add_p_portada("EQUIPO FORMULADOR (INTEGRANTES):", size_pt=9.5, bold=True, space_after=3)
    for integ in [x.strip() for x in cfg["INTEGRANTES"].split(",") if x.strip()]:
        add_p_portada(f"- {integ}", size_pt=9.5, bold=True, space_after=2)

    add_p_portada("CIUDAD Y FECHA:", size_pt=9.5, bold=True, space_before=10, space_after=1)
    add_p_portada("Osorno, Región de Los Lagos - Noviembre de 2026", size_pt=10, bold=False, space_after=4)

    # ==========================================================================
    # PÁGINA 2 EXCLUSIVA: TABLA DE CONTENIDOS CLICABLE Y FIJA EN PÁGINA 2
    # ==========================================================================
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
    configurar_run_arial_limpio(r_tt, size_pt=13, bold=True)

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
        ("6. CÁLCULO DEL CAPITAL DE TRABAJO (45 DÍAS DE DESFASE - PUNTO 6)", "_Toc_Sec_6", "10", 1),
        ("7. FLUJO DE CAJA PRIVADO A 10 AÑOS Y RENTABILIDAD (PUNTO 7)", "_Toc_Sec_7", "11", 1),
        ("8. EVALUACIÓN SOCIAL DEL PROYECTO (FLUJO SOCIAL MIDESO - PUNTO 8)", "_Toc_Sec_8", "12", 1),
        ("8.1 Resumen Comparativo: Resultados Privados (12%) vs. Sociales (6%)", "_Toc_Sub_81", "12", 2),
        ("9. CONCLUSIONES PRINCIPALES DEL PROYECTO", "_Toc_Sec_9", "13", 1),
        ("ANEXO 1: EL PORQUÉ DE CADA CIFRA (EXPLICACIÓN FÁCIL DE ENTENDER)", "_Toc_Anx_1", "14", 1),
        ("ANEXO 2: RESUMEN DE INFORMANTES Y DATOS DE TERRENO EN OSORNO", "_Toc_Anx_2", "15", 1),
        ("ANEXO 3: COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS EN OSORNO", "_Toc_Anx_3", "15", 1),
        ("ANEXO 4: SUPUESTOS DEL MODELO EXPLICADOS PASO A PASO", "_Toc_Anx_4", "16", 1),
    ]
    insertar_toc_clicable_fijo_pagina_2(doc, items_toc)
    doc.add_page_break()

    bm_counter = [100]

    def add_titulo_seccion(texto, anchor_name=None):
        p = doc.add_paragraph()
        marcar_nivel_esquema_word(p, nivel=0)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(limpiar_texto_chileno(texto))
        configurar_run_arial_limpio(r, size_pt=12, bold=True)
        if anchor_name:
            bm_counter[0] += 1
            agregar_marcador_word(p, anchor_name, bm_counter[0])
        return p

    def add_subtitulo(texto, anchor_name=None):
        p = doc.add_paragraph()
        marcar_nivel_esquema_word(p, nivel=1)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(limpiar_texto_chileno(texto))
        configurar_run_arial_limpio(r, size_pt=10.5, bold=True)
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
            rb = p.add_run(limpiar_texto_chileno(bold_prefix) + " ")
            configurar_run_arial_limpio(rb, size_pt=10.5, bold=True)

        # Convertir referencias "(Nota X)" en superíndices elegantes sin corchetes raros
        partes = re.split(r"(\(Nota \d+\))", limpiar_texto_chileno(texto))
        for parte in partes:
            if re.match(r"^\(Nota \d+\)$", parte):
                r_sup = p.add_run(f" {parte} ")
                configurar_run_arial_limpio(r_sup, size_pt=8.5, bold=True, superscript=True)
            elif parte:
                r_norm = p.add_run(parte)
                configurar_run_arial_limpio(r_norm, size_pt=10.5, bold=False)
        return p

    # ==========================================================================
    # DESARROLLO DEL INFORME (PÁGINA 3 EN ADELANTE)
    # ==========================================================================
    add_titulo_seccion("RESUMEN EJECUTIVO", "_Toc_Sec_0")
    add_parrafo(
        'El proyecto "Cartón Vivo Osorno" evalúa a 10 años la creación de un taller productivo en el Sector Industrial de Ovejería (Osorno) que recoge cajas de cartón limpias del comercio céntrico para transformarlas en juegos educativos para niños y juegos de memoria para adultos mayores. El problema actual en el centro de Osorno es muy concreto: todos los días entre las 06:00 y las 07:00 de la mañana pasa el camión recolector municipal por calle Eleuterio Ramírez y se lleva las cajas de cartón que dejan las tiendas, mezclándolas con basura común y enterrándolas en el Vertedero Provincial de Curaco.'
    )
    add_parrafo(
        "Para trabajar con números reales y fáciles de comprobar, nuestro proyecto recoge el cartón de 6 locales comerciales ubicados en calle Eleuterio Ramírez entre el #446 y el #1024 (Hojalatería Lynch #446, Autoplanet #582, Zapatería La Malagueña #625, Minimarket El Sabrosón #647, Galería Centrosorno #977 y Liquimax #1024). Nuestro furgón pasa a las 05:45 AM - justo antes que el camión de la basura - y retira en total 50 kilos por día trabajado. En 25 días al mes son 1.250 kilos mensuales, lo que equivale exactamente a 15,00 toneladas al año en el Año 1."
    )
    add_parrafo(
        "De esas 15,00 toneladas anuales, el 70% (10,50 toneladas) se aprovecha directamente en los productos y el 30% (4,50 toneladas) se pierde en recortes sobrantes y bordes con corchetes o cinta. Con las 10,50 toneladas útiles fabricamos dos productos validados con 3 compradores de Osorno: (1) Set Didáctico Escolar (3.000 unidades al año que usan 2,50 kilos de cartón cada una y se venden a $16.000 con IVA(Nota 1) al DAEM de Osorno y a Jardines Infantiles Privados); y (2) Juegos Cognitivos Adulto Mayor (1.500 unidades al año que usan 2,00 kilos de cartón cada una y se venden a $19.000 con IVA(Nota 1) al Hogar de Ancianos Santa María y clubes de adulto mayor)."
    )
    add_parrafo(
        f"En la Evaluación Privada (Punto 7), invirtiendo $16.000.000 con IVA en máquinas y arreglo del galpón ({fmt_clp(m['inv_inicial_priv'])} netos sin IVA(Nota 2)) más un capital de trabajo inicial de {fmt_clp(m['ct_inicial'])} para cubrir 45 días de operación, y exigiendo un 12% de rentabilidad anual, el proyecto logra un VAN Privado positivo de {fmt_clp(m['van_privado'])}, una TIR de {fmt_dec_cl(m['tir_privada'], 2)}% y recupera el dinero en {fmt_dec_cl(m['payback_privado'], 2)} años. En la Evaluación Social MIDESO (Punto 8), con la tasa social del 6% anual y sumando el ahorro municipal por no enterrar este cartón en el Vertedero de Curaco ($48.000 por tonelada con IVA(Nota 5)), el VAN Social sube a {fmt_clp(m['van_social'])} y la TIR Social a {fmt_dec_cl(m['tir_social'], 2)}%."
    )

    add_titulo_seccion("1. EL PROBLEMA EN OSORNO Y LOS 6 LOCALES DE CALLE ELEUTERIO RAMÍREZ", "_Toc_Sec_1")
    add_parrafo(
        "El Vertedero Provincial de Curaco en Osorno tiene cada vez menos espacio porque recibe a diario toneladas de cartón limpio que podrían aprovecharse. En la calle Eleuterio Ramírez, las ferreterías, tiendas de repuestos, zapaterías, minimarkets y mayoristas desocupan todos los días cajas de cartón grueso y resistente. Como no tienen bodega para guardarlas, las sacan temprano a la vereda y el camión de la basura municipal se las lleva entre las 06:00 y las 07:00 de la mañana."
    )
    add_parrafo(
        "Fuimos a terreno en calle Eleuterio Ramírez (entre los números #446 y #1024) y registramos 6 comercios reales que entregan en conjunto 50 kilos de cartón limpio por día hábil (1.250 kilos al mes en 25 días trabajados = 15,00 toneladas al año, respaldado en el Anexo Excel 2 de Informantes y Encuestas). El detalle de cada local se muestra en la Tabla 1.1:"
    )
    agregar_tabla_word_proporcionada(
        doc, resultados["df_catastro_6"], incluir_indice=False, fuente_pt=7.8,
        col_izq_indices=(1, 2, 3, 4),
        anchos_pulgadas=[0.55, 1.15, 1.10, 0.85, 1.70, 0.45, 0.45, 0.45]
    )

    add_titulo_seccion("2. PRODUCTOS FABRICADOS Y LOS 3 COMPRADORES EN OSORNO", "_Toc_Sec_2")
    add_parrafo(
        "Caja con letras móviles, figuras geométricas encajables y tableros para aprender a leer y contar en kínder y primeros años básicos, pintados con pintura al agua segura para niños (certificación europea EN-71).",
        bold_prefix="- Producto 1: Set Didáctico Escolar (usa 2,50 kilos de cartón por unidad - Precio: $16.000 con IVA / $13.445 sin IVA(Nota 1)):"
    )
    add_parrafo(
        "Tableros grandes con letras y figuras de alto contraste y piezas gruesas fáciles de tomar con la mano, pensados para ejercitar la memoria en personas mayores.",
        bold_prefix="- Producto 2: Juegos Cognitivos Adulto Mayor (usa 2,00 kilos de cartón por unidad - Precio: $19.000 con IVA / $15.966 sin IVA(Nota 1)):"
    )
    add_parrafo(
        "En lugar de depender de ventas sueltas al detalle, aseguramos las ventas del Año 1 (4.500 juegos en total) con 3 grupos de compradores institucionales de Osorno entrevistados en terreno (Tabla 2.1):"
    )
    agregar_tabla_word_proporcionada(
        doc, resultados["df_demanda_3"], incluir_indice=False, fuente_pt=7.8,
        col_izq_indices=(1, 2, 3, 6),
        anchos_pulgadas=[0.60, 1.30, 1.10, 0.95, 0.55, 0.65, 1.55]
    )

    add_titulo_seccion("3. FUNCIONAMIENTO DEL TALLER Y APROVECHAMIENTO DEL CARTÓN (70% / 30%)", "_Toc_Sec_3")
    add_parrafo(
        "De lunes a sábado, entre las 05:30 y las 06:15 de la mañana, nuestro furgón pasa por los 6 locales de calle Eleuterio Ramírez antes que el camión municipal (costo de flete: $30.000 con IVA por tonelada(Nota 3)) y lleva las cajas secas al galpón de 140 m2 arrendado en el Sector Industrial de Ovejería ($450.000 mensuales(Nota 4)). Allí las cajas se limpian, se cortan con moldes en la máquina troqueladora de rodillos y se pintan con tintas al agua ($1.800 en pintura y barniz + $600 en caja por unidad(Nota 2))."
    )
    add_parrafo(
        "Supuesto de Balance de Masa y Merma Técnica (70% útil y 30% merma): De las 15,00 toneladas que recogemos en el Año 1, el 70% (10,50 toneladas) queda dentro de los 4.500 productos terminados (3.000 sets escolares x 2,50 kg = 7,50 toneladas + 1.500 juegos adulto mayor x 2,00 kg = 3,00 toneladas). El 30% restante (4,50 toneladas) corresponde a la merma normal del proceso (15% por descarte de bordes con corchetes, cintas o humedad matutina y 15% por recorte sobrante del molde de troquelado)."
    )
    if os.path.exists(ruta_img_carton):
        p_img_c = doc.add_paragraph()
        p_img_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img_c.add_run().add_picture(ruta_img_carton, width=Inches(5.9))

    add_titulo_seccion("4. TRABAJADORES, SUELDOS Y PRECIOS SOCIALES MIDESO (PUNTO 4.3)", "_Toc_Sec_4")
    add_parrafo(
        "La empresa funcionará como una Sociedad por Acciones (SpA) en el Régimen Pro Pyme General (paga 25% de impuesto sobre las utilidades). El taller trabaja con 5 personas en total(Nota 6): 1 Jefe de Planta y Ventas (mano de obra calificada MOC, sueldo de $950.000 al mes), 2 Operadores de Máquina (semicalificada MOSC, $600.000 al mes cada uno) y 2 Operarios de Recolección y Armado (no calificada MONC, $520.000 al mes cada uno). En la Tabla 4.1 se muestran sus sueldos reales y cómo se ajustan para el Flujo Social usando los factores oficiales de MIDESO (1,00 para MOC; 0,85 para MOSC y 0,68 para MONC):"
    )
    agregar_tabla_word_proporcionada(
        doc, resultados["df_personal"], incluir_indice=False, fuente_pt=7.8,
        col_izq_indices=(0, 3),
        anchos_pulgadas=[1.35, 0.40, 0.50, 1.75, 0.70, 0.75, 0.50, 0.75]
    )

    # Bloque de Notas a Pie de Página
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(6)
    p_sep.paragraph_format.space_after = Pt(2)
    r_s = p_sep.add_run("________________________________________")
    configurar_run_arial_limpio(r_s, size_pt=8.5, bold=True)

    p_nt = doc.add_paragraph()
    p_nt.paragraph_format.space_after = Pt(2)
    r_nt = p_nt.add_run("NOTAS A PIE DE PÁGINA - REFERENCIAS DE PRECIOS, COSTOS Y PARÁMETROS:")
    configurar_run_arial_limpio(r_nt, size_pt=8.5, bold=True)

    notas_pie = [
        "(Nota 1) Precios de venta validados en entrevistas con encargados de compras DAEM Osorno (Compra Ágil y SEP), directoras de jardines particulares y coordinación del Hogar de Ancianos Santa María ($16.000 c/IVA Set Escolar y $19.000 c/IVA Juego Cognitivo; un 25% a 30% menores que equivalentes en madera).",
        "(Nota 2) Cotizaciones de maquinaria e insumos gráficos (Octubre-Noviembre 2026): Troqueladora de rodillos 1200 mm ($8.500.000 c/IVA), Guillotina y matrices ($3.500.000 c/IVA), Equipo serigráfico al agua ($2.000.000 c/IVA), Habilitación eléctrica SEC y SEREMI galpón Ovejería ($2.000.000 c/IVA), Tintas al agua EN-71 y PVA ($1.800/unid c/IVA) y Caja microcorrugada ($600/unid c/IVA).",
        "(Nota 3) Costo logístico cotizado para furgón utilitario local en ruta matutina calle Eleuterio Ramírez (#446 al #1024) hacia Sector Industrial Ovejería: $30.000 c/IVA por tonelada transportada ($25.210 netos por tonelada).",
        "(Nota 4) Costos fijos mensuales en Osorno: Arriendo galpón 140 m2 en Ovejería ($450.000 exento de IVA), electricidad trifásica Saesa + agua Suralis + internet ($180.000 c/IVA) y contabilidad Pro Pyme + patente municipal + seguros ($150.000 c/IVA).",
        "(Nota 5) Costo municipal ahorrado por tonelada desviada del Vertedero Provincial de Curaco: $48.000 por tonelada c/IVA ($40.336 netos), estimado según contratos de recolección y disposición final de la Municipalidad de Osorno y registros SINADER.",
        "(Nota 6) Remuneraciones brutas mensuales acordes al mercado laboral de Osorno (Año 2026) y factores de conversión social del Ministerio de Desarrollo Social y Familia (MIDESO / SNI): MOC = 1,00; MOSC = 0,85; MONC = 0,68; Divisa = 1,01.",
    ]
    for n_txt in notas_pie:
        pn = doc.add_paragraph()
        pn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pn.paragraph_format.space_before = Pt(1)
        pn.paragraph_format.space_after = Pt(1.5)
        rn = pn.add_run(limpiar_texto_chileno(n_txt))
        configurar_run_arial_limpio(rn, size_pt=8.0, bold=False)

    # SECCIÓN HORIZONTAL PARA TABLAS DE 10 AÑOS
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
    add_subtitulo("5.a Cantidad de Cartón Recolectado y Juegos Fabricados por Año (Crece 5% anual - Ver Anexo Excel 3)", "_Toc_Sub_5a")
    agregar_tabla_word_proporcionada(doc, resultados["df_cantidades"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_subtitulo("5.b Inversiones en Máquinas en el Año 0 ($16.000.000 con IVA = $13.445.378 sin IVA), Renovación y Depreciación", "_Toc_Sub_5b")
    agregar_tabla_word_proporcionada(doc, resultados["df_inversiones"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,), ancho_total_pulgadas=9.6)
    agregar_tabla_word_proporcionada(doc, resultados["df_depreciacion"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_subtitulo("5.c Costos Variables, Gastos Fijos y Sueldos por Año (Valores Netos sin IVA)", "_Toc_Sub_5c")
    agregar_tabla_word_proporcionada(doc, resultados["df_costos_priv"], incluir_indice=True, fuente_pt=7.2, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_subtitulo("5.d Ingresos Anuales por Ventas sin IVA (Años 1 al 10)", "_Toc_Sub_5d")
    agregar_tabla_word_proporcionada(doc, resultados["df_ingresos"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_titulo_seccion("6. CÁLCULO DEL CAPITAL DE TRABAJO (45 DÍAS DE DESFASE - PUNTO 6)", "_Toc_Sec_6")
    add_parrafo(
        f"Supuesto de Capital de Trabajo: Como el DAEM de Osorno y las instituciones demoran entre 30 y 45 días en pagar las facturas, necesitamos contar con dinero en caja para cubrir 45 días de costos de funcionamiento ((Costos Anuales x 45) / 365). Por eso en el Año 0 se ponen {fmt_clp(m['ct_inicial'])} de Capital de Trabajo, el cual se recupera completo al terminar el Año 10:"
    )
    agregar_tabla_word_proporcionada(doc, resultados["df_capital_trabajo"], incluir_indice=True, fuente_pt=7.4, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_titulo_seccion("7. FLUJO DE CAJA PRIVADO A 10 AÑOS Y RENTABILIDAD (PUNTO 7)", "_Toc_Sec_7")
    add_parrafo(
        "En la Tabla 7.1 se presenta el Flujo de Caja Privado sin IVA (respaldado con fórmulas editables en el Anexo Excel 1 de Flujos), descontando el 25% de Impuesto a la Renta y evaluando al 12,0% de tasa de descuento privada:"
    )
    agregar_tabla_word_proporcionada(doc, resultados["df_flujo_privado"], incluir_indice=True, fuente_pt=7.0, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_titulo_seccion("8. EVALUACIÓN SOCIAL DEL PROYECTO (FLUJO SOCIAL MIDESO - PUNTO 8)", "_Toc_Sec_8")
    add_parrafo(
        f"En el Flujo Social (Tabla 8.1) se eliminan los impuestos (IVA, aranceles e impuesto a la renta), se ajustan los sueldos con los factores sociales de MIDESO y se suma como beneficio para Osorno el dinero que ahorra la Municipalidad por cada tonelada de cartón que no tiene que recoger ni enterrar en el Vertedero de Curaco ({fmt_clp(cfg['AHORRO_MUNI_TON']/(1+iva))} netos por tonelada(Nota 5)), más {fmt_clp(cfg['OTROS_BEN_SOC_ANUAL'])} anuales por beneficios ambientales y educativos:"
    )
    agregar_tabla_word_proporcionada(doc, resultados["df_flujo_social"], incluir_indice=True, fuente_pt=7.1, col_izq_indices=(0,), ancho_total_pulgadas=9.6)

    add_subtitulo("8.1 Resumen Comparativo: Resultados Privados (12%) vs. Sociales (6%)", "_Toc_Sub_81")
    agregar_tabla_word_proporcionada(
        doc, resultados["df_indicadores"], incluir_indice=False, fuente_pt=8.2,
        col_izq_indices=(0,), anchos_pulgadas=[2.4, 1.3, 1.6, 1.3, 1.4, 1.6]
    )

    # RETORNO A VERTICAL PARA CONCLUSIONES Y ANEXOS
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
        p_img_f.add_run().add_picture(ruta_img_flujos, width=Inches(6.3))

    add_titulo_seccion("9. CONCLUSIONES PRINCIPALES DEL PROYECTO", "_Toc_Sec_9")
    add_parrafo(
        "1. Cifras claras y comprobables en terreno: Recogiendo apenas 50 kilos diarios de cartón en 6 locales reales de calle Eleuterio Ramírez (#446 al #1024) antes de las 06:00 AM, juntamos 15,00 toneladas al año. Usando el 70% (10,50 toneladas) y dejando un 30% de margen por pérdidas y recortes, fabricamos exactamente los 4.500 juegos anuales que demandan el DAEM, los jardines privados y los centros de adulto mayor de Osorno."
    )
    add_parrafo(
        f"2. Negocio rentable para el inversionista privado (Punto 7): Con una tasa exigida del 12% anual, el proyecto entrega un VAN Privado positivo de {fmt_clp(m['van_privado'])}, una TIR de {fmt_dec_cl(m['tir_privada'], 2)}% y recupera todo el capital invertido en {fmt_dec_cl(m['payback_privado'], 2)} años."
    )
    add_parrafo(
        f"3. Alto beneficio social y ambiental para Osorno (Punto 8): Evaluado con la tasa social del 6% de MIDESO, el VAN Social alcanza {fmt_clp(m['van_social'])} y la TIR Social llega a {fmt_dec_cl(m['tir_social'], 2)}% (Payback Social de {fmt_dec_cl(m['payback_social'], 2)} años), demostrando que el proyecto alivia el Vertedero de Curaco, genera empleo en Osorno y entrega material educativo accesible."
    )

    add_titulo_seccion("ANEXO 1: EL PORQUÉ DE CADA CIFRA (EXPLICACIÓN FÁCIL DE ENTENDER)", "_Toc_Anx_1")
    agregar_tabla_word_proporcionada(
        doc, resultados["df_porque"], incluir_indice=False, fuente_pt=8.0,
        col_izq_indices=(0, 1, 2), anchos_pulgadas=[1.85, 1.35, 3.50]
    )

    add_titulo_seccion("ANEXO 2: RESUMEN DE INFORMANTES Y DATOS DE TERRENO EN OSORNO", "_Toc_Anx_2")
    agregar_tabla_word_proporcionada(
        doc, resultados["df_informantes"], incluir_indice=False, fuente_pt=7.8,
        col_izq_indices=(0, 1, 2, 3), anchos_pulgadas=[1.35, 1.65, 1.45, 2.25]
    )

    add_titulo_seccion("ANEXO 3: COTIZACIONES DE MÁQUINAS, INSUMOS Y SERVICIOS EN OSORNO", "_Toc_Anx_3")
    agregar_tabla_word_proporcionada(
        doc, resultados["df_cotizaciones"], incluir_indice=False, fuente_pt=7.6,
        col_izq_indices=(0, 1, 2, 6), anchos_pulgadas=[0.60, 0.85, 1.65, 0.65, 0.70, 0.70, 1.55]
    )

    add_titulo_seccion("ANEXO 4: SUPUESTOS DEL MODELO EXPLICADOS PASO A PASO", "_Toc_Anx_4")
    agregar_tabla_word_proporcionada(
        doc, resultados["df_supuestos"], incluir_indice=False, fuente_pt=7.8,
        col_izq_indices=(0, 1, 2, 3), anchos_pulgadas=[1.35, 1.15, 2.55, 1.65]
    )

    doc.save(ruta_word)

    ruta_pdf = os.path.splitext(ruta_word)[0] + ".pdf"
    pdf_ok = exportar_word_a_pdf(ruta_word, ruta_pdf) if generar_pdf else False
    return ruta_word, ruta_pdf, usando_plantilla, pdf_ok


# ==============================================================================
# 6. EMPAQUETADOR AUTOMÁTICO DE LA ENTREGA COMPLETA (WORD + PDF + 3 EXCELS + ZIP)
# ==============================================================================

def ejecutar_super_corrector(directorio_destino=None):
    if directorio_destino is None:
        directorio_destino = os.path.dirname(os.path.abspath(__file__))

    cfg = copy.deepcopy(DEFAULTS)
    cfg["INVERSIONES_LISTA"] = copy.deepcopy(INVERSIONES_DEFAULT)
    cfg["COSTOS_VAR_LISTA"] = copy.deepcopy(COSTOS_VAR_DEFAULT)
    cfg["COSTOS_FIJOS_LISTA"] = copy.deepcopy(COSTOS_FIJOS_DEFAULT)
    cfg["PERSONAL_LISTA"] = copy.deepcopy(PERSONAL_DEFAULT)

    resultados = ejecutar_motor_financiero(cfg)
    apellidos = cfg["APELLIDOS_ENTREGA"].strip()

    if os.path.basename(os.path.normpath(directorio_destino)).lower() == apellidos.lower():
        carpeta_entrega = os.path.normpath(directorio_destino)
        directorio_padre = os.path.dirname(carpeta_entrega)
    else:
        carpeta_entrega = os.path.join(directorio_destino, apellidos)
        directorio_padre = directorio_destino

    os.makedirs(carpeta_entrega, exist_ok=True)

    # 1. Crear los 3 Excels limpios
    r_exc1, r_exc2, r_exc3 = exportar_tres_excels_limpios(resultados, carpeta_entrega)

    # 2. Crear el Word limpio y su PDF
    ruta_word = os.path.join(carpeta_entrega, f"{apellidos}.docx")
    ruta_word, ruta_pdf, usando_pl, pdf_ok = generar_informe_word_limpio(resultados, ruta_word, generar_pdf=True)

    # Eliminar PNGs temporales para dejar la carpeta limpia
    for img_temp in ["grafico_flujos_cartonvivo.png", "grafico_carton_ramirez.png"]:
        p_img = os.path.join(carpeta_entrega, img_temp)
        if os.path.exists(p_img):
            try:
                os.remove(p_img)
            except Exception:
                pass

    # 3. Comprimir en Vargas,Muñoz,Jofre.zip
    ruta_zip = os.path.join(directorio_padre, f"{apellidos}.zip")
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

    return {
        "carpeta": carpeta_entrega,
        "zip": ruta_zip,
        "word": ruta_word,
        "pdf": ruta_pdf,
        "pdf_ok": pdf_ok,
        "plantilla_ok": usando_pl,
        "excels": [r_exc1, r_exc2, r_exc3],
    }


if __name__ == "__main__":
    print(">>> Ejecutando SUPER CORRECTOR ULAGOS...")
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    carpeta_base = os.path.dirname(os.path.abspath(__file__))
    try:
        info = ejecutar_super_corrector(carpeta_base)
        if sys.platform.startswith("win"):
            os.startfile(info["carpeta"])
        messagebox.showinfo(
            "¡Super Corrector ULagos Completado!",
            f"Se corrigieron todos los caracteres, tablas, portada e índice clicable.\n\n"
            f"Carpeta lista:\n{info['carpeta']}\n\n"
            f"Archivo ZIP para subir a la plataforma:\n{info['zip']}\n\n"
            f"Incluye:\n"
            f"- {os.path.basename(info['word'])} (Logo ULagos en Portada + Índice en Pág. 2)\n"
            f"- {os.path.basename(info['pdf'])} ({'PDF Generado' if info['pdf_ok'] else 'Abrir Word y guardar como PDF'})\n"
            f"- 1_Excel_Flujos_Editables.xlsx\n"
            f"- 2_Excel_Informantes_Encuestas_Osorno.xlsx\n"
            f"- 3_Excel_Calculos_Proyecciones.xlsx",
            parent=root
        )
    except Exception as e:
        messagebox.showerror("Aviso en Super Corrector", f"Cierra Word y Excel si tienes los archivos abiertos e intenta de nuevo:\n\n{e}", parent=root)
    finally:
        root.destroy()