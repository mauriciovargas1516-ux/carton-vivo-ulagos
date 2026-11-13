#  Suite de Formulación y Evaluación de Proyectos 2026

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/)
[![Lanzar Suite en Streamlit](https://img.shields.io/badge/_ABRIR_SUITE_INTERACTIVA_EN_STREAMLIT-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://github.com/mauriciovargas1516-ux/PROYECTOS-2026)

 **Acceso directo a la consola web en vivo:**  
**[https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/](https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/)**

## 1. Propósito del Sistema

Este repositorio integra un motor matematico financiero y socioeconomico con un sistema de automatizacion documental. Permite evaluar iniciativas de inversion en tiempo real bajo criterios privados y bajo la normativa vigente del Sistema Nacional de Inversiones (SNI) del Ministerio de Desarrollo Social y Familia (MIDESO) de Chile.

A partir de los supuestos ingresados en la interfaz web, el sistema recalcula instantaneamente los indicadores de decision y compila un paquete descargable en formato `.zip` que contiene el informe tecnico en Word, su version en PDF y tres libros de calculo en Excel con formulas dinamicas vinculadas.

---

## 2. Capacidades Principales

### Evaluacion Privada y Simulador Financiero
* **Flujo de Caja Puro y Financiado:** Proyeccion a 10 o 20 anos, permitiendo alternar entre financiamiento 100% con capital propio o endeudamiento bancario (calculo automatico de tabla de amortizacion metodo frances, intereses y escudo tributario).
* **Tributacion SII:** Soporte para tasa fija de Primera Categoria o secuencia escalonada transitoria del Regimen Pro Pyme General (Circular SII N° 53/2025: 12,5%, 15% y 25%).
* **Capital de Trabajo:** Estimacion automatica mediante el metodo del periodo de desfase (ciclo de cobro institucional y comercial).
* **Indicadores de Rentabilidad y Riesgo:** Valor Actual Neto (VAN), Tasa Interna de Retorno (TIR), Periodo de Recuperacion (Payback) y analisis de sensibilidad multiescenario (pesimista en ingresos, incremento de costos y crecimiento cero).

### Evaluacion Social y Ambiental (Normativa SNI / MIDESO 2026)
* **Precios Sociales Vigentes 2026:** Aplicacion directa de la Tasa Social de Descuento (TSD 5,5%), factores de correccion social de la mano de obra (Calificada 0,97; Semicalificada 0,95; No Calificada 0,91) y factor social de la divisa (1,00).
* **Enfoque Costo-Beneficio y Costo-Eficiencia:** Calculo simultaneo de Valor Actual Neto Social (VANS), Tasa Interna de Retorno Social (TIRS), Valor Actual de Costos (VAC) y Costo Anual Equivalente (CAE total y por tonelada o unidad gestionada).
* **Economia Circular y Ley REP (Ley N° 20.920):** Modelamiento de balance fisico de masa (fraccion valorizada incorporada en productos versus fraccion residual derivada a gestores autorizados) y cuantificacion de liberacion de recursos municipales por ahorro en recoleccion y disposicion final en vertedero.
* **Valorizacion de Gases de Efecto Invernadero (GEI):** Cuantificacion del beneficio climatico por reduccion de emisiones aplicando el Precio Social del Carbono 2026 ($71.801 CLP / t CO2eq) y factores de conversion de metano.
* **Ficha Resumen IDI / BIP:** Conversion automatica de presupuestos, costos e indicadores a Unidades de Fomento (UF oficial Precios Sociales 2026).

### Microeconomia Aplicada (Valor Social de la Produccion)
* **Equilibrio de Mercado sin y con Proyecto:** Resolucion analitica y grafica de funciones lineales de oferta y demanda ante la entrada de un proyecto productivo.
* **Enfoque de Eficiencia y por Agentes:** Desglose exacto del Valor Social de la Produccion (VSP) en Aumento en Disponibilidad (AD) y Liberacion de Recursos (LR), junto con la variacion en el Excedente del Consumidor y Excedente del Productor.

---

## 3. Automatizacion de Entregables (.ZIP)

Al ejecutar la exportacion desde la consola web, el sistema genera y empaqueta automaticamente los siguientes archivos:

1. **Informe Tecnico Completo (`.docx` y `.pdf`):**
   * Aplicacion automatica de la plantilla institucional oficial de la Universidad de Los Lagos.
   * Tabla de Contenidos dinamica con hipervinculos y numeracion de pagina actualizada.
   * Siete diagramas tecnicos generados en alta resolucion: Arbol de Problemas, Arbol de Objetivos, Balance Fisico de Masa, Croquis de Planta, Flujo de Proceso Productivo, Comparacion Anual de Flujos y Analisis de Sensibilidad.
   * Modelo de Gestion Estrategica (Business Model Canvas), Matriz de Marco Logico (Fin, Proposito, Componentes y Actividades) y Bibliografia en formato APA 7ª edicion.
2. **Libro 1 — `1_Excel_Flujos_Editables.xlsx`:**
   * Resumen de indicadores, Flujo de Caja Privado, Flujo de Beneficios Sociales Netos y Capital de Trabajo con formulas nativas de Excel (`NPV`, `IRR`, `SUM`, `ROUND`).
3. **Libro 2 — `2_Excel_Informantes_Encuestas_Osorno.xlsx`:**
   * Catastro de proveedores de residuo en terreno, validacion de demanda institucional, matriz de informantes, cotizaciones de mercado y memoria explicativa de supuestos.
4. **Libro 3 — `3_Excel_Calculos_Proyecciones.xlsx`:**
   * Proyecciones fisicas a 10 anos, calendario de inversiones, reinversiones, depreciacion lineal, estructura salarial a precios privados y sombra, e ingresos operacionales.

---

## 4. Estructura del Repositorio

| Archivo | Descripcion Tecnica |
| :--- | :--- |
| `consola_web.py` | Aplicacion principal en Streamlit. Integra los paneles de simulacion, calculadoras SNI 2026, formateo APA 7 y descarga del paquete `.zip`. |
| `2_corrector_formato_y_excels.py` | Motor financiero y socioeconomico en Python. Construye las matrices de flujos de caja, exporta los tres libros Excel y estructura el documento Word base. |
| `3_integrador_arboles_y_pauta_josue.py` | Motor grafico en Matplotlib y ensamblador XML (`python-docx`) para arboles causales, croquis de planta, Canvas, Marco Logico e indice clicable. |
| `1_simulador_interactivo_cartonvivo.py` | Version de escritorio con interfaz grafica local (`Tkinter`) y motor de respaldo. |
| `0_consola_maestra_total.py` | Modulo de gestion de filas dinamicas para inversiones, costos operativos y dotacion de personal. |
| `plantilla_hoja_carta_institucional_BLANCO.docx` | Plantilla institucional con diseno oficial de portada, encabezados y pies de pagina de la Universidad de Los Lagos. |
| `requirements.txt` | Dependencias de Python necesarias para el despliegue en Streamlit Cloud y ejecucion local. |
| `packages.txt` | Paquetes del sistema Linux (`libreoffice`) para la conversion automatizada de Word a PDF en el servidor. |

---

## 5. Caso de Estudio Base Incluido: Carton Vivo Osorno

El sistema incluye preconfigurado el proyecto **"Carton Vivo Osorno"**, una iniciativa de economia circular que recolecta 15,00 toneladas anuales de carton corrugado limpio en el eje comercial de calle Eleuterio Ramirez (antes del paso del camion recolector municipal hacia el Vertedero Provincial de Curaco) para transformarlo en una planta de 140 m2 en el sector industrial de Ovejeria en dos lineas de productos:
* **Set Didactico Escolar:** Orientado a escuelas municipales, jardines VTF (DAEM Osorno) y jardines infantiles particulares.
* **Juegos Cognitivos para Adulto Mayor:** Orientado a residencias y clubes de adulto mayor de la comuna.

### Resumen de Resultados del Caso Base (Parametros SNI 2026)

| Indicador | Evaluacion Privada (Tasa 12,0%) | Evaluacion Social SNI 2026 (TSD 5,5%) |
| :--- | :--- | :--- |
| **Inversion Inicial Activos Fijos (Ano 0)** | $13.445.377 (neto sin IVA) | $13.029.173 (precios sociales) |
| **Capital de Trabajo Inicial (45 dias)** | $6.961.008 | $6.202.345 |
| **VAN / VANS** | $61.910.740 (25% fijo) / $63.175.846 (Pro Pyme) | $154.991.495 |
| **TIR / TIRS** | 48,75% / 50,66% | 70,26% |
| **Payback (Recuperacion)** | 2,56 anos / 2,40 anos | 1,74 anos |

---

## 6. Instalacion y Ejecucion Local

1. Clonar el repositorio:
```bash
git clone https://github.com/mauriciovargas1516-ux/PROYECTOS-2026.git
cd PROYECTOS-2026
pip install -r requirements.txt
streamlit run consola_web.py
