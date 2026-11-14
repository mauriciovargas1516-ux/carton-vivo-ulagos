# Suite de Formulación y Evaluación Privada y Social de Proyectos 2026 — ULagos

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/)
[![Lanzar Consola en Streamlit](https://img.shields.io/badge/ABRIR_CONSOLA_DE_PROYECTOS_EN_STREAMLIT-1E3A8A?style=for-the-badge&logo=streamlit&logoColor=white)](https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://github.com/mauriciovargas1516-ux/PROYECTOS-2026)

Plataforma interactiva de ingeniería económica y automatización documental desarrollada en **Python (`Streamlit`, `python-docx`, `openpyxl`, `numpy-financial`, `matplotlib`)** para la formulación, evaluación financiera privada y evaluación social de proyectos bajo metodología del **Sistema Nacional de Inversiones (SNI / MIDESO 2026)**, aplicada al proyecto de economía circular **"Cartón Vivo Osorno"** (Ingeniería Comercial, Universidad de Los Lagos — Docente: Dra. Luz María Ferrada B.).

**Acceso directo a la consola interactiva en vivo:**  
https://proyectos-2026-ubfvxgpqmpp5z7dn6wlnn3.streamlit.app/

---

## Guía Rápida de Uso (Simulación en Vivo y Generación de Entregables en 1 Clic)

1. **Ajustar Supuestos o Cargar Presets Oficiales:** Usa los botones superiores para alternar entre los **Precios Sombra SNI 2026 (TSD 5,5%)** y el **Caso Base Original (TSD 6,0%)**, o modifica libremente tasas, demanda, precios y factores sociales (MOC, MOSC, MONC, Divisa).
2. **Agregar o Quitar Ítems Dinámicamente (`➕` / `🗑️`):** Puedes añadir o eliminar máquinas en el CAPEX, costos variables, gastos fijos del taller o cargos de personal; el banner superior recalculará en tiempo real el **VAN Privado, TIR Privada, VANS Social, TIRS Social y Payback**.
3. **Generar Paquete Completo (`.ZIP`):** Al presionar **"Generar Entrega Completa"**, el servidor ejecuta en cadena los motores financieros y documentales para empaquetar y descargar el archivo `Vargas,Muñoz,Jofre.zip` con el **Informe Word (`.docx`)**, el **Informe `.pdf`** y los **3 Anexos Excel (`.xlsx`)** actualizados al peso con tus cifras en pantalla.

---

## Impacto Cuantitativo y Multiplicadores de Velocidad (Pipeline Automatizado vs. Trabajo Manual)

| Etapa de Formulación, Modelamiento o Maquetación | Entregable / Volumen Procesado | Tiempo Manual (Word / Excel) | Tiempo Motor Cloud | Horas Ahorradas | % Mejora (Multiplicador) |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **1. Construcción de Flujo Privado (12%) y Flujo Social SNI (6% / 5,5%) a 10 Años** | Corrección de IVA, Aranceles, Precios Sombra (MOC/MOSC/MONC/Divisa) y Capital de Trabajo | 12,0 horas | **0,01 hrs** (2 seg) | 11,99 hrs/ciclo | **99,9% (21.600x más rápido)** |
| **2. Compilación de 3 Libros Excel Anexos con Fórmulas Nativas (`openpyxl`)** | Anexo 1 (Flujos), Anexo 2 (Catastro e Informantes) y Anexo 3 (Proyecciones y Sensibilidad) | 9,5 horas | **0,02 hrs** (3 seg) | 9,48 hrs/entrega | **99,8% (11.400x más rápido)** |
| **3. Redacción, Cuadratura de Tablas y Formato Institucional en Word (`python-docx`)** | Portada oficial ULagos, Tabla de Contenido con puntos líderes, 12 capítulos, Canvas y Marco Lógico | 18,0 horas | **0,03 hrs** (4 seg) | 17,97 hrs/informe | **99,8% (16.200x más rápido)** |
| **4. Recálculo Total ante Cambio de un Supuesto (Ej. Sueldo, Precio o TSD)** | Actualizar 3 Excels + más de 15 tablas del Word + gráficos + conclusiones sin descuadres | 4,5 horas (por cambio) | **0,02 hrs** (8 seg) | 4,48 hrs/iteración | **99,6% (2.025x más rápido)** |
| **TOTAL GENERACIÓN DE ENTREGA OFICIAL COMPLETA (`.ZIP`)** | **1 Word + 1 PDF + 3 Libros Excel + 7 Gráficos de Ingeniería** | **44,0 horas (5,5 jornadas)** | **0,20 hrs (12 min)** | **43,80 horas** | **99,5% (220,0x más rápido)** |

---

## Arquitectura de los 4 Motores en Python del Repositorio

| Archivo / Script | Rol en la Arquitectura del Sistema |
| :--- | :--- |
| **`0_consola_maestra_total.py` / `consola_web.py`** | Aplicación web en Streamlit con interfaz en modo claro fijo, tablas editables con botones dinámicos (`➕ Agregar` / `🗑️ Quitar`), recálculo instantáneo de VAN/TIR Privado y Social y empaquetador `.zip` en servidor. |
| **`1_simulador_interactivo_cartonvivo.py`** | Simulador de escritorio (`Tkinter` + `Matplotlib`) y motor base de evaluación privada y social a 10 años, cálculo de depreciaciones, valor residual, capital de trabajo por período de desfase (45 días) y liberación de recursos municipales (Vertedero de Curaco). |
| **`2_corrector_formato_y_excels.py`** | Generador automatizado de los **3 Libros Excel oficiales** (`Anexo_1_Flujos_Financieros.xlsx`, `Anexo_2_Informantes_y_Catastro.xlsx`, `Anexo_3_Proyecciones_y_Sensibilidad.xlsx`) y constructor del informe maestro en Word sobre la plantilla institucional de la Universidad de Los Lagos. |
| **`3_integrador_arboles_y_pauta_josue.py`** | Motor de inyección gráfica y estructural: inserta los diagramas de ingeniería (Árbol de Problemas, Árbol de Objetivos, Balance de Masa 70%/30%, Layout de Galpón en Ovejería y Flujo de Proceso), la Matriz de Marco Lógico (MML), el Business Model Canvas y compila el archivo `.pdf` y `.zip` final. |

---

## Instalación y Ejecución Local

```bash
git clone [https://github.com/mauriciovargas1516-ux/PROYECTOS-2026.git](https://github.com/mauriciovargas1516-ux/PROYECTOS-2026.git)
cd PROYECTOS-2026
pip install -r requirements.txt
streamlit run 0_consola_maestra_total.py
```
