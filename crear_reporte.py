"""
=============================================================================
SCRIPT GENERADOR DEL REPORTE DETALLADO DE INSTALACIÓN Y DESPLIEGUE DE SHINY
=============================================================================
Genera dos documentos formales:
  1. Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx (Microsoft Word)
  2. Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf  (Documento PDF)
=============================================================================
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

DIR_BASE = os.path.dirname(os.path.abspath(__file__))
DIR_IMG = os.path.join(DIR_BASE, "reporte_img")
os.makedirs(DIR_IMG, exist_ok=True)

RUTA_DOCX = os.path.join(DIR_BASE, "Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx")
RUTA_PDF = os.path.join(DIR_BASE, "Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf")
RUTA_GRAFICA = os.path.join(DIR_IMG, "grafica_histograma.png")
RUTA_FLUJO = os.path.join(DIR_IMG, "flujo_actividad.png")

# -----------------------------------------------------------------------------
# 1. Gráficas ilustrativas para el reporte
# -----------------------------------------------------------------------------
print("1. Generando gráficas para el reporte...")

# Gráfica del Histograma Básico
np.random.seed(42)
valores = np.random.normal(loc=50, scale=10, size=100)

fig, ax = plt.subplots(figsize=(7.5, 3.8), dpi=200)
n, bins, patches = ax.hist(valores, bins=15, color="#2563EB", edgecolor="black", alpha=0.75)
ax.axvline(np.mean(valores), color="#DC2626", linestyle="dashed", linewidth=2, label=f"Media (x̄) = {np.mean(valores):.2f}")
ax.axvline(np.mean(valores) + np.std(valores), color="#059669", linestyle="dotted", linewidth=1.8, label=f"+1 Desv. Est. (σ) = {np.std(valores):.2f}")
ax.axvline(np.mean(valores) - np.std(valores), color="#059669", linestyle="dotted", linewidth=1.8)

ax.set_title("Salida Visual del Dashboard: Histograma de Distribución Normal (N = 100)", fontsize=11, fontweight="bold", pad=10)
ax.set_xlabel("Valor de la Variable", fontsize=9.5)
ax.set_ylabel("Frecuencia Absoluta", fontsize=9.5)
ax.grid(axis="y", linestyle="--", alpha=0.7)
ax.legend(fontsize=8.5, loc="upper right")
plt.tight_layout()
plt.savefig(RUTA_GRAFICA, bbox_inches="tight")
plt.close()

# Diagrama de Flujo Paso a Paso
fig, ax = plt.subplots(figsize=(10.5, 2.4), dpi=200)
ax.axis("off")
pasos = [
    ("1. Instalación", "pip install shiny\nrsconnect-python\npandas, matplotlib", "#1E3A8A"),
    ("2. Estructura", "cd Dashboard\napp.py\nrequirements.txt", "#0284C7"),
    ("3. Desarrollo", "UI (Sidebar + KPIs)\nServer Reactivo\napp.run(port=8000)", "#0D9488"),
    ("4. Ejecución", "shiny run --reload\nhttp://127.0.0.1:8000\nPruebas locales", "#D97706"),
    ("5. Posit & GitHub", "connect.posit.cloud\nUsuario: rncm11\nDeploy en la Nube", "#059669"),
]

from matplotlib.patches import FancyBboxPatch
for i, (titulo, subtitulo, color) in enumerate(pasos):
    x = i * 2.15
    caja = FancyBboxPatch((x, 0.2), 1.8, 1.8, boxstyle="round,pad=0.1", facecolor=color, edgecolor="black", alpha=0.92)
    ax.add_patch(caja)
    ax.text(x + 0.9, 1.5, titulo, ha="center", va="center", color="white", fontsize=9.5, fontweight="bold")
    ax.text(x + 0.9, 0.85, subtitulo, ha="center", va="center", color="white", fontsize=8)
    if i < len(pasos) - 1:
        ax.annotate("", xy=(x + 2.12, 1.1), xytext=(x + 1.83, 1.1),
                    arrowprops=dict(arrowstyle="->", lw=2.2, color="#374151"))

ax.set_xlim(-0.2, len(pasos) * 2.15)
ax.set_ylim(0, 2.2)
plt.tight_layout()
plt.savefig(RUTA_FLUJO, bbox_inches="tight")
plt.close()
print("Gráficas guardadas correctamente.")

# -----------------------------------------------------------------------------
# 2. Documento Word (.docx)
# -----------------------------------------------------------------------------
print("2. Generando documento Word (.docx)...")
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc = Document()

for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

def sombrear_celda(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def agregar_codigo(doc, texto):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    c = table.cell(0, 0)
    sombrear_celda(c, "F1F5F9")
    c.text = ""
    cp = c.paragraphs[0]
    cp.paragraph_format.space_before = Pt(4)
    cp.paragraph_format.space_after = Pt(4)
    crun = cp.add_run(texto)
    crun.font.name = "Consolas"
    crun.font.size = Pt(8.5)
    crun.font.color.rgb = RGBColor(15, 23, 42)
    p_vacio = doc.add_paragraph()
    p_vacio.paragraph_format.space_before = Pt(2)
    p_vacio.paragraph_format.space_after = Pt(2)

# PORTADA FORMAL
p_uni = doc.add_paragraph()
p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_uni = p_uni.add_run("PROGRAMACIÓN EN CIENCIA DE DATOS\nREPORTE DE ACTIVIDAD PRÁCTICA")
run_uni.font.name = "Calibri"
run_uni.font.size = Pt(13)
run_uni.font.bold = True
run_uni.font.color.rgb = RGBColor(71, 85, 105)

p_tit = doc.add_paragraph()
p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_tit.paragraph_format.space_before = Pt(16)
p_tit.paragraph_format.space_after = Pt(12)
run_tit = p_tit.add_run("GUÍA INTEGRAL DE INSTALACIÓN DE SHINY, CONSTRUCCIÓN DE UN DASHBOARD BÁSICO Y PUBLICACIÓN EN POSIT CONNECT CLOUD")
run_tit.font.name = "Calibri"
run_tit.font.size = Pt(19)
run_tit.font.bold = True
run_tit.font.color.rgb = RGBColor(30, 64, 175)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_sub = p_sub.add_run("Proceso Paso a Paso: Desde la Instalación del Entorno hasta la Inicialización y Despliegue en la Nube")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(11)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(100, 116, 139)

doc.add_paragraph().paragraph_format.space_after = Pt(14)

# Metadatos
t_meta = doc.add_table(rows=5, cols=2)
t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
t_meta.columns[0].width = Inches(2.2)
t_meta.columns[1].width = Inches(4.3)
campos = [
    ("Actividad:", "Instalación de Shiny, Dashboard Básico y Despliegue en Posit"),
    ("Cuenta Posit Connect Cloud:", "rncm11 (https://connect.posit.cloud/rncm11)"),
    ("Repositorio Oficial GitHub:", "https://github.com/RolandoCastillo11/Shiny-Prog-CD"),
    ("Herramientas Principales:", "Python 3.11, Shiny for Python (v1.8.0), Posit Connect Cloud, Git"),
    ("Fecha:", "Octubre de 2026"),
]
for i, (k, v) in enumerate(campos):
    c0, c1 = t_meta.cell(i, 0), t_meta.cell(i, 1)
    sombrear_celda(c0, "E2E8F0")
    sombrear_celda(c1, "F8FAFC")
    r0 = c0.paragraphs[0].add_run(k)
    r0.font.bold = True
    r0.font.size = Pt(9.5)
    r1 = c1.paragraphs[0].add_run(v)
    r1.font.size = Pt(9.5)

doc.add_page_break()

# SECCIÓN 1: INTRODUCCIÓN Y OBJETIVOS
h1 = doc.add_heading("1. Introducción y Requisitos Previos", level=1)
h1.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "En la Ciencia de Datos, la comunicación de resultados es tan importante como el análisis estadístico o el entrenamiento "
    "de modelos. Tradicionalmente, crear una aplicación web interactiva requería programar tanto el frontend (HTML, CSS, JavaScript) "
    "como el backend (servidores Flask, Django o FastAPI). Shiny for Python simplifica este proceso radicalmente, permitiendo "
    "crear tableros interactivos completos programando exclusivamente en Python mediante un modelo reactivo."
)

doc.add_paragraph(
    "Requisitos Previos del Sistema:\n"
    "• Python 3.10 o superior (en esta práctica se utilizó Python 3.11.0).\n"
    "• Gestor de paquetes pip actualizado.\n"
    "• Terminal o consola del sistema operativo (PowerShell o CMD en Windows).\n"
    "• Cuenta activa en GitHub para control de versiones.\n"
    "• Cuenta gratuita en Posit Connect Cloud para el hosting en internet."
)

p_obj = doc.add_paragraph()
p_obj.add_run("Objetivos de la Práctica:\n").bold = True
p_obj.add_run("1. Instalar Shiny for Python y todas sus dependencias necesarias en la máquina local.\n")
p_obj.add_run("2. Crear la estructura del proyecto y desarrollar un Dashboard Básico que demuestre entradas interactivas, KPIs y visualización gráfica reactiva.\n")
p_obj.add_run("3. Explicar los métodos de inicialización y ejecución local de la aplicación en el navegador web.\n")
p_obj.add_run("4. Tramitar la cuenta en Posit Connect Cloud (cuenta: rncm11) y vincular el repositorio en GitHub.\n")
p_obj.add_run("5. Desplegar el dashboard en internet de forma pública y accesible globalmente.")

if os.path.exists(RUTA_FLUJO):
    doc.add_paragraph().paragraph_format.space_before = Pt(6)
    doc.add_picture(RUTA_FLUJO, width=Inches(6.4))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 1: Diagrama secuencial de las 5 fases desarrolladas en la práctica.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

# SECCIÓN 2: INSTALACIÓN PASO A PASO DE SHINY
h2 = doc.add_heading("2. Paso a Paso: Instalación de Shiny y Dependencias", level=1)
h2.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "La instalación de Shiny y su entorno se llevó a cabo desde la consola Windows PowerShell. A continuación se "
    "detallan los comandos ejecutados y su función técnica:"
)

p_i1 = doc.add_paragraph()
p_i1.add_run("Paso 2.1: Verificación de la versión de Python instalada\n").bold = True
p_i1.add_run("Antes de instalar cualquier librería, es fundamental confirmar que Python está accesible en la variable PATH del sistema:")
agregar_codigo(doc, "python --version\n# Salida: Python 3.11.0")

p_i2 = doc.add_paragraph()
p_i2.add_run("Paso 2.2: Instalación de Shiny for Python\n").bold = True
p_i2.add_run(
    "Shiny se instala mediante el gestor oficial pip. Al ejecutar esta instrucción, pip descarga tanto Shiny como sus motores "
    "internos: Uvicorn (servidor ASGI asíncrono para Python), Starlette (framework base de enrutamiento web), "
    "HTMLTools (generador de etiquetas HTML) y BSLib (estilos visuales modernos basados en Bootstrap):"
)
agregar_codigo(doc, "pip install shiny")

p_i3 = doc.add_paragraph()
p_i3.add_run("Paso 2.3: Instalación de librerías analíticas y del conector de Posit\n").bold = True
p_i3.add_run(
    "Para dotar al dashboard de capacidades de generación de datos, cálculo matricial y gráficos estadísticos, "
    "así como la herramienta de publicación en la nube, se instalaron los paquetes complementarios:"
)
agregar_codigo(doc, (
    "pip install numpy pandas matplotlib jinja2 rsconnect-python\n\n"
    "# Función de cada librería:\n"
    "# - numpy: Generación de distribuciones aleatorias (Normal, Uniforme, Exponencial)\n"
    "# - pandas: Manejo tabular de los datos y creación del DataFrame\n"
    "# - matplotlib: Renderizado del histograma de frecuencias\n"
    "# - jinja2: Motor de plantillas HTML requerido para formateo tabular seguro\n"
    "# - rsconnect-python: Herramienta CLI oficial de Posit para despliegue en la nube"
))

p_i4 = doc.add_paragraph()
p_i4.add_run("Paso 2.4: Comprobación y verificación de la instalación\n").bold = True
p_i4.add_run("Se verificó en terminal la correcta instalación y la versión disponible del CLI de Shiny:")
agregar_codigo(doc, (
    "shiny --version\n"
    "# Salida esperada: shiny, version 1.8.0\n\n"
    "rsconnect version\n"
    "# Salida esperada: 1.32.0"
))

# SECCIÓN 3: PREPARACIÓN DEL DIRECTORIO Y ARCHIVOS
h3 = doc.add_heading("3. Preparación del Directorio y Estructura del Proyecto", level=1)
h3.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Para mantener el código organizado de forma modular, se creó un directorio dedicado para la aplicación: Dashboard_Basico_Shiny."
)

p_d1 = doc.add_paragraph()
p_d1.add_run("Paso 3.1: Creación de la carpeta y navegación a la ruta de trabajo\n").bold = True
p_d1.add_run("En la consola de comandos se ejecutó:")
agregar_codigo(doc, (
    "# Crear la carpeta del proyecto\n"
    "mkdir Dashboard_Basico_Shiny\n\n"
    "# Ubicarse DENTRO de la carpeta (Paso crucial para evitar errores de 'app.py not found')\n"
    "cd Dashboard_Basico_Shiny"
))

p_d2 = doc.add_paragraph()
p_d2.add_run("Paso 3.2: Estructura de archivos del proyecto\n").bold = True
p_d2.add_run("Dentro de la carpeta se establecieron los archivos mínimos necesarios:")
agregar_codigo(doc, (
    "Dashboard_Basico_Shiny/\n"
    "├── app.py             # Código fuente interactivo completo (UI + Server)\n"
    "├── requirements.txt   # Lista de librerías para que Posit instale en la nube\n"
    "├── .gitignore         # Exclusión de archivos temporales (__pycache__, .venv)\n"
    "├── manifest.json      # Manifiesto de compilación generado por rsconnect\n"
    "└── README.md          # Documentación del proyecto para GitHub"
))

# SECCIÓN 4: DESARROLLO DEL CÓDIGO BÁSICO
h4 = doc.add_heading("4. Desarrollo y Explicación del Código del Dashboard Básico", level=1)
h4.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Siguiendo el objetivo de construir un dashboard básico, claro y sin sobrecargas, se desarrolló el "
    "'Generador de Distribuciones Estadísticas' en app.py. Este ejemplo permite comprender a fondo los tres "
    "pilares de una aplicación Shiny: la Interfaz de Usuario (UI), la Lógica del Servidor (Server) y la Inicialización."
)

doc.add_heading("4.1 Explicación de la Interfaz de Usuario (UI - app_ui)", level=2)
p_ui = doc.add_paragraph()
p_ui.add_run("• ui.page_sidebar: ").bold = True
p_ui.add_run("Genera una distribución moderna con una barra lateral izquierda de controles y un panel principal derecho.\n")
p_ui.add_run("• Controles de Entrada (Inputs):\n").bold = True
p_ui.add_run("  - ui.input_slider('n_muestras', ...): Deslizador que define cuántos números aleatorios generar (20 a 500).\n")
p_ui.add_run("  - ui.input_select('tipo_distribucion', ...): Menú desplegable para alternar entre Normal, Uniforme y Exponencial.\n")
p_ui.add_run("  - ui.input_radio_buttons('color_grafica', ...): Botones de selección única para cambiar el color del histograma.\n")
p_ui.add_run("• Tarjetas de Indicadores (ui.value_box): ").bold = True
p_ui.add_run("Muestran de forma visual y destacada tres KPIs en tiempo real: Media Aritmética (x̄), Desviación Estándar (σ) y Total de Muestras.\n")
p_ui.add_run("• Contenedores de Salida (Outputs):\n").bold = True
p_ui.add_run("  - ui.output_plot('grafico_histograma'): Espacio donde se dibuja el gráfico generado con Matplotlib.\n")
p_ui.add_run("  - ui.output_data_frame('tabla_primeros_datos'): Contenedor donde se despliegan las primeras 10 filas del conjunto de datos generado.")

doc.add_heading("4.2 Explicación de la Lógica del Servidor (Server)", level=2)
doc.add_paragraph(
    "El servidor contiene las funciones reactivas que responden a la interacción del usuario:\n"
    "1. Decorador @reactive.calc (Función generar_datos): Este cálculo reactivo genera el conjunto de datos con NumPy y Pandas. "
    "Shiny almacena el resultado en memoria caché y solo vuelve a calcularlo si el usuario mueve el slider o cambia de distribución, "
    "evitando procesamiento innecesario.\n"
    "2. Decoradores @render.text (kpi_media, kpi_desviacion, kpi_total): Extraen las métricas estadísticas del cálculo reactivo y "
    "las envían a los Value Boxes correspondientes.\n"
    "3. Decorador @render.plot (grafico_histograma): Genera la figura de Matplotlib dibujando las barras del histograma con el color "
    "seleccionado por el usuario.\n"
    "4. Decorador @render.data_frame (tabla_primeros_datos): Convierte el DataFrame de Pandas en un DataGrid interactivo de Shiny sin "
    "dependencias externas complejas."
)

if os.path.exists(RUTA_GRAFICA):
    doc.add_paragraph().paragraph_format.space_before = Pt(6)
    doc.add_picture(RUTA_GRAFICA, width=Inches(5.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 2: Histograma de frecuencias y líneas estadísticas generado por el Dashboard Básico.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

doc.add_heading("4.3 Código Fuente Completo de app.py", level=2)
doc.add_paragraph("A continuación se muestra el código completo implementado en el archivo app.py:")
agregar_codigo(doc, (
    "import numpy as np\n"
    "import pandas as pd\n"
    "import matplotlib.pyplot as plt\n"
    "from shiny import App, render, ui, reactive\n\n"
    "# -----------------------------------------------------------------------------\n"
    "# 1. Interfaz de Usuario (UI)\n"
    "# -----------------------------------------------------------------------------\n"
    "app_ui = ui.page_sidebar(\n"
    "    ui.sidebar(\n"
    "        ui.h4('Panel de Control', class_='fw-bold text-primary mb-3'),\n"
    "        ui.input_slider('n_muestras', 'Número de Muestras (N):', min=20, max=500, value=100, step=10),\n"
    "        ui.input_select('tipo_distribucion', 'Tipo de Distribución:',\n"
    "                        {'normal': 'Normal (Gaussiana)', 'uniforme': 'Uniforme', 'exponencial': 'Exponencial'}),\n"
    "        ui.input_radio_buttons('color_grafica', 'Color de las Barras:',\n"
    "                               {'#2563EB': 'Azul', '#059669': 'Verde', '#DC2626': 'Rojo', '#7C3AED': 'Púrpura'}),\n"
    "        ui.hr(),\n"
    "        ui.p('Dashboard básico para despliegue en Posit Connect Cloud.', class_='text-muted small'),\n"
    "        width=280,\n"
    "    ),\n"
    "    ui.h3('📊 Dashboard Básico: Generador de Distribuciones Estadísticas', class_='fw-bold mb-3'),\n"
    "    ui.layout_columns(\n"
    "        ui.value_box('Media Aritmética (x̄)', ui.output_text('kpi_media'), showcase=ui.div('📈', style='font-size: 2rem;'), theme='primary'),\n"
    "        ui.value_box('Desviación Estándar (σ)', ui.output_text('kpi_desviacion'), showcase=ui.div('📏', style='font-size: 2rem;'), theme='success'),\n"
    "        ui.value_box('Total de Observaciones', ui.output_text('kpi_total'), showcase=ui.div('🔢', style='font-size: 2rem;'), theme='info'),\n"
    "        col_widths=[4, 4, 4],\n"
    "    ),\n"
    "    ui.layout_columns(\n"
    "        ui.card(ui.card_header('Histograma de Frecuencias en Tiempo Real', class_='fw-bold'), ui.output_plot('grafico_histograma')),\n"
    "        ui.card(ui.card_header('Primeras 10 Observaciones del Conjunto', class_='fw-bold'), ui.output_data_frame('tabla_primeros_datos')),\n"
    "        col_widths=[7, 5],\n"
    "    ),\n"
    "    title='Dashboard Básico - Shiny',\n"
    "    fillable=True,\n"
    ")\n\n"
    "# -----------------------------------------------------------------------------\n"
    "# 2. Lógica del Servidor (Server)\n"
    "# -----------------------------------------------------------------------------\n"
    "def server(input, output, session):\n"
    "    @reactive.calc\n"
    "    def generar_datos():\n"
    "        n = input.n_muestras()\n"
    "        tipo = input.tipo_distribucion()\n"
    "        np.random.seed(42)\n"
    "        if tipo == 'normal': valores = np.random.normal(loc=50, scale=10, size=n)\n"
    "        elif tipo == 'uniforme': valores = np.random.uniform(low=20, high=80, size=n)\n"
    "        else: valores = np.random.exponential(scale=20, size=n)\n"
    "        return pd.DataFrame({'ID': range(1, n + 1), 'Valor': np.round(valores, 2)})\n\n"
    "    @render.text\n"
    "    def kpi_media(): return f\"{generar_datos()['Valor'].mean():.2f}\"\n\n"
    "    @render.text\n"
    "    def kpi_desviacion(): return f\"{generar_datos()['Valor'].std():.2f}\"\n\n"
    "    @render.text\n"
    "    def kpi_total(): return f\"{len(generar_datos())} datos\"\n\n"
    "    @render.plot\n"
    "    def grafico_histograma():\n"
    "        df = generar_datos()\n"
    "        color = input.color_grafica()\n"
    "        tipo = input.tipo_distribucion()\n"
    "        fig, ax = plt.subplots(figsize=(6, 4))\n"
    "        ax.hist(df['Valor'], bins=15, color=color, edgecolor='black', alpha=0.75)\n"
    "        ax.set_title(f'Histograma de Distribución: {tipo.capitalize()}', fontsize=12, fontweight='bold')\n"
    "        ax.set_xlabel('Valor de la Variable', fontsize=10)\n"
    "        ax.set_ylabel('Frecuencia Absoluta', fontsize=10)\n"
    "        ax.grid(axis='y', linestyle='--', alpha=0.7)\n"
    "        plt.tight_layout()\n"
    "        return fig\n\n"
    "    @render.data_frame\n"
    "    def tabla_primeros_datos():\n"
    "        df = generar_datos()\n"
    "        return render.DataGrid(df.head(10))\n\n"
    "# -----------------------------------------------------------------------------\n"
    "# 3. Inicialización y Ejecución de la App\n"
    "# -----------------------------------------------------------------------------\n"
    "app = App(app_ui, server)\n\n"
    "if __name__ == '__main__':\n"
    "    app.run(port=8000)"
))

# SECCIÓN 5: CÓMO SE INICIALIZA Y EJECUTA LA APP
doc.add_page_break()
h5 = doc.add_heading("5. Cómo se Inicializa y Ejecuta la Aplicación", level=1)
h5.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Una aplicación Shiny for Python puede inicializarse de dos formas en la computadora local. Ambas opciones "
    "levantan un servidor web local y permiten visualizar la interfaz desde el navegador:"
)

p_m1 = doc.add_paragraph()
p_m1.add_run("Método 1: Ejecución con el Servidor CLI de Shiny (Recomendado para desarrollo)\n").bold = True
p_m1.add_run(
    "El comando oficial de Shiny incluye la bandera --reload, la cual detecta automáticamente cualquier cambio guardado en app.py "
    "y recarga el navegador al instante sin necesidad de reiniciar la terminal:"
)
agregar_codigo(doc, (
    "# 1. Asegurarse de estar en la carpeta donde reside app.py:\n"
    "cd \"c:\\Users\\rolca\\Documents\\Prog. CD\\Dashboard_Basico_Shiny\"\n\n"
    "# 2. Ejecutar con el servidor de desarrollo de Shiny:\n"
    "shiny run --reload --port 8000 app.py"
))

p_m2 = doc.add_paragraph()
p_m2.add_run("Método 2: Ejecución directa con el intérprete de Python\n").bold = True
p_m2.add_run(
    "Gracias al bloque añadido al final del código (if __name__ == '__main__': app.run(port=8000)), la aplicación también "
    "puede levantarse como cualquier script convencional de Python:"
)
agregar_codigo(doc, (
    "# Ubicado dentro de Dashboard_Basico_Shiny:\n"
    "python app.py"
))

doc.add_paragraph(
    "Acceso en el Navegador Web:\n"
    "Una vez que la terminal muestra el mensaje 'Uvicorn running on http://127.0.0.1:8000', abre cualquier navegador "
    "(Chrome, Edge, Firefox) y navega a la dirección:\n"
    "👉 http://127.0.0.1:8000  (o http://localhost:8000)"
)

# SECCIÓN 6: RESOLUCIÓN DE INCIDENCIAS
h6 = doc.add_heading("6. Resolución de Incidencias Comunes Durante la Práctica", level=1)
h6.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Durante el desarrollo y prueba de aplicaciones web en terminal, es habitual toparse con errores de entorno o de rutas. "
    "A continuación se documentan las dos incidencias experimentadas y su solución definitiva:"
)

p_e1 = doc.add_paragraph()
p_e1.add_run("Incidencia 1: Error '.\\app.py not found'\n").bold = True
p_e1.add_run(
    "• Causa: Este error ocurre cuando se intenta ejecutar 'shiny run app.py' o 'python app.py' estando ubicado en la carpeta padre "
    "(por ejemplo en 'c:\\Users\\rolca\\Documents\\Prog. CD') en lugar de la subcarpeta donde realmente se encuentra el archivo app.py.\n"
    "• Solución: Cambiar de directorio en la terminal antes de lanzar la aplicación ejecutando: cd \"Dashboard_Basico_Shiny\"."
)

p_e2 = doc.add_paragraph()
p_e2.add_run("Incidencia 2: 'Import Jinja2 failed. DataFrame.style requires jinja2'\n").bold = True
p_e2.add_run(
    "• Causa: Al utilizar la directiva antigua @render.table con un DataFrame de Pandas, la función interna intenta invocar "
    "el formateador CSS DataFrame.style, el cual depende obligatoriamente del paquete externo Jinja2.\n"
    "• Solución Doble Aplicada:\n"
    "  1. Se actualizó app.py utilizando la directiva moderna @render.data_frame con render.DataGrid, la cual es nativa de Shiny, "
    "más rápida y no requiere Jinja2 para tablas básicas.\n"
    "  2. Se agregó 'jinja2>=3.1.0' al archivo requirements.txt para asegurar que los servidores de Posit cuenten con todas las dependencias."
)

# SECCIÓN 7: TRAMITACIÓN EN POSIT Y GITHUB
h7 = doc.add_heading("7. Tramitación de Cuenta en Posit Connect Cloud y Publicación", level=1)
h7.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Para cumplir con el objetivo de publicar la aplicación en internet y ponerla al alcance de cualquier usuario en el mundo, "
    "se realizaron los siguientes trámites y configuraciones:"
)

p_pos = doc.add_paragraph()
p_pos.add_run("7.1 Tramitación de la Cuenta en Posit Connect Cloud\n").bold = True
p_pos.add_run(
    "• Portal de acceso: https://connect.posit.cloud/\n"
    "• Nombre de usuario / Cuenta registrada: rncm11\n"
    "• Perfil y panel de aplicaciones activas: https://connect.posit.cloud/rncm11\n"
    "• Tipo de plan: Gratuito (Cloud Free Tier), ideal para prácticas académicas sin costo de servidor."
)

p_git = doc.add_paragraph()
p_git.add_run("7.2 Creación del Repositorio en GitHub\n").bold = True
p_git.add_run(
    "• Nombre del repositorio: Shiny-Prog-CD\n"
    "• URL pública del código: https://github.com/RolandoCastillo11/Shiny-Prog-CD\n"
    "• Comandos ejecutados para sincronizar el proyecto local con GitHub:\n"
)
agregar_codigo(doc, (
    "git init\n"
    "git add .\n"
    "git commit -m \"Refactor: Dashboard basico simplificado y reporte de instalacion actualizado\"\n"
    "git branch -M main\n"
    "git remote add origin https://github.com/RolandoCastillo11/Shiny-Prog-CD.git\n"
    "git push -u origin main"
))

p_pub = doc.add_paragraph()
p_pub.add_run("7.3 Publicación Final en la Nube de Posit\n").bold = True
p_pub.add_run(
    "Desde el panel web de https://connect.posit.cloud/rncm11:\n"
    "1. Se pulsa el botón 'Publish' situado en la parte superior derecha.\n"
    "2. Se selecciona la fuente de despliegue 'GitHub' y se autoriza la cuenta RolandoCastillo11.\n"
    "3. Se elige el repositorio Shiny-Prog-CD.\n"
    "4. Posit Connect Cloud detecta automáticamente app.py como punto de entrada y requirements.txt como dependencias.\n"
    "5. Al presionar 'Deploy', el clúster de Posit compila el contenedor Docker en Linux y genera la URL pública con protocolo seguro HTTPS."
)

# SECCIÓN 8: TABLA DE VERIFICACIÓN Y CONCLUSIONES
h8 = doc.add_heading("8. Resumen de Actividades y Verificación", level=1)
h8.style.font.color.rgb = RGBColor(30, 64, 175)

t_res = doc.add_table(rows=6, cols=3)
t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
t_res.columns[0].width = Inches(2.0)
t_res.columns[1].width = Inches(2.6)
t_res.columns[2].width = Inches(1.9)

encabezados = ["Actividad Requerida", "Herramienta / Evidencia", "Estado de Verificación"]
for j, enc in enumerate(encabezados):
    celda = t_res.cell(0, j)
    sombrear_celda(celda, "1E3A8A")
    pr = celda.paragraphs[0].add_run(enc)
    pr.font.bold = True
    pr.font.color.rgb = RGBColor(255, 255, 255)
    pr.font.size = Pt(9.5)

filas = [
    ("1. Instalación de Shiny", "pip install shiny (v1.8.0)", "Completada y Verificada ✓"),
    ("2. Dashboard Básico", "app.py (Inputs, KPIs, Histograma, Tabla)", "Completado y Funcionando ✓"),
    ("3. Inicialización Local", "shiny run --reload / python app.py", "Probado en http://127.0.0.1:8000 ✓"),
    ("4. Tramitación en Posit", "connect.posit.cloud/rncm11", "Cuenta Activa Tramitada ✓"),
    ("5. Sincronización GitHub", "RolandoCastillo11/Shiny-Prog-CD", "Código Publicado en Git ✓"),
]

for i, (f1, f2, f3) in enumerate(filas, start=1):
    c0, c1, c2 = t_res.cell(i, 0), t_res.cell(i, 1), t_res.cell(i, 2)
    sombrear_celda(c0, "F8FAFC" if i % 2 == 0 else "FFFFFF")
    sombrear_celda(c1, "F8FAFC" if i % 2 == 0 else "FFFFFF")
    sombrear_celda(c2, "ECFDF5")
    c0.paragraphs[0].add_run(f1).font.size = Pt(9)
    c1.paragraphs[0].add_run(f2).font.size = Pt(9)
    r3 = c2.paragraphs[0].add_run(f3)
    r3.font.size = Pt(9)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(5, 150, 105)

doc.add_paragraph().paragraph_format.space_before = Pt(12)
h9 = doc.add_heading("9. Conclusiones", level=1)
h9.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "1. Shiny for Python se consolida como una herramienta esencial para la ciencia de datos moderna, "
    "permitiendo transformar scripts analíticos estáticos en aplicaciones web dinámicas con un mínimo esfuerzo de código.\n\n"
    "2. La arquitectura basada en decoradores reactivos (@reactive.calc y @render) simplifica el manejo del estado "
    "de la aplicación, garantizando que los cálculos se actualicen exclusivamente cuando cambian los controles del usuario.\n\n"
    "3. La integración entre repositorios de GitHub y la plataforma Posit Connect Cloud permite implementar un flujo "
    "profesional de integración y despliegue continuo (CI/CD), asegurando que cualquier cambio en el código se publique en internet "
    "de forma automatizada, transparente y segura."
)

doc.save(RUTA_DOCX)
print(f"Documento Word guardado exitosamente en: {RUTA_DOCX}")

# -----------------------------------------------------------------------------
# 3. Documento PDF (.pdf) con ReportLab
# -----------------------------------------------------------------------------
print("3. Generando documento PDF (.pdf)...")
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, HRFlowable
)
from reportlab.pdfgen import canvas

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super(NumberedCanvas, self).__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super(NumberedCanvas, self).showPage()
        super(NumberedCanvas, self).save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        if self._pageNumber > 1:
            self.drawString(54, 755, "Reporte Práctico: Instalación de Shiny, Dashboard Básico e Inicialización")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 750, 558, 750)
        pie_txt = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(558, 36, pie_txt)
        self.drawString(54, 36, "Programación en Ciencia de Datos | Octubre 2026")
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(54, 46, 558, 46)
        self.restoreState()

doc_pdf = SimpleDocTemplate(
    RUTA_PDF,
    pagesize=letter,
    leftMargin=54,
    rightMargin=54,
    topMargin=54,
    bottomMargin=54,
)

styles = getSampleStyleSheet()

t_port = ParagraphStyle("TPort", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=17, leading=22, textColor=colors.HexColor("#1E3A8A"), alignment=1, spaceAfter=8)
sub_port = ParagraphStyle("SPort", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=14, textColor=colors.HexColor("#64748B"), alignment=1, spaceAfter=16)
h1_pdf = ParagraphStyle("H1Pdf", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=12.5, leading=16, textColor=colors.HexColor("#1E3A8A"), spaceBefore=11, spaceAfter=5)
h2_pdf = ParagraphStyle("H2Pdf", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=10, leading=13.5, textColor=colors.HexColor("#1E293B"), spaceBefore=8, spaceAfter=4)
p_pdf = ParagraphStyle("PPdf", parent=styles["Normal"], fontName="Helvetica", fontSize=8.8, leading=12.8, textColor=colors.HexColor("#334155"), spaceAfter=5)
c_pdf = ParagraphStyle("CPdf", parent=styles["Normal"], fontName="Courier", fontSize=7.8, leading=10.2, textColor=colors.HexColor("#0F172A"))

def caja_codigo_pdf(texto):
    t = Table([[Paragraph(f"<font color='#0F172A'>{texto.replace(chr(10), '<br/>').replace(' ', '&nbsp;')}</font>", c_pdf)]], colWidths=[504])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t

story = []

# Portada PDF
story.append(Paragraph("<b>PROGRAMACIÓN EN CIENCIA DE DATOS</b><br/><font color='#64748B' size='8.5'>REPORTE DE ACTIVIDAD PRÁCTICA</font>", ParagraphStyle("Port", alignment=1, fontSize=11, leading=14, fontName="Helvetica-Bold", textColor=colors.HexColor("#334155"))))
story.append(Spacer(1, 10))
story.append(Paragraph("GUÍA INTEGRAL DE INSTALACIÓN DE SHINY, CONSTRUCCIÓN DE UN DASHBOARD BÁSICO Y DESPLIEGUE EN POSIT", t_port))
story.append(Paragraph("Proceso Paso a Paso: Desde la Instalación de Librerías hasta la Inicialización y Publicación en la Nube", sub_port))
story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceBefore=2, spaceAfter=10))

# Tabla de Metadatos
datos_meta = [
    [Paragraph("<b>Actividad:</b>", p_pdf), Paragraph("Instalación de Shiny, Dashboard Básico y Publicación en Internet", p_pdf)],
    [Paragraph("<b>Cuenta en Posit:</b>", p_pdf), Paragraph("rncm11 (https://connect.posit.cloud/rncm11)", p_pdf)],
    [Paragraph("<b>Repositorio GitHub:</b>", p_pdf), Paragraph("https://github.com/RolandoCastillo11/Shiny-Prog-CD", p_pdf)],
    [Paragraph("<b>Herramientas:</b>", p_pdf), Paragraph("Python 3.11, Shiny for Python (v1.8.0), Posit Connect Cloud, Git", p_pdf)],
    [Paragraph("<b>Fecha:</b>", p_pdf), Paragraph("Octubre de 2026", p_pdf)],
]
t_m = Table(datos_meta, colWidths=[125, 379])
t_m.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E2E8F0")),
    ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#F8FAFC")),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94A3B8")),
    ("TOPPADDING", (0, 0), (-1, -1), 3.5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ("LEFTPADDING", (0, 0), (-1, -1), 7),
    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
]))
story.append(t_m)
story.append(Spacer(1, 8))

# 1. INTRODUCCIÓN
story.append(Paragraph("1. Introducción y Requisitos Previos", h1_pdf))
story.append(Paragraph(
    "El propósito de esta práctica es familiarizarse con el framework <b>Shiny for Python</b>, instalando su entorno "
    "completo, estructurando un <b>dashboard básico e interactivo</b>, documentando su inicialización y publicándolo "
    "en internet mediante <b>Posit Connect Cloud</b> (cuenta: <code>rncm11</code>) y <b>GitHub</b>.", p_pdf
))

if os.path.exists(RUTA_FLUJO):
    story.append(Spacer(1, 3))
    story.append(Image(RUTA_FLUJO, width=500, height=114))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 1: Fases secuenciales de la actividad práctica.</i></font>", ParagraphStyle("Cap1", alignment=1, spaceAfter=5)))

# 2. INSTALACIÓN
story.append(Paragraph("2. Paso a Paso: Instalación de Shiny y Dependencias", h1_pdf))
story.append(Paragraph(
    "Se realizaron los siguientes pasos en la consola Windows PowerShell:<br/>"
    "1. <b>Verificación de Python:</b> <code>python --version</code> (Python 3.11.0 detectado).<br/>"
    "2. <b>Instalación de Shiny:</b> <code>pip install shiny</code> (instala el servidor ASGI Uvicorn, Starlette, HTMLTools y BSLib).<br/>"
    "3. <b>Instalación de Librerías Analíticas:</b> <code>pip install numpy pandas matplotlib jinja2 rsconnect-python</code>.<br/>"
    "4. <b>Comprobación de Versiones:</b> <code>shiny --version</code> (retorna versión 1.8.0) y <code>rsconnect version</code> (retorna 1.32.0).", p_pdf
))

# 3. ESTRUCTURA Y DASHBOARD
story.append(PageBreak())
story.append(Paragraph("3. Preparación del Directorio y Desarrollo del Dashboard Básico", h1_pdf))
story.append(Paragraph(
    "Se creó la carpeta <code>Dashboard_Basico_Shiny/</code> para alojar <code>app.py</code>, <code>requirements.txt</code> "
    "y el control de versiones. El dashboard básico explora distribuciones estadísticas con entradas, KPIs, gráfica y tabla:", p_pdf
))

if os.path.exists(RUTA_GRAFICA):
    story.append(Spacer(1, 3))
    story.append(Image(RUTA_GRAFICA, width=460, height=165))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 2: Histograma de frecuencias generado reactivamente por el Dashboard.</i></font>", ParagraphStyle("Cap2", alignment=1, spaceAfter=5)))

story.append(Paragraph("<b>Código fuente representativo de app.py:</b>", h2_pdf))
story.append(caja_codigo_pdf(
    "from shiny import App, render, ui, reactive\n"
    "import numpy as np, pandas as pd, matplotlib.pyplot as plt\n\n"
    "# --- 1. Interfaz de Usuario (UI) ---\n"
    "app_ui = ui.page_sidebar(\n"
    "    ui.sidebar(\n"
    "        ui.input_slider('n_muestras', 'Muestras (N):', min=20, max=500, value=100, step=10),\n"
    "        ui.input_select('tipo_distribucion', 'Distribución:', {'normal':'Normal','uniforme':'Uniforme','exponencial':'Exponencial'}),\n"
    "        ui.input_radio_buttons('color_grafica', 'Color:', {'#2563EB':'Azul','#059669':'Verde','#DC2626':'Rojo'}),\n"
    "        width=280,\n"
    "    ),\n"
    "    ui.h3('Dashboard Básico: Generador de Distribuciones Estadísticas'),\n"
    "    ui.layout_columns(\n"
    "        ui.value_box('Media (x̄)', ui.output_text('kpi_media'), theme='primary'),\n"
    "        ui.value_box('Desv. Est. (σ)', ui.output_text('kpi_desviacion'), theme='success'),\n"
    "        ui.value_box('Total Muestras', ui.output_text('kpi_total'), theme='info'),\n"
    "        col_widths=[4, 4, 4],\n"
    "    ),\n"
    "    ui.layout_columns(\n"
    "        ui.card(ui.card_header('Histograma'), ui.output_plot('grafico_histograma')),\n"
    "        ui.card(ui.card_header('Primeras 10 Observaciones'), ui.output_data_frame('tabla_primeros_datos')),\n"
    "        col_widths=[7, 5],\n"
    "    ),\n"
    "    title='Dashboard Básico - Shiny',\n"
    ")\n\n"
    "# --- 2. Lógica del Servidor (Server) ---\n"
    "def server(input, output, session):\n"
    "    @reactive.calc\n"
    "    def generar_datos():\n"
    "        n, tipo = input.n_muestras(), input.tipo_distribucion()\n"
    "        np.random.seed(42)\n"
    "        if tipo == 'normal': vals = np.random.normal(50, 10, n)\n"
    "        elif tipo == 'uniforme': vals = np.random.uniform(20, 80, n)\n"
    "        else: vals = np.random.exponential(20, n)\n"
    "        return pd.DataFrame({'ID': range(1, n+1), 'Valor': np.round(vals, 2)})\n\n"
    "    @render.text\n"
    "    def kpi_media(): return f\"{generar_datos()['Valor'].mean():.2f}\"\n"
    "    @render.text\n"
    "    def kpi_desviacion(): return f\"{generar_datos()['Valor'].std():.2f}\"\n"
    "    @render.text\n"
    "    def kpi_total(): return f\"{len(generar_datos())} datos\"\n\n"
    "    @render.plot\n"
    "    def grafico_histograma():\n"
    "        fig, ax = plt.subplots(figsize=(6, 4))\n"
    "        ax.hist(generar_datos()['Valor'], bins=15, color=input.color_grafica(), edgecolor='black')\n"
    "        return fig\n\n"
    "    @render.data_frame\n"
    "    def tabla_primeros_datos(): return render.DataGrid(generar_datos().head(10))\n\n"
    "# --- 3. Inicialización ---\n"
    "app = App(app_ui, server)\n"
    "if __name__ == '__main__':\n"
    "    app.run(port=8000)"
))

# 4. INICIALIZACIÓN Y EJECUCIÓN LOCAL
story.append(Paragraph("4. Cómo se Inicializa y Ejecuta la Aplicación", h1_pdf))
story.append(Paragraph(
    "Para correr la aplicación es indispensable ubicarse en el directorio del script:<br/>"
    "<code>cd \"Dashboard_Basico_Shiny\"</code><br/>"
    "<b>Opción A (Servidor Shiny con recarga en caliente):</b> <code>shiny run --reload --port 8000 app.py</code><br/>"
    "<b>Opción B (Ejecución directa en Python):</b> <code>python app.py</code><br/>"
    "La consola confirma: <i>Uvicorn running on http://127.0.0.1:8000</i>. Al abrir esa URL en el navegador se interactúa con el dashboard.", p_pdf
))

# 5. RESOLUCIÓN DE INCIDENCIAS
story.append(Paragraph("5. Resolución de Incidencias Durante la Práctica", h1_pdf))
story.append(Paragraph(
    "• <b>Error '.\\app.py not found':</b> Ocurre si se ejecuta el comando fuera de la carpeta <code>Dashboard_Basico_Shiny</code>. Solución: <code>cd Dashboard_Basico_Shiny</code>.<br/>"
    "• <b>Error 'Import Jinja2 failed':</b> Ocurría por <code>@render.table</code> llamando a <code>DataFrame.style</code>. Solución: Se migró a <code>@render.data_frame</code> con <code>render.DataGrid</code> (nativo de Shiny) y se añadió <code>jinja2>=3.1.0</code> a <code>requirements.txt</code>.", p_pdf
))

# 6. POSIT Y GITHUB
story.append(Paragraph("6. Tramitación de Cuenta en Posit y Publicación a Internet", h1_pdf))
story.append(Paragraph(
    "• <b>Cuenta en Posit Connect Cloud:</b> <u>https://connect.posit.cloud/rncm11</u> (Usuario: <code>rncm11</code>).<br/>"
    "• <b>Repositorio en GitHub:</b> <u>https://github.com/RolandoCastillo11/Shiny-Prog-CD</u>.<br/>"
    "• <b>Publicación en la Nube:</b> En <i>connect.posit.cloud</i> se pulsa <b>Publish</b> -> <b>GitHub</b> -> <b>RolandoCastillo11/Shiny-Prog-CD</b>. "
    "Posit detecta <code>app.py</code> y <code>requirements.txt</code>, compila el contenedor seguro y entrega el enlace público HTTPS.", p_pdf
))

# TABLA RESUMEN
story.append(Paragraph("7. Tabla de Verificación de la Práctica", h1_pdf))
datos_act = [
    [Paragraph("<b>Actividad</b>", p_pdf), Paragraph("<b>Herramienta / Evidencia</b>", p_pdf), Paragraph("<b>Estado</b>", p_pdf)],
    [Paragraph("1. Instalación de Shiny", p_pdf), Paragraph("pip install shiny (v1.8.0)", p_pdf), Paragraph("<font color='#059669'><b>✓ Completada</b></font>", p_pdf)],
    [Paragraph("2. Dashboard Básico", p_pdf), Paragraph("app.py (Inputs, KPIs, Histograma)", p_pdf), Paragraph("<font color='#059669'><b>✓ Completado</b></font>", p_pdf)],
    [Paragraph("3. Inicialización Local", p_pdf), Paragraph("shiny run / python app.py (:8000)", p_pdf), Paragraph("<font color='#059669'><b>✓ En Ejecución</b></font>", p_pdf)],
    [Paragraph("4. Cuenta en Posit", p_pdf), Paragraph("connect.posit.cloud/rncm11", p_pdf), Paragraph("<font color='#059669'><b>✓ Tramitada</b></font>", p_pdf)],
    [Paragraph("5. Repositorio GitHub", p_pdf), Paragraph("Shiny-Prog-CD", p_pdf), Paragraph("<font color='#059669'><b>✓ Sincronizado</b></font>", p_pdf)],
]
t_act = Table(datos_act, colWidths=[125, 224, 155])
t_act.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94A3B8")),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
]))
story.append(t_act)

doc_pdf.build(story, canvasmaker=NumberedCanvas)
print(f"Documento PDF guardado exitosamente en: {RUTA_PDF}")
print("DOCUMENTOS REGENERADOS CON ÉXITO.")
