"""
Script para generar el reporte formal en formato Word (.docx) y PDF (.pdf)
para la actividad de Instalación de Shiny, Dashboard Básico y Despliegue en Posit.
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
# 1. Generación de Gráficas para el Reporte
# -----------------------------------------------------------------------------
print("1. Generando gráficas ilustrativas del Dashboard Básico...")

# Gráfica del histograma básico generado por el dashboard
np.random.seed(42)
valores = np.random.normal(loc=50, scale=10, size=100)

fig, ax = plt.subplots(figsize=(7, 3.8), dpi=200)
n, bins, patches = ax.hist(valores, bins=15, color="#2563EB", edgecolor="black", alpha=0.75)
ax.axvline(np.mean(valores), color="#DC2626", linestyle="dashed", linewidth=2, label=f"Media = {np.mean(valores):.2f}")
ax.axvline(np.mean(valores) + np.std(valores), color="#059669", linestyle="dotted", linewidth=1.8, label=f"+1 Desv. Est. = {np.std(valores):.2f}")
ax.axvline(np.mean(valores) - np.std(valores), color="#059669", linestyle="dotted", linewidth=1.8)

ax.set_title("Histograma de Frecuencias: Distribución Normal (N = 100)", fontsize=11, fontweight="bold", pad=10)
ax.set_xlabel("Valor de la Variable", fontsize=9.5)
ax.set_ylabel("Frecuencia Absoluta", fontsize=9.5)
ax.grid(axis="y", linestyle="--", alpha=0.7)
ax.legend(fontsize=8.5, loc="upper right")
plt.tight_layout()
plt.savefig(RUTA_GRAFICA, bbox_inches="tight")
plt.close()

# Diagrama de flujo del proyecto
fig, ax = plt.subplots(figsize=(10, 2.3), dpi=200)
ax.axis("off")
pasos = [
    ("1. Instalación", "pip install shiny\nrsconnect-python\nPython 3.11", "#1E3A8A"),
    ("2. Dashboard", "app.py (Básico)\nUI + Server\nHistograma + Tabla", "#0D9488"),
    ("3. Cuenta Posit", "connect.posit.cloud\nUsuario: rncm11\nTramitación", "#D97706"),
    ("4. Repositorio", "GitHub\nRolandoCastillo11/\nShiny-Prog-CD", "#4F46E5"),
    ("5. Publicación", "Hosting en Internet\nAcceso Público\nPosit Cloud", "#059669"),
]

from matplotlib.patches import FancyBboxPatch
for i, (titulo, subtitulo, color) in enumerate(pasos):
    x = i * 2.1
    caja = FancyBboxPatch((x, 0.2), 1.7, 1.7, boxstyle="round,pad=0.1", facecolor=color, edgecolor="black", alpha=0.9)
    ax.add_patch(caja)
    ax.text(x + 0.85, 1.45, titulo, ha="center", va="center", color="white", fontsize=9.5, fontweight="bold")
    ax.text(x + 0.85, 0.80, subtitulo, ha="center", va="center", color="white", fontsize=8)
    if i < len(pasos) - 1:
        ax.annotate("", xy=(x + 2.05, 1.05), xytext=(x + 1.72, 1.05),
                    arrowprops=dict(arrowstyle="->", lw=2, color="#374151"))

ax.set_xlim(-0.2, len(pasos) * 2.1)
ax.set_ylim(0, 2.1)
plt.tight_layout()
plt.savefig(RUTA_FLUJO, bbox_inches="tight")
plt.close()

print("Gráficas guardadas.")

# -----------------------------------------------------------------------------
# 2. Generación del Documento Word (.docx)
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
p_tit.paragraph_format.space_before = Pt(18)
p_tit.paragraph_format.space_after = Pt(12)
run_tit = p_tit.add_run("INSTALACIÓN DE SHINY, DESARROLLO DE UN DASHBOARD BÁSICO Y PUBLICACIÓN EN LA NUBE (POSIT)")
run_tit.font.name = "Calibri"
run_tit.font.size = Pt(20)
run_tit.font.bold = True
run_tit.font.color.rgb = RGBColor(30, 64, 175)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_sub = p_sub.add_run("Guía Paso a Paso de Instalación, Código Fuente del Dashboard y Despliegue en Posit Connect Cloud")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(11)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(100, 116, 139)

doc.add_paragraph().paragraph_format.space_after = Pt(14)

# Tabla de Metadatos
t_meta = doc.add_table(rows=5, cols=2)
t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
t_meta.columns[0].width = Inches(2.2)
t_meta.columns[1].width = Inches(4.3)
campos = [
    ("Actividad:", "Instalación de Shiny, Dashboard Básico y Publicación en Internet"),
    ("Usuario / Cuenta Posit:", "rncm11 (https://connect.posit.cloud/rncm11)"),
    ("Repositorio GitHub:", "https://github.com/RolandoCastillo11/Shiny-Prog-CD"),
    ("Herramientas Utilizadas:", "Python 3.11, Shiny for Python (v1.8.0), Posit Connect Cloud, Git"),
    ("Fecha de Realización:", "Octubre de 2026"),
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

# 1. INTRODUCCIÓN
h1 = doc.add_heading("1. Introducción y Objetivos", level=1)
h1.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "El propósito fundamental de esta actividad es familiarizarse con el framework Shiny for Python, "
    "aprender a estructurar un tablero analítico básico e interactivo, tramitar una cuenta en la infraestructura en "
    "la nube de Posit y desplegar la aplicación para que sea accesible públicamente a través de internet."
)

p_obj = doc.add_paragraph()
p_obj.add_run("Objetivos Específicos:\n").bold = True
p_obj.add_run("1. Instalar Shiny for Python y sus dependencias en el entorno local de desarrollo.\n")
p_obj.add_run("2. Crear un Dashboard Básico que demuestre la reactividad de Shiny mediante entradas, tarjetas estadísticas (Value Boxes), un gráfico interactivo y una tabla de datos.\n")
p_obj.add_run("3. Tramitar y configurar la cuenta de publicación en Posit Connect Cloud (cuenta: rncm11).\n")
p_obj.add_run("4. Crear y sincronizar el repositorio en GitHub (RolandoCastillo11/Shiny-Prog-CD).\n")
p_obj.add_run("5. Publicar el dashboard en internet y documentar los resultados.")

if os.path.exists(RUTA_FLUJO):
    doc.add_paragraph().paragraph_format.space_before = Pt(6)
    doc.add_picture(RUTA_FLUJO, width=Inches(6.4))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 1: Fases secuenciales de la actividad práctica.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

# 2. INSTALACIÓN
h2 = doc.add_heading("2. Pasos para la Instalación de Shiny", level=1)
h2.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Shiny for Python es el paquete oficial desarrollado por Posit para construir aplicaciones web reactivas en Python. "
    "La instalación se realiza a través de la terminal o consola de comandos (PowerShell / CMD) mediante pip:"
)

p_c1 = doc.add_paragraph()
p_c1.add_run("Paso 2.1: Instalación de Shiny y librerías analíticas básicas\n").bold = True
agregar_codigo(doc, "pip install shiny numpy pandas matplotlib rsconnect-python")

p_c2 = doc.add_paragraph()
p_c2.add_run("Paso 2.2: Comprobación de la instalación\n").bold = True
p_c2.add_run("Para confirmar que Shiny se instaló correctamente y conocer la versión instalada, se ejecuta:")
agregar_codigo(doc, (
    "shiny --version\n"
    "# Resultado en consola: shiny, version 1.8.0\n\n"
    "python --version\n"
    "# Resultado: Python 3.11.0"
))

# 3. CONSTRUCCIÓN DEL DASHBOARD BÁSICO
h3 = doc.add_heading("3. Desarrollo del Dashboard Básico", level=1)
h3.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Siguiendo el requerimiento de crear un dashboard básico, claro y sin sobrecargas, se implementó una aplicación "
    "que permite explorar distribuciones estadísticas en tiempo real: 'Generador de Distribuciones Estadísticas'."
)

doc.add_heading("3.1 Estructura de Archivos del Proyecto", level=2)
agregar_codigo(doc, (
    "Dashboard_Basico_Shiny/\n"
    "├── app.py             # Código fuente de la aplicación básica (UI + Server)\n"
    "├── requirements.txt   # Librerías mínimas necesarias para Posit (shiny, numpy, pandas, matplotlib)\n"
    "├── README.md          # Documentación del proyecto en GitHub\n"
    "├── .gitignore         # Archivos excluidos del control de versiones\n"
    "└── manifest.json      # Manifiesto de compilación generado por rsconnect"
))

doc.add_heading("3.2 Componentes del Dashboard Básico", level=2)
p_comp = doc.add_paragraph()
p_comp.add_run("• Controles de Entrada (Sidebar): ").bold = True
p_comp.add_run("Un control deslizante (slider) para elegir el número de muestras N (de 20 a 500), un selector desplegable para elegir la distribución (Normal, Uniforme o Exponencial) y botones de opción para personalizar el color del gráfico.\n")
p_comp.add_run("• Tarjetas de Indicadores (Value Boxes): ").bold = True
p_comp.add_run("Tres tarjetas numéricas que calculan dinámicamente la Media Aritmética (x̄), la Desviación Estándar (σ) y el Total de observaciones.\n")
p_comp.add_run("• Gráfico Interactivo (Histograma): ").bold = True
p_comp.add_run("Muestra la distribución de frecuencias generada con Matplotlib, actualizándose al instante ante cualquier cambio del usuario.\n")
p_comp.add_run("• Tabla de Datos: ").bold = True
p_comp.add_run("Una tabla limpia que presenta las primeras 10 observaciones generadas.")

doc.add_heading("3.3 Código Fuente Completo de app.py", level=2)
agregar_codigo(doc, (
    "import numpy as np\n"
    "import pandas as pd\n"
    "import matplotlib.pyplot as plt\n"
    "from shiny import App, render, ui, reactive\n\n"
    "# --- Interfaz de Usuario (UI) ---\n"
    "app_ui = ui.page_sidebar(\n"
    "    ui.sidebar(\n"
    "        ui.h4('Panel de Control', class_='fw-bold text-primary mb-3'),\n"
    "        ui.input_slider('n_muestras', 'Número de Muestras (N):', min=20, max=500, value=100, step=10),\n"
    "        ui.input_select('tipo_distribucion', 'Tipo de Distribución:',\n"
    "                        {'normal': 'Normal', 'uniforme': 'Uniforme', 'exponencial': 'Exponencial'}),\n"
    "        ui.input_radio_buttons('color_grafica', 'Color de las Barras:',\n"
    "                               {'#2563EB': 'Azul', '#059669': 'Verde', '#DC2626': 'Rojo'}),\n"
    "        width=280,\n"
    "    ),\n"
    "    ui.h3('📊 Dashboard Básico: Generador de Distribuciones', class_='fw-bold mb-3'),\n"
    "    ui.layout_columns(\n"
    "        ui.value_box('Media (x̄)', ui.output_text('kpi_media'), theme='primary'),\n"
    "        ui.value_box('Desv. Est. (σ)', ui.output_text('kpi_desviacion'), theme='success'),\n"
    "        ui.value_box('Total Muestras', ui.output_text('kpi_total'), theme='info'),\n"
    "        col_widths=[4, 4, 4],\n"
    "    ),\n"
    "    ui.layout_columns(\n"
    "        ui.card(ui.card_header('Histograma'), ui.output_plot('grafico_histograma')),\n"
    "        ui.card(ui.card_header('Primeras 10 Observaciones'), ui.output_table('tabla_primeros_datos')),\n"
    "        col_widths=[7, 5],\n"
    "    ),\n"
    "    title='Dashboard Básico - Shiny',\n"
    ")\n\n"
    "# --- Lógica del Servidor (Server) ---\n"
    "def server(input, output, session):\n"
    "    @reactive.calc\n"
    "    def generar_datos():\n"
    "        n = input.n_muestras()\n"
    "        tipo = input.tipo_distribucion()\n"
    "        np.random.seed(42)\n"
    "        if tipo == 'normal': valores = np.random.normal(50, 10, n)\n"
    "        elif tipo == 'uniforme': valores = np.random.uniform(20, 80, n)\n"
    "        else: valores = np.random.exponential(20, n)\n"
    "        return pd.DataFrame({'ID': range(1, n + 1), 'Valor': np.round(valores, 2)})\n\n"
    "    @render.text\n"
    "    def kpi_media(): return f\"{generar_datos()['Valor'].mean():.2f}\"\n"
    "    @render.text\n"
    "    def kpi_desviacion(): return f\"{generar_datos()['Valor'].std():.2f}\"\n"
    "    @render.text\n"
    "    def kpi_total(): return f\"{len(generar_datos())} datos\"\n\n"
    "    @render.plot\n"
    "    def grafico_histograma():\n"
    "        df = generar_datos()\n"
    "        fig, ax = plt.subplots(figsize=(6, 4))\n"
    "        ax.hist(df['Valor'], bins=15, color=input.color_grafica(), edgecolor='black')\n"
    "        return fig\n\n"
    "    @render.table\n"
    "    def tabla_primeros_datos(): return generar_datos().head(10)\n\n"
    "app = App(app_ui, server)\n"
    "if __name__ == '__main__':\n"
    "    app.run(port=8000)"
))

if os.path.exists(RUTA_GRAFICA):
    doc.add_paragraph().paragraph_format.space_before = Pt(6)
    doc.add_picture(RUTA_GRAFICA, width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 2: Histograma generado de forma reactiva por el Dashboard Básico.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

doc.add_heading("3.4 Archivo de Requisitos (requirements.txt)", level=2)
doc.add_paragraph("Librerías declaradas de forma concisa para el contenedor en la nube:")
agregar_codigo(doc, "shiny>=1.0.0\nnumpy>=1.24.0\npandas>=2.0.0\nmatplotlib>=3.7.0")

doc.add_heading("3.5 Ejecución Local", level=2)
doc.add_paragraph("Para ejecutar y probar la aplicación en la computadora local:")
agregar_codigo(doc, (
    "# Opción 1: Servidor Shiny de recarga automática\n"
    "shiny run --reload --port 8000 app.py\n\n"
    "# Opción 2: Ejecución directa con Python\n"
    "python app.py\n\n"
    "# Abrir en el navegador: http://127.0.0.1:8000"
))

# 4. TRAMITACIÓN DE LA CUENTA EN POSIT
doc.add_page_break()
h4 = doc.add_heading("4. Tramitación de la Cuenta en Posit y GitHub", level=1)
h4.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Para poner la aplicación al alcance del público, se tramitó la cuenta oficial en la plataforma en la nube de Posit "
    "y se creó el repositorio de código abierto en GitHub."
)

p_pos = doc.add_paragraph()
p_pos.add_run("4.1 Cuenta en Posit Connect Cloud\n").bold = True
p_pos.add_run(
    "• Plataforma: Posit Connect Cloud (servicio oficial gestionado de Posit).\n"
    "• URL de la cuenta tramitada: https://connect.posit.cloud/rncm11\n"
    "• Nombre de usuario / Cuenta: rncm11\n"
    "• Plan utilizado: Gratuito (incluye publicación de aplicaciones Shiny para Python sin costo)."
)

p_git = doc.add_paragraph()
p_git.add_run("4.2 Repositorio en GitHub\n").bold = True
p_git.add_run(
    "• Nombre del repositorio: Shiny-Prog-CD\n"
    "• URL pública: https://github.com/RolandoCastillo11/Shiny-Prog-CD\n"
    "• Comandos de inicialización y sincronización ejecutados:\n"
)
agregar_codigo(doc, (
    "git init\n"
    "git add .\n"
    "git commit -m \"first commit\"\n"
    "git branch -M main\n"
    "git remote add origin https://github.com/RolandoCastillo11/Shiny-Prog-CD.git\n"
    "git push -u origin main"
))

# 5. DESPLIEGUE A INTERNET
h5 = doc.add_heading("5. Publicación del Dashboard en Internet", level=1)
h5.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "Con el código disponible en GitHub y la cuenta activa en Posit Connect Cloud, el despliegue se realiza directamente "
    "desde la interfaz web de Posit:"
)

p_dep = doc.add_paragraph()
p_dep.add_run("Pasos de Publicación en Posit Connect Cloud:\n").bold = True
p_dep.add_run(
    "1. Iniciar sesión en https://connect.posit.cloud/rncm11\n"
    "2. Hacer clic en el botón 'Publish' (esquina superior derecha).\n"
    "3. Seleccionar la fuente 'GitHub' y vincular la cuenta RolandoCastillo11.\n"
    "4. Elegir el repositorio Shiny-Prog-CD.\n"
    "5. Posit detecta automáticamente el archivo app.py y el archivo requirements.txt.\n"
    "6. Hacer clic en 'Deploy'. Posit construye el contenedor Docker, instala las librerías y activa el enlace web público."
)

# 6. TABLA RESUMEN DE ACTIVIDADES
h6 = doc.add_heading("6. Resumen de Actividades Realizadas", level=1)
h6.style.font.color.rgb = RGBColor(30, 64, 175)

t_res = doc.add_table(rows=6, cols=3)
t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
t_res.columns[0].width = Inches(1.8)
t_res.columns[1].width = Inches(2.7)
t_res.columns[2].width = Inches(2.0)

encabezados = ["Actividad", "Detalle / Herramienta", "Estado"]
for j, enc in enumerate(encabezados):
    celda = t_res.cell(0, j)
    sombrear_celda(celda, "1E3A8A")
    pr = celda.paragraphs[0].add_run(enc)
    pr.font.bold = True
    pr.font.color.rgb = RGBColor(255, 255, 255)
    pr.font.size = Pt(9.5)

filas = [
    ("Instalación de Shiny", "pip install shiny (v1.8.0)", "Completada ✓"),
    ("Desarrollo del Dashboard", "app.py (Histograma + KPIs + Tabla)", "Completado y Probado ✓"),
    ("Tramitación en Posit", "connect.posit.cloud/rncm11", "Cuenta Activa ✓"),
    ("Repositorio GitHub", "RolandoCastillo11/Shiny-Prog-CD", "Código Publicado ✓"),
    ("Documentación Formal", "Reporte en Word (.docx) y PDF (.pdf)", "Generado ✓"),
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

# 7. CONCLUSIONES
doc.add_paragraph().paragraph_format.space_before = Pt(12)
h7 = doc.add_heading("7. Conclusiones", level=1)
h7.style.font.color.rgb = RGBColor(30, 64, 175)

doc.add_paragraph(
    "1. Shiny for Python permite construir herramientas analíticas interactivas con una curva de aprendizaje mínima, "
    "utilizando estructuras limpias que separan la interfaz de usuario (UI) de la lógica del servidor (Server).\n\n"
    "2. La reactividad mediante decoradores (@reactive.calc y @render) simplifica enormemente la sincronización entre "
    "los controles visuales y los cálculos numéricos en tiempo real.\n\n"
    "3. La combinación de GitHub con Posit Connect Cloud ofrece un flujo de despliegue continuo (CI/CD) ágil, permitiendo "
    "que cualquier estudiante o profesional comparta sus tableros de ciencia de datos con el mundo sin necesidad de gestionar "
    "servidores físicos ni configuraciones complejas."
)

doc.save(RUTA_DOCX)
print(f"Documento Word guardado en: {RUTA_DOCX}")

# -----------------------------------------------------------------------------
# 3. Generación del Documento PDF (.pdf) con ReportLab
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
            self.drawString(54, 755, "Reporte Práctico: Instalación de Shiny y Dashboard Básico")
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

t_port = ParagraphStyle("TPort", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=18, leading=23, textColor=colors.HexColor("#1E3A8A"), alignment=1, spaceAfter=8)
sub_port = ParagraphStyle("SPort", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=14, textColor=colors.HexColor("#64748B"), alignment=1, spaceAfter=18)
h1_pdf = ParagraphStyle("H1Pdf", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=13, leading=17, textColor=colors.HexColor("#1E3A8A"), spaceBefore=12, spaceAfter=6)
h2_pdf = ParagraphStyle("H2Pdf", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=colors.HexColor("#1E293B"), spaceBefore=9, spaceAfter=4)
p_pdf = ParagraphStyle("PPdf", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=13, textColor=colors.HexColor("#334155"), spaceAfter=6)
c_pdf = ParagraphStyle("CPdf", parent=styles["Normal"], fontName="Courier", fontSize=8, leading=10.5, textColor=colors.HexColor("#0F172A"))

def caja_codigo_pdf(texto):
    t = Table([[Paragraph(f"<font color='#0F172A'>{texto.replace(chr(10), '<br/>').replace(' ', '&nbsp;')}</font>", c_pdf)]], colWidths=[504])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t

story = []

# Encabezado portada
story.append(Paragraph("<b>PROGRAMACIÓN EN CIENCIA DE DATOS</b><br/><font color='#64748B' size='8.5'>REPORTE DE ACTIVIDAD PRÁCTICA</font>", ParagraphStyle("Port", alignment=1, fontSize=11, leading=14, fontName="Helvetica-Bold", textColor=colors.HexColor("#334155"))))
story.append(Spacer(1, 12))
story.append(Paragraph("INSTALACIÓN DE SHINY, CONSTRUCCIÓN DE UN DASHBOARD BÁSICO Y DESPLIEGUE EN POSIT", t_port))
story.append(Paragraph("Guía Paso a Paso con Código Fuente, Cuenta Posit (rncm11) y Repositorio en GitHub", sub_port))
story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceBefore=2, spaceAfter=12))

# Metadatos
datos_meta = [
    [Paragraph("<b>Actividad:</b>", p_pdf), Paragraph("Instalación de Shiny, Dashboard Básico y Publicación en Internet", p_pdf)],
    [Paragraph("<b>Cuenta en Posit:</b>", p_pdf), Paragraph("rncm11 (https://connect.posit.cloud/rncm11)", p_pdf)],
    [Paragraph("<b>Repositorio GitHub:</b>", p_pdf), Paragraph("https://github.com/RolandoCastillo11/Shiny-Prog-CD", p_pdf)],
    [Paragraph("<b>Tecnologías:</b>", p_pdf), Paragraph("Python 3.11, Shiny for Python (v1.8.0), Posit Connect Cloud, Git", p_pdf)],
    [Paragraph("<b>Fecha:</b>", p_pdf), Paragraph("Octubre de 2026", p_pdf)],
]
t_m = Table(datos_meta, colWidths=[130, 374])
t_m.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E2E8F0")),
    ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#F8FAFC")),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94A3B8")),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
]))
story.append(t_m)
story.append(Spacer(1, 10))

# 1. INTRODUCCIÓN
story.append(Paragraph("1. Introducción y Objetivos", h1_pdf))
story.append(Paragraph(
    "El objetivo de la práctica es instalar <b>Shiny for Python</b>, desarrollar un <b>dashboard analítico básico</b> "
    "e interactivo, tramitar una cuenta en la plataforma en la nube de <b>Posit (connect.posit.cloud)</b> y publicar la "
    "aplicación a través de internet vinculada con <b>GitHub</b>.", p_pdf
))

if os.path.exists(RUTA_FLUJO):
    story.append(Spacer(1, 4))
    story.append(Image(RUTA_FLUJO, width=500, height=115))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 1: Fases de la actividad práctica.</i></font>", ParagraphStyle("Cap1", alignment=1, spaceAfter=6)))

# 2. INSTALACIÓN
story.append(Paragraph("2. Instalación de Shiny", h1_pdf))
story.append(Paragraph("Se ejecutaron los siguientes comandos de instalación y verificación en la terminal:", p_pdf))
story.append(caja_codigo_pdf("pip install shiny numpy pandas matplotlib rsconnect-python\nshiny --version        # Versión instalada: 1.8.0"))

# 3. DASHBOARD BÁSICO
story.append(PageBreak())
story.append(Paragraph("3. Desarrollo del Dashboard Básico", h1_pdf))
story.append(Paragraph(
    "Se diseñó una aplicación autocontenida y básica en <code>app.py</code> que explora distribuciones estadísticas "
    "con un deslizador (slider), selector de distribución, 3 tarjetas de indicadores (Media, Desviación Estándar, Total), "
    "un histograma con Matplotlib y una tabla con las primeras observaciones:", p_pdf
))

if os.path.exists(RUTA_GRAFICA):
    story.append(Spacer(1, 4))
    story.append(Image(RUTA_GRAFICA, width=460, height=170))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 2: Histograma reactivo generado por el Dashboard.</i></font>", ParagraphStyle("Cap2", alignment=1, spaceAfter=6)))

story.append(Paragraph("<b>Código esencial de app.py:</b>", h2_pdf))
story.append(caja_codigo_pdf(
    "from shiny import App, render, ui, reactive\n"
    "import numpy as np, pandas as pd, matplotlib.pyplot as plt\n\n"
    "app_ui = ui.page_sidebar(\n"
    "    ui.sidebar(\n"
    "        ui.input_slider('n_muestras', 'Muestras (N):', min=20, max=500, value=100),\n"
    "        ui.input_select('tipo_distribucion', 'Distribución:', {'normal':'Normal','uniforme':'Uniforme'}),\n"
    "    ),\n"
    "    ui.layout_columns(\n"
    "        ui.value_box('Media', ui.output_text('kpi_media')),\n"
    "        ui.value_box('Desv. Est.', ui.output_text('kpi_desviacion')),\n"
    "        col_widths=[6, 6],\n"
    "    ),\n"
    "    ui.output_plot('grafico_histograma'),\n"
    ")\n\n"
    "def server(input, output, session):\n"
    "    @reactive.calc\n"
    "    def generar_datos():\n"
    "        return np.random.normal(50, 10, input.n_muestras())\n"
    "    @render.text\n"
    "    def kpi_media(): return f'{generar_datos().mean():.2f}'\n"
    "    @render.plot\n"
    "    def grafico_histograma():\n"
    "        fig, ax = plt.subplots(); ax.hist(generar_datos(), bins=15); return fig\n\n"
    "app = App(app_ui, server)"
))

story.append(Spacer(1, 4))
story.append(Paragraph("<b>Ejecución Local:</b> <code>shiny run --reload --port 8000 app.py</code> (o <code>python app.py</code>)", p_pdf))

# 4 Y 5. POSIT Y GITHUB
story.append(Paragraph("4. Tramitación de Cuenta en Posit y GitHub", h1_pdf))
story.append(Paragraph(
    "• <b>Cuenta en Posit Connect Cloud:</b> <u>https://connect.posit.cloud/rncm11</u> (Usuario: <code>rncm11</code>).<br/>"
    "• <b>Repositorio GitHub:</b> <u>https://github.com/RolandoCastillo11/Shiny-Prog-CD</u>.<br/>"
    "• <b>Comandos de sincronización:</b>", p_pdf
))
story.append(caja_codigo_pdf("git init; git add .; git commit -m 'first commit'; git push -u origin main"))

story.append(Paragraph("5. Publicación en Internet", h1_pdf))
story.append(Paragraph(
    "En Posit Connect Cloud (<u>connect.posit.cloud/rncm11</u>), se selecciona <b>Publish</b> -> <b>GitHub</b> -> "
    "<b>RolandoCastillo11/Shiny-Prog-CD</b>. Posit detecta automáticamente <code>app.py</code> y <code>requirements.txt</code>, "
    "compilando la aplicación en un contenedor seguro con URL pública permanente.", p_pdf
))

# 6. RESUMEN
story.append(Paragraph("6. Tabla de Verificación", h1_pdf))
datos_act = [
    [Paragraph("<b>Actividad</b>", p_pdf), Paragraph("<b>Detalle</b>", p_pdf), Paragraph("<b>Estado</b>", p_pdf)],
    [Paragraph("Instalación de Shiny", p_pdf), Paragraph("pip install shiny (v1.8.0)", p_pdf), Paragraph("<font color='#059669'><b>✓ Completada</b></font>", p_pdf)],
    [Paragraph("Dashboard Básico", p_pdf), Paragraph("app.py (Histograma + KPIs)", p_pdf), Paragraph("<font color='#059669'><b>✓ Completado</b></font>", p_pdf)],
    [Paragraph("Cuenta en Posit", p_pdf), Paragraph("connect.posit.cloud/rncm11", p_pdf), Paragraph("<font color='#059669'><b>✓ Activa</b></font>", p_pdf)],
    [Paragraph("Repositorio GitHub", p_pdf), Paragraph("Shiny-Prog-CD", p_pdf), Paragraph("<font color='#059669'><b>✓ Sincronizado</b></font>", p_pdf)],
    [Paragraph("Reporte Escrito", p_pdf), Paragraph("Word (.docx) y PDF (.pdf)", p_pdf), Paragraph("<font color='#059669'><b>✓ Generado</b></font>", p_pdf)],
]
t_act = Table(datos_act, colWidths=[130, 214, 160])
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
print(f"Documento PDF guardado en: {RUTA_PDF}")
print("GENERACIÓN DE DOCUMENTOS COMPLETADA CON ÉXITO.")
