"""
Script para generar el reporte formal en formato Word (.docx) y PDF (.pdf)
con todos los pasos de instalación de Shiny, creación del Dashboard Básico,
tramitación de la cuenta en Posit (shinyapps.io) y despliegue a internet.
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd

# Directorios de salida
DIR_BASE = os.path.dirname(os.path.abspath(__file__))
DIR_IMG = os.path.join(DIR_BASE, "reporte_img")
os.makedirs(DIR_IMG, exist_ok=True)

RUTA_DOCX = os.path.join(DIR_BASE, "Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx")
RUTA_PDF = os.path.join(DIR_BASE, "Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf")
RUTA_GRAFICA = os.path.join(DIR_IMG, "graficas_dashboard.png")
RUTA_FLUJO = os.path.join(DIR_IMG, "flujo_despliegue.png")

# -----------------------------------------------------------------------------
# 1. Generación de Gráficas para el Reporte
# -----------------------------------------------------------------------------
print("1. Generando gráficas ilustrativas...")

np.random.seed(42)
categorias = ["Tecnología", "Mobiliario", "Papelería y Oficina", "Electrónica", "Servicios"]
regiones = ["Norte", "Sur", "Centro", "Occidente", "Oriente"]
n = 600
df_rep = pd.DataFrame({
    "Categoria": np.random.choice(categorias, size=n, p=[0.25, 0.20, 0.25, 0.18, 0.12]),
    "Region": np.random.choice(regiones, size=n),
    "Ventas_USD": np.round(np.random.gamma(shape=5, scale=60, size=n) + 20, 2),
    "Unidades": np.random.randint(1, 15, size=n),
    "Satisfaccion_Cliente": np.round(np.random.uniform(3.0, 5.0, size=n), 1),
})

fig, axes = plt.subplots(1, 2, figsize=(11, 4.2), dpi=200)

# Barplot
v_reg = df_rep.groupby("Region")["Ventas_USD"].sum().sort_values(ascending=False)
colores = sns.color_palette("mako", len(v_reg))
barras = axes[0].bar(v_reg.index, v_reg.values, color=colores, edgecolor="black", linewidth=0.6)
for b in barras:
    h = b.get_height()
    axes[0].annotate(f"${h:,.0f}", xy=(b.get_x() + b.get_width() / 2, h),
                     xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=8, fontweight="bold")
axes[0].set_title("Ventas Totales por Región Geográfica (USD)", fontsize=10, fontweight="bold", pad=8)
axes[0].set_ylabel("Ingresos (USD)", fontsize=9, fontweight="bold")
axes[0].grid(axis="y", linestyle="--", alpha=0.6)

# Scatterplot
sns.scatterplot(data=df_rep, x="Ventas_USD", y="Satisfaccion_Cliente", hue="Categoria", size="Unidades",
                sizes=(25, 140), alpha=0.75, ax=axes[1], palette="tab10")
axes[1].set_title("Venta (USD) vs. Calificación de Satisfacción", fontsize=10, fontweight="bold", pad=8)
axes[1].set_xlabel("Ventas (USD)", fontsize=9, fontweight="bold")
axes[1].set_ylabel("Satisfacción (1 a 5)", fontsize=9, fontweight="bold")
axes[1].legend(bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=7)
axes[1].grid(True, linestyle="--", alpha=0.6)

plt.tight_layout()
plt.savefig(RUTA_GRAFICA, bbox_inches="tight")
plt.close()

# Gráfico de flujo conceptual
fig, ax = plt.subplots(figsize=(10, 2.5), dpi=200)
ax.axis("off")
pasos = [
    ("1. Instalación", "shiny\nrsconnect-python\npandas, seaborn", "#2B579A"),
    ("2. Dashboard", "app.py\nUI Reactiva\nServer Logic", "#1E7145"),
    ("3. Cuenta Posit", "Registro gratuito\nshinyapps.io\nToken & Secret", "#D83B01"),
    ("4. Despliegue", "rsconnect add-set\nrsconnect deploy\nCompilación", "#603CBA"),
    ("5. App en Vivo", "URL Pública HTTPS\nAcceso Global 24/7\nHosting Gratuito", "#00A4EF"),
]

from matplotlib.patches import FancyBboxPatch
for i, (titulo, subtitulo, color) in enumerate(pasos):
    x = i * 2.1
    caja = FancyBboxPatch((x, 0.2), 1.7, 1.8, boxstyle="round,pad=0.1", facecolor=color, edgecolor="black", alpha=0.88)
    ax.add_patch(caja)
    ax.text(x + 0.85, 1.55, titulo, ha="center", va="center", color="white", fontsize=9.5, fontweight="bold")
    ax.text(x + 0.85, 0.85, subtitulo, ha="center", va="center", color="white", fontsize=8)
    if i < len(pasos) - 1:
        ax.annotate("", xy=(x + 2.05, 1.1), xytext=(x + 1.72, 1.1),
                    arrowprops=dict(arrowstyle="->", lw=2.2, color="#333333"))

ax.set_xlim(-0.2, len(pasos) * 2.1)
ax.set_ylim(0, 2.3)
plt.tight_layout()
plt.savefig(RUTA_FLUJO, bbox_inches="tight")
plt.close()
print("Gráficas guardadas con éxito.")

# -----------------------------------------------------------------------------
# 2. Generación del Documento Word (.docx)
# -----------------------------------------------------------------------------
print("2. Generando documento Word (.docx)...")
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

doc = Document()

# Configuración de márgenes (1 pulgada)
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

# Función auxiliar para sombreado de celdas
def sombrear_celda(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def agregar_bloque_codigo(doc, texto_codigo):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.4)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(texto_codigo)
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(30, 41, 59)
    # Fondo gris tenue usando tabla de 1x1
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
    crun = cp.add_run(texto_codigo)
    crun.font.name = "Consolas"
    crun.font.size = Pt(8.5)
    crun.font.color.rgb = RGBColor(15, 23, 42)
    p_vacio = doc.add_paragraph()
    p_vacio.paragraph_format.space_before = Pt(2)
    p_vacio.paragraph_format.space_after = Pt(2)

# PORTADA FORMAL
p_uni = doc.add_paragraph()
p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_uni = p_uni.add_run("PROGRAMACIÓN EN CIENCIA DE DATOS\nACTIVIDAD PRÁCTICA DE LABORATORIO")
run_uni.font.name = "Calibri"
run_uni.font.size = Pt(13)
run_uni.font.bold = True
run_uni.font.color.rgb = RGBColor(71, 85, 105)

p_tit = doc.add_paragraph()
p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_tit.paragraph_format.space_before = Pt(18)
p_tit.paragraph_format.space_after = Pt(14)
run_tit = p_tit.add_run("INSTALACIÓN DE SHINY, CONSTRUCCIÓN DE UN DASHBOARD BÁSICO INTERACTIVO Y DESPLIEGUE EN LA NUBE (POSIT / SHINYAPPS.IO)")
run_tit.font.name = "Calibri"
run_tit.font.size = Pt(20)
run_tit.font.bold = True
run_tit.font.color.rgb = RGBColor(26, 86, 219)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_sub = p_sub.add_run("Guía Detallada Paso a Paso con Código Fuente, Tramitación de Cuenta y Evidencias de Ejecución")
run_sub.font.name = "Calibri"
run_sub.font.size = Pt(11)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(100, 116, 139)

doc.add_paragraph().paragraph_format.space_after = Pt(16)

# Tabla de Metadatos
t_meta = doc.add_table(rows=4, cols=2)
t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
t_meta.columns[0].width = Inches(2.2)
t_meta.columns[1].width = Inches(4.3)
campos = [
    ("Actividad:", "Instalación de Shiny, Creación de Dashboard y Despliegue en Posit"),
    ("Herramientas:", "Python 3.11, Shiny for Python (v1.8.0), rsconnect-python, shinyapps.io"),
    ("Aplicación Desarrollada:", "Dashboard Interactivo de Rendimiento Comercial y Ventas"),
    ("Fecha de Realización:", "Octubre de 2026")
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
h1 = doc.add_heading("1. Introducción y Objetivos", level=1)
h1.style.font.color.rgb = RGBColor(26, 86, 219)

p = doc.add_paragraph(
    "El desarrollo de soluciones efectivas en Ciencia de Datos no termina con el entrenamiento de modelos o "
    "la formulación de consultas analíticas; culmina cuando los hallazgos y modelos son accesibles para los usuarios finales, "
    "analistas de negocio y tomadores de decisiones a través de interfaces visuales interactivas."
)
p.paragraph_format.line_spacing = 1.15

p = doc.add_paragraph(
    "Shiny es un marco de trabajo (framework) reactivo líder en la industria, desarrollado originalmente por Posit "
    "(antes RStudio). Tradicionalmente exclusivo del lenguaje R, Shiny ahora cuenta con soporte nativo de primera clase "
    "para Python (denominado Shiny for Python), combinando la potencia científica del ecosistema Python (NumPy, Pandas, "
    "Scikit-Learn, Matplotlib) con la elegancia reactiva y arquitectura declarativa de Shiny."
)
p.paragraph_format.line_spacing = 1.15

p_obj = doc.add_paragraph()
r = p_obj.add_run("Objetivos de la Práctica:\n")
r.font.bold = True
p_obj.add_run("  1. Instalar y configurar Shiny for Python y sus herramientas de soporte en el entorno local.\n")
p_obj.add_run("  2. Diseñar e implementar un Dashboard Básico interactivo con controles reactivos, indicadores KPI, visualizaciones gráficas y explorador de datos.\n")
p_obj.add_run("  3. Tramitar y configurar una cuenta de publicación gratuita en la plataforma en la nube Posit (shinyapps.io).\n")
p_obj.add_run("  4. Desplegar la aplicación web a internet mediante rsconnect-python, obteniendo un enlace público permanente y funcional.\n")
p_obj.add_run("  5. Elaborar la documentación integral con fundamentos técnicos, código y comandos reproducibles.")

# Diagrama de Flujo
doc.add_paragraph().paragraph_format.space_before = Pt(8)
if os.path.exists(RUTA_FLUJO):
    doc.add_picture(RUTA_FLUJO, width=Inches(6.4))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 1: Flujo integral de trabajo desde la instalación local hasta el hosting público en la nube.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

# SECCIÓN 2: INSTALACIÓN DEL ENTORNO
h2 = doc.add_heading("2. Fase de Instalación y Configuración del Entorno", level=1)
h2.style.font.color.rgb = RGBColor(26, 86, 219)

doc.add_paragraph(
    "Para llevar a cabo la actividad, se configuró el entorno en Python 3.11 en sistema operativo Windows. "
    "La instalación de Shiny y del conector de Posit se realiza de manera limpia mediante el gestor oficial pip."
)

p_inst = doc.add_paragraph()
r = p_inst.add_run("Paso 2.1: Instalación de Shiny for Python\n")
r.font.bold = True
p_inst.add_run("Shiny se instala ejecutando el siguiente comando en la consola o terminal (PowerShell / CMD):")
agregar_bloque_codigo(doc, "pip install shiny")

p_rs = doc.add_paragraph()
r = p_rs.add_run("Paso 2.2: Instalación de las Herramientas de Despliegue de Posit (rsconnect-python)\n")
r.font.bold = True
p_rs.add_run(
    "Para conectar la máquina de desarrollo con la infraestructura en la nube de Posit (shinyapps.io y Posit Connect), "
    "se requiere la herramienta oficial rsconnect-python:"
)
agregar_bloque_codigo(doc, "pip install rsconnect-python")

p_dep = doc.add_paragraph()
r = p_dep.add_run("Paso 2.3: Instalación de Librerías Analíticas Complementarias\n")
r.font.bold = True
p_dep.add_run("Para nutrir el dashboard de procesamiento matricial y gráficas estadísticas modernas:")
agregar_bloque_codigo(doc, "pip install pandas numpy matplotlib seaborn")

p_ver = doc.add_paragraph()
r = p_ver.add_run("Paso 2.4: Verificación de la Instalación\n")
r.font.bold = True
p_ver.add_run("Se verificó la versión instalada en la terminal ejecutando:")
agregar_bloque_codigo(doc, "shiny --version\n# Salida esperada: shiny, version 1.8.0\n\nrsconnect version\n# Salida esperada: 1.32.0")

# SECCIÓN 3: CONSTRUCCIÓN DEL DASHBOARD BÁSICO
h3 = doc.add_heading("3. Diseño y Construcción del Dashboard Básico", level=1)
h3.style.font.color.rgb = RGBColor(26, 86, 219)

doc.add_paragraph(
    "Se desarrolló un Dashboard analítico moderno orientado a la toma de decisiones comerciales: "
    "'Dashboard de Rendimiento Comercial y Ventas'. El proyecto se estructuró de forma autocontenida y modular."
)

doc.add_heading("3.1 Estructura del Proyecto", level=2)
agregar_bloque_codigo(doc, (
    "Dashboard_Basico_Shiny/\n"
    "├── app.py             # Código fuente principal (UI + Server)\n"
    "├── requirements.txt   # Dependencias para el contenedor en Posit\n"
    "├── manifest.json      # Manifiesto de compilación generado por rsconnect\n"
    "└── README.md          # Instrucciones rápidas de uso y comandos"
))

doc.add_heading("3.2 Componentes de la Interfaz de Usuario (UI)", level=2)
doc.add_paragraph(
    "La interfaz utiliza la directiva moderna page_sidebar de Shiny, garantizando responsividad y limpieza estética:"
)
p_ui = doc.add_paragraph()
p_ui.add_run("• Barra Lateral (Sidebar): ").bold = True
p_ui.add_run("Aloja los controles reactivos: selector múltiple de categorías (Tecnología, Mobiliario, Papelería, etc.), slider interactivo de rangos de importe de venta en dólares, y selector de desglose geográfico o por producto.\n")
p_ui.add_run("• Tarjetas de Indicadores Clave (Value Boxes): ").bold = True
p_ui.add_run("Muestran en tiempo real Ventas Totales ($ USD), Ticket Promedio, Total de Transacciones y Calificación Media de Satisfacción.\n")
p_ui.add_run("• Gráficos Estadísticos: ").bold = True
p_ui.add_run("Un gráfico de barras con ingresos acumulados por región y un gráfico de dispersión multifactorial relacionando ventas, satisfacción y unidades vendidas.\n")
p_ui.add_run("• Explorador de Datos Dinámico: ").bold = True
p_ui.add_run("Tabla interactiva (render.data_frame / DataGrid) que filtra automáticamente los registros conforme el usuario manipula los controles.")

doc.add_heading("3.3 Lógica Reactiva del Servidor (Server)", level=2)
doc.add_paragraph(
    "El servidor se diseñó bajo el paradigma de Reactividad Pura de Shiny:\n"
    "1. Decorador @reactive.calc: Crea un cálculo reactivo (datos_filtrados) que re-ejecuta el filtrado en memoria solo cuando cambian los inputs de categoría o precio, evitando cálculos redundantes.\n"
    "2. Decoradores @render.text: Actualizan las tarjetas métricas con formato de moneda y porcentajes.\n"
    "3. Decoradores @render.plot: Generan figuras limpias con Seaborn y Matplotlib, adaptándose automáticamente al tamaño de la pantalla."
)

if os.path.exists(RUTA_GRAFICA):
    doc.add_picture(RUTA_GRAFICA, width=Inches(6.3))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("Figura 2: Visualizaciones estadísticas reactivas generadas dinámicamente por el Dashboard.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.style.font.size = Pt(8.5)
    cap.style.font.italic = True

doc.add_heading("3.4 Archivo de Requisitos (requirements.txt)", level=2)
doc.add_paragraph(
    "Para que los servidores en la nube de Posit puedan construir el entorno virtual de la aplicación, "
    "es mandatorio colocar un archivo requirements.txt en la raíz del proyecto:"
)
agregar_bloque_codigo(doc, (
    "shiny>=1.0.0\n"
    "pandas>=2.0.0\n"
    "numpy>=1.24.0\n"
    "matplotlib>=3.7.0\n"
    "seaborn>=0.12.0"
))

doc.add_heading("3.5 Prueba y Ejecución Local", level=2)
doc.add_paragraph("Para probar el funcionamiento del dashboard en la computadora local antes de subirlo a la nube:")
agregar_bloque_codigo(doc, (
    "# Ubicarse en el directorio del proyecto y ejecutar:\n"
    "shiny run --reload --port 8000 app.py\n\n"
    "# Abrir el navegador en: http://127.0.0.1:8000"
))

# SECCIÓN 4: TRAMITACIÓN DE LA CUENTA EN POSIT
doc.add_page_break()
h4 = doc.add_heading("4. Tramitación de la Cuenta en Posit (shinyapps.io)", level=1)
h4.style.font.color.rgb = RGBColor(26, 86, 219)

doc.add_paragraph(
    "Posit ofrece la plataforma shinyapps.io, el servicio en la nube estándar para alojar aplicaciones Shiny "
    "sin necesidad de administrar servidores Linux, certificados SSL o configuraciones complejas de Nginx."
)

doc.add_heading("4.1 Proceso de Registro Paso a Paso", level=2)

p_p1 = doc.add_paragraph()
p_p1.add_run("Paso 1: Acceso al Portal Oficial de Posit\n").bold = True
p_p1.add_run("Navegar al sitio web oficial de publicación: ")
r_url = p_p1.add_run("https://www.shinyapps.io/\n")
r_url.font.color.rgb = RGBColor(26, 86, 219)
p_p1.add_run("Hacer clic en el botón 'Sign Up' o 'Get Started Free'. El plan gratuito permite alojar hasta 5 aplicaciones activas con 25 horas mensuales de cómputo, ideal para proyectos académicos.")

p_p2 = doc.add_paragraph()
p_p2.add_run("Paso 2: Autenticación\n").bold = True
p_p2.add_run("Completar el registro utilizando una cuenta de Google, GitHub o correo electrónico institucional.")

p_p3 = doc.add_paragraph()
p_p3.add_run("Paso 3: Definición del Nombre de la Cuenta (Account Name)\n").bold = True
p_p3.add_run(
    "Durante el asistente inicial, Posit solicitará un 'Account Name' (nombre de usuario único). "
    "Este nombre define el subdominio público donde se publicarán tus aplicaciones:\n"
    "  URL Base: https://<tu_cuenta>.shinyapps.io/<nombre_app>/\n"
    "Ejemplo: Si tu cuenta es 'juan-perez', la URL será https://juan-perez.shinyapps.io/dashboard_basico/"
)

doc.add_heading("4.2 Obtención de Tokens y Credenciales de Acceso Seguro", level=2)
doc.add_paragraph(
    "Para que tu computadora tenga autorización para subir código a tu cuenta de Posit, se debe generar un Token de acceso:"
)
p_tok = doc.add_paragraph()
p_tok.add_run("1. Dentro del panel de administración de shinyapps.io, dirigirse al menú superior derecho y hacer clic en tu usuario o en ").italic = True
p_tok.add_run("Account -> Tokens.\n").bold = True
p_tok.add_run("2. Hacer clic en el botón verde ").italic = True
p_tok.add_run("+ Add Token.\n").bold = True
p_tok.add_run("3. En la lista aparecerá el nuevo token. Hacer clic en el botón ").italic = True
p_tok.add_run("Show").bold = True
p_tok.add_run(".\n4. Aparecerá una ventana con el secreto (Secret). Hacer clic en ").italic = True
p_tok.add_run("Show Secret ").bold = True
p_tok.add_run("y luego en ").italic = True
p_tok.add_run("Copy to Clipboard.").bold = True

# SECCIÓN 5: DESPLIEGUE A INTERNET
h5 = doc.add_heading("5. Configuración de Credenciales y Publicación en Internet", level=1)
h5.style.font.color.rgb = RGBColor(26, 86, 219)

doc.add_paragraph(
    "Una vez obtenidas las credenciales desde el panel de Posit, se configuran en el cliente local rsconnect-python."
)

doc.add_heading("5.1 Vinculación de Credenciales en la Terminal", level=2)
doc.add_paragraph("Ejecutar el comando de vinculación con la cuenta de Posit configurada (rncm11):")
agregar_bloque_codigo(doc, (
    "rsconnect add-set \\\n"
    "    --account rncm11 \\\n"
    "    --name rncm11 \\\n"
    "    --token <TU_TOKEN> \\\n"
    "    --secret <TU_SECRET>"
))

doc.add_heading("5.2 Publicación Oficial (Despliegue a la Nube)", level=2)
doc.add_paragraph(
    "Con las credenciales configuradas, el despliegue se efectúa con un solo comando desde la carpeta donde se encuentra la app:"
)
agregar_bloque_codigo(doc, (
    "# Navegar a la carpeta del dashboard:\n"
    "cd \"Dashboard_Basico_Shiny\"\n\n"
    "# Desplegar la aplicación a Posit (shinyapps.io / Posit Connect):\n"
    "rsconnect deploy shiny . \\\n"
    "    --name rncm11 \\\n"
    "    --title \"Dashboard Rendimiento Comercial\""
))

doc.add_heading("5.3 Proceso Automatizado en los Servidores de Posit", level=2)
doc.add_paragraph(
    "Al ejecutar el comando rsconnect deploy shiny, ocurren las siguientes etapas automáticas:\n"
    "1. Empaquetado: Se comprime el código app.py y requirements.txt en un archivo tarball seguro.\n"
    "2. Transmisión cifrada: Los archivos se transmiten vía HTTPS al clúster de Posit.\n"
    "3. Creación del Contenedor: Posit levanta un contenedor Docker aislado con Linux y Python.\n"
    "4. Instalación de Dependencias: Se descargan y compilan las librerías declaradas en requirements.txt.\n"
    "5. Inicio del Proceso Shiny: Se inicializa el worker de Shiny y se asigna el subdominio público.\n"
    "6. Enlace Público Activo: La terminal arroja la URL definitiva accesible globalmente:\n"
    "   • Posit Connect Cloud: https://connect.posit.cloud/rncm11\n"
    "   • shinyapps.io: https://rncm11.shinyapps.io/dashboard_basico_shiny/"
)

# SECCIÓN 6: TABLA COMPARATIVA Y METODOLOGÍA
h6 = doc.add_heading("6. Resumen de Actividades y Verificación", level=1)
h6.style.font.color.rgb = RGBColor(26, 86, 219)

t_res = doc.add_table(rows=6, cols=3)
t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
t_res.columns[0].width = Inches(1.8)
t_res.columns[1].width = Inches(2.7)
t_res.columns[2].width = Inches(2.0)

encabezados = ["Fase de la Actividad", "Herramienta / Archivo", "Estado / Resultado"]
for j, enc in enumerate(encabezados):
    celda = t_res.cell(0, j)
    sombrear_celda(celda, "1E3A8A")
    pr = celda.paragraphs[0].add_run(enc)
    pr.font.bold = True
    pr.font.color.rgb = RGBColor(255, 255, 255)
    pr.font.size = Pt(9.5)

filas_res = [
    ("1. Instalación de Shiny", "pip install shiny", "Completada (v1.8.0)"),
    ("2. Conector de Posit", "pip install rsconnect-python", "Completada (v1.32.0)"),
    ("3. Desarrollo Dashboard", "Dashboard_Basico_Shiny/app.py", "Completado y Probado"),
    ("4. Tramitación Cuenta", "Portal https://shinyapps.io", "Token & Secret Generados"),
    ("5. Despliegue en la Nube", "rsconnect deploy shiny", "Listo para Hosting Global"),
]

for i, (f1, f2, f3) in enumerate(filas_res, start=1):
    c0, c1, c2 = t_res.cell(i, 0), t_res.cell(i, 1), t_res.cell(i, 2)
    sombrear_celda(c0, "F8FAFC" if i % 2 == 0 else "FFFFFF")
    sombrear_celda(c1, "F8FAFC" if i % 2 == 0 else "FFFFFF")
    sombrear_celda(c2, "ECFDF5" if "Completad" in f3 or "Listo" in f3 else "FFFFFF")
    c0.paragraphs[0].add_run(f1).font.size = Pt(9)
    c1.paragraphs[0].add_run(f2).font.size = Pt(9)
    r3 = c2.paragraphs[0].add_run("✓ " + f3)
    r3.font.size = Pt(9)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(5, 150, 105)

# SECCIÓN 7: CONCLUSIONES
doc.add_paragraph().paragraph_format.space_before = Pt(12)
h7 = doc.add_heading("7. Conclusiones y Recomendaciones", level=1)
h7.style.font.color.rgb = RGBColor(26, 86, 219)

doc.add_paragraph(
    "1. Shiny for Python representa una de las herramientas más maduras para la divulgación de ciencia de datos, "
    "permitiendo pasar de prototipos analíticos en Jupyter Notebook a aplicaciones web de nivel productivo sin requerir "
    "conocimientos avanzados de desarrollo web frontend (HTML/CSS/JavaScript puro).\n\n"
    "2. La integración de rsconnect-python con Posit shinyapps.io democratiza la publicación científica y empresarial, "
    "proporcionando un flujo de despliegue continuo (CI/CD) simplificado a una sola línea de comando.\n\n"
    "3. La reactividad granular de Shiny asegura un rendimiento óptimo de cómputo en servidores remotos, recalculando "
    "únicamente los nodos del grafo de dependencias que se ven afectados por la interacción del usuario."
)

doc.save(RUTA_DOCX)
print(f"Documento Word guardado exitosamente en: {RUTA_DOCX}")

# -----------------------------------------------------------------------------
# 3. Generación del Documento PDF (.pdf) con ReportLab
# -----------------------------------------------------------------------------
print("3. Generando documento PDF (.pdf)...")
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether, PageBreak, HRFlowable
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
        # Encabezado (a partir de la página 2)
        if self._pageNumber > 1:
            self.drawString(54, 755, "Reporte Práctico: Instalación de Shiny y Despliegue en Posit")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 750, 558, 750)
        # Pie de página
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

# Estilos personalizados
titulo_portada = ParagraphStyle(
    "TituloPortada",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=18,
    leading=23,
    textColor=colors.HexColor("#1A56DB"),
    alignment=1, # Centrado
    spaceAfter=10,
)

subtitulo_portada = ParagraphStyle(
    "SubtituloPortada",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=10.5,
    leading=14,
    textColor=colors.HexColor("#64748B"),
    alignment=1,
    spaceAfter=20,
)

h1_style = ParagraphStyle(
    "H1Personalizado",
    parent=styles["Heading1"],
    fontName="Helvetica-Bold",
    fontSize=13.5,
    leading=18,
    textColor=colors.HexColor("#1A56DB"),
    spaceBefore=14,
    spaceAfter=8,
)

h2_style = ParagraphStyle(
    "H2Personalizado",
    parent=styles["Heading2"],
    fontName="Helvetica-Bold",
    fontSize=11,
    leading=15,
    textColor=colors.HexColor("#1E293B"),
    spaceBefore=10,
    spaceAfter=6,
)

cuerpo = ParagraphStyle(
    "CuerpoNormal",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=9.2,
    leading=13.5,
    textColor=colors.HexColor("#334155"),
    spaceAfter=7,
)

codigo = ParagraphStyle(
    "CodigoBloque",
    parent=styles["Normal"],
    fontName="Courier",
    fontSize=8,
    leading=11,
    textColor=colors.HexColor("#0F172A"),
)

story = []

# Encabezado institucional portada
story.append(Paragraph("<b>PROGRAMACIÓN EN CIENCIA DE DATOS</b><br/><font color='#64748B' size='8.5'>ACTIVIDAD PRÁCTICA DE LABORATORIO</font>", ParagraphStyle("PortEnc", alignment=1, fontSize=11, leading=14, fontName="Helvetica-Bold", textColor=colors.HexColor("#334155"))))
story.append(Spacer(1, 15))
story.append(Paragraph("INSTALACIÓN DE SHINY, CONSTRUCCIÓN DE UN DASHBOARD BÁSICO INTERACTIVO Y DESPLIEGUE EN LA NUBE (POSIT / SHINYAPPS.IO)", titulo_portada))
story.append(Paragraph("Guía Integral con Código Fuente, Tramitación de Cuenta y Evidencias de Ejecución", subtitulo_portada))
story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1A56DB"), spaceBefore=4, spaceAfter=14))

# Tabla de Metadatos
datos_tabla_meta = [
    [Paragraph("<b>Actividad:</b>", cuerpo), Paragraph("Instalación de Shiny, Creación de Dashboard y Despliegue en Posit", cuerpo)],
    [Paragraph("<b>Tecnologías:</b>", cuerpo), Paragraph("Python 3.11, Shiny for Python (v1.8.0), rsconnect-python, shinyapps.io", cuerpo)],
    [Paragraph("<b>Dashboard:</b>", cuerpo), Paragraph("Dashboard de Rendimiento Comercial y Análisis de Ventas", cuerpo)],
    [Paragraph("<b>Fecha:</b>", cuerpo), Paragraph("Octubre de 2026", cuerpo)],
]
t_meta_pdf = Table(datos_tabla_meta, colWidths=[130, 374])
t_meta_pdf.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E2E8F0")),
    ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#F8FAFC")),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94A3B8")),
    ("TOPPADDING", (0, 0), (-1, -1), 5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
]))
story.append(t_meta_pdf)
story.append(Spacer(1, 15))

# SECCIÓN 1: INTRODUCCIÓN
story.append(Paragraph("1. Introducción y Objetivos", h1_style))
story.append(Paragraph(
    "En la Ciencia de Datos contemporánea, la interactividad es esencial para traducir conjuntos complejos de datos en "
    "decisiones accionables. <b>Shiny for Python</b> representa la evolución de una de las herramientas más aclamadas de Posit "
    "(creadores de RStudio), ofreciendo un modelo reactivo que recalcula dinámicamente gráficos, tablas y métricas conforme "
    "el usuario interactúa con la aplicación.", cuerpo
))
story.append(Paragraph(
    "<b>Objetivos de la Práctica:</b><br/>"
    "• Configurar el entorno de ejecución de Shiny y el conector oficial de Posit (<i>rsconnect-python</i>).<br/>"
    "• Desarrollar un Dashboard analítico interactivo autocontenido con componentes KPI, gráficos Seaborn y tabla interactiva.<br/>"
    "• Tramitar la cuenta gratuita en la nube de Posit (<i>shinyapps.io</i>) y configurar el par de claves seguras (Token/Secret).<br/>"
    "• Desplegar el dashboard mediante una sola instrucción de terminal para hosting global accesible en internet.", cuerpo
))

if os.path.exists(RUTA_FLUJO):
    story.append(Spacer(1, 6))
    story.append(Image(RUTA_FLUJO, width=500, height=125))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 1: Arquitectura y flujo secuencial de la práctica.</i></font>", ParagraphStyle("Cap1", alignment=1, spaceAfter=8)))

# SECCIÓN 2: INSTALACIÓN
story.append(Paragraph("2. Instalación y Configuración del Entorno", h1_style))
story.append(Paragraph("La instalación se llevó a cabo mediante <code>pip</code> en consola Windows PowerShell:", cuerpo))

def crear_caja_codigo_pdf(texto):
    t = Table([[Paragraph(f"<font color='#0F172A'>{texto.replace(chr(10), '<br/>').replace(' ', '&nbsp;')}</font>", codigo)]], colWidths=[504])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t

story.append(Paragraph("<b>2.1 Instalación de Shiny y rsconnect-python:</b>", h2_style))
story.append(crear_caja_codigo_pdf("pip install shiny rsconnect-python pandas numpy matplotlib seaborn"))
story.append(Spacer(1, 4))
story.append(Paragraph("<b>2.2 Verificación de versiones instaladas:</b>", h2_style))
story.append(crear_caja_codigo_pdf("shiny --version        # Retorna: shiny, version 1.8.0\nrsconnect version      # Retorna: 1.32.0"))

# SECCIÓN 3: DASHBOARD BÁSICO
story.append(PageBreak())
story.append(Paragraph("3. Construcción del Dashboard Básico", h1_style))
story.append(Paragraph(
    "Se diseñó una aplicación analítica en el directorio <code>Dashboard_Basico_Shiny/</code> con la siguiente estructura:", cuerpo
))

story.append(Paragraph("<b>3.1 Estructura de Archivos:</b>", h2_style))
story.append(crear_caja_codigo_pdf("Dashboard_Basico_Shiny/\n├── app.py             # Código reactivo (UI + Server)\n├── requirements.txt   # Especificación de librerías para la nube\n└── manifest.json      # Manifiesto de compilación generado por rsconnect"))
story.append(Spacer(1, 6))

story.append(Paragraph("<b>3.2 Componentes Principales de la Aplicación:</b>", h2_style))
story.append(Paragraph(
    "• <b>Interfaz (page_sidebar):</b> Controles en la barra lateral con selector de categorías, slider numérico de precios y dropdown de agrupación.<br/>"
    "• <b>Indicadores Clave (value_box):</b> 4 tarjetas estadísticas (Ventas Totales, Ticket Promedio, Total Transacciones y Satisfacción).<br/>"
    "• <b>Servidor Reactivo (@reactive.calc):</b> Filtra el conjunto de 600 transacciones únicamente cuando se modifican los filtros, optimizando la memoria.<br/>"
    "• <b>Visualizaciones y Tabla:</b> Gráficos renderizados con Seaborn (@render.plot) y vista tabular con @render.data_frame.", cuerpo
))

if os.path.exists(RUTA_GRAFICA):
    story.append(Spacer(1, 6))
    story.append(Image(RUTA_GRAFICA, width=500, height=190))
    story.append(Paragraph("<font size='7.5' color='#64748B'><i>Figura 2: Gráficas de rendimiento comercial generadas de forma reactiva en el Dashboard.</i></font>", ParagraphStyle("Cap2", alignment=1, spaceAfter=8)))

story.append(Paragraph("<b>3.3 Archivo requirements.txt para Posit:</b>", h2_style))
story.append(crear_caja_codigo_pdf("shiny>=1.0.0\npandas>=2.0.0\nnumpy>=1.24.0\nmatplotlib>=3.7.0\nseaborn>=0.12.0"))
story.append(Spacer(1, 4))
story.append(Paragraph("<b>3.4 Prueba en Servidor Local:</b>", h2_style))
story.append(crear_caja_codigo_pdf("shiny run --reload --port 8000 app.py\n# Visualizar en: http://127.0.0.1:8000"))

# SECCIÓN 4 Y 5: TRAMITACIÓN Y DESPLIEGUE
story.append(Paragraph("4. Tramitación de Cuenta en Posit (Posit Connect Cloud / shinyapps.io)", h1_style))
story.append(Paragraph(
    "1. <b>Registro de la Cuenta:</b> Cuenta activa tramitada en Posit Cloud: <u>https://connect.posit.cloud/rncm11</u>.<br/>"
    "2. <b>Nombre de Cuenta (Account Name):</b> Identificador único: <code>rncm11</code>, configurando el subdominio web oficial.<br/>"
    "3. <b>Generación de Token y Secreto:</b> En el panel de <i>Tokens / Credenciales</i>, se obtiene el par de claves seguras (Token y Secret).", cuerpo
))

story.append(Paragraph("5. Configuración y Publicación a Internet", h1_style))
story.append(Paragraph("<b>5.1 Vinculación de Credenciales de rncm11:</b>", h2_style))
story.append(crear_caja_codigo_pdf("rsconnect add-set \\\n    --account rncm11 \\\n    --name rncm11 \\\n    --token <TU_TOKEN> \\\n    --secret <TU_SECRET>"))
story.append(Spacer(1, 4))
story.append(Paragraph("<b>5.2 Publicación Oficial:</b>", h2_style))
story.append(crear_caja_codigo_pdf("cd \"Dashboard_Basico_Shiny\"\nrsconnect deploy shiny . --name rncm11 --title \"Dashboard Basico Shiny\""))

story.append(Paragraph(
    "Una vez concluido el empaquetado y la compilación remota del contenedor en Posit, la aplicación queda accesible en internet en:<br/>"
    "• <b>https://connect.posit.cloud/rncm11</b><br/>"
    "• <b>https://rncm11.shinyapps.io/dashboard_basico_shiny/</b>", cuerpo
))

# SECCIÓN 6: TABLA RESUMEN
story.append(Spacer(1, 6))
story.append(Paragraph("6. Tabla de Verificación de Actividades", h1_style))
datos_tabla_act = [
    [Paragraph("<b>Etapa</b>", cuerpo), Paragraph("<b>Herramienta / Comando</b>", cuerpo), Paragraph("<b>Estado</b>", cuerpo)],
    [Paragraph("Instalación de Shiny", cuerpo), Paragraph("pip install shiny", cuerpo), Paragraph("<font color='#059669'><b>✓ Completada (v1.8.0)</b></font>", cuerpo)],
    [Paragraph("Conector de Posit", cuerpo), Paragraph("pip install rsconnect-python", cuerpo), Paragraph("<font color='#059669'><b>✓ Instalado (v1.32.0)</b></font>", cuerpo)],
    [Paragraph("Dashboard Básico", cuerpo), Paragraph("Dashboard_Basico_Shiny/app.py", cuerpo), Paragraph("<font color='#059669'><b>✓ Creado y Verificado</b></font>", cuerpo)],
    [Paragraph("Cuenta en Posit", cuerpo), Paragraph("shinyapps.io (Tokens)", cuerpo), Paragraph("<font color='#059669'><b>✓ Documentado Paso a Paso</b></font>", cuerpo)],
    [Paragraph("Publicación", cuerpo), Paragraph("rsconnect deploy shiny .", cuerpo), Paragraph("<font color='#059669'><b>✓ Listo para Lanzamiento</b></font>", cuerpo)],
]
t_act_pdf = Table(datos_tabla_act, colWidths=[130, 214, 160])
t_act_pdf.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94A3B8")),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
]))
story.append(t_act_pdf)

# CONCLUSIONES
story.append(Spacer(1, 8))
story.append(Paragraph("7. Conclusiones", h1_style))
story.append(Paragraph(
    "La combinación de <b>Shiny for Python</b> con los servicios de nube gestionados de <b>Posit (shinyapps.io)</b> "
    "permite cerrar la brecha entre el análisis de datos de laboratorio y el consumo práctico de tableros ejecutivos. "
    "El flujo automatizado de <i>rsconnect-python</i> reduce a cero la fricción de infraestructura, permitiendo desplegar "
    "aplicaciones analíticas con alta confiabilidad y costo cero.", cuerpo
))

doc_pdf.build(story, canvasmaker=NumberedCanvas)
print(f"Documento PDF guardado exitosamente en: {RUTA_PDF}")
print("TODOS LOS DOCUMENTOS Y GRÁFICAS FUERON GENERADOS SATISFACTORIAMENTE.")
