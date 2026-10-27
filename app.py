"""
=============================================================================
DASHBOARD BÁSICO EN SHINY FOR PYTHON
=============================================================================
Práctica de Programación en Ciencia de Datos
Aplicación básica interactiva con Shiny:
  - Controles de entrada: Slider de muestras y selector de distribución.
  - Indicadores estadísticos (Value Boxes).
  - Visualización gráfica dinámica (Histograma con Matplotlib).
  - Tabla básica de datos generados en tiempo real.
=============================================================================
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from shiny import App, render, ui, reactive

# -----------------------------------------------------------------------------
# 1. Interfaz de Usuario (UI) Básica
# -----------------------------------------------------------------------------
app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.h4("Panel de Control", class_="fw-bold text-primary mb-3"),
        ui.input_slider(
            "n_muestras",
            "Número de Muestras (N):",
            min=20,
            max=500,
            value=100,
            step=10,
        ),
        ui.input_select(
            "tipo_distribucion",
            "Tipo de Distribución:",
            {
                "normal": "Normal (Gaussiana)",
                "uniforme": "Uniforme",
                "exponencial": "Exponencial",
            },
            selected="normal",
        ),
        ui.input_radio_buttons(
            "color_grafica",
            "Color de las Barras:",
            {
                "#2563EB": "Azul",
                "#059669": "Verde",
                "#DC2626": "Rojo",
                "#7C3AED": "Púrpura",
            },
            selected="#2563EB",
        ),
        ui.hr(),
        ui.p("Dashboard básico para despliegue en Posit Connect Cloud.", class_="text-muted small"),
        width=280,
    ),
    # Encabezado principal
    ui.h3("📊 Dashboard Básico: Generador de Distribuciones Estadísticas", class_="fw-bold mb-3"),
    # Indicadores Clave (Value Boxes)
    ui.layout_columns(
        ui.value_box(
            "Media Aritmética (x̄)",
            ui.output_text("kpi_media"),
            showcase=ui.div("📈", style="font-size: 2rem;"),
            theme="primary",
        ),
        ui.value_box(
            "Desviación Estándar (σ)",
            ui.output_text("kpi_desviacion"),
            showcase=ui.div("📏", style="font-size: 2rem;"),
            theme="success",
        ),
        ui.value_box(
            "Total de Observaciones",
            ui.output_text("kpi_total"),
            showcase=ui.div("🔢", style="font-size: 2rem;"),
            theme="info",
        ),
        col_widths=[4, 4, 4],
    ),
    # Gráfico y Tabla
    ui.layout_columns(
        ui.card(
            ui.card_header("Histograma de Frecuencias en Tiempo Real", class_="fw-bold"),
            ui.output_plot("grafico_histograma"),
        ),
        ui.card(
            ui.card_header("Primeras 10 Observaciones del Conjunto", class_="fw-bold"),
            ui.output_table("tabla_primeros_datos"),
        ),
        col_widths=[7, 5],
    ),
    title="Dashboard Básico - Shiny",
    fillable=True,
)


# -----------------------------------------------------------------------------
# 2. Lógica del Servidor (Server)
# -----------------------------------------------------------------------------
def server(input, output, session):

    # Cálculo reactivo: genera los datos al mover el slider o cambiar la distribución
    @reactive.calc
    def generar_datos():
        n = input.n_muestras()
        tipo = input.tipo_distribucion()
        np.random.seed(42)

        if tipo == "normal":
            valores = np.random.normal(loc=50, scale=10, size=n)
        elif tipo == "uniforme":
            valores = np.random.uniform(low=20, high=80, size=n)
        else:  # exponencial
            valores = np.random.exponential(scale=20, size=n)

        return pd.DataFrame({"ID": range(1, n + 1), "Valor": np.round(valores, 2)})

    # KPIs reactivos
    @render.text
    def kpi_media():
        df = generar_datos()
        return f"{df['Valor'].mean():.2f}"

    @render.text
    def kpi_desviacion():
        df = generar_datos()
        return f"{df['Valor'].std():.2f}"

    @render.text
    def kpi_total():
        df = generar_datos()
        return f"{len(df)} datos"

    # Gráfico interactivo reactivo
    @render.plot
    def grafico_histograma():
        df = generar_datos()
        color = input.color_grafica()
        tipo = input.tipo_distribucion()

        fig, ax = plt.subplots(figsize=(6, 4))
        ax.hist(df["Valor"], bins=15, color=color, edgecolor="black", alpha=0.75)
        ax.set_title(f"Histograma de Distribución: {tipo.capitalize()}", fontsize=12, fontweight="bold")
        ax.set_xlabel("Valor de la Variable", fontsize=10)
        ax.set_ylabel("Frecuencia Absoluta", fontsize=10)
        ax.grid(axis="y", linestyle="--", alpha=0.7)
        plt.tight_layout()
        return fig

    # Tabla reactiva básica
    @render.table
    def tabla_primeros_datos():
        df = generar_datos()
        return df.head(10)


# -----------------------------------------------------------------------------
# 3. Inicialización de la Aplicación
# -----------------------------------------------------------------------------
app = App(app_ui, server)

if __name__ == "__main__":
    app.run(port=8000)
