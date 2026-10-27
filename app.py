"""
=============================================================================
DASHBOARD BÁSICO EN SHINY FOR PYTHON: ANÁLISIS DE VENTAS Y RENDIMIENTO COMERCIAL
=============================================================================
Este dashboard interactivo permite:
  1. Filtrar registros por categoría comercial y rango de importe de ventas.
  2. Consultar tarjetas de indicadores clave (KPI Value Boxes).
  3. Visualizar gráficos estadísticos dinámicos (barras y dispersión).
  4. Explorar la tabla de datos interactiva filtrada en tiempo real.
=============================================================================
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from shiny import App, render, ui, reactive

# Configuración estética de Seaborn y Matplotlib
sns.set_theme(style="whitegrid", palette="deep")
plt.rcParams.update({"font.sans-serif": "DejaVu Sans", "font.size": 10})

# -----------------------------------------------------------------------------
# Generación de Datos de Ejemplo (Dataset Autocontenido y Realista)
# -----------------------------------------------------------------------------
np.random.seed(42)
n_registros = 600

categorias = ["Tecnología", "Mobiliario", "Papelería y Oficina", "Electrónica", "Servicios"]
regiones = ["Norte", "Sur", "Centro", "Occidente", "Oriente"]

fechas = pd.date_range(start="2024-01-01", periods=180, freq="D")

datos = pd.DataFrame({
    "ID_Transaccion": [f"TRX-{1000 + i}" for i in range(n_registros)],
    "Fecha": np.random.choice(fechas, size=n_registros),
    "Categoria": np.random.choice(categorias, size=n_registros, p=[0.25, 0.20, 0.25, 0.18, 0.12]),
    "Region": np.random.choice(regiones, size=n_registros),
    "Ventas_USD": np.round(np.random.gamma(shape=5, scale=60, size=n_registros) + 20, 2),
    "Unidades": np.random.randint(1, 15, size=n_registros),
    "Satisfaccion_Cliente": np.round(np.random.uniform(3.0, 5.0, size=n_registros), 1),
})

# -----------------------------------------------------------------------------
# Interfaz de Usuario (UI) con diseño moderno y reactivo
# -----------------------------------------------------------------------------
app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.h4("Panel de Filtros", class_="fw-bold text-primary mb-3"),
        ui.p("Ajusta los parámetros para actualizar los KPIs y visualizaciones:", class_="text-muted small"),
        ui.hr(),
        ui.input_checkbox_group(
            "filtro_categorias",
            "Seleccionar Categorías:",
            choices={cat: cat for cat in categorias},
            selected=categorias,
        ),
        ui.hr(),
        ui.input_slider(
            "filtro_precio",
            "Rango de Ventas (USD):",
            min=float(datos["Ventas_USD"].min()),
            max=float(datos["Ventas_USD"].max()),
            value=[float(datos["Ventas_USD"].min()), float(datos["Ventas_USD"].max())],
            step=10.0,
            pre="$",
        ),
        ui.hr(),
        ui.input_select(
            "agrupador_grafico",
            "Desglosar Gráfica por:",
            choices={"Region": "Región Geográfica", "Categoria": "Categoría de Producto"},
            selected="Region",
        ),
        ui.hr(),
        ui.div(
            ui.p("Aplicación desplegada con ", ui.strong("Shiny for Python"), " en ", ui.strong("Posit"), "."),
            class_="text-secondary small mt-auto",
        ),
        width=300,
    ),
    # Encabezado principal del Dashboard
    ui.div(
        ui.h2("📊 Dashboard de Rendimiento Comercial y Ventas", class_="fw-bold text-dark mb-1"),
        ui.p("Monitoreo interactivo de ingresos, volumen de transacciones y satisfacción del cliente.", class_="text-muted mb-4"),
    ),
    # Fila de Indicadores Clave de Desempeño (KPIs)
    ui.layout_columns(
        ui.value_box(
            "Ventas Totales",
            ui.output_text("kpi_ventas_totales"),
            showcase=ui.div("💰", style="font-size: 2.5rem;"),
            theme="primary",
        ),
        ui.value_box(
            "Ticket Promedio",
            ui.output_text("kpi_ticket_promedio"),
            showcase=ui.div("📈", style="font-size: 2.5rem;"),
            theme="success",
        ),
        ui.value_box(
            "Total Transacciones",
            ui.output_text("kpi_total_transacciones"),
            showcase=ui.div("🛒", style="font-size: 2.5rem;"),
            theme="info",
        ),
        ui.value_box(
            "Satisfacción Media",
            ui.output_text("kpi_satisfaccion"),
            showcase=ui.div("⭐", style="font-size: 2.5rem;"),
            theme="warning",
        ),
        col_widths=[3, 3, 3, 3],
    ),
    # Fila de Visualizaciones Gráficas
    ui.layout_columns(
        ui.card(
            ui.card_header("Ingresos Totales por Grupo Seleccionado", class_="fw-bold"),
            ui.output_plot("grafico_barras_ingresos"),
            full_screen=True,
        ),
        ui.card(
            ui.card_header("Relación Ventas vs. Satisfacción del Cliente", class_="fw-bold"),
            ui.output_plot("grafico_dispersion"),
            full_screen=True,
        ),
        col_widths=[6, 6],
    ),
    # Fila de Tabla de Datos
    ui.card(
        ui.card_header("Explorador de Transacciones Filtradas (Primeros 100 Registros)", class_="fw-bold"),
        ui.output_data_frame("tabla_registros"),
        full_screen=True,
    ),
    title="Dashboard Comercial - Shiny for Python",
    fillable=True,
)

# -----------------------------------------------------------------------------
# Lógica del Servidor (Server)
# -----------------------------------------------------------------------------
def server(input, output, session):

    # Cálculo reactivo para filtrar el DataFrame según los controles del usuario
    @reactive.calc
    def datos_filtrados():
        cats = input.filtro_categorias()
        rango = input.filtro_precio()

        # Validación si no hay categorías seleccionadas
        if not cats:
            return pd.DataFrame(columns=datos.columns)

        df = datos[
            (datos["Categoria"].isin(cats)) &
            (datos["Ventas_USD"] >= rango[0]) &
            (datos["Ventas_USD"] <= rango[1])
        ]
        return df

    # --- Salidas de KPIs ---
    @render.text
    def kpi_ventas_totales():
        df = datos_filtrados()
        if df.empty:
            return "$0.00"
        return f"${df['Ventas_USD'].sum():,.2f} USD"

    @render.text
    def kpi_ticket_promedio():
        df = datos_filtrados()
        if df.empty:
            return "$0.00"
        return f"${df['Ventas_USD'].mean():,.2f} USD"

    @render.text
    def kpi_total_transacciones():
        df = datos_filtrados()
        return f"{len(df):,} ventas"

    @render.text
    def kpi_satisfaccion():
        df = datos_filtrados()
        if df.empty:
            return "0.0 / 5.0"
        return f"{df['Satisfaccion_Cliente'].mean():.2f} / 5.0"

    # --- Salidas Gráficas ---
    @render.plot
    def grafico_barras_ingresos():
        df = datos_filtrados()
        fig, ax = plt.subplots(figsize=(7, 4.2))

        if df.empty:
            ax.text(0.5, 0.5, "No hay datos para los filtros seleccionados",
                    ha="center", va="center", fontsize=12, color="gray")
            ax.set_axis_off()
            return fig

        agrupador = input.agrupador_grafico()
        ventas_grupo = df.groupby(agrupador)["Ventas_USD"].sum().sort_values(ascending=False)

        colores = sns.color_palette("mako", len(ventas_grupo))
        barras = ax.bar(ventas_grupo.index, ventas_grupo.values, color=colores, edgecolor="black", linewidth=0.6)

        # Etiquetas con montos en cada barra
        for barra in barras:
            altura = barra.get_height()
            ax.annotate(
                f"${altura:,.0f}",
                xy=(barra.get_x() + barra.get_width() / 2, altura),
                xytext=(0, 4),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=9,
                fontweight="bold",
            )

        ax.set_title(f"Ventas Totales Facturadas por {agrupador}", fontsize=12, fontweight="bold", pad=12)
        ax.set_xlabel(agrupador, fontsize=10, fontweight="bold")
        ax.set_ylabel("Ingresos Totales (USD)", fontsize=10, fontweight="bold")
        ax.yaxis.set_major_formatter("${x:,.0f}")
        plt.xticks(rotation=20, ha="right")
        plt.tight_layout()
        return fig

    @render.plot
    def grafico_dispersion():
        df = datos_filtrados()
        fig, ax = plt.subplots(figsize=(7, 4.2))

        if df.empty:
            ax.text(0.5, 0.5, "No hay datos para los filtros seleccionados",
                    ha="center", va="center", fontsize=12, color="gray")
            ax.set_axis_off()
            return fig

        sns.scatterplot(
            data=df,
            x="Ventas_USD",
            y="Satisfaccion_Cliente",
            hue="Categoria",
            size="Unidades",
            sizes=(30, 200),
            alpha=0.75,
            ax=ax,
            palette="tab10",
        )

        ax.set_title("Venta (USD) vs. Calificación de Satisfacción", fontsize=12, fontweight="bold", pad=12)
        ax.set_xlabel("Importe de Venta (USD)", fontsize=10, fontweight="bold")
        ax.set_ylabel("Satisfacción (1 a 5)", fontsize=10, fontweight="bold")
        ax.xaxis.set_major_formatter("${x:,.0f}")
        ax.set_ylim(2.5, 5.3)
        ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", borderaxespad=0, fontsize=8)
        plt.tight_layout()
        return fig

    # --- Salida de Tabla de Datos ---
    @render.data_frame
    def tabla_registros():
        df = datos_filtrados().copy()
        if df.empty:
            return render.DataGrid(pd.DataFrame({"Mensaje": ["No se encontraron registros con los filtros actuales."]}))

        df["Fecha"] = df["Fecha"].dt.strftime("%Y-%m-%d")
        df["Ventas_USD"] = df["Ventas_USD"].apply(lambda v: f"${v:,.2f}")
        return render.DataGrid(df.head(100), selection_mode="rows")


# -----------------------------------------------------------------------------
# Instancia de la Aplicación Shiny
# -----------------------------------------------------------------------------
app = App(app_ui, server)

if __name__ == "__main__":
    app.run(port=8000)

