# Shiny-Prog-CD

## Dashboard Básico en Shiny for Python y Despliegue en Posit Connect Cloud

Este proyecto contiene una aplicación web analítica interactiva desarrollada con **Shiny for Python**, lista para ser ejecutada localmente o desplegada en la infraestructura en la nube de **Posit Connect Cloud** (cuenta: `rncm11`) o **shinyapps.io**.

---

## 📁 Estructura del Directorio

```text
Dashboard_Basico_Shiny/
├── app.py                                            # Código fuente de la aplicación (UI + Server)
├── requirements.txt                                  # Librerías necesarias para el contenedor en Posit
├── manifest.json                                     # Manifiesto de compilación generado con rsconnect
├── Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx # Reporte formal en formato Word (.docx)
├── Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf  # Reporte formal en formato PDF listo para entrega
├── crear_reporte.py                                  # Script generador de los documentos y gráficas
└── reporte_img/                                      # Figuras y diagramas de arquitectura incluidos en el reporte
    ├── flujo_despliegue.png
    └── graficas_dashboard.png
```

---

## 🚀 1. Ejecución en Servidor Local

Para probar el dashboard en tu computadora:

1. Abre tu terminal o PowerShell en esta carpeta:
   ```bash
   cd "Dashboard_Basico_Shiny"
   ```
2. Ejecuta la aplicación con el servidor de desarrollo de Shiny:
   ```bash
   shiny run --reload --port 8000 app.py
   ```
3. Abre tu navegador web en:
   👉 **`http://127.0.0.1:8000`**

---

## ☁️ 2. Tramitar Cuenta y Publicar en Posit (shinyapps.io)

### Paso A: Registrarse en Posit
1. Entra a **[https://www.shinyapps.io/](https://www.shinyapps.io/)** y crea tu cuenta gratuita (**Sign Up** con Google, GitHub o correo).
2. Durante el registro, elige tu **Account Name** (nombre de usuario), por ejemplo: `mi-usuario`. Este nombre formará parte de la URL de tu aplicación.

### Paso B: Obtener Token y Secret
1. En el panel de **shinyapps.io**, ve al menú superior derecho: **Account** -> **Tokens**.
2. Haz clic en el botón verde **+ Add Token**.
3. En la fila del token, haz clic en **Show**, luego en **Show Secret**, y finalmente en **Copy to Clipboard**.

### Paso C: Vincular tus Credenciales en la Terminal
En tu terminal local ejecuta:
```bash
rsconnect add-set --account <TU_CUENTA> --name <TU_CUENTA> --token <TU_TOKEN> --secret <TU_SECRET>
```

### Paso D: Desplegar la Aplicación a Internet
Desde la carpeta `Dashboard_Basico_Shiny`:
```bash
rsconnect deploy shiny . --name <TU_CUENTA> --title "Dashboard Basico Shiny"
```

El proceso subirá tu código, instalará las dependencias en un contenedor seguro de Posit y te entregará el enlace público global:
🌐 **`https://<TU_CUENTA>.shinyapps.io/dashboard_basico_shiny/`**

---

## 📄 3. Documento Escrito para Entrega

En esta misma carpeta encontrarás los documentos generados con formato académico institucional, portada, tablas, código y capturas:
- **`Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx`** (Microsoft Word editable)
- **`Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf`** (PDF de alta calidad)
