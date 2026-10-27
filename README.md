# Shiny-Prog-CD

## Dashboard Básico en Shiny for Python y Despliegue en Posit Connect Cloud

Este proyecto contiene una aplicación analítica básica e interactiva desarrollada con **Shiny for Python**, lista para ser ejecutada localmente o desplegada en la infraestructura en la nube de **Posit Connect Cloud** (cuenta: [`rncm11`](https://connect.posit.cloud/rncm11)).

---

## 📁 Estructura del Directorio

```text
Dashboard_Basico_Shiny/
├── app.py                                            # Código del Dashboard Básico (UI + Server)
├── requirements.txt                                  # Librerías mínimas necesarias para Posit
├── manifest.json                                     # Manifiesto de compilación generado con rsconnect
├── Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx # Reporte formal de entrega en Word (.docx)
├── Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf  # Reporte formal de entrega en PDF (.pdf)
├── crear_reporte.py                                  # Script generador de los documentos
└── reporte_img/                                      # Figuras y diagramas incluidos en el reporte
    ├── flujo_actividad.png
    └── grafica_histograma.png
```

---

## 🚀 1. Ejecución Local

Para probar el dashboard básico en tu computadora:

1. Abre tu terminal en esta carpeta:
   ```bash
   cd "Dashboard_Basico_Shiny"
   ```
2. Ejecuta la aplicación:
   ```bash
   shiny run --reload --port 8000 app.py
   # o directamente:
   python app.py
   ```
3. Abre tu navegador web en:
   👉 **`http://127.0.0.1:8000`**

---

## ☁️ 2. Despliegue en Posit Connect Cloud

1. Inicia sesión en: **[https://connect.posit.cloud/rncm11](https://connect.posit.cloud/rncm11)**.
2. Haz clic en **Publish** (arriba a la derecha).
3. Selecciona **GitHub** y elige el repositorio: **`RolandoCastillo11/Shiny-Prog-CD`**.
4. Haz clic en **Deploy**. Posit detectará automáticamente `app.py` y `requirements.txt` y publicará la aplicación con enlace público HTTPS.

---

## 📄 3. Documentos Escritos

Los reportes detallados con portada académica, código fuente, explicaciones y capturas se encuentran listos para entrega:
- **`Reporte_Instalacion_Shiny_y_Despliegue_Posit.docx`** (Microsoft Word)
- **`Reporte_Instalacion_Shiny_y_Despliegue_Posit.pdf`** (Formato PDF)
