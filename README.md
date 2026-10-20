# 🕒 Formateador y Procesador de Horarios (Streamlit)

Aplicación web desarrollada en **Python** y **Streamlit** diseñada para automatizar la limpieza, estructuración, corrección de saltos de página en reportes PDF masivos generados con el software Inspector y agrupación de turnos en archivos de Excel administrativos para recursos humanos.

---

## 🚀 Características Principales

* **Procesamiento Masivo de Reportes:** Automatiza la lectura y conversión de datos entre formatos PDF y hojas de cálculo generados con la herramienta de reportes de inspector 2007.
* **Corrección de Saltos de Página (`util_corregir_nitro`):** Corrige el desbordamiento de páginas y reestructura las celdas y formatos originales de forma automatizada.
* **Agrupación Inteligente de Horarios:** Organiza y formatea automáticamente los registros aplicando bordes limpios y alineación centrada.
* **Interfaz Web Interactiva:** Accesible tanto desde equipos de escritorio como desde dispositivos móviles mediante red local o despliegue en la nube.

---

## 🛠️ Tecnologías y Librerías Utilizadas

* **Python** (Lógica principal de procesamiento)
* **Streamlit** (Interfaz gráfica web)
* **Pandas** (Manipulación y análisis de datos tabulares)
* **Openpyxl** (Manipulación avanzada de archivos Excel y estilos)
* **PyPDF / PDFPlumber** (Extracción de texto y manejo de documentos PDF)

---
## 📂 Estructura del Repositorio

```text
├── FormateadorV3.py        # Archivo principal de la interfaz de Streamlit
├── procesamiento.py        # Módulo con la lógica de negocio y funciones pesadas
├── requirements.txt        # Dependencias y librerías del proyecto
└── LICENSE                 # Licencia MIT del proyecto
