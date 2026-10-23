# 🕒 Formateador y Procesador de Horarios (Streamlit)

Aplicación web hecha en **Python** y **Streamlit** que automatiza el trabajo con los reportes de horarios de **Inspector 2007**: lee los reportes convertidos a Excel, calcula el turno de cada empleado, lo cruza con el export de MySQL y genera el archivo final para Recursos Humanos. Incluye además un conjunto de utilidades independientes (corrección de conversiones, filtros, agrupación y análisis gráfico).

---

## 🚀 Qué hace

La app tiene tres pestañas:

### 1. Generar reporte (flujo principal)

1. **Sube los archivos**
   - Horarios del personal **ACTIVO** (Inspector, `.xlsx`, Hoja1).
   - Horarios del personal **EN FIRMA** (Inspector, `.xlsx`) — *opcional*. Sin él se trabaja solo con el personal activo.
   - Export de **MySQL** (`.csv` o `.xlsx`), con la tarjeta en la columna A.
2. **Resultado:** métricas rápidas (empleados, en firma, turnos por revisar, coincidencias con el CSV), conteo por turno, avisos y listas de control (filas del CSV sin coincidencia y empleados de Inspector que no están en el CSV).
3. **Revisar y corregir horarios:** tabla editable para corregir horas y marcas de día (`*` = labora, `D` = descanso, vacío = fuera de su jornada) y reevaluar los turnos.
4. **Generar archivo final:** descarga `REPORTE_HORARIOS_FINAL.xlsx` con las columnas del CSV más los horarios de Inspector (incluida TOLERANCIAS).

Los turnos se clasifican como **MATUTINO, VESPERTINO, NOCTURNO, JORNADA ACUMULADA, JORNADA ACUMULADA NOCTURNA** o **REVISAR** cuando no cumplen ninguna regla.

### 2. Utilidades

| Utilidad | Para qué sirve |
|---|---|
| **Corregir salto de página de Nitro** | Sincroniza el PDF original con el Excel convertido por Nitro y reinserta la primera línea de empleado que se omite en cada salto de página, conservando formatos y celdas combinadas. |
| **Filtrar por quincenas** | Compara el archivo final contra la Quincena 1 (y opcionalmente la 2) y omite a quienes no aparecen. |
| **Agrupar por tipo de horario** | Junta a los empleados con el mismo horario y agrega el tipo (por ejemplo «MATUTINO LUNES A VIERNES BASIFICADOS 1»). Solo reordena y colorea; no agrega ni borra información. |
| **Filtrar por lista de tarjetas** | Extrae del reporte final solo las tarjetas indicadas, copiando columnas y formatos. |
| **Obtener personal EN FIRMA** | Cruza el personal inactivo del sistema con la quincena para saber a quién agregar en el reporte de Inspector. |
| **Buscar texto en exportado** | Copia las columnas A a D de las filas cuya columna elegida contiene el texto buscado. |
| **Dar formato a la quincena** | Reacomoda las columnas del archivo de la quincena al layout necesario. |

### 3. Análisis y Gráficos

Sube el archivo final para ver métricas generales y un desglose interactivo por turno (gráfica de barras con Plotly), con descarga en Excel del turno seleccionado.

---

## ✨ Detalles de uso

- **Barra de progreso** en los procesos pesados (reporte principal y corrección de Nitro), con el avance por etapa o por página del PDF.
- **Botón «🧹 Limpiar»** (arriba a la derecha): borra caché, resultados y archivos cargados para empezar de cero. Útil entre un reporte y otro.
- Los archivos se procesan en memoria durante la sesión; la app no los guarda en el repositorio.
- Existe una guía visual de uso en [`GUIA.html`](GUIA.html).

---

## 🛠️ Tecnologías

- **Python**: lógica de procesamiento.
- **Streamlit**: interfaz web (se requiere la versión 1.36 o superior).
- **Pandas**: manejo de datos tabulares.
- **Openpyxl**: lectura y escritura de Excel con estilos.
- **PyPDF**: extracción de texto de PDFs.
- **Plotly** y **Altair**: gráficos interactivos.

---

## ▶️ Cómo ejecutarla

```bash
git clone https://github.com/DanyFlow/Formateador_Horarios_Streamlit.git
cd Formateador_Horarios_Streamlit
pip install -r requirements.txt
streamlit run FormateadorV3.py
```

Se abre en `http://localhost:8501`. También se puede usar desde otros equipos o el celular dentro de la misma red local, o desplegarla en Streamlit Community Cloud.

---

## 📂 Estructura del repositorio

```
├── .devcontainer/          # Configuración para Codespaces / VS Code
├── FormateadorV3.py        # Interfaz de Streamlit (pestañas, botones, gráficos)
├── procesamiento.py        # Lógica de negocio: lectura de Inspector, turnos, cruces y utilidades
├── GUIA.html               # Guía visual de uso
├── requirements.txt        # Dependencias del proyecto
├── LICENSE.txt             # Licencia MIT
└── README.md
```

---

## 🔒 Privacidad

Los reportes contienen datos de personal. **No subas al repositorio** PDFs, Excel ni CSV reales; mantenlos solo en tu equipo y cárgalos desde la interfaz cuando los necesites.

---

## 📄 Licencia

Distribuido bajo la licencia **MIT**. Consulta [`LICENSE.txt`](LICENSE.txt).
