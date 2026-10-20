"""App Streamlit para generar el reporte de horarios del personal.

Ejecutar:  streamlit run Formateador.py
"""

import streamlit as st

import procesamiento as p

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="Reporte de horarios", page_icon="📋", layout="wide"
)
st.title("Reporte de horarios del personal")

tab_rep, tab_util = st.tabs(["Generar reporte", "Utilidades"])


# ----------------------------------------------------------------------------
# PESTAÑA 1: flujo principal
# ----------------------------------------------------------------------------
with tab_rep:
  st.subheader("1. Sube los archivos")
  st.caption(
      "Reportes de Inspector en Excel (.xlsx, Hoja1). El export de MySQL puede"
      " ser .csv o .xlsx. Sin el reporte EN FIRMA se trabaja solo con el"
      " personal activo; quienes solo estén en ese reporte saldrán como «sin"
      " coincidencia»."
  )

  c1, c2, c3 = st.columns(3)
  f_act = c1.file_uploader(
      "Horarios del personal ACTIVO (Inspector)", type=["xlsx"], key="f_act"
  )
  f_fir = c2.file_uploader(
      "Horarios del personal EN FIRMA (Inspector) — opcional",
      type=["xlsx"],
      key="f_fir",
  )
  f_csv = c3.file_uploader(
      "Export de MySQL (columna A = tarjeta)", type=["csv", "xlsx"], key="f_csv"
  )

  if st.button("Procesar", type="primary", disabled=not (f_act and f_csv)):
    try:
      with st.spinner("Procesando…"):
        st.session_state["res"] = p.procesar(f_act, f_fir, f_csv)
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001 - se muestra al usuario
      st.error(f"No pude procesar los archivos: {e}")

  res = st.session_state.get("res")
  if res:
    datos, cruce = res["datos"], res["cruce"]

    # ---- Resumen rápido
    st.subheader("2. Resultado")
    n_rev = int((datos["TURNO"] == "REVISAR").sum())
    m = st.columns(5)
    m[0].metric("Empleados en Inspector", len(datos))
    m[1].metric("En firma (inactivos)", int(datos["_INACTIVO"].sum()))
    m[2].metric("Turnos por REVISAR", n_rev)
    m[3].metric("Filas del CSV con coincidencia", cruce["coincidencias"])
    m[4].metric(
        "Filas del CSV sin coincidencia", len(cruce["sin_coincidencia"])
    )
    if cruce["vacias"]:
      st.caption(f"Se ignoraron {cruce['vacias']} filas vacías del CSV.")

    conteo = (
        datos["TURNO"]
        .value_counts()
        .rename_axis("Turno")
        .reset_index(name="Personal")
    )
    st.dataframe(conteo, hide_index=True)

    if res["avisos"]:
      with st.expander(f"Avisos ({len(res['avisos'])})", expanded=True):
        for a in res["avisos"]:
          st.warning(a)

    # ---- Listas de control
    with st.expander(
        "Filas del CSV sin coincidencia en Inspector"
        f" ({len(cruce['sin_coincidencia'])})"
    ):
      st.caption(
          "No pasan al archivo final. Revisa si la tarjeta está mal capturada o"
          " si falta el empleado en alguno de los reportes de Inspector."
      )
      st.dataframe(cruce["sin_coincidencia"], hide_index=True)
    with st.expander(
        "Empleados de Inspector que no están en el CSV"
        f" ({len(cruce['no_en_csv'])})"
    ):
      st.caption(
          "Tienen horario en Inspector pero no aparecen en el export de MySQL."
      )
      st.dataframe(cruce["no_en_csv"], hide_index=True)

    # ---- Corrección de horarios (macro 6)
    st.subheader("3. Revisar y corregir horarios")
    st.caption(
        "Corrige horas o marcas de día (* = labora, D = descanso, vacío = fuera"
        " de su jornada) y pulsa «Aplicar cambios y reevaluar». Un error común:"
        " en jornada acumulada nocturna, el lunes debe quedar vacío."
    )
    solo = st.toggle("Mostrar solo los REVISAR", value=n_rev > 0, key="solo_rev")
    vista = datos[datos["TURNO"] == "REVISAR"] if solo else datos
    cols_edit = ["ENTRADA", "SALIDA", *p.DIAS, "HORAS DE TRABAJO"]
    editado = None
    if solo and vista.empty:
      st.success("No hay horarios en REVISAR.")
    else:
      cfg = {
          d: st.column_config.SelectboxColumn(
              d, options=["*", "D"], width="small"
          )
          for d in p.DIAS
      }
      cfg["_HORARIO"] = st.column_config.TextColumn("Horario (Inspector)")
      for c in ("ENTRADA", "SALIDA", "HORAS DE TRABAJO"):
        cfg[c] = st.column_config.TextColumn(
            c, help="Formato HH:MM", width="small"
        )
      editado = st.data_editor(
          vista[[
              "TARJETA",
              "NOMBRE",
              "_HORARIO",
              *cols_edit,
              "TURNO",
              "TOLERANCIAS",
          ]],
          column_config=cfg,
          disabled=["TARJETA", "NOMBRE", "_HORARIO", "TURNO", "TOLERANCIAS"],
          hide_index=True,
          key=f"editor_{res['version']}_{int(solo)}",
      )
    if st.button("Aplicar cambios y reevaluar turnos"):
      if editado is not None:
        datos.loc[editado.index, cols_edit] = editado[cols_edit]
      res["datos"] = p.evaluar_turnos(datos)
      res["version"] += 1
      res["final"] = None
      st.rerun()

    # ---- Archivo final
    st.subheader("4. Generar archivo final")
    if n_rev:
      st.info(
          f"Todavía hay {n_rev} turno(s) en REVISAR; saldrán así en el archivo"
          " final."
      )
    if cruce["coincidencias"] == 0:
      st.error(
          "Ninguna fila del CSV coincide con los empleados de Inspector; revisa"
          " la columna A del CSV."
      )
    elif st.button("Generar archivo final"):
      res["final"] = p.exportar_final(res["datos"], res["csv"])
    if res.get("final"):
      st.download_button(
          "Descargar archivo final (.xlsx)",
          res["final"],
          file_name="REPORTE_HORARIOS_FINAL.xlsx",
          mime=XLSX,
      )
      st.caption(
          f"{cruce['coincidencias']} fila(s): columnas del CSV + ENTRADA a"
          " TOTAL de Inspector (TOLERANCIAS incluida). Solo valores."
      )


# ----------------------------------------------------------------------------
# PESTAÑA 2: utilidades independientes
# ----------------------------------------------------------------------------
def _guardar(clave_estado, valor):
  st.session_state[clave_estado] = valor


with tab_util:
  # ---- Corrector de conversión Nitro (PDF -> Excel)
  st.subheader("Corregir salto de página en conversión Nitro (PDF → Excel)")
  st.caption(
      "Sincroniza los empleados del PDF original con el Excel convertido por"
      " Nitro, reinsertando únicamente la primera línea de empleado omitida en"
      " cada salto de página manteniendo intacto el formato."
  )
  u_pdf = st.file_uploader("PDF Original de Inspector", type=["pdf"], key="u_nitro_pdf")
  u_excel = st.file_uploader("Excel Convertido por Nitro", type=["xlsx"], key="u_nitro_excel")
  if st.button(
      "Corregir conversion de Nitro",
      key="u_nitro_btn",
      disabled=not (u_pdf and u_excel),
  ):
    try:
      with st.spinner("Procesando conversión Nitro…"):
        _guardar("u_nitro_res", p.util_corregir_nitro(u_pdf, u_excel))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar los archivos: {e}")

  if "u_nitro_res" in st.session_state:
    bytes_nitro, n_insertados = st.session_state["u_nitro_res"]
    st.success(
        "¡Proceso completado con éxito! Se reinsertaron"
        f" {n_insertados} empleados omitidos por Nitro."
    )
    st.download_button(
        "Descargar Excel Corregido",
        bytes_nitro,
        file_name="REPORTE_ASISTENCIA_CORREGIDO.xlsx",
        mime=XLSX,
        key="u_nitro_dl",
    )

  st.divider()

# ---- Filtrar archivo final por quincenas
  st.subheader("Filtrar archivo final por pago en quincenas")
  st.caption(
      "Compara la columna B del archivo final contra la columna A de la Quincena 1 y, "
      "si se proporciona, contra la Quincena 2. Genera un nuevo archivo con sus formatos "
      "originales omitiendo a los empleados que no aparezcan en las quincenas."
  )
  uf_q1 = st.file_uploader("Archivo de Quincena 1 (Hoja1)", type=["xlsx"], key="uf_q1")
  uf_q2 = st.file_uploader("Archivo de Quincena 2 (Hoja1) — opcional", type=["xlsx"], key="uf_q2")
  uf_fin = st.file_uploader("Archivo Final o Agrupado", type=["xlsx"], key="uf_fin")

  if st.button("Filtrar por quincenas", key="uf_btn", disabled=not (uf_q1 and uf_fin)):
    try:
      with st.spinner("Procesando y filtrando registros..."):
        bytes_uf, copiadas, total = p.util_filtrar_por_quincenas(uf_q1, uf_q2, uf_fin)
        _guardar("uf_res", (bytes_uf, copiadas, total))
    except ValueError as e:
      st.error(str(e))
    except Exception as e:
      st.error(f"No pude procesar los archivos: {e}")

  if "uf_res" in st.session_state:
    bytes_uf, copiadas, total = st.session_state["uf_res"]
    st.success(
        f"¡Filtrado completado con éxito! Se conservaron {copiadas} de {total} empleados "
        f"(se omitieron {total - copiadas} registros sin coincidencia en las quincenas)."
    )
    st.download_button(
        "Descargar archivo filtrado",
        bytes_uf,
        file_name="REPORTE_FILTRADO_QUINCENAS.xlsx",
        mime=XLSX,
        key="uf_dl",
    )

  # ---- Agrupar empleados por tipo de horario
  st.subheader("Agrupar empleados por tipo de horario")
  st.caption(
      "Sube el reporte final. Junta a los empleados con el mismo horario"
      " (columnas J a U) en un archivo nuevo con una sola hoja (AGRUPADO): copia"
      " las columnas A a X y agrega en la Y el tipo de horario, por ejemplo"
      " «MATUTINO LUNES A VIERNES BASIFICADOS 1» o «NOCTURNO A (L-Mi-V)"
      " BASIFICADOS 1». Solo reordena; no agrega ni borra información."
  )
  u_agr = st.file_uploader(
      "Reporte final (REPORTE_HORARIOS_FINAL.xlsx)",
      type=["xlsx"],
      key="u_agr_file",
  )
  if st.button(
      "Agrupar por tipo de horario", key="u_agr_btn", disabled=u_agr is None
  ):
    try:
      with st.spinner("Agrupando…"):
        _guardar("u_agr_res", p.util_agrupar_horarios(u_agr))
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar el archivo: {e}")
  if "u_agr_res" in st.session_state:
    bytes_agr, resumen_agr, avisos_agr, n_agr = st.session_state["u_agr_res"]
    st.success(
        f"{n_agr} empleados agrupados en {len(resumen_agr)} tipos de horario."
        " Verificado: mismas filas y valores que el original."
    )
    for a in avisos_agr:
      st.warning(a)
    st.caption(
        "Colores en el archivo: azul suave = grupos alternos; amarillo = días"
        " irregulares; naranja = sin tolerancia; rojo = tolerancia distinta a"
        " las 3 principales."
    )
    st.dataframe(resumen_agr, hide_index=True)
    st.download_button(
        "Descargar reporte agrupado",
        bytes_agr,
        file_name="REPORTE_HORARIOS_AGRUPADO.xlsx",
        mime=XLSX,
        key="u_agr_dl",
    )

  st.divider()

  

  # ---- Obtener personal en firma (macro 2)
  st.subheader("Obtener personal EN FIRMA")
  st.caption(
      "Cruza el personal inactivo del sistema (export de MySQL) con la"
      " quincena. La lista resultante es la que hay que agregar a mano en el"
      " reporte de Inspector (con «mostrar empleados inactivos»)."
  )
  u2a = st.file_uploader(
      "Export de MySQL del personal inactivo (tarjeta en la columna A)",
      type=["csv", "xlsx"],
      key="u2_todos",
  )
  u2b = st.file_uploader(
      "Archivo de la quincena (Hoja1)", type=["xlsx"], key="u2_quin"
  )
  if st.button(
      "Obtener personal en firma", key="u2_btn", disabled=not (u2a and u2b)
  ):
    try:
      _guardar("u2_res", p.util_personal_en_firma(u2a, u2b))
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar los archivos: {e}")
  if "u2_res" in st.session_state:
    bytes2, vista2 = st.session_state["u2_res"]
    st.success(f"{len(vista2)} persona(s) inactivas con pago en la quincena.")
    st.dataframe(vista2, hide_index=True)
    st.download_button(
        "Descargar lista EN FIRMA",
        bytes2,
        file_name="EN_FIRMA.xlsx",
        mime=XLSX,
        key="u2_dl",
    )

  st.divider()

  # ---- Buscar «gua» (macro 8)
  st.subheader("Buscar «gua» en exportado del sistema")
  st.caption(
      "Copia las columnas A a D de las filas cuya columna E contiene el texto"
      " buscado."
  )
  u8 = st.file_uploader(
      "Excel exportado del sistema (Hoja1)", type=["xlsx"], key="u8_file"
  )
  ca, cb = st.columns(2)
  texto = ca.text_input("Texto a buscar", value="gua", key="u8_txt")
  col_n = cb.number_input(
      "Número de columna donde buscar (E = 5)",
      min_value=1,
      value=5,
      step=1,
      key="u8_col",
  )
  if st.button(
      "Ejecutar búsqueda", key="u8_btn", disabled=u8 is None or not texto
  ):
    try:
      _guardar("u8_res", p.util_buscar_gua(u8, texto, int(col_n)))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar el archivo: {e}")
  if "u8_res" in st.session_state:
    datos8, vista8 = st.session_state["u8_res"]
    st.success(f"{len(vista8)} fila(s) encontradas.")
    st.dataframe(vista8, hide_index=True)
    st.download_button(
        "Descargar resultado",
        datos8,
        file_name="RESULTADO_GUA.xlsx",
        mime=XLSX,
        key="u8_dl",
    )

  st.divider()

  # ---- Formato de quincena (macro 9)
  st.subheader("Dar formato necesario a la quincena")
  st.caption(
      "Reacomoda las columnas del archivo de la quincena (origen A:U → destino"
      " E:X; la columna H del origen no se copia)."
  )
  u9 = st.file_uploader(
      "Excel de la quincena (Hoja1)", type=["xlsx"], key="u9_file"
  )
  if st.button("Dar formato", key="u9_btn", disabled=u9 is None):
    try:
      _guardar("u9_res", p.util_formato_quincena(u9))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar el archivo: {e}")
  if "u9_res" in st.session_state:
    st.success(f"Listo: {st.session_state['u9_res'][1]} fila(s) reacomodadas.")
    st.download_button(
        "Descargar quincena con formato",
        st.session_state["u9_res"][0],
        file_name="QUINCENA_FORMATO.xlsx",
        mime=XLSX,
        key="u9_dl",
    )

# -------Firma
import base64
import streamlit as st

# ------- Firma flotante en la esquina inferior izquierda
url_imagen_drive = "https://drive.google.com/file/d/1oCCVTyUfMABigBHvsnSGwfOof8gxKFaV/view?usp=drive_link"
enlace_perfil = "https://share.streamlit.io/user/danyflowt"

firma_html = f"""
<style>
.firma-flotante-link {{
    position: fixed;
    bottom: 20px;
    left: 20px;
    z-index: 999999;
    text-decoration: none;
}}

.firma-flotante {{
    display: flex;
    align-items: center;
    background: #06b6d4;
    color: #ffffff;
    width: 44px;
    height: 44px;
    padding: 6px;
    border-radius: 50px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    border: 1px solid #0891b2;
    font-family: var(--font, sans-serif);
    font-size: 13px;
    overflow: hidden;
    transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.3s ease, background 0.3s ease, padding 0.3s ease;
    cursor: pointer;
    white-space: nowrap;
}}

.firma-flotante:hover {{
    width: max-content; 
    background: #0891b2; 
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.35);
    padding-right: 16px; 
}}

.firma-flotante img {{
    width: 30px;
    height: 30px;
    border-radius: 50%;
    object-fit: cover;
    flex-shrink: 0;
    border: 1.5px solid rgba(255, 255, 255, 0.9);
}}

.firma-texto {{
    margin-left: 12px;
    opacity: 0;
    transition: opacity 0.25s ease 0.1s;
    display: flex;
    flex-direction: column;
    justify-content: center;
    text-align: left;
    line-height: 1.3;
}}

.firma-flotante:hover .firma-texto {{
    opacity: 1;
}}

.firma-titulo {{
    font-weight: 600;
    font-size: 12px;
    color: #ffffff;
}}

.firma-sub {{
    font-size: 10px;
    color: #ecfeff;
}}
</style>

<a href="{enlace_perfil}" target="_blank" class="firma-flotante-link" title="Visitar perfil de Streamlit">
    <div class="firma-flotante">
        <img src="{url_imagen_drive}" alt="Dany FloVel">
        <div class="firma-texto">
            <span class="firma-titulo">Dany FloVel</span>
            <span class="firma-sub">⚡ Ver perfil ⚡</span>
        </div>
    </div>
</a>
"""

st.markdown(firma_html, unsafe_allow_html=True)
