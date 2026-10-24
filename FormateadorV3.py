"""App Streamlit para generar el reporte de horarios del personal.

Ejecutar:  streamlit run Formateador.py
"""



import altair as alt

import pandas as pd

import streamlit as st

import plotly.express as px

import procesamiento as p

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="Reporte de horarios", page_icon="📋", layout="wide"
)


def _crear_barra(texto):
  """Barra de progreso + callback que la actualiza."""
  barra = st.progress(0.0, text=texto)
  return barra, lambda frac, t="": barra.progress(frac, text=t)


def limpiar_todo():
  """Borra caché, resultados y archivos cargados para empezar de cero."""
  st.cache_data.clear()
  st.cache_resource.clear()
  for k in list(st.session_state.keys()):
    del st.session_state[k]


c_tit, c_lim = st.columns([6, 1], vertical_alignment="center")
c_tit.title("Reporte de horarios del personal 📇")
c_lim.button(
    "🧹 Limpiar",
    on_click=limpiar_todo,
    help="Borra resultados y archivos cargados para empezar de cero.",
    use_container_width=True,
)

tab_rep, tab_util, tab_analisis = st.tabs(["Generar reporte", "Utilidades", "Análisis y Gráficos"])


# ----------------------------------------------------------------------------
# PESTAÑA 1: flujo principal
# ----------------------------------------------------------------------------
with tab_rep:
  st.subheader("1 - Dar formato a reportes de inspector PDF->EXCEL 📚")
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
    barra, cb = _crear_barra("Iniciando…")
    try:
      st.session_state["res"] = p.procesar(f_act, f_fir, f_csv, progreso=cb)
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001 - se muestra al usuario
      st.error(f"No pude procesar los archivos: {e}")
    finally:
      barra.empty()

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
  st.subheader("1 - Corregir salto de página en conversión Nitro (PDF → Excel) 📝")
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
    barra, cb = _crear_barra("Iniciando…")
    try:
      _guardar("u_nitro_res", p.util_corregir_nitro(u_pdf, u_excel, progreso=cb))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar los archivos: {e}")
    finally:
      barra.empty()

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
      # ---- Agrupar empleados por tipo de horario
  st.subheader("2 - Agrupar empleados por tipo de horario 👥")
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

# ---- Filtrar archivo final por quincenas
  st.subheader("3 - Filtrar archivo final por pago en quincenas 🗂️")
  st.caption(
      "Compara el archivo generado en el pasoa anterior con los trabajadores activos en la quincena a subir, con la opcion de subir una segunda quincena para descartar trabajadores que no aparezcan por encontrarse en descanso olbigatorio"
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

  # ---- Agrupar por tipo de horario (Y) y UR (I)
  st.subheader("4 - Agrupar por tipo de horario y UR 🗃️")
  st.caption(
      "Sube el reporte final ya agrupado por tipo de horario y re ordena de acuerdo a UR y Horario con 2 opciones de descarga distintas"
  )
  u_ur = st.file_uploader(
      "📥 Opción 1 - Reporte final agrupado (columna Y = TIPO DE HORARIO)",
      type=["xlsx"],
      key="u_ur_file",
  )
  if st.button(
      "Agrupar por tipo de horario y UR", key="u_ur_btn", disabled=u_ur is None
  ):
    barra, cb = _crear_barra("Iniciando…")
    try:
      _guardar("u_ur_res", p.util_agrupar_por_ur(u_ur, progreso=cb))
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar el archivo: {e}")
    finally:
      barra.empty()
  if "u_ur_res" in st.session_state:
    bytes_ur, resumen_ur, avisos_ur, n_ur = st.session_state["u_ur_res"]
    st.success(
        f"{n_ur} empleados agrupados en {len(resumen_ur)} grupos (tipo de"
        " horario + UR). Verificado: mismas filas y valores que el original,"
        " salvo la columna I."
    )
    for a in avisos_ur:
      st.warning(a)
    st.dataframe(resumen_ur, hide_index=True)
    st.download_button(
        "Descargar reporte agrupado por UR (una sola hoja)",
        bytes_ur,
        file_name="REPORTE_HORARIOS_UR.xlsx",
        mime=XLSX,
        key="u_ur_dl",
    )

  # Segunda opción: pestañas por tipo GENERAL + UR
  st.markdown("---")
  st.markdown(
      " 📥 Opción 2 — Varias pestañas (tipo general + UR)**  \n"
      "Unifica las variantes numeradas y de tolerancia "
  )
  if st.button(
      "Agrupar por tipo general + UR (varias pestañas)",
      key="u_ur_pest_btn",
      disabled=u_ur is None,
  ):
    barra, cb = _crear_barra("Iniciando…")
    try:
      _guardar(
          "u_ur_pest_res", p.util_agrupar_por_ur_pestanas(u_ur, progreso=cb)
      )
    except ValueError as e:
      st.error(str(e))
    except Exception as e:  # noqa: BLE001
      st.error(f"No pude procesar el archivo: {e}")
    finally:
      barra.empty()

  if "u_ur_pest_res" in st.session_state:
    bytes_pest, resumen_pest, avisos_pest, n_pest = st.session_state[
        "u_ur_pest_res"
    ]
    st.success(
        f"{n_pest} empleados agrupados en {len(resumen_pest)} pestañas"
        " (tipo general + UR). Verificado: mismas filas y valores que el"
        " original, salvo la columna I."
    )
    for a in avisos_pest:
      st.warning(a)
    st.dataframe(resumen_pest, hide_index=True)
    st.download_button(
        "Descargar reporte con pestañas (tipo general + UR)",
        bytes_pest,
        file_name="REPORTE_HORARIOS_UR_PESTANAS.xlsx",
        mime=XLSX,
        key="u_ur_pest_dl",
    )

  st.divider()

  # ---- Filtrar archivo final por tarjetas específicas
  st.subheader("5 - Filtrar archivo final por lista de tarjetas 📑")
  st.caption(
      "Sube el reporte final e ingresa las tarjetas de los empleados que deseas "
      "extraer separadas por comas (ejemplo: 1023, 1024, 1050). Se generará un nuevo "
      "archivo copiando todas las columnas y formatos originales de las coincidencias."
  )
  u_ft_file = st.file_uploader("Reporte Final o Agrupado (.xlsx)", type=["xlsx"], key="u_ft_file")
  u_ft_txt = st.text_area("Tarjetas a buscar (separadas por comas)", value="", placeholder="Ejemplo: 1023, 1024, 1050", key="u_ft_txt")

  if st.button("Filtrar por tarjetas", key="u_ft_btn", disabled=not (u_ft_file and u_ft_txt.strip())):
    try:
      with st.spinner("Buscando y filtrando tarjetas..."):
        _guardar("u_ft_res", p.util_filtrar_por_tarjetas(u_ft_file, u_ft_txt))
    except ValueError as e:
      st.error(str(e))
    except Exception as e:
      st.error(f"No pude procesar el archivo: {e}")

  if "u_ft_res" in st.session_state:
    bytes_ft, encontradas, total_buscadas, no_encontradas, df_vista_ft = st.session_state["u_ft_res"]
    st.success(f"¡Filtrado completado! Se encontraron {encontradas} fila(s) de las {total_buscadas} tarjetas ingresadas.")
    
    if no_encontradas:
      st.warning(f"Las siguientes {len(no_encontradas)} tarjetas no se encontraron en el archivo: {', '.join(no_encontradas)}")

    if not df_vista_ft.empty:
      st.dataframe(df_vista_ft, hide_index=True)

    st.download_button(
        "Descargar reporte filtrado por tarjetas",
        bytes_ft,
        file_name="REPORTE_FILTRADO_TARJETAS.xlsx",
        mime=XLSX,
        key="u_ft_dl",
    )

  st.divider()

  # ---- Buscar «gua» (macro 8)
  st.subheader("6 - Buscar personal con horario Guarderia 🏡")
  st.caption(
      "Busca de manera resumida a los empleados con horario de guarderia del reporte final"
  )
  u8 = st.file_uploader(
      "Reporte final", type=["xlsx"], key="u8_file"
  )
  ca, cb = st.columns(2)
  texto = ca.text_input("Texto a buscar", value="Guardería", key="u8_txt")
  col_n = cb.number_input(
      "Número de columna donde buscar (U = 21)",
      min_value=1,
      value=21,
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

# ----------------------------------------------------------------------------
# PESTAÑA 3: análisis gráfico interactivo y descarga por turno (Columna T)
# ----------------------------------------------------------------------------
with tab_analisis:
  st.subheader("1 - Análisis de datos, métricas y gráficos por Turno 📊")
  st.caption(
      "Sube tu archivo final para visualizar el desglose de la columna AA (Descr) "
      "frente a la columna AB (Tipo de Trabajador) filtrado por Turno (Columna T)."
  )

  u_ana = st.file_uploader("Sube tu archivo para análisis", type=["xlsx"], key="u_ana_file")

  if u_ana is not None:
    try:
      df_ana, c_turno, c_descr, c_tipotrab = p.util_analizar_reporte(u_ana)
      st.success(f"¡Archivo cargado con éxito! {len(df_ana)} registros procesados.")

      # Métricas generales arriba
      m1, m2, m3 = st.columns(3)
      m1.metric("Total de Registros", len(df_ana))
      turnos_unicos = sorted(df_ana[c_turno].dropna().astype(str).unique())
      m2.metric("Turnos Detectados (Col. T)", len(turnos_unicos))
      m3.metric("Descripciones Únicas (Col. AA)", df_ana[c_descr].dropna().nunique())

      st.divider()

      # Selector de Turno (Columna T)
      st.subheader("Desglose interactivo por Turno")
      turno_sel = st.selectbox("Selecciona el Turno (Columna T):", options=turnos_unicos, key="sel_turno_ana")

      if turno_sel:
        df_turno = df_ana[df_ana[c_turno].astype(str) == str(turno_sel)]

        # Agrupación: número de registros por Descripción (AA) y Tipo de Trabajador (AB)
        tabla_resumen = df_turno.groupby([c_descr, c_tipotrab]).size().reset_index(name="CANTIDAD")

        st.markdown(f"**Registros para el turno `{turno_sel}`** ({len(df_turno)} empleados en total):")
        st.dataframe(tabla_resumen, hide_index=True)

        # Gráfico interactivo completo con Plotly Express
        if not tabla_resumen.empty:
          st.markdown("##### 📊 Distribución por Descripción (AA) y Tipo de Trabajador (AB)")
          
          # Calculamos un alto dinámico para que siempre haya espacio suficiente si hay muchas categorías
          num_categorias = tabla_resumen[c_descr].nunique()
          alto_grafico = max(450, num_categorias * 35)

          # Gráfico de barras horizontales (evita que los nombres largos de la columna AA se corten)
          fig = px.bar(
              tabla_resumen,
              y=c_descr,
              x="CANTIDAD",
              color=c_tipotrab,
              orientation="h",
              title=f"Desglose de personal - Turno: {turno_sel}",
              labels={
                  c_descr: "Descripción (Columna AA)",
                  "CANTIDAD": "Número de Empleados",
                  c_tipotrab: "Tipo de Trabajador (Columna AB)"
              },
              text="CANTIDAD" # Muestra la cifra exacta sobre/dentro de cada barra
          )

          # Personalización visual de ejes, márgenes y leyenda
          fig.update_layout(
              height=alto_grafico,
              yaxis={'categoryorder': 'total ascending'}, # Ordena de mayor a menor cantidad
              xaxis_title="Número de Empleados",
              yaxis_title="Descripción (Columna AA)",
              legend_title_text="Tipo de Trabajador",
              margin=dict(l=20, r=20, t=50, b=50),
              hovermode="closest"
          )
          fig.update_traces(textposition="outside")

          st.plotly_chart(fig, use_container_width=True)

        st.divider()

        # Botón opcional para descargar únicamente ese turno
        bytes_turno, n_turno = p.exportar_por_turno(df_ana, c_turno, turno_sel)
        st.download_button(
            label=f"📥 Descargar reporte en Excel para el turno {turno_sel} ({n_turno} registros)",
            data=bytes_turno,
            file_name=f"REPORTE_TURNO_{str(turno_sel).replace(' ', '_')}.xlsx",
            mime=XLSX,
            key=f"dl_turno_{turno_sel}"
        )

    except Exception as e:
      st.error(f"No se pudo procesar el archivo para análisis: {e}")

# -------Firma
import base64
import streamlit as st

# ------- Firma flotante en la esquina superior izquierda
url_imagen_drive = "https://i.redd.it/snoovatar/avatars/98e5bfa6-95f3-46d8-8754-7a64c0099533.png"
enlace_perfil = "https://share.streamlit.io/user/danyflow"

firma_html = f"""
<style>
.firma-flotante-link {{
    position: fixed;
    top: 20px;
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
