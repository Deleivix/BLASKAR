import os
import tempfile
import pandas as pd
import streamlit as st
from horas import (
    process_file, Persist, open_as_xlsx,
    find_all_proyectos_positions, extract_recurso_line,
    normalize_project, read_cell
)

st.set_page_config(page_title="Transformador Excel → IPI (web)", layout="wide")
st.title("Transformador Excel → IPI (web)")

persist = Persist.load()

def discover(path):
    """Devuelve (recursos, proyectos{codigo:nombre}) sin generar salida."""
    xlsx, wb = open_as_xlsx(path)
    ws = wb.active
    pos = find_all_proyectos_positions(ws)
    recursos = set()
    proyectos = {}
    for i, r in enumerate(pos):
        recursos.add(extract_recurso_line(ws, r) or "RECURSO DESCONOCIDO")
        r1 = r
        r2 = pos[i+1] - 2 if i < len(pos) - 1 else ws.max_row
        for rr in range(r1, r2 + 1):
            t = read_cell(ws, rr, 1) or read_cell(ws, rr, 2)
            p = normalize_project(t or "")
            if p:
                proyectos[p[0]] = p[1]
    return sorted(recursos), dict(sorted(proyectos.items()))

up = st.file_uploader("Sube el .xls/.xlsx", type=["xls", "xlsx"])

if up:
    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, up.name)
        with open(in_path, "wb") as f:
            f.write(up.read())

        recursos, proyectos = discover(in_path)

        # ---------- Clasificación con casillas (código + nombre) ----------
        with st.expander("Clasificar proyectos (persistente)", expanded=True):
            st.write("Marca tipo por fila. Si marcas ambos, se guarda **CONSTRUCCIÓN**.")

            rows = []
            for cod, nom in proyectos.items():
                rows.append({
                    "Código": cod,
                    "Nombre": nom,
                    "CONSTRUCCIÓN": persist.tipos.get(cod, "") == "CONSTRUCCION",
                    "REPARACIÓN": persist.tipos.get(cod, "") == "REPARACION",
                })
            df = pd.DataFrame(rows)

            edited = st.data_editor(
                df,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "CONSTRUCCIÓN": st.column_config.CheckboxColumn(),
                    "REPARACIÓN": st.column_config.CheckboxColumn(),
                },
                disabled=["Código", "Nombre"],
                key="df_clasificacion",
            )

            if st.button("Guardar clasificación"):
                for _, row in edited.iterrows():
                    cod = str(row["Código"])
                    nom = str(row["Nombre"])
                    con = bool(row["CONSTRUCCIÓN"])
                    rep = bool(row["REPARACIÓN"])
                    persist.nombres[cod] = nom
                    if con and not rep:
                        persist.tipos[cod] = "CONSTRUCCION"
                    elif rep and not con:
                        persist.tipos[cod] = "REPARACION"
                    elif con and rep:
                        persist.tipos[cod] = "CONSTRUCCION"
                    else:
                        persist.tipos[cod] = ""
                persist.asked_clasif = True
                persist.save()
                st.success("Clasificación guardada.")

        # ---------- Exclusiones (opcional) ----------
        with st.expander("Exclusiones persistentes"):
            col1, col2 = st.columns(2)

            with col1:
                st.caption("Recursos a excluir")
                rec_map = {r: st.checkbox(r, value=(r in persist.excluir_recursos)) for r in recursos}

            with col2:
                st.caption("Proyectos a excluir")
                proy_map = {
                    cod: st.checkbox(f"{cod} - {proyectos[cod]}", value=(cod in persist.excluir_proyectos))
                    for cod in proyectos.keys()
                }

            if st.button("Guardar exclusiones"):
                persist.excluir_recursos = [k for k, v in rec_map.items() if v]
                persist.excluir_proyectos = [k for k, v in proy_map.items() if v]
                persist.asked_excl = True
                persist.save()
                st.success("Exclusiones guardadas.")

        # ---------- Procesar ----------
        if st.button("Procesar y generar IPI"):
            out_path = process_file(in_path, persist)
            with open(out_path, "rb") as f:
                st.download_button(
                    "Descargar _IPI.xlsx",
                    data=f.read(),
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
