import io, os, tempfile
import streamlit as st
from horas import (
    process_file, Persist, open_as_xlsx,
    find_all_proyectos_positions, extract_recurso_line, normalize_project, read_cell
)

st.set_page_config(page_title="Transformador Excel → IPI", layout="wide")
st.title("Transformador Excel → IPI (web)")

persist = Persist.load()

def discover(path):
    """Devuelve recursos y proyectos detectados sin generar salida."""
    xlsx, wb = open_as_xlsx(path); ws = wb.active
    pos = find_all_proyectos_positions(ws)
    recursos=set(); proyectos={}
    for i,r in enumerate(pos):
        recursos.add(extract_recurso_line(ws,r) or "RECURSO DESCONOCIDO")
        r1, r2 = r, (pos[i+1]-2 if i < len(pos)-1 else ws.max_row)
        for rr in range(r1, r2+1):
            t = read_cell(ws, rr, 1) or read_cell(ws, rr, 2)
            p = normalize_project(t or "")
            if p: proyectos[p[0]] = p[1]
    return sorted(recursos), dict(sorted(proyectos.items()))

up = st.file_uploader("Sube el .xls/.xlsx", type=["xls","xlsx"])

if up:
    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, up.name)
        with open(in_path, "wb") as f: f.write(up.read())

        # Descubrir para configurar
        recursos, proyectos = discover(in_path)

        with st.expander("Clasificar proyectos (persistente)"):
            st.write("Marca tipo por lista rápida.")
            con_ini = [k for k,v in persist.tipos.items() if v=="CONSTRUCCION"]
            rep_ini = [k for k,v in persist.tipos.items() if v=="REPARACION"]
            col1, col2 = st.columns(2)
            with col1:
                con_sel = st.multiselect("CONSTRUCCIÓN", list(proyectos.keys()), default=con_ini, key="con")
            with col2:
                rep_sel = st.multiselect("REPARACIÓN", list(proyectos.keys()), default=rep_ini, key="rep")
            if st.button("Guardar clasificación"):
                for k in proyectos.keys():
                    persist.tipos[k] = "CONSTRUCCION" if k in con_sel else ("REPARACION" if k in rep_sel else "")
                    persist.nombres[k] = proyectos[k]
                persist.asked_clasif = True
                persist.save()
                st.success("Clasificación guardada.")

        with st.expander("Exclusiones persistentes"):
            exc_rec = st.multiselect("Recursos a excluir", recursos, default=persist.excluir_recursos)
            exc_proy = st.multiselect("Proyectos a excluir", list(proyectos.keys()), default=persist.excluir_proyectos,
                                      format_func=lambda k: f"{k} - {proyectos.get(k,'')}")
            if st.button("Guardar exclusiones"):
                persist.excluir_recursos = exc_rec
                persist.excluir_proyectos = exc_proy
                persist.asked_excl = True
                persist.save()
                st.success("Exclusiones guardadas.")

        if st.button("Procesar y generar IPI"):
            out_path = process_file(in_path, persist)  # reutiliza tu lógica
            with open(out_path, "rb") as f:
                st.download_button(
                    "Descargar _IPI.xlsx",
                    data=f.read(),
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
