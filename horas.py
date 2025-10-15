#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, re, json, tempfile
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

# --- Tkinter opcional (la web/servidor no lo tiene) ---
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None
    filedialog = messagebox = ttk = None

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ======================================================================================
# Configuración
# ======================================================================================

JSON_PATH = os.environ.get(
    "JSON_PATH",
    os.path.join(os.path.dirname(__file__), "clasificacion_proyectos.json"),
)
SHEET_SALIDA = "IPI"

COLOR_HEADER       = "D9D9D9"  # gris
COLOR_CONSTRUCCION = "9DC3E6"  # azul
COLOR_REPARACION   = "C6E0B4"  # verde

RE_PROYECTOS_TAG = re.compile(r"^\s*Proyectos:\s*$", re.I)
RE_COD_PROY      = re.compile(r"^\s*(\d{5})\s*[-–]\s*(.+)$", re.U)
RE_TIPO_HORA     = re.compile(r"^\s*\d+\s*-\s*HORA", re.I)
RE_HHMM          = re.compile(r"^\s*(\d{1,2})\s*:\s*([0-5]\d)\s*$")

GUI_ENABLED = tk is not None

# ======================================================================================
# Utilidades
# ======================================================================================

def hhmm_to_minutes(text: str) -> int:
    if not text:
        return 0
    m = RE_HHMM.match(str(text))
    if not m:
        return 0
    return int(m.group(1)) * 60 + int(m.group(2))

def minutes_to_hhmm(total_min: int) -> str:
    return f"{total_min//60:02d}:{total_min%60:02d}"

def normalize_project(text: str) -> Optional[Tuple[str, str]]:
    if not text:
        return None
    s = str(text).strip()
    m = RE_COD_PROY.match(s)
    if m:
        return m.group(1), m.group(2).strip()
    m2 = re.match(r"^\s*(\d{5})\s+(.*)$", s)
    if m2:
        return m2.group(1), m2.group(2).strip()
    return None

def read_cell(ws, r, c):
    v = ws.cell(row=r, column=c).value
    if v is None:
        return ""
    return v.strip() if isinstance(v, str) else str(v).strip()

# ======================================================================================
# Filtro de filas “basura” (cabeceras repetidas en mitad del archivo)
# ======================================================================================

MONTHS_ES = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","setiembre","octubre","noviembre","diciembre"
]
RE_MONTH_BANNER = re.compile(rf"^\s*({'|'.join(MONTHS_ES)})\s+de\s+\d{{4}}\s*$", re.I)
DOW_TOKENS = {"lu","lu.","ma","ma.","mi","mi.","ju","ju.","vi","vi.","sá","sá.","sa","sa.","do","do.","dom","dom."}

def _row_has_month_banner(ws, r:int)->bool:
    for c in range(1, ws.max_column+1):
        if RE_MONTH_BANNER.match(read_cell(ws, r, c) or ""):
            return True
    return False

def _row_is_header_recurso(ws, r:int)->bool:
    a = (read_cell(ws, r, 1) or "").lower()
    b = (read_cell(ws, r, 2) or "").lower()
    return a.startswith("recurso") and b.startswith("tipo")

def _row_is_header_dow(ws, r:int, day_start:int, n_days:int)->bool:
    cnt = 0
    for c in range(day_start, min(ws.max_column, day_start + n_days - 1) + 1):
        v = (read_cell(ws, r, c) or "").strip().lower()
        if v in DOW_TOKENS:
            cnt += 1
    return cnt >= 5

def is_garbage_row(ws, r:int, day_start:int, n_days:int)->bool:
    return (
        _row_has_month_banner(ws, r) or
        _row_is_header_recurso(ws, r) or
        _row_is_header_dow(ws, r, day_start, n_days)
    )

# ======================================================================================
# Datos
# ======================================================================================

@dataclass
class RowData:
    recurso: str
    proyecto_codigo: Optional[str] = None
    proyecto_nombre: Optional[str] = None
    tipo_proyecto: Optional[str] = None
    tipo_imputacion: Optional[str] = None
    horas_por_dia: List[str] = field(default_factory=list)

@dataclass
class Persist:
    tipos: Dict[str, str] = field(default_factory=dict)      # codigo -> tipo
    nombres: Dict[str, str] = field(default_factory=dict)    # codigo -> nombre
    excluir_proyectos: List[str] = field(default_factory=list)
    excluir_recursos: List[str] = field(default_factory=list)
    asked_clasif: bool = False
    asked_excl: bool = False

    @classmethod
    def load(cls):
        if not os.path.exists(JSON_PATH):
            return cls()
        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            obj = cls()
            obj.tipos = d.get("tipos", {})
            obj.nombres = d.get("nombres", {})
            obj.excluir_proyectos = d.get("excluir_proyectos", [])
            obj.excluir_recursos = d.get("excluir_recursos", [])
            obj.asked_clasif = d.get("asked_clasif", False)
            obj.asked_excl = d.get("asked_excl", False)
            return obj
        except Exception:
            return cls()

    def save(self):
        with open(JSON_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "tipos": self.tipos,
                "nombres": self.nombres,
                "excluir_proyectos": self.excluir_proyectos,
                "excluir_recursos": self.excluir_recursos,
                "asked_clasif": self.asked_clasif,
                "asked_excl": self.asked_excl,
            }, f, ensure_ascii=False, indent=2)

# ======================================================================================
# Lectura Excel
# ======================================================================================

def open_as_xlsx(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return path, load_workbook(path, data_only=True)
    if ext == ".xls":
        # 1) pandas + xlrd; 2) Excel COM
        try:
            import pandas as pd
            xls = pd.ExcelFile(path, engine="xlrd")
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx"); tmp_path = tmp.name; tmp.close()
            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                for sh in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sh, header=None, dtype=str, engine="xlrd")
                    df.to_excel(writer, sheet_name=sh, index=False, header=False)
            return tmp_path, load_workbook(tmp_path, data_only=True)
        except Exception as e:
            try:
                import win32com.client as win32
                excel = win32.Dispatch("Excel.Application"); excel.Visible = False
                wb = excel.Workbooks.Open(os.path.abspath(path))
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx"); tmp_path = tmp.name; tmp.close()
                wb.SaveAs(tmp_path, FileFormat=51); wb.Close(False); excel.Quit()
                return tmp_path, load_workbook(tmp_path, data_only=True)
            except Exception:
                raise RuntimeError("Para .xls: usa pandas+xlrd>=2.0.1 o Excel (pywin32).") from e
    raise ValueError("Extensión no soportada")

def find_all_proyectos_positions(ws) -> List[int]:
    res = []
    for r in range(1, ws.max_row+1):
        for c in range(1, ws.max_column+1):
            if RE_PROYECTOS_TAG.match(read_cell(ws, r, c) or ""):
                res.append(r); break
    return res

def extract_recurso_line(ws, proyectos_row: int) -> str:
    r = proyectos_row - 1
    parts = [read_cell(ws, r, c) for c in range(1, 7)]
    return re.sub(r"\s+", " ", " ".join([p for p in parts if p])).strip()

def detect_day_grid(ws) -> Tuple[int, int]:
    """Devuelve (n_dias, col_inicio_dia1) buscando una fila con la secuencia 1..N."""
    for r in range(1, min(80, ws.max_row)+1):
        c = 1
        while c <= ws.max_column:
            v = read_cell(ws, r, c)
            if v.isdigit() and int(v) == 1:
                start = c
                run = 1
                expect = 2
                k = c + 1
                while k <= ws.max_column:
                    v2 = read_cell(ws, r, k)
                    if v2.isdigit() and int(v2) == expect:
                        run += 1; expect += 1; k += 1
                    else:
                        break
                if run >= 28:  # 28..31
                    return run, start
                c = k
            else:
                c += 1
    return 31, 4  # fallback conservador

# ======================================================================================
# Parseo del bloque
# ======================================================================================

def parse_block(ws, r1, r2, recurso, n_days, day_start) -> List[RowData]:
    """
    Solo filas de imputación entre r1..r2.
    - Proyecto en col 1/2; tipo en la misma fila o posteriores (col 1/2/3).
    - Anti-resumen: índice N de "N-HORA ..." debe crecer por proyecto; si baja, se ignoran
      filas restantes hasta el siguiente proyecto.
    - Horas por columnas absolutas: desde day_start durante n_days.
    """
    rows = []
    proyecto_actual = None
    seq_max = 0
    descartar_hasta_proyecto = False

    def first_tipo_in_cols(r, cols=(1, 2, 3)):
        for c in cols:
            t = read_cell(ws, r, c)
            if RE_TIPO_HORA.match(t or ""):
                return t.strip(), c
        return None, None

    for r in range(r1, r2+1):
        if is_garbage_row(ws, r, day_start, n_days):
            continue

        # ¿Nueva línea con proyecto?
        proj = None; proj_col = None
        for c in (1, 2):
            t = read_cell(ws, r, c)
            p = normalize_project(t)
            if p:
                proj, proj_col = p, c
                break

        if proj:
            proyecto_actual = proj
            seq_max = 0
            descartar_hasta_proyecto = False
            cols_inline = tuple(x for x in (proj_col+1, proj_col+2, 3) if 1 <= x <= max(3, proj_col+2))
            tipo, _ = first_tipo_in_cols(r, cols_inline)
            if tipo:
                m = re.match(r"^\s*(\d+)\s*-\s*", tipo); idx = int(m.group(1)) if m else 0
                seq_max = idx
                horas = [read_cell(ws, r, c) for c in range(day_start, day_start+n_days)]
                rows.append(RowData(
                    recurso=recurso,
                    proyecto_codigo=proyecto_actual[0],
                    proyecto_nombre=proyecto_actual[1],
                    tipo_imputacion=tipo,
                    horas_por_dia=horas
                ))
            continue

        if not proyecto_actual or descartar_hasta_proyecto:
            continue

        # Fila con tipo
        tipo, _ = first_tipo_in_cols(r)
        if tipo:
            m = re.match(r"^\s*(\d+)\s*-\s*", tipo); idx = int(m.group(1)) if m else 0
            if idx <= seq_max:
                # retroceso del índice => sección de resumen. Saltar hasta nuevo proyecto
                descartar_hasta_proyecto = True
                continue
            seq_max = idx
            horas = [read_cell(ws, r, c) for c in range(day_start, day_start+n_days)]
            rows.append(RowData(
                recurso=recurso,
                proyecto_codigo=proyecto_actual[0],
                proyecto_nombre=proyecto_actual[1],
                tipo_imputacion=tipo,
                horas_por_dia=horas
            ))
    return rows

# ======================================================================================
# Salida IPI
# ======================================================================================

def style_header(ws, row_idx, n_days):
    fill = PatternFill("solid", fgColor=COLOR_HEADER)
    bold = Font(bold=True)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    headers = ["RECURSO", "PROYECTO", "TIPO IMPUTACIÓN"] + [str(i) for i in range(1, n_days+1)] + ["TOTAL", "TOTAL DEC", "TIPO PROYECTO"]
    for c, t in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=c, value=t)
        cell.fill = fill; cell.font = bold; cell.alignment = align; cell.border = border
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 22

def style_row(ws, row_idx, color, n_cols):
    if not color:
        return
    fill = PatternFill("solid", fgColor=color)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    for c in range(1, n_cols+1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = fill; cell.border = border

def build_output(wb_out, rows: List[RowData], persist: Persist, n_days: int):
    if SHEET_SALIDA in wb_out.sheetnames:
        del wb_out[SHEET_SALIDA]
    ws = wb_out.create_sheet(SHEET_SALIDA)
    style_header(ws, 1, n_days)
    r0 = 2
    for rd in rows:
        rd.tipo_proyecto = persist.tipos.get(rd.proyecto_codigo, "")
        ws.cell(row=r0, column=1, value=rd.recurso)
        ws.cell(row=r0, column=2, value=f"{rd.proyecto_codigo} - {rd.proyecto_nombre}")
        ws.cell(row=r0, column=3, value=rd.tipo_imputacion or "")

        vals = rd.horas_por_dia[:n_days]
        for i, v in enumerate(vals, start=4):
            ws.cell(row=r0, column=i, value=v or "")
        col_total = 3 + n_days + 1
        col_total_dec = 3 + n_days + 2
        col_tipo = 3 + n_days + 3

        total_min = sum(hhmm_to_minutes(v) for v in vals)
        ws.cell(row=r0, column=col_total, value=minutes_to_hhmm(total_min))
        ws.cell(row=r0, column=col_total_dec, value=round(total_min/60.0, 2))
        ws.cell(row=r0, column=col_tipo, value=rd.tipo_proyecto or "")

        color = COLOR_CONSTRUCCION if rd.tipo_proyecto == "CONSTRUCCION" else (COLOR_REPARACION if rd.tipo_proyecto == "REPARACION" else None)
        style_row(ws, r0, color, col_tipo)
        r0 += 1

    # Totales por tipo de imputación
    sumas: Dict[str, int] = {}
    for rd in rows:
        if not rd.tipo_imputacion:
            continue
        total_min = sum(hhmm_to_minutes(x) for x in rd.horas_por_dia[:n_days])
        sumas[rd.tipo_imputacion] = sumas.get(rd.tipo_imputacion, 0) + total_min
    ws.cell(row=r0+1, column=1, value="TOTALES POR TIPO DE IMPUTACIÓN")
    style_row(ws, r0+1, COLOR_HEADER, 3 + n_days + 3)
    rr = r0 + 2
    for k, v in sorted(sumas.items()):
        ws.cell(row=rr, column=1, value=k)
        ws.cell(row=rr, column=2, value=minutes_to_hhmm(v))
        ws.cell(row=rr, column=3, value=round(v/60.0, 2))
        rr += 1

    # Totales por tipo de proyecto
    sum_tipo = {"CONSTRUCCION": 0, "REPARACION": 0}
    for rd in rows:
        if not rd.tipo_imputacion or rd.tipo_proyecto not in sum_tipo:
            continue
        sum_tipo[rd.tipo_proyecto] += sum(hhmm_to_minutes(x) for x in rd.horas_por_dia[:n_days])

    ws.cell(row=rr + 1, column=1, value="TOTALES POR TIPO DE PROYECTO")
    style_row(ws, rr + 1, COLOR_HEADER, 3 + n_days + 3)
    rtp = rr + 2
    for key in ("CONSTRUCCION", "REPARACION"):
        ws.cell(row=rtp, column=1, value=key)
        ws.cell(row=rtp, column=2, value=minutes_to_hhmm(sum_tipo[key]))
        ws.cell(row=rtp, column=3, value=round(sum_tipo[key] / 60.0, 2))
        rtp += 1

# ======================================================================================
# Pipeline
# ======================================================================================

def collect_discovered(recursos: List[str], proyectos: Dict[str, str], persist: Persist):
    changed = False
    for cod, nom in proyectos.items():
        if cod not in persist.tipos:
            persist.tipos[cod] = ""
            changed = True
        if persist.nombres.get(cod) != nom:
            persist.nombres[cod] = nom
            changed = True
    if changed:
        persist.save()

def process_file(input_path: str, persist: Persist) -> str:
    xlsx_path, wb = open_as_xlsx(input_path); ws = wb.active
    pos = find_all_proyectos_positions(ws)
    if not pos:
        raise RuntimeError("No se encontró 'Proyectos:'")

    n_days, day_start = detect_day_grid(ws)

    # bloques [Proyectos: .. hasta 2 filas antes del siguiente Proyectos:]
    bloques = []
    for i, r in enumerate(pos):
        r1 = r
        r2 = pos[i+1]-2 if i < len(pos)-1 else ws.max_row
        if r1 <= r2:
            bloques.append((r1, r2))

    all_rows = []; recursos = []; proyectos = {}
    for (r1, r2) in bloques:
        recurso = extract_recurso_line(ws, r1) or "RECURSO DESCONOCIDO"
        if recurso not in recursos:
            recursos.append(recurso)
        rows = parse_block(ws, r1, r2, recurso, n_days, day_start)
        for rd in rows:
            if rd.proyecto_codigo and rd.proyecto_nombre:
                proyectos[rd.proyecto_codigo] = rd.proyecto_nombre
        all_rows.extend(rows)

    collect_discovered(recursos, proyectos, persist)

    # exclusiones
    all_rows = [
        rd for rd in all_rows
        if rd.recurso not in persist.excluir_recursos
        and (rd.proyecto_codigo is None or rd.proyecto_codigo not in persist.excluir_proyectos)
    ]

    wb_out = Workbook(); wb_out.remove(wb_out.active)
    build_output(wb_out, all_rows, persist, n_days)
    base = os.path.splitext(os.path.basename(input_path))[0]
    out = os.path.join(os.path.dirname(input_path), f"{base}_IPI.xlsx")
    wb_out.save(out)
    return out

# ======================================================================================
# UI (solo si hay Tkinter)
# ======================================================================================

if GUI_ENABLED:

    def make_scrollframe(parent):
        canvas = tk.Canvas(parent, borderwidth=0)
        frame = ttk.Frame(canvas)
        v = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v.set)
        v.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        return frame

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Transformador Excel → IPI")
            self.geometry("980x560")
            self.persist = Persist.load()
            self.input_path = tk.StringVar(value="")
            self._build()

        def _build(self):
            frm = ttk.Frame(self, padding=10); frm.pack(fill="both", expand=True)
            row0 = ttk.Frame(frm); row0.pack(fill="x", pady=(0, 8))
            ttk.Label(row0, text="Archivo Excel (.xlsx/.xls):").pack(side="left")
            ttk.Entry(row0, textvariable=self.input_path, width=80).pack(side="left", padx=6, fill="x", expand=True)
            ttk.Button(row0, text="Explorar…", command=self.on_browse).pack(side="left")
            row1 = ttk.Frame(frm); row1.pack(fill="x", pady=(0, 8))
            ttk.Button(row1, text="Clasificar proyectos…", command=self.on_edit_clasif).pack(side="left")
            ttk.Button(row1, text="Definir exclusiones…", command=self.on_edit_excl).pack(side="left", padx=6)
            ttk.Button(row1, text="Procesar y generar IPI", command=self.on_process).pack(side="left", padx=12)
            ttk.Label(frm, text="1) Selecciona Excel. 2) Clasifica y excluye. 3) Genera IPI.").pack(anchor="w", pady=(12, 0))

        def on_browse(self):
            p = filedialog.askopenfilename(title="Selecciona Excel",
                                           filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
            if p:
                self.input_path.set(p)

        def on_edit_clasif(self):
            proys = sorted(self.persist.tipos.keys())
            if not proys:
                messagebox.showinfo("Aviso", "No hay proyectos detectados. Carga un Excel primero."); return
            dlg = tk.Toplevel(self); dlg.title("Clasificación de proyectos"); dlg.geometry("760x560")
            container = ttk.Frame(dlg, padding=10); container.pack(fill="both", expand=True)
            sf = make_scrollframe(container)
            ttk.Label(sf, text="CÓDIGO", width=10).grid(row=0, column=0, sticky="w")
            ttk.Label(sf, text="NOMBRE", width=50).grid(row=0, column=1, sticky="w")
            ttk.Label(sf, text="CONSTRUCCIÓN", width=14).grid(row=0, column=2, sticky="w")
            ttk.Label(sf, text="REPARACIÓN", width=12).grid(row=0, column=3, sticky="w")
            vars_con = {}; vars_rep = {}
            def toggle(c, who):
                if who == "CON":
                    if vars_con[c].get(): vars_rep[c].set(0)
                else:
                    if vars_rep[c].get(): vars_con[c].set(0)
            for i, cod in enumerate(proys, start=1):
                nombre = self.persist.nombres.get(cod, "")
                ttk.Label(sf, text=cod, width=10).grid(row=i, column=0, sticky="w")
                ttk.Label(sf, text=nombre, width=50).grid(row=i, column=1, sticky="w")
                v1 = tk.IntVar(value=1 if self.persist.tipos.get(cod, "") == "CONSTRUCCION" else 0)
                v2 = tk.IntVar(value=1 if self.persist.tipos.get(cod, "") == "REPARACION" else 0)
                vars_con[cod] = v1; vars_rep[cod] = v2
                ttk.Checkbutton(sf, variable=v1, command=lambda c=cod: toggle(c, "CON")).grid(row=i, column=2, padx=(6, 18))
                ttk.Checkbutton(sf, variable=v2, command=lambda c=cod: toggle(c, "REP")).grid(row=i, column=3)
            def save_close():
                for cod in proys:
                    if vars_con[cod].get() and not vars_rep[cod].get():
                        self.persist.tipos[cod] = "CONSTRUCCION"
                    elif vars_rep[cod].get() and not vars_con[cod].get():
                        self.persist.tipos[cod] = "REPARACION"
                    else:
                        self.persist.tipos[cod] = ""
                self.persist.asked_clasif = True; self.persist.save(); dlg.destroy()
            ttk.Button(container, text="Guardar", command=save_close).pack(pady=8)

        def on_edit_excl(self):
            if not self.input_path.get():
                messagebox.showwarning("Atención", "Selecciona primero un archivo Excel."); return
            try:
                xlsx, wb = open_as_xlsx(self.input_path.get()); ws = wb.active
                pos = find_all_proyectos_positions(ws)
                recursos = set(); proyectos = {}
                for i, r in enumerate(pos):
                    recurso = extract_recurso_line(ws, r) or "RECURSO DESCONOCIDO"; recursos.add(recurso)
                    r1 = r; r2 = pos[i+1]-2 if i < len(pos)-1 else min(ws.max_row, r+120)
                    for rr in range(r1, r2+1):
                        t = read_cell(ws, rr, 1) or read_cell(ws, rr, 2)
                        proj = normalize_project(t or "")
                        if proj: proyectos[proj[0]] = proj[1]
            except Exception as e:
                messagebox.showerror("Error", str(e)); return
            for k, v in proyectos.items():
                if self.persist.nombres.get(k) != v: self.persist.nombres[k] = v
            self.persist.save()
            recursos = sorted(recursos); proyectos = dict(sorted(proyectos.items()))
            dlg = tk.Toplevel(self); dlg.title("Exclusiones persistentes"); dlg.geometry("980x620")
            container = ttk.Frame(dlg, padding=10); container.pack(fill="both", expand=True)
            lf1 = ttk.Labelframe(container, text="Recursos a excluir"); lf1.pack(side="left", fill="both", expand=True, padx=(0, 6))
            sf1 = make_scrollframe(lf1); rec_vars = {}
            for i, r in enumerate(recursos):
                var = tk.IntVar(value=1 if r in self.persist.excluir_recursos else 0); rec_vars[r] = var
                ttk.Checkbutton(sf1, text=r, variable=var).grid(row=i, column=0, sticky="w", pady=2)
            lf2 = ttk.Labelframe(container, text="Proyectos a excluir"); lf2.pack(side="left", fill="both", expand=True, padx=(6, 0))
            sf2 = make_scrollframe(lf2); proy_vars = {}
            for i, (cod, nom) in enumerate(proyectos.items()):
                var = tk.IntVar(value=1 if cod in self.persist.excluir_proyectos else 0); proy_vars[cod] = var
                ttk.Checkbutton(sf2, text=f"{cod} - {nom}", variable=var).grid(row=i, column=0, sticky="w", pady=2)
            def save_close():
                self.persist.excluir_recursos  = [k for k, v in rec_vars.items() if v.get() == 1]
                self.persist.excluir_proyectos = [k for k, v in proy_vars.items() if v.get() == 1]
                self.persist.asked_excl = True; self.persist.save(); dlg.destroy()
            ttk.Button(container, text="Guardar exclusiones", command=save_close).pack(side="bottom", pady=8)

        def _ensure_prompts(self):
            need_class = any(v == "" for v in self.persist.tipos.values())
            if need_class or not self.persist.asked_clasif:
                if messagebox.askyesno("Clasificar proyectos", "Hay proyectos sin tipo o es la primera vez. ¿Clasificar ahora?"):
                    self.on_edit_clasif()
                else:
                    self.persist.asked_clasif = True; self.persist.save()
            if not self.persist.asked_excl:
                if messagebox.askyesno("Definir exclusiones", "¿Quieres definir exclusiones ahora?"):
                    self.on_edit_excl()
                else:
                    self.persist.asked_excl = True; self.persist.save()

        def on_process(self):
            if not self.input_path.get():
                messagebox.showwarning("Atención", "Selecciona un archivo primero."); return
            self._ensure_prompts()
            try:
                out = process_file(self.input_path.get(), self.persist)
            except Exception as e:
                messagebox.showerror("Error", str(e)); return
            messagebox.showinfo("Listo", f"Generado:\n{out}")

else:
    App = None  # Import seguro en entornos sin GUI

# ======================================================================================
# Lanzador
# ======================================================================================

if __name__ == "__main__":
    if not GUI_ENABLED:
        raise SystemExit("Tkinter no disponible. Usa la versión web.")
    if not os.path.exists(JSON_PATH):
        Persist().save()
    App().mainloop()
