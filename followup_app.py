import gc
import json
import os
import re
import textwrap
import threading
import tkinter as tk
from datetime import datetime
from itertools import islice
from queue import Empty, Queue
from tkinter import filedialog, messagebox, ttk

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib.patches import FancyBboxPatch
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt


def safe_div(numerator: float, denominator: float, *, multiplier: float = 1.0, fallback: float = np.nan) -> float:
    try:
        denom_value = float(denominator)
    except (TypeError, ValueError):
        denom_value = 0.0

    if not np.isfinite(denom_value) or abs(denom_value) < 1e-12:
        return fallback

    try:
        num_value = float(numerator)
    except (TypeError, ValueError):
        num_value = float(np.nan)

    if not np.isfinite(num_value):
        return fallback

    return (num_value / denom_value) * multiplier


plt.rcParams.update(
    {
        "axes.edgecolor": "#bbb",
        "axes.labelcolor": "#000000",
        "xtick.color": "#000000",
        "ytick.color": "#000000",
        "font.size": 10,
        "grid.color": "#d9d9d9",
        "grid.linestyle": ":",
        "axes.grid": True,
    }
)


class FollowUpApp:
    def __init__(self, master):
        self.master = master
        self.master.title("5G Follow-Up Generator - Huawei/VIVO")
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        self.csv_paths = []
        self.tasklines = []
        self.df: pd.DataFrame | None = None
        self.progress_queue: Queue | None = None
        self._worker_thread: threading.Thread | None = None
        self._is_generating = False
        self._tasklines_snapshot: list[tuple[pd.Timestamp, str]] = []
        self.build_interface()

    def validate_date_entry(self, event):
        widget = event.widget
        try:
            datetime.strptime(widget.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Data inválida", "Use o formato AAAA-MM-DD.")
            widget.focus_set()

    def validate_datetime_entry(self, event):
        widget = event.widget
        try:
            datetime.strptime(widget.get(), "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showerror("Data e hora inválidas", "Use o formato AAAA-MM-DD HH:MM.")
            widget.focus_set()

    def build_interface(self):
        frm = ttk.Frame(self.master, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Arquivos CSV (5G1 e 5G2):").pack(anchor="w")
        ttk.Button(frm, text="Selecionar Arquivos", command=self.select_csvs).pack(pady=4)
        self.files_label = ttk.Label(frm, text="Nenhum arquivo selecionado.")
        self.files_label.pack(anchor="w")

        self.filtros: dict[str, tk.Listbox] = {}
        for col in ["gNodeB Name", "Cell Name", "NR Cell ID", "Frequency Band"]:
            ttk.Label(frm, text=f"Filtro: {col}").pack(anchor="w")
            lb = tk.Listbox(frm, selectmode="multiple", height=4, exportselection=False)
            lb.pack(fill="x", pady=2)
            self.filtros[col] = lb

        ttk.Label(frm, text="Intervalo de Datas (AAAA-MM-DD):").pack(anchor="w", pady=(10, 0))
        dtf = ttk.Frame(frm)
        dtf.pack(fill="x")
        self.date_ini = ttk.Entry(dtf, width=12)
        self.date_ini.pack(side="left", padx=5)
        self.date_ini.bind("<FocusOut>", self.validate_date_entry)
        self.date_fim = ttk.Entry(dtf, width=12)
        self.date_fim.pack(side="left", padx=5)
        self.date_fim.bind("<FocusOut>", self.validate_date_entry)

        ttk.Label(frm, text="Filtro: Hora do dia").pack(anchor="w")
        self.hour_filter = tk.Listbox(frm, selectmode="multiple", height=6, exportselection=False)
        for h in range(24):
            self.hour_filter.insert("end", f"{h:02}:00")
        self.hour_filter.pack(fill="x", pady=(0, 6))

        ttk.Label(frm, text="Tasklines (evento + horário):").pack(anchor="w")
        taskf = ttk.Frame(frm)
        taskf.pack(fill="x")
        self.task_entry_time = ttk.Entry(taskf)
        self.task_entry_time.pack(side="left", padx=2)
        self.task_entry_time.bind("<FocusOut>", self.validate_datetime_entry)
        self.task_entry_label = ttk.Entry(taskf)
        self.task_entry_label.insert(0, "Ativação")
        self.task_entry_label.pack(side="left", padx=2)
        ttk.Button(taskf, text="+", width=3, command=self.add_taskline).pack(side="left", padx=2)

        self.task_listbox = tk.Listbox(frm, height=3)
        self.task_listbox.pack(fill="x", pady=4)
        self.task_listbox.bind("<Double-Button-1>", self.remove_taskline)

        ttk.Button(frm, text="Limpar Filtros", command=self.reset_filters).pack(pady=(0, 10))

        self.kpi_title_entry = ttk.Entry(frm)
        self.kpi_title_entry.insert(0, "Nome do Follow-Up (opcional)")
        self.kpi_title_entry.pack(fill="x", pady=(0, 10))

        exportf = ttk.Frame(frm)
        exportf.pack(fill="x", pady=10)
        self.export_png = tk.BooleanVar(value=True)
        self.export_pdf = tk.BooleanVar(value=True)
        self.export_ppt = tk.BooleanVar(value=False)
        ttk.Checkbutton(exportf, text="Exportar PNG", variable=self.export_png).pack(
            side="left", padx=5
        )
        ttk.Checkbutton(exportf, text="Exportar PDF", variable=self.export_pdf).pack(side="left", padx=5)
        ttk.Checkbutton(exportf, text="Exportar PPTX", variable=self.export_ppt).pack(side="left", padx=5)

        ttk.Label(exportf, text="Gráficos por página:").pack(side="left", padx=(20, 5))
        self.page_layout = ttk.Combobox(exportf, values=[4, 6], width=5)
        self.page_layout.set(4)
        self.page_layout.pack(side="left")

        ttk.Button(frm, text="Gerar Relatório", command=self.generate_report).pack(pady=10)

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_frame = ttk.Frame(frm)
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_label.pack(anchor="w")
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, maximum=100, variable=self.progress_var
        )
        self.progress_bar.pack(fill="x", pady=(2, 0))
        self.progress_frame.pack(fill="x")
        self.progress_frame.pack_forget()

    def reset_filters(self):
        for lb in self.filtros.values():
            lb.selection_clear(0, "end")
        self.hour_filter.selection_clear(0, "end")

    def select_csvs(self):
        paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if not paths:
            return

        self.csv_paths = list(paths)
        df = self.merge_csvs()
        self.df = df.copy()
        fnames = [os.path.basename(p) for p in self.csv_paths]
        self.files_label.config(text="Selecionados: " + ", ".join(fnames))

        for col, lb in self.filtros.items():
            lb.delete(0, "end")
            if col in df.columns:
                for val in sorted(df[col].dropna().unique()):
                    lb.insert("end", val)

        if "DATETIME" in df.columns:
            min_dt = df["DATETIME"].min()
            max_dt = df["DATETIME"].max()
            if pd.notna(min_dt) and pd.notna(max_dt):
                self.date_ini.delete(0, "end")
                self.date_ini.insert(0, min_dt.strftime("%Y-%m-%d"))
                self.date_fim.delete(0, "end")
                self.date_fim.insert(0, max_dt.strftime("%Y-%m-%d"))
                self.task_entry_time.delete(0, "end")
                self.task_entry_time.insert(0, min_dt.strftime("%Y-%m-%d %H:00"))

    def add_taskline(self):
        time = self.task_entry_time.get()
        label = self.task_entry_label.get()
        if time:
            try:
                ts = pd.to_datetime(time)
            except ValueError:
                messagebox.showerror("Data inválida", "Use o formato AAAA-MM-DD HH:MM.")
                return
            self.tasklines.append((ts, label))
            self.task_listbox.insert("end", f"{time} - {label}")

    def remove_taskline(self, event):
        sel = self.task_listbox.curselection()
        if sel:
            self.tasklines.pop(sel[0])
            self.task_listbox.delete(sel[0])

    def on_close(self):
        plt.close("all")
        self.master.quit()
        self.master.destroy()

    def get_filtro_nome(self):
        for nome in ["gNodeB Name", "Cell Name", "Frequency Band"]:
            lb = self.filtros.get(nome)
            if not lb:
                continue
            items = lb.curselection()
            if len(items) == 1:
                return lb.get(items[0])
        return None

    def aplicar_filtros(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None:
            return pd.DataFrame()

        filtered = df.copy()
        for col, lb in self.filtros.items():
            if col not in filtered.columns:
                continue
            selections = [lb.get(i) for i in lb.curselection()]
            if selections:
                filtered = filtered[filtered[col].isin(selections)]

        ini = pd.to_datetime(self.date_ini.get(), errors="coerce")
        fim = pd.to_datetime(self.date_fim.get(), errors="coerce")
        if pd.notna(ini) and pd.notna(fim) and "DATETIME" in filtered.columns:
            filtered = filtered[(filtered["DATETIME"] >= ini) & (filtered["DATETIME"] <= fim)]

        selected_hours = [self.hour_filter.get(i) for i in self.hour_filter.curselection()]
        if selected_hours and "DATETIME" in filtered.columns:
            filtered = filtered[filtered["DATETIME"].dt.strftime("%H:00").isin(selected_hours)]

        return filtered

    def clean_csv(self, path: str) -> pd.DataFrame:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        idx = next((i for i, line in enumerate(lines) if line.lower().startswith("time") or line.lower().startswith("day")), 0)
        df = pd.read_csv(path, skiprows=idx, skipfooter=1, engine="python")
        return df

    def merge_csvs(self) -> pd.DataFrame:
        dfs = [self.clean_csv(p) for p in self.csv_paths]
        df = pd.concat(dfs, axis=1)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        if "Time" in df.columns:
            df["DATETIME"] = pd.to_datetime(df["Time"], errors="coerce")
        elif "Day" in df.columns:
            df["DATETIME"] = pd.to_datetime(df["Day"], errors="coerce")
        return df

    def load_kpi_formulas(self) -> dict:
        path = os.path.join(os.getcwd(), "kpi_formulas.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def _replace_source_refs(self, formula: str) -> str:
        pattern = re.compile(r"'[^']+'\[([^\]]+)\]")

        def repl(match: re.Match) -> str:
            column = match.group(1)
            return f"df[{repr(column)}]"

        return pattern.sub(repl, formula)

    def _normalize_formula_output(self, value, kpi_name: str) -> pd.Series | None:
        if isinstance(value, pd.Series):
            numeric = pd.to_numeric(value, errors="coerce")
            numeric = numeric.dropna()
            return numeric if not numeric.empty else None

        if isinstance(value, pd.DataFrame):
            numeric_df = value.select_dtypes(include=[np.number])
            if numeric_df.empty:
                return None
            collapsed = numeric_df.sum(axis=0)
            collapsed = pd.to_numeric(collapsed, errors="coerce").dropna()
            return collapsed if not collapsed.empty else None

        if isinstance(value, dict):
            series = pd.Series(value, dtype="float64")
            series = pd.to_numeric(series, errors="coerce").dropna()
            return series if not series.empty else None

        if isinstance(value, (list, tuple, np.ndarray)):
            if len(value) == 0:
                return None
            series = pd.Series(value, dtype="float64")
            series.index = [f"{kpi_name} {idx}" for idx in range(1, len(series) + 1)]
            series = pd.to_numeric(series, errors="coerce").dropna()
            return series if not series.empty else None

        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            return None

        if not np.isfinite(numeric_value):
            return None

        return pd.Series({kpi_name: numeric_value}, dtype="float64")

    def _evaluate_formula(self, df_filt: pd.DataFrame, info: dict, kpi_name: str) -> pd.DataFrame:
        if "DATETIME" not in df_filt.columns:
            raise ValueError("Coluna DATETIME não encontrada no conjunto de dados filtrado.")

        grouped = df_filt.groupby("DATETIME", sort=True)
        if grouped.ngroups == 0:
            return pd.DataFrame()

        formula = self._replace_source_refs(info["formula"])

        safe_globals: dict[str, object] = {
            "np": np,
            "pd": pd,
            "min": min,
            "max": max,
            "abs": abs,
            "round": round,
            "SAFE_DIV": safe_div,
        }

        rows: list[pd.Series] = []
        index_values: list[pd.Timestamp] = []
        for timestamp, group in grouped:
            local_context = {"df": group}
            try:
                value = eval(formula, safe_globals, local_context)
            except Exception as exc:
                raise ValueError(
                    f"Erro na fórmula do KPI '{kpi_name}': {exc}"
                ) from exc

            series = self._normalize_formula_output(value, kpi_name)
            if series is None:
                continue

            rows.append(series)
            index_values.append(pd.to_datetime(timestamp))

        if not rows:
            return pd.DataFrame()

        df_result = pd.DataFrame(rows, index=index_values)
        df_result = df_result.sort_index()
        df_result = df_result.replace([np.inf, -np.inf], np.nan)
        df_result = df_result.dropna(how="all")
        if df_result.empty:
            return df_result

        df_result = df_result.astype(np.float32)
        return df_result

    @staticmethod
    def _chunked(iterable, size):
        iterator = iter(iterable)
        while True:
            chunk = list(islice(iterator, size))
            if not chunk:
                break
            yield chunk

    def _is_percentage_metric(self, formula: str) -> bool:
        formula = formula or ""
        normalized = re.sub(r"\s+", "", formula)
        if re.search(r"(?i)multiplier=100(?!\d)", normalized):
            return True

        # Captures expressions such as *100, *100.0 or 100*(...) without matching 1000, 10000...
        if re.search(r"\*100(?!\d)", normalized):
            return True

        if re.search(r"(?<!\d)100(?:\.0+)?\*", normalized):
            return True

        return False

    def _load_brand_assets(self) -> dict[str, str]:
        assets_dir = os.path.join(os.getcwd(), "assets")
        os.makedirs(assets_dir, exist_ok=True)
        expected = {
            "huawei": os.path.join(assets_dir, "huawei.png"),
            "vivo": os.path.join(assets_dir, "vivo.png"),
            "background": os.path.join(assets_dir, "mapa-fundo.png"),
        }

        missing = [name for name, path in expected.items() if not os.path.exists(path)]
        if missing:
            missing_names = ", ".join(missing)
            raise FileNotFoundError(
                "Arquivos de branding ausentes: "
                f"{missing_names}. Coloque-os na pasta 'assets/'."
            )

        return expected

    def _gather_report_metadata(self, df_filt: pd.DataFrame) -> dict:
        placeholder_name = "Nome do Follow-Up (opcional)"

        filters_summary: dict[str, str] = {}
        for col, lb in self.filtros.items():
            selections = [lb.get(i) for i in lb.curselection()]
            filters_summary[col] = ", ".join(selections) if selections else "All"

        date_ini = self.date_ini.get().strip()
        date_fim = self.date_fim.get().strip()
        if date_ini and date_fim:
            date_range = f"{date_ini} to {date_fim}"
        elif date_ini or date_fim:
            date_range = date_ini or date_fim
        else:
            date_range = "All"

        selected_hours = [self.hour_filter.get(i) for i in self.hour_filter.curselection()]
        hour_summary = ", ".join(selected_hours) if selected_hours else "All"

        tasklines = []
        for idx, (ts, label) in enumerate(self.tasklines, start=1):
            ts_fmt = ts.strftime("%Y-%m-%d %H:%M")
            tasklines.append(f"{idx}. {label} ({ts_fmt})")

        report_name_entry = self.kpi_title_entry.get().strip()
        report_name = (
            report_name_entry
            if report_name_entry and report_name_entry != placeholder_name
            else ""
        )

        return {
            "report_name": report_name,
            "filters": filters_summary,
            "date_range": date_range,
            "hours": hour_summary,
            "tasklines": tasklines,
            "total_registros": len(df_filt),
        }

    def _add_cover_annotation(self, ax, image_path: str, xy: tuple[float, float], zoom: float):
        if not os.path.exists(image_path):
            return
        image = plt.imread(image_path)
        image_box = OffsetImage(image, zoom=zoom)
        box = AnnotationBbox(image_box, xy, frameon=False)
        ax.add_artist(box)

    def _build_pdf_cover(self, pdf: PdfPages, metadata: dict, logos: dict[str, str], page_size: tuple[float, float]):
        fig, ax = plt.subplots(figsize=page_size)
        ax.axis("off")
        fig.patch.set_facecolor("#ffffff")
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)

        background = plt.imread(logos["background"])
        ax.imshow(
            background,
            extent=[0.5, 1.0, 0.0, 1.0],
            aspect="auto",
            alpha=0.95,
            zorder=0,
            origin="upper",
        )

        self._add_cover_annotation(ax, logos["huawei"], (0.12, 0.94), 0.42)
        self._add_cover_annotation(ax, logos["vivo"], (0.28, 0.94), 0.42)

        ax.text(
            0.08,
            0.80,
            "FOLLOW UP",
            fontsize=40,
            fontweight="bold",
            color="#c3002f",
            va="top",
            ha="left",
        )
        ax.text(
            0.08,
            0.68,
            "5G",
            fontsize=68,
            fontweight="bold",
            color="#c3002f",
            va="top",
            ha="left",
        )

        box = FancyBboxPatch(
            (0.08, 0.07),
            0.34,
            0.26,
            boxstyle="round,pad=0.08",
            linewidth=1.4,
            edgecolor="#c3002f",
            facecolor="#fff7f7",
        )
        ax.add_patch(box)

        summary_lines = []
        if metadata["report_name"]:
            summary_lines.append(metadata["report_name"])
        summary_lines.append(f"Date: {metadata['date_range']}")
        summary_lines.append(f"Hour: {metadata['hours']}")
        label_map = {
            "gNodeB Name": "gNodeB",
            "Cell Name": "Cell",
            "NR Cell ID": "NR Cell ID",
            "Frequency Band": "Frequency Band",
        }
        for key in label_map:
            if key in metadata["filters"]:
                summary_lines.append(f"{label_map[key]}: {metadata['filters'][key]}")

        if metadata["tasklines"]:
            summary_lines.append("Task Lines:")
            summary_lines.extend(f"  {line}" for line in metadata["tasklines"])

        ax.text(
            0.1,
            0.32,
            "\n".join(summary_lines),
            fontsize=11,
            color="#4b1a20",
            va="top",
            ha="left",
            linespacing=1.25,
        )

        pdf.savefig(fig, dpi=400)
        plt.close(fig)

    def _add_pdf_header(self, fig: plt.Figure, metadata: dict, logos: dict[str, str]):
        header_height = 0.1
        fig.subplots_adjust(top=0.82)

        left_ax = fig.add_axes([0.02, 0.86, 0.12, header_height])
        left_ax.axis("off")
        left_ax.set_zorder(-1)
        if os.path.exists(logos["huawei"]):
            left_ax.patch.set_alpha(0.0)
            left_ax.imshow(plt.imread(logos["huawei"]), alpha=0.38, zorder=-1)

        right_ax = fig.add_axes([0.86, 0.86, 0.12, header_height])
        right_ax.axis("off")
        right_ax.set_zorder(-1)
        if os.path.exists(logos["vivo"]):
            right_ax.patch.set_alpha(0.0)
            right_ax.imshow(plt.imread(logos["vivo"]), alpha=0.38, zorder=-1)

        label_map = {
            "gNodeB Name": "gNodeB",
            "Cell Name": "Cell",
            "NR Cell ID": "NR Cell ID",
            "Frequency Band": "Frequency Band",
        }
        filters_summary = [f"Hour: {metadata['hours']}"]
        for key in label_map:
            value = metadata["filters"].get(key)
            if value is not None:
                filters_summary.append(f"{label_map[key]}: {value}")

        header_text = "\n".join(filters_summary)
        fig.text(
            0.5,
            0.94,
            header_text,
            ha="center",
            va="top",
            fontsize=6,
            color="#333333",
        )

    def _add_ppt_cover(self, prs: Presentation, metadata: dict, logos: dict[str, str]):
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)

        if os.path.exists(logos["background"]):
            slide.shapes.add_picture(
                logos["background"],
                prs.slide_width - Inches(8.0),
                Inches(0),
                width=Inches(8.0),
                height=prs.slide_height,
            )

        if os.path.exists(logos["huawei"]):
            slide.shapes.add_picture(logos["huawei"], Inches(0.9), Inches(0.2), height=Inches(1.6))
        if os.path.exists(logos["vivo"]):
            slide.shapes.add_picture(logos["vivo"], Inches(2.4), Inches(0.28), height=Inches(1.4))

        title_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.5), prs.slide_width - Inches(6.0), Inches(2.2))
        title_frame = title_box.text_frame
        title_frame.clear()
        title_frame.text = "FOLLOW UP"
        title_frame.paragraphs[0].font.size = Pt(38)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = RGBColor(195, 0, 47)
        title_frame.paragraphs[0].alignment = 0
        para = title_frame.add_paragraph()
        para.text = "5G"
        para.font.size = Pt(64)
        para.font.bold = True
        para.font.color.rgb = RGBColor(195, 0, 47)
        para.alignment = 0

        summary_shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.9),
            prs.slide_height - Inches(2.7),
            Inches(3.3),
            Inches(2.5),
        )
        fill = summary_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 245, 245)
        line = summary_shape.line
        line.color.rgb = RGBColor(195, 0, 47)
        line.width = Pt(2)

        summary_frame = summary_shape.text_frame
        summary_frame.clear()
        if metadata["report_name"]:
            title_para = summary_frame.paragraphs[0]
            title_para.text = metadata["report_name"]
            title_para.font.size = Pt(20)
            title_para.font.bold = True
            title_para.font.color.rgb = RGBColor(75, 26, 32)
        else:
            summary_frame.text = ""

        para = summary_frame.add_paragraph()
        para.text = f"Date: {metadata['date_range']}"
        para.font.size = Pt(14)
        para.font.color.rgb = RGBColor(75, 26, 32)

        para = summary_frame.add_paragraph()
        para.text = f"Hour: {metadata['hours']}"
        para.font.size = Pt(14)
        para.font.color.rgb = RGBColor(75, 26, 32)

        label_map = {
            "gNodeB Name": "gNodeB",
            "Cell Name": "Cell",
            "NR Cell ID": "NR Cell ID",
            "Frequency Band": "Frequency Band",
        }
        for key, label in label_map.items():
            value = metadata["filters"].get(key)
            if value is not None:
                para = summary_frame.add_paragraph()
                para.text = f"{label}: {value}"
                para.font.size = Pt(13)
                para.font.color.rgb = RGBColor(75, 26, 32)

        if metadata["tasklines"]:
            para = summary_frame.add_paragraph()
            para.text = "Task Lines:"
            para.font.size = Pt(13)
            para.font.color.rgb = RGBColor(75, 26, 32)
            for line_text in metadata["tasklines"]:
                bullet = summary_frame.add_paragraph()
                bullet.text = line_text
                bullet.level = 1
                bullet.font.size = Pt(12)
                bullet.font.color.rgb = RGBColor(75, 26, 32)

    def _render_chart(
        self,
        ax: plt.Axes,
        data: pd.DataFrame,
        kpi_name: str,
        is_percent: bool,
        chart_type: str,
        *,
        style_scale: float = 1.0,
        tasklines: list[tuple[pd.Timestamp, str]] | None = None,
    ) -> None:
        fig = ax.figure
        fig.patch.set_facecolor("#ffffff")
        ax.set_facecolor("#faf3f4")

        for spine in ax.spines.values():
            spine.set_visible(False)

        chart_kind = (chart_type or "line").lower()
        line_color = "#52000f"
        fill_color = "#791021"
        linewidth = 2.4 * max(0.65, style_scale)

        plot_df = data.sort_index().astype(float)
        if plot_df.empty:
            return

        if is_percent:
            finite_values = plot_df.replace([np.inf, -np.inf], np.nan).stack(dropna=True)
            if not finite_values.empty and finite_values.max() <= 1.5:
                plot_df = plot_df * 100.0

        x_values = plot_df.index
        legend_labels: list[str] = []

        if chart_kind == "stacked_bar":
            date_numbers = mdates.date2num(x_values)
            if len(date_numbers) > 1:
                spacing = np.diff(np.sort(date_numbers))
                bar_width = float(np.min(spacing)) * 0.7
            else:
                bar_width = 0.25
            bottoms = np.zeros(len(plot_df))
            palette = [
                "#8b1d3b",
                "#c9433c",
                "#f0624d",
                "#f79352",
                "#fbc87d",
                "#fde0a5",
            ]
            for idx, column in enumerate(plot_df.columns):
                values = plot_df[column].fillna(0).to_numpy()
                color = palette[idx % len(palette)]
                ax.bar(
                    x_values,
                    values,
                    width=bar_width,
                    bottom=bottoms,
                    color=color,
                    edgecolor="#ffffff",
                    linewidth=0.4,
                )
                bottoms = bottoms + values
                legend_labels.append(str(column))
        elif chart_kind == "bar":
            date_numbers = mdates.date2num(x_values)
            if len(date_numbers) > 1:
                spacing = np.diff(np.sort(date_numbers))
                bar_width = float(np.min(spacing)) * 0.6
            else:
                bar_width = 0.18
            column = plot_df.columns[0]
            ax.bar(
                x_values,
                plot_df[column].to_numpy(),
                width=bar_width,
                color=line_color,
                alpha=0.82,
                edgecolor=line_color,
            )
            legend_labels.append(str(column))
        else:
            column = plot_df.columns[0]
            ax.plot(x_values, plot_df[column].to_numpy(), linewidth=linewidth, color=line_color)
            if chart_kind == "area":
                ax.fill_between(x_values, plot_df[column].to_numpy(), color=fill_color, alpha=0.2)
            legend_labels.append(str(column))

        ymin, ymax = ax.get_ylim()
        if np.isfinite(ymin) and np.isfinite(ymax):
            padding = (ymax - ymin) * 0.1 if ymax != ymin else max(abs(ymax), 1.0) * 0.1
        else:
            padding = 1.0
        ax.set_ylim(ymin - padding * 0.12, ymax + padding)

        for ts, label in tasklines or []:
            if x_values.min() <= ts <= x_values.max():
                ax.axvline(ts, color="#8c8c8c", linestyle="--", linewidth=1.0)
                ax.text(
                    ts,
                    ax.get_ylim()[1] - padding * 0.25,
                    label,
                    color="#5f5f5f",
                    rotation=90,
                    ha="right",
                    va="top",
                    fontsize=max(6, 7.2 * style_scale),
                    bbox={
                        "facecolor": "#ffffff",
                        "edgecolor": "none",
                        "alpha": 0.7,
                        "pad": 1,
                    },
                )

        ax.set_xlabel("", color="#000000")
        ax.set_ylabel("", color="#000000")
        ax.tick_params(axis="both", colors="#000000", labelsize=max(7, 9 * style_scale))

        if is_percent:
            ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=1))

        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m %H:%M"))
        ax.margins(x=0.01)

        wrapped_title = textwrap.fill(kpi_name, width=40 if style_scale >= 1.0 else 34)
        title_font = max(8, 11 * style_scale)
        ax.set_title(
            wrapped_title,
            fontweight="bold",
            fontsize=title_font,
            color="#6c0b14",
            pad=20 * style_scale,
        )
        title = ax.title
        title.set_y(1.04 if style_scale < 1.0 else 1.05)

        if legend_labels and chart_kind in {"stacked_bar", "bar"} and len(plot_df.columns) > 1:
            ax.legend(
                legend_labels,
                loc="upper center",
                bbox_to_anchor=(0.5, 1.18),
                ncol=min(len(legend_labels), 4),
                frameon=False,
                fontsize=max(6, 7.5 * style_scale),
            )

        for label in ax.get_xticklabels():
            label.set_rotation(24)
            label.set_horizontalalignment("right")

    def generate_report(self):
        if self._is_generating:
            messagebox.showinfo(
                "Processando",
                "Um relatório já está sendo gerado. Aguarde a finalização antes de iniciar outro.",
            )
            return

        if self.df is None:
            messagebox.showerror("Erro", "Importe os CSVs antes.")
            return

        should_export_png = self.export_png.get()
        should_export_pdf = self.export_pdf.get()
        should_export_ppt = self.export_ppt.get()
        if not any([should_export_png, should_export_pdf, should_export_ppt]):
            messagebox.showerror(
                "Formato não selecionado",
                "Selecione pelo menos um formato de exportação (PNG, PDF ou PPTX).",
            )
            return

        kpis = self.load_kpi_formulas()
        if not kpis:
            messagebox.showerror("Erro", "Arquivo kpi_formulas.json não encontrado.")
            return

        df_filt = self.aplicar_filtros(self.df).copy()
        if df_filt.empty:
            messagebox.showerror("Sem dados", "Nenhum dado encontrado com os filtros selecionados.")
            return

        metadata = self._gather_report_metadata(df_filt)
        self._tasklines_snapshot = list(self.tasklines)
        filtro_nome = self.get_filtro_nome()

        data_hoje = datetime.now().strftime("%Y%m%d")
        followup_name = self.kpi_title_entry.get().strip().replace(" ", "_")
        followup_suffix = f"_{followup_name}" if followup_name else ""
        base_folder = f"FollowUp_5G{followup_suffix}_{data_hoje}"
        out_dir = os.path.join(os.getcwd(), base_folder)
        os.makedirs(out_dir, exist_ok=True)

        per_page = int(self.page_layout.get())

        self.progress_queue = Queue()
        self._is_generating = True
        self._show_progress("Preparando geração...", 0)

        worker_args = (
            df_filt,
            metadata,
            kpis,
            bool(should_export_png),
            bool(should_export_pdf),
            bool(should_export_ppt),
            out_dir,
            followup_suffix,
            data_hoje,
            filtro_nome,
            list(self._tasklines_snapshot),
            per_page,
        )

        self._worker_thread = threading.Thread(
            target=self._run_report_generation,
            args=worker_args,
            daemon=True,
        )
        self._worker_thread.start()
        self.master.after(120, self._poll_progress_queue)

    def _run_report_generation(
        self,
        df_filt: pd.DataFrame,
        metadata: dict,
        kpis: dict,
        should_export_png: bool,
        should_export_pdf: bool,
        should_export_ppt: bool,
        out_dir: str,
        followup_suffix: str,
        data_hoje: str,
        filtro_nome: str | None,
        tasklines: list[tuple[pd.Timestamp, str]],
        per_page: int,
    ) -> None:
        charts: list[dict[str, object]] = []
        try:
            total_kpis = max(len(kpis), 1)
            for processed_kpis, (kpi_name, info) in enumerate(kpis.items(), start=1):
                progress_pct = (processed_kpis / total_kpis) * 70
                self._queue_progress(
                    f"Gerando gráfico {processed_kpis}/{total_kpis}...",
                    progress_pct,
                )
                try:
                    result_df = self._evaluate_formula(df_filt, info, kpi_name)
                except Exception as exc:
                    print(f"Erro no KPI {kpi_name}: {exc}")
                    continue

                if result_df.empty:
                    continue

                result_df = result_df.astype(np.float32)
                is_percent = self._is_percentage_metric(info.get("formula", ""))
                chart_type = info.get("chart_type", "line")

                fig, ax = plt.subplots(figsize=(9.4, 5.6))
                self._render_chart(
                    ax,
                    result_df,
                    kpi_name,
                    is_percent,
                    chart_type,
                    style_scale=1.0,
                    tasklines=tasklines,
                )
                fig.subplots_adjust(left=0.06, right=0.98, bottom=0.14, top=0.86)

                nome_suffix = f"_{filtro_nome}" if filtro_nome else ""
                nome_base = (
                    f"FollowUp_5G{followup_suffix}_{kpi_name.replace(' ', '_')}{nome_suffix}_{data_hoje}"
                )
                img_path = os.path.join(out_dir, nome_base + ".png")
                fig.savefig(img_path, dpi=400, bbox_inches="tight", pad_inches=0.25)
                charts.append(
                    {
                        "image_path": img_path,
                        "data": result_df.copy(),
                        "name": kpi_name,
                        "is_percent": is_percent,
                        "chart_type": chart_type,
                    }
                )
                plt.close(fig)

            if not charts:
                self.progress_queue.put(
                    {"type": "done", "status": "error", "message": "Nenhum gráfico foi gerado."}
                )
                return

            try:
                logos = self._load_brand_assets()
            except Exception as exc:
                self.progress_queue.put(
                    {"type": "done", "status": "error", "message": str(exc)}
                )
                return

            generated_messages: list[str] = []
            rows = 2
            cols = 2 if per_page == 4 else 3
            page_size = (11.69, 8.27)

            if should_export_pdf:
                self._queue_progress("Montando PDF...", 75)
                pdf_path = os.path.join(
                    out_dir, f"FollowUp_5G{followup_suffix}_{data_hoje}.pdf"
                )
                style_scale = 0.8 if per_page == 4 else 0.62
                with PdfPages(pdf_path) as pdf:
                    self._build_pdf_cover(pdf, metadata, logos, page_size)
                    for chunk in self._chunked(charts, per_page):
                        fig, axs = plt.subplots(rows, cols, figsize=page_size)
                        axs_array = np.array(axs).reshape(rows * cols)
                        for ax in axs_array:
                            ax.clear()
                            ax.axis("off")
                        for ax, chart in zip(axs_array, chunk):
                            ax.axis("on")
                            self._render_chart(
                                ax,
                                chart["data"],
                                chart["name"],
                                bool(chart["is_percent"]),
                                chart.get("chart_type", "line"),
                                style_scale=style_scale,
                                tasklines=tasklines,
                            )
                        fig.subplots_adjust(
                            left=0.08,
                            right=0.97,
                            bottom=0.12,
                            top=0.78,
                            wspace=0.38,
                            hspace=0.58,
                        )
                        self._add_pdf_header(fig, metadata, logos)
                        pdf.savefig(fig, dpi=400, bbox_inches="tight")
                        plt.close(fig)
                generated_messages.append(f"PDF: {os.path.basename(pdf_path)}")

            for chart in charts:
                chart.pop("data", None)

            if should_export_ppt:
                self._queue_progress("Montando PPTX...", 85)
                ppt_path = os.path.join(
                    out_dir, f"FollowUp_5G{followup_suffix}_{data_hoje}.pptx"
                )
                prs = Presentation()
                prs.slide_width = Inches(13.33)
                prs.slide_height = Inches(7.5)
                blank = prs.slide_layouts[6]

                self._add_ppt_cover(prs, metadata, logos)

                margin_w = Inches(0.4)
                margin_h = Inches(0.4)
                spacing_w = Inches(0.2)
                spacing_h = Inches(0.2)

                usable_width = prs.slide_width - 2 * margin_w - (cols - 1) * spacing_w
                usable_height = prs.slide_height - 2 * margin_h - (rows - 1) * spacing_h
                img_width = usable_width / cols
                img_height = usable_height / rows

                for chunk in self._chunked([chart["image_path"] for chart in charts], per_page):
                    slide = prs.slides.add_slide(blank)
                    for idx, img_path in enumerate(chunk):
                        row = idx // cols
                        col = idx % cols
                        left = margin_w + col * (img_width + spacing_w)
                        top = margin_h + row * (img_height + spacing_h)
                        slide.shapes.add_picture(
                            img_path, left, top, width=img_width, height=img_height
                        )
                prs.save(ppt_path)
                generated_messages.append(f"PPTX: {os.path.basename(ppt_path)}")

            if should_export_png:
                generated_messages.append("PNGs gerados com sucesso.")

            self._queue_progress("Finalizando...", 100)
            self.progress_queue.put(
                {"type": "done", "status": "success", "messages": generated_messages}
            )

        except PermissionError:
            self.progress_queue.put(
                {
                    "type": "done",
                    "status": "error",
                    "message": "O arquivo está aberto em outro programa. Feche-o e tente novamente.",
                }
            )
        except Exception as exc:
            self.progress_queue.put(
                {"type": "done", "status": "error", "message": str(exc)}
            )
        finally:
            if not should_export_png:
                for chart in charts:
                    img_path = chart.get("image_path")
                    if img_path and os.path.exists(img_path):
                        try:
                            os.remove(img_path)
                        except OSError:
                            pass
            charts.clear()
            plt.close("all")
            gc.collect()

    def _queue_progress(self, message: str, value: float) -> None:
        if self.progress_queue is not None:
            self.progress_queue.put(
                {"type": "progress", "message": message, "value": float(value)}
            )

    def _poll_progress_queue(self):
        queue = self.progress_queue
        if queue is None:
            return

        done_event = False
        try:
            while True:
                event = queue.get_nowait()
                event_type = event.get("type")
                if event_type == "progress":
                    self._show_progress(event.get("message", ""), event.get("value", 0.0))
                elif event_type == "done":
                    done_event = True
                    status = event.get("status")
                    messages = event.get("messages", [])
                    error_message = event.get("message", "")
                    self._is_generating = False
                    self._worker_thread = None
                    self._hide_progress()
                    if status == "success":
                        info_lines = "\n".join(messages)
                        messagebox.showinfo(
                            "Concluído",
                            "Relatório gerado com sucesso!\n" + info_lines
                            if info_lines
                            else "Relatório gerado com sucesso!",
                        )
                    else:
                        messagebox.showerror(
                            "Erro",
                            error_message or "Ocorreu um erro na geração do relatório.",
                        )
                    break
        except Empty:
            pass

        if done_event:
            self.progress_queue = None
        elif self._is_generating:
            self.master.after(120, self._poll_progress_queue)


    def _show_progress(self, message: str, value: float):
        self.progress_frame.pack(fill="x", pady=(5, 0))
        self.progress_label.config(text=message)
        self.progress_var.set(max(0.0, min(100.0, value)))
        self.master.update_idletasks()

    def _hide_progress(self):
        self.progress_var.set(0.0)
        self.progress_label.config(text="")
        self.progress_frame.pack_forget()


if __name__ == "__main__":
    root = tk.Tk()
    app = FollowUpApp(root)
    root.mainloop()
