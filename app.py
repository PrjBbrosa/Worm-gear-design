
import os, json, math
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.patches as mpatches
import matplotlib.patheffects as patheffects
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import font_manager
from matplotlib.colors import Normalize
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

from src.utils import load_json
from src.worm_model import compute_worm_cycle
from src.export_xlsx import export_cycle_xlsx

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
APP_TITLE = "WormGear Studio — 蜗轮蜗杆设计校核"

# Apple-inspired colour palette
CLR_BG       = "#F5F5F7"   # light grey background
CLR_CARD     = "#FFFFFF"   # white cards
CLR_ACCENT   = "#007AFF"   # iOS blue
CLR_ACCENT2  = "#34C759"   # iOS green
CLR_TEXT      = "#1D1D1F"   # near-black text
CLR_TEXT2     = "#6E6E73"   # secondary text
CLR_BORDER   = "#D2D2D7"   # subtle border
CLR_INPUT_BG = "#FFFFFF"
CLR_SECTION  = "#F2F2F7"   # section bg
CLR_WORM     = "#FF3B30"   # red for worm
CLR_WHEEL    = "#007AFF"   # blue for wheel
CLR_DIM      = "#8E8E93"   # dimension lines

# ---------------------------------------------------------------------------
# Font setup
# ---------------------------------------------------------------------------
def setup_fonts(root=None):
    candidates = [
        "PingFang SC", "Hiragino Sans GB", "STHeiti", "Heiti SC",
        "Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
        "WenQuanYi Zen Hei", "Arial Unicode MS"
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = None
    for name in candidates:
        if name in available:
            chosen = name
            break
    if chosen is None:
        chosen = "DejaVu Sans"

    if root is not None:
        try:
            for font_name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
                tkfont.nametofont(font_name).configure(family=chosen, size=11)
        except Exception:
            pass

    matplotlib.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False
    return chosen


def list_materials(folder):
    out = []
    if not os.path.isdir(folder):
        return out
    for fn in os.listdir(folder):
        if fn.lower().endswith(".json"):
            out.append(os.path.join(folder, fn))
    return sorted(out)


# ---------------------------------------------------------------------------
# Rounded-rect card helper for matplotlib
# ---------------------------------------------------------------------------
def _rounded_rect(ax, x, y, w, h, r=0.08, **kwargs):
    from matplotlib.patches import FancyBboxPatch
    patch = FancyBboxPatch((x, y), w, h,
                           boxstyle=f"round,pad={r}",
                           **kwargs)
    ax.add_patch(patch)
    return patch


# ===================================================================
# Main Application
# ===================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self._font_family = setup_fonts(self)
        self._setup_style()
        self.title(APP_TITLE)
        self.geometry("1520x920")
        self.minsize(1200, 750)

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.steel_paths = list_materials(os.path.join(self.base_dir, "materials", "metals"))
        self.poly_paths  = list_materials(os.path.join(self.base_dir, "materials", "polymers"))
        if not self.steel_paths or not self.poly_paths:
            raise RuntimeError("materials 目录缺失，请保持项目完整解压。")

        self.steel_db = {os.path.basename(p): p for p in self.steel_paths}
        self.poly_db  = {os.path.basename(p): p for p in self.poly_paths}

        self.steel = load_json(self.steel_db.get("37CrS4.json", self.steel_paths[0]))
        self.wheel = load_json(self.poly_db.get("PA66_modified_draft.json", self.poly_paths[0]))

        # Default values for all inputs
        self._defaults = {
            # Drive
            "T1_Nm": "6.0", "n1_rpm": "3000", "ratio": "25",
            "life_h": "3000", "steps": "720",
            # Worm geometry
            "z1": "2", "mn_mm": "2.5", "q": "10", "x1": "0.0",
            "alpha_n_deg": "20", "rho_f_mm": "0.6",
            "da1_mm": "", "df1_mm": "", "d1_mm": "",
            # Wheel geometry
            "z2": "", "x2": "0.0", "d2_mm": "", "da2_mm": "", "df2_mm": "",
            "a_target_mm": "", "b2_mm": "18",
            # Common
            "cross_angle_deg": "90", "mu": "0.06",
            # Correction factors
            "KA": "1.10", "KV": "1.05", "KHb": "1.00", "KFb": "1.00",
            # Environment
            "temp_C": "80",
        }
        self.inputs = {}
        self.res = None
        self.sn_rows = []
        self._cloud_cbar = None

        self._build_menu()
        self._build_ui()
        self._auto_calc_worm()
        self._auto_calc_wheel()
        self.refresh_geom_plot()

    # ------------------------------------------------------------------
    # Style
    # ------------------------------------------------------------------
    def _setup_style(self):
        self.configure(bg=CLR_BG)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background=CLR_BG)
        style.configure("Card.TFrame", background=CLR_CARD)
        style.configure("TLabel", background=CLR_BG, foreground=CLR_TEXT, font=("", 11))
        style.configure("Card.TLabel", background=CLR_CARD, foreground=CLR_TEXT, font=("", 11))
        style.configure("Header.TLabel", background=CLR_BG, foreground=CLR_TEXT, font=("", 13, "bold"))
        style.configure("CardHeader.TLabel", background=CLR_CARD, foreground=CLR_TEXT, font=("", 12, "bold"))
        style.configure("Sub.TLabel", background=CLR_CARD, foreground=CLR_TEXT2, font=("", 10))
        style.configure("Unit.TLabel", background=CLR_CARD, foreground=CLR_DIM, font=("", 10))
        style.configure("Status.TLabel", background=CLR_BG, foreground=CLR_ACCENT, font=("", 10))

        style.configure("TLabelframe", background=CLR_CARD, borderwidth=0, relief="flat")
        style.configure("TLabelframe.Label", background=CLR_CARD, foreground=CLR_TEXT, font=("", 11, "bold"))

        # Apple-style buttons
        style.configure("Accent.TButton", padding=(14, 7), font=("", 11, "bold"))
        style.map("Accent.TButton",
                  background=[("active", "#005EC4"), ("!active", CLR_ACCENT)],
                  foreground=[("active", "#FFF"), ("!active", "#FFF")])
        style.configure("TButton", padding=(10, 6), font=("", 10))
        style.map("TButton", background=[("active", "#E8E8ED")])

        style.configure("Small.TButton", padding=(6, 3), font=("", 9))

        style.configure("TNotebook", background=CLR_BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(20, 10), font=("", 11))
        style.map("TNotebook.Tab",
                  background=[("selected", CLR_CARD), ("!selected", CLR_SECTION)],
                  foreground=[("selected", CLR_ACCENT), ("!selected", CLR_TEXT2)])

        style.configure("Treeview", rowheight=28, font=("", 10), background=CLR_CARD,
                        fieldbackground=CLR_CARD, borderwidth=0)
        style.configure("Treeview.Heading", font=("", 10, "bold"), background=CLR_SECTION)
        style.map("Treeview", background=[("selected", "#D1E8FF")])

        # Entry style
        style.configure("TEntry", fieldbackground=CLR_INPUT_BG, borderwidth=1,
                        relief="solid", padding=(6, 4))

    # ------------------------------------------------------------------
    # Menu
    # ------------------------------------------------------------------
    def _build_menu(self):
        m = tk.Menu(self, bg=CLR_CARD, fg=CLR_TEXT, activebackground=CLR_ACCENT,
                    activeforeground="#FFF", bd=0)
        fm = tk.Menu(m, tearoff=0, bg=CLR_CARD, fg=CLR_TEXT)
        fm.add_command(label="  导出 XLSX（曲线）...", command=self.export_xlsx)
        fm.add_separator()
        fm.add_command(label="  退出", command=self.destroy)
        m.add_cascade(label="  文件  ", menu=fm)
        self.config(menu=m)

    # ------------------------------------------------------------------
    # UI skeleton
    # ------------------------------------------------------------------
    def _build_ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=(6, 12))

        self.tab_geom = ttk.Frame(self.nb)
        self.tab_mat  = ttk.Frame(self.nb)
        self.tab_res  = ttk.Frame(self.nb)
        self.tab_fat  = ttk.Frame(self.nb)

        self.nb.add(self.tab_geom, text="  几何参数  ")
        self.nb.add(self.tab_mat,  text="  材料与S-N  ")
        self.nb.add(self.tab_res,  text="  应力与效率  ")
        self.nb.add(self.tab_fat,  text="  寿命校核  ")

        self._build_geom_tab()
        self._build_mat_tab()
        self._build_res_tab()
        self._build_fat_tab()

    # ==================================================================
    # Helper: create a labelled entry row inside a card
    # ==================================================================
    def _entry(self, parent, key, label, unit="", width=14, readonly=False):
        row = tk.Frame(parent, bg=CLR_CARD)
        row.pack(fill="x", padx=12, pady=3)
        lbl = tk.Label(row, text=label, bg=CLR_CARD, fg=CLR_TEXT, font=("", 10),
                       anchor="w", width=20)
        lbl.pack(side="left")
        var = tk.StringVar(value=self._defaults.get(key, ""))
        state = "readonly" if readonly else "normal"
        ent = tk.Entry(row, textvariable=var, width=width, font=("", 10),
                       relief="solid", bd=1, bg=CLR_INPUT_BG if not readonly else CLR_SECTION,
                       fg=CLR_TEXT, highlightthickness=1,
                       highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER,
                       state=state)
        ent.pack(side="left", padx=(4, 0))
        if unit:
            tk.Label(row, text=unit, bg=CLR_CARD, fg=CLR_DIM, font=("", 9),
                     width=6, anchor="w").pack(side="left", padx=(6, 0))
        self.inputs[key] = var
        return row

    def _auto_btn(self, parent, text, command):
        """Small auto-calculate button placed inline."""
        btn = tk.Button(parent, text=text, command=command,
                        font=("", 9), fg=CLR_ACCENT, bg=CLR_CARD,
                        activebackground="#E8E8ED", activeforeground=CLR_ACCENT,
                        relief="flat", bd=0, padx=6, pady=1, cursor="hand2")
        btn.pack(side="left", padx=(6, 0))
        return btn

    def _section_label(self, parent, text, bg=CLR_CARD):
        lbl = tk.Label(parent, text=text, bg=bg, fg=CLR_TEXT,
                       font=("", 11, "bold"), anchor="w")
        lbl.pack(fill="x", padx=12, pady=(10, 2))

    def _separator(self, parent):
        sep = tk.Frame(parent, bg=CLR_BORDER, height=1)
        sep.pack(fill="x", padx=12, pady=6)

    # ==================================================================
    # Tab 1: Geometry input (left: worm + wheel cards, right: diagram)
    # ==================================================================
    def _build_geom_tab(self):
        # Main horizontal split
        left = tk.Frame(self.tab_geom, bg=CLR_BG)
        left.pack(side="left", fill="y", padx=(8, 4), pady=8)
        right = tk.Frame(self.tab_geom, bg=CLR_BG)
        right.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=8)

        # Scrollable left panel
        left_outer = tk.Frame(left, bg=CLR_BG)
        left_outer.pack(fill="both", expand=True)

        self.geom_canvas = tk.Canvas(left_outer, width=480, highlightthickness=0, bg=CLR_BG, bd=0)
        geom_scroll = ttk.Scrollbar(left_outer, orient="vertical", command=self.geom_canvas.yview)
        self.geom_canvas.configure(yscrollcommand=geom_scroll.set)
        geom_scroll.pack(side="right", fill="y")
        self.geom_canvas.pack(side="left", fill="both", expand=True)

        scroll_frame = tk.Frame(self.geom_canvas, bg=CLR_BG)
        self._geom_window = self.geom_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_frame.bind("<Configure>", lambda e: self.geom_canvas.configure(scrollregion=self.geom_canvas.bbox("all")))
        self.geom_canvas.bind("<Configure>", lambda e: self.geom_canvas.itemconfigure(self._geom_window, width=e.width))

        # Mouse wheel
        def _on_mw(event):
            if self.nb.index(self.nb.select()) == 0:
                self.geom_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.geom_canvas.bind_all("<MouseWheel>", _on_mw)
        # Linux scroll
        self.geom_canvas.bind_all("<Button-4>", lambda e: self.geom_canvas.yview_scroll(-3, "units"))
        self.geom_canvas.bind_all("<Button-5>", lambda e: self.geom_canvas.yview_scroll(3, "units"))

        # ---- Drive Parameters Card ----
        drive_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0, highlightbackground=CLR_BORDER,
                              highlightthickness=1)
        drive_card.pack(fill="x", padx=6, pady=(6, 4))
        self._section_label(drive_card, "驱动参数")
        self._entry(drive_card, "T1_Nm", "输入扭矩 T₁", "N·m")
        self._entry(drive_card, "n1_rpm", "蜗杆转速 n₁", "rpm")
        r = self._entry(drive_card, "ratio", "传动比 i", "")
        self._auto_btn(r, "⟲ 由z算i", self._calc_ratio_from_z)
        self._entry(drive_card, "life_h", "目标寿命", "h")
        self._entry(drive_card, "steps", "相位点数", "点")
        tk.Frame(drive_card, bg=CLR_CARD, height=6).pack()

        # ---- Worm (蜗杆) Card ----
        worm_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0, highlightbackground=CLR_WORM,
                             highlightthickness=2)
        worm_card.pack(fill="x", padx=6, pady=4)

        worm_header = tk.Frame(worm_card, bg=CLR_CARD)
        worm_header.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(worm_header, text="● 蜗杆参数", bg=CLR_CARD, fg=CLR_WORM,
                 font=("", 12, "bold")).pack(side="left")
        tk.Button(worm_header, text="⟲ 自动计算蜗杆尺寸", command=self._auto_calc_worm,
                  font=("", 9, "bold"), fg="#FFF", bg=CLR_WORM,
                  activebackground="#D32F2F", activeforeground="#FFF",
                  relief="flat", bd=0, padx=10, pady=3, cursor="hand2").pack(side="right")

        self._section_label(worm_card, "基本参数")
        self._entry(worm_card, "z1", "头数 z₁", "")
        self._entry(worm_card, "mn_mm", "法向模数 mₙ", "mm")
        self._entry(worm_card, "q", "直径系数 q", "")
        self._entry(worm_card, "x1", "变位系数 x₁", "")
        self._entry(worm_card, "alpha_n_deg", "法向压力角 αₙ", "deg")
        self._entry(worm_card, "rho_f_mm", "齿根圆角 ρf", "mm")

        self._separator(worm_card)
        self._section_label(worm_card, "计算尺寸（自动）")
        self._entry(worm_card, "d1_mm", "分度圆直径 d₁", "mm", readonly=True)
        self._entry(worm_card, "da1_mm", "齿顶圆直径 dₐ₁", "mm", readonly=True)
        self._entry(worm_card, "df1_mm", "齿根圆直径 df₁", "mm", readonly=True)
        tk.Frame(worm_card, bg=CLR_CARD, height=6).pack()

        # ---- Wheel (蜗轮) Card ----
        wheel_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0, highlightbackground=CLR_WHEEL,
                              highlightthickness=2)
        wheel_card.pack(fill="x", padx=6, pady=4)

        wheel_header = tk.Frame(wheel_card, bg=CLR_CARD)
        wheel_header.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(wheel_header, text="● 蜗轮参数", bg=CLR_CARD, fg=CLR_WHEEL,
                 font=("", 12, "bold")).pack(side="left")
        tk.Button(wheel_header, text="⟲ 自动计算蜗轮尺寸", command=self._auto_calc_wheel,
                  font=("", 9, "bold"), fg="#FFF", bg=CLR_WHEEL,
                  activebackground="#005EC4", activeforeground="#FFF",
                  relief="flat", bd=0, padx=10, pady=3, cursor="hand2").pack(side="right")

        self._section_label(wheel_card, "基本参数")
        r = self._entry(wheel_card, "z2", "齿数 z₂", "")
        self._auto_btn(r, "⟲ 由i算z₂", self._calc_z2_from_ratio)
        self._entry(wheel_card, "x2", "变位系数 x₂", "")
        self._entry(wheel_card, "b2_mm", "齿宽 b₂", "mm")
        r2 = self._entry(wheel_card, "a_target_mm", "目标中心距 a", "mm")
        self._auto_btn(r2, "⟲ 自动算a", self._calc_center_distance)

        self._separator(wheel_card)
        self._section_label(wheel_card, "计算尺寸（自动）")
        self._entry(wheel_card, "d2_mm", "分度圆直径 d₂", "mm", readonly=True)
        self._entry(wheel_card, "da2_mm", "齿顶圆直径 dₐ₂", "mm", readonly=True)
        self._entry(wheel_card, "df2_mm", "齿根圆直径 df₂", "mm", readonly=True)
        tk.Frame(wheel_card, bg=CLR_CARD, height=6).pack()

        # ---- Common Parameters Card ----
        common_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0, highlightbackground=CLR_BORDER,
                               highlightthickness=1)
        common_card.pack(fill="x", padx=6, pady=4)
        self._section_label(common_card, "公共参数与修正系数")
        self._entry(common_card, "cross_angle_deg", "交错角 Σ", "deg")
        self._entry(common_card, "mu", "摩擦系数 μ", "")
        self._separator(common_card)
        self._entry(common_card, "KA", "使用系数 Kₐ", "")
        self._entry(common_card, "KV", "动载系数 Kᵥ", "")
        self._entry(common_card, "KHb", "齿宽载荷系数 KHβ", "")
        self._entry(common_card, "KFb", "齿根载荷系数 KFβ", "")
        self._entry(common_card, "temp_C", "工作温度", "°C")
        tk.Frame(common_card, bg=CLR_CARD, height=6).pack()

        # ---- Action buttons ----
        btn_frame = tk.Frame(scroll_frame, bg=CLR_BG)
        btn_frame.pack(fill="x", padx=6, pady=(8, 12))

        tk.Button(btn_frame, text="更新示意图", command=self._on_refresh_diagram,
                  font=("", 11), fg=CLR_TEXT, bg=CLR_CARD,
                  activebackground="#E8E8ED", relief="flat", bd=0,
                  padx=16, pady=8, cursor="hand2",
                  highlightbackground=CLR_BORDER, highlightthickness=1).pack(side="left", padx=(0, 8))

        tk.Button(btn_frame, text="计算并绘图", command=self.run,
                  font=("", 11, "bold"), fg="#FFF", bg=CLR_ACCENT,
                  activebackground="#005EC4", activeforeground="#FFF",
                  relief="flat", bd=0, padx=20, pady=8, cursor="hand2").pack(side="left")

        # ---- Right side: diagram ----
        diag_card = tk.Frame(right, bg=CLR_CARD, bd=0, highlightbackground=CLR_BORDER,
                             highlightthickness=1)
        diag_card.pack(fill="both", expand=True)

        fig = Figure(figsize=(8, 7), dpi=100, facecolor=CLR_CARD)
        self.ax_geom = fig.add_subplot(111)
        self.canvas_geom = FigureCanvasTkAgg(fig, master=diag_card)
        self.canvas_geom.get_tk_widget().pack(fill="both", expand=True, padx=2, pady=2)

        self.geom_check_var = tk.StringVar(value="几何校核信息：等待输入参数。")
        tk.Label(right, textvariable=self.geom_check_var, bg=CLR_BG, fg=CLR_ACCENT,
                 font=("", 10), anchor="w").pack(fill="x", padx=4, pady=(4, 0))

    # ==================================================================
    # Auto-calculation functions
    # ==================================================================
    def _safe_float(self, key, default=0.0):
        try:
            return float(self.inputs[key].get())
        except (ValueError, KeyError):
            return default

    def _auto_calc_worm(self):
        """Calculate worm dimensions from basic parameters."""
        mn = self._safe_float("mn_mm", 2.5)
        q = self._safe_float("q", 10)
        x1 = self._safe_float("x1", 0.0)

        d1 = (q + 2.0 * x1) * mn
        da1 = d1 + 2.0 * mn
        df1 = d1 - 2.4 * mn

        self.inputs["d1_mm"].set(f"{d1:.3f}")
        self.inputs["da1_mm"].set(f"{da1:.3f}")
        self.inputs["df1_mm"].set(f"{df1:.3f}")

    def _auto_calc_wheel(self):
        """Calculate wheel dimensions from basic parameters."""
        mn = self._safe_float("mn_mm", 2.5)
        x2 = self._safe_float("x2", 0.0)
        z2_txt = self.inputs.get("z2", tk.StringVar(value=""))
        try:
            z2 = int(float(z2_txt.get()))
        except (ValueError, TypeError):
            ratio = self._safe_float("ratio", 25)
            z1 = int(self._safe_float("z1", 2))
            z2 = int(round(ratio * z1))
            self.inputs["z2"].set(str(z2))

        d2 = (z2 + 2.0 * x2) * mn
        da2 = d2 + 2.0 * mn * (1.0 + x2)
        df2 = d2 - 2.0 * mn * (1.2 - x2)

        self.inputs["d2_mm"].set(f"{d2:.3f}")
        self.inputs["da2_mm"].set(f"{da2:.3f}")
        self.inputs["df2_mm"].set(f"{df2:.3f}")

    def _calc_ratio_from_z(self):
        z1 = int(self._safe_float("z1", 2))
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            return
        if z1 > 0:
            self.inputs["ratio"].set(f"{z2 / z1:.2f}")

    def _calc_z2_from_ratio(self):
        ratio = self._safe_float("ratio", 25)
        z1 = int(self._safe_float("z1", 2))
        z2 = int(round(ratio * z1))
        self.inputs["z2"].set(str(z2))
        self._auto_calc_wheel()

    def _calc_center_distance(self):
        mn = self._safe_float("mn_mm", 2.5)
        q = self._safe_float("q", 10)
        x1 = self._safe_float("x1", 0.0)
        x2 = self._safe_float("x2", 0.0)
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            ratio = self._safe_float("ratio", 25)
            z1 = int(self._safe_float("z1", 2))
            z2 = int(round(ratio * z1))
        a = 0.5 * mn * (q + z2 + 2.0 * (x1 + x2))
        self.inputs["a_target_mm"].set(f"{a:.3f}")

    def _on_refresh_diagram(self):
        self._auto_calc_worm()
        self._auto_calc_wheel()
        self.refresh_geom_plot()

    # ==================================================================
    # Geometry diagram - proper engineering cross-section
    # ==================================================================
    def refresh_geom_plot(self):
        ax = self.ax_geom
        ax.clear()
        ax.set_aspect("equal", adjustable="datalim")
        ax.axis("off")
        fig = ax.get_figure()
        fig.set_facecolor(CLR_CARD)
        ax.set_facecolor(CLR_CARD)

        # Read parameters
        mn = self._safe_float("mn_mm", 2.5)
        q = self._safe_float("q", 10)
        x1 = self._safe_float("x1", 0.0)
        x2 = self._safe_float("x2", 0.0)
        z1 = int(self._safe_float("z1", 2))
        ratio = self._safe_float("ratio", 25)
        z2 = int(round(ratio * z1))
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            pass

        d1 = (q + 2.0 * x1) * mn
        da1 = d1 + 2.0 * mn
        df1 = d1 - 2.4 * mn
        d2 = (z2 + 2.0 * x2) * mn
        da2 = d2 + 2.0 * mn * (1.0 + x2)
        df2 = d2 - 2.0 * mn * (1.2 - x2)
        a_calc = 0.5 * (d1 + d2)
        b2 = self._safe_float("b2_mm", 18)

        a_target_txt = self.inputs.get("a_target_mm", tk.StringVar(value="")).get().strip()
        a_target = float(a_target_txt) if a_target_txt else None
        a_display = a_target if a_target else a_calc

        # Update check text
        if a_target is not None:
            delta = a_target - a_calc
            xsum_need = a_target / mn - 0.5 * (q + z2) if mn > 0 else 0
            self.geom_check_var.set(
                f"几何校核：a_calc = {a_calc:.3f} mm,  a_target = {a_target:.3f} mm,  "
                f"Δa = {delta:+.3f} mm,  建议 x₁+x₂ ≈ {xsum_need:.4f}")
        else:
            self.geom_check_var.set(f"几何校核：a_calc = {a_calc:.3f} mm，未设定目标中心距。")

        # --- Drawing scale: normalize so wheel d2 maps to a fixed visual size ---
        scale = 60.0 / max(d2, 1)  # wheel pitch circle = 60 visual units

        r1  = 0.5 * d1  * scale
        ra1 = 0.5 * da1 * scale
        rf1 = 0.5 * df1 * scale
        r2  = 0.5 * d2  * scale
        ra2 = 0.5 * da2 * scale
        rf2 = 0.5 * df2 * scale
        a_s = a_display * scale

        # Centers
        cx_wheel, cy_wheel = 0, 0
        cx_worm, cy_worm = 0, a_s  # worm above wheel

        # ---- Draw wheel (blue) cross-section ----
        # Tip circle
        ax.add_patch(mpatches.Circle((cx_wheel, cy_wheel), ra2, fill=False,
                     edgecolor=CLR_WHEEL, linewidth=1.5, linestyle="-", alpha=0.6))
        # Pitch circle
        ax.add_patch(mpatches.Circle((cx_wheel, cy_wheel), r2, fill=False,
                     edgecolor=CLR_WHEEL, linewidth=2.0, linestyle="-"))
        # Root circle
        ax.add_patch(mpatches.Circle((cx_wheel, cy_wheel), rf2, fill=False,
                     edgecolor=CLR_WHEEL, linewidth=1.5, linestyle="--", alpha=0.6))
        # Center mark
        ax.plot(cx_wheel, cy_wheel, "+", color=CLR_WHEEL, markersize=10, markeredgewidth=1.5)

        # Draw some teeth on the wheel (top portion where mesh occurs)
        n_teeth_draw = min(z2, 30)  # draw up to 30 teeth
        tooth_angles = np.linspace(0, 2 * np.pi, z2, endpoint=False)
        for i, ang in enumerate(tooth_angles):
            if i >= n_teeth_draw and i < z2 - 2:
                continue
            # Simplified tooth profile as trapezoid
            hw = np.pi * mn * scale / (2 * d2 * scale) * r2  # half tooth width at pitch
            ang_hw = hw / r2 if r2 > 0 else 0.05
            # Tooth tip arc
            a1 = ang - ang_hw * 0.7
            a2 = ang + ang_hw * 0.7
            # Draw tooth outline
            pts = []
            pts.append((cx_wheel + rf2 * np.cos(ang - ang_hw * 1.0),
                       cy_wheel + rf2 * np.sin(ang - ang_hw * 1.0)))
            pts.append((cx_wheel + ra2 * np.cos(ang - ang_hw * 0.5),
                       cy_wheel + ra2 * np.sin(ang - ang_hw * 0.5)))
            pts.append((cx_wheel + ra2 * np.cos(ang + ang_hw * 0.5),
                       cy_wheel + ra2 * np.sin(ang + ang_hw * 0.5)))
            pts.append((cx_wheel + rf2 * np.cos(ang + ang_hw * 1.0),
                       cy_wheel + rf2 * np.sin(ang + ang_hw * 1.0)))
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            ax.plot(xs, ys, color=CLR_WHEEL, linewidth=0.8, alpha=0.5)

        # ---- Draw worm (red) axial cross-section (side view as rectangle + teeth) ----
        worm_len = b2 * scale * 1.8  # visual worm length
        half_len = worm_len / 2

        # Worm body (between root circles)
        rect_x = cx_worm - half_len
        rect_y = cy_worm - rf1
        ax.add_patch(mpatches.FancyBboxPatch(
            (rect_x, rect_y), worm_len, 2 * rf1,
            boxstyle="round,pad=0.5", facecolor=CLR_WORM, alpha=0.08,
            edgecolor=CLR_WORM, linewidth=1.0))

        # Worm tip outline
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm + ra1, cy_worm + ra1], color=CLR_WORM, linewidth=1.0, linestyle="--", alpha=0.5)
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm - ra1, cy_worm - ra1], color=CLR_WORM, linewidth=1.0, linestyle="--", alpha=0.5)

        # Worm pitch lines
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm + r1, cy_worm + r1], color=CLR_WORM, linewidth=1.8)
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm - r1, cy_worm - r1], color=CLR_WORM, linewidth=1.8)

        # Worm root lines
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm + rf1, cy_worm + rf1], color=CLR_WORM, linewidth=1.0, linestyle="--", alpha=0.5)
        ax.plot([cx_worm - half_len, cx_worm + half_len],
                [cy_worm - rf1, cy_worm - rf1], color=CLR_WORM, linewidth=1.0, linestyle="--", alpha=0.5)

        # Worm axial teeth (sinusoidal-ish threads)
        pitch_axial = np.pi * mn * scale  # axial pitch
        n_threads = int(worm_len / pitch_axial) + 2
        for t in range(n_threads):
            tx = cx_worm - half_len + t * pitch_axial
            if tx < cx_worm - half_len - pitch_axial or tx > cx_worm + half_len + pitch_axial:
                continue
            # Thread profile (top)
            xs_t = [tx, tx + pitch_axial * 0.15, tx + pitch_axial * 0.35, tx + pitch_axial * 0.5]
            ys_top = [cy_worm + rf1, cy_worm + ra1, cy_worm + ra1, cy_worm + rf1]
            ys_bot = [cy_worm - rf1, cy_worm - ra1, cy_worm - ra1, cy_worm - rf1]
            ax.plot(xs_t, ys_top, color=CLR_WORM, linewidth=0.9, alpha=0.6,
                    clip_on=True)
            ax.plot(xs_t, ys_bot, color=CLR_WORM, linewidth=0.9, alpha=0.6,
                    clip_on=True)

        # Worm center axis
        ax.plot([cx_worm - half_len - 5, cx_worm + half_len + 5],
                [cy_worm, cy_worm], color=CLR_WORM, linewidth=0.8, linestyle="-.",
                alpha=0.5)
        ax.plot(cx_worm, cy_worm, "+", color=CLR_WORM, markersize=10, markeredgewidth=1.5)

        # ===== Dimension annotations =====
        dim_color = CLR_DIM
        dim_fontsize = 8

        # Center distance (vertical, between centers)
        ax.annotate("", xy=(cx_worm + half_len + 12, cy_worm),
                    xytext=(cx_worm + half_len + 12, cy_wheel),
                    arrowprops=dict(arrowstyle="<->", color=dim_color, lw=1.3))
        ax.text(cx_worm + half_len + 14, (cy_worm + cy_wheel) / 2,
                f"a = {a_display:.2f}",
                fontsize=dim_fontsize, color=dim_color, va="center", rotation=90)

        # Worm d1 (vertical bracket)
        bx = cx_worm - half_len - 8
        ax.annotate("", xy=(bx, cy_worm + r1), xytext=(bx, cy_worm - r1),
                    arrowprops=dict(arrowstyle="<->", color=CLR_WORM, lw=1.2))
        ax.text(bx - 2, cy_worm, f"d₁={d1:.2f}", fontsize=dim_fontsize, color=CLR_WORM,
                va="center", ha="right")

        # Worm da1
        bx2 = cx_worm - half_len - 16
        ax.annotate("", xy=(bx2, cy_worm + ra1), xytext=(bx2, cy_worm - ra1),
                    arrowprops=dict(arrowstyle="<->", color=CLR_WORM, lw=1.0, alpha=0.6))
        ax.text(bx2 - 2, cy_worm, f"dₐ₁={da1:.2f}", fontsize=7, color=CLR_WORM,
                va="center", ha="right", alpha=0.7)

        # Wheel d2 (horizontal bracket)
        by = cy_wheel - ra2 - 8
        ax.annotate("", xy=(cx_wheel - r2, by), xytext=(cx_wheel + r2, by),
                    arrowprops=dict(arrowstyle="<->", color=CLR_WHEEL, lw=1.2))
        ax.text(cx_wheel, by - 3, f"d₂={d2:.2f}", fontsize=dim_fontsize, color=CLR_WHEEL,
                va="top", ha="center")

        # Wheel da2
        by2 = cy_wheel - ra2 - 16
        ax.annotate("", xy=(cx_wheel - ra2, by2), xytext=(cx_wheel + ra2, by2),
                    arrowprops=dict(arrowstyle="<->", color=CLR_WHEEL, lw=1.0, alpha=0.6))
        ax.text(cx_wheel, by2 - 3, f"dₐ₂={da2:.2f}", fontsize=7, color=CLR_WHEEL,
                va="top", ha="center", alpha=0.7)

        # Tooth width b2
        ax.annotate("", xy=(cx_worm - half_len, cy_worm + ra1 + 6),
                    xytext=(cx_worm + half_len, cy_worm + ra1 + 6),
                    arrowprops=dict(arrowstyle="<->", color=dim_color, lw=1.0))
        ax.text(cx_worm, cy_worm + ra1 + 9, f"b₂={b2:.1f}", fontsize=dim_fontsize,
                color=dim_color, ha="center", va="bottom")

        # Labels
        ax.text(cx_worm, cy_worm + ra1 + 18, "蜗杆 (Worm)", fontsize=11, color=CLR_WORM,
                ha="center", va="bottom", fontweight="bold")
        ax.text(cx_wheel, cy_wheel - ra2 - 26, "蜗轮 (Wheel)", fontsize=11, color=CLR_WHEEL,
                ha="center", va="top", fontweight="bold")

        # Module & basic info legend
        info_text = (
            f"mₙ = {mn} mm    z₁ = {z1}    z₂ = {z2}\n"
            f"x₁ = {x1:.3f}    x₂ = {x2:.3f}\n"
            f"i = {ratio:.1f}    αₙ = {self._safe_float('alpha_n_deg', 20):.1f}°"
        )
        ax.text(0.02, 0.02, info_text, transform=ax.transAxes, fontsize=9,
                color=CLR_TEXT2, va="bottom", ha="left",
                bbox=dict(boxstyle="round,pad=0.4", facecolor=CLR_SECTION, edgecolor=CLR_BORDER,
                          alpha=0.9))

        # Legend in top-right
        legend_items = [
            mpatches.Patch(facecolor=CLR_WORM, alpha=0.3, edgecolor=CLR_WORM, label="蜗杆"),
            mpatches.Patch(facecolor=CLR_WHEEL, alpha=0.3, edgecolor=CLR_WHEEL, label="蜗轮"),
        ]
        ax.legend(handles=legend_items, loc="upper right", fontsize=9, framealpha=0.9,
                  edgecolor=CLR_BORDER)

        # Title
        ax.set_title("蜗轮蜗杆啮合截面示意图", fontsize=13, color=CLR_TEXT, pad=10, fontweight="bold")

        # Tight margins
        ax.autoscale_view()
        margin = max(ra2, a_s + ra1) * 0.3
        ax.set_xlim(-(ra2 + margin + 20), ra2 + margin + 20)
        ax.set_ylim(-(ra2 + margin + 30), a_s + ra1 + margin + 22)

        fig.tight_layout()
        self.canvas_geom.draw()

    # ==================================================================
    # Tab 2: Materials & S-N
    # ==================================================================
    def _build_mat_tab(self):
        top = tk.Frame(self.tab_mat, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=10, pady=10)

        # Worm material card
        g1 = tk.Frame(top, bg=CLR_CARD, highlightbackground=CLR_BORDER, highlightthickness=1)
        g1.pack(fill="x", pady=(0, 8))
        self._section_label(g1, "蜗杆材料（默认 37CrS4）")
        row = tk.Frame(g1, bg=CLR_CARD)
        row.pack(fill="x", padx=12, pady=8)
        tk.Label(row, text="选择材料", bg=CLR_CARD, fg=CLR_TEXT).pack(side="left")
        self.steel_var = tk.StringVar(
            value="37CrS4.json" if "37CrS4.json" in self.steel_db else list(self.steel_db.keys())[0])
        self.steel_cb = ttk.Combobox(row, textvariable=self.steel_var,
                                     values=list(self.steel_db.keys()), state="readonly", width=35)
        self.steel_cb.pack(side="left", padx=8)
        tk.Button(row, text="加载", command=self.load_steel, font=("", 10),
                  fg=CLR_ACCENT, bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left")
        tk.Button(row, text="导入 JSON...", command=self.import_steel, font=("", 10),
                  fg=CLR_TEXT2, bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left", padx=8)
        self.steel_text = tk.Text(g1, height=4, wrap="word", bg=CLR_SECTION, fg=CLR_TEXT,
                                  relief="flat", bd=0, font=("", 10), padx=8, pady=4)
        self.steel_text.pack(fill="x", padx=12, pady=(0, 10))

        # Wheel material card
        g2 = tk.Frame(top, bg=CLR_CARD, highlightbackground=CLR_BORDER, highlightthickness=1)
        g2.pack(fill="both", expand=True, pady=(0, 4))
        self._section_label(g2, "蜗轮材料（支持温度相关模量 & S-N 曲线）")

        row2 = tk.Frame(g2, bg=CLR_CARD)
        row2.pack(fill="x", padx=12, pady=8)
        tk.Label(row2, text="基础模板", bg=CLR_CARD, fg=CLR_TEXT).pack(side="left")
        self.wheel_var = tk.StringVar(
            value="PA66_modified_draft.json" if "PA66_modified_draft.json" in self.poly_db
            else list(self.poly_db.keys())[0])
        self.wheel_cb = ttk.Combobox(row2, textvariable=self.wheel_var,
                                     values=list(self.poly_db.keys()), state="readonly", width=35)
        self.wheel_cb.pack(side="left", padx=8)
        tk.Button(row2, text="加载", command=self.load_wheel, font=("", 10),
                  fg=CLR_ACCENT, bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left")
        tk.Button(row2, text="导入 JSON...", command=self.import_wheel, font=("", 10),
                  fg=CLR_TEXT2, bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left", padx=8)

        # E(T) row
        row3 = tk.Frame(g2, bg=CLR_CARD)
        row3.pack(fill="x", padx=12, pady=4)
        tk.Label(row3, text="E(T) 数据点 (°C:GPa)", bg=CLR_CARD, fg=CLR_TEXT,
                 font=("", 10)).pack(side="left")
        self.Et_var = tk.StringVar(value=self._format_Et())
        tk.Entry(row3, textvariable=self.Et_var, width=60, font=("", 10),
                 relief="solid", bd=1, bg=CLR_INPUT_BG, fg=CLR_TEXT,
                 highlightthickness=1, highlightcolor=CLR_ACCENT,
                 highlightbackground=CLR_BORDER).pack(side="left", padx=8)
        tk.Label(row3, text="例: 23:3.0,60:2.4,80:2.0", bg=CLR_CARD,
                 fg=CLR_DIM, font=("", 9)).pack(side="left")

        # S-N table
        sn_frame = tk.Frame(g2, bg=CLR_CARD)
        sn_frame.pack(fill="both", padx=12, pady=6, expand=True)
        tk.Label(sn_frame, text="S-N 曲线数据", bg=CLR_CARD, fg=CLR_TEXT,
                 font=("", 10, "bold"), anchor="w").pack(anchor="w")

        cols = ("temp_C", "N", "contact_MPa", "root_MPa")
        self.sn_table = ttk.Treeview(sn_frame, columns=cols, show="headings", height=6)
        self.sn_table.heading("temp_C", text="温度 °C")
        self.sn_table.heading("N", text="循环次数 N")
        self.sn_table.heading("contact_MPa", text="接触许用 MPa")
        self.sn_table.heading("root_MPa", text="齿根许用 MPa")
        for col in cols:
            self.sn_table.column(col, width=150, anchor="center")
        self.sn_table.pack(side="left", fill="both", expand=True)
        table_scroll = ttk.Scrollbar(sn_frame, orient="vertical", command=self.sn_table.yview)
        self.sn_table.configure(yscrollcommand=table_scroll.set)
        table_scroll.pack(side="left", fill="y")

        # S-N input row
        ctrl = tk.Frame(g2, bg=CLR_CARD)
        ctrl.pack(fill="x", padx=12, pady=6)
        self.sn_temp_var = tk.StringVar(value=self._defaults.get("temp_C", "80"))
        self.sn_N_var = tk.StringVar(value="1e6")
        self.sn_contact_var = tk.StringVar(value="110")
        self.sn_root_var = tk.StringVar(value="45")
        for var, w, hint in [
            (self.sn_temp_var, 8, "°C"), (self.sn_N_var, 12, "N"),
            (self.sn_contact_var, 12, "接触MPa"), (self.sn_root_var, 12, "齿根MPa")
        ]:
            tk.Entry(ctrl, textvariable=var, width=w, font=("", 10), relief="solid", bd=1,
                     bg=CLR_INPUT_BG, fg=CLR_TEXT, highlightthickness=1,
                     highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER).pack(side="left", padx=(0, 4))
        tk.Button(ctrl, text="添加行", command=self.add_sn_row, font=("", 10),
                  fg=CLR_ACCENT, bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left", padx=4)
        tk.Button(ctrl, text="删除选中", command=self.delete_sn_rows, font=("", 10),
                  fg="#FF3B30", bg=CLR_CARD, relief="flat", cursor="hand2").pack(side="left")
        tk.Button(ctrl, text="应用材料卡", command=self.apply_wheel_card, font=("", 10, "bold"),
                  fg="#FFF", bg=CLR_ACCENT2, activebackground="#28A745",
                  activeforeground="#FFF", relief="flat", cursor="hand2",
                  padx=12).pack(side="left", padx=12)

        self.wheel_text = tk.Text(g2, height=6, wrap="word", bg=CLR_SECTION, fg=CLR_TEXT,
                                  relief="flat", bd=0, font=("", 10), padx=8, pady=4)
        self.wheel_text.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self._load_sn_table_from_wheel()
        self._refresh_material_texts()

    # ==================================================================
    # Tab 3: Results
    # ==================================================================
    def _build_res_tab(self):
        top = tk.Frame(self.tab_res, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=8, pady=8)
        fig = Figure(figsize=(10, 7), dpi=100, facecolor=CLR_CARD)
        self.ax_p = fig.add_subplot(231)
        self.ax_s = fig.add_subplot(232)
        self.ax_t = fig.add_subplot(233)
        self.ax_eta = fig.add_subplot(234)
        self.ax_cloud = fig.add_subplot(235, projection="3d")
        self.ax_legend = fig.add_subplot(236)
        self.ax_legend.axis("off")
        self.canvas_res = FigureCanvasTkAgg(fig, master=top)
        self.canvas_res.get_tk_widget().pack(fill="both", expand=True)
        tk.Label(top, text="轻量代理模型结果（含 KISSsoft 风格修正系数）",
                 bg=CLR_BG, fg=CLR_TEXT2, font=("", 10)).pack(anchor="w", padx=6, pady=4)

    # ==================================================================
    # Tab 4: Fatigue summary
    # ==================================================================
    def _build_fat_tab(self):
        top = tk.Frame(self.tab_fat, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=10, pady=10)

        cards = tk.Frame(top, bg=CLR_BG)
        cards.pack(fill="x")

        # Worm output card
        worm_box = tk.Frame(cards, bg=CLR_CARD, highlightbackground=CLR_WORM, highlightthickness=2)
        worm_box.pack(side="left", fill="both", expand=True, padx=(0, 6))
        tk.Label(worm_box, text="蜗杆输出参数", bg=CLR_CARD, fg=CLR_WORM,
                 font=("", 11, "bold")).pack(padx=8, pady=(8, 4), anchor="w")
        self.worm_out = ttk.Treeview(worm_box, columns=("k", "v"), show="headings", height=7)
        self.worm_out.heading("k", text="参数")
        self.worm_out.heading("v", text="值")
        self.worm_out.column("k", width=180, anchor="w")
        self.worm_out.column("v", width=220, anchor="w")
        self.worm_out.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # Wheel output card
        wheel_box = tk.Frame(cards, bg=CLR_CARD, highlightbackground=CLR_WHEEL, highlightthickness=2)
        wheel_box.pack(side="left", fill="both", expand=True, padx=(6, 0))
        tk.Label(wheel_box, text="蜗轮输出参数", bg=CLR_CARD, fg=CLR_WHEEL,
                 font=("", 11, "bold")).pack(padx=8, pady=(8, 4), anchor="w")
        self.wheel_out = ttk.Treeview(wheel_box, columns=("k", "v"), show="headings", height=7)
        self.wheel_out.heading("k", text="参数")
        self.wheel_out.heading("v", text="值")
        self.wheel_out.column("k", width=180, anchor="w")
        self.wheel_out.column("v", width=220, anchor="w")
        self.wheel_out.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # Fatigue text
        self.fat_text = tk.Text(top, wrap="word", bd=0, relief="flat", bg=CLR_CARD,
                                fg=CLR_TEXT, font=("", 10), padx=12, pady=8)
        self.fat_text.pack(fill="both", expand=True, padx=2, pady=(10, 0))
        self.fat_text.insert("1.0", '点击"计算并绘图"后，此处显示雨流计数的疲劳损伤与安全系数汇总。')

    # ==================================================================
    # Tree helper
    # ==================================================================
    def _fill_tree_kv(self, tree, rows):
        for i in tree.get_children():
            tree.delete(i)
        for k, v in rows:
            tree.insert("", "end", values=(k, v))

    # ==================================================================
    # Material helpers
    # ==================================================================
    def _format_Et(self):
        pts = self.wheel.get("elastic_T", {}).get("points_C_GPa", [])
        return ",".join([f"{int(p[0])}:{p[1]}" for p in pts])

    def _build_sn_rows_from_legacy(self):
        temp = self._safe_float("temp_C", 80)
        contact = sorted(self.wheel.get("SN", {}).get("contact_allow_MPa_vs_N", []), key=lambda p: p[0])
        root = sorted(self.wheel.get("SN", {}).get("root_allow_MPa_vs_N", []), key=lambda p: p[0])
        all_n = sorted({float(p[0]) for p in contact} | {float(p[0]) for p in root})
        rows = []
        for n in all_n:
            c = next((float(v[1]) for v in contact if float(v[0]) == n), 0.0)
            r = next((float(v[1]) for v in root if float(v[0]) == n), 0.0)
            rows.append({"temp_C": temp, "N": n, "contact_MPa": c, "root_MPa": r})
        return rows

    def _load_sn_table_from_wheel(self):
        for iid in self.sn_table.get_children():
            self.sn_table.delete(iid)
        rows = self.wheel.get("SN", {}).get("table")
        if not rows:
            rows = self._build_sn_rows_from_legacy()
        self.sn_rows = []
        for row in rows:
            one = {
                "temp_C": float(row.get("temp_C", self._defaults.get("temp_C", 80))),
                "N": float(row.get("N", 1e6)),
                "contact_MPa": float(row.get("contact_MPa", 0)),
                "root_MPa": float(row.get("root_MPa", 0))
            }
            self.sn_rows.append(one)
            self.sn_table.insert("", "end",
                values=(f"{one['temp_C']:.1f}", f"{one['N']:.3e}",
                        f"{one['contact_MPa']:.3f}", f"{one['root_MPa']:.3f}"))

    def add_sn_row(self):
        try:
            row = {
                "temp_C": float(self.sn_temp_var.get()),
                "N": float(self.sn_N_var.get()),
                "contact_MPa": float(self.sn_contact_var.get()),
                "root_MPa": float(self.sn_root_var.get())
            }
        except ValueError:
            messagebox.showerror("输入错误", "S-N 行输入需为数值。")
            return
        self.sn_rows.append(row)
        self.sn_table.insert("", "end",
            values=(f"{row['temp_C']:.1f}", f"{row['N']:.3e}",
                    f"{row['contact_MPa']:.3f}", f"{row['root_MPa']:.3f}"))

    def delete_sn_rows(self):
        sel = self.sn_table.selection()
        if not sel:
            return
        idxs = sorted([self.sn_table.index(i) for i in sel], reverse=True)
        for iid in sel:
            self.sn_table.delete(iid)
        for idx in idxs:
            if 0 <= idx < len(self.sn_rows):
                self.sn_rows.pop(idx)

    def load_steel(self):
        self.steel = load_json(self.steel_db[self.steel_var.get()])
        self._refresh_material_texts()

    def import_steel(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        dst = os.path.join(self.base_dir, "materials", "metals", os.path.basename(path))
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        with open(dst, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.steel_db[os.path.basename(dst)] = dst
        self.steel_cb["values"] = list(self.steel_db.keys())
        self.steel_var.set(os.path.basename(dst))
        self.load_steel()

    def load_wheel(self):
        self.wheel = load_json(self.poly_db[self.wheel_var.get()])
        self.Et_var.set(self._format_Et())
        self._load_sn_table_from_wheel()
        self._refresh_material_texts()

    def import_wheel(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        dst = os.path.join(self.base_dir, "materials", "polymers", os.path.basename(path))
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        with open(dst, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.poly_db[os.path.basename(dst)] = dst
        self.wheel_cb["values"] = list(self.poly_db.keys())
        self.wheel_var.set(os.path.basename(dst))
        self.load_wheel()

    def apply_wheel_card(self):
        pts_et = []
        for part in self.Et_var.get().split(","):
            part = part.strip()
            if not part or ":" not in part:
                continue
            a, b = part.split(":", 1)
            pts_et.append([float(a.strip()), float(b.strip())])
        self.wheel["elastic_T"] = {"points_C_GPa": pts_et}
        table = []
        for row in self.sn_rows:
            table.append({
                "temp_C": float(row["temp_C"]),
                "N": float(row["N"]),
                "contact_MPa": float(row["contact_MPa"]),
                "root_MPa": float(row["root_MPa"])
            })
        table = sorted(table, key=lambda x: (x["temp_C"], x["N"]))
        self.wheel["SN"] = self.wheel.get("SN", {})
        self.wheel["SN"]["table"] = table
        t_ref = self._safe_float("temp_C", 80)
        near = sorted(table, key=lambda x: abs(x["temp_C"] - t_ref))
        if near:
            t_pick = near[0]["temp_C"]
            same_t = [r for r in table if abs(r["temp_C"] - t_pick) < 1e-9]
            self.wheel["SN"]["contact_allow_MPa_vs_N"] = [
                [r["N"], r["contact_MPa"]] for r in same_t if r["contact_MPa"] > 0]
            self.wheel["SN"]["root_allow_MPa_vs_N"] = [
                [r["N"], r["root_MPa"]] for r in same_t if r["root_MPa"] > 0]
        self._refresh_material_texts()
        messagebox.showinfo("完成", "蜗轮材料卡已应用（仅当前会话）。")

    def _refresh_material_texts(self):
        self.steel_text.delete("1.0", "end")
        self.steel_text.insert("1.0", json.dumps(self.steel, ensure_ascii=False, indent=2))
        self.wheel_text.delete("1.0", "end")
        self.wheel_text.insert("1.0", json.dumps(self.wheel, ensure_ascii=False, indent=2))

    # ==================================================================
    # Compute
    # ==================================================================
    def _collect_inputs(self):
        """Collect all input values into a flat dict for the compute engine."""
        d = {}
        for k, v in self.inputs.items():
            d[k] = v.get()
        # Map b2_mm -> b_mm for backward compat with compute engine
        d["b_mm"] = d.get("b2_mm", "18")
        return d

    def run(self):
        try:
            self._auto_calc_worm()
            self._auto_calc_wheel()
            inp = self._collect_inputs()
            res = compute_worm_cycle(inp, self.steel, self.wheel)
            self.res = res
            self.plot_results(res)
            self.update_fatigue(res)
            self.nb.select(self.tab_res)
        except Exception as e:
            messagebox.showerror("计算失败", str(e))
            import traceback
            traceback.print_exc()

    def plot_results(self, res):
        phi = res["phi"]
        deg = phi * 180 / np.pi
        for ax in [self.ax_p, self.ax_s, self.ax_t, self.ax_eta, self.ax_cloud, self.ax_legend]:
            ax.clear()

        # Style helper
        def _style_ax(ax, title, xlabel, ylabel):
            ax.set_title(title, fontsize=10, fontweight="bold", color=CLR_TEXT)
            ax.set_xlabel(xlabel, fontsize=9, color=CLR_TEXT2)
            ax.set_ylabel(ylabel, fontsize=9, color=CLR_TEXT2)
            ax.tick_params(labelsize=8)
            ax.grid(True, alpha=0.2)
            ax.set_facecolor("#FAFBFC")

        self.ax_p.fill_between(deg, res["p_contact_MPa"], alpha=0.15, color=CLR_ACCENT)
        self.ax_p.plot(deg, res["p_contact_MPa"], color=CLR_ACCENT, linewidth=1.5)
        _style_ax(self.ax_p, "接触应力 p(φ)", "φ (deg)", "MPa")

        self.ax_s.fill_between(deg, res["sigma_root_MPa"], alpha=0.15, color=CLR_WORM)
        self.ax_s.plot(deg, res["sigma_root_MPa"], color=CLR_WORM, linewidth=1.5)
        _style_ax(self.ax_s, "齿根应力 σF(φ)", "φ (deg)", "MPa")

        self.ax_t.fill_between(deg, res["T2_Nm"], alpha=0.15, color=CLR_ACCENT2)
        self.ax_t.plot(deg, res["T2_Nm"], color=CLR_ACCENT2, linewidth=1.5)
        _style_ax(self.ax_t, "输出扭矩 T₂(φ)", "φ (deg)", "N·m")

        self.ax_eta.plot(deg, res["eta"], color=CLR_ACCENT, linewidth=1.5, label="η")
        self.ax_eta.plot(deg, res["Nc_proxy"], color="#FF9500", linewidth=1.5, label="Nc")
        _style_ax(self.ax_eta, "效率 η & 接触数 Nc", "φ (deg)", "-")
        self.ax_eta.legend(fontsize=8, framealpha=0.9)

        # 3D stress cloud
        width_pts = np.linspace(-1.0, 1.0, 36)
        PHI, BW = np.meshgrid(phi, width_pts)
        base = np.interp(PHI[0], phi, res["sigma_root_MPa"])
        base2d = np.tile(base, (len(width_pts), 1))
        spread = 1.0 + 0.22 * (BW ** 2) + 0.10 * np.sin(2.0 * PHI)
        sigma_3d = base2d * spread
        norm = Normalize(vmin=float(np.min(sigma_3d)), vmax=float(np.max(sigma_3d)))
        face_colors = matplotlib.cm.inferno(norm(sigma_3d))
        self.ax_cloud.plot_surface(
            PHI * 180.0 / np.pi, BW, sigma_3d,
            rstride=1, cstride=1, linewidth=0, antialiased=True,
            facecolors=face_colors, shade=False)
        self.ax_cloud.set_title("3D齿根应力云图", fontsize=10, fontweight="bold", color=CLR_TEXT)
        self.ax_cloud.set_xlabel("φ (deg)", fontsize=8)
        self.ax_cloud.set_ylabel("齿宽", fontsize=8)
        self.ax_cloud.set_zlabel("σF MPa", fontsize=8)
        self.ax_cloud.view_init(elev=24, azim=-130)

        sm = matplotlib.cm.ScalarMappable(cmap="inferno", norm=norm)
        sm.set_array([])
        if self._cloud_cbar is not None:
            try:
                self._cloud_cbar.remove()
            except Exception:
                pass
        self._cloud_cbar = self.canvas_res.figure.colorbar(sm, ax=self.ax_cloud, shrink=0.60, pad=0.08)
        self._cloud_cbar.set_label("应力 MPa", fontsize=8)

        # Legend / info panel
        self.ax_legend.axis("off")
        self.ax_legend.set_facecolor(CLR_CARD)
        m = res["meta"]
        info = (
            f"计算摘要\n"
            f"─────────────────\n"
            f"z₁ = {m['z1']}   z₂ = {m['z2']}\n"
            f"d₁ = {m['d1_mm']:.2f} mm\n"
            f"d₂ = {m['d2_mm']:.2f} mm\n"
            f"a  = {m['a_mm']:.2f} mm\n"
            f"η₀ = {m['eta0']:.3f}\n"
            f"γ  = {m['gamma_deg']:.2f}°\n"
            f"─────────────────\n"
            f"接触应力峰值: {float(np.max(res['p_contact_MPa'])):.1f} MPa\n"
            f"齿根应力峰值: {float(np.max(res['sigma_root_MPa'])):.1f} MPa"
        )
        self.ax_legend.text(0.05, 0.95, info, transform=self.ax_legend.transAxes,
                           fontsize=9, va="top", ha="left", color=CLR_TEXT,
                           family="monospace",
                           bbox=dict(boxstyle="round,pad=0.5", facecolor=CLR_SECTION,
                                     edgecolor=CLR_BORDER, alpha=0.9))

        self.canvas_res.figure.tight_layout()
        self.canvas_res.draw()

    def update_fatigue(self, res):
        m = res["meta"]
        z1 = m["z1"]
        worm_rows = [
            ("头数 z₁", f"{z1:d}"),
            ("分度圆 d₁", f"{m['d1_mm']:.2f} mm"),
            ("齿顶圆 dₐ₁", f"{m['da1_mm']:.2f} mm"),
            ("齿根圆 df₁", f"{m['df1_mm']:.2f} mm"),
            ("变位系数 x₁", f"{m['x1']:.3f}"),
            ("导程角 γ", f"{m['gamma_deg']:.2f}°"),
            ("效率 η₀", f"{m['eta0']:.3f}"),
            ("齿根应力峰值", f"{np.max(res['sigma_root_MPa']):.2f} MPa"),
            ("齿根安全系数", "-" if m.get("SF_root") is None else f"{m['SF_root']:.2f}")
        ]
        wheel_rows = [
            ("齿数 z₂", f"{m['z2']:d}"),
            ("分度圆 d₂", f"{m['d2_mm']:.2f} mm"),
            ("齿顶圆 dₐ₂", f"{m['da2_mm']:.2f} mm"),
            ("齿根圆 df₂", f"{m['df2_mm']:.2f} mm"),
            ("变位系数 x₂", f"{m['x2']:.3f}"),
            ("中心距 a", f"{m['a_mm']:.2f} mm"),
            ("接触应力峰值", f"{np.max(res['p_contact_MPa']):.2f} MPa"),
            ("齿面安全系数", "-" if m.get("SF_contact") is None else f"{m['SF_contact']:.2f}"),
            ("中心距偏差 Δa", "-" if m.get("delta_a_mm") is None else f"{m['delta_a_mm']:.2f} mm")
        ]
        self._fill_tree_kv(self.worm_out, worm_rows)
        self._fill_tree_kv(self.wheel_out, wheel_rows)

        lines = []
        lines.append("【雨流计数 + Miner 累积损伤（基于齿根应力代理）】")
        lines.append(f"  累积损伤 D ≈ {m['damage_root']:.3e}（D < 1 视为满足目标寿命）")
        if m.get("SF_root") is not None:
            lines.append(f"  齿根安全系数 SF_root ≈ {m['SF_root']:.2f}")
        else:
            lines.append("  齿根安全系数：未提供蜗轮 root SN 数据")
        if m.get("SF_contact") is not None:
            lines.append(f"  齿面安全系数 SF_contact ≈ {m['SF_contact']:.2f}")
        else:
            lines.append("  齿面安全系数：未提供蜗轮 contact SN 数据")
        lines.append("")
        lines.append("【关键几何与材料参数】")
        lines.append(f"  z₁ = {m['z1']},  z₂ = {m['z2']}")
        lines.append(f"  d₁ = {m['d1_mm']:.2f} mm,  d₂ = {m['d2_mm']:.2f} mm,  a = {m['a_mm']:.2f} mm")
        if m.get("a_target_mm") is not None:
            lines.append(f"  目标中心距 a_target = {m['a_target_mm']:.2f} mm,  偏差 Δa = {m['delta_a_mm']:.2f} mm")
        lines.append(f"  x₁ = {m['x1']:.3f},  x₂ = {m['x2']:.3f}")
        lines.append(f"  γ = {m['gamma_deg']:.2f}°")
        lines.append(f"  E' = {m['Eprime_GPa']:.2f} GPa")
        lines.append(f"  η₀ = {m['eta0']:.3f}")
        lines.append(f"  K系数: KA={m['KA']:.3f}  KV={m['KV']:.3f}  KHβ={m['KHb']:.3f}  KFβ={m['KFb']:.3f}")
        self.fat_text.delete("1.0", "end")
        self.fat_text.insert("1.0", "\n".join(lines))

    # ==================================================================
    # Export
    # ==================================================================
    def export_xlsx(self):
        if self.res is None:
            messagebox.showwarning("提示", "请先计算一次再导出。")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        inputs = self._collect_inputs()
        export_cycle_xlsx(path, inputs, self.steel, self.wheel, self.res)
        messagebox.showinfo("完成", f"已导出：\n{path}")


if __name__ == "__main__":
    App().mainloop()
