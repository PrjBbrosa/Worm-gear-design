
import os, json, math
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.patches as mpatches
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import font_manager
from matplotlib.colors import Normalize
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

from src.utils import load_json
from src.worm_model import compute_worm_cycle
from src.export_xlsx import export_cycle_xlsx

# =====================================================================
# i18n bilingual dictionary
# =====================================================================
LANG_ZH = {
    "app_title": "WormGear Studio \u2014 \u8717\u8f6e\u8717\u6746\u8bbe\u8ba1\u6821\u6838",
    "tab_geom": "  \u51e0\u4f55\u53c2\u6570  ",
    "tab_mat": "  \u6750\u6599\u4e0eS-N  ",
    "tab_res": "  \u5e94\u529b\u4e0e\u6548\u7387  ",
    "tab_fat": "  \u5bff\u547d\u6821\u6838  ",
    "tab_formula": "  \u516c\u5f0f\u8bf4\u660e  ",
    "menu_file": "  \u6587\u4ef6  ",
    "menu_export": "  \u5bfc\u51fa XLSX\uff08\u66f2\u7ebf\uff09...",
    "menu_exit": "  \u9000\u51fa",
    "drive_params": "\u9a71\u52a8\u53c2\u6570",
    "T1_Nm": "\u8f93\u5165\u626d\u77e9 T1", "n1_rpm": "\u8717\u6746\u8f6c\u901f n1",
    "ratio": "\u4f20\u52a8\u6bd4 i", "life_h": "\u76ee\u6807\u5bff\u547d", "steps": "\u76f8\u4f4d\u70b9\u6570",
    "calc_ratio": "  \u7531z\u7b97i  ", "worm_params": "\u8717\u6746\u53c2\u6570",
    "auto_calc_worm": "  \u81ea\u52a8\u8ba1\u7b97\u8717\u6746\u5c3a\u5bf8  ",
    "basic_params": "\u57fa\u672c\u53c2\u6570", "z1": "\u5934\u6570 z1", "mn_mm": "\u6cd5\u5411\u6a21\u6570 mn",
    "q": "\u76f4\u5f84\u7cfb\u6570 q", "x1": "\u53d8\u4f4d\u7cfb\u6570 x1", "beta_deg": "\u8717\u6746\u87ba\u65cb\u89d2 beta",
    "alpha_n_deg": "\u6cd5\u5411\u538b\u529b\u89d2 an", "rho_f_mm": "\u9f7f\u6839\u5706\u89d2 rf",
    "computed_dims": "\u8ba1\u7b97\u5c3a\u5bf8\uff08\u81ea\u52a8\uff09",
    "d1_mm": "\u5206\u5ea6\u5706\u76f4\u5f84 d1", "da1_mm": "\u9f7f\u9876\u5706\u76f4\u5f84 da1",
    "df1_mm": "\u9f7f\u6839\u5706\u76f4\u5f84 df1",
    "gamma_deg": "\u5bfc\u7a0b\u89d2 gamma", "px_mm": "\u8f74\u5411\u9f7f\u8ddd px",
    "pz_mm": "\u5bfc\u7a0b pz", "L_worm_mm": "\u8717\u6746\u957f\u5ea6 L",
    "wheel_params": "\u8717\u8f6e\u53c2\u6570",
    "auto_calc_wheel": "  \u81ea\u52a8\u8ba1\u7b97\u8717\u8f6e\u5c3a\u5bf8  ",
    "z2": "\u9f7f\u6570 z2", "calc_z2": "  \u7531i\u7b97z2  ",
    "x2": "\u53d8\u4f4d\u7cfb\u6570 x2", "b2_mm": "\u9f7f\u5bbd b2",
    "a_target_mm": "\u76ee\u6807\u4e2d\u5fc3\u8ddd a", "calc_a": "  \u81ea\u52a8\u7b97a  ",
    "d2_mm": "\u5206\u5ea6\u5706\u76f4\u5f84 d2", "da2_mm": "\u9f7f\u9876\u5706\u76f4\u5f84 da2",
    "df2_mm": "\u9f7f\u6839\u5706\u76f4\u5f84 df2",
    "common_params": "\u516c\u5171\u53c2\u6570\u4e0e\u4fee\u6b63\u7cfb\u6570",
    "cross_angle_deg": "\u4ea4\u9519\u89d2 Sigma", "mu": "\u6469\u64e6\u7cfb\u6570 mu",
    "KA": "\u4f7f\u7528\u7cfb\u6570 KA", "KV": "\u52a8\u8f7d\u7cfb\u6570 KV",
    "KHb": "\u9f7f\u5bbd\u8f7d\u8377\u7cfb\u6570 KHb", "KFb": "\u9f7f\u6839\u8f7d\u8377\u7cfb\u6570 KFb",
    "temp_C": "\u5de5\u4f5c\u6e29\u5ea6",
    "btn_refresh": "\u66f4\u65b0\u793a\u610f\u56fe", "btn_calc": "\u8ba1\u7b97\u5e76\u7ed8\u56fe",
    "geom_check_wait": "\u51e0\u4f55\u6821\u6838\u4fe1\u606f\uff1a\u7b49\u5f85\u8f93\u5165\u53c2\u6570\u3002",
    "diagram_title": "Worm-Wheel Focused Mesh Section",
    "worm_label": "Worm", "wheel_label": "Wheel",
    "mat_worm_title": "\u8717\u6746\u6750\u6599",
    "mat_select": "\u9009\u62e9\u6750\u6599", "mat_load": "\u52a0\u8f7d", "mat_import": "\u5bfc\u5165 JSON...",
    "mat_name": "\u6750\u6599\u540d\u79f0", "mat_standard": "\u6807\u51c6", "mat_E": "\u5f39\u6027\u6a21\u91cf E",
    "mat_nu": "\u6cca\u677e\u6bd4 nu", "mat_yield": "\u5c48\u670d\u5f3a\u5ea6 Rp0.2",
    "mat_tensile": "\u62c9\u4f38\u5f3a\u5ea6 Rm", "mat_hardness": "\u786c\u5ea6 HRC",
    "mat_notes": "\u5907\u6ce8",
    "mat_wheel_title": "\u8717\u8f6e\u6750\u6599",
    "mat_base_template": "\u57fa\u7840\u6a21\u677f",
    "mat_Et_label": "E(T) \u6570\u636e\u70b9 (C:GPa)",
    "mat_Et_hint": "\u4f8b: 23:3.0,60:2.4,80:2.0",
    "sn_title": "S-N \u66f2\u7ebf\u6570\u636e",
    "sn_add": "\u6dfb\u52a0\u884c", "sn_delete": "\u5220\u9664\u9009\u4e2d",
    "sn_apply": "\u5e94\u7528\u6750\u6599\u5361",
    "mat_save_json": "\u4fdd\u5b58\u4e3a JSON...",
    "res_footer": "\u8f7b\u91cf\u4ee3\u7406\u6a21\u578b\u7ed3\u679c\uff08\u542b KISSsoft \u98ce\u683c\u4fee\u6b63\u7cfb\u6570\uff09",
    "worm_output": "\u8717\u6746\u8f93\u51fa\u53c2\u6570", "wheel_output": "\u8717\u8f6e\u8f93\u51fa\u53c2\u6570",
    "fat_wait": '\u70b9\u51fb"\u8ba1\u7b97\u5e76\u7ed8\u56fe"\u540e\uff0c\u6b64\u5904\u663e\u793a\u75b2\u52b3\u635f\u4f24\u4e0e\u5b89\u5168\u7cfb\u6570\u6c47\u603b\u3002',
    "formula_title": "\u516c\u5f0f\u4e0e\u53c2\u6570\u8bf4\u660e\uff08\u5ba1\u6838\u7528\uff09",
    "lang_toggle": "EN / \u4e2d",
    "N_m": "N\u00b7m", "rpm": "rpm", "mm": "mm", "deg": "deg", "h": "h",
    "pts": "\u70b9", "C": "\u00b0C",
}

LANG_EN = {
    "app_title": "WormGear Studio \u2014 Worm Gear Design & Check",
    "tab_geom": "  Geometry  ", "tab_mat": "  Material & S-N  ",
    "tab_res": "  Stress & Efficiency  ", "tab_fat": "  Fatigue Check  ", "tab_formula": "  Formula Notes  ",
    "menu_file": "  File  ", "menu_export": "  Export XLSX (curves)...",
    "menu_exit": "  Exit",
    "drive_params": "Drive Parameters",
    "T1_Nm": "Input Torque T1", "n1_rpm": "Worm Speed n1",
    "ratio": "Gear Ratio i", "life_h": "Target Life", "steps": "Phase Points",
    "calc_ratio": "  Calc i from z  ", "worm_params": "Worm Parameters",
    "auto_calc_worm": "  Auto-Calc Worm Dims  ",
    "basic_params": "Basic Parameters", "z1": "No. of Starts z1",
    "mn_mm": "Normal Module mn", "q": "Diameter Factor q", "x1": "Profile Shift x1", "beta_deg": "Worm Helix Angle beta",
    "alpha_n_deg": "Normal Press. Angle an", "rho_f_mm": "Root Fillet rf",
    "computed_dims": "Computed Dimensions",
    "d1_mm": "Pitch Dia. d1", "da1_mm": "Tip Dia. da1", "df1_mm": "Root Dia. df1",
    "gamma_deg": "Lead Angle gamma", "px_mm": "Axial Pitch px",
    "pz_mm": "Lead pz", "L_worm_mm": "Worm Length L",
    "wheel_params": "Wheel Parameters",
    "auto_calc_wheel": "  Auto-Calc Wheel Dims  ",
    "z2": "No. of Teeth z2", "calc_z2": "  Calc z2 from i  ",
    "x2": "Profile Shift x2", "b2_mm": "Face Width b2",
    "a_target_mm": "Target Center Dist. a", "calc_a": "  Auto-Calc a  ",
    "d2_mm": "Pitch Dia. d2", "da2_mm": "Tip Dia. da2", "df2_mm": "Root Dia. df2",
    "common_params": "Common Params & Correction Factors",
    "cross_angle_deg": "Shaft Angle Sigma", "mu": "Friction Coeff. mu",
    "KA": "Application Factor KA", "KV": "Dynamic Factor KV",
    "KHb": "Load Dist. Factor KHb", "KFb": "Root Load Factor KFb",
    "temp_C": "Operating Temp.",
    "btn_refresh": "Refresh Diagram", "btn_calc": "Calculate & Plot",
    "geom_check_wait": "Geometry check: waiting for input.",
    "diagram_title": "Worm-Wheel Focused Mesh Section",
    "worm_label": "Worm", "wheel_label": "Wheel",
    "mat_worm_title": "Worm Material",
    "mat_select": "Select Material", "mat_load": "Load", "mat_import": "Import JSON...",
    "mat_name": "Material Name", "mat_standard": "Standard", "mat_E": "Elastic Modulus E",
    "mat_nu": "Poisson Ratio nu", "mat_yield": "Yield Strength Rp0.2",
    "mat_tensile": "Tensile Strength Rm", "mat_hardness": "Hardness HRC",
    "mat_notes": "Notes",
    "mat_wheel_title": "Wheel Material",
    "mat_base_template": "Base Template",
    "mat_Et_label": "E(T) Data Points (C:GPa)",
    "mat_Et_hint": "e.g. 23:3.0,60:2.4,80:2.0",
    "sn_title": "S-N Curve Data",
    "sn_add": "Add Row", "sn_delete": "Delete Selected",
    "sn_apply": "Apply Material Card",
    "mat_save_json": "Save as JSON...",
    "res_footer": "Lightweight proxy model results (KISSsoft-style correction factors)",
    "worm_output": "Worm Output", "wheel_output": "Wheel Output",
    "fat_wait": 'Click "Calculate & Plot" to see fatigue damage & safety factor summary.',
    "formula_title": "Formulas and Parameter Notes (for review)",
    "lang_toggle": "EN / \u4e2d",
    "N_m": "N\u00b7m", "rpm": "rpm", "mm": "mm", "deg": "deg", "h": "h",
    "pts": "pts", "C": "\u00b0C",
}

# =====================================================================
# Colour palette
# =====================================================================
CLR_BG       = "#F5F5F7"
CLR_CARD     = "#FFFFFF"
CLR_ACCENT   = "#007AFF"
CLR_ACCENT2  = "#34C759"
CLR_TEXT      = "#1D1D1F"
CLR_TEXT2     = "#6E6E73"
CLR_BORDER   = "#D2D2D7"
CLR_INPUT_BG = "#FFFFFF"
CLR_SECTION  = "#F2F2F7"
CLR_WORM     = "#E8453C"
CLR_WHEEL    = "#007AFF"
CLR_DIM      = "#636366"
CLR_BTN_BG   = "#E5E5EA"

# =====================================================================
# Font setup
# =====================================================================
_MPL_FONT = None

def setup_fonts(root=None):
    global _MPL_FONT
    candidates = [
        "PingFang SC", "Hiragino Sans GB", "STHeiti", "Heiti SC",
        "Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
        "WenQuanYi Zen Hei", "Arial Unicode MS",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = None
    for name in candidates:
        if name in available:
            chosen = name
            break
    if chosen is None:
        chosen = "DejaVu Sans"

    _MPL_FONT = chosen
    if root is not None:
        try:
            for fn in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
                tkfont.nametofont(fn).configure(family=chosen, size=11)
        except Exception:
            pass

    matplotlib.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False
    # Force matplotlib to find font for CJK
    matplotlib.rcParams["font.family"] = "sans-serif"
    return chosen


def list_materials(folder):
    out = []
    if not os.path.isdir(folder):
        return out
    for fn in os.listdir(folder):
        if fn.lower().endswith(".json"):
            out.append(os.path.join(folder, fn))
    return sorted(out)


# =====================================================================
# Main Application
# =====================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self._font_family = setup_fonts(self)
        self._lang = LANG_ZH
        self._lang_code = "zh"
        self._setup_style()
        self.title(self._t("app_title"))
        self.geometry("1560x940")
        self.minsize(1200, 750)

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.steel_paths = list_materials(os.path.join(self.base_dir, "materials", "metals"))
        self.poly_paths  = list_materials(os.path.join(self.base_dir, "materials", "polymers"))
        if not self.steel_paths or not self.poly_paths:
            raise RuntimeError("materials directory missing.")

        self.steel_db = {os.path.basename(p): p for p in self.steel_paths}
        self.poly_db  = {os.path.basename(p): p for p in self.poly_paths}
        self.steel = load_json(self.steel_db.get("37CrS4.json", self.steel_paths[0]))
        self.wheel = load_json(self.poly_db.get("PA66_modified_draft.json", self.poly_paths[0]))

        self._defaults = {
            "T1_Nm": "6.0", "n1_rpm": "3000", "ratio": "25",
            "life_h": "3000", "steps": "720",
            "z1": "2", "mn_mm": "2.5", "q": "10", "x1": "0.0", "beta_deg": "11.5",
            "alpha_n_deg": "20", "rho_f_mm": "0.6",
            "da1_mm": "", "df1_mm": "", "d1_mm": "",
            "gamma_deg": "", "px_mm": "", "pz_mm": "", "L_worm_mm": "",
            "z2": "", "x2": "0.0", "d2_mm": "", "da2_mm": "", "df2_mm": "",
            "a_target_mm": "", "b2_mm": "18",
            "cross_angle_deg": "90", "mu": "0.06",
            "KA": "1.10", "KV": "1.05", "KHb": "1.00", "KFb": "1.00",
            "temp_C": "80",
        }
        self.inputs = {}
        self.res = None
        self.sn_rows = []
        self._cloud_cbar = None

        # Track all labelled widgets for language refresh
        self._i18n_widgets = []

        self._build_menu()
        self._build_ui()
        self._auto_calc_worm()
        self._auto_calc_wheel()
        self.refresh_geom_plot()

    # --- i18n helper ---
    def _t(self, key):
        return self._lang.get(key, key)

    def _toggle_lang(self):
        if self._lang_code == "zh":
            self._lang = LANG_EN
            self._lang_code = "en"
        else:
            self._lang = LANG_ZH
            self._lang_code = "zh"
        # Refresh all tracked widgets
        for widget, key in self._i18n_widgets:
            try:
                widget.configure(text=self._t(key))
            except Exception:
                pass
        self.title(self._t("app_title"))
        # Refresh tabs
        self.nb.tab(self.tab_geom, text=self._t("tab_geom"))
        self.nb.tab(self.tab_mat, text=self._t("tab_mat"))
        self.nb.tab(self.tab_res, text=self._t("tab_res"))
        self.nb.tab(self.tab_fat, text=self._t("tab_fat"))
        self.nb.tab(self.tab_formula, text=self._t("tab_formula"))
        self._refresh_formula_views()
        self.refresh_geom_plot()

    def _track(self, widget, key):
        """Register widget for i18n refresh."""
        self._i18n_widgets.append((widget, key))
        return widget

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
        style.configure("TLabel", background=CLR_BG, foreground=CLR_TEXT, font=("", 11))
        style.configure("TLabelframe", background=CLR_CARD, borderwidth=0, relief="flat")
        style.configure("TLabelframe.Label", background=CLR_CARD, foreground=CLR_TEXT,
                        font=("", 11, "bold"))
        style.configure("TButton", padding=(10, 6), font=("", 10))
        style.map("TButton", background=[("active", "#E8E8ED")])
        style.configure("TNotebook", background=CLR_BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(20, 10), font=("", 11))
        style.map("TNotebook.Tab",
                  background=[("selected", CLR_CARD), ("!selected", CLR_SECTION)],
                  foreground=[("selected", CLR_ACCENT), ("!selected", CLR_TEXT2)])
        style.configure("Treeview", rowheight=28, font=("", 10), background=CLR_CARD,
                        fieldbackground=CLR_CARD, borderwidth=0)
        style.configure("Treeview.Heading", font=("", 10, "bold"), background=CLR_SECTION)
        style.map("Treeview", background=[("selected", "#D1E8FF")])

    # ------------------------------------------------------------------
    # Menu
    # ------------------------------------------------------------------
    def _build_menu(self):
        m = tk.Menu(self, bg=CLR_CARD, fg=CLR_TEXT, activebackground=CLR_ACCENT,
                    activeforeground="#FFF", bd=0)
        fm = tk.Menu(m, tearoff=0, bg=CLR_CARD, fg=CLR_TEXT)
        fm.add_command(label=self._t("menu_export"), command=self.export_xlsx)
        fm.add_separator()
        fm.add_command(label=self._t("menu_exit"), command=self.destroy)
        m.add_cascade(label=self._t("menu_file"), menu=fm)
        self.config(menu=m)

    # ------------------------------------------------------------------
    # Button helpers with good contrast
    # ------------------------------------------------------------------
    def _make_btn(self, parent, text_key, command, style="normal", **pack_kw):
        """Create a visible, high-contrast button."""
        if style == "accent":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 11, "bold"), fg="#FFFFFF", bg=CLR_ACCENT,
                            activebackground="#005EC4", activeforeground="#FFFFFF",
                            relief="raised", bd=1, padx=16, pady=7, cursor="hand2")
        elif style == "worm":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 10, "bold"), fg="#FFFFFF", bg=CLR_WORM,
                            activebackground="#C0392B", activeforeground="#FFFFFF",
                            relief="raised", bd=1, padx=10, pady=4, cursor="hand2")
        elif style == "wheel":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 10, "bold"), fg="#FFFFFF", bg=CLR_WHEEL,
                            activebackground="#005EC4", activeforeground="#FFFFFF",
                            relief="raised", bd=1, padx=10, pady=4, cursor="hand2")
        elif style == "green":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 10, "bold"), fg="#FFFFFF", bg=CLR_ACCENT2,
                            activebackground="#28A745", activeforeground="#FFFFFF",
                            relief="raised", bd=1, padx=12, pady=4, cursor="hand2")
        elif style == "danger":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 10), fg="#FFFFFF", bg="#FF3B30",
                            activebackground="#C0392B", activeforeground="#FFFFFF",
                            relief="raised", bd=1, padx=8, pady=3, cursor="hand2")
        elif style == "link":
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 10, "bold"), fg=CLR_ACCENT, bg=CLR_CARD,
                            activebackground=CLR_SECTION, activeforeground=CLR_ACCENT,
                            relief="flat", bd=0, padx=6, pady=2, cursor="hand2")
        else:
            # normal - dark text on visible background
            btn = tk.Button(parent, text=self._t(text_key), command=command,
                            font=("", 11), fg=CLR_TEXT, bg=CLR_BTN_BG,
                            activebackground="#D1D1D6", activeforeground=CLR_TEXT,
                            relief="raised", bd=1, padx=14, pady=7, cursor="hand2")
        self._track(btn, text_key)
        if pack_kw:
            btn.pack(**pack_kw)
        return btn

    def _auto_btn(self, parent, text_key, command):
        """Inline auto-calc button - visible blue text on white."""
        btn = tk.Button(parent, text=self._t(text_key), command=command,
                        font=("", 9, "bold"), fg=CLR_ACCENT, bg=CLR_CARD,
                        activebackground=CLR_SECTION, activeforeground="#005EC4",
                        relief="groove", bd=1, padx=4, pady=1, cursor="hand2")
        btn.pack(side="left", padx=(6, 0))
        self._track(btn, text_key)
        return btn

    # ------------------------------------------------------------------
    # UI skeleton
    # ------------------------------------------------------------------
    def _build_ui(self):
        # Language toggle in top-right
        top_bar = tk.Frame(self, bg=CLR_BG)
        top_bar.pack(fill="x", padx=12, pady=(4, 0))
        lang_btn = tk.Button(top_bar, text=self._t("lang_toggle"), command=self._toggle_lang,
                             font=("", 10, "bold"), fg=CLR_ACCENT, bg=CLR_CARD,
                             activebackground=CLR_SECTION, relief="groove", bd=1,
                             padx=10, pady=2, cursor="hand2")
        lang_btn.pack(side="right")
        self._track(lang_btn, "lang_toggle")

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        self.tab_geom = ttk.Frame(self.nb)
        self.tab_mat  = ttk.Frame(self.nb)
        self.tab_res  = ttk.Frame(self.nb)
        self.tab_fat  = ttk.Frame(self.nb)
        self.tab_formula = ttk.Frame(self.nb)

        self.nb.add(self.tab_geom, text=self._t("tab_geom"))
        self.nb.add(self.tab_mat,  text=self._t("tab_mat"))
        self.nb.add(self.tab_res,  text=self._t("tab_res"))
        self.nb.add(self.tab_fat,  text=self._t("tab_fat"))
        self.nb.add(self.tab_formula, text=self._t("tab_formula"))

        self._build_geom_tab()
        self._build_mat_tab()
        self._build_res_tab()
        self._build_fat_tab()
        self._build_formula_tab()

    # ==================================================================
    # Helpers
    # ==================================================================
    def _entry(self, parent, key, label_key, unit="", width=14, readonly=False):
        row = tk.Frame(parent, bg=CLR_CARD)
        row.pack(fill="x", padx=12, pady=3)
        lbl = tk.Label(row, text=self._t(label_key), bg=CLR_CARD, fg=CLR_TEXT,
                       font=("", 10), anchor="w", width=22)
        lbl.pack(side="left")
        self._track(lbl, label_key)
        var = tk.StringVar(value=self._defaults.get(key, ""))
        state = "readonly" if readonly else "normal"
        ent = tk.Entry(row, textvariable=var, width=width, font=("", 10),
                       relief="solid", bd=1,
                       bg=CLR_INPUT_BG if not readonly else CLR_SECTION,
                       fg=CLR_TEXT, highlightthickness=1,
                       highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER,
                       state=state)
        ent.pack(side="left", padx=(4, 0))
        if unit:
            tk.Label(row, text=unit, bg=CLR_CARD, fg=CLR_DIM, font=("", 9),
                     width=6, anchor="w").pack(side="left", padx=(6, 0))
        self.inputs[key] = var
        return row

    def _section_label(self, parent, text_key, bg=CLR_CARD):
        lbl = tk.Label(parent, text=self._t(text_key), bg=bg, fg=CLR_TEXT,
                       font=("", 11, "bold"), anchor="w")
        lbl.pack(fill="x", padx=12, pady=(10, 2))
        self._track(lbl, text_key)

    def _separator(self, parent):
        tk.Frame(parent, bg=CLR_BORDER, height=1).pack(fill="x", padx=12, pady=6)

    # ==================================================================
    # Tab 1: Geometry
    # ==================================================================
    def _build_geom_tab(self):
        left = tk.Frame(self.tab_geom, bg=CLR_BG)
        left.pack(side="left", fill="y", padx=(8, 4), pady=8)
        right = tk.Frame(self.tab_geom, bg=CLR_BG)
        right.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=8)

        # Scrollable left panel
        left_outer = tk.Frame(left, bg=CLR_BG)
        left_outer.pack(fill="both", expand=True)
        self.geom_canvas = tk.Canvas(left_outer, width=500, highlightthickness=0, bg=CLR_BG, bd=0)
        geom_scroll = ttk.Scrollbar(left_outer, orient="vertical", command=self.geom_canvas.yview)
        self.geom_canvas.configure(yscrollcommand=geom_scroll.set)
        geom_scroll.pack(side="right", fill="y")
        self.geom_canvas.pack(side="left", fill="both", expand=True)
        scroll_frame = tk.Frame(self.geom_canvas, bg=CLR_BG)
        self._geom_window = self.geom_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_frame.bind("<Configure>",
            lambda e: self.geom_canvas.configure(scrollregion=self.geom_canvas.bbox("all")))
        self.geom_canvas.bind("<Configure>",
            lambda e: self.geom_canvas.itemconfigure(self._geom_window, width=e.width))
        self.geom_canvas.bind_all("<Button-4>",
            lambda e: self.geom_canvas.yview_scroll(-3, "units"))
        self.geom_canvas.bind_all("<Button-5>",
            lambda e: self.geom_canvas.yview_scroll(3, "units"))

        # ---- Drive Card ----
        drive_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0,
                              highlightbackground=CLR_BORDER, highlightthickness=1)
        drive_card.pack(fill="x", padx=6, pady=(6, 4))
        self._section_label(drive_card, "drive_params")
        self._entry(drive_card, "T1_Nm", "T1_Nm", self._t("N_m"))
        self._entry(drive_card, "n1_rpm", "n1_rpm", self._t("rpm"))
        r = self._entry(drive_card, "ratio", "ratio", "")
        self._auto_btn(r, "calc_ratio", self._calc_ratio_from_z)
        self._entry(drive_card, "life_h", "life_h", self._t("h"))
        self._entry(drive_card, "steps", "steps", self._t("pts"))
        tk.Frame(drive_card, bg=CLR_CARD, height=6).pack()

        # ---- Worm Card ----
        worm_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0,
                             highlightbackground=CLR_WORM, highlightthickness=2)
        worm_card.pack(fill="x", padx=6, pady=4)
        worm_header = tk.Frame(worm_card, bg=CLR_CARD)
        worm_header.pack(fill="x", padx=12, pady=(8, 0))
        self._track(
            tk.Label(worm_header, text=self._t("worm_params"), bg=CLR_CARD, fg=CLR_WORM,
                     font=("", 12, "bold")),
            "worm_params").pack(side="left")
        self._make_btn(worm_header, "auto_calc_worm", self._auto_calc_worm,
                       style="worm", side="right")

        self._section_label(worm_card, "basic_params")
        self._entry(worm_card, "z1", "z1", "")
        self._entry(worm_card, "mn_mm", "mn_mm", self._t("mm"))
        self._entry(worm_card, "q", "q", "")
        self._entry(worm_card, "x1", "x1", "")
        self._entry(worm_card, "beta_deg", "beta_deg", self._t("deg"))
        self._entry(worm_card, "alpha_n_deg", "alpha_n_deg", self._t("deg"))
        self._entry(worm_card, "rho_f_mm", "rho_f_mm", self._t("mm"))

        self._separator(worm_card)
        self._section_label(worm_card, "computed_dims")
        self._entry(worm_card, "d1_mm", "d1_mm", self._t("mm"), readonly=True)
        self._entry(worm_card, "da1_mm", "da1_mm", self._t("mm"), readonly=True)
        self._entry(worm_card, "df1_mm", "df1_mm", self._t("mm"), readonly=True)
        self._entry(worm_card, "gamma_deg", "gamma_deg", self._t("deg"), readonly=True)
        self._entry(worm_card, "px_mm", "px_mm", self._t("mm"), readonly=True)
        self._entry(worm_card, "pz_mm", "pz_mm", self._t("mm"), readonly=True)
        self._entry(worm_card, "L_worm_mm", "L_worm_mm", self._t("mm"), readonly=True)
        tk.Frame(worm_card, bg=CLR_CARD, height=6).pack()

        # ---- Wheel Card ----
        wheel_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0,
                              highlightbackground=CLR_WHEEL, highlightthickness=2)
        wheel_card.pack(fill="x", padx=6, pady=4)
        wheel_header = tk.Frame(wheel_card, bg=CLR_CARD)
        wheel_header.pack(fill="x", padx=12, pady=(8, 0))
        self._track(
            tk.Label(wheel_header, text=self._t("wheel_params"), bg=CLR_CARD, fg=CLR_WHEEL,
                     font=("", 12, "bold")),
            "wheel_params").pack(side="left")
        self._make_btn(wheel_header, "auto_calc_wheel", self._auto_calc_wheel,
                       style="wheel", side="right")

        self._section_label(wheel_card, "basic_params")
        r = self._entry(wheel_card, "z2", "z2", "")
        self._auto_btn(r, "calc_z2", self._calc_z2_from_ratio)
        self._entry(wheel_card, "x2", "x2", "")
        self._entry(wheel_card, "b2_mm", "b2_mm", self._t("mm"))
        r2 = self._entry(wheel_card, "a_target_mm", "a_target_mm", self._t("mm"))
        self._auto_btn(r2, "calc_a", self._calc_center_distance)

        self._separator(wheel_card)
        self._section_label(wheel_card, "computed_dims")
        self._entry(wheel_card, "d2_mm", "d2_mm", self._t("mm"), readonly=True)
        self._entry(wheel_card, "da2_mm", "da2_mm", self._t("mm"), readonly=True)
        self._entry(wheel_card, "df2_mm", "df2_mm", self._t("mm"), readonly=True)
        tk.Frame(wheel_card, bg=CLR_CARD, height=6).pack()

        # ---- Common Card ----
        common_card = tk.Frame(scroll_frame, bg=CLR_CARD, bd=0,
                               highlightbackground=CLR_BORDER, highlightthickness=1)
        common_card.pack(fill="x", padx=6, pady=4)
        self._section_label(common_card, "common_params")
        self._entry(common_card, "cross_angle_deg", "cross_angle_deg", self._t("deg"))
        self._entry(common_card, "mu", "mu", "")
        self._separator(common_card)
        self._entry(common_card, "KA", "KA", "")
        self._entry(common_card, "KV", "KV", "")
        self._entry(common_card, "KHb", "KHb", "")
        self._entry(common_card, "KFb", "KFb", "")
        self._entry(common_card, "temp_C", "temp_C", self._t("C"))
        tk.Frame(common_card, bg=CLR_CARD, height=6).pack()

        # ---- Action buttons ----
        btn_frame = tk.Frame(scroll_frame, bg=CLR_BG)
        btn_frame.pack(fill="x", padx=6, pady=(8, 12))
        self._make_btn(btn_frame, "btn_refresh", self._on_refresh_diagram,
                       style="normal", side="left", padx=(0, 8))
        self._make_btn(btn_frame, "btn_calc", self.run,
                       style="accent", side="left")

        # ---- Right side: diagram ----
        diag_card = tk.Frame(right, bg=CLR_CARD, bd=0,
                             highlightbackground=CLR_BORDER, highlightthickness=1)
        diag_card.pack(fill="both", expand=True)
        fig = Figure(figsize=(8, 7), dpi=100, facecolor=CLR_CARD)
        gs = fig.add_gridspec(2, 1, height_ratios=[1.15, 0.95], hspace=0.18)
        self.ax_geom_mesh = fig.add_subplot(gs[0, 0])
        self.ax_geom_axial = fig.add_subplot(gs[1, 0])
        self.canvas_geom = FigureCanvasTkAgg(fig, master=diag_card)
        self.canvas_geom.get_tk_widget().pack(fill="both", expand=True, padx=2, pady=2)
        self.geom_check_var = tk.StringVar(value=self._t("geom_check_wait"))
        tk.Label(right, textvariable=self.geom_check_var, bg=CLR_BG, fg=CLR_ACCENT,
                 font=("", 10), anchor="w").pack(fill="x", padx=4, pady=(4, 0))

    # ==================================================================
    # Auto-calc
    # ==================================================================
    def _safe_float(self, key, default=0.0):
        try:
            return float(self.inputs[key].get())
        except (ValueError, KeyError):
            return default

    def _auto_calc_worm(self):
        mn = self._safe_float("mn_mm", 2.5)
        x1 = self._safe_float("x1", 0.0)
        z1 = int(self._safe_float("z1", 2))
        q = self._safe_float("q", 10)
        beta_deg = self._safe_float("beta_deg", 0.0)
        beta_rad = math.radians(beta_deg)

        if 0.5 < beta_deg < 80.0 and abs(math.tan(beta_rad)) > 1e-6:
            d1 = z1 * mn / math.tan(beta_rad)
            q = d1 / mn - 2.0 * x1 if mn > 0 else q
            self.inputs["q"].set(f"{q:.3f}")
        else:
            d1 = (q + 2.0 * x1) * mn
            beta_rad = math.atan2(z1 * mn, d1) if d1 > 0 else math.radians(5.0)
            beta_deg = math.degrees(beta_rad)
            self.inputs["beta_deg"].set(f"{beta_deg:.3f}")

        da1 = d1 + 2.0 * mn
        df1 = d1 - 2.4 * mn
        gamma = math.degrees(math.atan2(z1 * mn, d1)) if d1 > 0 else beta_deg
        px = mn * math.pi
        pz = px * z1
        L_worm = pz * 3.0 + 2.0 * mn

        self.inputs["d1_mm"].set(f"{d1:.3f}")
        self.inputs["da1_mm"].set(f"{da1:.3f}")
        self.inputs["df1_mm"].set(f"{df1:.3f}")
        self.inputs["gamma_deg"].set(f"{gamma:.3f}")
        self.inputs["px_mm"].set(f"{px:.3f}")
        self.inputs["pz_mm"].set(f"{pz:.3f}")
        self.inputs["L_worm_mm"].set(f"{L_worm:.2f}")

    def _auto_calc_wheel(self):
        mn = self._safe_float("mn_mm", 2.5)
        x2 = self._safe_float("x2", 0.0)
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError, KeyError):
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
        x2 = self._safe_float("x2", 0.0)
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            ratio = self._safe_float("ratio", 25)
            z1 = int(self._safe_float("z1", 2))
            z2 = int(round(ratio * z1))
        d1 = self._safe_float("d1_mm", 0.0)
        if d1 <= 0:
            q = self._safe_float("q", 10)
            x1 = self._safe_float("x1", 0.0)
            d1 = (q + 2.0 * x1) * mn
        d2 = (z2 + 2.0 * x2) * mn
        a = 0.5 * (d1 + d2)
        self.inputs["a_target_mm"].set(f"{a:.3f}")

    def _on_refresh_diagram(self):
        self._auto_calc_worm()
        self._auto_calc_wheel()
        self.refresh_geom_plot()

    # ==================================================================
    # Geometry diagram (uses ASCII-safe labels to avoid CJK garble)
    # ==================================================================
    def refresh_geom_plot(self):
        ax1 = self.ax_geom_mesh
        ax2 = self.ax_geom_axial
        ax1.clear()
        ax2.clear()
        for ax in (ax1, ax2):
            ax.set_aspect("equal", adjustable="box")
            ax.axis("off")
            ax.set_facecolor(CLR_CARD)
        fig = ax1.get_figure()
        fig.set_facecolor(CLR_CARD)

        fp = {"fontfamily": _MPL_FONT or "DejaVu Sans"}
        ds = 8

        mn = self._safe_float("mn_mm", 2.5)
        x1 = self._safe_float("x1", 0.0)
        x2 = self._safe_float("x2", 0.0)
        z1 = int(self._safe_float("z1", 2))
        ratio = self._safe_float("ratio", 25)
        z2 = int(round(ratio * z1))
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            pass
        d1 = self._safe_float("d1_mm", 0.0)
        if d1 <= 0:
            q = self._safe_float("q", 10)
            d1 = (q + 2.0 * x1) * mn
        da1 = d1 + 2.0 * mn
        df1 = d1 - 2.4 * mn
        d2 = (z2 + 2.0 * x2) * mn
        da2 = d2 + 2.0 * mn * (1.0 + x2)
        df2 = d2 - 2.0 * mn * (1.2 - x2)
        a_calc = 0.5 * (d1 + d2)
        b2 = self._safe_float("b2_mm", 18)
        alpha_n = self._safe_float("alpha_n_deg", 20)
        beta_deg = self._safe_float("beta_deg", 0.0)
        if beta_deg <= 0 and d1 > 0:
            beta_deg = math.degrees(math.atan2(z1 * mn, d1))
        gamma = math.degrees(math.atan2(z1 * mn, d1)) if d1 > 0 else beta_deg
        px = mn * math.pi
        pz = px * z1
        a_target_txt = self.inputs.get("a_target_mm", tk.StringVar(value="")).get().strip()
        a_target = float(a_target_txt) if a_target_txt else None
        a_display = a_target if a_target else a_calc
        if a_target is not None:
            self.geom_check_var.set(
                f"a_calc={a_calc:.3f} mm, a_target={a_target:.3f} mm, "
                f"da={a_target - a_calc:+.3f} mm, beta={beta_deg:.2f} deg")
        else:
            self.geom_check_var.set(f"a_calc={a_calc:.3f} mm, beta={beta_deg:.2f} deg")

        # ---------- Section A: normal mesh section ----------
        scale = 68.0 / max(d2, 1)
        r2 = 0.5 * d2 * scale
        ra2 = 0.5 * da2 * scale
        rf2 = max(0.5 * df2 * scale, 0.8)
        cx_wheel, cy_wheel = 0.0, 28.0
        cx_worm, cy_worm = 0.0, cy_wheel - a_display * scale
        rw = max(0.5 * d1 * scale, 10.0)
        tooth_depth = max((ra2 - rf2) * 0.95, 1.5)

        # Wheel outline in attachment-like style
        n_draw_teeth = max(10, min(z2, 28))
        phi = np.linspace(0, 2 * np.pi, 1000, endpoint=False)
        mod = 0.5 * (1.0 + np.cos(n_draw_teeth * phi))
        radius = r2 + tooth_depth * (mod ** 1.9) - 0.48 * tooth_depth
        xw = cx_wheel + radius * np.cos(phi)
        yw = cy_wheel + radius * np.sin(phi)
        ax1.fill(xw, yw, facecolor="#EEF1F6", edgecolor="#626670", linewidth=1.4, zorder=2)
        ax1.add_patch(mpatches.Circle((cx_wheel, cy_wheel), r2 * 0.33,
                                      facecolor=CLR_CARD, edgecolor="#626670", linewidth=1.2, zorder=3))
        ax1.add_patch(mpatches.Circle((cx_wheel, cy_wheel), r2,
                                      fill=False, edgecolor="#8C8F97", linewidth=1.0,
                                      linestyle=(0, (4, 3)), zorder=2))
        ax1.plot([cx_wheel - ra2 * 1.18, cx_wheel + ra2 * 1.18], [cy_wheel, cy_wheel],
                 color="#9CA0A8", linewidth=0.9, linestyle=(0, (4, 2)))
        ax1.plot([cx_wheel, cx_wheel], [cy_wheel - ra2 * 1.18, cy_wheel + ra2 * 1.18],
                 color="#9CA0A8", linewidth=0.9, linestyle=(0, (4, 2)))

        # Worm generating arc and teeth traces below wheel
        lower_r = max(rw * 1.55, ra2 * 0.95)
        ax1.add_patch(mpatches.Arc((cx_worm, cy_worm), 2 * lower_r, 2 * lower_r * 0.86,
                                   theta1=24, theta2=156, color="#7B7F88", linewidth=1.2))
        ax1.add_patch(mpatches.Arc((cx_worm, cy_worm), 2 * (lower_r - tooth_depth * 0.7), 2 * (lower_r - tooth_depth * 0.7) * 0.86,
                                   theta1=24, theta2=156, color="#A1A6B0", linewidth=1.0, linestyle=(0, (5, 3))))
        ts = np.linspace(0.18, 0.82, 7)
        for t in ts:
            ang = math.radians(24 + 132 * t)
            rr = lower_r - 0.15 * tooth_depth * math.cos(5 * t * np.pi)
            x0 = cx_worm + rr * math.cos(ang)
            y0 = cy_worm + 0.86 * rr * math.sin(ang)
            dx = -math.sin(ang)
            dy = math.cos(ang)
            ax1.plot([x0 - 2.8 * dx, x0 + 2.8 * dx], [y0 - 2.8 * dy, y0 + 2.8 * dy],
                     color="#5E636D", linewidth=1.1)

        # Contact marker and action line
        contact = (cx_wheel, cy_wheel - r2)
        ax1.scatter([contact[0]], [contact[1]], s=28, color="#111111", zorder=6)
        ax1.plot([cx_wheel + r2 * 0.14, cx_worm + lower_r * 0.52],
                 [cy_wheel - r2 * 0.87, cy_worm + lower_r * 0.38],
                 color="#7E838D", linewidth=1.0)

        # Dimension and labels
        ax1.annotate("", xy=(ra2 * 1.22, cy_wheel), xytext=(ra2 * 1.22, cy_worm),
                     arrowprops=dict(arrowstyle="<->", color=CLR_DIM, lw=1.1))
        ax1.text(ra2 * 1.26, 0.5 * (cy_wheel + cy_worm), f"a={a_display:.2f} mm",
                 fontsize=ds, color=CLR_DIM, rotation=90, va="center", **fp)
        ax1.text(cx_wheel, cy_wheel + ra2 * 1.26, self._t("wheel_label"),
                 fontsize=10, color="#4A5D8C", fontweight="bold", ha="center", **fp)
        ax1.text(cx_worm, cy_worm - lower_r * 0.55, self._t("worm_label"),
                 fontsize=10, color="#8A4D49", fontweight="bold", ha="center", **fp)
        ax1.text(0.02, 0.03, "Section A: normal mesh section", transform=ax1.transAxes,
                 fontsize=8, color=CLR_TEXT2, ha="left", **fp)
        lim_x = ra2 * 1.35
        ax1.set_xlim(-lim_x, lim_x)
        ax1.set_ylim(cy_worm - lower_r * 0.88, cy_wheel + ra2 * 1.36)

        # ---------- Section B: axial section ----------
        scale2 = 62.0 / max(d2, 1)
        r1 = 0.5 * d1 * scale2
        ra1 = 0.5 * da1 * scale2
        r2b = 0.5 * d2 * scale2
        cx2, cy2 = 0.0, 0.0
        worm_len = b2 * scale2 * 1.85
        half_len = worm_len / 2.0
        wheel_cy = a_display * scale2

        ax2.add_patch(mpatches.FancyBboxPatch(
            (cx2 - half_len, cy2 - ra1), worm_len, 2 * ra1,
            boxstyle="round,pad=0.22,rounding_size=4",
            facecolor="#FFF0EE", edgecolor="#A35B55", linewidth=1.2))
        ax2.plot([cx2 - half_len, cx2 + half_len], [cy2 + r1, cy2 + r1], color="#A35B55", linewidth=1.7)
        ax2.plot([cx2 - half_len, cx2 + half_len], [cy2 - r1, cy2 - r1], color="#A35B55", linewidth=1.7)
        pitch_s = max(px * scale2, 2.4)
        slope = math.tan(math.radians(beta_deg)) * 0.28
        for i in range(-4, 5):
            x0 = cx2 + i * pitch_s * 0.58
            ax2.plot([x0 - pitch_s * 0.42, x0 + pitch_s * 0.42],
                     [cy2 - ra1 * 0.82 + slope * (x0 - cx2), cy2 + ra1 * 0.82 + slope * (x0 - cx2)],
                     color="#B96E67", linewidth=0.9, alpha=0.85)

        ax2.add_patch(mpatches.Arc((cx2, wheel_cy), 2 * r2b * 1.05, 2 * r2b * 0.86,
                                   theta1=196, theta2=344, color="#4F74B3", linewidth=2.0))
        ax2.add_patch(mpatches.Arc((cx2, wheel_cy), 2 * r2b * 0.88, 2 * r2b * 0.72,
                                   theta1=196, theta2=344, color="#7A98CA", linewidth=1.2, linestyle=(0, (4, 2))))
        ax2.scatter([0], [wheel_cy - r2b * 0.42], s=24, color="#111111", zorder=6)

        ax2.annotate("", xy=(half_len + 8, cy2), xytext=(half_len + 8, wheel_cy),
                     arrowprops=dict(arrowstyle="<->", color=CLR_DIM, lw=1.0))
        ax2.text(half_len + 10, 0.5 * wheel_cy, f"a={a_display:.2f}", fontsize=ds,
                 color=CLR_DIM, rotation=90, va="center", **fp)
        ax2.annotate("", xy=(-half_len, cy2 - ra1 - 5), xytext=(half_len, cy2 - ra1 - 5),
                     arrowprops=dict(arrowstyle="<->", color=CLR_DIM, lw=1.0))
        ax2.text(0, cy2 - ra1 - 7.8, f"b2={b2:.1f} mm", fontsize=ds, color=CLR_DIM, ha="center", **fp)
        ax2.text(-half_len + 2, cy2 + ra1 + 4.5, f"beta={beta_deg:.2f} deg", fontsize=8, color="#8A4D49", **fp)
        ax2.text(0.02, 0.03, "Section B: axial section", transform=ax2.transAxes,
                 fontsize=8, color=CLR_TEXT2, ha="left", **fp)

        info_text = (
            f"mn={mn:.3f}  z1={z1}  z2={z2}\n"
            f"alpha_n={alpha_n:.1f} deg  beta={beta_deg:.2f} deg\n"
            f"gamma={gamma:.2f} deg  px={px:.2f} mm  pz={pz:.2f} mm"
        )
        ax2.text(0.98, 0.98, info_text, transform=ax2.transAxes, fontsize=8,
                 color=CLR_TEXT2, va="top", ha="right",
                 bbox=dict(boxstyle="round,pad=0.35", facecolor=CLR_SECTION,
                           edgecolor=CLR_BORDER, alpha=0.96),
                 family="monospace")
        ax2.set_xlim(-half_len - 16, half_len + 18)
        ax2.set_ylim(cy2 - ra1 - 14, wheel_cy + r2b + 10)
        ax1.set_title(self._t("diagram_title"), fontsize=13, color=CLR_TEXT, pad=6,
                      fontweight="bold", **fp)
        self.canvas_geom.draw()

    # ==================================================================
    # Tab 2: Materials (GUI form instead of raw JSON)
    # ==================================================================
    def _build_mat_tab(self):
        top = tk.Frame(self.tab_mat, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=10, pady=10)

        # ---- Worm material card (form-based) ----
        g1 = tk.Frame(top, bg=CLR_CARD, highlightbackground=CLR_WORM, highlightthickness=2)
        g1.pack(fill="x", pady=(0, 8))

        g1_header = tk.Frame(g1, bg=CLR_CARD)
        g1_header.pack(fill="x", padx=12, pady=(8, 4))
        self._track(
            tk.Label(g1_header, text=self._t("mat_worm_title"), bg=CLR_CARD, fg=CLR_WORM,
                     font=("", 12, "bold")),
            "mat_worm_title").pack(side="left")

        # Selector row
        sel_row = tk.Frame(g1, bg=CLR_CARD)
        sel_row.pack(fill="x", padx=12, pady=4)
        self._track(tk.Label(sel_row, text=self._t("mat_select"), bg=CLR_CARD, fg=CLR_TEXT,
                             font=("", 10)), "mat_select").pack(side="left")
        self.steel_var = tk.StringVar(
            value="37CrS4.json" if "37CrS4.json" in self.steel_db else list(self.steel_db.keys())[0])
        self.steel_cb = ttk.Combobox(sel_row, textvariable=self.steel_var,
                                     values=list(self.steel_db.keys()), state="readonly", width=30)
        self.steel_cb.pack(side="left", padx=8)
        self._make_btn(sel_row, "mat_load", self.load_steel, style="link", side="left")
        self._make_btn(sel_row, "mat_import", self.import_steel, style="link", side="left", padx=4)

        # Form fields for worm material
        form1 = tk.Frame(g1, bg=CLR_CARD)
        form1.pack(fill="x", padx=12, pady=(4, 8))

        self.steel_fields = {}
        fields = [
            ("mat_name", "name", 25), ("mat_standard", "standard", 20),
            ("mat_E", "E_GPa", 12), ("mat_nu", "nu", 12),
            ("mat_yield", "Rp02_MPa", 12), ("mat_tensile", "Rm_MPa", 12),
            ("mat_hardness", "HRC", 12),
        ]

        # Two-column layout
        col_left = tk.Frame(form1, bg=CLR_CARD)
        col_left.pack(side="left", fill="both", expand=True)
        col_right = tk.Frame(form1, bg=CLR_CARD)
        col_right.pack(side="left", fill="both", expand=True)

        for idx, (label_key, field_key, width) in enumerate(fields):
            parent = col_left if idx < 4 else col_right
            row = tk.Frame(parent, bg=CLR_CARD)
            row.pack(fill="x", pady=2)
            lbl = tk.Label(row, text=self._t(label_key), bg=CLR_CARD, fg=CLR_TEXT,
                           font=("", 10), width=16, anchor="w")
            lbl.pack(side="left")
            self._track(lbl, label_key)
            var = tk.StringVar()
            tk.Entry(row, textvariable=var, width=width, font=("", 10), relief="solid", bd=1,
                     bg=CLR_INPUT_BG, fg=CLR_TEXT, highlightthickness=1,
                     highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER).pack(side="left", padx=4)
            self.steel_fields[field_key] = var

        # Notes field
        notes_row = tk.Frame(form1, bg=CLR_CARD)
        notes_row.pack(fill="x", pady=2)
        nlbl = tk.Label(notes_row, text=self._t("mat_notes"), bg=CLR_CARD, fg=CLR_TEXT,
                        font=("", 10), width=16, anchor="w")
        nlbl.pack(side="left")
        self._track(nlbl, "mat_notes")
        self.steel_notes_var = tk.StringVar()
        tk.Entry(notes_row, textvariable=self.steel_notes_var, width=60, font=("", 10),
                 relief="solid", bd=1, bg=CLR_INPUT_BG, fg=CLR_TEXT,
                 highlightthickness=1, highlightcolor=CLR_ACCENT,
                 highlightbackground=CLR_BORDER).pack(side="left", padx=4)

        # Save button for worm material
        save_row1 = tk.Frame(g1, bg=CLR_CARD)
        save_row1.pack(fill="x", padx=12, pady=(0, 8))
        self._make_btn(save_row1, "sn_apply", self._apply_steel_form, style="worm", side="left")
        self._make_btn(save_row1, "mat_save_json", self._save_steel_json, style="normal", side="left", padx=8)

        # ---- Wheel material card ----
        g2 = tk.Frame(top, bg=CLR_CARD, highlightbackground=CLR_WHEEL, highlightthickness=2)
        g2.pack(fill="both", expand=True, pady=(0, 4))

        g2_header = tk.Frame(g2, bg=CLR_CARD)
        g2_header.pack(fill="x", padx=12, pady=(8, 4))
        self._track(
            tk.Label(g2_header, text=self._t("mat_wheel_title"), bg=CLR_CARD, fg=CLR_WHEEL,
                     font=("", 12, "bold")),
            "mat_wheel_title").pack(side="left")

        sel_row2 = tk.Frame(g2, bg=CLR_CARD)
        sel_row2.pack(fill="x", padx=12, pady=4)
        self._track(tk.Label(sel_row2, text=self._t("mat_base_template"), bg=CLR_CARD,
                             fg=CLR_TEXT, font=("", 10)), "mat_base_template").pack(side="left")
        self.wheel_var = tk.StringVar(
            value="PA66_modified_draft.json" if "PA66_modified_draft.json" in self.poly_db
            else list(self.poly_db.keys())[0])
        self.wheel_cb = ttk.Combobox(sel_row2, textvariable=self.wheel_var,
                                     values=list(self.poly_db.keys()), state="readonly", width=30)
        self.wheel_cb.pack(side="left", padx=8)
        self._make_btn(sel_row2, "mat_load", self.load_wheel, style="link", side="left")
        self._make_btn(sel_row2, "mat_import", self.import_wheel, style="link", side="left", padx=4)

        # Wheel form fields
        form2 = tk.Frame(g2, bg=CLR_CARD)
        form2.pack(fill="x", padx=12, pady=4)
        wheel_fields_def = [
            ("mat_name", "w_name", 25), ("mat_nu", "w_nu", 12),
        ]
        self.wheel_fields = {}
        for label_key, field_key, width in wheel_fields_def:
            row = tk.Frame(form2, bg=CLR_CARD)
            row.pack(fill="x", pady=2)
            lbl = tk.Label(row, text=self._t(label_key), bg=CLR_CARD, fg=CLR_TEXT,
                           font=("", 10), width=20, anchor="w")
            lbl.pack(side="left")
            self._track(lbl, label_key)
            var = tk.StringVar()
            tk.Entry(row, textvariable=var, width=width, font=("", 10), relief="solid", bd=1,
                     bg=CLR_INPUT_BG, fg=CLR_TEXT, highlightthickness=1,
                     highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER).pack(side="left", padx=4)
            self.wheel_fields[field_key] = var

        # E(T) row
        row3 = tk.Frame(g2, bg=CLR_CARD)
        row3.pack(fill="x", padx=12, pady=4)
        lbl3 = tk.Label(row3, text=self._t("mat_Et_label"), bg=CLR_CARD, fg=CLR_TEXT,
                        font=("", 10), width=20, anchor="w")
        lbl3.pack(side="left")
        self._track(lbl3, "mat_Et_label")
        self.Et_var = tk.StringVar(value=self._format_Et())
        tk.Entry(row3, textvariable=self.Et_var, width=50, font=("", 10), relief="solid", bd=1,
                 bg=CLR_INPUT_BG, fg=CLR_TEXT, highlightthickness=1,
                 highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER).pack(side="left", padx=4)
        hint_lbl = tk.Label(row3, text=self._t("mat_Et_hint"), bg=CLR_CARD, fg=CLR_DIM, font=("", 9))
        hint_lbl.pack(side="left", padx=4)
        self._track(hint_lbl, "mat_Et_hint")

        # S-N table
        sn_frame = tk.Frame(g2, bg=CLR_CARD)
        sn_frame.pack(fill="both", padx=12, pady=6, expand=True)
        self._track(
            tk.Label(sn_frame, text=self._t("sn_title"), bg=CLR_CARD, fg=CLR_TEXT,
                     font=("", 10, "bold"), anchor="w"),
            "sn_title").pack(anchor="w")

        cols = ("temp_C", "N", "contact_MPa", "root_MPa")
        self.sn_table = ttk.Treeview(sn_frame, columns=cols, show="headings", height=5)
        self.sn_table.heading("temp_C", text="Temp C")
        self.sn_table.heading("N", text="Cycles N")
        self.sn_table.heading("contact_MPa", text="Contact MPa")
        self.sn_table.heading("root_MPa", text="Root MPa")
        for col in cols:
            self.sn_table.column(col, width=140, anchor="center")
        self.sn_table.pack(side="left", fill="both", expand=True)
        ttk.Scrollbar(sn_frame, orient="vertical",
                      command=self.sn_table.yview).pack(side="left", fill="y")

        # S-N input
        ctrl = tk.Frame(g2, bg=CLR_CARD)
        ctrl.pack(fill="x", padx=12, pady=6)
        self.sn_temp_var = tk.StringVar(value=self._defaults.get("temp_C", "80"))
        self.sn_N_var = tk.StringVar(value="1e6")
        self.sn_contact_var = tk.StringVar(value="110")
        self.sn_root_var = tk.StringVar(value="45")
        for var, w in [(self.sn_temp_var, 8), (self.sn_N_var, 12),
                       (self.sn_contact_var, 12), (self.sn_root_var, 12)]:
            tk.Entry(ctrl, textvariable=var, width=w, font=("", 10), relief="solid", bd=1,
                     bg=CLR_INPUT_BG, fg=CLR_TEXT, highlightthickness=1,
                     highlightcolor=CLR_ACCENT, highlightbackground=CLR_BORDER).pack(side="left", padx=(0, 4))
        self._make_btn(ctrl, "sn_add", self.add_sn_row, style="link", side="left", padx=4)
        self._make_btn(ctrl, "sn_delete", self.delete_sn_rows, style="danger", side="left", padx=4)
        self._make_btn(ctrl, "sn_apply", self.apply_wheel_card, style="green", side="left", padx=8)

        self._load_sn_table_from_wheel()
        self._populate_steel_form()
        self._populate_wheel_form()

    # --- Steel form populate / apply ---
    def _populate_steel_form(self):
        d = self.steel
        self.steel_fields["name"].set(d.get("name", ""))
        self.steel_fields["standard"].set(d.get("standard", ""))
        e = d.get("elastic", {})
        self.steel_fields["E_GPa"].set(str(e.get("E_GPa", 210.0)))
        self.steel_fields["nu"].set(str(e.get("nu", 0.3)))
        self.steel_fields["Rp02_MPa"].set(str(d.get("Rp02_MPa", "")))
        self.steel_fields["Rm_MPa"].set(str(d.get("Rm_MPa", "")))
        self.steel_fields["HRC"].set(str(d.get("HRC", "")))
        self.steel_notes_var.set(d.get("notes", ""))

    def _apply_steel_form(self):
        """Apply form fields back to steel dict."""
        self.steel["name"] = self.steel_fields["name"].get()
        self.steel["standard"] = self.steel_fields["standard"].get()
        try:
            self.steel["elastic"] = {
                "E_GPa": float(self.steel_fields["E_GPa"].get() or 210),
                "nu": float(self.steel_fields["nu"].get() or 0.3),
            }
        except ValueError:
            pass
        for k in ["Rp02_MPa", "Rm_MPa", "HRC"]:
            v = self.steel_fields[k].get().strip()
            if v:
                try:
                    self.steel[k] = float(v)
                except ValueError:
                    self.steel[k] = v
        self.steel["notes"] = self.steel_notes_var.get()
        messagebox.showinfo("OK", "Worm material applied (session only).")

    def _save_steel_json(self):
        self._apply_steel_form()
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.steel, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("OK", f"Saved: {path}")

    def _populate_wheel_form(self):
        d = self.wheel
        self.wheel_fields["w_name"].set(d.get("name", ""))
        self.wheel_fields["w_nu"].set(str(d.get("nu", 0.4)))

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
        self._track(
            tk.Label(top, text=self._t("res_footer"), bg=CLR_BG, fg=CLR_TEXT2, font=("", 10)),
            "res_footer").pack(anchor="w", padx=6, pady=4)

    # ==================================================================
    # Tab 4: Fatigue
    # ==================================================================
    def _build_fat_tab(self):
        top = tk.Frame(self.tab_fat, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=10, pady=10)
        cards = tk.Frame(top, bg=CLR_BG)
        cards.pack(fill="x")

        worm_box = tk.Frame(cards, bg=CLR_CARD, highlightbackground=CLR_WORM, highlightthickness=2)
        worm_box.pack(side="left", fill="both", expand=True, padx=(0, 6))
        self._track(
            tk.Label(worm_box, text=self._t("worm_output"), bg=CLR_CARD, fg=CLR_WORM,
                     font=("", 11, "bold")),
            "worm_output").pack(padx=8, pady=(8, 4), anchor="w")
        self.worm_out = ttk.Treeview(worm_box, columns=("k", "v"), show="headings", height=9)
        self.worm_out.heading("k", text="Param")
        self.worm_out.heading("v", text="Value")
        self.worm_out.column("k", width=180, anchor="w")
        self.worm_out.column("v", width=220, anchor="w")
        self.worm_out.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        wheel_box = tk.Frame(cards, bg=CLR_CARD, highlightbackground=CLR_WHEEL, highlightthickness=2)
        wheel_box.pack(side="left", fill="both", expand=True, padx=(6, 0))
        self._track(
            tk.Label(wheel_box, text=self._t("wheel_output"), bg=CLR_CARD, fg=CLR_WHEEL,
                     font=("", 11, "bold")),
            "wheel_output").pack(padx=8, pady=(8, 4), anchor="w")
        self.wheel_out = ttk.Treeview(wheel_box, columns=("k", "v"), show="headings", height=9)
        self.wheel_out.heading("k", text="Param")
        self.wheel_out.heading("v", text="Value")
        self.wheel_out.column("k", width=180, anchor="w")
        self.wheel_out.column("v", width=220, anchor="w")
        self.wheel_out.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.fat_text = tk.Text(top, wrap="word", bd=0, relief="flat", bg=CLR_CARD,
                                fg=CLR_TEXT, font=("", 10), padx=12, pady=8)
        self.fat_text.pack(fill="both", expand=True, padx=2, pady=(10, 0))
        self.fat_text.insert("1.0", self._t("fat_wait"))

    # ==================================================================
    # Tab 5: Formula Notes
    # ==================================================================
    def _build_formula_tab(self):
        top = tk.Frame(self.tab_formula, bg=CLR_BG)
        top.pack(fill="both", expand=True, padx=10, pady=10)

        card = tk.Frame(top, bg=CLR_CARD, highlightbackground=CLR_BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)

        self._track(
            tk.Label(card, text=self._t("formula_title"), bg=CLR_CARD, fg=CLR_TEXT,
                     font=("", 12, "bold"), anchor="w"),
            "formula_title").pack(fill="x", padx=12, pady=(10, 6))
        body = tk.Frame(card, bg=CLR_CARD)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        left = tk.Frame(body, bg=CLR_CARD)
        left.pack(side="left", fill="both", expand=True)
        right = tk.Frame(body, bg=CLR_CARD)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.formula_nb = ttk.Notebook(left)
        self.formula_nb.pack(fill="both", expand=True)

        self.formula_tabs = {}
        for key in ("geom", "stress", "fatigue"):
            tab = ttk.Frame(self.formula_nb)
            self.formula_nb.add(tab, text=key)
            tree = ttk.Treeview(tab, columns=("formula", "params", "purpose"), show="headings", height=14)
            tree.heading("formula", text="Formula")
            tree.heading("params", text="Parameters")
            tree.heading("purpose", text="Purpose")
            tree.column("formula", width=290, anchor="w")
            tree.column("params", width=210, anchor="w")
            tree.column("purpose", width=260, anchor="w")
            vs = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vs.set)
            tree.pack(side="left", fill="both", expand=True)
            vs.pack(side="left", fill="y")
            self.formula_tabs[key] = tree

        fig = Figure(figsize=(5.8, 4.8), dpi=100, facecolor=CLR_CARD)
        self.ax_formula_diag = fig.add_subplot(111)
        self.canvas_formula = FigureCanvasTkAgg(fig, master=right)
        self.canvas_formula.get_tk_widget().pack(fill="both", expand=True)

        self.formula_hint = tk.Label(right, bg=CLR_CARD, fg=CLR_TEXT2, anchor="w",
                                     justify="left", font=("", 10))
        self.formula_hint.pack(fill="x", pady=(6, 0))
        self._refresh_formula_views()

    def _refresh_formula_views(self):
        if not hasattr(self, "formula_tabs"):
            return
        if self._lang_code == "zh":
            tab_names = {
                "geom": "几何与啮合",
                "stress": "载荷与强度",
                "fatigue": "寿命与损伤",
            }
            rows = {
                "geom": [
                    ("d1=(q+2*x1)*mn 或 d1=z1*mn/tan(beta)", "mn,q,x1,z1,beta", "计算蜗杆分度圆直径"),
                    ("da1=d1+2*mn; df1=d1-2.4*mn", "d1,mn", "蜗杆齿顶/齿根圆直径"),
                    ("d2=(z2+2*x2)*mn; da2=d2+2*mn*(1+x2)", "z2,x2,mn", "蜗轮主尺寸"),
                    ("a=0.5*(d1+d2)", "d1,d2", "中心距"),
                    ("gamma=atan(z1*mn/d1); px=pi*mn; pz=z1*px", "z1,mn,d1", "导程角与导程参数"),
                ],
                "stress": [
                    ("E'=2/((1-nu1^2)/E1+(1-nu2^2)/E2)", "E1,E2,nu1,nu2", "赫兹接触等效弹性模量"),
                    ("Ft=2*T1*1000/d1", "T1,d1", "圆周力"),
                    ("Fn=Ft/(cos(alpha_n)*sin(gamma))", "Ft,alpha_n,gamma", "法向力"),
                    ("rho_eq=0.5*d1*sin(alpha_n)*cos(gamma)", "d1,alpha_n,gamma", "等效曲率半径"),
                    ("p_base=0.418*sqrt(Fn*KH*E'*1000/(rho_eq*b))", "Fn,KH,E',rho_eq,b", "接触应力基值"),
                    ("sigma_base=Fn*KF*YF/(b*mn)", "Fn,KF,YF,b,mn", "齿根应力基值"),
                ],
                "fatigue": [
                    ("N_life=n1*60*life_h/i", "n1,life_h,i", "寿命循环次数"),
                    ("SF_root=sigma_allow(N_life)/max(sigma_F)", "SN表,sigma_F", "齿根安全系数"),
                    ("SF_contact=p_allow(N_life)/max(p)", "SN表,p", "接触安全系数"),
                    ("sigma_amp=(sigma_F_max-sigma_F_min)/2", "sigma_F", "等效应力幅"),
                    ("D=n_cycles_total/N_allow(sigma_amp)", "sigma_amp,SN表", "Miner累积损伤"),
                ],
            }
            hint = "右图是参数示意图：标注了 d1、d2、a、beta、alpha_n 和啮合接触点。"
        else:
            tab_names = {
                "geom": "Geometry",
                "stress": "Load & Stress",
                "fatigue": "Life & Damage",
            }
            rows = {
                "geom": [
                    ("d1=(q+2*x1)*mn or d1=z1*mn/tan(beta)", "mn,q,x1,z1,beta", "worm pitch diameter"),
                    ("da1=d1+2*mn; df1=d1-2.4*mn", "d1,mn", "worm tip/root diameters"),
                    ("d2=(z2+2*x2)*mn; da2=d2+2*mn*(1+x2)", "z2,x2,mn", "wheel key diameters"),
                    ("a=0.5*(d1+d2)", "d1,d2", "center distance"),
                    ("gamma=atan(z1*mn/d1); px=pi*mn; pz=z1*px", "z1,mn,d1", "lead angle/pitch/lead"),
                ],
                "stress": [
                    ("E'=2/((1-nu1^2)/E1+(1-nu2^2)/E2)", "E1,E2,nu1,nu2", "equivalent modulus"),
                    ("Ft=2*T1*1000/d1", "T1,d1", "tangential force"),
                    ("Fn=Ft/(cos(alpha_n)*sin(gamma))", "Ft,alpha_n,gamma", "normal force"),
                    ("rho_eq=0.5*d1*sin(alpha_n)*cos(gamma)", "d1,alpha_n,gamma", "equiv. curvature radius"),
                    ("p_base=0.418*sqrt(Fn*KH*E'*1000/(rho_eq*b))", "Fn,KH,E',rho_eq,b", "contact stress baseline"),
                    ("sigma_base=Fn*KF*YF/(b*mn)", "Fn,KF,YF,b,mn", "root stress baseline"),
                ],
                "fatigue": [
                    ("N_life=n1*60*life_h/i", "n1,life_h,i", "life cycles"),
                    ("SF_root=sigma_allow(N_life)/max(sigma_F)", "SN,sigma_F", "root safety factor"),
                    ("SF_contact=p_allow(N_life)/max(p)", "SN,p", "contact safety factor"),
                    ("sigma_amp=(sigma_F_max-sigma_F_min)/2", "sigma_F", "equivalent stress amplitude"),
                    ("D=n_cycles_total/N_allow(sigma_amp)", "sigma_amp,SN", "Miner cumulative damage"),
                ],
            }
            hint = "The sketch marks d1, d2, a, beta, alpha_n and the mesh contact point."

        for idx, key in enumerate(("geom", "stress", "fatigue")):
            self.formula_nb.tab(idx, text=tab_names[key])
            tree = self.formula_tabs[key]
            for iid in tree.get_children():
                tree.delete(iid)
            for r in rows[key]:
                tree.insert("", "end", values=r)

        self.formula_hint.configure(text=hint)
        self._draw_formula_schematic()

    def _draw_formula_schematic(self):
        if not hasattr(self, "ax_formula_diag"):
            return
        ax = self.ax_formula_diag
        ax.clear()
        ax.set_aspect("equal", adjustable="box")
        ax.axis("off")
        fp = {"fontfamily": _MPL_FONT or "DejaVu Sans"}

        mn = self._safe_float("mn_mm", 2.5)
        z1 = int(self._safe_float("z1", 2))
        ratio = self._safe_float("ratio", 25)
        z2 = int(round(ratio * z1))
        try:
            z2 = int(float(self.inputs["z2"].get()))
        except (ValueError, TypeError):
            pass
        x1 = self._safe_float("x1", 0.0)
        x2 = self._safe_float("x2", 0.0)
        d1 = self._safe_float("d1_mm", 0.0)
        if d1 <= 0:
            d1 = (self._safe_float("q", 10) + 2 * x1) * mn
        d2 = (z2 + 2 * x2) * mn
        beta = self._safe_float("beta_deg", 11.5)
        alpha = self._safe_float("alpha_n_deg", 20.0)
        a = 0.5 * (d1 + d2)

        s = 72.0 / max(d2, 1)
        rw = 0.5 * d1 * s
        rg = 0.5 * d2 * s
        c1 = (0.0, 0.0)
        c2 = (0.0, a * s)

        ax.add_patch(mpatches.Circle(c1, rw, fill=False, edgecolor="#A35B55", linewidth=1.8))
        ax.add_patch(mpatches.Circle(c2, rg, fill=False, edgecolor="#4F74B3", linewidth=1.8))
        ax.plot([c1[0], c2[0]], [c1[1], c2[1]], color=CLR_DIM, linewidth=1.0, linestyle=(0, (4, 3)))
        cp = (0.0, c2[1] - rg)
        ax.scatter([cp[0]], [cp[1]], s=30, color="#111111")
        ax.text(cp[0] + 2, cp[1] + 2, "P", fontsize=9, color="#111111", **fp)

        ax.annotate("", xy=(rw + 8, c1[1]), xytext=(rw + 8, c2[1]),
                    arrowprops=dict(arrowstyle="<->", color=CLR_DIM, lw=1.0))
        ax.text(rw + 10, 0.5 * c2[1], "a", fontsize=9, color=CLR_DIM, rotation=90, va="center", **fp)
        ax.annotate("", xy=(-rw, c1[1] - rw - 8), xytext=(rw, c1[1] - rw - 8),
                    arrowprops=dict(arrowstyle="<->", color="#A35B55", lw=1.0))
        ax.text(0, c1[1] - rw - 11, "d1", fontsize=9, color="#A35B55", ha="center", **fp)
        ax.annotate("", xy=(-rg, c2[1] + rg + 8), xytext=(rg, c2[1] + rg + 8),
                    arrowprops=dict(arrowstyle="<->", color="#4F74B3", lw=1.0))
        ax.text(0, c2[1] + rg + 10, "d2", fontsize=9, color="#4F74B3", ha="center", **fp)

        ang = math.radians(90 - alpha)
        ax.plot([cp[0], cp[0] + 25 * math.cos(ang)], [cp[1], cp[1] + 25 * math.sin(ang)],
                color="#6C707A", linewidth=1.0)
        ax.text(cp[0] + 12, cp[1] + 13, "alpha_n", fontsize=8, color="#6C707A", **fp)
        ax.text(c1[0] - rw - 8, c1[1] - 6, f"beta={beta:.1f} deg", fontsize=8, color="#A35B55", **fp)
        ax.text(c1[0] + rw * 0.1, c2[1] + rg * 0.15, f"z1={z1}, z2={z2}", fontsize=8, color=CLR_TEXT2, **fp)
        ax.set_xlim(-rg - 20, rg + 30)
        ax.set_ylim(-rw - 18, c2[1] + rg + 18)
        self.canvas_formula.draw()

    # ==================================================================
    # Helpers
    # ==================================================================
    def _fill_tree_kv(self, tree, rows):
        for i in tree.get_children():
            tree.delete(i)
        for k, v in rows:
            tree.insert("", "end", values=(k, v))

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
                "temp_C": float(row.get("temp_C", 80)),
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
            messagebox.showerror("Error", "S-N values must be numeric.")
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
        self._populate_steel_form()

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
        self._populate_wheel_form()

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
        # Update name/nu from form
        self.wheel["name"] = self.wheel_fields["w_name"].get()
        try:
            self.wheel["nu"] = float(self.wheel_fields["w_nu"].get() or 0.4)
        except ValueError:
            pass
        table = []
        for row in self.sn_rows:
            table.append({
                "temp_C": float(row["temp_C"]), "N": float(row["N"]),
                "contact_MPa": float(row["contact_MPa"]), "root_MPa": float(row["root_MPa"])
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
        messagebox.showinfo("OK", "Wheel material card applied (session only).")

    # ==================================================================
    # Compute
    # ==================================================================
    def _collect_inputs(self):
        d = {}
        for k, v in self.inputs.items():
            d[k] = v.get()
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
            messagebox.showerror("Error", str(e))
            import traceback
            traceback.print_exc()

    def plot_results(self, res):
        phi = res["phi"]
        deg = phi * 180 / np.pi
        for ax in [self.ax_p, self.ax_s, self.ax_t, self.ax_eta, self.ax_cloud, self.ax_legend]:
            ax.clear()

        fp = {"fontfamily": _MPL_FONT or "DejaVu Sans"}

        def _style_ax(ax, title, xlabel, ylabel):
            ax.set_title(title, fontsize=10, fontweight="bold", color=CLR_TEXT, **fp)
            ax.set_xlabel(xlabel, fontsize=9, color=CLR_TEXT2)
            ax.set_ylabel(ylabel, fontsize=9, color=CLR_TEXT2)
            ax.tick_params(labelsize=8)
            ax.grid(True, alpha=0.2)
            ax.set_facecolor("#FAFBFC")

        self.ax_p.fill_between(deg, res["p_contact_MPa"], alpha=0.15, color=CLR_ACCENT)
        self.ax_p.plot(deg, res["p_contact_MPa"], color=CLR_ACCENT, linewidth=1.5)
        _style_ax(self.ax_p, "Contact Stress p(phi)", "phi (deg)", "MPa")

        self.ax_s.fill_between(deg, res["sigma_root_MPa"], alpha=0.15, color=CLR_WORM)
        self.ax_s.plot(deg, res["sigma_root_MPa"], color=CLR_WORM, linewidth=1.5)
        _style_ax(self.ax_s, "Root Stress sF(phi)", "phi (deg)", "MPa")

        self.ax_t.fill_between(deg, res["T2_Nm"], alpha=0.15, color=CLR_ACCENT2)
        self.ax_t.plot(deg, res["T2_Nm"], color=CLR_ACCENT2, linewidth=1.5)
        _style_ax(self.ax_t, "Output Torque T2(phi)", "phi (deg)", "N*m")

        self.ax_eta.plot(deg, res["eta"], color=CLR_ACCENT, linewidth=1.5, label="eta")
        self.ax_eta.plot(deg, res["Nc_proxy"], color="#FF9500", linewidth=1.5, label="Nc")
        _style_ax(self.ax_eta, "Efficiency & Contact No.", "phi (deg)", "-")
        self.ax_eta.legend(fontsize=8, framealpha=0.9)

        # 3D cloud
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
        self.ax_cloud.set_title("3D Root Stress", fontsize=10, fontweight="bold", color=CLR_TEXT)
        self.ax_cloud.set_xlabel("phi (deg)", fontsize=8)
        self.ax_cloud.set_ylabel("Width", fontsize=8)
        self.ax_cloud.set_zlabel("sF MPa", fontsize=8)
        self.ax_cloud.view_init(elev=24, azim=-130)

        sm = matplotlib.cm.ScalarMappable(cmap="inferno", norm=norm)
        sm.set_array([])
        if self._cloud_cbar is not None:
            try:
                self._cloud_cbar.remove()
            except Exception:
                pass
        self._cloud_cbar = self.canvas_res.figure.colorbar(sm, ax=self.ax_cloud, shrink=0.60, pad=0.08)
        self._cloud_cbar.set_label("Stress MPa", fontsize=8)

        self.ax_legend.axis("off")
        self.ax_legend.set_facecolor(CLR_CARD)
        m = res["meta"]
        info = (
            f"Summary\n"
            f"-------------------\n"
            f"z1={m['z1']}  z2={m['z2']}\n"
            f"d1={m['d1_mm']:.2f} mm\n"
            f"d2={m['d2_mm']:.2f} mm\n"
            f"a ={m['a_mm']:.2f} mm\n"
            f"eta0={m['eta0']:.3f}\n"
            f"beta={m.get('beta_deg', m['gamma_deg']):.2f} deg\n"
            f"gamma={m['gamma_deg']:.2f} deg\n"
            f"px={m['px_mm']:.2f} mm\n"
            f"pz={m['pz_mm']:.2f} mm\n"
            f"-------------------\n"
            f"Contact peak: {float(np.max(res['p_contact_MPa'])):.1f} MPa\n"
            f"Root peak: {float(np.max(res['sigma_root_MPa'])):.1f} MPa"
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
        worm_rows = [
            ("z1", f"{m['z1']}"),
            ("d1", f"{m['d1_mm']:.2f} mm"),
            ("da1", f"{m['da1_mm']:.2f} mm"),
            ("df1", f"{m['df1_mm']:.2f} mm"),
            ("x1", f"{m['x1']:.3f}"),
            ("q", f"{m.get('q', 0.0):.3f}"),
            ("beta", f"{m.get('beta_deg', m['gamma_deg']):.2f} deg"),
            ("gamma", f"{m['gamma_deg']:.2f} deg"),
            ("px", f"{m['px_mm']:.2f} mm"),
            ("pz (lead)", f"{m['pz_mm']:.2f} mm"),
            ("L (worm)", f"{m['L_worm_mm']:.2f} mm"),
            ("eta0", f"{m['eta0']:.3f}"),
            ("Root stress peak", f"{np.max(res['sigma_root_MPa']):.2f} MPa"),
            ("SF_root", "-" if m.get("SF_root") is None else f"{m['SF_root']:.2f}"),
        ]
        wheel_rows = [
            ("z2", f"{m['z2']}"),
            ("d2", f"{m['d2_mm']:.2f} mm"),
            ("da2", f"{m['da2_mm']:.2f} mm"),
            ("df2", f"{m['df2_mm']:.2f} mm"),
            ("x2", f"{m['x2']:.3f}"),
            ("a", f"{m['a_mm']:.2f} mm"),
            ("Contact stress peak", f"{np.max(res['p_contact_MPa']):.2f} MPa"),
            ("SF_contact", "-" if m.get("SF_contact") is None else f"{m['SF_contact']:.2f}"),
            ("delta_a", "-" if m.get("delta_a_mm") is None else f"{m['delta_a_mm']:.2f} mm"),
        ]
        self._fill_tree_kv(self.worm_out, worm_rows)
        self._fill_tree_kv(self.wheel_out, wheel_rows)

        lines = []
        lines.append("Rainflow + Miner Cumulative Damage (root stress proxy)")
        lines.append(f"  D = {m['damage_root']:.3e} (D < 1 = OK)")
        if m.get("SF_root") is not None:
            lines.append(f"  SF_root = {m['SF_root']:.2f}")
        else:
            lines.append("  SF_root: no root SN data")
        if m.get("SF_contact") is not None:
            lines.append(f"  SF_contact = {m['SF_contact']:.2f}")
        else:
            lines.append("  SF_contact: no contact SN data")
        lines.append("")
        lines.append("Key Parameters:")
        lines.append(f"  z1={m['z1']}, z2={m['z2']}")
        lines.append(f"  d1={m['d1_mm']:.2f}, d2={m['d2_mm']:.2f}, a={m['a_mm']:.2f} mm")
        if m.get("a_target_mm") is not None:
            lines.append(f"  a_target={m['a_target_mm']:.2f}, da={m['delta_a_mm']:.2f} mm")
        lines.append(f"  x1={m['x1']:.3f}, x2={m['x2']:.3f}")
        lines.append(f"  q={m.get('q', 0.0):.3f}, beta={m.get('beta_deg', m['gamma_deg']):.2f} deg")
        lines.append(f"  gamma={m['gamma_deg']:.2f} deg, px={m['px_mm']:.2f} mm")
        lines.append(f"  E'={m['Eprime_GPa']:.2f} GPa, eta0={m['eta0']:.3f}")
        lines.append(f"  KA={m['KA']:.3f} KV={m['KV']:.3f} KHb={m['KHb']:.3f} KFb={m['KFb']:.3f}")
        self.fat_text.delete("1.0", "end")
        self.fat_text.insert("1.0", "\n".join(lines))

    # ==================================================================
    # Export
    # ==================================================================
    def export_xlsx(self):
        if self.res is None:
            messagebox.showwarning("Info", "Please calculate first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        inputs = self._collect_inputs()
        export_cycle_xlsx(path, inputs, self.steel, self.wheel, self.res)
        messagebox.showinfo("OK", f"Exported:\n{path}")


if __name__ == "__main__":
    App().mainloop()
