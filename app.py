
import os, json
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.patches
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import font_manager
from matplotlib.colors import Normalize
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

from src.utils import load_json
from src.worm_model import compute_worm_cycle
from src.export_xlsx import export_cycle_xlsx

APP_TITLE = "Light Worm Gear Tool v3 — 蜗轮蜗杆详细版（可扩展）"

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
            default_font = tkfont.nametofont("TkDefaultFont")
            default_font.configure(family=chosen, size=10)
            text_font = tkfont.nametofont("TkTextFont")
            text_font.configure(family=chosen, size=10)
            menu_font = tkfont.nametofont("TkMenuFont")
            menu_font.configure(family=chosen, size=10)
        except Exception:
            pass

    matplotlib.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False

def list_materials(folder):
    out=[]
    if not os.path.isdir(folder):
        return out
    for fn in os.listdir(folder):
        if fn.lower().endswith(".json"):
            out.append(os.path.join(folder, fn))
    return sorted(out)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        setup_fonts(self)
        self._setup_style()
        self.title(APP_TITLE)
        self.geometry("1420x860")

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.steel_paths = list_materials(os.path.join(self.base_dir, "materials", "metals"))
        self.poly_paths  = list_materials(os.path.join(self.base_dir, "materials", "polymers"))
        if not self.steel_paths or not self.poly_paths:
            raise RuntimeError("materials 目录缺失，请保持项目完整解压。")

        self.steel_db = {os.path.basename(p): p for p in self.steel_paths}
        self.poly_db  = {os.path.basename(p): p for p in self.poly_paths}

        self.steel = load_json(self.steel_db.get("37CrS4.json", self.steel_paths[0]))
        self.wheel = load_json(self.poly_db.get("PA66_modified_draft.json", self.poly_paths[0]))

        self.inputs = {
            "T1_Nm":"6.0",
            "n1_rpm":"3000",
            "ratio":"25",
            "z1":"2",
            "mn_mm":"2.5",
            "q":"10",
            "x1":"0.0",
            "x2":"0.0",
            "a_target_mm":"",
            "b_mm":"18",
            "alpha_n_deg":"20",
            "cross_angle_deg":"90",
            "rho_f_mm":"0.6",
            "mu":"0.06",
            "KA":"1.10",
            "KV":"1.05",
            "KHb":"1.00",
            "KFb":"1.00",
            "temp_C":"80",
            "life_h":"3000",
            "steps":"720"
        }
        self.res = None
        self.sn_rows = []
        self._cloud_cbar = None

        self._build_menu()
        self._build_ui()
        self.refresh_geom_plot()

    def _setup_style(self):
        self.configure(bg="#eef2f7")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#eef2f7")
        style.configure("TLabelframe", background="#eef2f7", borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background="#eef2f7", foreground="#1f2d3d")
        style.configure("TLabel", background="#eef2f7", foreground="#1f2d3d")
        style.configure("TButton", padding=(10, 6))
        style.configure("TNotebook", background="#eef2f7", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 8))
        style.configure("Treeview", rowheight=24)
        style.map("TButton", background=[("active", "#d7e3f4")])

    def _build_menu(self):
        m = tk.Menu(self)
        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="导出 XLSX（曲线）...", command=self.export_xlsx)
        fm.add_separator()
        fm.add_command(label="退出", command=self.destroy)
        m.add_cascade(label="文件", menu=fm)
        self.config(menu=m)

    def _build_ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.tab_geom = ttk.Frame(self.nb)
        self.tab_mat  = ttk.Frame(self.nb)
        self.tab_res  = ttk.Frame(self.nb)
        self.tab_fat  = ttk.Frame(self.nb)

        self.nb.add(self.tab_geom, text="1) 参数输入")
        self.nb.add(self.tab_mat,  text="2) 材料与S-N")
        self.nb.add(self.tab_res,  text="3) 应力与效率结果")
        self.nb.add(self.tab_fat,  text="4) 寿命校核汇总")

        self._build_geom_tab()
        self._build_mat_tab()
        self._build_res_tab()
        self._build_fat_tab()

    def _entry(self, parent, key, label, unit=""):
        row = ttk.Frame(parent); row.pack(fill="x", padx=10, pady=4)
        ttk.Label(row, text=label, width=24).pack(side="left")
        var = tk.StringVar(value=self.inputs.get(key,""))
        ttk.Entry(row, textvariable=var, width=18).pack(side="left")
        if unit:
            ttk.Label(row, text=unit, width=10).pack(side="left", padx=(8,0))
        self.inputs[key] = var

    def _build_geom_tab(self):
        left = ttk.Frame(self.tab_geom); left.pack(side="left", fill="y", padx=8, pady=8)
        right = ttk.Frame(self.tab_geom); right.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        left_card = ttk.Labelframe(left, text="参数输入（滚动）")
        left_card.pack(fill="y", expand=True, padx=6, pady=6)
        self.geom_canvas = tk.Canvas(left_card, width=430, highlightthickness=0, bg="#eef2f7")
        self.geom_canvas.pack(side="left", fill="y", expand=True)
        geom_scroll = ttk.Scrollbar(left_card, orient="vertical", command=self.geom_canvas.yview)
        geom_scroll.pack(side="left", fill="y")
        self.geom_canvas.configure(yscrollcommand=geom_scroll.set)
        geom_inner = ttk.Frame(self.geom_canvas)
        self._geom_window = self.geom_canvas.create_window((0, 0), window=geom_inner, anchor="nw")

        def _on_inner_configure(event):
            self.geom_canvas.configure(scrollregion=self.geom_canvas.bbox("all"))

        def _on_canvas_configure(event):
            self.geom_canvas.itemconfigure(self._geom_window, width=event.width)

        geom_inner.bind("<Configure>", _on_inner_configure)
        self.geom_canvas.bind("<Configure>", _on_canvas_configure)
        self.geom_canvas.bind_all("<MouseWheel>", self._on_mousewheel_geom)

        grp = ttk.Labelframe(geom_inner, text="输入参数（带单位）"); grp.pack(fill="x", padx=6, pady=6)
        self._entry(grp, "T1_Nm", "输入扭矩 T1", "N·m")
        self._entry(grp, "n1_rpm", "蜗杆转速 n1", "rpm")
        self._entry(grp, "ratio", "传动比 i", "-")
        self._entry(grp, "z1", "蜗杆头数 z1", "-")
        self._entry(grp, "mn_mm", "法向模数 mn", "mm")
        self._entry(grp, "q", "蜗杆直径系数 q", "-")
        self._entry(grp, "x1", "蜗杆变位系数 x1", "-")
        self._entry(grp, "x2", "蜗轮变位系数 x2", "-")
        self._entry(grp, "a_target_mm", "目标中心距 a(可空)", "mm")
        self._entry(grp, "b_mm", "齿宽 b", "mm")
        self._entry(grp, "alpha_n_deg", "法向压力角 αn", "deg")
        self._entry(grp, "cross_angle_deg", "交错角 ε", "deg")
        self._entry(grp, "rho_f_mm", "齿根圆角 ρf(等效)", "mm")
        self._entry(grp, "mu", "摩擦系数 μ(估)", "-")
        self._entry(grp, "KA", "使用系数 KA", "-")
        self._entry(grp, "KV", "动载系数 KV", "-")
        self._entry(grp, "KHb", "齿宽载荷系数 KHβ", "-")
        self._entry(grp, "KFb", "齿根载荷系数 KFβ", "-")
        self._entry(grp, "temp_C", "工作温度", "°C")
        self._entry(grp, "life_h", "目标寿命", "h")
        self._entry(grp, "steps", "相位点数", "点")

        btns = ttk.Frame(geom_inner); btns.pack(fill="x", padx=6, pady=6)
        ttk.Button(btns, text="自动生成中心距/变位", command=self.autofill_center_shift).pack(side="left")
        ttk.Button(btns, text="更新示意图", command=self.refresh_geom_plot).pack(side="left")
        ttk.Button(btns, text="计算并绘图", command=self.run).pack(side="left", padx=8)

        fig = Figure(figsize=(7.2, 5.8), dpi=100)
        self.ax_geom = fig.add_subplot(111)
        self.canvas_geom = FigureCanvasTkAgg(fig, master=right)
        self.canvas_geom.get_tk_widget().pack(fill="both", expand=True)
        self.geom_check_var = tk.StringVar(value="几何校核信息：等待输入。")
        ttk.Label(right, textvariable=self.geom_check_var, foreground="#2b4b6f").pack(anchor="w", padx=6, pady=(4,0))
        ttk.Label(right, text="示意图用于帮助新手对应参数（非真实齿形比例）。", foreground="#555").pack(anchor="w", padx=6, pady=(6,0))

    def _on_mousewheel_geom(self, event):
        if self.nb.index(self.nb.select()) != self.nb.index(self.tab_geom):
            return
        self.geom_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_mat_tab(self):
        top = ttk.Frame(self.tab_mat); top.pack(fill="both", expand=True, padx=10, pady=10)

        g1 = ttk.Labelframe(top, text="蜗杆材料（默认 37CrS4，可导入 JSON）")
        g1.pack(fill="x", pady=6)
        row = ttk.Frame(g1); row.pack(fill="x", padx=10, pady=8)
        ttk.Label(row, text="选择材料").pack(side="left")
        self.steel_var = tk.StringVar(value="37CrS4.json" if "37CrS4.json" in self.steel_db else list(self.steel_db.keys())[0])
        self.steel_cb = ttk.Combobox(row, textvariable=self.steel_var, values=list(self.steel_db.keys()), state="readonly", width=40)
        self.steel_cb.pack(side="left", padx=8)
        ttk.Button(row, text="加载", command=self.load_steel).pack(side="left")
        ttk.Button(row, text="导入材料JSON...", command=self.import_steel).pack(side="left", padx=8)
        self.steel_text = tk.Text(g1, height=6, wrap="word")
        self.steel_text.pack(fill="x", padx=10, pady=(0,10))

        g2 = ttk.Labelframe(top, text="蜗轮材料（特殊材料：PA66改性，支持输入/导入SN）")
        g2.pack(fill="both", expand=True, pady=6)
        row2 = ttk.Frame(g2); row2.pack(fill="x", padx=10, pady=8)
        ttk.Label(row2, text="基础模板").pack(side="left")
        self.wheel_var = tk.StringVar(value="PA66_modified_draft.json" if "PA66_modified_draft.json" in self.poly_db else list(self.poly_db.keys())[0])
        self.wheel_cb = ttk.Combobox(row2, textvariable=self.wheel_var, values=list(self.poly_db.keys()), state="readonly", width=40)
        self.wheel_cb.pack(side="left", padx=8)
        ttk.Button(row2, text="加载", command=self.load_wheel).pack(side="left")
        ttk.Button(row2, text="导入材料JSON...", command=self.import_wheel).pack(side="left", padx=8)

        row3 = ttk.Frame(g2); row3.pack(fill="x", padx=10, pady=6)
        ttk.Label(row3, text="E(T) 点 (°C:GPa) 例 23:3.0,60:2.4,80:2.0").pack(side="left")
        self.Et_var = tk.StringVar(value=self._format_Et())
        ttk.Entry(row3, textvariable=self.Et_var, width=70).pack(side="left", padx=8)

        row4 = ttk.Frame(g2); row4.pack(fill="both", padx=10, pady=6, expand=True)
        ttk.Label(row4, text="S-N 输入表（温度、循环次数、接触应力、齿根应力）").pack(anchor="w")

        cols = ("temp_C", "N", "contact_MPa", "root_MPa")
        self.sn_table = ttk.Treeview(row4, columns=cols, show="headings", height=8)
        self.sn_table.heading("temp_C", text="温度 °C")
        self.sn_table.heading("N", text="循环次数 N")
        self.sn_table.heading("contact_MPa", text="接触许用 MPa")
        self.sn_table.heading("root_MPa", text="齿根许用 MPa")
        self.sn_table.column("temp_C", width=120, anchor="center")
        self.sn_table.column("N", width=180, anchor="center")
        self.sn_table.column("contact_MPa", width=180, anchor="center")
        self.sn_table.column("root_MPa", width=180, anchor="center")
        self.sn_table.pack(side="left", fill="both", expand=True)

        table_scroll = ttk.Scrollbar(row4, orient="vertical", command=self.sn_table.yview)
        self.sn_table.configure(yscrollcommand=table_scroll.set)
        table_scroll.pack(side="left", fill="y")

        ctrl = ttk.Frame(g2); ctrl.pack(fill="x", padx=10, pady=6)
        self.sn_temp_var = tk.StringVar(value=self.inputs["temp_C"].get())
        self.sn_N_var = tk.StringVar(value="1e6")
        self.sn_contact_var = tk.StringVar(value="110")
        self.sn_root_var = tk.StringVar(value="45")
        ttk.Entry(ctrl, textvariable=self.sn_temp_var, width=10).pack(side="left")
        ttk.Entry(ctrl, textvariable=self.sn_N_var, width=14).pack(side="left", padx=(6, 0))
        ttk.Entry(ctrl, textvariable=self.sn_contact_var, width=14).pack(side="left", padx=(6, 0))
        ttk.Entry(ctrl, textvariable=self.sn_root_var, width=14).pack(side="left", padx=(6, 0))
        ttk.Button(ctrl, text="添加行", command=self.add_sn_row).pack(side="left", padx=8)
        ttk.Button(ctrl, text="删除选中行", command=self.delete_sn_rows).pack(side="left")
        ttk.Button(ctrl, text="应用蜗轮材料卡到当前模型", command=self.apply_wheel_card).pack(side="left", padx=8)

        self.wheel_text = tk.Text(g2, height=10, wrap="word")
        self.wheel_text.pack(fill="both", expand=True, padx=10, pady=(0,10))

        self._load_sn_table_from_wheel()
        self._refresh_material_texts()

    def _build_res_tab(self):
        top = ttk.Frame(self.tab_res); top.pack(fill="both", expand=True, padx=8, pady=8)
        fig = Figure(figsize=(9.5, 6.6), dpi=100)
        self.ax_p = fig.add_subplot(231)
        self.ax_s = fig.add_subplot(232)
        self.ax_t = fig.add_subplot(233)
        self.ax_eta = fig.add_subplot(234)
        self.ax_cloud = fig.add_subplot(235, projection="3d")
        self.ax_legend = fig.add_subplot(236)
        self.ax_legend.axis("off")
        self.canvas_res = FigureCanvasTkAgg(fig, master=top)
        self.canvas_res.get_tk_widget().pack(fill="both", expand=True)
        ttk.Label(top, text="当前为轻量代理模型曲线（含KISSsoft风格修正系数与3D应力云图代理）。", foreground="#555").pack(anchor="w", padx=6, pady=6)

    def _build_fat_tab(self):
        top = ttk.Frame(self.tab_fat)
        top.pack(fill="both", expand=True, padx=10, pady=10)

        cards = ttk.Frame(top)
        cards.pack(fill="x")
        worm_box = ttk.Labelframe(cards, text="蜗杆输出参数")
        wheel_box = ttk.Labelframe(cards, text="蜗轮输出参数")
        worm_box.pack(side="left", fill="both", expand=True, padx=(0, 6))
        wheel_box.pack(side="left", fill="both", expand=True, padx=(6, 0))

        self.worm_out = ttk.Treeview(worm_box, columns=("k", "v"), show="headings", height=7)
        self.worm_out.heading("k", text="参数")
        self.worm_out.heading("v", text="值")
        self.worm_out.column("k", width=180, anchor="w")
        self.worm_out.column("v", width=220, anchor="w")
        self.worm_out.pack(fill="both", expand=True, padx=8, pady=8)

        self.wheel_out = ttk.Treeview(wheel_box, columns=("k", "v"), show="headings", height=7)
        self.wheel_out.heading("k", text="参数")
        self.wheel_out.heading("v", text="值")
        self.wheel_out.column("k", width=180, anchor="w")
        self.wheel_out.column("v", width=220, anchor="w")
        self.wheel_out.pack(fill="both", expand=True, padx=8, pady=8)

        self.fat_text = tk.Text(top, wrap="word", bd=0, relief="flat", bg="#f8fafc")
        self.fat_text.pack(fill="both", expand=True, padx=2, pady=(10, 0))
        self.fat_text.insert("1.0", "点击“计算并绘图”后，这里会显示雨流计数的疲劳损伤与安全系数汇总。")

    def _fill_tree_kv(self, tree, rows):
        for i in tree.get_children():
            tree.delete(i)
        for k, v in rows:
            tree.insert("", "end", values=(k, v))

    # ---- helpers ----
    def _format_Et(self):
        pts = self.wheel.get("elastic_T", {}).get("points_C_GPa", [])
        return ",".join([f"{int(p[0])}:{p[1]}" for p in pts])

    def _build_sn_rows_from_legacy(self):
        temp = float(self.inputs["temp_C"].get() or 80.0)
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
                "temp_C": float(row.get("temp_C", self.inputs["temp_C"].get() or 80)),
                "N": float(row.get("N", 1e6)),
                "contact_MPa": float(row.get("contact_MPa", 0)),
                "root_MPa": float(row.get("root_MPa", 0))
            }
            self.sn_rows.append(one)
            self.sn_table.insert(
                "",
                "end",
                values=(f"{one['temp_C']:.1f}", f"{one['N']:.3e}", f"{one['contact_MPa']:.3f}", f"{one['root_MPa']:.3f}")
            )

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
        self.sn_table.insert(
            "",
            "end",
            values=(f"{row['temp_C']:.1f}", f"{row['N']:.3e}", f"{row['contact_MPa']:.3f}", f"{row['root_MPa']:.3f}")
        )

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
        path = filedialog.askopenfilename(filetypes=[("JSON","*.json")])
        if not path: return
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
        path = filedialog.askopenfilename(filetypes=[("JSON","*.json")])
        if not path: return
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
        pts_et=[]
        for part in self.Et_var.get().split(","):
            part=part.strip()
            if not part or ":" not in part:
                continue
            a,b = part.split(":",1)
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
        # 兼容旧模型字段：默认抽取当前温度附近的一组
        t_ref = float(self.inputs["temp_C"].get() or 80.0)
        near = sorted(table, key=lambda x: abs(x["temp_C"] - t_ref))
        if near:
            t_pick = near[0]["temp_C"]
            same_t = [r for r in table if abs(r["temp_C"] - t_pick) < 1e-9]
            self.wheel["SN"]["contact_allow_MPa_vs_N"] = [[r["N"], r["contact_MPa"]] for r in same_t if r["contact_MPa"] > 0]
            self.wheel["SN"]["root_allow_MPa_vs_N"] = [[r["N"], r["root_MPa"]] for r in same_t if r["root_MPa"] > 0]
        self._refresh_material_texts()
        messagebox.showinfo("完成", "蜗轮材料卡已应用（仅当前会话）。")

    def _refresh_material_texts(self):
        self.steel_text.delete("1.0","end")
        self.steel_text.insert("1.0", json.dumps(self.steel, ensure_ascii=False, indent=2))
        self.wheel_text.delete("1.0","end")
        self.wheel_text.insert("1.0", json.dumps(self.wheel, ensure_ascii=False, indent=2))

    def refresh_geom_plot(self):
        ax = self.ax_geom
        ax.clear()
        ax.set_aspect("equal", adjustable="box")
        ax.axis("off")

        mn = float(self.inputs["mn_mm"].get() or 2.5)
        q  = float(self.inputs["q"].get() or 10)
        ratio = float(self.inputs["ratio"].get() or 25)
        z1 = int(float(self.inputs["z1"].get() or 2))
        z2 = int(round(ratio*z1))
        x1 = float(self.inputs["x1"].get() or 0.0)
        x2 = float(self.inputs["x2"].get() or 0.0)
        d1 = (q + 2.0*x1)*mn
        d2 = (z2 + 2.0*x2)*mn
        a = 0.5*(d1+d2)
        a_target_txt = self.inputs["a_target_mm"].get().strip()
        a_target = float(a_target_txt) if a_target_txt else None
        xsum_need = None
        if a_target is not None:
            xsum_need = a_target/mn - 0.5*(q + z2)
            delta = a_target - a
            self.geom_check_var.set(
                f"几何校核：a_calc={a:.3f} mm, a_target={a_target:.3f} mm, Δa={delta:.3f} mm, 建议 x1+x2≈{xsum_need:.3f}"
            )
        else:
            self.geom_check_var.set(f"几何校核：a_calc={a:.3f} mm，未设目标中心距。")
        b = float(self.inputs["b_mm"].get() or 18)
        eps = float(self.inputs["cross_angle_deg"].get() or 90)

        R2 = 1.0
        R1 = (d1/d2)*R2 if d2>0 else 0.3
        cx2, cy2 = 0.0, 0.0
        cx1, cy1 = - (a/d2)*2.2, 0.0

        ax.add_patch(matplotlib.patches.Circle((cx2, cy2), R2, fill=False, lw=2))
        ax.add_patch(matplotlib.patches.Circle((cx1, cy1), R1, fill=False, lw=2))

        ax.text(cx2, cy2+R2+0.12, "蜗轮 (d2)", ha="center")
        ax.text(cx1, cy1+R1+0.12, "蜗杆 (d1)", ha="center")

        ax.annotate("", xy=(cx1, cy1-1.25), xytext=(cx2, cy2-1.25),
                    arrowprops=dict(arrowstyle="<->", lw=1.8))
        ax.text((cx1+cx2)/2, cy2-1.42, f"中心距 a≈{a:.1f} mm", ha="center")

        ax.text(cx2+R2+0.25, cy2, f"mn={mn:.2f}mm\nz1={z1}, z2≈{z2}\nx1={x1:.3f}, x2={x2:.3f}", va="center")
        ax.annotate("", xy=(cx2+0.2, cy2+R2), xytext=(cx2+0.2, cy2+R2+0.5),
                    arrowprops=dict(arrowstyle="<->", lw=1.5))
        ax.text(cx2+0.28, cy2+R2+0.25, f"齿宽 b={b:.1f}mm", va="center")
        ax.text(cx1, cy1- R1-0.35, f"交错角 ε={eps:.1f}°", ha="center")

        ax.plot([cx1-1.3, cx2+1.3], [0,0], lw=1, alpha=0.5)

        # 右下角：齿型啮合截面参考图（虚线=无变位，实线=有变位）
        inset = ax.inset_axes([0.52, 0.05, 0.44, 0.42])
        inset.set_title("啮合截面细节（参考）", fontsize=9)
        inset.set_facecolor("#fafbfd")

        x = np.linspace(-1.2, 1.2, 700)
        pitch = 0.35
        saw = ((x / pitch) % 1.0)
        tri = np.where(saw < 0.5, 2.0 * saw, 2.0 * (1.0 - saw))

        # 无变位基准（虚线）
        wheel_base = -0.22 + 0.16 * tri
        worm_base = 0.22 - 0.16 * np.roll(tri, 45)
        inset.plot(x, wheel_base, "--", color="#9aa6b2", lw=1.2, label="蜗轮 无变位")
        inset.plot(x, worm_base, "--", color="#9aa6b2", lw=1.2, label="蜗杆 无变位")

        # 有变位（实线）：用 x1/x2 改变齿厚与相位
        dx1 = 0.08 * x1
        dx2 = 0.08 * x2
        tri_wheel = np.where((((x + dx2) / pitch) % 1.0) < 0.5, 2.0 * (((x + dx2) / pitch) % 1.0), 2.0 * (1.0 - (((x + dx2) / pitch) % 1.0)))
        tri_worm = np.where((((x - dx1) / pitch) % 1.0) < 0.5, 2.0 * (((x - dx1) / pitch) % 1.0), 2.0 * (1.0 - (((x - dx1) / pitch) % 1.0)))
        wheel_shift = -0.22 + (0.16 + 0.03 * x2) * tri_wheel
        worm_shift = 0.22 - (0.16 + 0.03 * x1) * tri_worm

        inset.plot(x, wheel_shift, "-", color="#2463a6", lw=1.8, label="蜗轮 有变位")
        inset.plot(x, worm_shift, "-", color="#bd3b1b", lw=1.8, label="蜗杆 有变位")

        gap = worm_shift - wheel_shift
        min_idx = int(np.argmin(gap))
        inset.scatter([x[min_idx]], [(wheel_shift[min_idx] + worm_shift[min_idx]) * 0.5], s=22, color="#2a9d8f", zorder=5)
        inset.text(x[min_idx] + 0.04, (wheel_shift[min_idx] + worm_shift[min_idx]) * 0.5 + 0.02, "近似啮合点", fontsize=8)
        inset.text(-1.15, 0.31, f"x1={x1:.3f}, x2={x2:.3f}", fontsize=8, color="#334155")
        inset.set_xlim(-1.2, 1.2)
        inset.set_ylim(-0.38, 0.38)
        inset.set_xticks([])
        inset.set_yticks([])
        inset.grid(alpha=0.25, linestyle=":")
        inset.legend(loc="lower right", fontsize=7, frameon=False)

        ax.set_xlim(-3.0, 2.5)
        ax.set_ylim(-2.0, 2.2)
        self.canvas_geom.draw()

    def autofill_center_shift(self):
        try:
            mn = float(self.inputs["mn_mm"].get() or 2.5)
            q = float(self.inputs["q"].get() or 10)
            ratio = float(self.inputs["ratio"].get() or 25)
            z1 = int(float(self.inputs["z1"].get() or 2))
            z2 = int(round(ratio * z1))
            x1 = float(self.inputs["x1"].get() or 0.0)
            x2 = float(self.inputs["x2"].get() or 0.0)
            a_target_txt = self.inputs["a_target_mm"].get().strip()
            if not a_target_txt:
                a_calc = 0.5 * mn * (q + z2 + 2 * (x1 + x2))
                self.inputs["a_target_mm"].set(f"{a_calc:.3f}")
            else:
                a_target = float(a_target_txt)
                xsum_need = a_target / mn - 0.5 * (q + z2)
                self.inputs["x2"].set(f"{xsum_need - x1:.4f}")
            self.refresh_geom_plot()
        except Exception as e:
            messagebox.showerror("自动生成失败", str(e))

    def run(self):
        try:
            inp = {k:v.get() for k,v in self.inputs.items()}
            res = compute_worm_cycle(inp, self.steel, self.wheel)
            self.res = res
            self.plot_results(res)
            self.update_fatigue(res)
            self.nb.select(self.tab_res)
        except Exception as e:
            messagebox.showerror("计算失败", str(e))

    def plot_results(self, res):
        phi = res["phi"]
        deg = phi*180/np.pi
        self.ax_p.clear(); self.ax_s.clear(); self.ax_t.clear(); self.ax_eta.clear(); self.ax_cloud.clear(); self.ax_legend.clear()

        self.ax_p.plot(deg, res["p_contact_MPa"])
        self.ax_p.set_title("接触应力（代理）p(φ)")
        self.ax_p.set_xlabel("φ (deg)"); self.ax_p.set_ylabel("MPa")

        self.ax_s.plot(deg, res["sigma_root_MPa"])
        self.ax_s.set_title("齿根应力（代理）σF(φ)")
        self.ax_s.set_xlabel("φ (deg)"); self.ax_s.set_ylabel("MPa")

        self.ax_t.plot(deg, res["T2_Nm"])
        self.ax_t.set_title("输出扭矩波动 T2(φ)")
        self.ax_t.set_xlabel("φ (deg)"); self.ax_t.set_ylabel("N·m")

        self.ax_eta.plot(deg, res["eta"])
        self.ax_eta.plot(deg, res["Nc_proxy"])
        self.ax_eta.set_title("效率 η(φ) 与接触数代理 Nc(φ)")
        self.ax_eta.set_xlabel("φ (deg)"); self.ax_eta.set_ylabel("-")
        self.ax_eta.legend(["η","Nc"], loc="best", fontsize=8)

        # 3D 应力云图（代理）：环向位置×齿宽位置的应力分布
        width_pts = np.linspace(-1.0, 1.0, 36)
        PHI, BW = np.meshgrid(phi, width_pts)
        base = np.interp(PHI[0], phi, res["sigma_root_MPa"])
        base2d = np.tile(base, (len(width_pts), 1))
        spread = 1.0 + 0.22 * (BW**2) + 0.10 * np.sin(2.0 * PHI)
        sigma_3d = base2d * spread
        norm = Normalize(vmin=float(np.min(sigma_3d)), vmax=float(np.max(sigma_3d)))
        face_colors = matplotlib.cm.inferno(norm(sigma_3d))
        self.ax_cloud.plot_surface(
            PHI * 180.0 / np.pi,
            BW,
            sigma_3d,
            rstride=1,
            cstride=1,
            linewidth=0,
            antialiased=True,
            facecolors=face_colors,
            shade=False
        )
        self.ax_cloud.set_title("3D齿根应力云图（代理）")
        self.ax_cloud.set_xlabel("相位 φ (deg)")
        self.ax_cloud.set_ylabel("齿宽归一化")
        self.ax_cloud.set_zlabel("σF (MPa)")
        self.ax_cloud.view_init(elev=24, azim=-130)

        sm = matplotlib.cm.ScalarMappable(cmap="inferno", norm=norm)
        sm.set_array([])
        if self._cloud_cbar is not None:
            try:
                self._cloud_cbar.remove()
            except Exception:
                pass
        self._cloud_cbar = self.canvas_res.figure.colorbar(sm, ax=self.ax_cloud, shrink=0.60, pad=0.08)
        self._cloud_cbar.set_label("应力 MPa")
        self.ax_legend.axis("off")
        self.ax_legend.text(
            0.02,
            0.88,
            "说明：\n- 3D云图用于快速识别\n  相位+齿宽方向热点\n- 当前为轻量代理分布",
            va="top"
        )

        self.canvas_res.figure.tight_layout()
        self.canvas_res.draw()

    def update_fatigue(self, res):
        m = res["meta"]
        z1 = int(float(self.inputs["z1"].get() or 2))
        worm_rows = [
            ("头数 z1", f"{z1:d}"),
            ("分度圆 d1", f"{m['d1_mm']:.2f} mm"),
            ("变位系数 x1", f"{m['x1']:.3f}"),
            ("导程角 γ", f"{m['gamma_deg']:.2f} deg"),
            ("效率 η0", f"{m['eta0']:.3f}"),
            ("齿根应力峰值", f"{np.max(res['sigma_root_MPa']):.2f} MPa"),
            ("齿根安全系数", "-" if m.get("SF_root") is None else f"{m['SF_root']:.2f}")
        ]
        wheel_rows = [
            ("齿数 z2", f"{m['z2']:d}"),
            ("分度圆 d2", f"{m['d2_mm']:.2f} mm"),
            ("变位系数 x2", f"{m['x2']:.3f}"),
            ("中心距 a", f"{m['a_mm']:.2f} mm"),
            ("接触应力峰值", f"{np.max(res['p_contact_MPa']):.2f} MPa"),
            ("齿面安全系数", "-" if m.get("SF_contact") is None else f"{m['SF_contact']:.2f}"),
            ("中心距偏差 Δa", "-" if m.get("delta_a_mm") is None else f"{m['delta_a_mm']:.2f} mm")
        ]
        self._fill_tree_kv(self.worm_out, worm_rows)
        self._fill_tree_kv(self.wheel_out, wheel_rows)

        lines=[]
        lines.append("【雨流计数 + Miner 累积损伤（基于齿根应力代理）】")
        lines.append(f"- 累积损伤 D ≈ {m['damage_root']:.3e} （经验上 D<1 视为满足目标寿命）")
        if m.get("SF_root") is not None:
            lines.append(f"- 齿根安全系数 SF_root(峰值+SN) ≈ {m['SF_root']:.2f}")
        else:
            lines.append("- 齿根安全系数：未提供蜗轮 root SN 数据")
        if m.get("SF_contact") is not None:
            lines.append(f"- 齿面安全系数 SF_contact(峰值+SN) ≈ {m['SF_contact']:.2f}")
        else:
            lines.append("- 齿面安全系数：未提供蜗轮 contact SN 数据")
        lines.append("")
        lines.append("【关键几何/材料（核对）】")
        lines.append(f"- z2 ≈ {m['z2']}")
        lines.append(f"- d1 ≈ {m['d1_mm']:.2f} mm, d2 ≈ {m['d2_mm']:.2f} mm, a ≈ {m['a_mm']:.2f} mm")
        if m.get("a_target_mm") is not None:
            lines.append(f"- 目标中心距 a_target ≈ {m['a_target_mm']:.2f} mm, 偏差 Δa ≈ {m['delta_a_mm']:.2f} mm")
        lines.append(f"- 变位系数 x1={m['x1']:.3f}, x2={m['x2']:.3f}")
        lines.append(f"- 导程角 γ ≈ {m['gamma_deg']:.2f} deg")
        lines.append(f"- E' ≈ {m['Eprime_GPa']:.2f} GPa")
        lines.append(f"- 名义效率 η0 ≈ {m['eta0']:.3f}")
        lines.append(f"- K系数: KA={m['KA']:.3f}, KV={m['KV']:.3f}, KHβ={m['KHb']:.3f}, KFβ={m['KFb']:.3f}")
        self.fat_text.delete("1.0","end")
        self.fat_text.insert("1.0", "\n".join(lines))

    def export_xlsx(self):
        if self.res is None:
            messagebox.showwarning("提示", "请先计算一次，再导出。")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        inputs = {k:v.get() for k,v in self.inputs.items()}
        export_cycle_xlsx(path, inputs, self.steel, self.wheel, self.res)
        messagebox.showinfo("完成", f"已导出：\n{path}")

if __name__ == "__main__":
    App().mainloop()
