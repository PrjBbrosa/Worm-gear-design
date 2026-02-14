# Copilot instructions for Worm-gear-design

Short, actionable guidance to get productive with this repo.

- **Project type:** Desktop GUI analysis tool (Tkinter) — not a web service. Entry point: `app.py`.
- **Run locally:** Install requirements and run the GUI:

  pip install -r requirements.txt
  python app.py

- **System notes:** Uses `tkinter` and matplotlib `TkAgg` backend; ensure system Python has `tk` (macOS: usually installed with Python from Homebrew/Apple). `openpyxl` is required only for Excel export.

- **High-level data flow:**
  - UI (`app.py`) reads material JSON files from `materials/metals` and `materials/polymers` via `src.utils.load_json`.
  - User inputs (stored as strings) are passed to `src.worm_model.compute_worm_cycle(inputs, steel, wheel)` which converts strings to numeric values, computes phase-resolved arrays and a `meta` dict.
  - Results are plotted in the GUI and can be exported via `src.export_xlsx.export_cycle_xlsx(path, inputs, steel, wheel, res)`.

- **Key files to inspect/edit:**
  - `app.py` — GUI, i18n (`LANG_ZH`, `LANG_EN`), widget tracking (`_track`) and material selection logic.
  - `src/worm_model.py` — core numeric model; returns arrays (`phi`, `p_contact_MPa`, `sigma_root_MPa`, `T2_Nm`, `eta`, `Nc_proxy`) and `meta` (units: mm, MPa, N/m, etc.).
  - `src/export_xlsx.py` — Excel export (uses `openpyxl`).
  - `materials/*/*.json` — material cards. Common keys: `elastic` (E_GPa, nu), `elastic_T.points_C_GPa` (for temperature interpolation), and `SN` with `contact_allow_MPa_vs_N` and `root_allow_MPa_vs_N` arrays.

- **Project-specific conventions & gotchas:**
  - Inputs are read from UI as strings; `compute_worm_cycle` expects string-valued `inp` and parses numbers internally. Prefer reusing `compute_worm_cycle` rather than reimplementing conversions.
  - Material temperature-dependent modulus: either `elastic.E_GPa` (scalar) or `elastic_T.points_C_GPa` (list of [C, GPa]) — code uses linear interpolation.
  - S-N curves are expressed as arrays of pairs `[N_cycles, MPa]` (note order), and `_interp_sn` performs log10 interpolation on N.
  - GUI uses simple proxy models (sinusoidal stiffness/torque ripple); this is a trend tool, not FEM-accurate. Changes to formulas impact many downstream plots and exports.

- **Examples**
  - Invoke model from code (same signature GUI uses):

```py
from src.worm_model import compute_worm_cycle
from src.utils import load_json

steel = load_json('materials/metals/37CrS4.json')
wheel = load_json('materials/polymers/PA66_modified_draft.json')
inputs = {'T1_Nm':'6.0','n1_rpm':'3000','ratio':'25', 'z1':'2', 'mn_mm':'2.5', 'steps':'720'}
res = compute_worm_cycle(inputs, steel, wheel)
print(res['meta']['SF_root'], res['meta']['damage_root'])
```

- **Dependency notes**
  - `requirements.txt` contains `matplotlib`, `numpy`, `openpyxl`. `openpyxl` only needed for `Export XLSX` menu.

- **Developer workflow tips**
  - To investigate numeric changes, add small unit tests that call `compute_worm_cycle` with deterministic inputs (no GUI). There is currently no test suite.
  - For modifying material schema, update both `app.py` material-loading logic and `src/worm_model.py` interpolation helpers (`_interp_Et`, `_interp_sn`).
  - When changing plotting layout, update `app.py` UI builder and call sites that expect keys in `res` and `meta`.

- If anything in the above is unclear or you'd like examples in Chinese or additional CI/test scaffolding, tell me which parts to expand.
