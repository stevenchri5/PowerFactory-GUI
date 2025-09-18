# Ōtaki Grid – Solar Penetration (GUI + PowerFactory Backend)

## Overview
A Tkinter GUI (`gui_app.py`) controls PV penetration per suburb and runs a Quasi‑Dynamic Simulation (QDS) in DIgSILENT PowerFactory via `otaki_sim.py`. The GUI plots PV/load (kW) and transformer/line loading (%) and allows exporting current details to CSV.

- GUI window: sliders per suburb, results panel, graphs, and two map panes.
- Backend: connects to PF, activates the project/study case, registers monitors, runs QDS, and returns time‑series/results.

## Files
- **gui_app.py** — Tkinter UI, plotting, CSV export, and interaction with backend.
- **otaki_sim.py** — PowerFactory setup, monitoring, overrides, QDS execution, and results extraction.

## Requirements
- Windows with **DIgSILENT PowerFactory 2025 SP1**.
- Python **3.9** (matching PowerFactory’s embedded Python).
- Packages: `matplotlib`, `Pillow` (optional for images), `tkinter` (standard).

### PowerFactory
- Project: `ENGR489 Otaki Grid Base Solar and Bat(1)`
- Study Case: `Study Case`
- QDS timing: 1‑hour step, full‑day period (configurable).

## How It Works (Flow)
1. **GUI launch** → builds slider rows for each `PV_CONFIG` key and seeds states.
2. **User sets sliders** → GUI prepares overrides, shows pending % in results.
3. **Run** → GUI thread calls backend wrapper to: apply overrides, build monitored variables, prepare results, run QDS, and extract all series.
4. **Results back** → GUI updates slider states, results labels, and plots on suburb click.

## Key Concepts & Data
### Suburbs & Display Names
- `SUBURB_FULL` maps PV code (e.g., `OTBa_PV`) to user‑friendly names.
- Alphabetical ordering uses Unicode‑safe normalization.

### PV Configuration
`PV_CONFIG[pv_key] = {bus, load, pline, tx, homes}`  
- `homes` defines inverter count at 100% slider.  
- `bus` → `*.ElmTerm` for p.u. voltages.  
- `load` → `*.ElmLod` for kW demand.  
- `pline`/`tx` → line/transformer for loading %.

### Overrides
- **Inverters**: `PV_INV_OVERRIDES[pv_key] → ElmPvsys.ngnum`  
- **Panels/inverter**: `PV_PANEL_OVERRIDES[pv_key] → ElmPvsys.npnum/nPnum`  
- **Panel wattage**: `PANEL_WATT = 240 W` used to quantise kW/inverter.

### Results Structure
`RESULTS = { "bus":{}, "load":{}, "pv":{}, "tx":{}, "line":{}, "pv_meta":{} }`

Per element:
- **bus**: `{ "t":[…], "u_pu":[…], "u_pu_min", "u_pu_min_hour", "u_pu_max", "u_pu_max_hour" }`
- **load/pv**: `{ "t":[…], "P_W":[…] }`
- **tx/line**: `{ "t":[…], "loading_pct":[…] }`
- **pv_meta**: `{ pv_key: {"inverters":N, "panels_per_inv":M, "kw_per_inv":K} }`

## GUI Details (Selected Sections)
- **3.1.5**: Run/Settings/Export buttons row. `EXPORT CSV` triggers `_export_to_csv` (writes current details).  
- **3.1.6–3.1.8**: Scrollable sliders panel; each suburb has two rows: name + kW/inverter entry; slider + % entry; result labels.  
- **3.4.0**: `_build_slider_row` constructs one two‑row control group, with tool‑tips and live kW label.  
- **3.6.0**: `_update_kw_label` computes installed kW = `% × homes × kW/inverter`.  
- **3.7.0**: `_update_percent_entry` keeps `%` entry in sync with slider.  
- **3.9.0**: `_on_suburb_clicked` highlights button, resolves dataset names, and calls `_plot_curves`.  
- **3.9.5**: `_plot_curves` draws PV, load, tx%, line%; auto‑scales axes; sets ticks/grid; supports vertical lines for min/max p.u. hours.  
- **Graphs**: Matplotlib embedded via `FigureCanvasTkAgg`; twin y‑axes (kW left, % right).  
- **Maps**: Two canvases; paths configured via `load_map_images` (absolute paths).

## Backend Details (Selected Sections)
- **1.0**: Inserts PF Python path and imports `powerfactory` module.  
- **2.1–2.3**: Connects PF, activates project and study case.  
- **2.4–2.5**: Refreshes `PV_INV_OVERRIDES` and `PV_PANEL_OVERRIDES` from model.  
- **2.6–2.7**: Helpers to populate `pv_meta` with inverter counts and panels/kW per inverter.  
- **3.5–3.6**: Applies overrides to `ElmPvsys.ngnum` and `npnum/nPnum`.  
- **4.1**: Builds **monitored** dict for buses, loads, PVs, transformers, and lines.  
- **3.1/3.2** (prepare/run): Creates results file, adds variables, sets QDS timing, executes QDS.  
- **4.2**: Extracts all time‑series and computes bus min/max p.u.

## Run Book
1. Open PowerFactory; ensure project/study case exist and are consistent with `PV_CONFIG` naming.  
2. Run `gui_app.py`.  
3. Adjust sliders/kW per inverter; click **RUN**.  
4. Click a suburb to view curves; use **EXPORT CSV** if needed.

## CSV Export (GUI)
- Exports current settings and summary metrics per suburb (installed kW, % values, etc.).  
- Hourly series (PV, load, tx%, line%) are included when available.

## Troubleshooting
- **No graphs**: Install Matplotlib.  
- **No PF connection**: Check `DIG_PATH`, project/study names.  
- **Empty curves**: Ensure monitors matched model element names (see `*_LIST` and `PV_CONFIG`).  
- **Scaling odd**: Verify units (kW vs W) and that series are appended correctly.

## Extending
- Add suburbs: update `PV_CONFIG`, `*_LIST` sets, and `SUBURB_FULL`.  
- Add metrics to plots: include new variables in `build_monitored_dict` and extract them in `extract_qds_results`.

---

© 2025 Chris Stevenson – ENGR489 Honors Project
