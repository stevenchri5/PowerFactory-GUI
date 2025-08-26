#====================================================================================================
# 1.0 ✅ Set Up Environment
#====================================================================================================


import os, sys, time
DIG_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP1\Python\3.9"
sys.path.append(DIG_PATH)
os.environ['PATH'] += ';' + DIG_PATH
import powerfactory as pf


# 1.1.0 Project and Study Case Names-----------------------------------------------------------------
PROJECT_NAME    = "ENGR489 Otaki Grid Base Solar and Bat(1)"
STUDY_CASE_NAME = "Study Case"


# 1.2.0 Quasi-Dynamic Simulation Timing Settings ----------------------------------------------------
QDS_STEP_SIZE   = 1   # Step size in hours
QDS_STEP_UNIT   = 2   # 0 = seconds, 1 = minutes, 2 = hours, 3 = days
QDS_CALC_PERIOD = 0   # 0 = full day, 12 = 12 hours


# 1.3.0 Print Results in Terminal ------------------------------------------------------------------
PRINT_BUS_HOURLY        = False
PRINT_LOAD_HOURLY       = False
PRINT_PV_HOURLY         = False
PRINT_TX_HOURLY         = False
PRINT_BUS_MIN_MAX       = False
PRINT_LINE_HOURLY       = False
PRINT_VARIABLE_CHECKS   = False
PRINT_PV_META           = True
PRINT_PV_OVERRIDES      = False


# 1.4.0 PV panel wattage (per panel)
PANEL_WATT = 240.0  # W


# 1.5.0 Results Store and Monitor Registry
RESULTS   = {"bus": {}, "load": {}, "pv": {}, "tx": {},"line": {}}
ASSOC     = {}   # load ➜ {"bus": <bus>, "pv": [..], "tx": [..], "line": [..]}
MONITORED = []   # list of (selector, [vars]) for GUI

# 1.6.0 PV inverter overrides (GUI sets before run)
PV_INV_OVERRIDES = {}


# 1.5.0 Full Bus Element List-------------------------------------------------------------------------
BUS_LIST = {
    "OTBa_0.415",
    "OTBb_0.415",
    "OTBc_0.415",
    "OTCa_0.415",
    "OTCb_0.415",
    "OTIa_0.415",
    "OTKa_0.415",
    "OTKb_0.415",
    "OTKc_0.415",
    "OTS_0.415",
    "RGUa_0.415",
    "RGUb_0.415",
    "TRE_0.415",
    "WTVa_0.415",
    "WTVb_0.415",
    "WTVc_0.415",
}


# 1.6.0 Full Load Element List------------------------------------------------------------------------ 
LOAD_LIST = {
    "Otaki Beach A",
    "Otaki Beach B",
    "Otaki Beach C",
    "Otaki Commercial A",
    "Otaki Commercial B",
    "Otaki Industrial A",
    "Otaki School",
    "Otaki Town A",
    "Otaki Town B",
    "Otaki Town C",
    "Rangiuru Rd A",
    "Rangiuru Rd B",
    "Te Rauparaha St",
    "Waitohu Valley A",
    "Waitohu Valley B",
    "Waitohu Valley C",
}


#1.7.0 Full PV Element list---------------------------------------------------------------------------
PV_LIST = {
    "OTBa_PV",
    "OTBb_PV",
    "OTBc_PV",
    "OTCa_PV",
    "OTCb_PV",
    "OTIa_PV",
    "OTKa_PV",
    "OTKb_PV",
    "OTKc_PV",
    "OTS_PV",
    "RGUa_PV",
    "RGUb_PV",
    "TRE_PV",
    "WTVa_PV",
    "WTVb_PV",
    "WTVc_PV",
}


#1.8.0 Full TX Element list--------------------------------------------------------------------------
TX_LIST = {
    "OTB_T1",
    "OTB_T2",
    "OTB_T3",
    "OTCa_T1",
    "OTCb_T1",
    "OTI_T1",
    "OTK_T1",
    "OTK_T2",
    "OTK_T3",
    "OTS_T1",
    "RGU_T1",
    "RGU_T2",
    "TRE_T1",
    "WTV_T1",
    "WTV_T2",
    "WTV_T3",
}


#1.9.0 Full Line Element list------------------------------------------------------------------------
LINE_LIST = {
    "OTBa_pline_OTBa_0.415",
    "OTBb_pline_OTBb_0.415",
    "OTBc_pline_OTBc_0.145",
    "OTCa_pline_OTCa_0.415",
    "OTCb_pline_OTCb_0.415",
    "OTIa_pline_OTIa_0.415",
    "OTKa_pline_OTKa_0.415",
    "OTKb_pline_OTKb_0.415",
    "OTKc_pline_OTKc_0.415",
    "OTS_pline_OTS_0.415",
    "RGUa_pline_RGUa_0.415",
    "RGUb_pline_RGUb_0.415",
    "TRE_pline_TRE_0.415",
    "WTVa_pline_WTVa_0.415",
    "WTVb_pline_WTVb_0.415",
    "WTVc_pline_WTVc_0.415",
}

# 1.10.0 mapping for GUI use
PV_CONFIG = {
    "OTBa_PV": {"bus": "OTBa_0.415", "load": "OTBa", "pline": "OTBa_pline_OTBa_0.415", "tx": "OTB_T1", "homes": 100},
    "OTBb_PV": {"bus": "OTBb_0.415", "load": "OTBb", "pline": "OTBb_pline_OTBb_0.415", "tx": "OTB_T2", "homes": 100},
    "OTBc_PV": {"bus": "OTBc_0.415", "load": "OTBc", "pline": "OTBc_pline_OTBc_0.145", "tx": "OTB_T3", "homes": 100},
    "OTCa_PV": {"bus": "OTCa_0.415", "load": "OTCa", "pline": "OTCa_pline_OTCa_0.415", "tx": "OTCa_T1", "homes": 1},
    "OTCb_PV": {"bus": "OTCb_0.415", "load": "OTCb", "pline": "OTCb_pline_OTCb_0.415", "tx": "OTCb_T1", "homes": 1},
    "OTIa_PV": {"bus": "OTIa_0.415", "load": "OTIa", "pline": "OTIa_pline_OTIa_0.415", "tx": "OTI_T1", "homes": 1},
    "OTKa_PV": {"bus": "OTKa_0.415", "load": "OTKa", "pline": "OTKa_pline_OTKa_0.415", "tx": "OTK_T1", "homes": 100},
    "OTKb_PV": {"bus": "OTKb_0.415", "load": "OTKb", "pline": "OTKb_pline_OTKb_0.415", "tx": "OTK_T2", "homes": 100},
    "OTKc_PV": {"bus": "OTKc_0.415", "load": "OTKc", "pline": "OTKc_pline_OTKc_0.415", "tx": "OTK_T3", "homes": 100},
    "OTS_PV":  {"bus": "OTS_0.415",  "load": "OTS",  "pline": "OTS_pline_OTS_0.415",  "tx": "OTS_T1",  "homes": 1},
    "RGUa_PV": {"bus": "RGUa_0.415", "load": "RGUa", "pline": "RGUa_pline_RGUa_0.415", "tx": "RGU_T1", "homes": 100},
    "RGUb_PV": {"bus": "RGUb_0.415", "load": "RGUb", "pline": "RGUb_pline_RGUb_0.415", "tx": "RGU_T2", "homes": 100},
    "TRE_PV":  {"bus": "TRE_0.415",  "load": "TRE",  "pline": "TRE_pline_TRE_0.415",  "tx": "TRE_T1",  "homes": 100},
    "WTVa_PV": {"bus": "WTVa_0.415", "load": "WTVa", "pline": "WTVa_pline_WTVa_0.415", "tx": "WTV_T1", "homes": 100},
    "WTVb_PV": {"bus": "WTVb_0.415", "load": "WTVb", "pline": "WTVb_pline_WTVb_0.415", "tx": "WTV_T2", "homes": 100},
    "WTVc_PV": {"bus": "WTVc_0.415", "load": "WTVc", "pline": "WTVc_pline_WTVc_0.415", "tx": "WTV_T3", "homes": 100},
}


#====================================================================================================
# 2.0 ✅ Connect, Activate Project & Study Case ✅ Working Perfectly, Don't fucking touch
#====================================================================================================


#2.1 Connect to PowerFactory
print("2.1.0 🔌    Connecting to PowerFactory…")
app = pf.GetApplication()
if not app:
    print("2.1.0 ❌ Could not connect to PowerFactory."); sys.exit(1)
print("2.1.0 ✅     Connected.")
print()


#2.2 Activate project
print(f"2.2.0 📂    Activating project: {PROJECT_NAME}")
if app.ActivateProject(PROJECT_NAME) != 0:
    print(f"2.2.0 ❌ Could not activate project '{PROJECT_NAME}'."); sys.exit(1)
print(f"2.2.0 ✅     Project activated: {PROJECT_NAME}")
print()


# 2.3 Activate study case
print(f"2.3.0 🔎    Looking for study case: {STUDY_CASE_NAME}")
study_folder = app.GetProjectFolder("study")
active_case = None
for case in study_folder.GetContents("*.IntCase", 1):
    if case.loc_name == STUDY_CASE_NAME:
        case.Activate()
        active_case = case
        break
if not active_case:
    print(f"2.3.0 ❌ Study case '{STUDY_CASE_NAME}' not found."); sys.exit(1)
print(f"2.3.0 ✅     Study Case activated: {active_case.loc_name}")
print()


# 2.4.0 ▶️    Refresh PV inverter overrides from model
def refresh_pv_overrides_from_model(app):
    print("2.4.0 ▶️    Read current PV inverter counts.")
    found = 0
    for code in PV_LIST:
        for p in (app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []):
            PV_INV_OVERRIDES[p.loc_name] = int(getattr(p, "ngnum", 0) or 0)
            found += 1
            if PRINT_PV_OVERRIDES:
                print(f"2.4.1 ✅     {p.loc_name}: ngnum={PV_INV_OVERRIDES[p.loc_name]}")
    print(f"2.4.0 ✅     Refreshed {found} PV entries.")
    print()
refresh_pv_overrides_from_model(app)


# 2.5.0 get_inverter_counts
def get_inverter_counts():
    global RESULTS
    RESULTS.setdefault("pv_meta", {})
    pv_meta = RESULTS["pv_meta"]
    pv_meta.clear()
    for pv_key, meta in PV_CONFIG.items():
        homes = int(meta.get("homes", 0))
        pv_meta[pv_key] = {"inverters": homes}
        print(f"[otaki_sim] {pv_key}: inverters={homes}")
    return pv_meta


#====================================================================================================
# 3.0 ✅ Quasi‑Dynamic Simulation Setup
#====================================================================================================


# 3.2 Results file setup ----------------------------------------------------------------------------
def prepare_quasi_dynamic(app, monitored_vars,
                          step_size=QDS_STEP_SIZE,
                          step_unit=QDS_STEP_UNIT,
                          period=QDS_CALC_PERIOD):
    print("3.1.0 ▶️    Results Setup Begin.")
    qds = app.GetFromStudyCase("ComStatsim")
    if not qds:
        raise RuntimeError("3.1.0 ❌ [Results Setup] ComStatsim not found.")
    res = qds.results
    if not res:
        raise RuntimeError("3.1.0 ❌ [Results Setup] QDS results object missing.")
    res.Clear()
    print("3.1.0 ✅     Results Setup ready.")
    print()

    print("3.1.2 ▶️    Adding Variables to Check.")
    for sel, var_list in monitored_vars.items():
        elems = app.GetCalcRelevantObjects(sel) or []
        for e in elems:
            res.AddVars(e, *var_list)
            if PRINT_VARIABLE_CHECKS:
                print(f"3.1.2 ➕    Checking Variables {e.loc_name}: {var_list}")
    print("3.1.2 ✅     Variable Checking Done.")
    print()

    print("3.1.3 ▶️    Configure Quasi-Dynamic Simulation Timing.")
    qds.stepSize   = step_size
    qds.stepUnit   = step_unit
    qds.calcPeriod = period
    print(f"3.1.3 ✅     Timing: period={period}, step={step_size}, unit={step_unit}")

    print("3.1.4 ✅     Quasi-Dynamic Simulation Ready to Run.")
    print()
    return res, qds


# 3.3 Execute QDS — runs the quasi‑dynamic simulation ---------------------------------------------
def run_quasi_dynamic(qds):
    print("3.2.0 ▶️    Quasi-Dynamic Simulation Running…")
    ok = (qds.Execute() == 0)
    print("3.2.0 ✅     Quasi-Dynamic Simulation Finished. --------------------- " if ok else 
          "3.2.0 ❌ Quasi-Dynamic Simulation Failed.")
    print()
    return ok


# 3.4 Read results — loads time series for one element/variable ------------------------------------
def get_dynamic_results(app, res, elm_selector, var_name, verbose=True):
    elems = app.GetCalcRelevantObjects(elm_selector) or []
    name = elems[0].loc_name if elems else elm_selector
    if verbose:
        print(f"3.3.0 ▶️    Results For {name}")
    if not elems:
        if verbose: print(f"3.3.0 ⚠️ [Read Results] Element not found for '{elm_selector}' ({var_name}).")
        return [], []
    elm = elems[0]
    app.ResLoadData(res)
    col = app.ResGetIndex(res, elm, var_name)
    if col < 0:
        if verbose: print(f"3.3.0 ⚠️ [Read Results] Var '{var_name}' not in results for {name}.")
        return [], []
    n = app.ResGetValueCount(res, 0)
    if verbose:
        print(f"3.3.0 ℹ️ Reading Results For {name}, Variable={var_name}, Rows={n}")
    t, v = [], []
    for i in range(n):
        t.append(app.ResGetData(res, i, -1)[1])
        v.append(app.ResGetData(res, i, col)[1])
    if verbose:
        print(f"3.3.0 ✅     Results For {name} Extracted.")
    return t, v


# 3.5 Apply PV inverter overrides (set ElmPvsys.ngnum)
def apply_pv_inverter_overrides(app, overrides):
    print("3.4.0 ▶️    Apply PV Inverter Overrides.")
    if not overrides:
        print("3.4.0 ✅     No overrides provided."); return
    for code, count in overrides.items():
        pvs = app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []
        if not pvs:
            print(f"3.4.0 ⚠️    No PV matched '{code}'"); continue
        for p in pvs:
            old = getattr(p, "ngnum", None)
            try:
                p.ngnum = int(count)
                if PRINT_PV_OVERRIDES:
                    print(f"3.4.4 ✅     {p.loc_name}: ngnum {old} → {p.ngnum}")
            except Exception as e:
                print(f"3.4.0 ❌    {p.loc_name}: failed to set ngnum: {e}")
    print("3.4.0 ✅     Overrides applied.")
    print()


#====================================================================================================
# 4.0 ✅ Quasi-Dynamic Simulation Core (Wrapped for GUI Use)
#====================================================================================================


# 4.1 Build full monitored variable dict -----------------------------------------------------------
def build_monitored_dict():
    print("4.0.0 ▶️    Quasi-Dynamic Simulation Setup.")
    bus_selectors  = {f"{bus}.ElmTerm": ["m:u1"] for bus in BUS_LIST}             # p.u. voltage
    load_selectors = {f"{code}*.ElmLod": ["m:P:bus1"] for code in LOAD_LIST}      # kW demand
    pv_selectors   = {f"{code}*.ElmPvsys": ["m:Psum:bus1"] for code in PV_LIST}   # PV output
    tx_selectors   = {f"{code}*.ElmTr2": ["m:loading:bus1"] for code in TX_LIST}  # % loading
    line_selectors = {f"{name}.ElmLne":   ["m:loading:bus1"] for name in LINE_LIST}  # % loading

    monitored = {}
    monitored.update(bus_selectors)
    monitored.update(load_selectors)
    monitored.update(pv_selectors)
    monitored.update(tx_selectors)
    monitored.update(line_selectors)

    MONITORED[:] = [(k, v) for k, v in monitored.items()]                         # update global list for GUI access
    print(f"4.0.3 ℹ️    Register Monitors {len(MONITORED)} entries. ------------")
    print()
    return monitored


# 4.2 Extract all results (same as before) -----------------------------------------------
def extract_qds_results(app, res):
    # 4.2 BUS Hourly p.u. voltage (24hrs)-----------------------------------------------------------
    print("4.2.0 ▶️     Hourly p.u. Voltages Begin.")
    for bus in BUS_LIST:
        sel = f"{bus}.ElmTerm"
        t, u = get_dynamic_results(app, res, sel, "m:u1", verbose=PRINT_BUS_HOURLY)
        if not u:
            print(f"4.2.0 ⚠️ No voltage data for {bus}")
            continue
        RESULTS["bus"][bus] = {"t": t[:24], "u_pu": u[:24]}
        if PRINT_BUS_HOURLY:
            for hr in range(min(24, len(u))):
                print(f"{hr:02d} {bus} = {u[hr]:.4f} p.u")
    print("4.2.0 ✅     Hourly p.u. Voltages Done. ------------------------------")
    print()


# 4.2.1 Bus 24hr Min/Max Voltage p.u-------------------------------------------------------------
    print("4.2.1 ▶️    Bus 24hr Min/Max p.u Begin.")
    for bus in BUS_LIST:
        rec = RESULTS["bus"].get(bus)
        if not rec:
            print(f"4.2.1 ⚠️ No data for {bus}"); continue
        u = rec.get("u_pu", [])[:24]
        if not u:
            print(f"4.2.1 ⚠️ Empty voltage list for {bus}"); continue
        u_min = min(u); h_min = u.index(u_min)
        u_max = max(u); h_max = u.index(u_max)
        rec.update({
            "u_pu_min": u_min, "u_pu_min_hour": h_min,
            "u_pu_max": u_max, "u_pu_max_hour": h_max
        })
        if PRINT_BUS_MIN_MAX:
            print(f"4.2.1 📌 {bus} min={u_min:.4f} @ {h_min:02d}h, max={u_max:.4f} @ {h_max:02d}h")
    print("4.2.1 ✅     Bus 24hr Min/Max p.u Done. -------------------------------")
    print()


# 4.3 LOAD Hourly demand (24hrs)---------------------------------------------------------------
    print("4.3.0 ▶️   Hourly Load Demand Begin.")
    for load in LOAD_LIST:
        sel = f"{load}*.ElmLod"
        loads = app.GetCalcRelevantObjects(sel) or []
        if not loads:
            print(f"4.3.0 ⚠️ No loads matched '{load}'"); continue
        for ld in loads:
            tP, P = get_dynamic_results(app, res, ld.loc_name + ".ElmLod", "m:P:bus1", verbose=PRINT_LOAD_HOURLY)
            if not P:
                print(f"4.3.0 ⚠️ No demand data for {ld.loc_name}"); continue
            RESULTS["load"][ld.loc_name] = {"t": tP[:24], "P_W": P[:24]}
            if PRINT_LOAD_HOURLY:
                for hr in range(min(24, len(P))):
                    print(f"{hr:02d} {ld.loc_name} = {P[hr]:.4f} W")
                print()
    print("4.3.0 ✅     Hourly Load Demand Done. ---------------------------------")
    print()


# 4.4 PV Hourly Production (24hrs)---------------------------------------------------------------
    print("4.4.0 ▶️   Hourly PV Array Production Begin.")
    for pv in PV_LIST:
        pvs = app.GetCalcRelevantObjects(f"{pv}*.ElmPvsys") or []
        if not pvs:
            print(f"4.4.0 ⚠️ No PV objects found for '{pv}'"); continue
        for p in pvs:
            sel = f"{p.loc_name}.ElmPvsys"
            tP, P = get_dynamic_results(app, res, sel, "m:Psum:bus1", verbose=PRINT_PV_HOURLY)
            if not P:
                print(f"4.4.0 ⚠️ No production data for {p.loc_name}"); continue
            RESULTS["pv"][p.loc_name] = {"t": tP[:24], "P_W": P[:24]}
            if PRINT_PV_HOURLY:
                for hr in range(min(24, len(P))):
                    print(f"{hr:02d} {p.loc_name} = {P[hr]/1000:.2f} kW")
                print()
    print("4.4.0 ✅     Hourly PV Production Done. --------------------------------")
    print()


# 4.4.1 PV Nameplate & Inverters
    print("4.4.1 ▶️    PV Variables.")
    if "pv_meta" not in RESULTS: RESULTS["pv_meta"] = {}
    for code in PV_LIST:
        pvs = app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []
        if not pvs:
            print(f"4.4.1 ⚠️   No PV Found '{code}'"); continue
        for p in pvs:
            inv_objs = app.GetCalcRelevantObjects(p.loc_name + ".ElmInv") or []
            inv_count = len(inv_objs) if inv_objs else int(getattr(p, "ngnum", 0) or 0)
            par_mods = int(getattr(p, "npnum", 0) or 0)
            rating_kw_calc = (inv_count * par_mods * float(PANEL_WATT)) / 1000.0
            RESULTS["pv_meta"][p.loc_name] = {
                "rating_kW_calc": rating_kw_calc,
                "inverters": inv_count,
                "panels_per_inverter": par_mods,
            }
            if PRINT_PV_META:
                prefix = f"4.4.1 ✅     {p.loc_name}: "
                pad = " " * len(prefix)
                print(f"{prefix}Nameplate Rating = {rating_kw_calc:.2f} kW,")
                print(f"{pad} Inverters in Parallel = {inv_count},")
                print(f"{pad} Panels per inverter = {par_mods}")
    print("4.4.1 ✅     PV Nameplate & Inverters Done. ----------------------------")
    print()


# 4.5 TRANSFORMER Loading ------------------------------------------------------------------------
    print("4.5.0 ▶️   Hourly Tx Loading Begin.")
    for tx in TX_LIST:
        tx_objs = app.GetCalcRelevantObjects(f"{tx}*.ElmTr2") or []
        if not tx_objs:
            print(f"4.5.0 ⚠️ No Transformer Found '{tx}'"); continue
        for t in tx_objs:
            tP, P = get_dynamic_results(app, res, t.loc_name + ".ElmTr2", "m:loading:bus1", verbose=PRINT_TX_HOURLY)
            if not P:
                print(f"4.5.0 ⚠️ No Loading Data for {t.loc_name}"); continue
            RESULTS["tx"][t.loc_name] = {"t": tP[:24], "loading_pct": P[:24]}
            if PRINT_TX_HOURLY:
                for hr in range(min(24, len(P))):
                    print(f"{hr:02d} {t.loc_name} = {P[hr]:.2f} %")
                print()
    print("4.5.0 ✅     Hourly Tx Loading Done. ------------------------------------")
    print()


# 4.6 LINE Loading --------------------------------------------------------------------------------
    print("4.6.0 ▶️    Hourly Line Loading Begin.")
    if "line" not in RESULTS: RESULTS["line"] = {}
    for line in LINE_LIST:
        lines = app.GetCalcRelevantObjects(f"{line}.ElmLne") or []
        if not lines:
            print(f"4.6.0 ⚠️ No Line Found '{line}'"); continue
        for ln in lines:
            tL, L = get_dynamic_results(app, res, ln.loc_name + ".ElmLne", "m:loading:bus1", verbose=PRINT_LINE_HOURLY)
            if not L:
                print(f"4.6.0 ⚠️ No Loading Data for Line {ln.loc_name}"); continue
            RESULTS["line"][ln.loc_name] = {"t": tL[:24], "loading_pct": L[:24]}
            if PRINT_LINE_HOURLY:
                for hr in range(min(24, len(L))):
                    print(f"{hr:02d} {ln.loc_name} Loading = {L[hr]:.2f} %")
                print()
    print("4.6.0 ✅     Hourly Line Loading Done. ----------------------------------")
    print()


# 4.7 BUILD ASSOCIATIONS --------------------------------------------------------------------------
    print("4.7.0 ▶️    Build Associations Begin.")
    ASSOC.clear()
    all_pv  = app.GetCalcRelevantObjects("*.ElmPvsys") or []
    all_tx2 = app.GetCalcRelevantObjects("*.ElmTr2")   or []
    all_ln  = app.GetCalcRelevantObjects("*.ElmLne")   or []

    for code in LOAD_LIST:
        loads = app.GetCalcRelevantObjects(f"{code}*.ElmLod") or []
        for ld in loads:
            bus_term = getattr(ld, "bus1", None)
            bus_name = getattr(bus_term, "loc_name", None) if bus_term else None

            pv_names = [pv.loc_name for pv in all_pv
                        if getattr(pv, "bus1", None) is bus_term]

            tx_names = [tx.loc_name for tx in all_tx2
                        if (getattr(tx, "buslv", None) is bus_term)
                        or (getattr(tx, "bus1", None) is bus_term)
                        or (getattr(tx, "bus2", None) is bus_term)]

            line_names = []
            for ln in all_ln:
                b1 = getattr(ln, "bus1", None)
                b2 = getattr(ln, "bus2", None)
                if b1 is bus_term or b2 is bus_term:
                    line_names.append(ln.loc_name)
                elif code in ln.loc_name and ln.loc_name not in line_names:
                    line_names.append(ln.loc_name)

            ASSOC[ld.loc_name] = {"bus": bus_name, "pv": pv_names, "tx": tx_names, "line": line_names}
    print(f"4.7.0 ✅     Build Associations: {len(ASSOC)} loads mapped to bus/pv/tx/line.")
    print()


# 4.3 MAIN ENTRY POINT -------------------------------------------------------------------------------
def run_simulation(pv_overrides=None):
    """
    Executes the full QDS simulation. Accepts optional PV inverter overrides.
    Updates global RESULTS and ASSOC. Returns True if OK, False otherwise.
    """
    if pv_overrides:
        apply_pv_inverter_overrides(app, pv_overrides)
    monitored = build_monitored_dict()
    res, qds = prepare_quasi_dynamic(app, monitored)

    print("4.1.0 ▶️    Execute QDS Begin.")
    ok = run_quasi_dynamic(qds)
    print(f"4.1.0 ✅     Execute QDS Status = {ok} -------------------------------")
    print()
    if not ok:
        return False

    extract_qds_results(app, res)
    return True


run_simulation()