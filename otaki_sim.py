# Otaki_sim.py
# ENGR489 PowerFactory model - Mo Working Code


# =====================================================================================================================================================
# =====================================================================================================================================================
# 1.0 ---------- Setup & configuration ----------  # imports, PowerFactory environment, user settings, PV mapping
# =====================================================================================================================================================
# =====================================================================================================================================================

# 1.1 Standard library imports  
import sys                                                                  
import os                                                                   
import json                                                                 
import math                   
import datetime as dt         
from collections import defaultdict
import numpy as np   

# 1.2 PowerFactory import 
DIG_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP1\Python\3.9"   # PF’s bundled Python site-packages path (adjust if your install differs)
if os.path.isdir(DIG_PATH) and DIG_PATH not in sys.path:                    # only append if the folder exists and isn’t already on sys.path
    sys.path.append(DIG_PATH)                                               # allow `import powerfactory` to resolve when launched within PF
    os.environ["PATH"] += ";" + DIG_PATH                                    # also extend PATH for any PF-side DLL lookups


try:
    import powerfactory                                                     # guarded import; if this works we’re likely running within PF’s Python
except Exception:
    powerfactory = None                                                     # we’ll check this before any PF calls  # sentinel so code can detect lack of PF at runtime


# 1.3 Initial settings  
USE_QDS_FIRST = True                                                        # prefer quasi-dynamic simulation (time series) if the study case supports it
QDS_TSTART_S  = 0.0                                                         # start time in seconds from midnight (0 = 00:00)
QDS_TSTOP_S   = 24*3600.0                                                   # stop time in seconds (24 hours window)
QDS_DT_S      = 3600.0                                                      # simulation step size in seconds (1 hour steps)
V_OK_MIN      = 0.95                                                        # lower acceptable voltage limit in per-unit
V_OK_MAX      = 1.05                                                        # upper acceptable voltage limit in per-unit
PDF_DEFAULT   = os.path.abspath("otaki_results.pdf")                        # default output path for generated PDF report

# 1.6 PV configuration (keys must match GUI)                                # GUI keys mapped to PF bus + category
PV_CONFIG = {
    # suburb_key : {"bus": "<bus name in PF>", "type": "res"/"com"/"ind"/"school"}
    "OTBa_PV": {"bus": "OTBa", "type": "res"},                              # Ōtaki Beach A, residential
    "OTBb_PV": {"bus": "OTBb", "type": "res"},                              # Ōtaki Beach B, residential
    "OTBc_PV": {"bus": "OTBc", "type": "res"},                              # Ōtaki Beach C, residential
    "OTCa_PV": {"bus": "OTCa", "type": "com"},                              # Ōtaki Commercial A
    "OTCb_PV": {"bus": "OTCb", "type": "com"},                              # Ōtaki Commercial B
    "OTIa_PV": {"bus": "OTIa", "type": "ind"},                              # Ōtaki Industrial
    "OTKa_PV": {"bus": "OTKa", "type": "res"},                              # Ōtaki Township A
    "OTKb_PV": {"bus": "OTKb", "type": "res"},                              # Ōtaki Township B
    "OTKc_PV": {"bus": "OTKc", "type": "res"},                              # Ōtaki Township C
    "OTS_PV" : {"bus": "OTS",  "type": "school"},                           # Ōtaki College (school)
    "RGUa_PV": {"bus": "RGUa", "type": "res"},                              # Rangiuru A
    "RGUb_PV": {"bus": "RGUb", "type": "res"},                              # Rangiuru B
    "TRE_PV" : {"bus": "TRE",  "type": "res"},                              # Te Roto
    "WTVa_PV": {"bus": "WTVa", "type": "res"},                              # Waitohu Valley A
    "WTVb_PV": {"bus": "WTVb", "type": "res"},                              # Waitohu Valley B
    "WTVc_PV": {"bus": "WTVc", "type": "res"},                              # Waitohu Valley C
}

# 1.7 Capacity rules (kW) for slider meaning  # maps slider percentage to applied PV capacity by site type
"""
residential: slider 0..100 maps to 0..100 inverters at 6 kW each (npar)
commercial/industrial: max 100 kW (fixed P setpoint)
school: max 25 kW (fixed P setpoint)
"""
RES_KW_PER_INVERTER = 6.0                                                   # each residential percent is 1 inverter & 6 kW
COM_MAX_KW   = 100.0                                                        # cap for commercial sites
IND_MAX_KW   = 100.0                                                        # cap for industrial sites
SCH_MAX_KW   = 25.0                                                         # cap for school site

# 1.8 Helper function to safely fetch attributes from PowerFactory objects
def _get_attr(obj, *candidates):
    """Return the first non-None attribute among the candidate names."""
    for name in candidates:                                                 # Loop through each candidate attribute name provided
        try:
            val = obj.GetAttribute(name)                                    # Try to get the attribute from the PowerFactory object
            if val is not None:                                             # If the attribute exists and is not None, return it immediately
                return val
        except Exception:                                                   # If the attribute does not exist or can't be accessed, ignore the error and continue
            pass
    return None                                                             # If none of the candidates worked, return None as a safe fallback




# =====================================================================================================================================================
# =====================================================================================================================================================
# 2.0 ---------- PowerFactory Connection 
# =====================================================================================================================================================
# =====================================================================================================================================================

# 2.1 Connect to PowerFactory, make sure a project is active, return the Application object
def connect_and_activate():
    """
    Start/attach to PowerFactory and return the Application object.
    Uses GetApplicationExt to ensure proper error propagation.  :contentReference[oaicite:6]{index=6}
    """
    if powerfactory is None:                                                # make sure this only runs inside PF python bit
        raise RuntimeError("This module must be run inside the PowerFactory Python environment.")
    app = powerfactory.GetApplicationExt()                                  # raises if PF cannot start :contentReference[oaicite:7]{index=7}  # request app handle with better error propagation
    app.Show()                                                              # make UI visible (optional)  :contentReference[oaicite:8]{index=8}  # bring the PF GUI to the foreground
    prj = app.GetActiveProject()                                            # fetch the currently active project (if any)
    if not prj:                                                             # no project open 
        raise RuntimeError("No active project. Please open your project/study case.")
    return app                                                              # return the live Application object for further API calls

# 2.2 Make sure a Study Case is active and return it
def get_studycase(app):
    sc = app.GetActiveStudyCase()                                           # ask PF for the active Study Case on this project
    if not sc:                                                              # none selected, you will have to activate one
        raise RuntimeError("No active Study Case. Please activate a study case.")
    return sc                                                               # return the Study Case object

# 2.3 Build a name,terminal map for all buses wanted
def get_bus_objects(app):
    buses = {}                                                              # prepare result dict: {loc_name: ElmTerm}
    for term in app.GetCalcRelevantObjects("*.ElmTerm") or []:              # :contentReference[oaicite:9]{index=9}  # iterate LV/HV terminals in the active calc scope
        buses[term.loc_name] = term                                         # map human-readable name to the terminal object
    return buses                                                            # return the mapping for quick lookups by name

# 2.3.1 Bus-name resolver — map configured names to actual PF terminal names
def _resolve_bus_name(config_bus: str, bus_dict: dict):
    """
    Return the actual terminal name from bus_dict that best matches config_bus.
    Priority: exact match -> common LV suffix -> startswith -> contains -> None.
    """
    # 1) Exact match
    if config_bus in bus_dict:
        return config_bus

    # 2) Common LV suffix used in many models (e.g., "OTKa" -> "OTKa_0.415")
    cand = f"{config_bus}_0.415"
    if cand in bus_dict:
        return cand

    # 3) Name starts with configured token (e.g., "OTKa" matches "OTKa LV Bus")
    for name in bus_dict:
        if name.startswith(config_bus):
            return name

    # 4) Fuzzy contains as last resort
    for name in bus_dict:
        if config_bus in name:
            return name

    # 5) No match
    return None

# 2.4 Find the PV object that the GUI is asking for
def get_pv_objects(app):
    """
    Return dict {pv_key: ElmGenstat} matching PV_CONFIG entries by name pattern.
    You can adapt matching to your actual naming in the model.
    """
    pv_dict = {}                                                            # result mapping: GUI pv_key → PF object
                                                                            # Try PV systems, then generic stat gens (model-dependent)
    candidates = (app.GetCalcRelevantObjects("*.ElmPvsys") or []
                  ) + (app.GetCalcRelevantObjects("*.ElmGenstat") or [])    # :contentReference[oaicite:10]{index=10}  # include static generators as fallback
    by_name = {obj.loc_name: obj for obj in candidates}                     # quick index by object display name
    for key, meta in PV_CONFIG.items():                                     # iterate all configured PV keys expected by the GUI
        guess = key                                                         # first guess: object named exactly like the pv_key
        obj = by_name.get(guess)                                            # try direct match
        if obj is None:                                                     # if not found, try an alternative naming scheme
            alt = f"{meta['bus']}_PV"                                       # e.g., busname_PV
            obj = by_name.get(alt)                                          # look up the alternate name
        if obj:                                                             # if either guess worked, record the mapping
            pv_dict[key] = obj
    missing = [k for k in PV_CONFIG.keys() if k not in pv_dict]             # anything not resolved is considered missing
    return pv_dict, missing                                                 # return both the found objects and a list of missing pv_keys

# 2.5 Find lines, transformers, and loads that are relevant to the active calculation
def discover_network_elements(app):
    """Find lines, trafos, loads used by the active calculation."""
    lines  = app.GetCalcRelevantObjects("*.ElmLne") or []                   # :contentReference[oaicite:11]{index=11}  # line elements
    traf2  = app.GetCalcRelevantObjects("*.ElmTr2") or []                   # :contentReference[oaicite:12]{index=12}  # 2-winding transformers
    traf3  = app.GetCalcRelevantObjects("*.ElmTr3") or []                   # :contentReference[oaicite:13]{index=13}  # 3-winding transformers
    trafos = (traf2 or []) + (traf3 or [])                                  # merge into a single list
    loads  = app.GetCalcRelevantObjects("*.ElmLod") or []                   # :contentReference[oaicite:14]{index=14}  # load elements
    return {"lines": lines, "trafos": trafos, "loads": loads}               # package everything into a dict for convenience



# =====================================================================================================================================================
# =====================================================================================================================================================
# 3.0 ---------- Change PV to match sliders                                 # apply GUI slider values to the PF PV objects
# =====================================================================================================================================================
# =====================================================================================================================================================

# 3.1 Map 0..100% → 0..100 parallel inverters (≈6 kW each)
def _apply_residential_pv(pv_obj, percent):
    npar = int(round(max(0.0, min(100.0, percent))))                        # Lock percent to [0,100], round, and treat as number of inverters
    try:
        pv_obj.npar = npar                                                  # set number of parallel units on the PF PV object
    except Exception:
        pass                                                                # if the object doesn’t expose npar, swallow and continue (defensive)

# 3.2 For non-resi sites, write a kW (or MW) setpoint
def _apply_fixed_kw(pv_obj, kw):
    """
    For commercial/industrial/school — directly set active power setpoint (pgini) in kW.
    Some models use MW; if so, divide by 1000.
    """
    try:                                                                    # Detect unit by magnitude of the rated power if available
        pv_obj.pgini = kw                                                   # prefer kW if the model expects kW
    except Exception:
        try:
            pv_obj.pgini = kw / 1000.0                                      # fallback: some templates define pgini in MW
        except Exception:
            pass                                                            # if neither assignment works, ignore (object mismatch)

# 3.3 Apply sliders, return any missing PV keys
def set_pv_penetrations(app, sliders):
    pv_objs, missing = get_pv_objects(app)                                  # resolve pv_key → PF object; also gather list of missing keys
    for key, percent in sliders.items():                                    # iterate each GUI slider input
        pv = pv_objs.get(key)                                               # fetch the corresponding PF object (if found)
        if not pv:                                                          # skip keys that weren’t resolved (will be in 'missing')
            continue
        typ = PV_CONFIG.get(key, {}).get("type", "res")                     # look up configured site type; default to residential
        if typ == "res":                                                    # residential → use inverter count via npar
            _apply_residential_pv(pv, percent)
        elif typ == "com":                                                  # commercial → scale up to COM_MAX_KW by percent
            _apply_fixed_kw(pv, COM_MAX_KW * (max(0.0, 
            min(100.0, percent)) / 100.0))
        elif typ == "ind":                                                  # industrial → scale up to IND_MAX_KW by percent
            _apply_fixed_kw(pv, IND_MAX_KW * (max(0.0, 
            min(100.0, percent)) / 100.0))
        elif typ == "school":                                               # school → scale up to SCH_MAX_KW by percent
            _apply_fixed_kw(pv, SCH_MAX_KW * (max(0.0, 
            min(100.0, percent)) / 100.0))
    return missing                                                          # inform caller which pv_keys didn’t map to PF objects



# =====================================================================================================================================================
# =====================================================================================================================================================
# 4.0 ---------- Extraction utilities (series, min/max, energy) ----------  # helpers for processing time-series from PF
# =====================================================================================================================================================
# =====================================================================================================================================================

# 4.1 series_minmax: Given a [(t, u), ...] series, return dict with min/max values and their timestamps
def series_minmax(series):
    if not series:                                                          # empty / None → no data to process
        return None                                                         # signal absence of values
    vmin = min(series, key=lambda x: x[1])                                  # find the (t, u) pair with the smallest u
    vmax = max(series, key=lambda x: x[1])                                  # find the (t, u) pair with the largest u
    return {"u_min": vmin[1], "t_min_s": vmin[0], 
    "u_max": vmax[1], "t_max_s": vmax[0]}                                   # package results with value + time (seconds)

# 4.2 energy_kwh: Trapezoidal integrate a power series [(t, kW)] to get energy in kWh
def energy_kwh(series_kw):
    if not series_kw or len(series_kw) < 2:                                 # need at least two points to integrate
        return 0.0                                                          # no area under a single point
    e = 0.0                                                                 # accumulator for energy (kWh)
    for (t0, p0), (t1, p1) in zip(series_kw, series_kw[1:]):                # iterate consecutive point pairs
        dt_h = max(0.0, (t1 - t0) / 3600.0)                                 # time step in hours (guard against negative/unsorted)
        e += 0.5 * (p0 + p1) * dt_h                                         # trapezoid area: average power × duration
    return e                                                                # total energy in kWh

# 4.3 Convert MW series to kW if magnitudes look too small to be kW
def maybe_MW_to_kW(vals):                                                   # If median magnitude < 5, assume MW and convert to kW
    if not vals:                                                            # empty input → nothing to do
        return vals                                                         # return as-is
    mags = sorted(abs(v) for _, v in vals)                                  # absolute magnitudes of the values for robust median
    med = mags[len(mags) // 2]                                              # median magnitude
    if med < 5.0:                                                           # typical kW medians are >> 5; <5 suggests MW scale
        return [(t, v * 1000.0) for t, v in vals]                           # convert MW → kW by ×1000
    return vals                                                             # otherwise keep original units


# 4.4 Build a list of second-based timestamps for a 24 h day (inclusive end)
def _times_24h(dt_s: float = 3600.0):
    """
    Return [0, dt_s, 2*dt_s, ..., 24*3600] in seconds.
    Snapshot mode usually uses 3600 s (hourly), so callers often slice off the last item.
    """
    total = int(24 * 3600)                                                  # total simulation span in seconds (24 hours)
    steps = int(round(total / dt_s))                                        # number of uniform steps across the day
    return [i * dt_s for i in range(steps + 1)]                             # include the final 24h mark so _times_24h()[:-1] → 0..23 hours


# 4.5 Smooth daylight scalar (0..1) peaking near 12:00
def _gaussian_pv_scaler(hour: int) -> float:
    """
    Simple bell-shaped daylight multiplier. ~0 at night, ~1 at noon.
    Non-zero roughly 06:00–18:00 with a peak at ~12:00.
    """
    if hour < 6 or hour > 18:                                               # outside daylight window → zero PV
        return 0.0                                                          # return no generation
    sigma = 3.5                                                             # width of the bell (larger = flatter midday shoulder)
    x = (hour - 12.0) / sigma                                               # normalised distance from noon
    return float(math.exp(-0.5 * x * x))                                    # Gaussian curve value in [0,1]


# 4.6 Locate the Load Flow command (ComLdf) in the active Study Case
def _get_ldf(app):
    """
    Try common ways to obtain a ComLdf object. Raise if none is found.
    """
    ldf = None                                                              # will hold the found ComLdf command 
    try:                                                                    # Direct fetch from the Study Case (works when there is a single ComLdf)
        ldf = app.GetFromStudyCase("ComLdf")                                # ask the study case for a ComLdf object
    except Exception:
        ldf = None                                                          # ignore and try fallbacks
    if not ldf:                                                             # Fallback: search calculation-relevant objects by class pattern
        try:
            candidates = app.GetCalcRelevantObjects("*.ComLdf") or []       # look up any load-flow commands in scope
            if candidates:                                                  # if at least one exists
                ldf = candidates[0]                                         # pick the first (common single-command case)
        except Exception:
            ldf = None                                                      # still nothing; keep going

    if not ldf:                                                             # if no command was found at all
        raise RuntimeError("No Load Flow command (ComLdf) " \
        "found in the active Study Case.")                                  # bail with a clear message
    return ldf                                                              # return the command object ready to Execute()


# 4.7 Find and lightly configure a QDS command (ComInc/ComSim)
def _configure_qds(app):
    """
    Return a quasi-dynamic simulation command object if present.
    Tries ComInc first, then ComSim. Returns None if neither is available.
    """
    qds = None                                                              # will hold whichever QDS command we can find

    # Try incremental simulation command first (ComInc)
    try:
        qds = app.GetFromStudyCase("ComInc")                                # ask the study case for a ComInc
    except Exception:
        qds = None                                                          # ignore and try other options

    # Fallback: general simulation command (ComSim)
    if not qds:
        try:
            qds = app.GetFromStudyCase("ComSim")                            # ask the study case for a ComSim
        except Exception:
            qds = None                                                      # still nothing; try search

    # Last resort: search by class pattern amongst calc-relevant objects
    if not qds:
        try:
            for cls in ("*.ComInc", "*.ComSim"):                            # check both common QDS command classes
                objs = app.GetCalcRelevantObjects(cls) or []                # any instances available?
                if objs:
                    qds = objs[0]                                           # pick the first one found
                    break                                                   # stop searching once we’ve got one
        except Exception:
            qds = None                                                      # swallow and leave as None

    # Light-touch configuration; not all variants expose these attributes
    if qds:
        for attr, val in (("tstart", QDS_TSTART_S),                         # start time (seconds since midnight)
                          ("tstop",  QDS_TSTOP_S),                          # stop time (seconds since midnight)
                          ("dtgrid", QDS_DT_S)):                            # time step (seconds)
            try:
                setattr(qds, attr, val)                                     # set the attribute if it exists on this command
            except Exception:
                pass                                                        # safe to ignore — different templates expose different fields

    return qds                                                              # may be None; caller decides on snapshot fallback

# 4.8 ElmRes helpers — locate results table and read columns safely

def _get_elmres(app):
    """
    Return the first results table (ElmRes) in calc-relevant scope, or None.
    """
    try:
        res_list = app.GetCalcRelevantObjects("*.ElmRes") or []
        return res_list[0] if res_list else None
    except Exception:
        return None

def _elmres_series(res, col_name_candidates):
    """
    Try each candidate column name in ElmRes; return [(t, v), ...] if found, else [].
    Assumes column 0 is time (seconds). Tolerant to missing columns.
    """
    if not res:
        return []
    # Find a usable column index
    col_idx = -1
    for name in col_name_candidates:
        try:
            idx = res.FindColumn(name)  # returns -1 if not found
            if idx is not None and int(idx) >= 0:
                col_idx = int(idx)
                break
        except Exception:
            pass
    if col_idx < 0:
        return []

    # Read all rows: time = col 0, value = chosen column
    out = []
    try:
        nrows = int(res.GetRowCount() or 0)
    except Exception:
        nrows = 0

    for r in range(nrows):
        try:
            t = float(res.GetValue(r, 0))          # time column (seconds)
            v = res.GetValue(r, col_idx)           # the value we want
            if v is None:
                continue
            out.append((t, float(v)))
        except Exception:
            # ignore bad rows and keep going
            continue
    return out

# 4.9 Ensure QDS Variables set (auto builder)
def _ensure_qds_varset(app, qds, buses_map, net):
    """
    Make sure the QDS command has a Variables set with all channels we need,
    create/update it if required, and attach it to qds.p_resvar.
    """
    sc = app.GetActiveStudyCase()

    # Find or create the SetVar under the Study Case
    name = "QDS Variables (auto)"
    existing = sc.GetContents(f"{name}.SetVar")
    vset = existing[0] if existing else sc.CreateObject("SetVar", name)

    # Try to clear it (PF versions differ on the API name)
    for clear_fn in ("Clear", "Reset", "DeleteAll"):
        fn = getattr(vset, clear_fn, None)
        if fn:
            try:
                fn()
                break
            except Exception:
                pass

    # Helper to add one result channel, tolerating PF version differences
    def _add(vset_obj, pf_obj, var_name):
        for meth in ("Add", "AddVar", "AddVars", "AddVariable"):
            fn = getattr(vset_obj, meth, None)
            if fn:
                try:
                    # Most builds accept (object, "var"); some accept ("var", object)
                    try:
                        return fn(pf_obj, var_name)
                    except Exception:
                        return fn(var_name, pf_obj)
                except Exception:
                    pass
        # Last resort on some installs
        fn = getattr(vset_obj, "AddResVars", None)
        if fn:
            try:
                return fn(pf_obj, var_name)
            except Exception:
                pass
        # Don’t crash the run; just skip silently
        app.PrintWarn(f"6.4a ⚠️ Couldn’t add var '{var_name}' for {getattr(pf_obj, 'loc_name', pf_obj)}")

    # 4.9.1 Buses: per‑unit voltage (try m:u then m:u1)
    for _, term in buses_map.items():
        if not _add(vset, term, "m:u"):
            _add(vset, term, "m:u1")

    # 4.9.2 Lines & transformers: loading %
    for ln in net.get("lines", []):
        _add(vset, ln, "c:loading")
    for tr in net.get("trafos", []):
        _add(vset, tr, "c:loading")

    # 4.9.3 Loads & PVs: active power
    for ld in net.get("loads", []):
        if not _add(vset, ld, "m:Psum"):
            _add(vset, ld, "m:P")

    pv_objs, _ = get_pv_objects(app)
    for _key, pv in pv_objs.items():
        if not _add(vset, pv, "m:Psum"):
            _add(vset, pv, "m:P")

    # Attach the Variables set to the QDS command
    try:
        qds.p_resvar = vset
        app.PrintPlain(f"6.4a ✅ Attached Variables set: {vset.loc_name}")
    except Exception as e:
        app.PrintWarn(f"6.4a ⚠️ Couldn’t attach Variables set (p_resvar): {e}")

    return vset


# =====================================================================================================================================================
# =====================================================================================================================================================
# 5.0 ---------- Snapshot sweep (Fallback if QDS is playing silly buggers) 
# =====================================================================================================================================================
# =====================================================================================================================================================


# # 5.1 Run 24 hourly sim, gather voltages/loads/PV and store values
# def run_snapshot_24h(app):
#     """
#     Runs 24 hourly load flows and collects: bus pu, PV kW, loads kW, line/trafo loading %.
#     Uses a simple PV shape multiplier to emulate time-of-day if model lacks time-series.
#     """
#     ldf     = _get_ldf(app)                                                 # obtain a Load Flow command object for executing snapshots
#     buses   = get_bus_objects(app)                                          # map of bus name → terminal object for voltage reads
#     pv_objs, _ = get_pv_objects(app)                                        # pv_key → PV object (ignore missing list here)
#     net     = discover_network_elements(app)                                # discover lines/transformers/loads for loading & totals

# # 5.1.1 Storage for time-series that builds as sweep hours
#     bus_u      = {b: [] for b in buses}                                     # bus name → list of (t, pu) voltage tuples
#     pv_p_kw    = {k: [] for k in pv_objs}                                   # pv_key → list of (t, kW) output tuples
#     load_kw    = []                                                         # total load time-series (t, kW) across all loads
#     line_load  = {ln.loc_name: [] for ln in net["lines"]}                   # line name → list of (t, %) loading tuples
#     trafo_load = {tr.loc_name: [] for tr in net["trafos"]}                  # trafo name → list of (t, %) loading tuples

#     times = _times_24h()[:-1]                                               # 0..23 hours in seconds (exclude 24:00 end)
#     for idx, t in enumerate(times):                                         # iterate each hourly timestamp
#         hour = t // 3600                                                    # integer hour index (0..23)

# # 5.1.2 If model lacks PV dynamics, emulate a day profile with a smooth multiplier
#         day_mult = _gaussian_pv_scaler(hour)                                # scalar 0..1 shaping factor centred around midday

# # 5.1.3 Execute one snapshot load flow for this hour
#         ldf.Execute()                                                       # perform the network solution for current state  # :contentReference[oaicite:17]{index=17}

# # 5.1.4 Buses: capture voltage in per-unit
#         for name, term in buses.items():                                    # loop over every calc-relevant bus/terminal
#             try:
#                 u = term.GetAttribute("m:u")                                # read measured voltage (pu) attribute
#             except Exception:
#                 u = None                                                    # if attribute missing/unavailable, record nothing
#             if u is not None:
#                 bus_u[name].append((t, float(u)))                           # append (time, pu) to the bus’s series

# # 5.1.5 PV: capture active power (kW or MW depending on template), apply day profile for snapshot mode
#         for key, pv in pv_objs.items():                                     # iterate each PV object we resolved
#             p = pv.GetAttribute("m:Psum")                                   # try measured total active power first
#             if p is None:
#                 p = pv.GetAttribute("pgini")                                # fallback to initial setpoint if measured value absent
#             val = float(p) if p is not None else 0.0                        # ensure a numeric value
#             pv_p_kw[key].append((t, val * day_mult))                        # apply daylight multiplier and store (t, value)

# # 5.1.6 Loads: sum all ElmLod active powers
#         tot = 0.0                                                           # accumulator for total load
#         for ld in net["loads"]:                                             # iterate each load element
#             p = ld.GetAttribute("m:Psum")                                   # measured active power of the load
#             if p is None:
#                 continue                                                    # skip if attribute not available
#             tot += float(p)                                                 # add to the running total
#         load_kw.append((t, tot))                                            # store (t, total_kW_or_MW_depends)

# # 5.1.7 Assets: capture line/trafo loading percentages
#         for ln in net["lines"]:                                             # for each line
#             v = ln.GetAttribute("loading")                                  # loading percentage attribute
#             if v is not None:
#                 line_load[ln.loc_name].append((t, float(v)))                # store (t, %)
#         for tr in net["trafos"]:                                            # for each transformer
#             v = tr.GetAttribute("loading")                                  # loading percentage attribute
#             if v is not None:
#                 trafo_load[tr.loc_name].append((t, float(v)))               # store (t, %)

# # 5.1.8 Normalise PV/load units to kW (heuristic converter handles MW→kW)
#     for k, s in pv_p_kw.items():                                            # per PV series
#         pv_p_kw[k] = maybe_MW_to_kW(s)                                      # convert if series appears to be in MW
#     load_kw = maybe_MW_to_kW(load_kw)                                       # convert total load series if needed

# # 5.1.9 Assemble the structured report payload
#     report = {
#         "limits": {},                                                       # bus voltage min/max + OK flag per GUI suburb bus
#         "pv": {},                                                           # per-PV energy and peak information + series
#         "lines": {},                                                        # per-line maximum loading %
#         "trafos": {},                                                       # per-trafo maximum loading %
#         "load": {},                                                         # total load series + energy
#         "meta": {"mode": "snapshot", "dt_s": 3600.0}                        # metadata: mode and timestep used
#     }

# # 5.1.10 Compute bus min/max/OK flags (per GUI suburb bus)
#     for key, meta in PV_CONFIG.items():                                     # iterate GUI-configured suburbs
#         bus = meta["bus"]                                                   # the associated PF bus name
#         s = bus_u.get(bus, [])                                              # fetch that bus’s voltage series
#         mm = series_minmax(s)                                               # derive min/max and times
#         if mm:
#             mm["ok"] = (mm["u_min"] >= V_OK_MIN and 
#             mm["u_max"] <= V_OK_MAX)                                        # within limits → OK True/False
#             report["limits"][bus] = mm                                      # store under the bus name

# # 5.1.11 PV energy and peaks per pv_key
#     for key, s in pv_p_kw.items():                                          # per PV series
#         e = energy_kwh(s)                                                   # integrate to kWh over the day
#         pmax = max((v for _, v in s), default=0.0)                          # maximum instantaneous power across the series
#         report["pv"][key] = {"series_kW": s, "e_kWh": e, "p_kw_max": pmax}  # package PV metrics

# # 5.1.12 Line/trafo worst-case loading percentages
#     for name, s in line_load.items():                                       # per line series
#         if s:
#             report["lines"][name] = {"loading_max_pct": 
#                                      max(v for _, v in s)}                  # store the maximum observed %
#     for name, s in trafo_load.items():                                      # per transformer series
#         if s:
#             report["trafos"][name] = {"loading_max_pct": 
#                                       max(v for _, v in s)}                 # store the maximum observed %

# # 5.1.13 Total load series & daily energy
#     report["load"]["total_series_kW"] = load_kw                             # the unified total-load time-series (kW)
#     report["load"]["total_e_kWh"] = energy_kwh(load_kw)                     # integrated daily energy (kWh)

#     return report                                                           # hand back the complete snapshot-mode results bundle
# =====================================================================================================================================================
# =====================================================================================================================================================
# 5.0 ---------- Snapshot sweep (Fallback if QDS is playing silly buggers) 
# =====================================================================================================================================================
# =====================================================================================================================================================


# 5.1 Run 24 hourly sim, gather voltages/loads/PV and store values
def run_snapshot_24h(app):
    """
    Runs 24 hourly load flows and collects: bus pu, PV kW, loads kW, line/trafo loading %.
    Uses a simple PV shape multiplier to emulate time-of-day if model lacks time-series.
    """
    ldf     = _get_ldf(app)                                                 # obtain a Load Flow command object for executing snapshots
    buses   = get_bus_objects(app)                                          # map of bus name → terminal object for voltage reads
    pv_objs, _ = get_pv_objects(app)                                        # pv_key → PV object (ignore missing list here)
    net     = discover_network_elements(app)                                # discover lines/transformers/loads for loading & totals

    # 5.1.1 Storage for time-series that builds as sweep hours
    bus_u      = {b: [] for b in buses}                                     # bus name → list of (t, pu) voltage tuples
    pv_p_kw    = {k: [] for k in pv_objs}                                   # pv_key → list of (t, kW) output tuples
    load_kw    = []                                                         # total load time-series (t, kW) across all loads
    line_load  = {ln.loc_name: [] for ln in net["lines"]}                   # line name → list of (t, %) loading tuples
    trafo_load = {tr.loc_name: [] for tr in net["trafos"]}                  # trafo name → list of (t, %) loading tuples

    times = _times_24h()[:-1]                                               # 0..23 hours in seconds (exclude 24:00 end)
    for idx, t in enumerate(times):                                         # iterate each hourly timestamp
        hour = t // 3600                                                    # integer hour index (0..23)

        # 5.1.2 If model lacks PV dynamics, emulate a day profile with a smooth multiplier
        day_mult = _gaussian_pv_scaler(hour)                                # scalar 0..1 shaping factor centred around midday

        # 5.1.3 Execute one snapshot load flow for this hour
        ldf.Execute()                                                       # perform the network solution for current state

        # 5.1.4 Buses: capture voltage in per-unit
        for name, term in buses.items():                                    # loop over every calc-relevant bus/terminal
            # CHANGED: tolerant voltage read (tries measured + calculated variants)
            u = _get_attr(term, "m:u", "m:u1", "c:u")                       # NEW: robust voltage (pu) getter
            if u is not None:
                bus_u[name].append((t, float(u)))                           # append (time, pu) to the bus’s series

        # 5.1.5 PV: capture active power (kW or MW depending on template), apply day profile for snapshot mode
        for key, pv in pv_objs.items():                                     # iterate each PV object we resolved
            # CHANGED: tolerant PV power read (measured → setpoint → calculated)
            p = _get_attr(pv, "m:Psum", "pgini", "c:Psum")                  # NEW: robust PV active power getter
            val = float(p) if p is not None else 0.0                        # ensure a numeric value
            pv_p_kw[key].append((t, val * day_mult))                        # apply daylight multiplier and store (t, value)

        # 5.1.6 Loads: sum all ElmLod active powers
        tot = 0.0                                                           # accumulator for total load
        for ld in net["loads"]:                                             # iterate each load element
            # CHANGED: tolerant load power read (measured/calculated variants)
            p = _get_attr(ld, "m:Psum", "m:P", "c:Psum")                    # NEW: robust load active power getter
            if p is None:
                continue                                                    # skip if attribute not available
            tot += float(p)                                                 # add to the running total
        load_kw.append((t, tot))                                            # store (t, total_kW_or_MW_depends)

        # 5.1.7 Assets: capture line/trafo loading percentages
        for ln in net["lines"]:                                             # for each line
            # CHANGED: tolerant loading read (calculated → measured → bare)
            v = _get_attr(ln, "c:loading", "m:loading", "loading")          # NEW: robust line loading (%)
            if v is not None:
                line_load[ln.loc_name].append((t, float(v)))                # store (t, %)
        for tr in net["trafos"]:                                            # for each transformer
            # CHANGED: tolerant loading read (calculated → measured → bare)
            v = _get_attr(tr, "c:loading", "m:loading", "loading")          # NEW: robust transformer loading (%)
            if v is not None:
                trafo_load[tr.loc_name].append((t, float(v)))               # store (t, %)

    # 5.1.8 Normalise PV/load units to kW (heuristic converter handles MW→kW)
    for k, s in pv_p_kw.items():                                            # per PV series
        pv_p_kw[k] = maybe_MW_to_kW(s)                                      # convert if series appears to be in MW
    load_kw = maybe_MW_to_kW(load_kw)                                       # convert total load series if needed

    # 5.1.9 Assemble the structured report payload
    report = {
        "limits": {},                                                       # bus voltage min/max + OK flag per GUI suburb bus
        "pv": {},                                                           # per-PV energy and peak information + series
        "lines": {},                                                        # per-line maximum loading %
        "trafos": {},                                                       # per-trafo maximum loading %
        "load": {},                                                         # total load series + energy
        "meta": {"mode": "snapshot", "dt_s": 3600.0}                        # metadata: mode and timestep used
    }

# 5.1.10 Compute bus min/max/OK flags (per GUI suburb bus)
    # NOTE: use the resolver to map configured bus names to actual PF terminal names
    for key, meta in PV_CONFIG.items():                                     # iterate GUI-configured suburbs
        wanted = meta["bus"]                                                # the configured PF bus name
        real   = _resolve_bus_name(wanted, buses)                           # map to an actual terminal name if needed
        if not real:
            continue                                                        # nothing to record if we can't find a match

        s = bus_u.get(real, [])                                             # fetch that terminal’s voltage series
        mm = series_minmax(s)                                               # derive min/max and their times
        if mm:
            mm["ok"] = (mm["u_min"] >= V_OK_MIN and 
                        mm["u_max"] <= V_OK_MAX)                            # within limits → OK True/False
            report["limits"][real] = mm                                     # store under the RESOLVED terminal name


    # 5.1.11 PV energy and peaks per pv_key
    for key, s in pv_p_kw.items():                                          # per PV series
        e = energy_kwh(s)                                                   # integrate to kWh over the day
        pmax = max((v for _, v in s), default=0.0)                          # maximum instantaneous power across the series
        report["pv"][key] = {"series_kW": s, "e_kWh": e, "p_kw_max": pmax}  # package PV metrics

    # 5.1.12 Line/trafo worst-case loading percentages
    for name, s in line_load.items():                                       # per line series
        if s:
            report["lines"][name] = {"loading_max_pct": 
                                     max(v for _, v in s)}                  # store the maximum observed %
    for name, s in trafo_load.items():                                      # per transformer series
        if s:
            report["trafos"][name] = {"loading_max_pct": 
                                      max(v for _, v in s)}                 # store the maximum observed %

    # 5.1.13 Total load series & daily energy
    report["load"]["total_series_kW"] = load_kw                             # the unified total-load time-series (kW)
    report["load"]["total_e_kWh"] = energy_kwh(load_kw)                     # integrated daily energy (kWh)

    return report



# # =====================================================================================================================================================
# # =====================================================================================================================================================
# # 6.0 ---------- QDS path (if available) ----------  # attempt a quasi-dynamic simulation; fall back to snapshot post-processing
# # =====================================================================================================================================================
# # =====================================================================================================================================================

# # 6.1 run_qds_24h: Configure and run a 24-hour QDS sweep; if unsupported, raise so caller can use snapshot mode
# def run_qds_24h(app):
#     """
#     Attempt a quasi-dynamic 24h sweep. If no suitable object is present or ElmRes
#     writes no rows, raise to allow fallback to snapshot.
#     """
# # 6.1.1 Obtain a QDS command object (project dependent) or fail fast
#     qds = _configure_qds(app)                                               # helper should locate/prepare ComInc/ComSim (or equivalent) for QDS
#     if not qds:
#         raise RuntimeError("No QDS command object found (ComInc/ComSim).")  # let the caller decide to fall back

# # 6.1.2 Set the time grid on the QDS object where supported
#     for attr, val in (("tstart", QDS_TSTART_S), ("tstop", QDS_TSTOP_S), ("dtgrid", QDS_DT_S)):
#         try:
#             setattr(qds, attr, val)                                         # not all QDS variants expose these; ignore if they don’t
#         except Exception:
#             pass                                                            # keep going even if some attributes aren’t available

# # 6.1.3 Execute the QDS sweep (object type depends on your PF setup)
#     qds.Execute()                                                           # Run the sweep (object class depends on your setup)  # :contentReference[oaicite:18]{index=18}

# # 6.1.4 Post-processing: reuse robust snapshot readers to assemble the same report structure
#     # After QDS, read results directly from element attributes at final time,
#     # and also reconstruct time series via repeated Read of ElmRes if configured.
#     # Because QDS wiring varies widely, we’ll use the robust snapshot readers
#     # executed *after* QDS to collect the same signals across the 24h times.
#     return run_snapshot_24h(app)                                            # leverage existing logic to build the 'report' dict
# =====================================================================================================================================================
# =====================================================================================================================================================
# 6.0 ---------- QDS path (preferred & required) ----------  # run quasi-dynamic and read time-series from ElmRes
# =====================================================================================================================================================
# =====================================================================================================================================================

def run_qds_24h(app):
    """
    Run a 24-hour quasi-dynamic simulation and read results from ElmRes.
    Builds the same 'report' dict shape as snapshot mode:
      - report["limits"][<bus_name>] = {u_min, t_min_s, u_max, t_max_s, ok}
      - report["pv"][<pv_key>]       = {series_kW: [(t, kW)...], e_kWh, p_kw_max}
      - report["lines"][name]        = {loading_max_pct}
      - report["trafos"][name]       = {loading_max_pct}
      - report["load"]["total_series_kW"], report["load"]["total_e_kWh"]
    Raises RuntimeError if ElmRes is not found or has no usable columns.
    """

    # 6.1 Locate and lightly configure a QDS command
    qds = _configure_qds(app)                                               # ComInc/ComSim (project-dependent)
    if not qds:
        raise RuntimeError("No QDS command object found (ComInc/ComSim) in the active Study Case.")

    # 6.2 Execute the QDS sweep
    qds.Execute()                                                           # run the time sweep (00:00..24:00 in your settings)

    # 6.3 Fetch ElmRes (results table) — required for QDS time-series extraction
    res = _get_elmres(app)
    if not res:
        raise RuntimeError("No ElmRes results table found after QDS. Ensure your Study Case writes results to ElmRes.")

    # 6.4 Discover network objects for naming
    buses_map   = get_bus_objects(app)                                      # {loc_name: ElmTerm}
    pv_objs, _  = get_pv_objects(app)                                       # {pv_key: ElmPvsys/ElmGenstat}
    net         = discover_network_elements(app)                            # {"lines": [...], "trafos": [...], "loads": [...]}
    # 6.4a Ensure QDS Variables set (create/update + attach to qds.p_resvar)
    _ensure_qds_varset(app, qds, buses_map, net)

    # 6.5 Read series from ElmRes, assembling containers like snapshot mode
    bus_u      = {b: [] for b in buses_map}                                 # bus name -> [(t, pu)...]
    pv_p_kw    = {k: [] for k in pv_objs}                                   # pv_key -> [(t, kW_or_MW)...]
    line_load  = {ln.loc_name: [] for ln in net["lines"]}                   # line name -> [(t, %) ...]
    trafo_load = {tr.loc_name: [] for tr in net["trafos"]}                  # trafo name -> [(t, %) ...]
    load_total = []                                                         # [(t, kW_or_MW)...], summed across loads

    # 6.5.1 Voltages per bus — try common variants (:u, :u1, :ul)
    for bname in buses_map:
        series = _elmres_series(res, [f"{bname}:u", f"{bname}:u1", f"{bname}:ul", f"{bname}:c:u"])
        if series:
            bus_u[bname] = series

    # 6.5.2 PV power per pv_key — use the PV object's display name in ElmRes (loc_name)
    for pv_key, pv in pv_objs.items():
        pv_name = getattr(pv, "loc_name", pv_key)
        s = _elmres_series(res, [f"{pv_name}:Psum", f"{pv_name}:P", f"{pv_name}:c:Psum"])
        if s:
            pv_p_kw[pv_key] = s

    # 6.5.3 Line/transformer loading (%)
    for ln in net["lines"]:
        ln_name = ln.loc_name
        s = _elmres_series(res, [f"{ln_name}:loading", f"{ln_name}:c:loading"])
        if s:
            line_load[ln_name] = s

    for tr in net["trafos"]:
        tr_name = tr.loc_name
        s = _elmres_series(res, [f"{tr_name}:loading", f"{tr_name}:c:loading"])
        if s:
            trafo_load[tr_name] = s

    # 6.5.4 Total load series — sum all loads' P across time rows
    # We assume all load series share the same time vector (ElmRes row 0..N-1).
    # First, gather per-load series; then sum by row index.
    load_series_per_ld = []
    for ld in net["loads"]:
        ld_name = ld.loc_name
        s = _elmres_series(res, [f"{ld_name}:Psum", f"{ld_name}:P", f"{ld_name}:c:Psum"])
        if s:
            load_series_per_ld.append(s)

    if load_series_per_ld:
        # Use the time axis of the first load series; sum values at each index.
        base = load_series_per_ld[0]
        for i, (t, _) in enumerate(base):
            tot = 0.0
            for s in load_series_per_ld:
                if i < len(s) and abs(s[i][0] - t) < 1e-6:                  # align by index/time
                    tot += s[i][1]
            load_total.append((t, tot))

    # 6.6 Normalise PV/load units to kW (handles series that might be in MW)
    for k, s in list(pv_p_kw.items()):
        pv_p_kw[k] = maybe_MW_to_kW(s)
    load_total = maybe_MW_to_kW(load_total)

    # 6.7 Assemble the structured report payload
    report = {
        "limits": {},                                                       # bus voltage min/max + OK flag
        "pv": {},                                                           # per-PV energy/peaks + series
        "lines": {},                                                        # per-line max loading %
        "trafos": {},                                                       # per-trafo max loading %
        "load": {},                                                         # total load series + energy
        "meta": {"mode": "qds", "dt_s": QDS_DT_S}                           # record that this came from QDS
    }

    # 6.8 Compute bus min/max/OK using your resolver from Step 3
    for key, meta in PV_CONFIG.items():
        wanted = meta["bus"]
        real   = _resolve_bus_name(wanted, buses_map)                       # resolve configured name -> actual terminal
        if not real:
            continue
        s = bus_u.get(real, [])
        mm = series_minmax(s)
        if mm:
            mm["ok"] = (mm["u_min"] >= V_OK_MIN and mm["u_max"] <= V_OK_MAX)
            report["limits"][real] = mm

    # 6.9 PV energy and peaks per pv_key
    for key, s in pv_p_kw.items():
        if not s:
            continue
        e = energy_kwh(s)                                                   # kWh across the day
        pmax = max((v for _, v in s), default=0.0)                          # peak kW
        report["pv"][key] = {"series_kW": s, "e_kWh": e, "p_kw_max": pmax}

    # 6.10 Line/trafo worst-case loading percentages
    for name, s in line_load.items():
        if s:
            report["lines"][name] = {"loading_max_pct": max(v for _, v in s)}
    for name, s in trafo_load.items():
        if s:
            report["trafos"][name] = {"loading_max_pct": max(v for _, v in s)}

    # 6.11 Total load series & daily energy
    report["load"]["total_series_kW"] = load_total
    report["load"]["total_e_kWh"]     = energy_kwh(load_total)

    return report


# =====================================================================================================================================================
# =====================================================================================================================================================
# 7.0 ---------- Public entry point: set sliders, run, export PDF ----------  # main API used by the GUI
# =====================================================================================================================================================
# =====================================================================================================================================================


# # 7.1 Apply GUI slider values in PF, run 24 h analysis, write JSON/PDF, return results dict
# def set_penetrations_and_run(sliders, pdf_path=PDF_DEFAULT):
#     """
#     Main entry for GUI:
#       - set PV penetrations according to sliders {pv_key: percent}
#       - run 24-hour analysis (QDS first; snapshot fallback)
#       - return rich result dict (and write a PDF report)
#     """
#     app = connect_and_activate()                                            # ensure PowerFactory is running with a project open; get Application handle
#     _ = get_studycase(app)                                                  # verify a Study Case is active (raises if not); value unused, check only
#     missing = set_pv_penetrations(app, sliders)                             # push slider percentages into PF objects; returns any pv_keys not found

#     try:                                                                    # attempt preferred simulation path first
#         if USE_QDS_FIRST:                                                   # if configured to use quasi-dynamic simulation
#             results = run_qds_24h(app)                                      # run QDS for 24 h and post-process
#         else:                                                               # otherwise
#             results = run_snapshot_24h(app)                                 # do robust snapshot sweep instead
#     except Exception:                                                       # if QDS setup fails or execution errors,
#         results = run_snapshot_24h(app)                                     # fall back to snapshot mode to keep things resilient

#     results["missing_pv"] = list(missing)                                   # surface any PV objects not found so the GUI can warn the user

# 7.1 Apply GUI slider values in PF, run 24 h analysis, write JSON/PDF, return results dict
def set_penetrations_and_run(sliders, pdf_path=PDF_DEFAULT):
    """
    Main entry for GUI:
      - set PV penetrations according to sliders {pv_key: percent}
      - run 24-hour quasi-dynamic simulation (QDS only)
      - return rich result dict (and write a PDF report)
    """
    app = connect_and_activate()                                            # ensure PowerFactory is running with a project open; get Application handle
    _ = get_studycase(app)                                                  # verify a Study Case is active (raises if not); value unused, check only
    missing = set_pv_penetrations(app, sliders)                             # push slider percentages into PF objects; returns any pv_keys not found

    # --- Step 4 change: force QDS only, no snapshot fallback ---
    results = run_qds_24h(app)                                              # run quasi-dynamic simulation for 24 h

    results["missing_pv"] = list(missing)                                   # surface any PV objects not found so the GUI can warn the user
    return results


# # 7.1.1 Build legacy bus-keyed entries: { "<bus>": {"u_min":..., "u_max":..., "ok":...}, ... }
#     legacy = {}                                                             # temporary dict for legacy keys
#     for key, meta in PV_CONFIG.items():                                     # iterate configured GUI suburbs
#         bus = meta["bus"]                                                   # associated PF bus name
#         lim = results.get("limits", {}).get(bus)                            # look up computed limits for that bus
#         if lim:                                                             # if we have limits, copy out the essentials
#             legacy[bus] = {
#                 "u_min": lim.get("u_min"),
#                 "u_max": lim.get("u_max"),
#                 "ok":    lim.get("ok", True),
#             }

# # 7.1.2 Merge legacy keys into top-level results so old GUI code can still iterate results.items()
#     results.update(legacy)                                                  # adds/overwrites bus-named keys at the root of the dict
    # 7.1.1 Build legacy bus-keyed entries with RESOLVED terminal names
    try:
        # Build a fresh bus map here so the resolver has the same view
        buses_map = get_bus_objects(app)                                    # {loc_name: ElmTerm}
    except Exception:
        buses_map = {}

    legacy = {}                                                             # temporary dict for legacy keys
    for key, meta in PV_CONFIG.items():                                     # iterate configured GUI suburbs
        wanted = meta["bus"]                                                # configured bus (may be shorthand)
        real   = _resolve_bus_name(wanted, buses_map)                       # resolve to actual terminal name
        if not real:
            continue
        lim = results.get("limits", {}).get(real)                           # look up computed limits using RESOLVED name
        if lim:
            legacy[real] = {                                                # use RESOLVED bus name for legacy top-level key
                "u_min": lim.get("u_min"),
                "u_max": lim.get("u_max"),
                "ok":    lim.get("ok", True),
            }

    # 7.1.2 Merge legacy keys into top-level results so the GUI can still iterate results.items()
    results.update(legacy)                                                  # adds/overwrites bus-named keys at the root of the dict

# 7.1.3 Always include a JSON copy beside the PDF for programmatic reuse
    try:
        with open(os.path.splitext(pdf_path)[0] + ".json", 
                  "w", encoding="utf-8") as f:                              # write results next to the PDF
            json.dump(results, f, indent=2)                                 # pretty-print JSON for easy inspection
    except Exception:                                                       # file I/O failure shouldn’t crash the run
        pass

# 7.1.4 Export PDF report (best-effort; warn to console if it fails)
    try:
        export_report_pdf(results, pdf_path)                                # generate a human-readable PDF summary
    except Exception as e:
        # don't crash GUI if PDF export fails
        print(f"[WARN] PDF export failed: {e}")                             # log a warning; GUI remains responsive

# 7.1.5 Return the full results dict (contains 'limits', 'pv', 'lines', 'trafos', 'load', plus legacy bus keys)
    return results                                                          # GUI consumes this to update result labels/graphs


# $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
# =====================================================================================================================================================
# =====================================================================================================================================================
# 8.0 ---------- PDF export (matplotlib PdfPages) ----------  # generate a multi-page A4 report with tables and charts
# =====================================================================================================================================================
# =====================================================================================================================================================
# $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


# 8.1 export_report_pdf: Build and write the PDF report using matplotlib’s PdfPages
def export_report_pdf(results, pdf_path):
    """
    Multi-page PDF with summary tables and key charts.
    """
    import matplotlib.pyplot as plt                                         # plotting API for figures/axes
    from matplotlib.backends.backend_pdf import PdfPages                    # context manager to assemble multi-page PDFs

    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M")                       # timestamp string for the report header

# 8.1.1 fig_table: small helper to create a landscape A4 page with a centred table
    def fig_table(title, rows, col_labels):
        fig = plt.figure(figsize=(11.69, 8.27))                             # A4 landscape in inches (297×210 mm)
        fig.suptitle(title, fontsize=14)                                    # page title
        ax = fig.add_subplot(111)                                           # single full-page axes
        ax.axis("off")                                                      # hide axes frame/ticks
        table = ax.table(cellText=rows, 
                         colLabels=col_labels, loc="center")                # add the table widget
        table.auto_set_font_size(False)                                     # control font size manually for consistency
        table.set_fontsize(8)                                               # compact table font
        table.scale(1, 1.4)                                                 # increase row height slightly for readability
        return fig                                                          # return the fully built figure to be saved

# 8.1.2 Page 1 — Overview & voltage limits (per bus)
    rows = []                                                               # accumulate [bus, min, max, status] rows
    for key, meta in PV_CONFIG.items():                                     # iterate configured suburbs
        bus = meta["bus"]                                                   # PF bus name for this suburb
        lim = results.get("limits", {}).get(bus)                            # limits dict for that bus
        if not lim:
            continue                                                        # skip if not present in the results
        ok = "OK" if lim.get("ok") else "OUT"                               # text flag for status
        rows.append([bus, f"{lim['u_min']:.4f}", 
                     f"{lim['u_max']:.4f}", ok])                            # append a formatted row
    rows.sort(key=lambda r: r[0])                                           # sort rows alphabetically by bus name

    with PdfPages(pdf_path) as pdf:                                         # open the multipage PDF writer
        fig1 = fig_table(                                                   # build the first page table
            f"Ōtaki 24-h Summary — Voltage Limits (generated {ts})",
            rows,
            ["Bus", "Min pu", "Max pu", "Status"]
        )
        pdf.savefig(fig1); plt.close(fig1)                                  # write page then close the figure to free memory

# 8.1.3 Page 2 — Top line & transformer loading (bars) # Lines: pick the top 15 by max loading %
        line_items = sorted(                                                # (name, max_loading%) pairs sorted high→low
            ((name, d["loading_max_pct"]) for name, d in results.get("lines", {}).items()),
            key=lambda x: x[1], reverse=True
        )[:15]
        if line_items:                                                      # only draw if we have data
            names, vals = zip(*line_items)                                  # split names and values
            fig2 = plt.figure(figsize=(11.69, 8.27))                        # new A4 landscape figure
            ax2 = fig2.add_subplot(111)                                     # single axes
            ax2.set_title("Top Line Loading (%)")                           # chart title
            ax2.bar(names, vals)                                            # vertical bar chart
            ax2.set_xticklabels(names, rotation=60, ha="right")             # tilt labels to avoid overlap
            ax2.set_ylabel("%")                                             # y-axis label
            pdf.savefig(fig2); plt.close(fig2)                              # save and close

        # Transformers: pick the top 10 by max loading %
        trafo_items = sorted(                                               # (name, max_loading%) pairs sorted high→low
            ((name, d["loading_max_pct"]) for name, d in results.get("trafos", {}).items()),
            key=lambda x: x[1], reverse=True
        )[:10]
        if trafo_items:                                                     # only draw if we have data
            names, vals = zip(*trafo_items)                                 # split names and values
            fig3 = plt.figure(figsize=(11.69, 8.27))                        # new figure for trafos
            ax3 = fig3.add_subplot(111)                                     # single axes
            ax3.set_title("Top Transformer Loading (%)")                    # chart title
            ax3.bar(names, vals)                                            # bar chart
            ax3.set_xticklabels(names, rotation=60, ha="right")             # readable x labels
            ax3.set_ylabel("%")                                             # y-axis label
            pdf.savefig(fig3); plt.close(fig3)                              # save and close

# 8.1.4 Page 3 — PV production (total) and system load time-series,         # Build a unified time axis from the first PV series we find
        all_pv = results.get("pv", {})                                      # dict keyed by pv_key with series and metrics
        time_axis = None                                                    # will hold [t0, t1, ...] in seconds
        for d in all_pv.values():                                           # scan PV series
            s = d.get("series_kW", [])                                      # list of (t, kW) tuples
            if s:
                time_axis = [t for t, _ in s]                               # copy the time stamps
                break                                                       # stop after first non-empty series
        total_pv = []                                                       # will hold summed PV across all sites
        if time_axis:                                                       # only compute if we have a base axis
            for i, t in enumerate(time_axis):                               # walk the time indices
                tot = 0.0                                                   # running sum at this instant
                for d in all_pv.values():                                   # sum over every PV series
                    s = d.get("series_kW", [])
                    if i < len(s) and abs(s[i][0] - t) < 1e-6:              # align on index and check timestamp tolerance
                        tot += s[i][1]                                      # add kW at that instant
                total_pv.append((t, tot))                                   # store (t, total_kW)

        # Plot total PV vs hour
        if total_pv:
            fig4 = plt.figure(figsize=(11.69, 8.27))                        # new figure for PV time-series
            ax4 = fig4.add_subplot(111)                                     # single axes
            ax4.set_title("Total PV Production over 24h (kW)")              # chart title
            ax4.plot([t/3600.0 for t, _ in total_pv], 
                     [v for _, v in total_pv])                              # convert seconds→hours for x
            ax4.set_xlabel("Hour")                                          # x-axis label
            ax4.set_ylabel("kW")                                            # y-axis label
            pdf.savefig(fig4); plt.close(fig4)                              # save and close

        # Plot total system load vs hour
        load_series = results.get("load", {}).get("total_series_kW", [])    # list of (t, kW)
        if load_series:
            fig5 = plt.figure(figsize=(11.69, 8.27))                        # new figure for load time-series
            ax5 = fig5.add_subplot(111)                                     # single axes
            ax5.set_title("Total System Load over 24h (kW)")                # chart title
            ax5.plot([t/3600.0 for t, _ in load_series], 
                     [v for _, v in load_series])                           # time in hours
            ax5.set_xlabel("Hour")                                          # x-axis label
            ax5.set_ylabel("kW")                                            # y-axis label
            pdf.savefig(fig5); plt.close(fig5)                              # save and close

# 8.1.5 Page 4 — Energy summary table (per PV + totals)
        pv_rows = []                                                        # rows: [pv_key, energy_kWh, peak_kW]
        tot_pv_kwh = 0.0                                                    # accumulator for total PV energy
        for k, d in sorted(all_pv.items()):                                 # sorted for stable row order
            e = float(d.get("e_kWh", 0.0))                                  # energy (kWh) from results
            pmax = float(d.get("p_kw_max", 0.0))                            # peak kW from results
            pv_rows.append([k, f"{e:.1f}", f"{pmax:.1f}"])                  # add formatted row
            tot_pv_kwh += e                                                 # accumulate total PV energy
        load_kwh = float(results.get("load", {}).get("total_e_kWh", 0.0))   # total system load energy (kWh)
        pv_rows.append(["— Total PV —", f"{tot_pv_kwh:.1f}", ""])           # add a total PV row
        pv_rows.append(["— Total Load —", f"{load_kwh:.1f}", ""])           # add a total Load row

        fig6 = fig_table("Daily Energy (kWh) & PV Peaks (kW)", pv_rows, 
                         ["PV Key", "Energy (kWh)", "Peak (kW)"])           # build the table page
        pdf.savefig(fig6); plt.close(fig6)                                  # save the final page and close
