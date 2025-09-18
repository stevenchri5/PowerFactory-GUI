# otaki_sim.py
# 1.0 Set Up Environment
    # 1.1.0 Project and Study Case Names
    # 1.2.0 Quasi-Dynamic Simulation Timing Settings
    # 1.3.0 Print Results in Terminal
    # 1.4.0 PV panel wattage (per panel)
    # 1.5.0 Results Store and Monitor Registry
    # 1.6.0 PV inverter and panel overrides

# 2.0 Connect, Activate Project & Study Case
    # 2.1.0 Connect to PowerFactory
    # 2.2.0 Activate project
    # 2.3.0 Activate study case
    # 2.4.0 Refresh PV inverter overrides from model
    # 2.5.0 Refresh PV panels-per-inverter overrides from model
    # 2.6.0 get_inverter_counts
    # 2.7.0 get_panels_per_inverter

# 3.0 Quasi-Dynamic Simulation Setup
    # 3.2.0 prepare_quasi_dynamic (results file setup)
    # 3.3.0 run_quasi_dynamic (execute QDS)
    # 3.4.0 get_dynamic_results (read results)
    # 3.5.0 apply_pv_inverter_overrides
    # 3.6.0 apply_pv_panel_overrides

# 4.0 Quasi-Dynamic Simulation Core (GUI wrapper)
    # 4.1.0 build_monitored_dict
    # 4.2.0 extract_qds_results
        # 4.2.0 BUS Hourly p.u. voltage (24h)
        # 4.2.1 Bus 24h Min/Max Voltage p.u.
        # 4.2.2 LOAD Hourly demand (24h)
        # 4.2.3 PV Hourly Production (24h)
        # 4.2.4 PV Nameplate & Inverters
        # 4.2.5 TRANSFORMER Loading (24h)
        # 4.2.6 LINE Loading (24h)
        # 4.2.7 BUILD ASSOCIATIONS
    # 4.3.0 run_simulation (main entry)
    # 4.4.0 set_penetrations_and_run (GUI adapter)



#====================================================================================================
# 1.0  Set Up Environment
#====================================================================================================


import os, sys, time                                                                                # Import standard Python libraries for OS, system path, and timing
DIG_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP1\Python\3.9"                           # Path to PowerFactory Python API
sys.path.append(DIG_PATH)                                                                           # Add PowerFactory Python path to system path
os.environ['PATH'] += ';' + DIG_PATH                                                                # Append PowerFactory path to environment variables
import powerfactory as pf                                                                           # Import PowerFactory Python module


# 1.1.0 Project and Study Case Names-----------------------------------------------------------------
PROJECT_NAME    = "ENGR489 Otaki Grid Base Solar and Bat(1)"                                        # Name of the active PowerFactory project
STUDY_CASE_NAME = "Study Case"                                                                      # Name of the active study case


# 1.2.0 Quasi-Dynamic Simulation Timing Settings ----------------------------------------------------
QDS_STEP_SIZE   = 1                                                                                 # Step size for simulation (1 hour steps)
QDS_STEP_UNIT   = 2                                                                                 # Step unit: 0=seconds, 1=minutes, 2=hours, 3=days
QDS_CALC_PERIOD = 0                                                                                 # Simulation period: 0=full day, 12=12 hours, etc.


# 1.3.0 Print Results in Terminal ------------------------------------------------------------------
PRINT_BUS_HOURLY        = False                                                                     # Toggle printing hourly bus voltages
PRINT_LOAD_HOURLY       = False                                                                     # Toggle printing hourly load values
PRINT_PV_HOURLY         = False                                                                     # Toggle printing hourly PV generation
PRINT_TX_HOURLY         = False                                                                     # Toggle printing hourly transformer loadings
PRINT_BUS_MIN_MAX       = False                                                                     # Toggle printing bus min/max voltages
PRINT_LINE_HOURLY       = False                                                                     # Toggle printing hourly line loadings
PRINT_VARIABLE_CHECKS   = False                                                                     # Toggle printing debug variable checks
PRINT_PV_META           = False                                                                     # Toggle printing PV metadata
PRINT_PV_OVERRIDES      = False                                                                     # Toggle printing PV override information


# 1.4.0 PV panel wattage (per panel) ---------------------------------------------------------------
PANEL_WATT = 240.0                                                                                  # Define the power rating of each PV panel in watts


# 1.5.0 Results Store and Monitor Registry ---------------------------------------------------------
RESULTS   = {"bus": {}, "load": {}, "pv": {}, "tx": {},"line": {}}                                  # Dictionary to store simulation results for buses, loads, PV, transformers, and lines
ASSOC     = {}                                                                                      # Mapping of load names to their associated bus, PV, transformer, and line connections
MONITORED = []                                                                                      # List of monitored variables for GUI (each entry is a tuple: selector + list of variable names)


# 1.6.0 PV inverter and panel overrides (GUI sets before run) --------------------------------------
PV_INV_OVERRIDES   = {}                                                                             # Dictionary to store user-defined inverter overrides from GUI
PV_PANEL_OVERRIDES = {}                                                                             # Dictionary to store user-defined panel overrides from GUI


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


# 1.10.0 mapping for GUI use ------------------------------------------------------------------------
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
# 2.0  Connect, Activate Project & Study Case  (Working Perfectly, Don't fucking touch)
#====================================================================================================


#2.1 Connect to PowerFactory ------------------------------------------------------------------------
print("2.1.0     Connecting to PowerFactory…")                                                    # Print connection attempt message
app = pf.GetApplication()                                                                           # Get PowerFactory application object
if not app:                                                                                          # Check if connection failed
    print("2.1.0  Could not connect to PowerFactory."); sys.exit(1)                               # Exit if connection unsuccessful
print("2.1.0      Connected.")                                                                    # Print success message
print()                                                                                             # Print blank line for spacing


#2.2 Activate project-------------------------------------------------------------------------------
print(f"2.2.0     Activating project: {PROJECT_NAME}")                                            # Print project activation attempt
if app.ActivateProject(PROJECT_NAME) != 0:                                                          # Try activating project by name
    print(f"2.2.0  Could not activate project '{PROJECT_NAME}'."); sys.exit(1)                    # Exit if failed
print(f"2.2.0     Project activated: {PROJECT_NAME}")                                            # Print success message
print()                                                                                             # Print blank line for spacing


# 2.3 Activate study case --------------------------------------------------------------------------
print(f"2.3.0     Looking for study case: {STUDY_CASE_NAME}")                                     # Print study case search attempt
study_folder = app.GetProjectFolder("study")                                                        # Get study case folder from project
active_case = None                                                                                  # Initialize active case as None
for case in study_folder.GetContents("*.IntCase", 1):                                               # Loop through all study cases
    if case.loc_name == STUDY_CASE_NAME:                                                            # Check if case matches required name
        case.Activate()                                                                             # Activate matching study case
        active_case = case                                                                          # Save reference to active case
        break                                                                                       # Exit loop after activating
if not active_case:                                                                                 # If no matching study case found
    print(f"2.3.0  Study case '{STUDY_CASE_NAME}' not found."); sys.exit(1)                       # Exit with error
print(f"2.3.0      Study Case activated: {active_case.loc_name}")                                 # Print success message
print()                                                                                             # Print blank line for spacing


# 2.4.0 ▶    Refresh PV inverter overrides from model -----------------------------------------------
def refresh_pv_overrides_from_model(app):                                                           # Define function to refresh inverter overrides
    #print("2.4.0 ▶    Read current PV inverter counts.")                                          # Debug print (currently disabled)
    found = 0                                                                                       # Counter for how many PV entries found
    for code in PV_LIST:                                                                            # Loop through list of PV system codes
        for p in (app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []):                           # Get matching PV system objects
            PV_INV_OVERRIDES[p.loc_name] = int(getattr(p, "ngnum", 0) or 0)                         # Store inverter count override
            found += 1                                                                              # Increment counter
            if PRINT_PV_OVERRIDES:                                                                  # If debug printing enabled
                print(f"2.4.1      {p.loc_name}: ngnum={PV_INV_OVERRIDES[p.loc_name]}")           # Print inverter count
    #print(f"2.4.0      Refreshed {found} PV entries.")                                           # Debug print (disabled)
    #print()                                                                                        # Debug print (disabled)
refresh_pv_overrides_from_model(app)                                                                # Call function to load overrides into memory


# 2.5.0     Refresh PV panels-per-inverter overrides from model ------------------------------------
def refresh_pv_panels_from_model(app):                                                              # Define function to refresh panels-per-inverter
    #print("2.6.0     Read current PV panels per inverter.")                                      # Debug print (disabled)
    found = 0                                                                                       # Counter for PV entries
    for code in PV_LIST:                                                                            # Loop through list of PV system codes
        for p in (app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []):                           # Get matching PV system objects
            PV_PANEL_OVERRIDES[p.loc_name] = int(getattr(p, "npnum", 0) or 0)                       # Store panel count override
            found += 1                                                                              # Increment counter
            if PRINT_PV_OVERRIDES:                                                                  # If debug printing enabled
                print(f"2.6.1     {p.loc_name}: npnum={PV_PANEL_OVERRIDES[p.loc_name]}")         # Print panel count
    #print(f"2.6.0      Refreshed {found} PV entries.")                                           # Debug print (disabled)
    #print()                                                                                        # Debug print (disabled)
refresh_pv_panels_from_model(app)                                                                   # Call function to load panel overrides into memory


# 2.6.0 get_inverter_counts -------------------------------------------------------------------------
def get_inverter_counts():                                                                          # Define function to retrieve inverter counts
    global RESULTS                                                                                  # Use global RESULTS dictionary
    RESULTS.setdefault("pv_meta", {})                                                               # Ensure "pv_meta" key exists in RESULTS
    pv_meta = RESULTS["pv_meta"]                                                                    # Reference pv_meta dictionary
    pv_meta.clear()                                                                                 # Clear existing entries
    for pv_key, meta in PV_CONFIG.items():                                                          # Loop through PV configuration entries
        homes = int(meta.get("homes", 0))                                                           # Get number of homes with inverters
        pv_meta[pv_key] = {"inverters": homes}                                                      # Store inverter count in results
        print(f"[otaki_sim] {pv_key}: inverters={homes}")                                           # Print inverter info to terminal
    return pv_meta                                                                                  # Return inverter metadata dictionary


# 2.7.0 get_panels_per_inverter ---------------------------------------------------------------------
def get_panels_per_inverter():                                                                      # Define function to retrieve panels per inverter
    global RESULTS                                                                                  # Use global RESULTS dictionary
    RESULTS.setdefault("pv_meta", {})                                                               # Ensure "pv_meta" key exists in RESULTS
    pv_meta = RESULTS["pv_meta"]                                                                    # Reference pv_meta dictionary
    pv_meta.clear()                                                                                 # Clear existing entries
    for pv_key, meta in PV_CONFIG.items():                                                          # Loop through PV configuration entries
        nmods = int(PV_PANEL_OVERRIDES.get(pv_key, 0))                                              # Get number of panels per inverter
        kw = round((nmods * PANEL_WATT)/1000, 3)                                                    # Convert panels to kW per inverter (3 decimal places)
        pv_meta[pv_key] = {"panels_per_inv": nmods, "kw_per_inv": kw}                               # Store results
        if PRINT_PV_OVERRIDES:                                                                      # If debug printing enabled
            print(f"[otaki_sim] {pv_key}: panels={nmods}, kw/inv={kw}")                             # Print panel/inverter info
    return pv_meta                                                                                  # Return updated metadata dictionary


#====================================================================================================
# 3.0  Quasi-Dynamic Simulation Setup
#====================================================================================================


# 3.2 Results file setup ----------------------------------------------------------------------------
def prepare_quasi_dynamic(app, monitored_vars,
                          step_size=QDS_STEP_SIZE,
                          step_unit=QDS_STEP_UNIT,
                          period=QDS_CALC_PERIOD):                                                # Function to set up quasi-dynamic simulation results
    print("3.1.0     Results Setup Begin.")                                                     # Print start of results setup
    qds = app.GetFromStudyCase("ComStatsim")                                                      # Get quasi-dynamic simulation object
    if not qds:                                                                                   # Check if missing
        raise RuntimeError("3.1.0  [Results Setup] ComStatsim not found.")                       # Raise error if not found
    res = qds.results                                                                              # Get results object from QDS
    if not res:                                                                                    # Check if missing
        raise RuntimeError("3.1.0  [Results Setup] QDS results object missing.")                 # Raise error if missing
    res.Clear()                                                                                    # Clear existing results
    print("3.1.0      Results Setup ready.")                                                     # Print success
    print()                                                                                        # Blank line

    print("3.1.2     Adding Variables to Check.")                                                # Print variable check start
    for sel, var_list in monitored_vars.items():                                                   # Loop through monitored selectors
        elems = app.GetCalcRelevantObjects(sel) or []                                              # Get matching elements
        for e in elems:                                                                            # Loop through each element
            res.AddVars(e, *var_list)                                                              # Add variables to results
            if PRINT_VARIABLE_CHECKS:                                                              # If variable check printing enabled
                print(f"3.1.2     Checking Variables {e.loc_name}: {var_list}")                   # Print variable list
    print("3.1.2      Variable Checking Done.")                                                  # Print success
    print()                                                                                        # Blank line

    print("3.1.3     Configure Quasi-Dynamic Simulation Timing.")                                # Print timing setup start
    qds.stepSize   = step_size                                                                     # Set simulation step size
    qds.stepUnit   = step_unit                                                                     # Set step unit
    qds.calcPeriod = period                                                                        # Set simulation period
    print(f"3.1.3      Timing: period={period}, step={step_size}, unit={step_unit}")              # Print configured timing

    print("3.1.4      Quasi-Dynamic Simulation Ready to Run.")                                    # Print ready message
    print()                                                                                        # Blank line
    return res, qds                                                                                # Return results and QDS objects


# 3.3 Execute QDS — runs the quasi-dynamic simulation ----------------------------------------------
def run_quasi_dynamic(qds):                                                                        # Function to execute quasi-dynamic simulation
    print("3.2.0     Quasi-Dynamic Simulation Running…")                                         # Print start message
    ok = (qds.Execute() == 0)                                                                      # Execute simulation and check success
    print("3.2.0      Quasi-Dynamic Simulation Finished." if ok else                             # Print success message
          "3.2.0  Quasi-Dynamic Simulation Failed.")                                             # Or failure message
    print()                                                                                        # Blank line
    return ok                                                                                      # Return success flag


# 3.4 Read results — loads time series for one element/variable -------------------------------------
def get_dynamic_results(app, res, elm_selector, var_name, verbose=True):                           # Function to extract results for one element/variable
    elems = app.GetCalcRelevantObjects(elm_selector) or []                                         # Get matching elements
    name = elems[0].loc_name if elems else elm_selector                                            # Get name or fallback to selector
    if verbose:                                                                                    # If verbose output enabled
        print(f"3.3.0     Results For {name}")                                                   # Print header
    if not elems:                                                                                  # If no elements found
        if verbose: print(f"3.3.0  [Read Results] Element not found for '{elm_selector}' ({var_name}).")  # Warn missing element
        return [], []                                                                              # Return empty results
    elm = elems[0]                                                                                 # Take first element
    app.ResLoadData(res)                                                                           # Load result data
    col = app.ResGetIndex(res, elm, var_name)                                                      # Get index of requested variable
    if col < 0:                                                                                    # If variable not found
        if verbose: print(f"3.3.0  [Read Results] Var '{var_name}' not in results for {name}.")   # Warn missing variable
        return [], []                                                                              # Return empty results
    n = app.ResGetValueCount(res, 0)                                                               # Get number of data rows
    if verbose:                                                                                    # If verbose output enabled
        print(f"3.3.0  Reading Results For {name}, Variable={var_name}, Rows={n}")                # Print row count
    t, v = [], []                                                                                  # Initialize time and value lists
    for i in range(n):                                                                             # Loop through rows
        t.append(app.ResGetData(res, i, -1)[1])                                                    # Append time value
        v.append(app.ResGetData(res, i, col)[1])                                                   # Append variable value
    if verbose:                                                                                    # If verbose output enabled
        print(f"3.3.0     Results For {name} Extracted.")                                        # Print extraction success
    return t, v                                                                                    # Return time and value lists


# 3.5 Apply PV inverter overrides (set ElmPvsys.ngnum) ----------------------------------------------
def apply_pv_inverter_overrides(app, overrides):                                                    # Function to apply PV inverter overrides
    print("3.4.0     Apply PV Inverter Overrides.")                                               # Print start message
    if not overrides:                                                                               # If no overrides provided
        print("3.4.0      No overrides provided."); return                                         # Print message and exit
    for code, count in overrides.items():                                                           # Loop through overrides dictionary
        pvs = app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []                                # Find matching PV system objects
        if not pvs:                                                                                 # If none found
            print(f"3.4.0     No PV matched '{code}'"); continue                                  # Print warning and continue
        for p in pvs:                                                                               # Loop through each PV object
            old = getattr(p, "ngnum", None)                                                         # Get old inverter count
            try:                                                                                    # Attempt to update
                p.ngnum = int(count)                                                                # Set new inverter count
                if PRINT_PV_OVERRIDES:                                                              # If debug printing enabled
                    print(f"3.4.4      {p.loc_name}: ngnum {old} → {p.ngnum}")                    # Print change applied
            except Exception as e:                                                                  # Catch errors
                print(f"3.4.0     {p.loc_name}: failed to set ngnum: {e}")                         # Print failure message
    print("3.4.0      Overrides applied.")                                                         # Print completion message
    print()                                                                                         # Blank line


# 3.6 Apply PV panel overrides (set ElmPvsys.nPnum / npnum) -----------------------------------------
def apply_pv_panel_overrides(app, overrides):                                                       # Function to apply PV panel-per-inverter overrides
    print("3.6.0     Apply PV Panel-Per-Inverter Overrides.")                                     # Print start message
    if not overrides:                                                                               # If no overrides provided
        print("3.6.0      No overrides provided."); print(); return                                # Print message and exit
    for code, nmods in overrides.items():                                                           # Loop through overrides dictionary
        pvs = app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []                                # Find matching PV system objects
        if not pvs:                                                                                 # If none found
            print(f"3.6.0     No PV matched '{code}'"); continue                                  # Print warning and continue
        for p in pvs:                                                                               # Loop through each PV object
            old = int(getattr(p, "npnum", getattr(p, "nPnum", 0)) or 0)                             # Get old panel count
            try:                                                                                    # Attempt to update
                if hasattr(p, "npnum"):   p.npnum = int(nmods)                                      # Update npnum if attribute exists
                elif hasattr(p, "nPnum"): p.nPnum = int(nmods)                                      # Update nPnum if attribute exists
                else: raise AttributeError("no nPnum/npnum attr")                                   # Raise error if neither exists
                if PRINT_PV_OVERRIDES:                                                              # If debug printing enabled
                    print(f"3.6.4      {p.loc_name}: nPnum {old} → {int(nmods)}")                 # Print change applied
            except Exception as e:                                                                  # Catch errors
                print(f"3.6.0     {p.loc_name}: failed to set nPnum: {e}")                         # Print failure message
    print("3.6.0      Panel overrides applied."); print()                                          # Print completion message and blank line


#====================================================================================================
# 4.0  Quasi-Dynamic Simulation Core (Wrapped for GUI Use)
#====================================================================================================


# 4.1 Build full monitored variable dict ------------------------------------------------------------
def build_monitored_dict():                                                                         # Function to build monitored variable dictionary
    print("4.0.0     Quasi-Dynamic Simulation Setup.")                                            # Print setup start
    bus_selectors  = {f"{bus}.ElmTerm": ["m:u1"] for bus in BUS_LIST}                              # p.u. voltage selectors
    load_selectors = {f"{code}*.ElmLod": ["m:P:bus1"] for code in LOAD_LIST}                       # kW demand selectors
    pv_selectors   = {f"{code}*.ElmPvsys": ["m:Psum:bus1"] for code in PV_LIST}                    # PV output selectors
    tx_selectors   = {f"{code}*.ElmTr2": ["m:loading:bus1"] for code in TX_LIST}                   # % loading selectors
    line_selectors = {f"{name}.ElmLne":   ["m:loading:bus1"] for name in LINE_LIST}                # Line loading selectors

    monitored = {}                                                                                  # Initialize monitored dict
    monitored.update(bus_selectors)                                                                 # Add bus monitors
    monitored.update(load_selectors)                                                                # Add load monitors
    monitored.update(pv_selectors)                                                                  # Add PV monitors
    monitored.update(tx_selectors)                                                                  # Add transformer monitors
    monitored.update(line_selectors)                                                                # Add line monitors

    MONITORED[:] = [(k, v) for k, v in monitored.items()]                                           # Update global monitored list for GUI
    print(f"4.0.3     Register Monitors {len(MONITORED)} entries. ------------")                  # Print number of monitors registered
    print()                                                                                         # Blank line
    return monitored                                                                                # Return monitored dict


# 4.2 Extract all results ----------------------------------------------------------------------------
def extract_qds_results(app, res):                                                                  # Function to extract QDS results

    # 4.2 BUS Hourly p.u. voltage (24hrs) -----------------------------------------------------------
    print("4.2.0      Hourly p.u. Voltages Begin.")                                               # Print start
    for bus in BUS_LIST:                                                                            # Loop through buses
        sel = f"{bus}.ElmTerm"                                                                      # Selector string for bus
        t, u = get_dynamic_results(app, res, sel, "m:u1", verbose=PRINT_BUS_HOURLY)                 # Get bus voltage results
        if not u:                                                                                   # If no results found
            print(f"4.2.0  No voltage data for {bus}")                                            # Warn missing data
            continue                                                                                # Skip to next bus
        RESULTS["bus"][bus] = {"t": t[:24], "u_pu": u[:24]}                                         # Store first 24 hours of results
        if PRINT_BUS_HOURLY:                                                                        # If debug printing enabled
            for hr in range(min(24, len(u))):                                                       # Loop through 24 hours
                print(f"{hr:02d} {bus} = {u[hr]:.4f} p.u")                                          # Print voltage per hour
    print("4.2.0      Hourly p.u. Voltages Done. ------------------------------")                  # Print completion
    print()                                                                                         # Blank line


    # 4.2.1 Bus 24hr Min/Max Voltage p.u ------------------------------------------------------------
    print("4.2.1     Bus 24hr Min/Max p.u Begin.")                                                # Print start
    for bus in BUS_LIST:                                                                            # Loop through buses
        rec = RESULTS["bus"].get(bus)                                                               # Get recorded bus data
        if not rec:                                                                                 # If no record
            print(f"4.2.1  No data for {bus}"); continue                                          # Warn and continue
        u = rec.get("u_pu", [])[:24]                                                                # Extract first 24h voltage list
        if not u:                                                                                   # If empty list
            print(f"4.2.1  Empty voltage list for {bus}"); continue                               # Warn and continue
        u_min = min(u); h_min = u.index(u_min)                                                      # Find minimum and hour
        u_max = max(u); h_max = u.index(u_max)                                                      # Find maximum and hour
        rec.update({                                                                                # Update record with min/max values
            "u_pu_min": u_min, "u_pu_min_hour": h_min,
            "u_pu_max": u_max, "u_pu_max_hour": h_max
        })
        if PRINT_BUS_MIN_MAX:                                                                       # If debug printing enabled
            print(f"4.2.1 📌 {bus} min={u_min:.4f} @ {h_min:02d}h, max={u_max:.4f} @ {h_max:02d}h") # Print results
    print("4.2.1      Bus 24hr Min/Max p.u Done. -------------------------------")                 # Print completion
    print()                                                                                         # Blank line


    # 4.2.2 LOAD Hourly demand (24hrs) --------------------------------------------------------------
    print("4.2.2    Hourly Load Demand Begin.")                                                   # Print start
    for load in LOAD_LIST:                                                                          # Loop through load groups
        sel = f"{load}*.ElmLod"                                                                     # Selector for load
        loads = app.GetCalcRelevantObjects(sel) or []                                               # Get matching loads
        if not loads:                                                                               # If no matches
            print(f"4.2.2  No loads matched '{load}'"); continue                                  # Warn and continue
        for ld in loads:                                                                            # Loop through each load
            tP, P = get_dynamic_results(app, res, ld.loc_name + ".ElmLod", "m:P:bus1", verbose=PRINT_LOAD_HOURLY)  # Get load demand
            if not P:                                                                               # If no demand data
                print(f"4.2.2  No demand data for {ld.loc_name}"); continue                       # Warn and continue
            RESULTS["load"][ld.loc_name] = {"t": tP[:24], "P_W": P[:24]}                            # Store first 24h demand
            if PRINT_LOAD_HOURLY:                                                                   # If debug printing enabled
                for hr in range(min(24, len(P))):                                                   # Loop through 24h
                    print(f"{hr:02d} {ld.loc_name} = {P[hr]:.4f} kW")                               # Print hourly demand
                print()                                                                             # Extra blank line for spacing
    print("4.2.2      Hourly Load Demand Done. ---------------------------------")                # Print completion
    print()                                                                                         # Blank line


# 4.2.3 PV Hourly Production (24hrs) ---------------------------------------------------------------
    print("4.2.3    Hourly PV Array Production Begin.")                                          # Print start
    for pv in PV_LIST:                                                                             # Loop through PV groups
        pvs = app.GetCalcRelevantObjects(f"{pv}*.ElmPvsys") or []                                 # Get matching PV system objects
        if not pvs:                                                                                # If no PV found
            print(f"4.2.3  No PV objects found for '{pv}'"); continue                           # Warn and continue
        for p in pvs:                                                                              # Loop through PV objects
            sel = f"{p.loc_name}.ElmPvsys"                                                         # Selector string
            tP, P = get_dynamic_results(app, res, sel, "m:Psum:bus1", verbose=PRINT_PV_HOURLY)     # Get PV production results
            #print(f"[DEBUG] {sel} — Raw Psum values:")                                            # Debug print (disabled)
            #for i in range(min(24, len(P))):                                                      # Debug loop (disabled)
                #print(f"[DEBUG] Object: {p}, Class: {p.GetClassName()}, Name: {p.loc_name}")      # Debug info (disabled)

            if not P:                                                                              # If no production data
                print(f"4.2.3  No production data for {p.loc_name}"); continue                   # Warn and continue
            RESULTS["pv"][p.loc_name] = {"t": tP[:24], "P_W": P[:24]}                              # Store first 24h PV production
            #print(f"[DEBUG] Raw PV values for {p.loc_name}:", P[:24])                             # Debug print (disabled)
            if PRINT_PV_HOURLY:                                                                    # If debug printing enabled
                for hr in range(min(24, len(P))):                                                  # Loop through 24h
                    print(f"{hr:02d} {p.loc_name} = {P[hr]:.2f} kW")                               # Print PV production
                print()                                                                            # Extra blank line
    print("4.2.3      Hourly PV Production Done. --------------------------------")               # Print completion
    print()                                                                                        # Blank line


# 4.2.4 PV Nameplate & Inverters -------------------------------------------------------------------
    print("4.2.4     PV Variables.")                                                             # Print start
    if "pv_meta" not in RESULTS: RESULTS["pv_meta"] = {}                                           # Ensure pv_meta exists in results
    for code in PV_LIST:                                                                           # Loop through PV groups
        pvs = app.GetCalcRelevantObjects(f"{code}*.ElmPvsys") or []                               # Get PV system objects
        if not pvs:                                                                                # If none found
            print(f"4.2.4    No PV Found '{code}'"); continue                                    # Warn and continue
        for p in pvs:                                                                              # Loop through PV systems
            inv_objs = app.GetCalcRelevantObjects(p.loc_name + ".ElmInv") or []                    # Get inverter objects
            inv_count = len(inv_objs) if inv_objs else int(getattr(p, "ngnum", 0) or 0)            # Get inverter count
            par_mods = int(getattr(p, "npnum", 0) or 0)                                            # Get panels per inverter
            rating_kw_calc = (inv_count * par_mods * float(PANEL_WATT)) / 1000.0                   # Calculate rating in kW
            RESULTS["pv_meta"][p.loc_name] = {                                                     # Store metadata
                "rating_kW_calc": rating_kw_calc,
                "inverters": inv_count,
                "panels_per_inverter": par_mods,
            }
            if PRINT_PV_META:                                                                      # If debug printing enabled
                prefix = f"4.4.1      {p.loc_name}: "                                             # Prefix for formatting
                pad = " " * len(prefix)                                                            # Indentation pad
                print(f"{prefix}Nameplate Rating = {rating_kw_calc:.2f} kW,")                      # Print rating
                print(f"{pad} Inverters in Parallel = {inv_count},")                               # Print inverter count
                print(f"{pad} Panels per inverter = {par_mods}")                                   # Print panel count
    print("4.2.4      PV Nameplate & Inverters Done. ----------------------------")               # Print completion
    print()                                                                                        # Blank line


# 4.2.5 TRANSFORMER Loading ------------------------------------------------------------------------
    print("4.2.5    Hourly Tx Loading Begin.")                                                   # Print start
    for tx in TX_LIST:                                                                             # Loop through transformers
        tx_objs = app.GetCalcRelevantObjects(f"{tx}*.ElmTr2") or []                               # Get transformer objects
        if not tx_objs:                                                                            # If none found
            print(f"4.2.5  No Transformer Found '{tx}'"); continue                               # Warn and continue
        for t in tx_objs:                                                                          # Loop through transformers
            tP, P = get_dynamic_results(app, res, t.loc_name + 
            ".ElmTr2", "m:loading:bus1", verbose=PRINT_TX_HOURLY)                                  # Get transformer loading
            if not P:                                                                              # If no data
                print(f"4.2.5  No Loading Data for {t.loc_name}"); continue                      # Warn and continue
            RESULTS["tx"][t.loc_name] = {"t": tP[:24], "loading_pct": P[:24]}                      # Store 24h loading data
            if PRINT_TX_HOURLY:                                                                    # If debug printing enabled
                for hr in range(min(24, len(P))):                                                  # Loop through 24h
                    print(f"{hr:02d} {t.loc_name} = {P[hr]:.2f} %")                                # Print hourly loading
                print()                                                                            # Extra blank line
    print("4.2.5      Hourly Tx Loading Done. ------------------------------------")              # Print completion
    print()                                                                                        # Blank line


# 4.2.6 LINE Loading --------------------------------------------------------------------------------
    print("4.2.6     Hourly Line Loading Begin.")                                                # Print start
    if "line" not in RESULTS: RESULTS["line"] = {}                                                 # Ensure "line" key exists in results
    for line in LINE_LIST:                                                                         # Loop through line list
        lines = app.GetCalcRelevantObjects(f"{line}.ElmLne") or []                                 # Get matching line objects
        if not lines:                                                                              # If no matches
            print(f"4.2.6  No Line Found '{line}'"); continue                                    # Warn and continue
        for ln in lines:                                                                           # Loop through lines
            tL, L = get_dynamic_results(app, res, ln.loc_name + ".ElmLne", "m:loading:bus1", verbose=PRINT_LINE_HOURLY)  # Get line loading
            if not L:                                                                              # If no data
                print(f"4.2.6  No Loading Data for Line {ln.loc_name}"); continue                # Warn and continue
            RESULTS["line"][ln.loc_name] = {"t": tL[:24], "loading_pct": L[:24]}                   # Store 24h line loading data
            if PRINT_LINE_HOURLY:                                                                  # If debug printing enabled
                for hr in range(min(24, len(L))):                                                  # Loop through 24h
                    print(f"{hr:02d} {ln.loc_name} Loading = {L[hr]:.2f} %")                       # Print hourly loading
                print()                                                                            # Extra blank line
    print("4.2.6      Hourly Line Loading Done. ----------------------------------")              # Print completion
    print()                                                                                        # Blank line


# 4.2.7 BUILD ASSOCIATIONS --------------------------------------------------------------------------
    print("4.2.7     Build Associations (using PV_CONFIG).")                                     # Print start
    ASSOC.clear()                                                                                  # Clear associations dict

    for pv_key, cfg in PV_CONFIG.items():                                                          # Loop through PV config
        load_name = cfg.get("load")                                                                # Get load name
        if not load_name:                                                                          # Skip if no load
            continue

        if load_name not in ASSOC:                                                                 # If load not yet mapped
            ASSOC[load_name] = {                                                                   # Initialize mapping
                "bus": cfg.get("bus"),
                "pv": [],
                "tx": [],
                "line": [],
            }

        ASSOC[load_name]["pv"].append(pv_key)                                                      # Add PV to load mapping
        if cfg.get("tx"):                                                                          # If transformer exists
            ASSOC[load_name]["tx"].append(cfg["tx"])                                               # Add transformer to mapping
        if cfg.get("pline"):                                                                       # If line exists
            ASSOC[load_name]["line"].append(cfg["pline"])                                          # Add line to mapping

    print(f"4.2.7      Build Associations: {len(ASSOC)} loads mapped to bus/pv/tx/line.")         # Print mapping result
    print()                                                                                        # Blank line


# 4.3 MAIN ENTRY POINT -------------------------------------------------------------------------------
def run_simulation(pv_overrides=None):                                                             # Main entry point for QDS simulation
    """
    Executes the full QDS simulation. Accepts optional PV inverter overrides.
    Updates global RESULTS and ASSOC. Returns True if OK, False otherwise.
    """
    if pv_overrides:                                                                               # If overrides provided
        apply_pv_inverter_overrides(app, pv_overrides)                                             # Apply inverter overrides
        apply_pv_panel_overrides(app, PV_PANEL_OVERRIDES)                                          # Apply panel overrides
    monitored = build_monitored_dict()                                                             # Build monitored dictionary
    res, qds = prepare_quasi_dynamic(app, monitored)                                               # Prepare QDS

    print("4.1.0     Execute QDS Begin.")                                                        # Print execution start
    ok = run_quasi_dynamic(qds)                                                                    # Run QDS
    print(f"4.1.0      Execute QDS Status = {ok} -------------------------------")                # Print execution status
    print()                                                                                        # Blank line
    if not ok:                                                                                     # If QDS failed
        return False                                                                               # Return False

    extract_qds_results(app, res)                                                                  # Extract results
    return True                                                                                    # Return success flag


# 4.4 set_penetrations_and_run — adapter for GUI -----------------------------------------------------
def set_penetrations_and_run(pv_overrides=None):                                                   # Adapter for GUI Run button / threads
    # 4.4.1  Apply overrides here (SIM touches PF; GUI does not)
    try:                                                                                           # Attempt inverter overrides
        apply_pv_inverter_overrides(app, pv_overrides or PV_INV_OVERRIDES)                         # Apply inverter overrides
    except Exception as e:                                                                         # Catch errors
        print(f"4.4.1  inverter overrides failed: {e}")                                          # Print warning
    try:                                                                                           # Attempt panel overrides
        apply_pv_panel_overrides(app, PV_PANEL_OVERRIDES)                                          # Apply panel overrides
    except Exception as e:                                                                         # Catch errors
        print(f"4.4.1  panel overrides failed: {e}")                                             # Print warning

    # 4.4.2 ▶️ Run QDS and return RESULTS
    ok = run_simulation(pv_overrides=pv_overrides)                                                 # Run simulation
    if not ok:                                                                                     # If failed
        print("4.8.1  Simulation failed.")                                                       # Print failure
        return {}                                                                                  # Return empty dict
    print("4.8.1  Simulation completed, returning RESULTS.")                                      # Print success
    return dict(RESULTS)                                                                           # Return results dictionary

