# gui_app.py
# ENGR489 PowerFactory model

# 1.0 Enviroment Set UP
    # 1.1 Plotting and Graphing setup

# 2.0 UI constants and suburb display names
    # 2.1 UI Theme constents
    # 2.2 Suburb Full Name list
    # 2.3 Suburb Variable Defaults
    # 2.4 Result labels used for the results box

# 3.0 UI setup
    # 3.1.0 Create the main window, build layout, wire widgets
        # 3.1.1 Window basics
        # 3.1.2 Style / colour scheme
        # 3.1.3 Main layout: horizontal split (left sliders/run; right graphs/map)
        # 3.1.4 Left-hand window (sliders + RUN)
        # 3.1.5 RUN row (cog + RUN + EXPORT)
            # 3.1.5.1 Settings cog
            # 3.1.5.2 RUN button
            # 3.1.5.3 EXPORT button
        # 3.1.6 Scrollable suburb window area
            # 3.1.6.1 Grid widths (4 cols)
            # 3.1.6.2 Scrollregion update
            # 3.1.6.3 Mouse wheel setup
        # 3.1.7 Suburb slider rows + state stores
        # 3.1.8 App-wide state + inverter counts
            # 3.1.8.1 Build slider rows per suburb
        # 3.1.9 Load initial slider values and bus names
            # 3.1.9.1 Track last input signature
        # 3.1.10 Right: graphs + map vertically
        # 3.1.11 Graph frame (top)
            # 3.1.11.1 Graph area (Matplotlib)
        # 3.1.12 Map frame (middle)
            # 3.1.12.1 Fixed split container
            # 3.1.12.2 Two equal-width columns
            # 3.1.12.3 Left canvas
            # 3.1.12.4 Right canvas
            # 3.1.12.5 Image refs/paths
        # 3.1.13 Default pane heights
            # 3.1.13.1 Load both maps
        # 3.1.14 Metric selector for results box
    # 3.2.0 Passive check PV using backend lists (no PF calls)
    # 3.3.0 Build a deterministic snapshot of sliders to compare runs
    # 3.4.0 Build one two-row slider+labels group for a suburb
        # 3.4.1 Row A: suburb button
        # 3.4.2 Row A: kW label
        # 3.4.3 Seed slider variable
        # 3.4.4 Tooltip helper
        # 3.4.5 Row A: “kW per inverter” Entry
        # 3.4.6 Row B: slider rail + nudges
        # 3.4.7 Row B: slider (0–100%)
        # 3.4.8 Row B: % Entry
        # 3.4.9 Row A/B: result lines
        # 3.4.10 Seed initial UI state
    # 3.5.0  Run Button Handler
    # 3.6.0 Show calculated kW based on %, homes, and kW/inverter
    # 3.6.1 Sync slider widgets with backend RESULTS
    # 3.7.0 Keep the % Entry text synced with the slider ( 0 to 100)
    # 3.8.0 Print the current % to result line 1
    # 3.9.0 When Suburb Clicked Highlight selection, sync state, ensure curves, draw plot
        # 3.9.1 Visual selection highlight
        # 3.9.2 Sync state with slider
        # 3.9.3 Draw selected suburb curves
        # 3.9.4 Return button to default style colours
        # 3.9.5 Plot voltage, load, tx, and line data for a selected suburb
            # 3.9.5.1 Helpers and accumulators
            # 3.9.5.2 PV Data (kW)
            # 3.9.5.3 Load Data (kW)
            # 3.9.5.4 Transformer Loading (%)
            # 3.9.5.5 Line Loading (%)
            # 3.9.5.6 Dynamic scaling
            # 3.9.5.7 Final formatting and ticks
            # 3.9.5.8 Min/Max p.u. markers
            # 3.9.5.9 Draw
        # 3.9.6 Display message before any results exist
    # 3.10.0 Draw load/PV curves for selected suburb onto the axes
        # 3.10.1 Clear axes
        # 3.10.2 Left Y: load/PV (kW)
        # 3.10.3 Right Y: Tx/Line (%)
        # 3.10.4 Legend merge
        # 3.10.5 Title and refresh
        # 3.10.6 Ensure initial %/kW are populated from backend
    # 3.11.0 CSV Export summary + hourly PV/Load/Tx/Line
        # 3.11.1 CSV headers (summary + per-hour)
        # 3.11.2 pick save path
        # 3.11.3 CSV build and write rows

# 4.0 Cool colors setup
    # 4.1.0 open_settings
        # 4.1.1 Create the window
        # 4.1.2 Close handler restores cog
        # 4.1.3 Depress/disable cog while open
        # 4.1.4 Close box restores cog
    # 4.2.0 Compute bg/fg for the current scheme and repaint the widget tree
    # 4.3.0 Apply colours to supported widget types without breaking special cases

# 5.0 Load Map to GUI
    # 5.1.0 load the map images
    # 5.2.0 draw the maps on canvas

# 6.0 Simulation Run
    # 6.1.0 When Run clicked disable run button, get input and run
    # 6.2.0 Read all sliders and keep the cache in sync
        # 6.2.1 Keep cache in sync
    # 6.3.0 Run sim in background, update GUI on main thread
        # 6.3.1 Build panel overrides dict
        # 6.3.2 Run simulation
        # 6.3.3 Schedule GUI updates
    # 6.4.0 Refresh result labels and re-enable the RUN button

# 7.0 Results display
    # 7.1 Copy backend values (min/max/ok) into the GUI’s per-suburb cache
        # 7.1.1 Build Mapping from bus and PV Keys
        # 7.1.2 Iterate over bus results and update per-suburb state
        # 7.1.3 Link pv key to load/tx/line
    # 7.1.4 Build load_dict by matching suburb_state to load to results
    # 7.2 Format-specific text updates per suburb
    # 7.3 Read cache and redraw the two-line labels with hours

# 8.0 Main



# =================================================================================================
#====================================================================================================
# 1.0 Enviroment Set UP
#====================================================================================================
# =================================================================================================


import os                                                                                           # Standard library
import threading                                                                                    # Run long-running simulations off the main Tkinter thread
import tkinter as tk                                                                                # Main Tk GUI tools
from tkinter import ttk, messagebox                                                                 # TK Themed widgets and dialog boxes

try:                                                                                                # Try to get better image handling if pillow is available
    from PIL import Image, ImageTk                                                                  # Get pillow image handling tools
    PIL_AVAILABLE = True                                                                            # GUI can use Pillow
except Exception:                                                                                   # If Pillow isn't installed or fails to import, don't loose your shit
    PIL_AVAILABLE = False                                                                           # Disable image nice bits but keep the app running

import otaki_sim as sim                                                                             # Backend code for PowerFactory logic and return results
from copy import deepcopy                                                                           # Keep independent copies of lists etc for the GUI storage
import time                                                                                         # Time warp baby
import unicodedata


# 1.1 Plotting and Graphing setup ------------------------------------------------------------------
try:                                                                                                # Try to enable plotting inside the Tkinter window if Matplotlib is avaiable
    from matplotlib.figure import Figure                                                            # The plotting frame
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg                                 # Puts a frame into the Tk widget
    from matplotlib.ticker import MultipleLocator
    MPL_OK = True                                                                                   # GUI can show charts if Matplotlib imported successfully
except Exception:                                                                                   # If Matplotlib import fails, carry on without plots
    MPL_OK = False                                                                                  # Disable graphing so the rest of the app still runs


# =================================================================================================
#====================================================================================================
# 2.0 UI constants and suburb display names
#====================================================================================================
# =================================================================================================


# 2.1 UI Theme constents-------------------------------------------------------------------------------
APP_TITLE = "Ōtaki Grid – Solar Penetration GUI"                                                    # Show the title in the title bar
DEFAULT_BG = "#D9D9D9"                                                                            # Main background colour for the widgets (light grey)
OK_COLOUR = "#1E7D32"                                                                               # Colour used when voltages are within limits (green)
WARN_COLOUR = "#C62828"                                                                             # Colour used when voltages are out of limits (red)
ACCENT_GREEN = "#2ECC71"                                                                            # Accent colour for highlights, pressed buttons etc
MID_COLOUR = "#F57C00"                                                                              # Warning Orange


# 2.2 Suburb Full Name list ------------------------------------------------------------------------ 
SUBURB_FULL = {                                               
    "OTBa_PV": "Ōtaki Beach A",                                                                     # Label shown next to its slider
    "OTBb_PV": "Ōtaki Beach B",
    "OTBc_PV": "Ōtaki Beach C",
    "OTCa_PV": "Ōtaki Commercial A",
    "OTCb_PV": "Ōtaki Commercial B",
    "OTIa_PV": "Ōtaki Industrial",
    "OTKa_PV": "Ōtaki Township A",
    "OTKb_PV": "Ōtaki Township B",
    "OTKc_PV": "Ōtaki Township C",
    "OTS_PV" : "Ōtaki College",
    "RGUa_PV": "Rangiuru Rd A",
    "RGUb_PV": "Rangiuru Rd B",
    "TRE_PV" : "Te Rauparaha",
    "WTVa_PV": "Waitohu Valley A",
    "WTVb_PV": "Waitohu Valley B",
    "WTVc_PV": "Waitohu Valley C",
}


# 2.3 Suburb Variable Defaults ----------------------------------------------------------------------
SUBURB_DEFAULTS = {
    "pv_pct": 0.0,                                                                                  # Slider % PV penetration
    "pv_kw": 0.0,                                                                                   # Installed PV (kW) computed from slider
    "u_min": None,                                                                                  # Min pu voltage over the day
    "u_max": None,                                                                                  # Max pu voltage over the day
    "ok": None,                                                                                     # True if voltages within limits
    "load_curve": None,                                                                             # kW time series (from backend)
    "pv_profile": None,                                                                             # kW time series (from backend)
    "last_updated": None,                                                                           # Timestamp of last update
    "bus": None,                                                                                    # Associated PF bus name
    "tx_loading": None,                                                                             # Transformer loading (%) at summary point
    "line_loading": None,                                                                           # Line loading (%) at summary point
    "tx_overload": None,                                                                            # True if any Tx overload detected
    "line_overload": None,                                                                          # True if any line overload detected
    "tx_loading_curve": None,                                                                       # Transformer loading % time series
    "line_loading_curve": None,                                                                     # Line loading % time series
}


# 2.4 Result labels used for the results box -------------------------------------------------------
SHOW_METRIC_SELECTOR = False                                                                        # Drop down menu for graph, set true to enable, not sure if needed
METRIC_OPTIONS = [
    "Voltage p.u. (min over day)",                                                                  # Show the minimum pu voltage across the day
    "Voltage p.u. (max over day)",                                                                  # Show the maximum pu voltage across the day
    "OK/Limit flag",                                                                                # Show a status from backend (True/Green, False/red)
]


# =================================================================================================
# ==================================================================================================
# 3.0 ---------- UI setup ----------
# ==================================================================================================
# =================================================================================================

class App(tk.Tk):


# 3.1.0  Create the main window, build layout, wire widgets etc ------------------------------------
    def __init__(self):
        super().__init__()


# 3.1.1 Window basics ------------------------------------------------------------------------------
        self.title(APP_TITLE)                                                                       # Window title
        self.configure(bg=DEFAULT_BG)                                                               # Main ackground colour
        self.geometry("1400x900")                                                                   # Initial size
        self.minsize(1100, 700)                                                                     # Minimum usable size


# 3.1.2 Style / colour scheme ----------------------------------------------------------------------
        self.scheme = "Grey"                                                                        # Settings color scheme
        self.style = ttk.Style(self)                                                                # ttk theme thingy
        self.style.theme_use("default")                                                             # Use the default ttk theme


# 3.1.3 Main layout: Horizontal split windows, left for the sliders and run button, right top bottom for the graph and map
        main_pane = tk.PanedWindow(self, orient="horizontal", 
        sashrelief="raised", bg=DEFAULT_BG)                                                         # Make a horizontal Window 
        main_pane.pack(fill="both", expand=True)                                                    # pack it so it fills up the whole window and grows/shrinks with resizing


# 3.1.4 Left-hand window with the slider list and RUN button ---------------------------------------
        left_frame = tk.Frame(main_pane, bg=DEFAULT_BG)                                             # Builds a smaller frame inside the left side window
        main_pane.add(left_frame, minsize=350)                                                      # add the frame to the window; make a minimum width so it doesn’t go too small


# 3.1.5 RUN row (bottom-left): square cog on the left + wide RUN button on the right ---------------
        run_row = tk.Frame(left_frame, bg=DEFAULT_BG)                                               # container strip at the bottom
        run_row.pack(side="bottom", fill="x", pady=8)                                               # stick to the bottom of the left panel

        # Use grid so the cog can be a fixed-width square and the RUN can expand ------------------
        run_row.grid_columnconfigure(0, weight=0, minsize=54)                                       # col 0 = cog slot
        run_row.grid_columnconfigure(1, weight=1)                                                   # col 1 = RUN button expands to fill remaining width
        run_row.grid_rowconfigure(0, weight=1)                                                      # let buttons stretch to the row height


# 3.1.5.1 Settings cog (square, same vertical height as RUN) ---------------------------------------
        self.cog_btn = tk.Button(
            run_row,
            text="⚙",                                                                              # gear icon
            font=("Segoe UI Symbol", 14),                                                           # icon size
            bg=DEFAULT_BG, activebackground=DEFAULT_BG, fg="black",                                 # neutral look
            relief="raised", borderwidth=4, highlightthickness=2,                                   # match the visual weight of RUN
            width=2, height=1,                                                                      # narrow width
            command=self.open_settings                                                              # same handler as before
        )
        self.cog_btn.grid(row=0, column=0, sticky="nsew", 
                          padx=(6, 6), pady=6)                                                      # fill its cell; padding mirrors RUN


# 3.1.5.2 RUN button (big, expands across the rest of the row) ----------------------------------
        self.run_btn = tk.Button(                                                                 # Create RUN button widget
            run_row, text="RUN", font=("Segoe UI", 13, "bold"),                                   # Label text and font
            bg=ACCENT_GREEN, activebackground=ACCENT_GREEN, fg="white",                           # Button colours
            relief="raised", borderwidth=4, highlightthickness=2, padx=24, pady=6,                # Style and padding
            command=self.on_run_clicked                                                           # Command: run simulation
        )
        self.run_btn.grid(row=0, column=1, sticky="nsew",                                         # Place in grid layout
                          padx=(0, 6), pady=6)                                                    # Stretch horizontally, same vertical padding as cog
        self._settings_win = None                                                                 # Track active settings window (None = closed)


# 3.1.5.3 EXPORT button (raised look; resilient to theme/color changes) --------------------------
        self.export_btn = tk.Button(                                                              # Create EXPORT CSV button widget
            run_row, text="EXPORT CSV", font=("Segoe UI", 11, "bold"),                            # Label text and font
            bg="#455A64", activebackground="#455A64", fg="white",                            # Button colours (dark grey)
            relief="raised", overrelief="raised", borderwidth=4,                                  # Style and raised look
            highlightthickness=2, padx=12, pady=6, takefocus=1,                                   # Padding and focus
            command=self._export_to_csv                                                           # Command: export results to CSV
        )
        self.export_btn.grid(row=0, column=1, sticky="nsew", padx=(0, 6), pady=6)                 # Place export button in grid
        self.run_btn.grid_configure(row=0, column=2, sticky="nsew", padx=(0, 6), pady=6)          # Adjust RUN button placement
        run_row.grid_columnconfigure(2, weight=1)                                                 # Expand column 2 to stretch buttons


# 3.1.6 Scrollable suburbd window area -----------------------------------------------------------
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical")                                # Vertical scrollbar for the left panel
        scrollbar.pack(side="left", fill="y")                                                   # lock the scrollbar on the far left and stretch it vertically
        canvas = tk.Canvas(left_frame, bg=DEFAULT_BG, 
                           highlightthickness=0)                                                # canvas will host the scrolling content window
        canvas.pack(side="left", fill="both", expand=True)                                      # place canvas to the right of the scrollbar, make it resize with the window
        canvas.configure(yscrollcommand=scrollbar.set)                                          # canvas tells the scrollbar where the view is (for the thumb position)
        scrollbar.config(command=canvas.yview)                                                  # scrollbar controls the vertical view of the canvas when dragged
        self.sliders_frame = tk.LabelFrame(canvas, text="Solar " \
        "Penetration (%)", bg=DEFAULT_BG)                                                       # inner frame that actually holds slider rows
        self.sliders_frame_id = canvas.create_window((0, 0), 
        window=self.sliders_frame, anchor="nw")                                                 # put the inner frame at (0,0) inside the canvas


# 3.1.6.1 Grid widths across 4 columns where each slider row is laid out   ---------------------
        self.sliders_frame.columnconfigure(0, weight=0, minsize=0)                              # column 0: suburb button, fixed width
        self.sliders_frame.columnconfigure(1, weight=0, minsize=0)                              # column 1: kW label, fixed width
        self.sliders_frame.columnconfigure(2, weight=0, minsize=0)                              # column 2: % entry, fixed width
        self.sliders_frame.columnconfigure(3, weight=0, minsize=190)                            # column 3: results labels, 190 px so the text won't be cramped


# 3.1.6.2 Update scrollable area whenever content size changes -------------------------------- 
        def on_frame_config(event):                                                             # callback fired when sliders_frame changes size like adding rows
            canvas.configure(scrollregion=canvas.bbox("all"))                                   # set the scrollable bit to fit all the content
        self.sliders_frame.bind("<Configure>", on_frame_config)                                 # wire the size-change event to the callback


# 3.1.6.3 Mouse wheel setup -------------------------------------------------------------------                                         
        def _on_mousewheel(event):                                                              # wheel data
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 
                                "units")                                                        # scroll up/down by 1 unit per notch
        canvas.bind_all("<MouseWheel>", _on_mousewheel)                                         # listen for mouse wheel events anywhere in the app and route to this canvas


# 3.1.7 Suburb slider rows and their state storage --------------------------------------------------
        self.slider_vars     = {}                                                                 # Tk variables for each suburb slider (%)
        self.percent_entries = {}                                                                 # Entry widgets for direct % editing
        self.result_lines    = {}                                                                 # Matplotlib line handles per suburb
        self.result_labels   = {}                                                                 # Label widgets for results text per suburb
        self.suburb_buttons  = {}                                                                 # Button widgets for suburb selection
        self.inv_kw_vars     = {}                                                                 # kW/inverter vars
        self.inv_kw_entries  = {}                                                                 # kW/inverter entries
        self.inv_kw_start    = {}                                                                 # PF-start kW/inverter
        row_base = 0                                                                              # Grid row index for first slider block


# 3.1.9 Prefetch PF kW/inverter per suburb (on startup) --------------------------------------------
        self.inv_kw_start = {}                                                                    # Reset cache of starting kW/inverter
        for k, nmods in (sim.PV_PANEL_OVERRIDES or {}).items():                                   # Iterate PV panel overrides from sim
            self.inv_kw_start[k] = round((nmods * sim.PANEL_WATT)/1000, 3)                        # Compute kW per inverter from panels


# 3.1.8 App-wide state + pull inverter counts from model -------------------------------------------
        self.suburb_state = {}                                                                    # Master per-suburb state dict
        self.last_run_signature = None                                                            # Tracks last run inputs for change detection
        self.pv_inverters = dict(sim.PV_INV_OVERRIDES)                                            # otaki_sim already ran refresh_pv_overrides_from_model(app) at import
        #print("[gui] pv_inverters (from model):", self.pv_inverters)                             # Debug print of inverter counts


# 3.1.8.1 Build one two-row slider per suburb ------------------------------------------------------
        row_base = 0                                                                              # Reset grid row for building sliders
        def normalize_label(pv_key):                                                              # Helper: ASCII-normalize for sorting labels
            label = SUBURB_FULL[pv_key]                                                           # Lookup display label from key
            return unicodedata.normalize("NFKD", label).encode("ASCII", "ignore").decode()        # Strip accents for stable sort

        self.ordered_pv_keys = sorted(SUBURB_FULL, key=normalize_label)                           # Deterministic alphabetical order
        for pv_key in self.ordered_pv_keys:                                                       # Build rows in sorted order
            self._build_slider_row(self.sliders_frame, row_base, pv_key)                          # Create the 2-row slider block
            row_base += 2                                                                         # Advance to next block (two rows)


# 3.1.9 Load the storage with initial slider values and bus names ----------------------------------
        from copy import deepcopy                                                                 # Local import for clean cloning
        self.suburb_state = {}                                                                    # Reinitialize state store
        for pv_key in self.ordered_pv_keys:                                                       # use same alphabetical list
            item = deepcopy(SUBURB_DEFAULTS)                                                      # Clone default template
            item["pv_pct"] = float(self.slider_vars[pv_key].get())                                # Seed with current slider %
            item["pv_kw"]  = None                                                                 # Will be computed after results
            item["bus"]    = pv_key.replace("_PV", "_0.415")                                      # Derive associated bus name
            item["load_curve"] = None                                                             # Placeholder for load kW time-series
            item["pv_profile"] = None                                                             # Placeholder for PV kW time-series
            self.suburb_state[pv_key] = item                                                      # Save initialized record


# 3.1.9.1 Track last input signature separately (re-set after building) ----------------------------
        self.last_run_signature = None                                                              # make sure no old signature is left after initial UI build


# 3.1.10 Right: graphs + map stacked vertically                                                     # right-hand window stacks the graph on top and map on bottom
        right_frame = tk.Frame(main_pane, bg=DEFAULT_BG)                                            # Frame that sits in the right side of main window
        main_pane.add(right_frame)                                                                  # add the right frame as the second child of the main window

        vertical_pane = tk.PanedWindow(right_frame,                                                 # Create a vertical paned window
        orient="vertical", sashrelief="raised", bg=DEFAULT_BG)                                      # vertical split for graphs/map
        vertical_pane.pack(fill="both", expand=True)                                                # let the vertical pane fill the whole right side and resize with the window


# 3.1.11 Graph frame (top)                                                                          # labelled frame for the pgraph area
        self.graphs_frame = tk.LabelFrame(vertical_pane,                                            # Create labelled frame for graphs
        text="Graphs", bg=DEFAULT_BG)                                                               # top section title 
        self.metric_var = tk.StringVar(value=METRIC_OPTIONS[0])                                     # hidden default selection; no UI shown
        vertical_pane.add(self.graphs_frame, minsize=100)                                           # add to the vertical split, don’t let it shrink below ~100 px


# 3.1.11.1 Graph area (Matplotlib embedded in Tk) --------------------------------------------------
        if MPL_OK:                                                                                  # Only build plotting area if Matplotlib was imported successfully
            self.fig = Figure(figsize=(5, 3), dpi=100)                                              # Create a new Matplotlib Figure (canvas) with fixed size and resolution
            self.ax = self.fig.add_subplot(111)                                                     # Add one subplot (1 row, 1 col, first cell) → left axis for PV & demand
            self.ax2 = self.ax.twinx()                                                              # Create a second y-axis (right side) sharing the same x-axis for % loadings
            self.ax.set_title("Select a suburb to display curves")                                  # Set the main title of the graph
            self.ax.set_xlabel("Hour")                                                              # Label for the x-axis → simulation time in hours
            self.ax.set_ylabel("kW")                                                                # Label for the left y-axis → PV and demand values in kW
            self.ax2.set_ylabel("% Loading")                                                        # Label for the right y-axis → transformer/line loading in %
            self.ax.grid(True, linestyle="--", linewidth=0.5)                                       # Enable dashed grid lines for readability

            
            self.lines = {                                                                          # Dictionary to store line objects (plot handles) for later updates when a suburb is selected
                "pv": None,                                                                         # PV generation curve in kW
                "load": None,                                                                       # Load demand curve in kW
                "tx_pct": None,                                                                     # Transformer loading curve in %
                "line_pct": None,                                                                   # Line loading curve in %
            }

            self.canvas_mpl = FigureCanvasTkAgg(self.fig,                                           # Create Tk canvas for Matplotlib figure
            master=self.graphs_frame)                                                               # Embed the Matplotlib figure into the Tkinter frame
            self.canvas_mpl.get_tk_widget().pack(fill="both",                                       # Pack canvas widget to fill frame
            expand=True, padx=8, pady=(0, 8))                                                       # Pack canvas so it fills available space
        else:                                                                                       # If Matplotlib is not available
            ttk.Label(self.graphs_frame,                                                            # Place a text label inside the graphs frame
                    text="Matplotlib not installed.\nInstall it to see plots.").pack(padx=8, pady=8)  # Message shown instead of graph


# 3.1.12 Map frame (middle) — fixed 50/50 split (no adjustable sash) -------------------------------
        self.map_frame = tk.LabelFrame(vertical_pane, text="Maps", bg=DEFAULT_BG)                   # labelled container for both maps
        vertical_pane.add(self.map_frame, minsize=100)                                              # add to the vertical split


# 3.1.12.1 Fixed split using a simple Frame + grid (two equal columns) ------------------------------
        maps_row = tk.Frame(self.map_frame, bg=DEFAULT_BG)                                          # inner container row
        maps_row.pack(fill="both", expand=True)                                                     # occupy all available space


# 3.1.12.2 Make two equal-width columns that always share space 50/50 -------------------------------
        maps_row.grid_columnconfigure(0, weight=1, uniform="maps")                                  # left column stretches equally
        maps_row.grid_columnconfigure(1, weight=1, uniform="maps")                                  # right column stretches equally
        maps_row.grid_rowconfigure(0, weight=1)                                                     # row stretches vertically too


# 3.1.12.3 Left canvas (distribution/normal map) ----------------------------------------------------
        self.map_canvas_left = tk.Canvas(maps_row, bg="#BFBFBF", highlightthickness=0)            # left image area
        self.map_canvas_left.grid(row=0, column=0, sticky="nsew", padx=(0, 2), pady=0)              # fill left cell; tiny gap between maps


# 3.1.12.4 Right canvas (transmission single-line map) ----------------------------------------------
        self.map_canvas_right = tk.Canvas(maps_row, bg="#BFBFBF", highlightthickness=0)          # right image area
        self.map_canvas_right.grid(row=0, column=1, sticky="nsew", padx=(2, 0), pady=0)            # fill right cell


# 3.1.12.5 Image refs/paths so Tk doesn’t GC the bitmaps --------------------------------------------
        self._map_img_left = None                                                                   # PhotoImage for left canvas
        self._map_img_right = None                                                                  # PhotoImage for right canvas
        self._map_img_path_left = None                                                              # file path for left image
        self._map_img_path_right = None                                                             # file path for right image


# 3.1.13 Default pane heights — keep so maps are visible on startup --------------------------------
        self.after(80, lambda: vertical_pane.paneconfig(self.graphs_frame, height=600))             # graphs top ~220 px
        self.after(80, lambda: vertical_pane.paneconfig(self.map_frame,    height=100))             # maps middle ~400 px


# 3.1.13.1 Load both maps (absolute paths) ----------------------------------------------------------
        self.load_map_images(                                                                       # load both map images
            r"C:\Users\chris\OneDrive - Victoria University of Wellington - STUDENT\Vic\ENGR489\Artifact\Map.png",               # left: normal map
            r"C:\Users\chris\OneDrive - Victoria University of Wellington - STUDENT\Vic\ENGR489\Artifact\Single Line Map.png"    # right: single-line transmission map
        )


# 3.1.14 Metric selector for results boxe -----------------------------------------------------------

        if SHOW_METRIC_SELECTOR:                                                                        # Only build if enabled
            metric_bar = tk.Frame(self.graphs_frame, bg=DEFAULT_BG)                                     # Horizontal container for controls
            metric_bar.pack(side="top", fill="x", padx=8, pady=8)                                       # Fill width with padding
            ttk.Label(metric_bar, text="Display:").pack(side="left")                                    # Static label for combo
            self.metric_combo = ttk.Combobox(                                                            # Readonly dropdown
                metric_bar,
                values=METRIC_OPTIONS,
                textvariable=self.metric_var,
                state="readonly",
                width=28
            )
            self.metric_combo.pack(side="left", padx=8)                                                  # Place combo beside label
            self.metric_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_results())             # Refresh results on change


# 3.2.0 Passive check PV using backend lists (no PF calls) ----------------------------------------------

    def _check_pv_objects_on_startup(self):                                                             # Section 3.2.0 init readiness text
        for pv_key in self.ordered_pv_keys:                                                             # Iterate suburbs in alpha order
            l1, l2 = self.result_lines[pv_key]                                                           # Get the two status labels
            l1.configure(text="Ready", fg=OK_COLOUR)                                                     # Mark status as ready
            l2.configure(text=f"Bus: {pv_key.replace('_PV','_0.415')}", fg="black")                      # Show mapped bus name


# 3.3.0 Build a deterministic snapshot of sliders to compare runs ------------------

    def _input_signature(self):                                                                         # Section 3.3.0 build signature tuple
        return tuple(sorted((k, int(round(v.get()))) for k,                                             # Sort by key for stability
                            v in self.slider_vars.items()))                                              # Use integer percent values


# 3.4.0 Build one two-row slider+labels group for a suburb --------------------

    def _build_slider_row(self, parent, row_base, pv_key):                                              # Section 3.4.0 row builder

# 3.4.1 ROW A: Suburb button (click to plot curves) ---------------------------------------------------

        full_name = SUBURB_FULL.get(pv_key, pv_key)                                                     # Resolve display name
        name_btn = tk.Button(                                                                           # Create suburb button
            parent, text=full_name, anchor="w",
            relief="raised", bd=1, padx=6, pady=1,
            bg=DEFAULT_BG, activebackground="#CFCFCF",
            highlightthickness=1, highlightbackground="#BEBEBE",
            command=lambda k=pv_key: self._on_suburb_clicked(k)
        )
        name_btn.grid(row=row_base, column=0, sticky="w",                                               # Place in leftmost column
                      padx=(6, 0), pady=(4, 0))
        self.suburb_buttons[pv_key] = name_btn                                                          # Store handle by key


# 3.4.2 ROW A: kW label (installed capacity summary) --------------------------------------------------

        kw_lbl = tk.Label(parent, bd=1, relief="sunken", bg="white", fg="black",                        # Small numeric badge
                          width=8, padx=6, pady=2, anchor="e")
        kw_lbl.grid(row=row_base, column=1, sticky="e", padx=(2, 0), pady=(4, 0))                       # Right-align in col 1
        self.kw_labels = getattr(self, 'kw_labels', {})                                                 # Ensure dict exists
        self.kw_labels[pv_key] = kw_lbl                                                                 # Save label for updates


# 3.4.3 Seed slider variable with current inverter count ----------------------------------------------

        inv = int(self.pv_inverters.get(pv_key, 0))                                                     # Start from model inverter count
        var = tk.DoubleVar(value=inv)                                                                   # Backing variable for slider
        self.slider_vars[pv_key] = var                                                                  # Register var for this suburb


# 3.4.4 Tooltip helper (local) -------------------------------------------------------------------------

        def _suburb_tip_text(pv_key):                                                                   # Compose tooltip text
            meta = sim.PV_CONFIG.get(pv_key, {})                                                        # Get config record
            bus  = meta.get("bus", pv_key.replace("_PV", "_0.415"))                                     # Bus fallback mapping
            load = meta.get("load", "")                                                                 # Load name if any
            tx   = meta.get("tx", "")                                                                   # Transformer name if any
            line = meta.get("pline", "")                                                                # Line name if any
            return ("Click to display results\n"                                                        # Multiline tooltip body
                    f"Load: {load}\n"
                    f"PV: {pv_key}\n"
                    f"Tx: {tx}\n"
                    f"Bus: {bus}\n"
                    f"Line: {line}")

        def _tooltip(widget, text_or_fn):                                                               # Attach simple tooltip
            tip = tk.Toplevel(widget); tip.withdraw(); tip.overrideredirect(True)                       # Floating borderless window
            lbl = tk.Label(tip, text="", bg="#FFF9C4", relief="solid",                                # Tooltip label widget
                           borderwidth=1, padx=6, pady=2, justify="left"); lbl.pack()
            def _show(e):                                                                               # On enter: show near cursor
                try:
                    txt = text_or_fn() if callable(text_or_fn) else text_or_fn                          # Compute text if callable
                except Exception:
                    txt = str(text_or_fn)                                                               # Fallback to string
                lbl.config(text=txt)                                                                    # Set text on label
                tip.geometry(f"+{e.x_root+10}+{e.y_root+10}"); tip.deiconify()                          # Position and reveal
            def _hide(_): tip.withdraw()                                                                # On leave: hide tooltip
            widget.bind("<Enter>", _show); widget.bind("<Leave>", _hide)                                # Bind events

        _tooltip(name_btn, lambda k=pv_key: _suburb_tip_text(k))                                        # Live suburb tooltip
        _tooltip(kw_lbl, "Installed PV = % × homes × kW/inverter.")                                     # Explain kW badge


# 3.4.5 ROW A: “kW per inverter” Entry (quantised to panels) ---------------------------------------

        start_kw = float(self.inv_kw_start.get(pv_key, 6.0))                                            # Initial kW/inverter seed from PF (fallback 6.0)
        inv_kw_var = tk.DoubleVar(value=start_kw)                                                       # Backing Tk variable for entry
        self.inv_kw_vars[pv_key] = inv_kw_var                                                           # Register var for suburb
        inv_kw_ent = ttk.Entry(parent, width=8, justify="right")                                        # Numeric entry for kW/inverter
        inv_kw_ent.grid(row=row_base, column=2, sticky="e", padx=(2, 0), pady=(4, 0))                   # Place in col 2, right aligned
        self.inv_kw_entries[pv_key] = inv_kw_ent                                                        # Keep handle for later reads
        _tooltip(inv_kw_ent, f"Array size per inverter rounded for {sim.PANEL_WATT} W panels.")         # Explain rounding to panel size

        def _format_inv_kw():                                                                           # Parse/quantise entry to nearest panel
            raw = inv_kw_ent.get().strip().replace("kW","").strip()                                     # Strip units and spaces
            try:
                req_kw = max(0.0, float(raw)) if raw else float(inv_kw_var.get())                       # Use typed value or current var
            except ValueError:
                req_kw = float(inv_kw_var.get())                                                        # Fallback on parse error
            nmods   = int(round((req_kw * 1000.0) / sim.PANEL_WATT))                                    # panels/inverter
            show_kw = (nmods * sim.PANEL_WATT) / 1000.0                                                 # quantised kW
            inv_kw_var.set(show_kw)                                                                     # Sync var to quantised value
            inv_kw_ent.delete(0, tk.END)                                                                # Clear entry
            inv_kw_ent.insert(0, f"{show_kw:.2f} kW")                                                   # e.g., ~49.92 kW
            self._update_kw_label(pv_key)                                                               # Refresh installed kW badge

        inv_kw_ent.insert(0, f"{start_kw:g} kW")                                                        # Seed entry text with start value
        inv_kw_ent.bind("<Return>",   lambda *_: _format_inv_kw())                                      # Quantise on Enter
        inv_kw_ent.bind("<FocusOut>", lambda *_: _format_inv_kw())                                      # Quantise on blur


# 3.4.6 ROW B: Slider rail + left/right nudge buttons ----------------------------------------------

        NUDGE_W, NUDGE_H = 16, 16                                                                       # Arrow button box size
        def _nudge(k, d):                                                                               # Bump slider by delta d
            v = self.slider_vars[k]                                                                     # Get Tk variable
            new = max(0.0, min(100.0, float(v.get()) + d))                                              # Clamp to [0, 100]
            v.set(new); self._update_percent_entry(k); self._update_kw_label(k); self._touch_placeholder_capacity(k)  # Sync UI

        rail = tk.Frame(parent, bg=DEFAULT_BG)                                                          # Container row for slider + arrows
        rail.grid(row=row_base+1, column=0, columnspan=2, sticky="ew", padx=(6,0), pady=(0,4))          # span only cols 0–1
        rail.grid_columnconfigure(0, weight=0, minsize=NUDGE_W)                                         # Left arrow column
        rail.grid_columnconfigure(1, weight=1)                                                          # Slider column (expands)
        rail.grid_columnconfigure(2, weight=0, minsize=NUDGE_W)                                         # Right arrow column
        rail.grid_rowconfigure(0, minsize=NUDGE_H)                                                      # Row height matches arrow boxes

        left_wrap  = tk.Frame(rail, width=NUDGE_W, height=NUDGE_H, bg=DEFAULT_BG); left_wrap.grid(row=0, column=0, sticky="w"); left_wrap.pack_propagate(False)  # Left arrow holder
        right_wrap = tk.Frame(rail, width=NUDGE_W, height=NUDGE_H, bg=DEFAULT_BG); right_wrap.grid(row=0, column=2, sticky="e"); right_wrap.pack_propagate(False) # Right arrow holder

        self.style.configure("Arrow.TButton", padding=0)                                                # Compact arrow style
        ttk.Button(left_wrap,  text="<", style="Arrow.TButton", command=lambda k=pv_key: _nudge(k,-1)).pack(fill="both", expand=True)  # Left nudge

        slider = ttk.Scale(rail, orient="horizontal", from_=0.0, to=100.0,                              # Main penetration slider
                           variable=self.slider_vars[pv_key],
                           command=lambda _=None, k=pv_key: (
                               self._update_percent_entry(k),
                               self._update_kw_label(k),
                               self._touch_placeholder_capacity(k)
                           ))
        slider.grid(row=0, column=1, sticky="ew")                                                       # Stretch across center column

        ttk.Button(right_wrap, text=">", style="Arrow.TButton", command=lambda k=pv_key: _nudge(k,+1)).pack(fill="both", expand=True)  # Right nudge


# 3.4.7 ROW B: Slider (0–100%) ----------------------------------------------------------------------

        slider = ttk.Scale(rail, orient="horizontal", from_=0.0, to=100.0,                              # Duplicate slider (kept as in source)
                           variable=self.slider_vars[pv_key],
                           command=lambda _=None, k=pv_key: (
                               self._update_percent_entry(k),
                               self._update_kw_label(k),
                               self._touch_placeholder_capacity(k)
                           ))
        slider.grid(row=0, column=1, sticky="ew")                                                       # Place in center cell
        _tooltip(slider, "PV penetration (%) for this suburb.")                                         # Explain slider purpose
        ttk.Button(right_wrap, text=">", style="Arrow.TButton",
                   command=lambda k=pv_key: _nudge(k,+1)).pack(fill="both", expand=True)                # Mirror right nudge


# 3.4.8 ROW B: % Entry (at end of slider) -----------------------------------------------------------

        ent = ttk.Entry(parent, width=6, justify="right")                                          # Percent entry widget
        ent.insert(0, "0")                                                                         # Seed with 0
        ent.grid(row=row_base+1, column=2, sticky="e", padx=(2, 0), pady=(0, 4))                   # Place at end of row B
        self.percent_entries[pv_key] = ent                                                         # Keep handle by suburb key
        _tooltip(ent, "Type a % value (0–100).")                                                   # Tooltip: expected range

        def commit_entry(*_):                                                                      # Commit typed % into slider var
            try: val = float(ent.get().replace('%',''))                                            # Parse number (strip %)
            except ValueError: val = 0.0                                                           # Fallback to 0 on error
            val = max(0.0, min(100.0, val))                                                        # Clamp to [0, 100]
            var.set(val); self._update_percent_entry(pv_key); self._update_kw_label(pv_key); self._touch_placeholder_capacity(pv_key)  # Sync UI and placeholders
        ent.bind("<Return>", commit_entry); ent.bind("<FocusOut>", commit_entry)                   # Apply on Enter/blur


# 3.4.9 ROW A/B: Result lines -----------------------------------------------------------------------

        res_line1 = tk.Label(parent, bd=1, relief="sunken", bg="white", fg="black", padx=6, pady=2, anchor="w")  # Upper result label
        res_line1.grid(row=row_base, column=3, sticky="nsew", padx=(8, 6), pady=(4, 0))            # Span cell for row A
        self.result_labels[pv_key] = res_line1                                                      # Keep reference

        res_line2 = tk.Label(parent, bd=1, relief="sunken", bg="white", fg="black", padx=6, pady=2, anchor="w")  # Lower result label
        res_line2.grid(row=row_base + 1, column=3, sticky="nsew", padx=(8, 6), pady=(0, 4))        # Span cell for row B
        self.result_lines[pv_key] = (res_line1, res_line2)                                          # Tuple of both labels

        # Tooltips for p.u. result labels with actual bus name
        bus_name = sim.PV_CONFIG.get(pv_key, {}).get("bus", pv_key.replace("_PV", "_0.415"))       # Resolve bus name for tooltip
        _tooltip(res_line1, f"p.u. values for {bus_name} bus")                                      # Attach tooltip to line 1
        _tooltip(res_line2, f"p.u. values for {bus_name} bus")                                      # Attach tooltip to line 2


# 3.4.10 Seed initial UI state ----------------------------------------------------------------------

        self._update_percent_entry(pv_key)                                                          # Populate % entry from slider
        self._update_kw_label(pv_key)                                                               # Update installed kW badge
        self._touch_placeholder_capacity(pv_key)                                                    # Mark capacity as dirty


# 3.5.0  Run Button Handler -------------------------------------------------------------------------

    def _on_run_button(self):                                                                       # Launch simulation thread
        try:                                                                                        # Guard UI thread
            print("[gui] Run clicked — launching simulation thread")                                # Log action
            threading.Thread(target=self._run_model_thread, daemon=True).start()                    # Start background worker
        except Exception as e:                                                                      # Catch and display failures
            print("Run error:", e)                                                                  # Log error
            messagebox.showerror("Run failed", str(e))                                              # User-visible error dialog


# 3.6.0 Show calculated kW based on %, homes, and kW/inverter --------------------------------------

    def _update_kw_label(self, pv_key):                                                             # Compute and display installed kW
        pct = float(self.slider_vars[pv_key].get())                                                 # Current % penetration
        homes = int(sim.PV_CONFIG[pv_key]["homes"])                                                 # Homes for this suburb
        kw_per_inv = float(self.inv_kw_vars[pv_key].get())                                          # kW per inverter (quantised)
        total_kw = int(round((pct/100.0) * homes * kw_per_inv))                                     # Total installed kW (rounded)
        lbl = self.kw_labels.get(pv_key)                                                            # Lookup label widget
        if lbl:                                                                                     # If present, update text
            lbl.config(text=f"{total_kw} kW")                                                       # Show integer kW


# 3.6.1 Sync slider widgets with backend RESULTS ---------------------------------------------------

    def _update_sliders_from_results(self, *args):                                                    # Update slider positions from RESULTS
        try:                                                                                          # Safely handle missing/invalid data
            pv_meta = sim.RESULTS.get("pv_meta", {})                                                  # Get PV metadata dict from backend
            if not isinstance(pv_meta, dict):                                                         # Validate structure
                print(f"[gui] pv_meta not dict (got {type(pv_meta)}):", pv_meta)                      # Log unexpected type
                return                                                                                # Abort update

            # Temporarily block slider callbacks if they exist
            for pv_key, meta in pv_meta.items():                                                      # Iterate suburb metadata
                if not isinstance(meta, dict):                                                        # Skip malformed entries
                    continue                                                                          # Next item
                inv = meta.get("inverters")                                                           # Target inverter count
                if inv is not None and pv_key in self.slider_vars:                                    # Only if slider exists
                    # update without triggering _on_slider_change storms
                    var = self.slider_vars[pv_key]                                                    # Tk variable bound to slider
                    if var.get() != inv:                                                              # Avoid redundant sets
                        var.set(inv)                                                                  # Move slider silently

            print("[gui] sliders updated")                                                            # Trace completion
        except Exception as e:                                                                         # Catch-all guard
            print("update_sliders_from_results error:", e)                                            # Log error


# 3.7.0 Keep the % Entry text synced with the slider ( 0 to 100) --------------------------------------

    def _update_percent_entry(self, pv_key):                                                            # mirror the slider percentage into the tiny Entry field as “NN%”
        """Sync the tiny % Entry with the slider (integer 0..100)."""
        pct = int(round(float(self.slider_vars[pv_key].get())))                                         # take the slider float and round to an integer percent
        ent = self.percent_entries.get(pv_key)                                                          # fetch the Entry widget that displays the %
        if not ent:                                                                                     # safety: if it wasn’t created yet, bail
            return
        pct_str = f"{pct}%"                                                                             # format as a human-friendly percent string
        if ent.get() != pct_str:                                                                        # only rewrite if the text actually changed (avoids cursor jumps)
            ent.delete(0, "end")                                                                        # clear current entry text
            ent.insert(0, pct_str)                                                                      # insert the fresh “NN%” string


# 3.8.0 Print the current % to result line 1 -----------------------------------------------------------

    def _show_percent_in_results(self, pv_key):                                                         # optionally echo the live % on the first result line without wiping old results
        """Before/while running, show % on line 1, do not clear previous results."""
        l1, _ = self.result_lines[pv_key]                                                               # fetch the first (top) results label for this suburb
        pct = int(round(float(self.slider_vars[pv_key].get())))                                         # read and round the slider’s current value
        l1.configure(text=f"{pct}%", fg="black")                                                        # show “NN%” in neutral black to indicate a pending change
        pass                                                                                            # placeholder so the method is callable even if no live echo is desired


# 3.9.0 Highlight selection, sync state, ensure curves, draw plot -----------------------------------

    def _on_suburb_clicked(self, pv_key):                                                            # Handle suburb button click
        print(f"[gui] Suburb button clicked: {pv_key}")                                              # Trace selected suburb
        self.current_pv_key = pv_key                                                                 # <-- remember current selection
        st = self.suburb_state.get(pv_key)                                                           # Fetch cached state for suburb
        if not st:                                                                                   # If no cached state
            print(f"[gui] No cached state for {pv_key}")                                             # Log missing state
        else:                                                                                         # If state exists
            cached = st.get("u_min"), st.get("u_max")                                                # Pull cached min/max pu
            print(f"[gui] Cached voltage results for {pv_key}:", cached)                             # Trace cached tuple

        if not MPL_OK:                                                                               # If plotting unavailable
            messagebox.showinfo("Graphs", "Matplotlib not installed. Install it to see plots.")      # Inform user
            return                                                                                   # Bail out of handler

        normal = self._normal_button_colours()                                                       # Get default button colours
        for btn in self.suburb_buttons.values():                                                     # Reset all buttons
            try:                                                                                     # Best-effort styling
                btn.config(relief="raised", bg=normal["bg"], fg=normal["fg"],                        # Restore default look
                           activebackground=normal["activebackground"])
            except Exception:                                                                         # Some themes may fail
                pass                                                                                 # Ignore styling errors

        try:                                                                                         # Try primary highlight palette
            self.suburb_buttons[pv_key].config(                                                      # Style selected button
                relief="sunken", bg="#1976D2", fg="white", activebackground="#1976D2"
            )
        except Exception:                                                                             # Fallback highlight palette
            self.suburb_buttons[pv_key].config(                                                      # Style with fallback colours
                relief="sunken", bg="#4A90E2", fg="white", activebackground="#4A90E2"
            )

        st = self.suburb_state[pv_key]                                                               # Access mutable state
        st["pv_pct"] = float(self.slider_vars[pv_key].get())                                         # Sync latest slider %
        st["pv_kw"]  = int(round(st["pv_pct"])) * 6                                                  # Quick placeholder kW calc

        pv_key_str = pv_key                                                                          # Alias for clarity
        st = self.suburb_state.get(pv_key, {})                                                       # Safe fetch of state again
        code_to_name = {                                                                             # Map PF codes to names
            "OTBa": "Otaki Beach A", "OTBb": "Otaki Beach B", "OTBc": "Otaki Beach C",
            "OTCa": "Otaki Commercial A", "OTCb": "Otaki Commercial B",
            "OTIa": "Otaki Industrial A",
            "OTKa": "Otaki Town A", "OTKb": "Otaki Town B", "OTKc": "Otaki Town C",
            "OTS":  "Otaki School",
            "RGUa": "Rangiuru Rd A", "RGUb": "Rangiuru Rd B",
            "TRE":  "Te Rauparaha St",
            "WTVa": "Waitohu Valley A", "WTVb": "Waitohu Valley B", "WTVc": "Waitohu Valley C",
        }
        load_code = st.get("load") or sim.PV_CONFIG.get(pv_key_str, {}).get("load", "")             # Resolve load code
        pf_load_name = code_to_name.get(load_code, load_code)                                        # Human-friendly name

        load_data = sim.RESULTS.get("load", {}).get(pf_load_name, {})                                # Load demand series
        pv_data   = sim.RESULTS.get("pv", {}).get(pv_key_str, {})                                    # PV production series
        tx_data   = sim.RESULTS.get("tx", {}).get(st.get("tx", ""), {})                              # Transformer series
        line_data = sim.RESULTS.get("line", {}).get(st.get("line", ""), {})                          # Line loading series

        self._plot_curves(pv_key_str, load_data, pv_data, tx_data, line_data)                        # Draw/update plots


# 3.9.1 Visually indicate selection: reset all, then depress + recolour clicked ---------------------

        normal = self._normal_button_colours()                                                       # get default button colours to restore non-selected buttons
        for btn in self.suburb_buttons.values():                                                     # iterate all suburb buttons and reset their styling
            try:
                btn.config(relief="raised", bg=normal["bg"], fg=normal["fg"],
                           activebackground=normal["activebackground"])
            except Exception:
                pass                                                                                 # some themes/platforms may not support all options; ignore safely

        try:
            # Highlight colour (blue) for the selected suburb                                        # emphasise the chosen suburb
            self.suburb_buttons[pv_key].config(
                relief="sunken", bg="#1976D2", fg="white", activebackground="#1976D2"
            )
        except Exception:                                                                            # fallback palette if the primary colour fails (unlikely)
            self.suburb_buttons[pv_key].config(
                relief="sunken", bg="#4A90E2", fg="white", activebackground="#4A90E2"
            )


# 3.9.2 Keep state in sync with the current slider ---------------------------------------------------

        st = self.suburb_state[pv_key]                                                              # grab the cached state for this suburb
        st["pv_pct"] = float(self.slider_vars[pv_key].get())                                        # read current slider % (float 0..100)
        st["pv_kw"]  = int(round(st["pv_pct"])) * 6                                                 # derive kW capacity using 6 kW per 1% rule-of-thumb


# 3.9.3 Draw the selected suburb’s curves ------------------------------------------------------------

        pv_key_str = pv_key                                                                         # alias for clarity
        st = self.suburb_state.get(pv_key, {})                                                      # safe fetch of state dict

        # Resolve load code (e.g., "OTKa") → PF load name (e.g., "Otaki Town A")
        code_to_name = {
            "OTBa": "Otaki Beach A", "OTBb": "Otaki Beach B", "OTBc": "Otaki Beach C",
            "OTCa": "Otaki Commercial A", "OTCb": "Otaki Commercial B",
            "OTIa": "Otaki Industrial A",
            "OTKa": "Otaki Town A", "OTKb": "Otaki Town B", "OTKc": "Otaki Town C",
            "OTS":  "Otaki School",
            "RGUa": "Rangiuru Rd A", "RGUb": "Rangiuru Rd B",
            "TRE":  "Te Rauparaha St",
            "WTVa": "Waitohu Valley A", "WTVb": "Waitohu Valley B", "WTVc": "Waitohu Valley C",
        }

        # Prefer cached mapping; fall back to sim.PV_CONFIG (load code lives there). :contentReference[oaicite:2]{index=2}
        load_code = st.get("load") or sim.PV_CONFIG.get(pv_key_str, {}).get("load", "")            # resolve load code with fallback
        pf_load_name = code_to_name.get(load_code, load_code)                                      # map to human-friendly name

        # Pull data using PF object names saved by the backend. :contentReference[oaicite:3]{index=3}
        load_data = sim.RESULTS.get("load", {}).get(pf_load_name, {})                               # demand time-series dict
        pv_data   = sim.RESULTS.get("pv", {}).get(pv_key_str, {})                                   # PV production time-series dict
        tx_data   = sim.RESULTS.get("tx", {}).get(st.get("tx", ""), {})                             # transformer loading series
        line_data = sim.RESULTS.get("line", {}).get(st.get("line", ""), {})                         # line loading series

        self._plot_curves(pv_key_str, load_data, pv_data, tx_data, line_data)                       # render/update curves on plot


# 3.9.4 Return default button style colours ------------------------------------------------------

    def _normal_button_colours(self):                                                               # Used by _on_suburb_clicked to reset all suburb buttons.
        return {                                                                                   # return style dictionary
            "bg": self.cget("bg"),                                                                 # background: match window bg
            "fg": "black",                                                                         # foreground text colour
            "activebackground": self.cget("bg")                                                    # active background same as bg
        }


# 3.9.5 Plot voltage, load, tx, and line data for a selected suburb ------------------------------

    def _plot_curves(self, pv_key, load_data, pv_data, tx_data, line_data):                        # Plot curves for the selected suburb
        from matplotlib.ticker import MultipleLocator                                               # Import locator for tidy tick spacing
        self.ax.set_axis_on(); self.ax2.set_axis_on()                                              # Ensure both axes are visible
        self.ax.clear(); self.ax2.clear()                                                          # Clear previous plots on both axes

# 3.9.5.1 Helpers and accumulators ---------------------------------------------------------------
        def timestamps_to_hours(timestamps):                                                       # Convert UNIX timestamps to decimal hours
            import datetime                                                                        # Local import to avoid top-level dependency
            return [datetime.datetime.fromtimestamp(ts).hour +                                     # Hour component
                    datetime.datetime.fromtimestamp(ts).minute/60 for ts in timestamps]            # Minute as fraction of hour
        left_values = []                                                                           # kW series accumulator (PV + Load)
        right_values = []                                                                          # % series accumulator (Tx + Line)

# 3.9.5.2 PV Data (kW) --------------------------------------------------------------------------
        if pv_data:                                                                                # If PV data available
            t = pv_data.get("t", [])                                                               # Time vector (timestamps)
            P = pv_data.get("P_W", [])                                                             # PV power vector (kW)
            if t and P:                                                                            # Only if both vectors non-empty
                t_hours = timestamps_to_hours(t)                                                   # Convert timestamps to hours
                PkW = [p*1 for p in P]                                                             # Copy / cast list to kW values
                self.ax.plot(t_hours, PkW, label="PV (kW)", linewidth=2, color="blue")             # Plot PV on left axis
                left_values.extend(PkW)                                                            # Accumulate for autoscaling

# 3.9.5.3 Load Data (kW) ------------------------------------------------------------------------
        if load_data:                                                                              # If load data available
            t = load_data.get("t", [])                                                             # Time vector (timestamps)
            P = load_data.get("P_W", [])                                                           # Load power vector (kW)
            if t and P:                                                                            # Only if both vectors non-empty
                t_hours = timestamps_to_hours(t)                                                   # Convert timestamps to hours
                LkW = [p*1 for p in P]                                                             # Copy / cast list to kW values
                self.ax.plot(t_hours, LkW, label="Load (kW)", linewidth=2, color="orange")         # Plot Load on left axis
                left_values.extend(LkW)                                                            # Accumulate for autoscaling

# 3.9.5.4 Transformer Loading (%) --------------------------------------------------------------
        if tx_data:                                                                                # If transformer data available
            t = tx_data.get("t", [])                                                               # Time vector (timestamps)
            L = tx_data.get("loading_pct", [])                                                     # Transformer loading (%)
            if t and L:                                                                            # Only if both vectors non-empty
                t_hours = timestamps_to_hours(t)                                                   # Convert timestamps to hours
                self.ax2.plot(t_hours, L, label="Tx Loading(%)",                                   # Plot Tx loading on right axis
                              linestyle="--", linewidth=1, color="green")
                right_values.extend(L)                                                             # Accumulate for autoscaling

# 3.9.5.5 Line Loading (%) ---------------------------------------------------------------------
        if line_data:                                                                              # If line data available
            t = line_data.get("t", [])                                                             # Time vector (timestamps)
            L = line_data.get("loading_pct", [])                                                   # Line loading (%)
            if t and L:                                                                            # Only if both vectors non-empty
                t_hours = timestamps_to_hours(t)                                                   # Convert timestamps to hours
                self.ax2.plot(t_hours, L, label="Line Loading(%)",                                 # Plot Line loading on right axis
                              linestyle=":", linewidth=1, color="red")
                right_values.extend(L)                                                             # Accumulate for autoscaling


# 3.9.5.6 Dynamic scaling ---------------------------------------------------------------------------

        if left_values:                                                                                 # If there are kW-series values
            ymin, ymax = 0, 1000                                                                        # Start with default kW bounds
            ymin = min(ymin, min(left_values)); ymax = max(ymax, max(left_values))                      # Expand to fit data
            self.ax.set_ylim(ymin, ymax)                                                                # Apply left-axis limits
        if right_values:                                                                                # If there are % loading values
            ymin, ymax = 0, 150                                                                         # Start with default % bounds
            ymin = min(ymin, min(right_values)); ymax = max(ymax, max(right_values))                    # Expand to fit data
            self.ax2.set_ylim(ymin, ymax)                                                               # Apply right-axis limits


# 3.9.5.7 Final formatting and grid/ticks ------------------------------------------------------------

        self.ax.set_title(f"Results – {pv_key}")                                                        # Dynamic title with suburb key
        self.ax.set_xlabel("Hour of Day"); self.ax.set_ylabel("kW")                                     # Axis labels (left)
        self.ax2.set_ylabel("% Loading")                                                                # Right-axis label
        self.ax.set_xlim(0, 23)                                                                         # Constrain to 0..23 hours
        self.ax.set_xticks(list(range(0, 24, 2)))                                                       # Major ticks every 2 hours
        self.ax.yaxis.set_major_locator(MultipleLocator(100))                                           # kW ticks every 100
        self.ax.grid(True, linestyle="--", linewidth=0.5)                                               # Light dashed grid
        self.ax.legend(loc="upper left"); self.ax2.legend(loc="upper right")                            # Place legends
        self.ax.yaxis.set_label_position("left")                                                        # Ensure left label side
        self.ax2.yaxis.set_label_position("right"); self.ax2.yaxis.set_ticks_position("right")          # Right label/ticks


# 3.9.5.8 Vertical lines for Min/Max p.u. hours ------------------------------------------------------

        st = self.suburb_state.get(pv_key, {})                                                          # Current suburb state
        umin, tmin = st.get("u_min"), st.get("t_min")                                                   # Min p.u. and hour
        umax, tmax = st.get("u_max"), st.get("t_max")                                                   # Max p.u. and hour

        def _colour_for(pu):                                                                            # Choose colour based on p.u.
            if pu is None: return "grey"                                                                # Unknown → neutral grey
            if pu < 0.49 or pu > 1.051: return WARN_COLOUR                                              # Extreme out-of-bounds
            if 0.92 <= pu <= 1.02:      return OK_COLOUR                                                # Within preferred band
            if (0.919 <= pu < 0.92) or (1.02 < pu <= 1.051): return MID_COLOUR                          # Near-limit band
            return WARN_COLOUR                                                                           # Otherwise warn

        ymax_left = self.ax.get_ylim()[1]                                                               # Top of left-axis range
        if tmin is not None and umin is not None:                                                       # If min markers available
            c = _colour_for(umin)                                                                        # Colour for min marker
            self.ax.axvline(float(tmin), linestyle=":", color=c, linewidth=1, zorder=5)                 # Vertical min line
            self.ax.text(float(tmin), ymax_left, f"Min {umin:.2f}",                                     # Annotate min p.u.
                         color=c, rotation=90, va="top", ha="right", fontsize=8, zorder=6)
        if tmax is not None and umax is not None:                                                       # If max markers available
            c = _colour_for(umax)                                                                        # Colour for max marker
            self.ax.axvline(float(tmax), linestyle=":", color=c, linewidth=1, zorder=5)                 # Vertical max line
            self.ax.text(float(tmax), ymax_left, f"Max {umax:.2f}",                                     # Annotate max p.u.
                         color=c, rotation=90, va="top", ha="left", fontsize=8, zorder=6)


# 3.9.5.9 Draw ---------------------------------------------------------------------------------------

        self.canvas_mpl.draw()                                                                          # Repaint the Matplotlib canvas
        self.canvas_mpl.get_tk_widget().update_idletasks()                                              # Flush Tk events for smooth UI


# 3.9.6 Display message before any results exist ----------------------------------------------------

    def _show_plot_placeholder(self):                                                                # Draw placeholder when no results
        self.ax.clear(); self.ax2.clear()                                                            # Clear both axes
        self.ax.set_axis_off(); self.ax2.set_axis_off()                                              # Hide axes frames
        self.ax.text(0.5, 0.5,                                                                       # Centered instructional text
                     "Please run the simulation,\nthen click a suburb to display results.",
                     ha="center", va="center", transform=self.ax.transAxes)
        self.canvas_mpl.draw(); self.canvas_mpl.get_tk_widget().update_idletasks()                  # Refresh canvas/UI


# 3.10.0 Draw load/PV curves for selected suburb onto the axes -------------------------------------

    def _draw_suburb_curves(self, pv_key):                                                          # draw the currently selected suburb’s load and PV profiles
        st = self.suburb_state[pv_key]                                                              # fetch cached state (contains curves and labels) for this suburb
        full_name = SUBURB_FULL.get(pv_key, pv_key)                                                 # resolve human-readable suburb name for the plot title

        h = list(range(24))                                                                         # x-axis: 0..23 hours
        load = st.get("load_curve") or [0.0] * 24                                                   # y-series 1: load curve (fallback to zeros if missing)
        pv   = st.get("pv_profile") or [0.0] * 24                                                   # y-series 2: PV curve (fallback to zeros if missing)
        tx   = st.get("tx_pct")     or [0.0] * 24                                                   # y-series 3: transformer % loading (fallback to zeros)
        line = st.get("line_pct")   or [0.0] * 24                                                   # y-series 4: line % loading (fallback to zeros)

        
# 3.10.1 Clear both axes ---------------------------------------------------------------------------

        self.ax.clear()                                                                             # wipe previous content on left y-axis
        self.ax2.clear()                                                                            # wipe previous content on right y-axis


# 3.10.2 Left Y-Axis: Load and PV in kW ------------------------------------------------------------

        from matplotlib.ticker import MultipleLocator                                               # tick helper for clean y-steps

        self.ax.grid(True, linestyle="--", linewidth=0.5)                                           # light dashed grid
        load_line, = self.ax.plot(h, load, label="Load (kW)")                                       # plot demand curve
        pv_line,   = self.ax.plot(h, pv,   label="PV (kW)")                                         # plot PV curve

        self.ax.set_xlim(0, 23)                                                                     # 0–23 hours
        self.ax.set_xticks(list(range(0, 24, 2)))                                                   # every 2 hours
        self.ax.set_xlabel("Hour")                                                                  # x-axis label

        self.ax.set_ylabel("kW")                                                                    # left y-axis label
        self.ax.yaxis.set_major_locator(MultipleLocator(100))                                       # every 100 kW


# 3.10.3 Right Y-Axis: Transformer and Line loading in % -------------------------------------------

        tx_line, = self.ax2.plot(h, tx, label="Tx Load (%)",
                                  linestyle="--")                                                   # dashed line for transformer load
        ln_line, = self.ax2.plot(h, line, label="Line Load (%)", 
                                 linestyle=":")                                                     # dotted line for line load
        self.ax2.set_ylabel("% Loading")                                                            # right y-axis label
        self.ax2.set_ylim(0, 120)                                                                   # optionally clamp right axis to 0–120% (adjust as needed)


# 3.10.4 Legend: merge both axes' legends into one -------------------------------------------------

        lines = [load_line, pv_line, tx_line, ln_line]                                              # gather all lines
        labels = [line.get_label() for line in lines]                                               # get their labels
        self.ax.legend(lines, labels, loc="best")                                                   # show a unified legend


# 3.10.5 Title and refresh -------------------------------------------------------------------------
        self.ax.set_title(full_name)                                                                 # title = suburb name
        self.canvas_mpl.draw_idle()                                                                  # redraw figure without blocking


# 3.10.6 Ensure initial %/kW are populated from backend --------------------------------------------
    def _touch_placeholder_capacity(self, pv_key):                                                   # Create initial cached state for suburb using backend PV_CONFIG.
        if not hasattr(self, "suburb_state") or self.suburb_state is None:                           # Ensure state dict exists
            self.suburb_state = {}                                                                   # Init store if missing
        st = self.suburb_state.get(pv_key, {})                                                       # Get existing state or empty
        if "pv_pct" not in st:                                                                       # Use backend config: 1% per home
            homes = sim.PV_CONFIG.get(pv_key, {}).get("homes", 0)                                    # Read homes count from config
            st["pv_pct"] = homes                                                                     # Seed % = homes
            st["pv_kw"]  = int(round(homes)) * 6                                                     # Seed kW using 6 kW/% rule
        if "bus" not in st:                                                                          # If bus not set yet
            st["bus"] = sim.PV_CONFIG.get(pv_key, {}).get("bus")                                     # Copy bus from config
        self.suburb_state[pv_key] = st                                                                # Write back updated state


# 3.11.0 CSV Export summary + hourly PV/Load/Tx/Line ----------------------------------------------------
    def _export_to_csv(self):                                                                        # Export combined summary + timeseries to CSV
        import csv, datetime, tkinter as tk                                                          # Local imports for dialog and CSV
        from tkinter import filedialog as fd                                                         # Save-as dialog
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")                                  # Timestamp for rows


# 3.11.1 CSV headers (summary + per-hour) ----------------------------------------------------------
        headers = [                                                                                  # Column order for output rows
            "timestamp","pv_key","suburb","homes","slider_pct","kw_per_inv","installed_kw",
            "bus","tx_id","line_id","hour","pv_kw","load_kw","tx_pct","line_pct"
        ]

# 3.11.2 pick save path ----------------------------------------------------------------------------
        default = f"otaki_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"            # Suggested filename with timestamp
        path = fd.asksaveasfilename(defaultextension=".csv", initialfile=default,                    # Open Save dialog
                                    filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not path:                                                                                 # User cancelled
            return                                                                                   # Abort export

# 3.11.3 CSV build and write rows ---------------------------------------------------------------------
        try:                                                                                         # Guard file I/O
            with open(path, "w", newline="", encoding="utf-8") as f:                                 # Create CSV file
                writer = csv.writer(f); writer.writerow(headers)                                     # Write header row

                # load-code → PF name --------------- ----------------------------------------------
                code_to_name = {                                                                     # Mapping from codes to PF names
                    "OTBa": "Otaki Beach A","OTBb": "Otaki Beach B","OTBc": "Otaki Beach C",
                    "OTCa": "Otaki Commercial A","OTCb": "Otaki Commercial B","OTIa": "Otaki Industrial A",
                    "OTKa": "Otaki Town A","OTKb": "Otaki Town B","OTKc": "Otaki Town C",
                    "OTS": "Otaki School","RGUa": "Rangiuru Rd A","RGUb": "Rangiuru Rd B",
                    "TRE": "Te Rauparaha St","WTVa": "Waitohu Valley A","WTVb": "Waitohu Valley B","WTVc": "Waitohu Valley C",
                }

                for pv_key in self.ordered_pv_keys:                                                  # Iterate suburbs in display order
                    meta = sim.PV_CONFIG.get(pv_key, {})                                             # Config for this suburb
                    homes = int(meta.get("homes", 0))                                                # Homes count
                    slider_pct = float(self.slider_vars.get(pv_key, tk.DoubleVar(value=0.0)).get())  # Current slider %
                    kw_per_inv = float(self.inv_kw_vars.get(pv_key, tk.DoubleVar(value=6.0)).get())  # kW per inverter
                    installed_kw = round((slider_pct/100.0) * homes * kw_per_inv, 3)                 # Installed capacity (kW)
                    suburb = self.SUBURB_FULL.get(pv_key, pv_key) if hasattr(self, "SUBURB_FULL") else SUBURB_FULL.get(pv_key, pv_key)  # Suburb label

                    bus   = meta.get("bus", "")                                                      # Bus name
                    tx_id = meta.get("tx", "")                                                       # Transformer ID
                    ln_id = meta.get("pline", "")                                                    # Line ID

                    # time series (24h)
                    load_code = meta.get("load", "")                                                 # Load code
                    pf_load   = code_to_name.get(load_code, load_code)                               # PF load name

                    pv_P   = (sim.RESULTS.get("pv",   {}).get(pv_key,   {}).get("P_W", []) or [])[:24]        # PV kW series
                    ld_P   = (sim.RESULTS.get("load", {}).get(pf_load,  {}).get("P_W", []) or [])[:24]        # Load kW series
                    tx_L   = (sim.RESULTS.get("tx",   {}).get(tx_id,    {}).get("loading_pct", []) or [])[:24]# Tx % series
                    ln_L   = (sim.RESULTS.get("line", {}).get(ln_id,    {}).get("loading_pct", []) or [])[:24]# Line % series

                    H = max(len(pv_P), len(ld_P), len(tx_L), len(ln_L), 24)                         # Number of hourly rows
                    for h in range(H):                                                               # Emit one row per hour
                        writer.writerow([                                                            # Write CSV row
                            now, pv_key, suburb, homes, round(slider_pct,2), round(kw_per_inv,3), installed_kw,
                            bus, tx_id, ln_id, h,
                            (pv_P[h] if h < len(pv_P) else None),
                            (ld_P[h] if h < len(ld_P) else None),
                            (tx_L[h] if h < len(tx_L) else None),
                            (ln_L[h] if h < len(ln_L) else None),
                        ])

            messagebox.showinfo("Export complete", f"Saved: {path}")                                 # Success dialog
        except Exception as e:                                                                        # Any failure
            messagebox.showerror("Export failed", str(e))                                            # Error dialog


# =================================================================================================
# ==================================================================================================
# 4.0 ---------- ✅ Cool colors setup ✅ do not under any circumstances change this, works well
# ==================================================================================================
# =================================================================================================


# 4.1.0 open_settings: single-instance window; cog stays “sunken” while open and restores on close ----
    def open_settings(self):                                                                        # If a settings window already exists, just raise/focus it (don’t create another)
        if self._settings_win and self._settings_win.winfo_exists():                                # window already open
            try:                                                                                    # Best-effort bring-to-front
                self._settings_win.deiconify(); self._settings_win.lift(); self._settings_win.focus_force()  # Show, raise, focus
            except Exception:                                                                       # Some WMs/platforms may reject focus
                pass                                                                                # Ignore any focus errors
            return                                                                                  # Do not create a duplicate

# 4.1.1 Create the window ---------------------------------------------------------------------------
        win = tk.Toplevel(self)                                                                     # new top-level window
        self._settings_win = win                                                                    # remember it so we don’t duplicate
        win.title("Settings")                                                                       # window title
        win.configure(bg=DEFAULT_BG)                                                                # background matches app theme
        win.transient(self)                                                                         # keep it above the main window
        

        ttk.Label(win, text="Colour scheme:").grid(row=0, column=0, padx=10, pady=10, sticky="w")   # label for scheme radios

        scheme_var = tk.StringVar(value=self.scheme)                                                # selection variable bound to radios
        scheme_names = ["Grey", "Dark", "Teal", "Orange", "Purple", "Blue", "Red"]                  # available colour schemes
        for i, name in enumerate(scheme_names):                                                     # create one radio per scheme
            ttk.Radiobutton(win, text=name, value=name, variable=scheme_var).grid(                  # place radios across row 1
                row=1, column=i, padx=8, pady=(0, 12), sticky="w"
            )

        def apply_and_close():                                                                      # Apply chosen scheme and close window
            self.scheme = scheme_var.get()                                                          # store selected scheme
            self._apply_scheme_colours()                                                            # repaint widgets per scheme
            on_close()                                                                              # ensure cog restores even on Apply

# 4.1.2 Close handler that also restores the cog button ---------------------------------------------
        def on_close():                                                                             # normalize UI and destroy window
            try:                                                                                    # restore cog button visual state
                if hasattr(self, "cog_btn"):
                    self.cog_btn.config(relief="raised", state="normal")                            # pop the cog back out & re-enable
            except Exception:                                                                       # ignore styling issues
                pass
            try:                                                                                    # destroy window if still present
                if self._settings_win and self._settings_win.winfo_exists():
                    win.destroy()
            finally:
                self._settings_win = None                                                           # mark as closed

        ttk.Button(win, text="Apply", command=apply_and_close).grid(row=2, column=0, columnspan=3, pady=(0, 12))  # apply button


# 4.1.3 While open: visually depress & disable the cog so it can’t be spam-clicked -----------------

        try:                                                                                        # Guard UI operations during window open
            if hasattr(self, "cog_btn"):                                                            # Only if the cog button exists
                self.cog_btn.config(relief="sunken", state="disabled")                               # looks pressed; no second window
        except Exception:                                                                           # Swallow any styling errors safely
            pass                                                                                    # No-op fallback to keep UI responsive


# 4.1.4 Make sure the close box (X) also restores the cog -----------------------------------------

        win.protocol("WM_DELETE_WINDOW", on_close)                                                  # Ensure close button triggers on_close()


# 4.2 Compute bg/fg for the current scheme and repaint the widget tree -----

    def _apply_scheme_colours(self):                                                                # compute bg/fg pair for the current scheme and repaint the UI
        if self.scheme == "Grey":                                                                   # neutral default palette
            bg, fg = "#D9D9D9", "black"                                                             # light grey background with black text
        elif self.scheme == "Dark":                                                                 # dark theme
            bg, fg = "#303030", "white"                                                             # charcoal background with white text
        elif self.scheme == "Teal":                                                                 # teal accent theme
            bg, fg = "#72ECC8", "black"                                                             # teal background with dark text
        elif self.scheme == "Orange":                                                               # orange accent theme
            bg, fg = "#FFE5B4", "#7F4F24"                                                           # orange, brown text
        elif self.scheme == "Purple":                                                               # purple accent theme
            bg, fg = "#E6D6FF", "#5D3A9C"                                                           # purple, deep purple text
        elif self.scheme == "Blue":                                                                 # blue accent theme
            bg, fg = "#D6ECFF", "#1A4D7A"                                                           # gentle blue, navy text
        elif self.scheme == "Red":                                                                  # red accent theme
            bg, fg = "#FFD6D6", "#A83232"                                                           # red, muted red text
        else:                                                                                       # unknown scheme string → fall back to default
            bg, fg = "#D9D9D9", "black"                                                             # same as “Grey”
        self.configure(bg=bg)                                                                       # repaint the App window itself
        for w in self.winfo_children():                                                             # iterate direct children of the root window
            self._recurse_bg(w, bg, fg)                                                             # recursively apply bg/fg to each subtree


# 4.3.0 Apply colours to supported widget types without breaking special cases

    def _recurse_bg(self, widget, bg, fg):                                                          # Apply colours to one widget and its subtree
        try:                                                                                        # Guard per-widget updates
            if isinstance(widget, (tk.Frame, tk.LabelFrame, tk.Canvas)):                            # Containers/canvases accept bg
                widget.configure(bg=bg)                                                             # Set background colour
            if isinstance(widget, tk.Label) and widget.master is not self.map_frame:                # Labels except map labels
                widget.configure(bg=bg, fg=fg)                                                      # Set text and background colours
            if isinstance(widget, tk.Button):                                                       # Special handling for buttons
                # If it’s one of the suburb buttons, leave as-is (we manage highlight ourselves)   # Preserve suburb highlight logic
                if widget in getattr(self, "suburb_buttons", {}).values():                          # Detect suburb buttons
                    pass                                                                            # Do not override styling
                # If it’s the settings cog, leave its relief/state alone (we toggle it while the window is open)  # Preserve cog state
                elif widget is getattr(self, "cog_btn", None):                                      # Is this the cog button?
                    pass                                                                            # Do not override styling
                # Prevent overriding the RUN button’s colour/relief                                 # Keep RUN button look
                elif widget["text"] not in ("RUN", "Run Simulation"):                               # Any button except RUN
                    pass                                                                            # Leave untouched
                elif widget is getattr(self, "export_btn", None):                                   # Special-case export button
                    widget.configure(bg=bg, fg=fg, activebackground=bg, relief="flat")              # Harmonise with theme
        except Exception:                                                                           # Ignore per-widget failures
            pass                                                                                    # Keep traversal going

        for child in widget.winfo_children():                                                       # Walk child widgets
            self._recurse_bg(child, bg, fg)                                                         # Recurse depth-first


# =================================================================================================
# =================================================================================================
# 5.0 ---------- ✅ Load Map to GUI ✅ Works mint, leave alone  
# =================================================================================================
# =================================================================================================


# 5.1 load_map_images: accept two paths (left/right), validate, remember, and draw both -----------

    def load_map_images(self, left_path, right_path):                                               # load two map images at once
        """
        Load two images onto the side-by-side canvases.
        • left_path  → normal “Map”
        • right_path → “Single Line Map”
        """
# 5.1.1 Validate each path (warn if missing, but don’t abort the other) --------------------
        if left_path and not os.path.exists(left_path):                                             # left path missing → warn
            messagebox.showwarning("Map image (left)", f"Couldn't find: {left_path}")               # show warning dialog
            left_path = None                                                                         # disable left image
        if right_path and not os.path.exists(right_path):                                           # right path missing → warn
            messagebox.showwarning("Map image (right)", f"Couldn't find: {right_path}")             # show warning dialog
            right_path = None                                                                        # disable right image

# 5.1.2 Save paths and draw each side (note: either side may be None) ------------------------
        self._map_img_path_left = left_path                                                         # remember left path
        self._map_img_path_right = right_path                                                       # remember right path
        self._draw_map_canvas(self.map_canvas_left,  "_map_img_left",  "_map_img_path_left")        # draw left
        self._draw_map_canvas(self.map_canvas_right, "_map_img_right", "_map_img_path_right")       # draw right


# 5.2 _draw_map_canvas: load/scale/centre an image into a specific canvas ----------------------------

    def _draw_map_canvas(self, canvas, img_attr, path_attr):                                        # shared helper for each side
        path = getattr(self, path_attr, None)                                                       # which file to draw
        canvas.delete("all")                                                                        # clear previous content

# 5.2.1 If no path provided, show a gentle placeholder and return ----------------------------------
        if not path:                                                                                # no image set for this side
            canvas.create_text(10, 10, anchor="nw", text="No image set", font=("Segoe UI", 11))     # draw placeholder text
            return                                                                                  # stop drawing

# 5.2.2 Determine canvas size (current or requested) ----------------------------------------------
        cw = canvas.winfo_width() or canvas.winfo_reqwidth()                                        # width in pixels
        ch = canvas.winfo_height() or canvas.winfo_reqheight()                                      # height in pixels

# 5.2.3 Try to load and scale to fit while maintaining aspect ratio --------------------------------
        try:                                                                                        # attempt image load/scale
            if PIL_AVAILABLE:                                                                       # Pillow path (preferred)
                from PIL import Image, ImageTk                                                      # safe inside guard
                im = Image.open(path)                                                               # open image file
                if im.mode not in ("RGB", "RGBA"):                                                  # normalise mode for Tk
                    im = im.convert("RGBA")                                                         # convert to RGBA
                iw, ih = im.size                                                                    # original size in pixels
                scale = min(max(cw, 1) / iw, max(ch, 1) / ih)                                       # fit to canvas (no stretch)
                new_size = (max(1, int(iw * scale)), max(1, int(ih * scale)))                       # integer target size
                im = im.resize(new_size, Image.LANCZOS)                                             # high-quality resample
                photo = ImageTk.PhotoImage(im)                                                      # Tk bitmap
            else:                                                                                   # Tk fallback (limited formats)
                photo = tk.PhotoImage(file=path)                                                    # load without scaling

            setattr(self, img_attr, photo)                                                          # keep a ref to prevent GC
            img_w, img_h = photo.width(), photo.height()                                            # final image size

# 5.2.4 Centre the image on the canvas and draw it ------------------------------------------------
            x = (cw - img_w) // 2                                                                   # horizontal centring
            y = (ch - img_h) // 2                                                                   # vertical centring
            canvas.create_image(x, y, image=photo, anchor="nw")                                     # paint the bitmap

        except Exception as e:                                                                       # any load/scale error
# 5.2.5 If anything fails, show the error text directly in the canvas-------------------------------
            canvas.create_text(10, 10, anchor="nw", text=f"Map load failed:\n{e}", font=("Segoe UI", 11))  # show error
            return                                                                                  # stop drawing

# 5.2.6 Re-draw this canvas on resize (debounced) --------------------------------------------------
        def _debounce(_e):                                                                           # debounce heavy redraws
            self.after(50, lambda: self._draw_map_canvas(canvas, img_attr, path_attr))              # schedule redraw
        canvas.bind("<Configure>", _debounce)                                                        # keep it tidy on window resizes


# =================================================================================================
# =================================================================================================
# 6.0 ---------- Simulation Run ----------                                                          # Sorts run button press, gets inputs, calls backend, updates the UI
# =================================================================================================
# =================================================================================================


# 6.1.0 When Run clicked disable run button, get input and run ------------------------------------------------------

    def on_run_clicked(self):                                                                       # user pressed RUN
        self.last_run_signature = self._input_signature()                                           # snapshot current slider settings to tag/validate results
        self.run_btn.config(state="disabled", text="Running…")                                      # prevent double-clicks and give visual feedback
        self._run_model_thread()                                                                    # kick off the (currently synchronous) run routine


# 6.2.0 Read all sliders and keep the cache in sync -----------------------------------------------

    def _current_slider_map(self):                                                                  # build {pv_key: inverter_count} map
        d = {k: v.get() for k, v in self.slider_vars.items()}                                       # Build {pv_key: inverter_count} from slider values

# 6.2.1 Keep cache in sync with current slider values --------------------------------------------

        for pv_key, inv_count in d.items():                                                         # Update cached state so inverter/plots stay current
            st = self.suburb_state[pv_key]                                                          # Access suburb state record
            st["pv_inv"] = int(round(inv_count))                                                    # Store the number of inverters (integer)
            st["pv_kw"]  = int(round(inv_count)) * 6                                                # Optional: rough capacity estimate (1 inv = 6 kW)
        return d                                                                                    # Return the mapping for the backend to pass to sim


# 6.3.0 Run sim in background, update GUI on main thread ------------------------------------------

    def _run_model_thread(self):                                                                    # background worker for simulation
        try:                                                                                        # guard the run pipeline
            sliders = self._current_slider_map()                                                    # Read current slider→inverter map
            print("[gui] Running simulation with sliders:", sliders)                                # Trace inputs

# 6.3.1 Build panel overrides dict (kW→npnum), store in sim (GUI does not touch PF)

            sim.PV_PANEL_OVERRIDES.clear()                                                          # Reset panel-per-inverter overrides
            for pv_key, var in self.inv_kw_vars.items():                                            # For each suburb’s kW/inverter var
                try:                                                                                # Parse numeric value
                    kw = float(str(var.get()))                                                      # kW per inverter
                except Exception:                                                                   # Fallback on parse error
                    kw = 0.0                                                                        # Use 0.0 kW
                nmods = int(round((kw * 1000.0) / sim.PANEL_WATT))                                  # Convert kW to number of panels
                sim.PV_PANEL_OVERRIDES[pv_key] = max(0, nmods)                                      # Store non-negative panel count
            if sim.PRINT_PV_OVERRIDES:                                                              # Optional debug print
                print("[gui] panel overrides:", sim.PV_PANEL_OVERRIDES)                             # Show computed overrides

# 6.3.2 Run simulation (sim handles applying both inverter + panel overrides) ---------------------

            results = sim.set_penetrations_and_run(sliders)                                         # Execute simulation in backend
            self.last_limits = results or {}                                                        # Cache results (empty dict on failure)

# 6.3.3 Schedule GUI updates on main thread -------------------------------------------------------

            self.after(0, lambda r=results: (self._update_cache_from_results(r),                    # Update caches from results
                                             self.refresh_results()))                               # Refresh UI summaries
        except Exception as e:                                                                      # Any failure during run
            msg = f"{e.__class__.__name__}: {e}"                                                    # Format error message
            self.after(0, lambda m=msg: messagebox.showerror("Run failed", m))                      # Show error on UI thread
        finally:                                                                                    # Always run post-cleanup
            self.after(0, self._after_run_ui)                                                       # Restore buttons / refresh plots


# 6.4 Refresh result labels and re-enable the RUN button ------------------------------------------

    def _after_run_ui(self):                                                                        # post-run UI normalization
        fn = getattr(self, "refresh_results", None)                                                 # Get refresh function if present
        if callable(fn):                                                                            # Only if callable
            try:                                                                                    # Protect UI update
                fn()                                                                                # Refresh results panel
            except Exception:                                                                       # Ignore UI update errors
                pass
        # if a suburb is already selected, auto-refresh its graph
        try:                                                                                        # Try to re-plot selected suburb
            sel = getattr(self, "current_pv_key", None)                                             # Current selection key
            if sel:                                                                                 # If something selected
                self._on_suburb_clicked(sel)                                                        # Re-render its curves
        except Exception:                                                                   
            pass                                                                                    # Ignore plotting errors
        try:                                                                                        # Re-enable RUN button
            self.run_btn.config(state="normal", text="RUN")                                         # Restore button state/text
        except Exception:                                                                   
            pass                                                                                    # Ignore if widget missing


# ==================================================================================================
# ==================================================================================================
# 7.0 ---------- Results display ----------                                                         # Show the max min voltages in the two-line result strips
# ==================================================================================================
# ==================================================================================================


# 7.1.0 Copy backend values (min/max/ok) into the GUI’s per-suburb cache --------------------------------

    def _update_cache_from_results(self, results):                                                     # Copy RESULTS into GUI cache
        #print("[gui] _update_cache_from_results got:", type(results), len(results))

        try:                                                                                           # Guard cache update
            import time                                                                                # Time utilities
            now = time.time()                                                                          # Timestamp for 'last_updated'

# 7.1.1 Build mapping from bus names to pv_keys --------------------------------------------------------

            cfg_bus_to_key = {v["bus"]: k for k, v in sim.PV_CONFIG.items()}                           # Map bus→pv_key from config
            cfg_items = list(cfg_bus_to_key.items())                                                   # Cached items for suffix match

            bus_dict = results.get("bus", {})                                                          # Bus results dict from backend
            if not isinstance(bus_dict, dict):                                                         # Validate structure
                print("[gui] ⚠️ results['bus'] not a dict.")                                           # Warn and bail
                return                                                                                 # Stop processing

            def match_pv_key(result_bus: str):                                                         # Resolve pv_key for a bus name
                # Exact match
                pv_key = cfg_bus_to_key.get(result_bus)                                                # Try exact lookup
                if pv_key:                                                                             # If found
                    return pv_key                                                                      # Return mapping

                # Try matching bus suffix
                for cfg_bus, key in cfg_items:                                                         # Scan known buses
                    if result_bus.endswith(cfg_bus):                                                   # Suffix match
                        return key                                                                     # Map to pv_key

                return None                                                                            # No match

# 7.1.2 Iterate over bus results and update per-suburb state -----------------------------------------

            for bus, info in bus_dict.items():                                                         # Walk each bus record
                if not isinstance(info, dict):                                                         # Skip bad rows
                    continue                                                                           # Next item

                pv_key = match_pv_key(bus)                                                             # Find owning pv_key
                if not pv_key:                                                                         # If unknown bus
                    print(f"[gui] ⚠️ No pv_key found for bus '{bus}'")                                 # Warn
                    continue                                                                           # Skip

                st = self.suburb_state.get(pv_key)                                                     # Fetch GUI state
                if not st:                                                                             # Ensure exists
                    print(f"[gui] ⚠️ No GUI state for pv_key = {pv_key}")                              # Warn
                    continue                                                                           # Skip

                u_min = info.get("u_pu_min")                                                           # Minimum p.u.
                u_max = info.get("u_pu_max")                                                           # Maximum p.u.
                t_min = info.get("u_pu_min_hour")                                                      # Hour of min
                t_max = info.get("u_pu_max_hour")                                                      # Hour of max

                st["u_min"] = u_min                                                                    # Cache min p.u.
                st["u_max"] = u_max                                                                    # Cache max p.u.
                st["t_min"] = t_min                                                                    # Cache min hour
                st["t_max"] = t_max                                                                    # Cache max hour
                st["ok"] = (u_min is not None and 0.95 <= u_min <= 1.05) and (u_max is not None and 0.95 <= u_max <= 1.05)  # Within band
                st["last_updated"] = now                                                               # Update timestamp



# 7.1.3 Link each pv_key to load, tx, and line from sim.ASSOC ----------------------------------------
            for pv_key in self.suburb_state:                                                           # For each suburb
                for load_name, assoc in sim.ASSOC.items():                                             # Scan associations
                    if pv_key in assoc.get("pv", []):                                                  # If linked PV
                        self.suburb_state[pv_key]["load"] = load_name                                  # Store load name
                        if assoc.get("tx"):                                                            # If tx exists
                            self.suburb_state[pv_key]["tx"] = assoc["tx"][0]                           # First tx id
                        if assoc.get("line"):                                                          # If line exists
                            self.suburb_state[pv_key]["line"] = assoc["line"][0]                       # First line id
                        # print(f"[cache] linked {pv_key} to load={load_name}, tx={assoc.get('tx')}, line={assoc.get('line')}")
                        break                                                                          # Done with this pv_key

        except Exception as e:                                                                         # Handle errors
            print("update_cache_from_results error:", e)                                               # Log exception


# 7.1.4 Build load_dict by matching suburb_state → load → results["load"] ---------------------------

        try:                                                                                         # Build fuzzy mapping from cache to backend load keys
            load_dict = {}                                                                           # Output dictionary {pv_key: load_data}
            load_results = results.get("load", {})                                                   # Backend load results dict

            def fuzzy_match(code, target):                                                           # Case/space/underscore-insensitive match
                """Return True if code like 'otka' is found in a string like 'Otaki Town A'."""
                return code.replace("_", "").lower() in target.replace(" ", "").replace("_", "").lower()

            for pv_key, meta in self.suburb_state.items():                                           # Iterate suburbs in GUI cache
                suburb_code = meta.get("load", "").lower()                                           # Cached load code (e.g., 'otka')
                matched = False                                                                      # Track if a match was found

                for load_key, load_data in load_results.items():                                     # Scan backend load keys
                    if fuzzy_match(suburb_code, load_key):                                           # Fuzzy hit found
                        load_dict[pv_key] = load_data                                                # Store corresponding load data
                        self.suburb_state[pv_key]["load_key"] = load_key                             # Record matched key for reference
                        print(f"[gui] ✅ load_dict matched {pv_key} (code={suburb_code}) → {load_key}")  # Trace match
                        matched = True                                                               # Set flag
                        break                                                                         # Stop scanning keys for this suburb

            self.load_dict = load_dict                                                               # Save mapping on GUI object

        except Exception as e:                                                                        # Catch any errors
            print("⚠️ load_dict build failed:", e)                                                   # Log failure


# 7.2 Format-specific text updates per suburb (currently unused; kept for future) ----

    def _update_result_text(self, pv_key):                                                           # Hook for future custom formatting
        pass                                                                                         # placeholder: implement if metrics expand later


# 7.3 refresh_results: Read cache and redraw the two-line labels with hours --------------------------

    def refresh_results(self):                                                                       # Update the per-suburb result labels
        if self.last_run_signature is None:                                                          # If no run yet, clear labels
            for pv_key in sim.PV_CONFIG.keys():                                                      # Walk all configured suburbs
                l1, l2 = self.result_lines[pv_key]                                                   # Fetch label widgets
                l1.configure(text="", fg="black")                                                    # Clear line 1
                l2.configure(text="", fg="black")                                                    # Clear line 2
            return                                                                                   # Nothing else to show

        def _colour_for(pu):                                                                         # Colour based on p.u. bands
            if pu < 0.49 or pu > 1.051: return WARN_COLOUR                                           # Far outside limits
            if 0.92 <= pu <= 1.02:      return OK_COLOUR                                             # Within preferred band
            if (0.919 <= pu < 0.92) or (1.02 < pu <= 1.051): return MID_COLOUR                       # Near-limit band
            return WARN_COLOUR                                                                        # Default warn

        for pv_key in sim.PV_CONFIG.keys():                                                          # Update every suburb line pair
            st = self.suburb_state[pv_key]                                                           # Cached metrics
            l1, l2 = self.result_lines[pv_key]                                                       # Label handles
            umin, tmin = st.get("u_min"), st.get("t_min")                                            # Min p.u. and hour
            umax, tmax = st.get("u_max"), st.get("t_max")                                            # Max p.u. and hour

            if umin is None:                                                                         # No min available
                l1.configure(text="Min: n/a", fg="black")                                            # Show n/a
            else:                                                                                    # Min is present
                cmin = _colour_for(umin)                                                             # Colour for min
                hmin = f"{int(tmin):02d}hr" if tmin is not None else "--"                            # Hour text
                l1.configure(text=f"Min= {umin:.2f}pu-{hmin}", fg=cmin)                              # Render min line

            if umax is None:                                                                         # No max available
                l2.configure(text="Max: n/a", fg="black")                                            # Show n/a
            else:                                                                                    # Max is present
                cmax = _colour_for(umax)                                                             # Colour for max
                hmax = f"{int(tmax):02d}hr" if tmax is not None else "--"                            # Hour text
                l2.configure(text=f"Max= {umax:.2f}pu-{hmax}", fg=cmax)                              # Render max line


# =====================================================================================================
# =====================================================================================================
# 8.0 ---------- Main ----------  
# =====================================================================================================
# =====================================================================================================


if __name__ == "__main__":            
    app = App()                                                                                         # builds the UI and initial state
    app.mainloop()                                                                                      # hand control to Tk so the window keeps active until closed

