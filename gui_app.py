# gui_app.py
# ENGR489 PowerFactory model - Chris Working Code


# =====================================================================================================================================================
# =====================================================================================================================================================
# 1.0 Set UP
# =====================================================================================================================================================
# =====================================================================================================================================================


import os                                                   # Standard library
import threading                                            # Run long-running simulations off the main Tkinter thread
import tkinter as tk                                        # Main Tk GUI tools
from tkinter import ttk, messagebox                         # TK Themed widgets and dialog boxes

try:                                                        # Try to get better image handling if pillow is available
    from PIL import Image, ImageTk                          # Get pillow image handling tools
    PIL_AVAILABLE = True                                    # GUI can use Pillow
except Exception:                                           # If Pillow isn't installed or fails to import, don't loose your shit
    PIL_AVAILABLE = False                                   # Disable image nice bits but keep the app running

import otaki_sim as sim                                     # Backend code for PowerFactory logic and return results
from copy import deepcopy                                   # Keep independent copies of lists etc for the GUI storage
import time                                                 # Timeswarp baby


# 1.1 Plotting and Graphing setup
try:                                                        # Try to enable plotting inside the Tkinter window if Matplotlib is avaiable
    from matplotlib.figure import Figure                    # The plotting frame
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # Puts a frame into the Tk widget
    MPL_OK = True                                           # GUI can show charts if Matplotlib imported successfully
except Exception:                                           # If Matplotlib import fails, carry on without plots
    MPL_OK = False                                          # Disable graphing so the rest of the app still runs


# =====================================================================================================================================================
# =====================================================================================================================================================
# 2.0 UI constants and suburb display names
# =====================================================================================================================================================
# =====================================================================================================================================================

APP_TITLE = "Ōtaki Grid – Solar Penetration GUI"              # Show the title in the title bar
DEFAULT_BG = "#D9D9D9"                                      # Main background colour for the widgets (light grey)
OK_COLOUR = "#1E7D32"                                       # Colour used when voltages are within limits (green)
WARN_COLOUR = "#C62828"                                     # Colour used when voltages are out of limits (red)
ACCENT_GREEN = "#2ECC71"                                    # Accent colour for highlights, pressed buttons etc

# 2.1 Suburb Full Name list 
SUBURB_FULL = {                                               
    "OTBa_PV": "Ōtaki Beach A",                               # Label shown next to its slider
    "OTBb_PV": "Ōtaki Beach B",
    "OTBc_PV": "Ōtaki Beach C",
    "OTCa_PV": "Ōtaki Commercial A",
    "OTCb_PV": "Ōtaki Commercial B",
    "OTIa_PV": "Ōtaki Industrial",
    "OTKa_PV": "Ōtaki Township A",
    "OTKb_PV": "Ōtaki Township B",
    "OTKc_PV": "Ōtaki Township C",
    "OTS_PV" : "Ōtaki College",
    "RGUa_PV": "Rangiuru A",
    "RGUb_PV": "Rangiuru B",
    "TRE_PV" : "Te Roto",
    "WTVa_PV": "Waitohu Valley A",
    "WTVb_PV": "Waitohu Valley B",
    "WTVc_PV": "Waitohu Valley C",
}

# 2.2 Suburb Variable Defauls 
SUBURB_DEFAULTS = {
    "pv_pct": 0.0,                                          # Slider value as a percentage of rooftop PV penetration 
    "pv_kw": 0.0,                                           # Computed nstalled PV (kW) for that suburb
    "u_min": None,                                          # Minimum pu bus voltage over the day
    "u_max": None,                                          # Maximum pu bus voltage over the day
    "ok": None,                                             # Limit flag (True=in limits, False=out of limits) for quick colouring
    "load_curve": None,                                     # Time series of load (kW) for plotting if provided by backend
    "pv_profile": None,                                     # Time series of PV output (kW) for plotting if provided by backend
    "last_updated": None,                                   # Timestamp of the last time this suburb’s results were written to the storage
    "bus": None,                                            # The associated PowerFactory bus name for this suburb
}

# 2.3 Result labels used for the results box ---------------------------------------------------------------------------
SHOW_METRIC_SELECTOR = False                                # Drop down menu for graph, set true to enable, not sure if needed

METRIC_OPTIONS = [
    "Voltage p.u. (min over day)",                          # Show the minimum pu voltage across the day
    "Voltage p.u. (max over day)",                          # Show the maximum pu voltage across the day
    "OK/Limit flag",                                        # Show a status from backend (True/Green, False/red)
]


# =====================================================================================================================================================
# =====================================================================================================================================================
# 4.0 ---------- UI setup ----------
# =====================================================================================================================================================
# =====================================================================================================================================================


class App(tk.Tk):

# 3.1  Create the main window, build layout, wire widgets etc
    def __init__(self):
        super().__init__()

# 3.1.1 Window basics
        self.title(APP_TITLE)                                       # Window title
        self.configure(bg=DEFAULT_BG)                               # Main ackground colour
        self.geometry("1400x900")                                   # Initial size
        self.minsize(1100, 700)                                     # Minimum usable size

# 3.1.2 Style / colour scheme
        self.scheme = "Grey"                                        # Settings color scheme
        self.style = ttk.Style(self)                                # ttk theme thingy
        self.style.theme_use("default")                             # Use the default ttk theme

# 3.1.3 Main layout: Horizontal split windows, left for the sliders and run button, right top bottom for the graph and map
        main_pane = tk.PanedWindow(self, orient="horizontal", 
        sashrelief="raised", bg=DEFAULT_BG)                         # Make a horizontal Window 
        main_pane.pack(fill="both", expand=True)                    # pack it so it fills up the whole window and grows/shrinks with resizing

# 3.1.4 Left: suburb slider panel (with its own vertical scroll)    # left-hand window with the slider list and RUN button
        left_frame = tk.Frame(main_pane, bg=DEFAULT_BG)             # Builds a smaller frame inside the left side window
        main_pane.add(left_frame, minsize=350)                      # add the frame to the window; make a minimum width so it doesn’t go too small

# 3.1.5 RUN row (bottom-left): square cog on the left + wide RUN button on the right
        run_row = tk.Frame(left_frame, bg=DEFAULT_BG)                   # container strip at the bottom
        run_row.pack(side="bottom", fill="x", pady=8)                   # stick to the bottom of the left panel

        # Use grid so the cog can be a fixed-width square and the RUN can expand
        run_row.grid_columnconfigure(0, weight=0, minsize=54)           # col 0 = cog slot (~square; tweak minsize if needed)
        run_row.grid_columnconfigure(1, weight=1)                       # col 1 = RUN button expands to fill remaining width
        run_row.grid_rowconfigure(0, weight=1)                          # let buttons stretch to the row height

# 3.1.5.1 Settings cog (square, same vertical height as RUN)
        self.cog_btn = tk.Button(
            run_row,
            text="⚙",                                                   # gear glyph
            font=("Segoe UI Symbol", 14),                               # nice crisp icon size
            bg=DEFAULT_BG, activebackground=DEFAULT_BG, fg="black",     # neutral look; scheme handler may recolour later
            relief="raised", borderwidth=4, highlightthickness=2,       # match the visual weight of RUN
            width=2, height=1,                                          # narrow width; the row height makes it appear square
            command=self.open_settings                                  # same handler as before
        )
        self.cog_btn.grid(row=0, column=0, sticky="nsew", 
                          padx=(6, 6), pady=6)                          # fill its cell; padding mirrors RUN

# 3.1.5.2 RUN button (big, expands across the rest of the row)
        self.run_btn = tk.Button(
            run_row, text="RUN", font=("Segoe UI", 13, "bold"),
            bg=ACCENT_GREEN, activebackground=ACCENT_GREEN, fg="white",
            relief="raised", borderwidth=4, highlightthickness=2, padx=24, pady=6,
            command=self.on_run_clicked
        )
        self.run_btn.grid(row=0, column=1, sticky="nsew", 
                          padx=(0, 6), pady=6)                          # stretch horizontally; same vertical padding as cog
        self._settings_win = None  # <-- track the one live settings window (None means closed)

# 3.1.6 Scrollable suburbd window area
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical")    # Vertical scrollbar for the left panel
        scrollbar.pack(side="left", fill="y")                       # lock the scrollbar on the far left and stretch it vertically

        canvas = tk.Canvas(left_frame, bg=DEFAULT_BG, 
                           highlightthickness=0)                    # canvas will host the scrolling content window
        canvas.pack(side="left", fill="both", expand=True)          # place canvas to the right of the scrollbar, make it resize with the window

        canvas.configure(yscrollcommand=scrollbar.set)              # canvas tells the scrollbar where the view is (for the thumb position)
        scrollbar.config(command=canvas.yview)                      # scrollbar controls the vertical view of the canvas when dragged

        self.sliders_frame = tk.LabelFrame(canvas, text="Solar " \
        "Penetration (%)", bg=DEFAULT_BG)                           # inner frame that actually holds slider rows
        self.sliders_frame_id = canvas.create_window((0, 0), 
        window=self.sliders_frame, anchor="nw")                     # put the inner frame at (0,0) inside the canvas

# 3.1.6.1 Grid widths across 4 columns where each slider row is laid out  
        self.sliders_frame.columnconfigure(0, weight=0, minsize=0)    # column 0: suburb button, fixed width
        self.sliders_frame.columnconfigure(1, weight=0, minsize=0)    # column 1: kW label, fixed width
        self.sliders_frame.columnconfigure(2, weight=0, minsize=0)    # column 2: % entry, fixed width
        self.sliders_frame.columnconfigure(3, weight=0, minsize=190)  # column 3: results labels, 190 px so the text won't be cramped

# 3.1.6.2 Update scrollable area whenever content size changes  
        def on_frame_config(event):                                 # callback fired when sliders_frame changes size like adding rows
            canvas.configure(scrollregion=canvas.bbox("all"))       # set the scrollable bit to fit all the content
        self.sliders_frame.bind("<Configure>", on_frame_config)     # wire the size-change event to the callback

# 3.1.6.3 Mouse wheel setup                                         # enable mouse wheel scrolling over the canvas
        def _on_mousewheel(event):                                  # wheel data
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 
                                "units")                            # scroll up/down by 1 unit per notch
        canvas.bind_all("<MouseWheel>", _on_mousewheel)             # listen for mouse wheel events anywhere in the app and route to this canvas


# 3.1.7 Suburb slider rows and their state storage
        self.slider_vars     = {}                                   # slider control  (0 to 100%)
        self.percent_entries = {}                                   # Entry widget that mirrors the slider percentage text
        self.result_lines    = {}                                   # Top and bottom result line
        self.result_labels   = {}                                   # Alias to the first result line Label (convenience)
        self.suburb_buttons  = {}                                   # Button used to select that suburb plot its bits

        row_base = 0                                                # Each suburb uses two rows

# 3.1.8 Make sure App-wide state storage exists before building the rows 
        self.suburb_state = {}                                      # Storage holding current state for that suburb
        self.last_run_signature = None                              # Snapshot of the current slider values

# 3.1.8.1 Build one two-row slider per suburb
        for pv_key in sim.PV_CONFIG.keys():                         # Go through all configured PV objects from the backend
            self._build_slider_row(self.sliders_frame, 
                                   row_base, pv_key)                # Build the two-row UI block for the suburb
            row_base += 2                                           # Go forward two rows for the next suburb underneath
#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
# 3.1.9 Load the storage with initial slider values and bus names
        # NOTE: This duplicates self.suburb_state creation; kept to match current logic.
        from copy import deepcopy                                   # import here (local) to emphasise we’re copying the defaults per suburb
        self.suburb_state = {}                                      # reset cache to a fresh dict (overwrites the earlier empty dict)
        for pv_key, meta in sim.PV_CONFIG.items():                  # iterate suburbs again to initialise their cached state
            item = deepcopy(SUBURB_DEFAULTS)                        # start from a clean, independent copy of the defaults
            item["pv_pct"] = float(self.slider_vars[pv_key].get())  # read the current slider (initially 0.0)
            item["pv_kw"]  = int(round(item["pv_pct"])) * 6         # derive a coarse kW capacity: 6 kW per 1% slider step
            item["bus"]    = meta["bus"]                            # remember the associated LV bus name from backend config
            item["load_curve"] = item.get("load_curve", None)       # keep placeholders for plotting if/when data arrives
            item["pv_profile"] = item.get("pv_profile", None)       # keep placeholders for plotting if/when data arrives
            self.suburb_state[pv_key] = item                        # write the initialised state back into the cache

# 3.1.9.1 Track last input signature separately (re-set after building)
        self.last_run_signature = None                              # make sure no old signature is left after initial UI build


# 3.1.10 Right: graphs + map stacked vertically                     # right-hand window stacks the graph on top and map on bottom
        right_frame = tk.Frame(main_pane, bg=DEFAULT_BG)            # Frame that sits in the right side of main window
        main_pane.add(right_frame)                                  # add the right frame as the second child of the main window

        vertical_pane = tk.PanedWindow(right_frame, 
        orient="vertical", sashrelief="raised", bg=DEFAULT_BG)      # vertical split for graphs/map
        vertical_pane.pack(fill="both", expand=True)                # let the vertical pane fill the whole right side and resize with the window

# 3.1.11 Graph frame (top)                                          # labelled frame for the pgraph area
        self.graphs_frame = tk.LabelFrame(vertical_pane, 
        text="Graphs", bg=DEFAULT_BG)                 # top section title 
        self.metric_var = tk.StringVar(value=METRIC_OPTIONS[0])     # hidden default selection; no UI shown
        vertical_pane.add(self.graphs_frame, minsize=100)           # add to the vertical split, don’t let it shrink below ~100 px

# 3.1.11.1 Graph area (Matplotlib embedded in Tk)                   # Put in a Matplotlib canvas if available
        if MPL_OK:                                                  # only build plotting widgets if Matplotlib was imported successfully
        
            from matplotlib.figure import Figure as _Figure         # type: ignore
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg as _FigureCanvasTkAgg  # type: ignore

            self.fig = Figure(figsize=(5, 3), dpi=100)              # Reseve a Figure object (overall plot canvas)
            self.ax = self.fig.add_subplot(111)                     # Put a single subplot (1x1 grid, first axes)
            self.ax.set_title("Select a suburb to display curves")  # First up prompt
            self.ax.set_xlabel("Hour")                              # x-axis label for time in hours
            self.ax.set_ylabel("kW")                                # y-axis label for power values
            self.ax.grid(True, linestyle="--", linewidth=0.5)       # dashed grid
            self.canvas_mpl = FigureCanvasTkAgg(self.fig, 
            master=self.graphs_frame)                               # Help Matplotlib talk to the Tk widget
            self.canvas_mpl.get_tk_widget().pack(fill="both", 
            expand=True, padx=8, pady=(0, 8))                       # pack the canvas to fill the frame
        else:
# 3.1.11.2 Text if matplotlib not found               
            self.graph_fallback = tk.Label(                         # Label explaining how to hopfully enable plots 
                self.graphs_frame, text="Matplotlib not installed.\nInstall it to see plots.",
                bg=DEFAULT_BG
            )
            self.graph_fallback.pack(fill="both", 
            expand=True, padx=8, pady=(0, 8))                           # centre the message

# 3.1.12 Map frame (middle) — fixed 50/50 split (no adjustable sash)
        self.map_frame = tk.LabelFrame(vertical_pane, text="Maps", bg=DEFAULT_BG)           # labelled container for both maps
        vertical_pane.add(self.map_frame, minsize=100)                                       # add to the vertical split

# 3.1.12.1 Fixed split using a simple Frame + grid (two equal columns)
        maps_row = tk.Frame(self.map_frame, bg=DEFAULT_BG)                                   # inner container row
        maps_row.pack(fill="both", expand=True)                                              # occupy all available space

# 3.1.12.2 Make two equal-width columns that always share space 50/50
        maps_row.grid_columnconfigure(0, weight=1, uniform="maps")                           # left column stretches equally
        maps_row.grid_columnconfigure(1, weight=1, uniform="maps")                           # right column stretches equally
        maps_row.grid_rowconfigure(0, weight=1)                                              # row stretches vertically too

# 3.1.12.3 Left canvas (distribution/normal map)
        self.map_canvas_left = tk.Canvas(maps_row, bg="#BFBFBF", highlightthickness=0)       # left image area
        self.map_canvas_left.grid(row=0, column=0, sticky="nsew", padx=(0, 2), pady=0)       # fill left cell; tiny gap between maps

# 3.1.12.4 Right canvas (transmission single-line map)
        self.map_canvas_right = tk.Canvas(maps_row, bg="#BFBFBF", highlightthickness=0)      # right image area
        self.map_canvas_right.grid(row=0, column=1, sticky="nsew", padx=(2, 0), pady=0)      # fill right cell

# 3.1.12.5 Image refs/paths so Tk doesn’t GC the bitmaps
        self._map_img_left = None                                                            # PhotoImage for left canvas
        self._map_img_right = None                                                           # PhotoImage for right canvas
        self._map_img_path_left = None                                                       # file path for left image
        self._map_img_path_right = None                                                      # file path for right image

# 3.1.13 Default pane heights — keep so maps are visible on startup
        self.after(80, lambda: vertical_pane.paneconfig(self.graphs_frame, height=600))      # graphs top ~220 px
        self.after(80, lambda: vertical_pane.paneconfig(self.map_frame,    height=100))      # maps middle ~400 px

# 3.1.13.1 Load both maps (absolute paths)
        self.load_map_images(
            r"C:\Users\chris\OneDrive - Victoria University of Wellington - STUDENT\Vic\ENGR489\Artifact\Map.png",               # left: normal map
            r"C:\Users\chris\OneDrive - Victoria University of Wellington - STUDENT\Vic\ENGR489\Artifact\Single Line Map.png"    # right: single-line transmission map
        )

# 3.1.14 Metric selector (combobox) for results boxes — DISABLED by default
        if SHOW_METRIC_SELECTOR:
            metric_bar = tk.Frame(self.graphs_frame, bg=DEFAULT_BG)     # horizontal container
            metric_bar.pack(side="top", fill="x", padx=8, pady=8)       # only created if flag True
            ttk.Label(metric_bar, text="Display:").pack(side="left")
            self.metric_combo = ttk.Combobox(
                metric_bar,
                values=METRIC_OPTIONS,
                textvariable=self.metric_var,
                state="readonly",
                width=28
            )
            self.metric_combo.pack(side="left", padx=8)
            self.metric_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_results())

# 3.2 _check_pv_objects_on_startup: Query backend to see which PV objects exist in the PF model
    def _check_pv_objects_on_startup(self):                         # runs once at startup to sanity-check that each PV object exists in the PowerFactory model
# 3.2.1  Try to get PV object presence from backend
        try:                                                        # attempt to connect to PowerFactory and ask the backend which PVs are present
            import otaki_sim                                        # import here (local) so the GUI file can still import even if PF isn’t installed
            app = otaki_sim.connect_and_activate()                  # ensure PF is running and the correct Project/Study Case is active; returns the app handle
            _, missing_pvs = otaki_sim.get_pv_objects(app)          # query backend for PV objects; returns (present_dict, missing_list); we only need missing_list
        except Exception:                                           # if anything fails (no PF, wrong path, no project), fail safe
# 3.2.2 If backend fails, pessimistically mark all as missing
            missing_pvs = list(sim.PV_CONFIG.keys())                # treat every configured PV as “missing” so the UI makes that obvious to the user

# 3.2.2 Paint each suburb’s first result line as Found/Missing
        for pv_key, meta in sim.PV_CONFIG.items():                  # iterate every suburb key that the GUI knows about
            l1, l2 = self.result_lines[pv_key]                      # grab the two result label widgets for this suburb (top/bottom)
            if pv_key in missing_pvs:                               # if backend said this PV object is missing from the model
                l1.configure(text="Missing", fg=WARN_COLOUR)        # red  # show “Missing” in warning red on the first result line
            else:                                                   # otherwise it exists in the currently active PF study case
                l1.configure(text="Found", fg=OK_COLOUR)            # green  # show “Found” in OK green
            l2.configure(text="", fg="black")                       # clear the second line for a clean startup state (no stale text/colour)

# 3.3 _input_signature: Build a deterministic snapshot of sliders to compare runs
    def _input_signature(self):                                     # returns a compact, hashable record of current slider settings to tag results with
# 3.3.1 Deterministic tuple of (pv_key, pct_int) sorted by key
        items = sorted((k, int(round(v.get()))) for k, 
                       v in self.slider_vars.items())               # build (key, integer_percent) pairs and sort for stability
        return tuple(items)                                         # make it immutable so we can compare cached signatures before updating result boxes

# 3.4 _build_slider_row: Construct one two-row slider+labels group for a suburb
    def _build_slider_row(self, parent, row_base, pv_key):          # parent: container widget; row_base: top row index; pv_key: suburb backend key
        """
        Two-row group starting at 'row_base'.

        Row A (row_base):
            col 0: suburb button (click → show curves)
            col 1: kW box
            col 2: % box (aligned with slider right edge)
            col 3: result line 1 (just past slider)

        Row B (row_base+1):
            col 0–2: slider spans columns 0–2 fully
            col 3: result line 2
        """
        full_name = SUBURB_FULL.get(pv_key, pv_key)                 # look up human-readable name for this suburb; fall back to key if missing

#3.4.1 ROW A: Suburb button (click to plot curves for this suburb)
        name_btn = tk.Button(                                       # create a left-aligned button with the suburb’s name
            parent, text=full_name, bg=DEFAULT_BG, anchor="w",      # put text on the left; match app background
            relief="flat", padx=0,                                  # flat style blends with the panel; no extra horizontal padding
            command=lambda k=pv_key: self._on_suburb_clicked(k)     # clicking triggers plotting/highlight for this suburb
        )
        name_btn.grid(row=row_base, column=0, sticky="w", 
                      padx=(6, 0), pady=(4, 0))                     # place in row A, col 0, with small left/top padding
        self.suburb_buttons[pv_key] = name_btn                      # keep a handle so we can recolour/depress it when selected

# 3.4.2 kW label (shows derived capacity = int(%)*6 kW)
        kw_lbl = tk.Label(parent, bd=1, relief="sunken", 
                          bg="white", fg="black",                   # sunken box look for a “readout” feel
                          width=8, padx=6, pady=2, anchor="e")      # fixed width; right-aligned numbers; small padding
        kw_lbl.grid(row=row_base, column=1, sticky="e", 
                    padx=(2, 0), pady=(4, 0))                       # place in row A, col 1; hug the right edge
        self.kw_labels = getattr(self, 'kw_labels', {})             # ensure the dict exists (first time through)
        self.kw_labels[pv_key] = kw_lbl                             # store the label so _update_kw_label can target it later

# 3.4.3 % Entry box (text entry stays in sync with the slider)
        ent = ttk.Entry(parent, width=6, justify="right")           # small numeric entry field; right-aligned text
        ent.insert(0, "0")                                          # show 0 as the initial percentage
        ent.grid(row=row_base, column=2, sticky="e", 
                 padx=(2, 0), pady=(4, 0))                          # place in row A, col 2; align to the right
        self.percent_entries[pv_key] = ent                          # remember this entry to keep it synchronised with the slider

# 3.4.4 Results line 1 (top result strip)
        res_line1 = tk.Label(parent, bd=1, relief="sunken", 
                             bg="white", fg="black",                # create a readout-style label (sunken border) for the first results line
                             padx=6, pady=2, anchor="w")            # small internal padding; left-align text within the label
        res_line1.grid(row=row_base, column=3, sticky="nsew", 
                       padx=(8, 6), pady=(4, 0))                    # place in row A, col 3; allow it to stretch; add side/top padding
        self.result_labels[pv_key] = res_line1                      # keep a reference so other methods can update this label

# 3.4.5 ROW B: slider + result line 2 (bottom)
        var = tk.DoubleVar(value=0.0)                               # per-suburb variable backing the slider (float 0..100)
        self.slider_vars[pv_key] = var                              # store the variable so we can read/update it elsewhere

# 3.4.6 When the slider moves: sync % entry, show % in result line, update kW label, and keep placeholder PV capacity in step
        slider = ttk.Scale(                                         # create a horizontal slider for PV penetration %
            parent, from_=0, to=100, orient="horizontal", 
            variable=var,                                           # range 0..100%; bound to the DoubleVar above
            command=lambda _=None, k=pv_key: (                      # on movement, run a small pipeline of UI updates
                self._update_percent_entry(k),                      # 1) mirror slider value into the % Entry field
                self._show_percent_in_results(k),                   # 2) paint the live % onto result line 1 (non-destructive)
                self._update_kw_label(k),                           # 3) recompute and show kW capacity (= int(%) * 6)
                self._touch_placeholder_capacity(k)                 # 4) keep cached pv_pct/pv_kw in sync for plotting
            )
        )
        slider.grid(row=row_base + 1, column=0, columnspan=3,       # place the slider on row B spanning columns 0..2
                    sticky="ew", padx=(6, 0), pady=(0, 4))          # stretch horizontally; add side/bottom padding

# 3.4.7 Sync entry → slider (user types a number; normalise to 0..100 and push back to slider)
        def commit_entry(*_):                                       # handler to commit typed % back into the slider/state
            try:
                val = float(ent.get())                              # parse the Entry text as a number
            except ValueError:
                val = 0.0                                           # on bad input, fall back to 0
            val = max(0.0, min(100.0, val))                         # clamp to valid range 0..100
            var.set(val)                                            # update the slider’s DoubleVar (moves the slider)
            self._update_percent_entry(pv_key)                      # keep Entry text formatted (e.g., "37%")
            self._show_percent_in_results(pv_key)                   # refresh live % on result line 1
            self._update_kw_label(pv_key)                           # recompute kW label
            self._touch_placeholder_capacity(pv_key)                # sync cached pv_pct/pv_kw for plots

        ent.bind("<Return>", commit_entry)                          # pressing Enter commits the typed value
        ent.bind("<FocusOut>", commit_entry)                        # leaving the field also commits the value

# 3.4.8 Results line 2 (bottom result strip)
        res_line2 = tk.Label(parent, bd=1, relief="sunken", 
                             bg="white", fg="black",                
                             padx=6, pady=2, anchor="w")            # second readout-style label for additional results                
        res_line2.grid(row=row_base + 1, column=3, sticky="nsew",   # place in row B, col 3; allow it to stretch
                       padx=(8, 6), pady=(0, 4))                    # side/bottom padding to match the row’s spacing
        self.result_lines[pv_key] = (res_line1, res_line2)          # store both labels as a tuple for easy updates later

# 3.4.9  Seed initial state for this suburb’s UI bits
        self._update_percent_entry(pv_key)                          # initialise the % Entry to match the slider (e.g., "0%")
        self._show_percent_in_results(pv_key)                       # paint the initial % on line 1 without clearing anything
        self._update_kw_label(pv_key)                               # compute and show initial kW ("0 kW")
        self._touch_placeholder_capacity(pv_key)                    # set initial cached pv_pct/pv_kw so graphs can be drawn

# 3.5 _normal_button_colours: Provide platform-default colours to reset suburb buttons
    def _normal_button_colours(self):                               # utility: returns a consistent set of colours to “unselect” any suburb button
        """
        Return a dict of platform-default colours for tk.Button so we can restore
        buttons to normal when deselected.
        """
        try:
            return {"bg": "SystemButtonFace", "fg": "black", 
                    "activebackground": "SystemButtonFace"}         # Windows/ttk default palette
        except Exception:
# 3.5.1 Fallback for themes without SystemButtonFace
            return {"bg": DEFAULT_BG, "fg": "black", 
                    "activebackground": DEFAULT_BG}                 # fallback to app bg if theme doesn’t expose system colours

# 3.6 _update_kw_label: Show calculated kW capacity based on rounded slider %
    def _update_kw_label(self, pv_key):                             # update the small “kW” readout next to the % entry for a given suburb
        """Update the kW label above the percent entry."""
        pct = float(self.slider_vars[pv_key].get())                 # read the current slider value (float 0..100)
        pct_int = int(round(pct))                                   # round to nearest whole percent for display/derivation
        kw_val = pct_int * 6                                        # convert percent to coarse capacity using 6 kW per 1% rule
        lbl = self.kw_labels.get(pv_key)                            # get the label widget that shows the kW value
        if lbl:                                                     # if the label exists (it should)
            lbl.config(text=f"{kw_val} kW")                         # update the text to show the computed capacity

# 3.7 _update_percent_entry: Keep the % Entry text synced with the slider ( 0 to 100)
    def _update_percent_entry(self, pv_key):                        # mirror the slider percentage into the tiny Entry field as “NN%”
        """Sync the tiny % Entry with the slider (integer 0..100)."""
        pct = int(round(float(self.slider_vars[pv_key].get())))     # take the slider float and round to an integer percent
        ent = self.percent_entries.get(pv_key)                      # fetch the Entry widget that displays the %
        if not ent:                                                 # safety: if it wasn’t created yet, bail
            return
        pct_str = f"{pct}%"                                         # format as a human-friendly percent string
        if ent.get() != pct_str:                                    # only rewrite if the text actually changed (avoids cursor jumps)
            ent.delete(0, "end")                                    # clear current entry text
            ent.insert(0, pct_str)                                  # insert the fresh “NN%” string

# 3.8 _show_percent_in_results: Paint the current % onto result line 1 (non-destructive)
    def _show_percent_in_results(self, pv_key):                     # optionally echo the live % on the first result line without wiping old results
        """Before/while running, show % on line 1, do not clear previous results."""
        # NOTE: You can implement the live % display here; currently a no-op placeholder.
        # Example (uncomment to enable):
        # l1, _ = self.result_lines[pv_key]                         # fetch the first (top) results label for this suburb
        # pct = int(round(float(self.slider_vars[pv_key].get())))   # read and round the slider’s current value
        # l1.configure(text=f"{pct}%", fg="black")                  # show “NN%” in neutral black to indicate a pending change
        pass                                                        # placeholder so the method is callable even if no live echo is desired

# 3.9 _ensure_placeholder_curves: Seed/load simple 24-point load & PV arrays for plotting
    def _ensure_placeholder_curves(self, pv_key):                   # ensure the suburb has basic curves to plot if real data isn’t available yet
        """
        If we don’t have real data from PowerFactory yet, seed placeholder 24-point arrays.
        - Load: morning + evening peaks
        - PV: bell/half-sine from ~06:00–18:00, scaled by current slider %
        """
        import math                                                 # local import for maths functions (exp, sin, pi)
        st = self.suburb_state[pv_key]                              # grab the mutable state dict for this suburb

# 3.9.1 24 hours
        hours = list(range(24))                                     # simple 0..23 hour index (used for consistency; plotting uses range(24) directly later)

# 3.9.2 Placeholder load curve (kW): base + peaks (07–09, 18–21)
        if st["load_curve"] is None:                                # only create the synthetic load once (unless later replaced with real data)
            base = [2.0] * 24                                       # start with a flat 2 kW base across all hours
            for h in range(24):                                     # iterate each hour 0..23
                morning = 2.5 * math.exp(-0.5 * 
                                         ((h - 8) / 1.8) ** 2)      # Gaussian-like morning peak centred ~08:00
                evening = 3.0 * math.exp(-0.5 * 
                                         ((h - 19) / 2.2) ** 2)     # Gaussian-like evening peak centred ~19:00
                base[h] += morning + evening                        # add peaks to the base for this hour
            st["load_curve"] = base                                 # list of 24 floats representing kW load per hour

# 3.9.3 Placeholder PV profile (kW): half-sine from 06–18, scaled by pv_kw “capacity”
        if st["pv_profile"] is None:                                # only create synthetic PV once (unless we need to rescale later)
            pv_cap = max(0.0, float(st.get("pv_kw", 0.0)))          # take the current capacity proxy from state; clamp to non-negative
            pv = [0.0] * 24                                         # start with zeros for all hours
            for h in range(6, 19):                                  # daylight window inclusive of 06:00..18:00
                x = (h - 6) / 12.0                                  # normalise to 0..1 across the daylight window
                pv[h] = pv_cap * math.sin(math.pi * x)              # half-sine from 0→peak→0 scaled by capacity
            st["pv_profile"] = pv                                   # store the 24-point PV profile

# 3.9.4 Re-scale PV if user changed slider since last seed
        else:                                                       # if a profile already exists, rebuild it using the updated capacity (keeps shape, adjusts magnitude)
            pv_cap = max(0.0, float(st.get("pv_kw", 0.0)))          # recalc capacity proxy from current slider-derived kW
            pv = [0.0] * 24                                         # fresh zeroed list
            for h in range(6, 19):                                  # same daylight window
                x = (h - 6) / 12.0                                  # 0..1 normalised position in the day
                pv[h] = pv_cap * math.sin(math.pi * x)              # recompute half-sine amplitude
            st["pv_profile"] = pv                                   # write the updated PV profile back to state

# 3.10 _on_suburb_clicked: Highlight selection, sync state, ensure curves, draw plot
    def _on_suburb_clicked(self, pv_key):                           # user clicked a suburb name; make it look selected and show its curves
        """Handle suburb button click: ensure data, then draw one suburb’s curves + visual highlight."""
        if not MPL_OK:                                              # if plotting support isn’t available, inform the user and bail gracefully
            messagebox.showinfo("Graphs", "Matplotlib not installed. Install it to see plots.")
            return

# 3.10.1 Visually indicate selection: reset all, then depress + recolour clicked
        normal = self._normal_button_colours()                      # get default button colours to restore non-selected buttons
        for btn in self.suburb_buttons.values():                    # iterate all suburb buttons and reset their styling
            try:
                btn.config(relief="raised", bg=normal["bg"], fg=normal["fg"],
                           activebackground=normal["activebackground"])
            except Exception:
                pass                                                # some themes/platforms may not support all options; ignore safely

        try:
            # Highlight colour (blue) for the selected suburb       # emphasise the chosen suburb
            self.suburb_buttons[pv_key].config(
                relief="sunken", bg="#1976D2", fg="white", activebackground="#1976D2"
            )
        except Exception:                                           # fallback palette if the primary colour fails (unlikely)
            self.suburb_buttons[pv_key].config(
                relief="sunken", bg="#4A90E2", fg="white", activebackground="#4A90E2"
            )

# 3.10.2 Keep state in sync with the current slider
        st = self.suburb_state[pv_key]                              # grab the cached state for this suburb
        st["pv_pct"] = float(self.slider_vars[pv_key].get())        # read current slider % (float 0..100)
        st["pv_kw"]  = int(round(st["pv_pct"])) * 6                 # derive kW capacity using 6 kW per 1% rule-of-thumb

# 3.10.3 Ensure placeholder curves exist (until PowerFactory is wired up)
        self._ensure_placeholder_curves(pv_key)                     # seed or rescale synthetic load/PV curves based on current capacity

# 3.10.4 Draw the selected suburb’s curves
        self._draw_suburb_curves(pv_key)                            # clear the axes and plot this suburb’s load and PV profiles

# 3.11 _draw_suburb_curves: Render load/PV curves for selected suburb onto the axes
    def _draw_suburb_curves(self, pv_key):                           # draw the currently selected suburb’s load and PV profiles
        st = self.suburb_state[pv_key]                                # fetch cached state (contains curves and labels) for this suburb
        full_name = SUBURB_FULL.get(pv_key, pv_key)                   # resolve human-readable suburb name for the plot title

        h = list(range(24))                                           # x-axis: 0..23 hours
        load = st.get("load_curve") or [0.0] * 24                     # y-series 1: load curve (fallback to zeros if missing)
        pv   = st.get("pv_profile") or [0.0] * 24                     # y-series 2: PV curve (fallback to zeros if missing)

        self.ax.clear()                                               # wipe any previous plot content
        self.ax.grid(True, linestyle="--", linewidth=0.5)             # light dashed grid for readability
        self.ax.plot(h, load, label="Load (kW)")                      # plot load vs hour with legend label
        self.ax.plot(h, pv,   label="PV (kW)")                        # plot PV vs hour with legend label
        self.ax.set_xlim(0, 23)                                       # clamp x-axis to the 24-hour range
        self.ax.set_xticks([0, 4, 8, 12, 16, 20, 23])                 # choose legible tick marks across the day
        self.ax.set_xlabel("Hour")                                    # axis label for hours
        self.ax.set_ylabel("kW")                                      # axis label for power in kW
        self.ax.set_title(full_name)                                  # set the figure title to the suburb’s name
        self.ax.legend(loc="best")                                    # show legend in the best available spot
        self.canvas_mpl.draw_idle()                                   # ask Tkinter-embedded canvas to redraw without blocking

# 3.12 _touch_placeholder_capacity: Keep cached pv_pct/pv_kw aligned with the slider
    def _touch_placeholder_capacity(self, pv_key):                    # ensure the cache mirrors current slider % and derived kW
        """
        Keep cached pv_pct/pv_kw aligned with the slider.
        Safe to call during widget construction (auto-creates dict entry).
        """
# 3.12.1 ensure dict exists (paranoia)
        if not hasattr(self, 
                       "suburb_state") or self.suburb_state is None:  # if the cache isn’t initialised for some reason
            self.suburb_state = {}                                    # create an empty dict to hold per-suburb state

# 3.12.2 get/create this suburb's state (minimal baseline)
        st = self.suburb_state.get(pv_key, {                          # look up existing state or build a minimal one on the fly
            "pv_pct": 0.0,                                            # current slider percentage
            "pv_kw":  0.0,                                            # derived coarse capacity in kW
            "u_min": None, "u_max": None, "ok": None,                 # placeholders for backend voltage/flag results
            "load_curve": None, "pv_profile": None,                   # placeholders for time-series used in plots
            "bus": sim.PV_CONFIG.get(pv_key, {}).get("bus")           # convenience pointer to the associated LV bus name
        })

# 3.12.3 update from current slider 
        pct = float(self.slider_vars[pv_key].get())                   # read the live slider value as a float
        st["pv_pct"] = pct                                            # store the percent back into the state cache
        st["pv_kw"]  = int(round(pct)) * 6                            # convert to whole-percent then to kW (6 kW per 1%)

        self.suburb_state[pv_key] = st                                # write updated state back to the cache


# =====================================================================================================================================================
# =====================================================================================================================================================
# 4.0 ---------- Cool colors setup ----------
# =====================================================================================================================================================
# =====================================================================================================================================================


# 4.1 open_settings: single-instance window; cog stays “sunken” while open and restores on close
    def open_settings(self):
        # If a settings window already exists, just raise/focus it (don’t create another)
        if self._settings_win and self._settings_win.winfo_exists():                 # window already open
            try:
                self._settings_win.deiconify(); self._settings_win.lift(); self._settings_win.focus_force()
            except Exception:
                pass
            return

# 4.1.1 Create the window
        win = tk.Toplevel(self)                                                      # new top-level window
        self._settings_win = win                                                     # remember it so we don’t duplicate
        win.title("Settings")
        win.configure(bg=DEFAULT_BG)
        win.transient(self)                                                          # keep it above the main window
        

        ttk.Label(win, text="Colour scheme:").grid(row=0, column=0, padx=10, pady=10, sticky="w")

        scheme_var = tk.StringVar(value=self.scheme)
        scheme_names = ["Grey", "Dark", "Teal", "Orange", "Purple", "Blue", "Red"]
        for i, name in enumerate(scheme_names):
            ttk.Radiobutton(win, text=name, value=name, variable=scheme_var).grid(
                row=1, column=i, padx=8, pady=(0, 12), sticky="w"
            )

        def apply_and_close():
            self.scheme = scheme_var.get()
            self._apply_scheme_colours()
            on_close()                                                               # ensure cog restores even on Apply

# 4.1.2 Close handler that also restores the cog button
        def on_close():
            try:
                if hasattr(self, "cog_btn"):
                    self.cog_btn.config(relief="raised", state="normal")             # pop the cog back out & re-enable
            except Exception:
                pass
            try:
                if self._settings_win and self._settings_win.winfo_exists():
                    win.destroy()
            finally:
                self._settings_win = None                                            # mark as closed

        ttk.Button(win, text="Apply", command=apply_and_close).grid(row=2, column=0, columnspan=3, pady=(0, 12))

# 4.1.3 While open: visually depress & disable the cog so it can’t be spam-clicked
        try:
            if hasattr(self, "cog_btn"):
                self.cog_btn.config(relief="sunken", state="disabled")               # looks pressed; no second window
        except Exception:
            pass

# 4.1.4 Make sure the close box (X) also restores the cog
        win.protocol("WM_DELETE_WINDOW", on_close)

# 4.2 _apply_scheme_colours: Compute bg/fg for the current scheme and repaint the widget tree
    def _apply_scheme_colours(self):                                # compute bg/fg pair for the current scheme and repaint the UI
        if self.scheme == "Grey":                                   # neutral default palette
            bg, fg = "#D9D9D9", "black"                           # light grey background with black text
        elif self.scheme == "Dark":                                 # dark theme
            bg, fg = "#303030", "white"                           # charcoal background with white text
        elif self.scheme == "Teal":                                 # teal accent theme
            bg, fg = "#72ECC8", "black"                           # teal background with dark text
        elif self.scheme == "Orange":                               # orange accent theme
            bg, fg = "#FFE5B4", "#7F4F24"                       # orange, brown text
        elif self.scheme == "Purple":                               # purple accent theme
            bg, fg = "#E6D6FF", "#5D3A9C"                       # purple, deep purple text
        elif self.scheme == "Blue":                                 # blue accent theme
            bg, fg = "#D6ECFF", "#1A4D7A"                       # gentle blue, navy text
        elif self.scheme == "Red":                                  # red accent theme
            bg, fg = "#FFD6D6", "#A83232"                       # red, muted red text
        else:                                                       # unknown scheme string → fall back to default
            bg, fg = "#D9D9D9", "black"                           # same as “Grey”
        self.configure(bg=bg)                                       # repaint the App window itself
        for w in self.winfo_children():                             # iterate direct children of the root window
            self._recurse_bg(w, bg, fg)                             # recursively apply bg/fg to each subtree

# 4.3 _recurse_bg: Depth-first traversal to apply colours to supported widget types without breaking special cases
    def _recurse_bg(self, widget, bg, fg):
        try:
            if isinstance(widget, (tk.Frame, tk.LabelFrame, tk.Canvas)):
                widget.configure(bg=bg)
            if isinstance(widget, tk.Label) and widget.master is not self.map_frame:
                widget.configure(bg=bg, fg=fg)
            if isinstance(widget, tk.Button):
                # If it’s one of the suburb buttons, leave as-is (we manage highlight ourselves)
                if widget in getattr(self, "suburb_buttons", {}).values():
                    pass
                # If it’s the settings cog, leave its relief/state alone (we toggle it while the window is open)
                elif widget is getattr(self, "cog_btn", None):
                    pass
                # Prevent overriding the RUN button’s colour/relief
                elif widget["text"] not in ("RUN", "Run Simulation"):
                    widget.configure(bg=bg, fg=fg, activebackground=bg, relief="flat")
        except Exception:
            pass

        for child in widget.winfo_children():
            self._recurse_bg(child, bg, fg)



# # =====================================================================================================================================================
# # =====================================================================================================================================================
# # 5.0 ---------- Load Map to GUI ----------  
# # =====================================================================================================================================================
# # =====================================================================================================================================================

# 5.1 load_map_images: accept two paths (left/right), validate, remember, and draw both
    def load_map_images(self, left_path, right_path):                            # load two map images at once
        """
        Load two images onto the side-by-side canvases.
        • left_path  → normal “Map”
        • right_path → “Single Line Map”
        """
        # 5.1.1 Validate each path (warn if missing, but don’t abort the other)
        if left_path and not os.path.exists(left_path):
            messagebox.showwarning("Map image (left)", f"Couldn't find: {left_path}")
            left_path = None
        if right_path and not os.path.exists(right_path):
            messagebox.showwarning("Map image (right)", f"Couldn't find: {right_path}")
            right_path = None

        # 5.1.2 Save paths and draw each side (note: either side may be None)
        self._map_img_path_left = left_path
        self._map_img_path_right = right_path
        self._draw_map_canvas(self.map_canvas_left,  "_map_img_left",  "_map_img_path_left")   # draw left
        self._draw_map_canvas(self.map_canvas_right, "_map_img_right", "_map_img_path_right")  # draw right

# 5.2 _draw_map_canvas: load/scale/centre an image into a specific canvas
    def _draw_map_canvas(self, canvas, img_attr, path_attr):                     # shared helper for each side
        path = getattr(self, path_attr, None)                                    # which file to draw
        canvas.delete("all")                                                     # clear previous content

# 5.2.1 If no path provided, show a gentle placeholder and return
        if not path:
            canvas.create_text(10, 10, anchor="nw", text="No image set", font=("Segoe UI", 11))
            return

# 5.2.2 Determine canvas size (current or requested)
        cw = canvas.winfo_width() or canvas.winfo_reqwidth()                     # width in pixels
        ch = canvas.winfo_height() or canvas.winfo_reqheight()                   # height in pixels

# 5.2.3 Try to load and scale to fit while maintaining aspect ratio
        try:
            if PIL_AVAILABLE:
                from PIL import Image, ImageTk                                   # safe inside guard
                im = Image.open(path)                                            # open image file
                if im.mode not in ("RGB", "RGBA"):                               # normalise mode for Tk
                    im = im.convert("RGBA")
                iw, ih = im.size                                                 # original size in pixels
                scale = min(max(cw, 1) / iw, max(ch, 1) / ih)                    # fit to canvas (no stretch)
                new_size = (max(1, int(iw * scale)), max(1, int(ih * scale)))    # integer target size
                im = im.resize(new_size, Image.LANCZOS)                          # high-quality resample
                photo = ImageTk.PhotoImage(im)                                   # Tk bitmap
            else:
                # Tk fallback (PNG/GIF best); no scaling available here
                photo = tk.PhotoImage(file=path)

            setattr(self, img_attr, photo)                                       # keep a ref to prevent GC
            img_w, img_h = photo.width(), photo.height()                         # final image size

# 5.2.4 Centre the image on the canvas and draw it
            x = (cw - img_w) // 2                                                # horizontal centring
            y = (ch - img_h) // 2                                                # vertical centring
            canvas.create_image(x, y, image=photo, anchor="nw")                  # paint the bitmap

        except Exception as e:
# 5.2.5 If anything fails, show the error text directly in the canvas
            canvas.create_text(10, 10, anchor="nw", text=f"Map load failed:\n{e}", font=("Segoe UI", 11))
            return

# 5.2.6 Re-draw this canvas on resize (debounced)
        def _debounce(_e):
            self.after(50, lambda: self._draw_map_canvas(canvas, img_attr, path_attr))
        canvas.bind("<Configure>", _debounce)                                     # keep it tidy on window resizes


# =====================================================================================================================================================
# =====================================================================================================================================================
# 6.0 ---------- Simulation Run ----------                          # Sorts run button press, gets inputs, calls backend, updates the UI
# =====================================================================================================================================================
# =====================================================================================================================================================


# 6.1 on_run_clicked: Disable the button, capture input signature, and start the run
    def on_run_clicked(self):                                           # user pressed RUN
        self.last_run_signature = self._input_signature()               # snapshot current slider settings to tag/validate results
        self.run_btn.config(state="disabled", text="Running…")          # prevent double-clicks and give visual feedback
        self._run_model_thread()                                        # kick off the (currently synchronous) run routine

# 6.2 _current_slider_map: Read all sliders and keep the cache in sync
    def _current_slider_map(self):
        d = {k: v.get() for k, v in self.slider_vars.items()}           # build {pv_key: percent_float} from all DoubleVars
# 6.2.1 keep cache in sync with current slider values
        for pv_key, pct in d.items():                                   # update per-suburb cached state so kW/plots stay current
            st = self.suburb_state[pv_key]
            st["pv_pct"] = float(pct)                                   # store the slider %
            st["pv_kw"]  = int(round(float(pct))) * 6                   # derive coarse capacity (6 kW per 1%)
        return d                                                        # return the mapping for the backend call

# 6.3 _run_model_thread: Call backend with current sliders, process results, handle errors
    def _run_model_thread(self):
        try:
            sliders = self._current_slider_map()                        # gather inputs and sync cache
            results = sim.set_penetrations_and_run(sliders)             # run the Ōtaki simulation in the backend
            self.last_limits = results                                  # stash raw results for any later inspection
            self._update_cache_from_results(results)                    # write min/max/ok (etc.) into GUI cache (expects this helper to exist)

# 6.3.1 Check for missing PV objects in results
            missing_pvs = []                                            # legacy scan: looks for “reason: Missing PV object: <key>”
            for bus, info in results.items():                           # iterate top-level items (only useful if backend returned legacy shape)
                if isinstance(info, dict) and 'reason' in info and 'Missing PV object' in info['reason']:
                    missing_pvs.append(info['reason'].split(': ')[-1])  # extract pv_key from the reason string

# 6.3.2 After: results = sim.set_penetrations_and_run(sliders)
            missing_pvs = (results or {}).get("missing_pv", [])         # preferred new format: backend provides an explicit list
            if missing_pvs:                                             # if any PV objects were not found in the PF model
                self.after(0, lambda: messagebox.showwarning(           # show a non-blocking warning popup on the Tk event loop
                    "Missing PV Objects",
                    "These PV objects are missing in the PowerFactory model:\n" + ", ".join(missing_pvs)
                ))

        except Exception as e:                                          # any exception during run → show an error dialog (non-blocking)
            self.after(0, lambda: messagebox.showerror("Run failed", str(e)))
        finally:
            self.after(0, self._after_run_ui)                           # regardless of success/failure, re-enable the UI safely via Tk thread

# 6.4 _after_run_ui: Refresh result labels and re-enable the RUN button
    def _after_run_ui(self):
        self.refresh_results()                                          # repaint the per-suburb result strips based on updated cache
        self.run_btn.config(state="normal", text="RUN")                 # re-enable the RUN button and restore its label



# =====================================================================================================================================================
# =====================================================================================================================================================
# 7.0 ---------- Results display ----------                             # Show the max min voltages in the two-line result strips
# =====================================================================================================================================================
# =====================================================================================================================================================


# 7.1 _update_cache_from_results: Copy backend values (min/max/ok) into the GUI’s per-suburb cache
    def _update_cache_from_results(self, results):
        """
        Make the GUI cache match the backend results.
        - New backend: bus limits live under results["limits"].
        - Legacy fallback: if no "limits", assume top-level bus keys.
        """
        import time                                                     # local import for simple timestamping
        now = time.time()                                               # epoch seconds; used to mark when a suburb’s cache was last updated

# 7.1.1 Map bus -> pv_key once (PV_CONFIG comes from otaki_sim)
        bus_to_pv = {v["bus"]: k for k, v in sim.PV_CONFIG.items()}     # reverse map: LV bus name → suburb key (pv_key)

# 7.1.2 Prefer the new-place limits, else fall back to old top-level shape
        limits_dict = (results or {}).get("limits")                     # new backend: results contains a "limits" dict keyed by bus
        if not isinstance(limits_dict, dict):                           # if missing or wrong type, assume legacy shape (top-level bus keys)
            limits_dict = results or {}                                 # tolerate None by treating it as {}

        for bus, info in limits_dict.items():                           # iterate each bus → info block (expects dicts with u_min/u_max/ok)
            if not isinstance(info, dict):                              # skip anything that isn’t a dict (defensive)
                continue                                                # ignore non-dict entries
            pv_key = bus_to_pv.get(bus)                                 # find which suburb this bus corresponds to
            if not pv_key:                                              # if the bus isn’t in our PV_CONFIG mapping, ignore it
                continue                                                # unknown bus → skip
            st = self.suburb_state[pv_key]                              # pull the cached state record for this suburb
            st["u_min"] = info.get("u_min")                             # write minimum per-unit voltage if present
            st["u_max"] = info.get("u_max")                             # write maximum per-unit voltage if present
            st["ok"]    = info.get("ok", True)                          # default to True (OK) if backend didn’t include a flag
            st["last_updated"] = now                                    # remember when we last updated this suburb’s results

# 7.2 _update_result_text: (reserved) Format-specific text updates per suburb (currently unused; kept for future)
    def _update_result_text(self, pv_key):
        pass                                                            # placeholder: hook for custom text formatting if metrics become more complex later

# 7.3 refresh_results: Read cache + metric selection and repaint the two-line labels for each suburb
    def refresh_results(self):
        metric = self.metric_var.get()                                  # current selection from the “Display” combobox (not yet branching on it)
        if self.last_run_signature is None:                             # before any simulation, clear all result labels
            for pv_key in sim.PV_CONFIG.keys():                         # iterate all configured suburbs
                l1, l2 = self.result_lines[pv_key]                      # fetch the two label widgets
                l1.configure(text="", fg="black")                       # clear top line
                l2.configure(text="", fg="black")                       # clear bottom line
            return                                                      # nothing else to do until a run occurs

        for pv_key, meta in sim.PV_CONFIG.items():                      # per-suburb redraw
            st = self.suburb_state[pv_key]                              # stored values for this suburb
            l1, l2 = self.result_lines[pv_key]                          # top/bottom result labels

            if st["u_min"] is None and st["u_max"] is None:             # no data for this suburb yet
                l1.configure(text="no data", fg="black")                # show a no data message
                l2.configure(text="", fg="black")                       # blank second line
                continue                                                # move to next suburb

            # Choose colour from limit flag if present
            col = OK_COLOUR if st.get("ok", True) else WARN_COLOUR      # green when OK; red when out-of-limits

            if st["u_min"] is not None:                                 # if there is a minimum value
                l1.configure(text=f"Min = {st['u_min']:.4f} pu", fg=col)# render min with 4 dp in current colour scheme
            else:
                l1.configure(text="Min: n/a", fg="black")               # draw n/a when missing

            if st["u_max"] is not None:                                 # if there is a maximum value
                l2.configure(text=f"Max = {st['u_max']:.4f} pu", fg=col)# render max with 4 dp in current olour scheme
            else:
                l2.configure(text="Max: n/a", fg="black")               # draw n/a when missing


# =====================================================================================================================================================
# =====================================================================================================================================================
# 8.0 ---------- Main ----------  
# =====================================================================================================================================================
# =====================================================================================================================================================


if __name__ == "__main__":            
    app = App()                                                         # builds the UI and initial state
    app.mainloop()                                                      # hand control to Tk so the window keeps active until closed


# In the next few years, the number of installed solar projects in New Zealand is set to increase 
# dramatically (See Transpower’s System Security Forecast 2024). The New Zealand grid has not 
# experienced high levels of intermittent power supply, which does make a number of stakeholders 
# very nervous. In this project, the student will create a model in which the installed solar 
# capacity can be altered and the voltage profile of the power system can be compared. It will also 
# potentially explore the role batteries may play in correcting voltage and reactive power problems 
# faced by the grid. The idea of the project is to create a tool, which can help a user understand 
# the implications of installing a set capacity of solar, or perhaps understanding what the maximum 
# capacity could be before the installation "breaks" the local grid. You will create a model in 
# Powerfactory and ideally a script which can be used to control the Powerfactory model. The model 
# will be of a specific location (to be selected by the student and supervisor) and will capture 
# enough of the surrounding power network to understand the system behavior. The input variables 
# of interest will be (but more might be included) minimum and maximum test solar capacity and step 
# size, while the output will by voltage level at the various nodes, power output, MW and MVAr.
# Ideally, a battery bank will also be installed, which can be disconnected and reconnected to observe 
# how this fill effect the voltage profile 