"""
Graph Coloring via Simulated Annealing — Tkinter desktop version.

Assignment 2 requirements covered:
  1. Graph representation  -> GraphBench (adjacency list `self.edges`, click-to-build GUI)
  2. Initial solution      -> random_hues(): random color per vertex from a chosen palette size
  3. Simulated annealing   -> anneal_step()/heat_at(): configurable initial temperature,
                              cooling schedule (geometric / linear / logarithmic), iteration budget
  4. Output                -> final coloring drawn on the canvas + a vertex/color table +
                              total conflict count

Also supports loading a graph + SA parameters from an Excel file
(graph_input.xlsx) via the "Load Excel" button, using pandas + openpyxl.
That's the only extra dependency beyond the standard library — no
scikit-learn, no matplotlib.
"""

import math
import os
import random
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

INK_PALETTE = ["#a6392e", "#2f7d76", "#c98a2c", "#4a4e69", "#7a5c3e", "#3c3744"]
COLOR_NAMES = ["Red", "Teal", "Amber", "Indigo", "Umber", "Charcoal"]

BG_PAPER = "#f2ede0"
BG_CARD = "#ece5d3"
BG_CANVAS = "#fbf9f2"
INK = "#232017"
INK_SOFT = "#6b6350"
RULE = "#cfc4a8"
DANGER = "#a6392e"
TEAL = "#2f7d76"
INDIGO = "#4a4e69"


class GraphBench:
    """Holds the graph itself: vertices, edges, and the current coloring."""

    def __init__(self):
        self.vertices = []   # list of dict(id, x, y)
        self.edges = []      # list of (id_a, id_b)
        self.next_id = 0
        self.hues = {}       # vertex id -> color index

    def vertex_near(self, x, y, radius=16):
        best = None
        best_d = radius
        for v in self.vertices:
            d = math.hypot(v["x"] - x, v["y"] - y)
            if d < best_d:
                best_d = d
                best = v
        return best

    def linked(self, a, b):
        return any((e[0] == a and e[1] == b) or (e[0] == b and e[1] == a) for e in self.edges)

    def clash_count(self, hues=None):
        hues = hues if hues is not None else self.hues
        return sum(1 for a, b in self.edges if hues.get(a) == hues.get(b))

    def random_hues(self, k):
        return {v["id"]: random.randrange(k) for v in self.vertices}

    def add_vertex(self, x, y, label=None):
        vid = self.next_id
        self.next_id += 1
        # `label` is what's drawn/displayed (e.g. "A" from an Excel sheet);
        # `id` stays a plain internal integer used for edges/hues, exactly
        # as before. Defaulting label to str(vid) keeps existing manual /
        # sample-graph behavior pixel-for-pixel unchanged.
        self.vertices.append({"id": vid, "x": x, "y": y, "label": label if label is not None else str(vid)})
        return vid

    def remove_vertex(self, vid):
        self.vertices = [v for v in self.vertices if v["id"] != vid]
        self.edges = [e for e in self.edges if vid not in e]
        # Small Issue 2 fix: don't leave a stale color entry behind.
        self.hues.pop(vid, None)

    def toggle_edge(self, a, b):
        if self.linked(a, b):
            self.edges = [e for e in self.edges if not ((e[0] == a and e[1] == b) or (e[0] == b and e[1] == a))]
        else:
            self.edges.append((a, b))

    def remove_edge_near(self, x, y, tolerance=14):
        closest, closest_d = None, tolerance
        for e in self.edges:
            va = next((v for v in self.vertices if v["id"] == e[0]), None)
            vb = next((v for v in self.vertices if v["id"] == e[1]), None)
            if not va or not vb:
                continue
            mx, my = (va["x"] + vb["x"]) / 2, (va["y"] + vb["y"]) / 2
            d = math.hypot(mx - x, my - y)
            if d < closest_d:
                closest_d, closest = d, e
        if closest:
            self.edges.remove(closest)
            return True
        return False


def heat_at(step, t0, rate, kind, total_steps):
    if kind == "linear":
        return max(0.01, t0 * (1 - step / total_steps))
    if kind == "log":
        return t0 / (1 + rate * 20 * math.log(1 + step))
    return t0 * (rate ** step)  # geometric


# ===========================================================================
# SECTION: GRAPH INPUT FROM EXCEL
# Lets the graph (and the SA parameters) be loaded from graph_input.xlsx
# instead of clicked in by hand. Plain parsing / graph construction with no
# Tkinter in it, so it's reusable and easy to test on its own. Errors are
# raised as ValueError with a human-readable message; the GUI layer
# (App.on_load_excel_click) turns those into a messagebox.
# ===========================================================================

GRAPH_INPUT_FILE = "graph_input.xlsx"

REQUIRED_PARAMETERS = {
    "Number_of_Colors": int,
    "Initial_Temperature": float,
    "Cooling_Rate": float,
    "Iterations": int,
}


def load_excel_graph(filename=GRAPH_INPUT_FILE):
    """Read the Vertices and Edges sheets and return (vertex_ids, edge_pairs).

    vertex_ids : list[str]              -- e.g. ["A", "B", "C", "D", "E"]
    edge_pairs : list[tuple[str, str]]  -- e.g. [("A","B"), ("A","C"), ...]

    Raises FileNotFoundError if the file is missing, and ValueError (with a
    message meant to be shown directly to the user) for anything malformed:
    missing sheets, missing columns, blank vertex ids, or an edge that
    references a vertex that isn't in the Vertices sheet.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"'{filename}' was not found in the current folder.")

    try:
        workbook = pd.ExcelFile(filename, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"'{filename}' doesn't look like a valid Excel file.\n({exc})")

    for sheet in ("Vertices", "Edges"):
        if sheet not in workbook.sheet_names:
            raise ValueError(f"'{filename}' is missing the required '{sheet}' sheet.")

    vertices_df = workbook.parse("Vertices")
    if "Vertex_ID" not in vertices_df.columns:
        raise ValueError("The 'Vertices' sheet must have a 'Vertex_ID' column.")
    vertex_ids = [str(v).strip() for v in vertices_df["Vertex_ID"].dropna().tolist()]
    if not vertex_ids:
        raise ValueError("The 'Vertices' sheet has no vertex ids in it.")
    if len(set(vertex_ids)) != len(vertex_ids):
        raise ValueError("The 'Vertices' sheet has duplicate Vertex_ID values.")

    edges_df = workbook.parse("Edges")
    if not {"From", "To"}.issubset(edges_df.columns):
        raise ValueError("The 'Edges' sheet must have 'From' and 'To' columns.")
    edges_df = edges_df.dropna(subset=["From", "To"])
    edge_pairs = [(str(a).strip(), str(b).strip()) for a, b in zip(edges_df["From"], edges_df["To"])]

    known = set(vertex_ids)
    for a, b in edge_pairs:
        if a not in known or b not in known:
            bad = a if a not in known else b
            raise ValueError(f"The 'Edges' sheet references vertex '{bad}', which isn't in the 'Vertices' sheet.")

    return vertex_ids, edge_pairs


def load_parameters(filename=GRAPH_INPUT_FILE):
    """Read the Parameters sheet and return a dict with the four required
    Simulated Annealing parameters, cast to the right numeric type."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"'{filename}' was not found in the current folder.")

    try:
        params_df = pd.read_excel(filename, sheet_name="Parameters", engine="openpyxl")
    except ValueError:
        raise ValueError(f"'{filename}' is missing the required 'Parameters' sheet.")

    if not {"Parameter", "Value"}.issubset(params_df.columns):
        raise ValueError("The 'Parameters' sheet must have 'Parameter' and 'Value' columns.")

    raw = dict(zip(params_df["Parameter"], params_df["Value"]))
    parameters = {}
    for name, cast in REQUIRED_PARAMETERS.items():
        if name not in raw or pd.isna(raw[name]):
            raise ValueError(f"The 'Parameters' sheet is missing a value for '{name}'.")
        try:
            parameters[name] = cast(raw[name])
        except (TypeError, ValueError):
            raise ValueError(f"'{name}' in the 'Parameters' sheet must be a number (got {raw[name]!r}).")

    if not (2 <= parameters["Number_of_Colors"] <= 6):
        raise ValueError("'Number_of_Colors' must be between 2 and 6.")
    if parameters["Cooling_Rate"] <= 0:
        raise ValueError("'Cooling_Rate' must be a positive number.")
    if parameters["Iterations"] <= 0:
        raise ValueError("'Iterations' must be a positive number.")

    return parameters


def build_graph(bench, vertex_ids, edge_pairs, center_x=320, center_y=200, radius=150):
    """Populate `bench` (a GraphBench) from parsed vertex labels and edge
    pairs, laying vertices out evenly around a circle so the loaded graph is
    immediately readable. Mutates `bench` in place and also returns it."""
    bench.vertices = []
    bench.edges = []
    bench.hues = {}
    bench.next_id = 0

    n = len(vertex_ids)
    label_to_id = {}
    for i, label in enumerate(vertex_ids):
        angle = (2 * math.pi * i / n) - (math.pi / 2) if n else 0.0
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        vid = bench.add_vertex(x, y, label=label)
        label_to_id[label] = vid

    for a_label, b_label in edge_pairs:
        a_id, b_id = label_to_id[a_label], label_to_id[b_label]
        if a_id == b_id or bench.linked(a_id, b_id):
            continue  # skip self-loops and exact duplicate edges
        bench.edges.append((a_id, b_id))

    return bench


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Graph Coloring on the Annealing Bench")
        self.configure(bg=BG_PAPER)
        self.geometry("1180x900")

        self.bench = GraphBench()
        self.mode = "sketch"
        self.pending_link = None
        self.dragged = None

        # SA run state. Cooling rate / schedule / step budget are deliberately
        # NOT cached anywhere else — they're read live from their widgets on
        # every step (see _live_params), so pausing and changing them takes
        # effect the moment you resume instead of sticking to old values.
        self.running = False
        self.tick = 0
        self.heat = 25.0
        self.best_hues = None
        self.best_clashes = float("inf")
        self.history = []   # list of (step, cur, best)
        self.after_id = None

        self._build_ui()
        self._bind_bench_events()
        self._refresh_all()

    # ---------------------------------------------------------- UI building
    def _build_ui(self):
        header = tk.Frame(self, bg=BG_CARD, highlightbackground=INK, highlightthickness=1)
        header.pack(fill="x", padx=16, pady=(14, 10))
        tk.Label(header, text="Graph Coloring on the Annealing Bench", bg=BG_CARD, fg=INK,
                 font=("Georgia", 18, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        tk.Label(
            header,
            text="Sketch a graph, choose a palette size, then heat it and let it cool under a "
                 "schedule of your choosing.",
            bg=BG_CARD, fg=INK_SOFT, font=("Courier New", 10), wraplength=1000, justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 10))

        top = tk.Frame(self, bg=BG_PAPER)
        top.pack(fill="both", expand=False, padx=16)
        top.columnconfigure(0, weight=3)
        top.columnconfigure(1, weight=2)

        # ---- Bench card (left) ----
        bench_card = self._card(top, "Bench")
        bench_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        toolbar = tk.Frame(bench_card, bg=BG_CARD)
        toolbar.pack(fill="x", pady=(0, 8))
        self.mode_buttons = {}
        for key, label in [("sketch", "Sketch vertices"), ("link", "Link vertices"),
                            ("shift", "Shift positions"), ("erase", "Erase")]:
            b = tk.Button(toolbar, text=label, command=lambda k=key: self.set_mode(k),
                          bg=BG_PAPER, fg=INK, relief="solid", bd=1, font=("Courier New", 9, "bold"))
            b.pack(side="left", padx=(0, 6))
            self.mode_buttons[key] = b
        tk.Button(toolbar, text="Load sample", command=self.load_sample,
                  bg=BG_PAPER, fg=INK, relief="solid", bd=1, font=("Courier New", 9, "bold")).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="Load Excel", command=self.on_load_excel_click,
                  bg=BG_PAPER, fg=INK, relief="solid", bd=1, font=("Courier New", 9, "bold")).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="Wipe bench", command=self.wipe_bench,
                  bg=BG_PAPER, fg=INK, relief="solid", bd=1, font=("Courier New", 9, "bold")).pack(side="left")

        self.bench_canvas = tk.Canvas(bench_card, width=650, height=410, bg=BG_CANVAS,
                                       highlightbackground=INK, highlightthickness=1, cursor="crosshair")
        self.bench_canvas.pack(fill="both", expand=True)

        self.mode_caption = tk.Label(bench_card, text="", bg=BG_CARD, fg=INK_SOFT,
                                      font=("Courier New", 9), anchor="w", justify="left")
        self.mode_caption.pack(fill="x", pady=(6, 0))

        # ---- Controls card (right) ----
        ctrl_card = self._card(top, "Annealing schedule")
        ctrl_card.grid(row=0, column=1, sticky="nsew")

        self.k_var = tk.IntVar(value=3)
        self.t0_var = tk.DoubleVar(value=25.0)
        self.rate_var = tk.DoubleVar(value=0.995)
        self.budget_var = tk.IntVar(value=4000)
        self.frame_var = tk.IntVar(value=10)
        self.schedule_var = tk.StringVar(value="Geometric decay (T x rate^step)")
        self.pick_var = tk.StringVar(value="Uniform random vertex")

        self.swatch_row = tk.Frame(ctrl_card, bg=BG_CARD)
        self.k_scale = self._slider_row(ctrl_card, "Palette size", self.k_var, 2, 6, 1, self.on_palette_change)
        self.swatch_row.pack(fill="x", pady=(0, 8))

        self.t0_scale = self._slider_row(ctrl_card, "Starting heat", self.t0_var, 1, 100, 0.5)

        tk.Label(ctrl_card, text="Cooling schedule", bg=BG_CARD, fg=INK_SOFT,
                 font=("Courier New", 9)).pack(anchor="w")
        schedule_box = ttk.Combobox(ctrl_card, textvariable=self.schedule_var, state="readonly",
                                     values=["Geometric decay (T x rate^step)",
                                             "Linear decay (T shrinks to 0 by end)",
                                             "Logarithmic decay (T0 / (1 + c*ln(step)))"])
        schedule_box.pack(fill="x", pady=(2, 10))

        self.rate_scale = self._slider_row(ctrl_card, "Decay constant", self.rate_var, 0.001, 0.999, 0.001)
        self.budget_scale = self._slider_row(ctrl_card, "Step budget", self.budget_var, 200, 20000, 100)

        tk.Label(ctrl_card, text="Vertex pick strategy", bg=BG_CARD, fg=INK_SOFT,
                 font=("Courier New", 9)).pack(anchor="w")
        pick_box = ttk.Combobox(ctrl_card, textvariable=self.pick_var, state="readonly",
                                 values=["Uniform random vertex", "Favor vertices currently clashing"])
        pick_box.pack(fill="x", pady=(2, 10))

        self._slider_row(ctrl_card, "Playback (moves/frame)", self.frame_var, 1, 200, 1)

        run_row = tk.Frame(ctrl_card, bg=BG_CARD)
        run_row.pack(fill="x", pady=(6, 0))
        self.go_button = tk.Button(run_row, text="Light the burner", command=self.on_go_click,
                                    bg=TEAL, fg="white", relief="solid", bd=1, font=("Courier New", 10, "bold"))
        self.go_button.pack(side="left", fill="x", expand=True, padx=(0, 6))
        tk.Button(run_row, text="Reshuffle colors", command=self.on_reshuffle_click,
                  bg=BG_PAPER, fg=INK, relief="solid", bd=1, font=("Courier New", 10, "bold")
                  ).pack(side="left", fill="x", expand=True)

        # ---- Bottom row: log + stats ----
        bottom = tk.Frame(self, bg=BG_PAPER)
        bottom.pack(fill="both", expand=True, padx=16, pady=(10, 16))
        bottom.columnconfigure(0, weight=3)
        bottom.columnconfigure(1, weight=2)

        log_card = self._card(bottom, "Lab log")
        log_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.log_canvas = tk.Canvas(log_card, width=650, height=150, bg=BG_CANVAS,
                                     highlightbackground=INK, highlightthickness=1)
        self.log_canvas.pack(fill="both", expand=True)
        legend = tk.Frame(log_card, bg=BG_CARD)
        legend.pack(fill="x", pady=(6, 0))
        tk.Label(legend, text="\u2500 live clashes", fg=DANGER, bg=BG_CARD, font=("Courier New", 9)).pack(side="left", padx=(0, 14))
        tk.Label(legend, text="\u2500 best found", fg=TEAL, bg=BG_CARD, font=("Courier New", 9)).pack(side="left")

        stats_card = self._card(bottom, "Thermometer & tally")
        stats_card.grid(row=0, column=1, sticky="nsew")

        therm_row = tk.Frame(stats_card, bg=BG_CARD)
        therm_row.pack(fill="x", pady=(0, 10))
        self.therm_canvas = tk.Canvas(therm_row, width=40, height=130, bg=BG_CANVAS,
                                       highlightbackground=INK, highlightthickness=1)
        self.therm_canvas.pack(side="left")
        therm_text = tk.Frame(therm_row, bg=BG_CARD)
        therm_text.pack(side="left", padx=(14, 0), anchor="n")
        tk.Label(therm_text, text="Heat", bg=BG_CARD, fg=INK_SOFT, font=("Courier New", 9)).pack(anchor="w")
        self.heat_label = tk.Label(therm_text, text="25.00", bg=BG_CARD, fg=INK, font=("Georgia", 14, "bold"))
        self.heat_label.pack(anchor="w", pady=(0, 8))
        tk.Label(therm_text, text="Step", bg=BG_CARD, fg=INK_SOFT, font=("Courier New", 9)).pack(anchor="w")
        self.step_label = tk.Label(therm_text, text="0 / 4000", bg=BG_CARD, fg=INK, font=("Georgia", 14, "bold"))
        self.step_label.pack(anchor="w")

        tiles = tk.Frame(stats_card, bg=BG_CARD)
        tiles.pack(fill="x", pady=(0, 10))
        cur_tile = tk.Frame(tiles, bg=BG_CANVAS, highlightbackground=INK, highlightthickness=1)
        cur_tile.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Label(cur_tile, text="LIVE CLASHES", bg=BG_CANVAS, fg=INK_SOFT, font=("Courier New", 8)).pack(anchor="w", padx=8, pady=(6, 0))
        self.cur_tile_val = tk.Label(cur_tile, text="\u2013", bg=BG_CANVAS, fg=DANGER, font=("Georgia", 16, "bold"))
        self.cur_tile_val.pack(anchor="w", padx=8, pady=(0, 6))
        best_tile = tk.Frame(tiles, bg=BG_CANVAS, highlightbackground=INK, highlightthickness=1)
        best_tile.pack(side="left", fill="x", expand=True, padx=(5, 0))
        tk.Label(best_tile, text="BEST FOUND", bg=BG_CANVAS, fg=INK_SOFT, font=("Courier New", 8)).pack(anchor="w", padx=8, pady=(6, 0))
        self.best_tile_val = tk.Label(best_tile, text="\u2013", bg=BG_CANVAS, fg=TEAL, font=("Georgia", 16, "bold"))
        self.best_tile_val.pack(anchor="w", padx=8, pady=(0, 6))

        table_frame = tk.Frame(stats_card, bg=BG_CANVAS, highlightbackground=INK, highlightthickness=1)
        table_frame.pack(fill="both", expand=True, pady=(0, 8))
        self.color_table = ttk.Treeview(table_frame, columns=("vertex", "color"), show="headings", height=6)
        self.color_table.heading("vertex", text="Vertex")
        self.color_table.heading("color", text="Color")
        self.color_table.column("vertex", width=90, anchor="w")
        self.color_table.column("color", width=140, anchor="w")
        self.color_table.pack(fill="both", expand=True)

        self.note_label = tk.Label(stats_card, text="Sketch or load a sample, then light the burner.",
                                    bg=BG_CARD, fg=INK_SOFT, font=("Courier New", 9), wraplength=380, justify="left")
        self.note_label.pack(fill="x")

        self.set_mode("sketch")
        self._paint_swatches(self.k_var.get())

    def _card(self, parent, title):
        card = tk.Frame(parent, bg=BG_CARD, highlightbackground=INK, highlightthickness=1, padx=14, pady=12)
        tk.Label(card, text=title.upper(), bg=BG_CARD, fg=INDIGO, font=("Georgia", 10, "bold")).pack(anchor="w", pady=(0, 8))
        return card

    def _slider_row(self, parent, label, var, lo, hi, step, command=None):
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", pady=(0, 4))
        top = tk.Frame(row, bg=BG_CARD)
        top.pack(fill="x")
        tk.Label(top, text=label, bg=BG_CARD, fg=INK_SOFT, font=("Courier New", 9)).pack(side="left")
        val_label = tk.Label(top, text=str(var.get()), bg=BG_CARD, fg=INDIGO, font=("Courier New", 9, "bold"))
        val_label.pack(side="right")

        def on_change(_evt=None, fire_command=True):
            v = var.get()
            val_label.config(text=(f"{v:.3f}" if step < 1 else f"{int(v)}"))
            if command and fire_command:
                command()

        scale = tk.Scale(row, from_=lo, to=hi, resolution=step, orient="horizontal", variable=var,
                          showvalue=False, command=lambda _v: on_change(), bg=BG_CARD, highlightthickness=0,
                          troughcolor=BG_CANVAS, fg=INK)
        scale.pack(fill="x")
        on_change(fire_command=False)
        return scale

    # ---------------------------------------------------------- mode / bench events
    MODE_CAPTIONS = {
        "sketch": "Click empty bench space to place a vertex.",
        "link": "Click one vertex, then a second, to connect or disconnect them.",
        "shift": "Drag a vertex to reposition it.",
        "erase": "Click a vertex to remove it, or click near an edge midpoint to cut that link.",
    }

    def set_mode(self, mode):
        self.mode = mode
        for key, btn in self.mode_buttons.items():
            btn.config(bg=(INDIGO if key == mode else BG_PAPER), fg=("white" if key == mode else INK))
        self.mode_caption.config(text=self.MODE_CAPTIONS[mode])
        self.pending_link = None
        self.draw_bench()

    def _bind_bench_events(self):
        c = self.bench_canvas
        c.bind("<Motion>", self._on_bench_motion)
        c.bind("<Button-1>", self._on_bench_click)
        c.bind("<B1-Motion>", self._on_bench_drag)
        c.bind("<ButtonRelease-1>", self._on_bench_release)

    def _on_bench_motion(self, event):
        if self.mode == "link" and self.pending_link is not None:
            self.draw_bench(preview=(event.x, event.y))

    def _on_bench_click(self, event):
        x, y = event.x, event.y
        hit = self.bench.vertex_near(x, y)

        if self.mode == "sketch":
            if not hit:
                vid = self.bench.add_vertex(x, y)
                if self.bench.hues:
                    self.bench.hues[vid] = random.randrange(self.k_var.get())
                self._refresh_all()

        elif self.mode == "link":
            if hit:
                if self.pending_link is None:
                    self.pending_link = hit["id"]
                elif self.pending_link != hit["id"]:
                    self.bench.toggle_edge(self.pending_link, hit["id"])
                    self.pending_link = None
                self._refresh_all()

        elif self.mode == "shift":
            if hit:
                self.dragged = hit

        elif self.mode == "erase":
            if hit:
                self.bench.remove_vertex(hit["id"])
            else:
                self.bench.remove_edge_near(x, y)
            self._refresh_all()

    def _on_bench_drag(self, event):
        if self.mode == "shift" and self.dragged:
            self.dragged["x"], self.dragged["y"] = event.x, event.y
            self.draw_bench()

    def _on_bench_release(self, _event):
        self.dragged = None

    def load_sample(self):
        self.bench = GraphBench()
        cx, cy, R = 320, 200, 150
        rim = [(cx, cy - R), (cx + R * 0.85, cy - R * 0.2), (cx + R * 0.55, cy + R * 0.75),
               (cx - R * 0.55, cy + R * 0.75), (cx - R * 0.85, cy - R * 0.2)]
        for x, y in rim:
            self.bench.add_vertex(x, y)
        self.bench.add_vertex(cx, cy)
        self.bench.edges = [(0, 1), (1, 2), (2, 3), (3, 4), (4, 0),
                             (5, 0), (5, 1), (5, 2), (5, 3), (5, 4)]
        self.bench.hues = self.bench.random_hues(self.k_var.get())
        self.best_clashes = self.bench.clash_count()
        self.best_hues = dict(self.bench.hues)
        self.history = [(0, self.best_clashes, self.best_clashes)]
        self.tick = 0
        self.note_label.config(text="Sample loaded — a hub-and-rim graph. Light the burner when ready.")
        self._refresh_all()

    def on_load_excel_click(self):
        """'Load Excel' button: read graph_input.xlsx, rebuild the graph and
        parameters from it, and refresh the GUI — same end state as
        load_sample(), just sourced from a file instead of being hardcoded."""
        try:
            vertex_ids, edge_pairs = load_excel_graph()
            parameters = load_parameters()
        except FileNotFoundError as exc:
            messagebox.showerror("File not found", str(exc))
            return
        except ValueError as exc:
            messagebox.showerror("Invalid Excel file", str(exc))
            return
        except Exception as exc:  # anything unexpected (corrupt file, etc.)
            messagebox.showerror("Could not load Excel file", str(exc))
            return

        # Stop any in-progress run before we swap the graph out from under it.
        if self.running:
            self.running = False
            if self.after_id:
                self.after_cancel(self.after_id)
            self.go_button.config(text="Light the burner")

        self.bench = GraphBench()
        build_graph(self.bench, vertex_ids, edge_pairs)

        # Load the SA parameters via their Scale widgets (rather than the
        # tk variables directly) so each one's on-change callback — label
        # text, swatch repaint, palette remap — fires exactly as it would
        # if the user had dragged the slider by hand.
        self.k_scale.set(parameters["Number_of_Colors"])
        self.t0_scale.set(parameters["Initial_Temperature"])
        self.rate_scale.set(parameters["Cooling_Rate"])
        self.budget_scale.set(parameters["Iterations"])

        self.bench.hues = self.bench.random_hues(self.k_var.get())
        self.best_clashes = self.bench.clash_count()
        self.best_hues = dict(self.bench.hues)
        self.history = [(0, self.best_clashes, self.best_clashes)]
        self.tick = 0
        self.heat = self.t0_var.get()

        self.note_label.config(
            text=f"Loaded {len(vertex_ids)} vertices / {len(self.bench.edges)} edges "
                 f"from {GRAPH_INPUT_FILE}. Parameters applied. Ready to run.")
        self._refresh_all()

    def wipe_bench(self):
        self.bench = GraphBench()
        self.best_clashes = float("inf")
        self.best_hues = None
        self.history = []
        self.tick = 0
        self.note_label.config(text="Bench wiped. Sketch a fresh graph.")
        self._refresh_all()

    # ---------------------------------------------------------- controls
    def on_palette_change(self):
        k = self.k_var.get()
        self._paint_swatches(k)

        # Small Issue 1 fix: if the palette shrank, remap any vertex whose
        # color index no longer exists to a valid random color instead of
        # silently leaving a stale index in place.
        remapped = False
        for v in self.bench.vertices:
            vid = v["id"]
            if vid in self.bench.hues and self.bench.hues[vid] >= k:
                self.bench.hues[vid] = random.randrange(k)
                remapped = True
        if remapped:
            self.best_clashes = self.bench.clash_count()
            self.best_hues = dict(self.bench.hues)
            self.history = []
            self.note_label.config(text="Palette shrank — out-of-range colors were reassigned.")
        self._refresh_all()

    def _paint_swatches(self, k):
        for child in self.swatch_row.winfo_children():
            child.destroy()
        for i in range(k):
            tk.Frame(self.swatch_row, bg=INK_PALETTE[i], width=16, height=16,
                     highlightbackground=INK, highlightthickness=1).pack(side="left", padx=3)

    def on_reshuffle_click(self):
        if not self.bench.vertices:
            return
        self.bench.hues = self.bench.random_hues(self.k_var.get())
        self.best_clashes = float("inf")
        self.best_hues = None
        self.history = []
        self.tick = 0
        self.running = False
        self.heat = self.t0_var.get()
        self.go_button.config(text="Light the burner")
        self.note_label.config(text="Colors reshuffled. Ready to run.")
        self._refresh_all()

    # ---------------------------------------------------------- annealing
    def _live_params(self):
        schedule_map = {
            "Geometric decay (T x rate^step)": "geometric",
            "Linear decay (T shrinks to 0 by end)": "linear",
            "Logarithmic decay (T0 / (1 + c*ln(step)))": "log",
        }
        return {
            "t0": self.t0_var.get(),
            "rate": self.rate_var.get(),
            "schedule": schedule_map[self.schedule_var.get()],
            "budget": self.budget_var.get(),
        }

    def anneal_step(self):
        n = len(self.bench.vertices)
        if n == 0:
            return 0
        k = self.k_var.get()

        if self.pick_var.get() == "Favor vertices currently clashing":
            in_conflict = [
                v for v in self.bench.vertices
                if any((a == v["id"] or b == v["id"]) and self.bench.hues.get(a) == self.bench.hues.get(b)
                       for a, b in self.bench.edges)
            ]
            if in_conflict and random.random() < 0.85:
                chosen = random.choice(in_conflict)
            else:
                chosen = random.choice(self.bench.vertices)
        else:
            chosen = random.choice(self.bench.vertices)

        old_hue = self.bench.hues.get(chosen["id"])
        new_hue = random.randrange(k)
        if k > 1:
            while new_hue == old_hue:
                new_hue = random.randrange(k)

        before = after = 0
        for a, b in self.bench.edges:
            if a == chosen["id"] or b == chosen["id"]:
                other = self.bench.hues.get(b if a == chosen["id"] else a)
                if other == old_hue:
                    before += 1
                if other == new_hue:
                    after += 1
        delta = after - before
        accept = delta <= 0
        if not accept:
            accept = random.random() < math.exp(-delta / max(self.heat, 1e-6))
        if accept:
            self.bench.hues[chosen["id"]] = new_hue

        self.tick += 1
        p = self._live_params()
        self.heat = heat_at(self.tick, p["t0"], p["rate"], p["schedule"], p["budget"])

        cur = self.bench.clash_count()
        if cur < self.best_clashes:
            self.best_clashes = cur
            self.best_hues = dict(self.bench.hues)
        if self.tick % max(1, p["budget"] // 300) == 0 or self.tick == p["budget"]:
            self.history.append((self.tick, cur, self.best_clashes))
        return cur

    def on_go_click(self):
        if not self.bench.vertices:
            self.note_label.config(text="Sketch at least one vertex first (or load the sample).")
            return
        if self.running:
            self.running = False
            self.go_button.config(text="Resume burner")
            if self.after_id:
                self.after_cancel(self.after_id)
            return

        if not self.bench.hues or len(self.bench.hues) != len(self.bench.vertices) or self.tick == 0:
            self.bench.hues = self.bench.random_hues(self.k_var.get())
            self.best_clashes = self.bench.clash_count()
            self.best_hues = dict(self.bench.hues)
            self.history = [(0, self.best_clashes, self.best_clashes)]
            self.tick = 0
            self.heat = self.t0_var.get()
            # rate / schedule / step budget are read live each step (see
            # _live_params) so they don't need to be captured here.

        self.running = True
        self.go_button.config(text="Pause")
        self.note_label.config(text="Annealing under way — watch the thermometer and the lab log.")
        self._loop_tick()

    def _loop_tick(self):
        if not self.running:
            return
        per_frame = self.frame_var.get()
        cur = self.bench.clash_count()
        budget_now = self._live_params()["budget"]
        for _ in range(per_frame):
            if self.tick >= budget_now:
                break
            cur = self.anneal_step()
            budget_now = self._live_params()["budget"]

        self._refresh_all(cur)

        if self.tick >= budget_now:
            self.running = False
            self.bench.hues = dict(self.best_hues or self.bench.hues)
            self.go_button.config(text="Light the burner")
            self._refresh_all()
            n_edges = len(self.bench.edges)
            plural = "" if self.best_clashes == 1 else "es"
            self.note_label.config(
                text=f"Burner out. Best coloring applied: {self.best_clashes} clash{plural} across {n_edges} links.")
            return

        self.after_id = self.after(16, self._loop_tick)

    # ---------------------------------------------------------- drawing
    def _refresh_all(self, cur_override=None):
        self.draw_log()
        self.draw_thermometer()
        cur = cur_override if cur_override is not None else self.bench.clash_count()
        self.cur_tile_val.config(text=str(cur) if self.bench.vertices else "\u2013")
        self.best_tile_val.config(text=(str(self.best_clashes) if self.best_clashes != float("inf") else "\u2013"))
        self.refresh_canvas()

    def draw_bench(self, preview=None):
        c = self.bench_canvas
        c.delete("all")
        for a, b in self.bench.edges:
            va = next((v for v in self.bench.vertices if v["id"] == a), None)
            vb = next((v for v in self.bench.vertices if v["id"] == b), None)
            if not va or not vb:
                continue
            clash = a in self.bench.hues and self.bench.hues.get(a) == self.bench.hues.get(b)
            c.create_line(va["x"], va["y"], vb["x"], vb["y"],
                          fill=(DANGER if clash else "#a89a89"), width=(3 if clash else 1.4))
        if self.mode == "link" and self.pending_link is not None and preview:
            vs = next((v for v in self.bench.vertices if v["id"] == self.pending_link), None)
            if vs:
                c.create_line(vs["x"], vs["y"], preview[0], preview[1], fill=INDIGO, dash=(4, 3))
        for v in self.bench.vertices:
            hue = self.bench.hues.get(v["id"])
            fill = INK_PALETTE[hue] if hue is not None else "#dcd3b8"
            outline = "#c98a2c" if self.pending_link == v["id"] else INK
            width = 3 if self.pending_link == v["id"] else 1.5
            c.create_oval(v["x"] - 15, v["y"] - 15, v["x"] + 15, v["y"] + 15, fill=fill, outline=outline, width=width)
            c.create_text(v["x"], v["y"], text=str(v.get("label", v["id"])),
                         fill=(BG_CANVAS if hue is not None else INK), font=("Courier New", 9, "bold"))

    def draw_log(self):
        c = self.log_canvas
        c.delete("all")
        if len(self.history) < 2:
            return
        max_v = max(1, max(p[1] for p in self.history), max(p[2] for p in self.history))
        w, h, pad = 650, 150, 8

        def xy(i, val):
            x = pad + (i / (len(self.history) - 1)) * (w - 2 * pad)
            y = h - pad - (val / max_v) * (h - 2 * pad)
            return x, y

        cur_pts = [xy(i, p[1]) for i, p in enumerate(self.history)]
        best_pts = [xy(i, p[2]) for i, p in enumerate(self.history)]
        c.create_line(*[coord for pt in cur_pts for coord in pt], fill=DANGER, width=2)
        c.create_line(*[coord for pt in best_pts for coord in pt], fill=TEAL, width=2)

    def draw_thermometer(self):
        c = self.therm_canvas
        c.delete("all")
        t0 = self.t0_var.get() or 1
        frac = max(0.0, min(1.0, self.heat / t0))
        hot = (166, 57, 46)
        cold = (47, 125, 118)
        mix = tuple(round(hot[i] * frac + cold[i] * (1 - frac)) for i in range(3))
        color = f"#{mix[0]:02x}{mix[1]:02x}{mix[2]:02x}"
        fill_h = frac * 128
        c.create_rectangle(2, 130 - fill_h, 38, 130, fill=color, outline="")
        self.heat_label.config(text=f"{self.heat:.2f}")
        self.step_label.config(text=f"{self.tick} / {self._live_params()['budget']}")

    def refresh_canvas(self):
        """Redraw the bench canvas and the vertex/color table together.
        Kept as its own method so it can be called on its own right after
        loading a graph from Excel."""
        self.draw_bench()
        self.refresh_color_table()

    def refresh_color_table(self):
        for row in self.color_table.get_children():
            self.color_table.delete(row)
        for v in sorted(self.bench.vertices, key=lambda v: v["id"]):
            hue = self.bench.hues.get(v["id"])
            color_name = COLOR_NAMES[hue] if hue is not None else "(unset)"
            self.color_table.insert("", "end", values=(f"Vertex {v.get('label', v['id'])}", color_name))


if __name__ == "__main__":
    app = App()
    app.mainloop()
