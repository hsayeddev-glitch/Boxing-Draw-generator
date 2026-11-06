"""
Boxing Draw Generator – Maharashtra Boxing Association
======================================================
FINAL VERSION (v12) — Complete Feature Set

CORE FEATURES:
  • Smart bracket algorithm (R32 → R16 → QF → SF → F)
    based on player count (2–32 players supported)
  • Correct box counts per column (32/16/8/4/2 boxes)
  • BYE spread algorithm — no BYE-vs-BYE bouts
  • Nested folder system (unlimited depth, max 26 draws/folder)

DRAW MANAGEMENT:
  • Create / Edit / Delete / Rename draws
  • Add Entries — reopen create screen pre-filled
  • Tournament name editable on the View screen (used in PDF header)
  • Category name inline-editable on the View screen

BRACKET INTERACTION:
  • Per-player left-click drag-to-swap (Round 1 only)
  • Single left-click → select box for inline quick-edit
  • Quick-Edit panel: change Name + Club without any dialog
  • Right-click context menu:
      – Rename Boxer (name + club via dialog)
      – Swap Corners (RED ↔ BLUE)
      – Rename Bout Label
  • Zoom: [−] / [+] / [Fit] / [100%] buttons + Ctrl+Scroll wheel
    Range: 40 % – 250 % in 10 % steps

IMPORT:
  • Import from Excel (.xlsx / .xls) or CSV
  • Auto-detects header row (Name / Club / Gym / Team columns)
  • Preview dialog before loading into the entry grid

EXPORT:
  • Export as Text (.txt) — full bracket with round headers
  • Export as PDF (.pdf) — A4 portrait via reportlab
      – Dark-themed header with tournament name + category
      – All 5 columns, ghost columns for unused rounds
      – BYE shown with red cross, TBD shows "Winner of Bout X"
      – Per-box direct connector lines

DEPENDENCIES:
  Required:  tkinter  (stdlib)
  Optional:  openpyxl   → pip install openpyxl   (Excel import)
             reportlab  → pip install reportlab   (PDF export)
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import math
import random
import datetime
import csv

# ── optional Excel (openpyxl) support ────────────────────────
try:
    import openpyxl
    XLSX_OK = True
except ImportError:
    XLSX_OK = False
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────

MAX_DRAWS = 26

# Column index 0 = leftmost displayed column (ROUND OF 32).
# We always render exactly 5 display columns.
_ROUND_LABELS = [
    "ROUND OF 32",
    "ROUND OF 16",
    "QUARTER-FINALS",
    "SEMI-FINALS",
    "FINAL",
]
NUM_COLS = 5   # always draw 5 columns

# bracket_size → which display column the real Round-0 occupies
_BS_TO_COL = {32: 0, 16: 1, 8: 2, 4: 3, 2: 4}


# ─────────────────────────────────────────────────────────────
# DATA STRUCTURES
# ─────────────────────────────────────────────────────────────

class Player:
    def __init__(self, name: str, club: str):
        self.name = name.strip()
        self.club = club.strip()

    def __repr__(self):
        return f"{self.name} ({self.club})"


class Match:
    """player1 = RED corner, player2 = BLUE corner."""
    def __init__(self, bout_number: int, player1, player2):
        self.bout_number = bout_number
        self.bout_label  = f"Bout {bout_number}"
        self.player1     = player1   # Player | "BYE" | None
        self.player2     = player2
        self.winner      = None

        if player1 == "BYE":
            self.winner = player2
        elif player2 == "BYE":
            self.winner = player1

    def is_bye(self):
        return self.player1 == "BYE" or self.player2 == "BYE"

    def swap_corners(self):
        self.player1, self.player2 = self.player2, self.player1
        self.winner = None
        if self.player1 == "BYE":  self.winner = self.player2
        elif self.player2 == "BYE": self.winner = self.player1


class Draw:
    def __init__(self, category: str, players: list):
        self.category = category
        self.players  = list(players)
        self.rounds: list[list[Match]] = []
        self.bracket_size = 0
        self.bye_count    = 0
        self.total_bouts  = 0
        self._generate()

    @staticmethod
    def _nearest_pow2(n):
        if n <= 1: return 1
        return 2 ** math.ceil(math.log2(n))

    def _generate(self):
        n = len(self.players)
        self.bracket_size = self._nearest_pow2(max(n, 2))
        self.bye_count    = self.bracket_size - n
        bc = [1]

        # ── Spread BYEs so no bout has two BYEs ─────────────────
        # Algorithm:
        #   1. Shuffle the real players randomly.
        #   2. Place the seeded list so BYEs are distributed at
        #      even intervals across all slots — guaranteeing that
        #      adjacent pairs never both contain a BYE.
        #   3. If bye_count > bracket_size/2 (more BYEs than bouts)
        #      it is mathematically impossible to avoid BYE-vs-BYE;
        #      in that case we still spread as evenly as possible.
        players_shuffled = self.players[:]
        random.shuffle(players_shuffled)

        if self.bye_count == 0:
            seeded = players_shuffled
        else:
            # Place BYEs at evenly-spaced indices in the full bracket
            total   = self.bracket_size
            seeded  = [None] * total
            # Compute ideal positions for BYEs: spread evenly
            step    = total / self.bye_count
            bye_positions = set()
            for k in range(self.bye_count):
                pos = int(round(k * step + step / 2)) % total
                # Avoid placing two BYEs adjacent (same pair = indices 2i, 2i+1)
                while pos in bye_positions or (
                        pos % 2 == 1 and (pos - 1) in bye_positions) or (
                        pos % 2 == 0 and (pos + 1) in bye_positions):
                    pos = (pos + 1) % total
                bye_positions.add(pos)

            # Fill BYE positions; fill remaining with shuffled players
            player_iter = iter(players_shuffled)
            for i in range(total):
                if i in bye_positions:
                    seeded[i] = "BYE"
                else:
                    seeded[i] = next(player_iter)

        def make_round(parts):
            rnd = []
            for i in range(0, len(parts), 2):
                p1 = parts[i]
                p2 = parts[i+1] if i+1 < len(parts) else "BYE"
                rnd.append(Match(bc[0], p1, p2))
                bc[0] += 1
            return rnd

        cur = make_round(seeded)
        self.rounds.append(cur)
        while len(cur) > 1:
            cur = make_round([m.winner for m in cur])
            self.rounds.append(cur)
        self.total_bouts = bc[0] - 1

    def rebuild(self, new_players: list):
        self.players = list(new_players)
        self.rounds = []
        self.bracket_size = self.bye_count = self.total_bouts = 0
        self._generate()

    # ── per-player swap helper ────────────────────────────────
    def swap_players(self, r_idx: int,
                     m_a: int, slot_a: int,
                     m_b: int, slot_b: int):
        """
        Swap two individual player slots in round r_idx.
        slot: 0 = player1 (red), 1 = player2 (blue).
        After the swap, recompute BYE winners for affected matches.
        """
        rnd = self.rounds[r_idx]
        if m_a >= len(rnd) or m_b >= len(rnd):
            return

        def get_p(match, slot):
            return match.player1 if slot == 0 else match.player2

        def set_p(match, slot, val):
            if slot == 0: match.player1 = val
            else:         match.player2 = val

        pa = get_p(rnd[m_a], slot_a)
        pb = get_p(rnd[m_b], slot_b)
        set_p(rnd[m_a], slot_a, pb)
        set_p(rnd[m_b], slot_b, pa)

        # Recompute winners for touched matches
        for mi in {m_a, m_b}:
            m = rnd[mi]
            m.winner = None
            if m.player1 == "BYE":  m.winner = m.player2
            elif m.player2 == "BYE": m.winner = m.player1


class Folder:
    def __init__(self, name: str, parent=None):
        self.name     = name.strip()
        self.parent   = parent
        self.children: list["Folder"] = []
        self.draws:    list[Draw]     = []

    def full(self): return len(self.draws) >= MAX_DRAWS
    def remaining(self): return MAX_DRAWS - len(self.draws)
    def path(self) -> str:
        return (self.parent.path() + " / " + self.name
                if self.parent else self.name)


# ─────────────────────────────────────────────────────────────
# APPLICATION STATE
# ─────────────────────────────────────────────────────────────

root_folder          = Folder("Root")
active_folder: Folder | None = None


def _active_draws() -> list[Draw]:
    return active_folder.draws if active_folder else []


# ─────────────────────────────────────────────────────────────
# ROOT APPLICATION
# ─────────────────────────────────────────────────────────────

class BoxingApp(tk.Tk):
    BG         = "#1a1a2e"
    PANEL      = "#16213e"
    ACCENT     = "#e94560"
    GOLD       = "#f5a623"
    TEXT       = "#eaeaea"
    SUBTEXT    = "#a0a8b8"
    ENTRY_BG   = "#0f3460"
    ENTRY_FG   = "#ffffff"
    BTN_BG     = "#e94560"
    BTN_FG     = "#ffffff"
    BTN_HVR    = "#c73652"
    BRACKET_BG = "#ffffff"
    RED_CORNER = "#c0392b"
    BLU_CORNER = "#1a5276"

    def __init__(self):
        super().__init__()
        self.title("Boxing Draw Generator – Maharashtra Boxing Association")
        self.geometry("1300x880")
        self.resizable(True, True)
        self.configure(bg=self.BG)
        self._setup_styles()

        c = tk.Frame(self, bg=self.BG)
        c.pack(fill="both", expand=True)
        c.grid_rowconfigure(0, weight=1)
        c.grid_columnconfigure(0, weight=1)

        self.frames: dict[type, tk.Frame] = {}
        for Cls in (WelcomeScreen, FolderManagerScreen,
                    CreateDrawScreen, ViewDrawsScreen):
            f = Cls(c, self)
            self.frames[Cls] = f
            f.grid(row=0, column=0, sticky="nsew")

        self.show_frame(WelcomeScreen)

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",   background=self.BG)
        s.configure("TLabel",   background=self.BG, foreground=self.TEXT)
        s.configure("Treeview", background=self.PANEL, foreground=self.TEXT,
                    fieldbackground=self.PANEL, rowheight=26)
        s.configure("Treeview.Heading", background=self.ACCENT,
                    foreground=self.BTN_FG, font=("Georgia", 10, "bold"))
        s.map("Treeview", background=[("selected", self.ACCENT)])

    def show_frame(self, cls: type, **kwargs):
        f = self.frames[cls]
        if hasattr(f, "on_show"):
            f.on_show(**kwargs)
        f.tkraise()

    def make_btn(self, parent, text, command,
                 width=22, color=None, font_size=11):
        bg = color or self.BTN_BG
        b  = tk.Button(parent, text=text, command=command,
                       bg=bg, fg=self.BTN_FG, activebackground=self.BTN_HVR,
                       font=("Georgia", font_size, "bold"), relief="flat",
                       padx=10, pady=7, cursor="hand2", width=width)
        b.bind("<Enter>", lambda e: b.config(bg=self.BTN_HVR))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def top_bar(self, parent, title, back_target=None, back_label="← Back"):
        bar = tk.Frame(parent, bg=self.PANEL, pady=12)
        bar.pack(fill="x")
        if back_target:
            self.make_btn(bar, back_label,
                          lambda: self.show_frame(back_target),
                          width=10, color="#444").pack(side="left", padx=16)
        tk.Label(bar, text=title, font=("Georgia", 16, "bold"),
                 bg=self.PANEL, fg=self.GOLD).pack(side="left", padx=8)
        return bar


# ─────────────────────────────────────────────────────────────
# SCREEN 1 – WELCOME
# ─────────────────────────────────────────────────────────────

class WelcomeScreen(tk.Frame):
    def __init__(self, parent, app: BoxingApp):
        super().__init__(parent, bg=app.BG)
        self.app = app
        self._build()

    def _build(self):
        app = self.app
        tk.Frame(self, bg=app.ACCENT, height=8).pack(fill="x")
        logo = tk.Frame(self, bg=app.BG, pady=34)
        logo.pack(fill="x")
        tk.Label(logo, text="🥊", font=("Arial", 52),
                 bg=app.BG, fg=app.GOLD).pack()
        tk.Label(logo, text="Boxing Draw Generator",
                 font=("Georgia", 28, "bold"),
                 bg=app.BG, fg=app.GOLD).pack(pady=(6, 2))
        tk.Label(logo, text="Maharashtra Boxing Association",
                 font=("Georgia", 14), bg=app.BG, fg=app.SUBTEXT).pack()
        tk.Frame(self, bg=app.ACCENT, height=2).pack(fill="x", padx=80, pady=18)
        menu = tk.Frame(self, bg=app.BG)
        menu.pack(expand=True)
        for txt, cmd in [
            ("📁  Folder Manager",      lambda: app.show_frame(FolderManagerScreen)),
            ("➕  Create New Draw",     self._go_create),
            ("📋  View Existing Draws", self._go_view),
        ]:
            app.make_btn(menu, txt, cmd, width=30).pack(pady=9, ipadx=10, ipady=4)
        tk.Label(self,
                 text=f"Each folder holds up to {MAX_DRAWS} draws  •  © Maharashtra Boxing Association",
                 font=("Georgia", 9), bg=app.BG, fg=app.SUBTEXT
                 ).pack(side="bottom", pady=10)

    def _need_folder(self):
        if not active_folder:
            messagebox.showinfo("No Folder",
                "Open Folder Manager and select a folder first.")
            self.app.show_frame(FolderManagerScreen)
            return False
        return True

    def _go_create(self):
        if self._need_folder(): self.app.show_frame(CreateDrawScreen)

    def _go_view(self):
        if self._need_folder(): self.app.show_frame(ViewDrawsScreen)


# ─────────────────────────────────────────────────────────────
# SCREEN 2 – FOLDER MANAGER
# ─────────────────────────────────────────────────────────────

class FolderManagerScreen(tk.Frame):
    def __init__(self, parent, app: BoxingApp):
        super().__init__(parent, bg=app.BG)
        self.app = app
        self._cur: Folder = root_folder
        self._build()

    def _build(self):
        app = self.app
        bar = app.top_bar(self, "Folder Manager", WelcomeScreen)
        self.active_lbl = tk.Label(bar, text="", font=("Georgia", 9),
                                   bg=app.PANEL, fg=app.GOLD)
        self.active_lbl.pack(side="right", padx=16)

        self.crumb_frame = tk.Frame(self, bg=app.ENTRY_BG, pady=6)
        self.crumb_frame.pack(fill="x")

        top = tk.Frame(self, bg=app.BG)
        top.pack(fill="x", padx=24, pady=(12, 4))
        tk.Label(top, text="New Sub-folder:", font=("Georgia", 11, "bold"),
                 bg=app.BG, fg=app.GOLD).pack(side="left")
        self.new_name = tk.StringVar()
        tk.Entry(top, textvariable=self.new_name, font=("Georgia", 11),
                 bg=app.ENTRY_BG, fg=app.ENTRY_FG, insertbackground="white",
                 relief="flat", width=28).pack(side="left", padx=8, ipady=5)
        app.make_btn(top, "➕ Create", self._create_folder, width=12).pack(side="left")

        tk.Label(self, text="Contents", font=("Georgia", 10, "bold"),
                 bg=app.BG, fg=app.SUBTEXT).pack(anchor="w", padx=24, pady=(6, 2))

        lf = tk.Frame(self, bg=app.ENTRY_BG, padx=2, pady=2)
        lf.pack(fill="both", expand=True, padx=24)

        self.tree = ttk.Treeview(lf, show="tree headings",
                                 columns=("type", "info"), selectmode="browse")
        self.tree.heading("#0",   text="Name")
        self.tree.heading("type", text="Type")
        self.tree.heading("info", text="Info")
        self.tree.column("#0",   width=300)
        self.tree.column("type", width=120, anchor="center")
        self.tree.column("info", width=220, anchor="center")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-Button-1>", self._on_dbl)

        btn_row = tk.Frame(self, bg=app.BG)
        btn_row.pack(fill="x", padx=24, pady=10)
        for txt, cmd, col, w in [
            ("📂 Open Folder", self._open_folder, None,    16),
            ("✏️  Rename",     self._rename_item, "#445",  12),
            ("🗑  Delete",      self._delete_item, "#661",  12),
            ("⬆  Up",          self._go_up,       "#333",   8),
        ]:
            app.make_btn(btn_row, txt, cmd, width=w, color=col
                         ).pack(side="left", padx=(0, 6))

        self.status_lbl = tk.Label(self, text="", font=("Georgia", 9),
                                   bg=app.BG, fg=app.SUBTEXT)
        self.status_lbl.pack(anchor="w", padx=24, pady=(0, 6))

    def on_show(self, **_): self._refresh()

    def _refresh(self):
        for w in self.crumb_frame.winfo_children(): w.destroy()
        tk.Label(self.crumb_frame, text=f"  📂 {self._cur.path()}",
                 font=("Georgia", 10, "bold"),
                 bg=self.app.ENTRY_BG, fg=self.app.GOLD).pack(side="left", padx=6)

        self.tree.delete(*self.tree.get_children())
        for child in self._cur.children:
            mark = " ◀ ACTIVE" if child is active_folder else ""
            self.tree.insert("", "end", iid=f"F:{id(child)}",
                             text=f"📁  {child.name}{mark}",
                             values=("Folder",
                                     f"{len(child.children)} sub-folders, "
                                     f"{len(child.draws)}/{MAX_DRAWS} draws"))
        for i, d in enumerate(self._cur.draws):
            self.tree.insert("", "end", iid=f"D:{i}",
                             text=f"🥊  {d.category}",
                             values=("Draw", f"{len(d.players)} entries"))

        self.active_lbl.config(
            text=f"Active: {active_folder.path()}" if active_folder
            else "No folder selected")
        self.status_lbl.config(
            text="Root level" if self._cur is root_folder
            else f"Inside: {self._cur.path()}")

    def _sel_folder(self):
        sel = self.tree.selection()
        if not sel: return None
        iid = sel[0]
        if iid.startswith("F:"):
            oid = int(iid[2:])
            for c in self._cur.children:
                if id(c) == oid: return c
        return None

    def _on_dbl(self, _=None):
        f = self._sel_folder()
        if f: self._cur = f; self._refresh()

    def _go_up(self):
        if self._cur.parent: self._cur = self._cur.parent; self._refresh()

    def _create_folder(self):
        name = self.new_name.get().strip()
        if not name: messagebox.showerror("Error", "Enter a name."); return
        if any(c.name == name for c in self._cur.children):
            messagebox.showerror("Error", f"'{name}' already exists."); return
        self._cur.children.append(Folder(name, parent=self._cur))
        self.new_name.set(""); self._refresh()

    def _open_folder(self):
        global active_folder
        f = self._sel_folder() or self._cur
        active_folder = f; self._refresh()
        messagebox.showinfo("Folder Opened",
            f"Working in:\n{f.path()}\n"
            f"Draws: {len(f.draws)}/{MAX_DRAWS}")

    def _rename_item(self):
        f = self._sel_folder()
        if not f: messagebox.showwarning("No Folder", "Select a folder."); return
        new = simpledialog.askstring("Rename", f"New name for '{f.name}':", parent=self)
        if new and new.strip():
            new = new.strip()
            if any(c.name == new for c in self._cur.children if c is not f):
                messagebox.showerror("Error", f"'{new}' exists."); return
            f.name = new; self._refresh()

    def _delete_item(self):
        global active_folder
        f = self._sel_folder()
        if not f: messagebox.showwarning("No Folder", "Select a folder."); return
        total = self._count(f)
        if not messagebox.askyesno("Delete",
                f"Delete '{f.name}' ({total} draws)?\nCannot be undone."): return
        if active_folder and self._is_desc(active_folder, f): active_folder = None
        self._cur.children.remove(f); self._refresh()

    @staticmethod
    def _count(folder):
        return len(folder.draws) + sum(FolderManagerScreen._count(c) for c in folder.children)

    @staticmethod
    def _is_desc(cand, anc):
        if cand is anc: return True
        return any(FolderManagerScreen._is_desc(cand, c) for c in anc.children)


# ─────────────────────────────────────────────────────────────
# EXCEL / CSV IMPORT HELPERS
# ─────────────────────────────────────────────────────────────

def _parse_import_file(path: str) -> list[tuple[str, str]]:
    """
    Read .xlsx or .csv and return list of (name, club) tuples (max 32).

    Header detection (case-insensitive):
      - Name column  : header contains "name"
      - Club column  : header contains "club", "gym", or "team"
      - No headers   : column 0 = name, column 1 = club
    Blank name rows are skipped.
    """
    ext = path.lower().rsplit(".", 1)[-1]
    rows: list[list[str]] = []

    if ext == "csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                rows.append([str(c).strip() for c in row])
    elif ext in ("xlsx", "xlsm", "xltx", "xls"):
        if not XLSX_OK:
            raise ImportError(
                "openpyxl is not installed.\n"
                "Run:  pip install openpyxl\nthen restart.")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append([str(c).strip() if c is not None else "" for c in row])
        wb.close()
    else:
        raise ValueError(f"Unsupported file type '.{ext}'. Use .xlsx or .csv")

    if not rows:
        return []

    header = [h.lower() for h in rows[0]]
    name_col = next((i for i, h in enumerate(header) if "name" in h), None)
    club_col  = next((i for i, h in enumerate(header)
                      if any(k in h for k in ("club", "gym", "team"))), None)
    has_header = name_col is not None or club_col is not None
    data_rows  = rows[1:] if has_header else rows
    if name_col is None: name_col = 0
    if club_col is None: club_col = 1 if len(rows[0]) > 1 else 0

    result = []
    for row in data_rows:
        nm = row[name_col] if name_col < len(row) else ""
        cl = row[club_col] if club_col < len(row) else ""
        if nm and nm.lower() not in ("none", "nan", "null", ""):
            result.append((nm, cl))
        if len(result) >= 32:
            break
    return result


class _ImportPreviewDialog(tk.Toplevel):
    """Modal preview of parsed Excel/CSV rows before they go into the grid."""
    def __init__(self, parent, entries: list[tuple[str, str]], app: "BoxingApp"):
        super().__init__(parent)
        self.app       = app
        self.confirmed = False
        self.entries   = entries
        self.title("Excel / CSV Import Preview")
        self.configure(bg=app.BG)
        self.resizable(False, False)
        self.grab_set()
        self._build()
        self.wait_window()

    def _build(self):
        app = self.app
        tk.Label(self, text="📥  Import Preview",
                 font=("Georgia", 13, "bold"),
                 bg=app.BG, fg=app.GOLD).pack(pady=(14, 2), padx=20)
        tk.Label(self,
                 text=f"{len(self.entries)} entries found — confirm to load into the entry grid.",
                 font=("Georgia", 9), bg=app.BG, fg=app.SUBTEXT).pack(padx=20, pady=(0, 8))

        frame = tk.Frame(self, bg=app.ENTRY_BG)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        hdr = tk.Frame(frame, bg=app.ENTRY_BG)
        hdr.pack(fill="x")
        for col_txt, w in [("#", 4), ("Player Name", 24), ("Club / Gym", 22)]:
            tk.Label(hdr, text=col_txt, font=("Georgia", 9, "bold"),
                     bg=app.ACCENT, fg="white", width=w,
                     anchor="w", padx=4, pady=3).pack(side="left", padx=1)

        inner_h = min(len(self.entries) * 26 + 8, 300)
        cvs = tk.Canvas(frame, bg=app.ENTRY_BG, height=inner_h,
                        highlightthickness=0)
        sb  = ttk.Scrollbar(frame, orient="vertical", command=cvs.yview)
        cvs.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cvs.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(cvs, bg=app.ENTRY_BG)
        cvs.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: cvs.configure(scrollregion=cvs.bbox("all")))

        for i, (nm, cl) in enumerate(self.entries):
            bg_ = app.ENTRY_BG if i % 2 == 0 else "#0d2a50"
            row = tk.Frame(inner, bg=bg_)
            row.pack(fill="x")
            tk.Label(row, text=f"{i+1:>2}.", font=("Georgia", 9),
                     bg=bg_, fg=app.SUBTEXT, width=4).pack(side="left")
            tk.Label(row, text=nm, font=("Georgia", 9, "bold"),
                     bg=bg_, fg=app.TEXT, width=24, anchor="w").pack(side="left", padx=2)
            tk.Label(row, text=cl, font=("Georgia", 9),
                     bg=bg_, fg=app.SUBTEXT, width=22, anchor="w").pack(side="left", padx=2)

        btn_row = tk.Frame(self, bg=app.BG)
        btn_row.pack(pady=(0, 12))
        app.make_btn(btn_row, "✅  Confirm Import",
                     self._confirm, width=20).pack(side="left", padx=8)
        app.make_btn(btn_row, "❌  Cancel",
                     self.destroy, width=12, color="#555").pack(side="left", padx=8)

    def _confirm(self):
        self.confirmed = True
        self.destroy()


# ─────────────────────────────────────────────────────────────
# SCREEN 3 – CREATE / EDIT DRAW
# ─────────────────────────────────────────────────────────────

class CreateDrawScreen(tk.Frame):
    def __init__(self, parent, app: BoxingApp):
        super().__init__(parent, bg=app.BG)
        self.app = app
        self.player_entries: list[tuple[tk.StringVar, tk.StringVar]] = []
        self._editing_draw: Draw | None = None
        self._build()

    def _build(self):
        app = self.app
        bar = app.top_bar(self, "Create / Edit Draw", WelcomeScreen)
        self.folder_lbl = tk.Label(bar, text="", font=("Georgia", 10),
                                   bg=app.PANEL, fg=app.GOLD)
        self.folder_lbl.pack(side="right", padx=16)

        body = tk.Frame(self, bg=app.BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        # Left panel
        left = tk.Frame(body, bg=app.PANEL, padx=16, pady=14, width=318)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        tk.Label(left, text="▸ DRAW SETTINGS", font=("Georgia", 10, "bold"),
                 bg=app.PANEL, fg=app.ACCENT).pack(anchor="w", pady=(0, 6))
        tk.Label(left, text="Weight Category", font=("Georgia", 10, "bold"),
                 bg=app.PANEL, fg=app.GOLD).pack(anchor="w")
        self.cat_var = tk.StringVar()
        tk.Entry(left, textvariable=self.cat_var, font=("Georgia", 10),
                 bg=app.ENTRY_BG, fg=app.ENTRY_FG, insertbackground="white",
                 relief="flat", width=28).pack(pady=(3, 10), ipady=5)

        tk.Label(left, text="Number of Players (2–32)",
                 font=("Georgia", 10, "bold"),
                 bg=app.PANEL, fg=app.GOLD).pack(anchor="w")
        crow = tk.Frame(left, bg=app.PANEL)
        crow.pack(anchor="w", pady=(3, 4))
        self.count_var = tk.StringVar(value="4")
        tk.Entry(crow, textvariable=self.count_var, font=("Georgia", 10),
                 bg=app.ENTRY_BG, fg=app.ENTRY_FG, insertbackground="white",
                 relief="flat", width=5).pack(side="left", ipady=5)
        app.make_btn(crow, "Set ▶", self._set_count,
                     width=8, font_size=10).pack(side="left", padx=(8, 0))

        self.info_var = tk.StringVar()
        tk.Label(left, textvariable=self.info_var, font=("Georgia", 8),
                 bg=app.PANEL, fg=app.SUBTEXT, wraplength=275
                 ).pack(pady=(2, 4), anchor="w")

        app.make_btn(left, "⚡  Generate Draw", self._generate, width=24
                     ).pack(pady=5, fill="x")
        app.make_btn(left, "🔄  Reset", self._reset, width=24, color="#555"
                     ).pack(pady=3, fill="x")
        app.make_btn(left, "📥  Import from Excel / CSV",
                     self._import_excel, width=24, color="#1a3d3d"
                     ).pack(pady=3, fill="x")

        tk.Frame(left, bg=app.ACCENT, height=1).pack(fill="x", pady=10)

        tk.Label(left, text="▸ EXISTING DRAWS", font=("Georgia", 10, "bold"),
                 bg=app.PANEL, fg=app.ACCENT).pack(anchor="w", pady=(0, 4))
        self.draws_lb = tk.Listbox(
            left, font=("Georgia", 9), bg=app.ENTRY_BG, fg=app.TEXT,
            selectbackground=app.ACCENT, activestyle="none",
            relief="flat", borderwidth=0, highlightthickness=0, height=9)
        self.draws_lb.pack(fill="x", pady=(0, 6))
        act = tk.Frame(left, bg=app.PANEL)
        act.pack(fill="x")
        app.make_btn(act, "✏️ Rename", self._rename_draw,
                     width=11, color="#334", font_size=9).pack(side="left", padx=(0, 4))
        app.make_btn(act, "🗑 Delete", self._delete_draw,
                     width=11, color="#661", font_size=9).pack(side="left")

        # Right panel
        right = tk.Frame(body, bg=app.BG)
        right.pack(side="left", fill="both", expand=True)
        self.rtitle = tk.Label(right, text="Player Details",
                               font=("Georgia", 13, "bold"),
                               bg=app.BG, fg=app.GOLD)
        self.rtitle.pack(anchor="w", pady=(0, 6))
        hdr = tk.Frame(right, bg=app.BG)
        hdr.pack(fill="x")
        for col, w in [("#", 4), ("Player Name", 22), ("Club / Gym", 20)]:
            tk.Label(hdr, text=col, font=("Georgia", 10, "bold"),
                     bg=app.ACCENT, fg="white", width=w, anchor="w",
                     padx=6, pady=4).pack(side="left", padx=1)

        cf = tk.Frame(right, bg=app.BG)
        cf.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(cf, bg=app.BG, highlightthickness=0)
        sb2 = ttk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.pf = tk.Frame(self.canvas, bg=app.BG)
        self._cw = self.canvas.create_window((0, 0), window=self.pf, anchor="nw")
        self.pf.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
            lambda e: self.canvas.itemconfig(self._cw, width=e.width))

    def on_show(self, draw=None, **_):
        self._editing_draw = draw
        self.folder_lbl.config(
            text=f"Folder: {active_folder.path()}" if active_folder else "—")
        self._update_info()
        self._refresh_list()
        if draw is not None:
            self.cat_var.set(draw.category)
            self.rtitle.config(text=f"Add Entries to: {draw.category}")
            self.count_var.set(str(len(draw.players) + 4))
            self._set_count(prefill=draw.players)
        else:
            self.rtitle.config(text="Player Details")

    def _update_info(self):
        if active_folder:
            self.info_var.set(
                f"'{active_folder.name}'  •  "
                f"{len(active_folder.draws)}/{MAX_DRAWS}  •  "
                f"{active_folder.remaining()} remaining")
        else:
            self.info_var.set("No folder selected.")

    def _refresh_list(self):
        self.draws_lb.delete(0, "end")
        for i, d in enumerate(_active_draws()):
            self.draws_lb.insert("end", f"  {i+1}. {d.category}  [{len(d.players)}P]")

    def _set_count(self, prefill=None):
        try:
            n = int(self.count_var.get())
            if not (2 <= n <= 32): raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Enter 2–32."); return
        for w in self.pf.winfo_children(): w.destroy()
        self.player_entries.clear()
        app = self.app
        for i in range(n):
            row = tk.Frame(self.pf, bg=app.BG)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=f"{i+1:>2}.", font=("Georgia", 10),
                     bg=app.BG, fg=app.SUBTEXT, width=4).pack(side="left")
            nv, cv = tk.StringVar(), tk.StringVar()
            if prefill and i < len(prefill):
                nv.set(prefill[i].name); cv.set(prefill[i].club)
            bg_ = "#0a3d0a" if (prefill and i < len(prefill)) else app.ENTRY_BG

            name_e = tk.Entry(row, textvariable=nv, font=("Georgia", 10),
                     bg=bg_, fg=app.ENTRY_FG, insertbackground="white",
                     relief="flat", width=22)
            name_e.pack(side="left", padx=2, ipady=4)

            club_e = tk.Entry(row, textvariable=cv, font=("Georgia", 10),
                     bg=bg_, fg=app.ENTRY_FG, insertbackground="white",
                     relief="flat", width=20)
            club_e.pack(side="left", padx=2, ipady=4)

            # ── Autoscroll: when this row gets focus, scroll the
            #    canvas so it is fully visible.  We capture i in the
            #    closure via a default argument.
            def _scroll_to_row(event, row_widget=row):
                """Scroll the player canvas so the focused row is visible."""
                self.canvas.update_idletasks()
                # y position of the row inside the scrollable frame
                try:
                    row_y = row_widget.winfo_y()
                    row_h = row_widget.winfo_height()
                    total = self.pf.winfo_height()
                    if total <= 0:
                        return
                    # fraction of the canvas the row occupies
                    frac_top = row_y / total
                    frac_bot = (row_y + row_h) / total
                    # current visible region
                    top_v, bot_v = self.canvas.yview()
                    vis_h = bot_v - top_v
                    if frac_top < top_v:
                        self.canvas.yview_moveto(max(0.0, frac_top - 0.02))
                    elif frac_bot > bot_v:
                        self.canvas.yview_moveto(
                            min(1.0 - vis_h, frac_bot - vis_h + 0.02))
                except Exception:
                    pass

            name_e.bind("<FocusIn>", _scroll_to_row)
            club_e.bind("<FocusIn>", _scroll_to_row)

            self.player_entries.append((nv, cv))

    def _generate(self):
        if not active_folder:
            messagebox.showerror("Error", "No folder selected."); return
        cat = self.cat_var.get().strip()
        if not cat: messagebox.showerror("Error", "Enter a weight category."); return
        if not self.player_entries:
            messagebox.showerror("Error", "Set player count first."); return
        players = []
        for i, (nv, cv) in enumerate(self.player_entries):
            n, c = nv.get().strip(), cv.get().strip()
            if not n:
                messagebox.showerror("Error", f"Player {i+1} name is empty."); return
            players.append(Player(n, c or "—"))

        if self._editing_draw is not None:
            self._editing_draw.category = cat
            self._editing_draw.rebuild(players)
            draw = self._editing_draw
            messagebox.showinfo("Updated",
                f"'{cat}' updated!\n"
                f"Entries: {len(players)}   Bracket: {draw.bracket_size}   "
                f"Byes: {draw.bye_count}")
        else:
            if active_folder.full():
                messagebox.showwarning("Full",
                    f"'{active_folder.name}' has {MAX_DRAWS} draws."); return
            draw = Draw(cat, players)
            active_folder.draws.append(draw)
            messagebox.showinfo("Created",
                f"'{cat}' generated!\n"
                f"Entries: {len(players)}   Bracket: {draw.bracket_size}   "
                f"Byes: {draw.bye_count}   Bouts: {draw.total_bouts}")

        self._reset()
        self.app.show_frame(ViewDrawsScreen)

    def _reset(self):
        self.cat_var.set(""); self.count_var.set("4")
        self._editing_draw = None
        for w in self.pf.winfo_children(): w.destroy()
        self.player_entries.clear()
        self.rtitle.config(text="Player Details")
        self._update_info(); self._refresh_list()

    def _sel_idx(self):
        s = self.draws_lb.curselection(); return s[0] if s else None

    def _rename_draw(self):
        idx = self._sel_idx()
        if idx is None: messagebox.showwarning("No Selection", "Select a draw."); return
        draw = _active_draws()[idx]
        new = simpledialog.askstring("Rename", f"New name:", parent=self)
        if new and new.strip(): draw.category = new.strip(); self._refresh_list()

    def _delete_draw(self):
        idx = self._sel_idx()
        if idx is None: messagebox.showwarning("No Selection", "Select a draw."); return
        draw = _active_draws()[idx]
        if not messagebox.askyesno("Delete", f"Delete '{draw.category}'?"): return
        active_folder.draws.pop(idx); self._refresh_list(); self._update_info()

    def _import_excel(self):
        """Browse for Excel/CSV, parse, preview, then populate the player grid."""
        path = filedialog.askopenfilename(
            title="Import Players from Excel / CSV",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xltx *.xls"),
                ("CSV files",   "*.csv"),
                ("All files",   "*.*"),
            ])
        if not path:
            return
        try:
            entries = _parse_import_file(path)
        except Exception as exc:
            messagebox.showerror("Import Error", str(exc)); return
        if not entries:
            messagebox.showwarning("No Data",
                "No player entries found.\n\n"
                "Ensure the file has a 'Name' column (and optionally\n"
                "a 'Club' / 'Gym' / 'Team' column)."); return

        dlg = _ImportPreviewDialog(self, entries, self.app)
        if not dlg.confirmed:
            return

        # Repopulate grid with imported players
        prefill = [Player(nm, cl) for nm, cl in entries]
        self.count_var.set(str(len(prefill)))
        self._set_count(prefill=prefill)
        messagebox.showinfo("Imported",
            f"{len(prefill)} player(s) loaded from:\n{path.split('/')[-1]}\n\n"
            "Set the Weight Category, then click Generate Draw.")


# ─────────────────────────────────────────────────────────────
# SCREEN 4 – VIEW DRAWS
# ─────────────────────────────────────────────────────────────

class ViewDrawsScreen(tk.Frame):
    """
    Smart bracket that starts in the correct round based on bracket_size.
    Full per-player left-click drag.  Right-click → rename boxer / swap / label.
    PDF export via reportlab.
    """

    def __init__(self, parent, app: BoxingApp):
        super().__init__(parent, bg=app.BG)
        self.app = app
        self.current_draw_idx: int | None = None

        # zoom state
        self._zoom      = 1.0
        self._zoom_min  = 0.4
        self._zoom_max  = 2.5
        self._zoom_step = 0.1
        self._zoom_var  = None   # tk.StringVar, set in _build

        # drag state
        self._drag: dict = {}
        self._ghost: int | None = None

        # lookup tables rebuilt on every render
        self._boxes: dict = {}

        # right-click state
        self._ctx_match: Match | None = None
        self._ctx_slot:  int   | None = None

        self._build()

    def _build(self):
        app = self.app
        bar = app.top_bar(self, "Tournament Draws", WelcomeScreen)
        self.folder_lbl = tk.Label(bar, text="", font=("Georgia", 10),
                                   bg=app.PANEL, fg=app.GOLD)
        self.folder_lbl.pack(side="right", padx=8)
        app.make_btn(bar, "➕ New Draw",
                     lambda: app.show_frame(CreateDrawScreen),
                     width=13, color=app.GOLD).pack(side="right", padx=8)

        paned = tk.PanedWindow(self, orient="horizontal",
                               bg=app.BG, sashwidth=6)
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # Sidebar
        sidebar = tk.Frame(paned, bg=app.PANEL, width=262)
        paned.add(sidebar, minsize=224)

        # ── Tournament name (used in PDF header) ─────────────
        tk.Label(sidebar, text="🏆 Tournament Name",
                 font=("Georgia", 9, "bold"),
                 bg=app.PANEL, fg=app.GOLD).pack(pady=(10, 2), padx=8, anchor="w")
        self.tournament_name_var = tk.StringVar(
            value="Maharashtra Boxing Association")
        tourn_entry = tk.Entry(
            sidebar, textvariable=self.tournament_name_var,
            font=("Georgia", 10), bg=app.ENTRY_BG, fg=app.ENTRY_FG,
            insertbackground="white", relief="flat")
        tourn_entry.pack(fill="x", padx=8, pady=(0, 6), ipady=5)

        tk.Frame(sidebar, bg=app.ACCENT, height=1).pack(fill="x", padx=8, pady=(0, 6))

        tk.Label(sidebar, text="Weight Categories",
                 font=("Georgia", 11, "bold"),
                 bg=app.PANEL, fg=app.GOLD).pack(pady=(0, 4))

        self.draw_lb = tk.Listbox(
            sidebar, font=("Georgia", 11),
            bg=app.ENTRY_BG, fg=app.TEXT,
            selectbackground=app.ACCENT, activestyle="none",
            relief="flat", borderwidth=0, highlightthickness=0)
        self.draw_lb.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.draw_lb.bind("<<ListboxSelect>>", self._on_select)

        sb_act = tk.Frame(sidebar, bg=app.PANEL)
        sb_act.pack(fill="x", padx=8, pady=2)
        app.make_btn(sb_act, "✏️ Rename", self._rename_draw,
                     width=11, color="#334", font_size=9).pack(side="left", padx=(0, 4))
        app.make_btn(sb_act, "🗑 Delete",  self._delete_draw,
                     width=11, color="#661", font_size=9).pack(side="left")

        app.make_btn(sidebar, "➕  Add Entries", self._add_entries,
                     width=24, color="#1a4a1a").pack(pady=4, padx=8, fill="x")
        app.make_btn(sidebar, "💾  Export as Text", self._export_txt,
                     width=24).pack(pady=2, padx=8, fill="x")
        app.make_btn(sidebar, "📄  Export as PDF",  self._export_pdf,
                     width=24, color="#2c3e7a").pack(pady=2, padx=8, fill="x")

        # Bracket area
        right = tk.Frame(paned, bg=app.BG)
        paned.add(right, minsize=780)

        # ── Editable draw info bar ────────────────────────────
        info_bar = tk.Frame(right, bg=app.ACCENT, pady=4)
        info_bar.pack(fill="x")

        tk.Label(info_bar, text="Category:", font=("Georgia", 9, "bold"),
                 bg=app.ACCENT, fg="white").pack(side="left", padx=(10, 3))
        self.edit_category_var = tk.StringVar(value="—")
        cat_entry = tk.Entry(info_bar, textvariable=self.edit_category_var,
                             font=("Georgia", 10, "bold"),
                             bg="#c0102a", fg="white",
                             insertbackground="white", relief="flat",
                             width=20)
        cat_entry.pack(side="left", padx=(0, 8), ipady=2)
        cat_entry.bind("<Return>",     self._apply_category_edit)
        cat_entry.bind("<FocusOut>",   self._apply_category_edit)

        self.summary_var = tk.StringVar(value="Select a draw to view.")
        tk.Label(info_bar, textvariable=self.summary_var,
                 font=("Georgia", 9), bg=app.ACCENT, fg="white",
                 anchor="w").pack(side="left", padx=4)

        # ── Inline edit panel ─────────────────────────────────
        edit_panel = tk.Frame(right, bg=app.PANEL, pady=4)
        edit_panel.pack(fill="x")

        tk.Label(edit_panel, text="✏️  Quick Edit:",
                 font=("Georgia", 9, "bold"),
                 bg=app.PANEL, fg=app.GOLD).pack(side="left", padx=(10, 6))

        # Player name quick-edit
        tk.Label(edit_panel, text="Select a box then edit here →",
                 font=("Georgia", 8), bg=app.PANEL,
                 fg=app.SUBTEXT).pack(side="left", padx=(0, 6))

        tk.Label(edit_panel, text="Name:", font=("Georgia", 9),
                 bg=app.PANEL, fg=app.TEXT).pack(side="left")
        self._qe_name_var = tk.StringVar()
        self._qe_name_entry = tk.Entry(
            edit_panel, textvariable=self._qe_name_var,
            font=("Georgia", 10), bg=app.ENTRY_BG, fg=app.ENTRY_FG,
            insertbackground="white", relief="flat", width=18)
        self._qe_name_entry.pack(side="left", padx=(2, 6), ipady=3)

        tk.Label(edit_panel, text="Club:", font=("Georgia", 9),
                 bg=app.PANEL, fg=app.TEXT).pack(side="left")
        self._qe_club_var = tk.StringVar()
        self._qe_club_entry = tk.Entry(
            edit_panel, textvariable=self._qe_club_var,
            font=("Georgia", 10), bg=app.ENTRY_BG, fg=app.ENTRY_FG,
            insertbackground="white", relief="flat", width=16)
        self._qe_club_entry.pack(side="left", padx=(2, 6), ipady=3)

        def _make_apply_btn():
            b = tk.Button(edit_panel, text="✅ Apply",
                          command=self._apply_quick_edit,
                          bg="#1a4a1a", fg="white",
                          font=("Georgia", 9, "bold"), relief="flat",
                          padx=8, pady=2, cursor="hand2")
            b.bind("<Enter>", lambda e: b.config(bg="#2a6a2a"))
            b.bind("<Leave>", lambda e: b.config(bg="#1a4a1a"))
            return b
        _make_apply_btn().pack(side="left", padx=(0, 4))

        # Bind Return on edit fields too
        self._qe_name_entry.bind("<Return>", lambda e: self._apply_quick_edit())
        self._qe_club_entry.bind("<Return>", lambda e: self._apply_quick_edit())

        # Track which box is selected for editing
        self._selected_box_key: tuple | None = None   # (col, m_idx, slot)

        # ── Zoom toolbar ──────────────────────────────────────
        ztb = tk.Frame(right, bg=app.PANEL, pady=5)
        ztb.pack(fill="x")

        tk.Label(ztb, text="🔍", font=("Arial", 11),
                 bg=app.PANEL, fg=app.GOLD).pack(side="left", padx=(10, 2))

        def _zb(txt, cmd):
            b = tk.Button(ztb, text=txt, command=cmd,
                          bg=app.ENTRY_BG, fg=app.TEXT,
                          font=("Helvetica", 9, "bold"),
                          relief="flat", padx=10, pady=3, cursor="hand2")
            b.bind("<Enter>", lambda e: b.config(bg=app.ACCENT))
            b.bind("<Leave>", lambda e: b.config(bg=app.ENTRY_BG))
            b.pack(side="left", padx=2)
            return b

        _zb("−", self._zoom_out)
        self._zoom_var = tk.StringVar(value="100%")
        tk.Label(ztb, textvariable=self._zoom_var,
                 font=("Courier", 9, "bold"), width=6,
                 bg=app.PANEL, fg=app.GOLD, anchor="center").pack(side="left")
        _zb("+", self._zoom_in)
        tk.Frame(ztb, bg=app.SUBTEXT, width=1).pack(
            side="left", fill="y", padx=8, pady=3)
        _zb("⬜ Fit",  self._zoom_fit)
        _zb("100%",   self._zoom_reset)

        tk.Label(ztb,
                 text="Ctrl+Scroll=zoom  •  Drag=swap  •  Right-click=options",
                 font=("Georgia", 7), bg=app.PANEL, fg=app.SUBTEXT
                 ).pack(side="right", padx=10)

        bf = tk.Frame(right, bg="#cccccc")
        bf.pack(fill="both", expand=True)

        h_sc = ttk.Scrollbar(bf, orient="horizontal")
        v_sc = ttk.Scrollbar(bf, orient="vertical")
        self.bc = tk.Canvas(bf, bg=app.BRACKET_BG, highlightthickness=0,
                            xscrollcommand=h_sc.set, yscrollcommand=v_sc.set)
        h_sc.config(command=self.bc.xview)
        v_sc.config(command=self.bc.yview)
        h_sc.pack(side="bottom", fill="x")
        v_sc.pack(side="right",  fill="y")
        self.bc.pack(fill="both", expand=True)

        # Right-click context menu [CHANGE 3]
        self.ctx = tk.Menu(self, tearoff=0, bg=app.PANEL, fg=app.TEXT,
                           activebackground=app.ACCENT, font=("Georgia", 10))
        self.ctx.add_command(label="✏️  Rename Boxer",
                             command=self._ctx_rename_boxer)
        self.ctx.add_command(label="🔄  Swap Corners (Red ↔ Blue)",
                             command=self._ctx_swap)
        self.ctx.add_command(label="🏷️  Rename Bout Label",
                             command=self._ctx_rename_bout)

        self.bc.bind("<Button-3>",       self._on_right_click)
        self.bc.bind("<ButtonPress-1>",   self._drag_start)
        self.bc.bind("<B1-Motion>",       self._drag_motion)
        self.bc.bind("<ButtonRelease-1>", self._drag_end)
        self.bc.bind("<ButtonPress-1>",   self._on_left_click_select, add="+")
        # Ctrl+scroll zoom (Windows & Linux)
        self.bc.bind("<Control-MouseWheel>", self._on_mousewheel)
        self.bc.bind("<Control-Button-4>",   lambda e: self._zoom_in())
        self.bc.bind("<Control-Button-5>",   lambda e: self._zoom_out())

    # ── Lifecycle ─────────────────────────────────────────────

    def on_show(self, **_):
        self.folder_lbl.config(
            text=f"Folder: {active_folder.path()}" if active_folder else "—")
        self._refresh_list()

    def _refresh_list(self):
        self.draw_lb.delete(0, "end")
        for i, d in enumerate(_active_draws()):
            self.draw_lb.insert("end", f"  {i+1:>2}. {d.category}")
        ads = _active_draws()
        if (self.current_draw_idx is not None
                and self.current_draw_idx < len(ads)):
            self.draw_lb.selection_set(self.current_draw_idx)
            self._render_bracket(ads[self.current_draw_idx])
            self.edit_category_var.set(ads[self.current_draw_idx].category)
        else:
            self.current_draw_idx = None
            self.bc.delete("all")
            self.summary_var.set("Select a draw to view.")
            self.edit_category_var.set("—")

    def _on_select(self, _=None):
        sel = self.draw_lb.curselection()
        if not sel: return
        self.current_draw_idx = sel[0]
        draw = _active_draws()[self.current_draw_idx]
        self.edit_category_var.set(draw.category)
        self._selected_box_key = None
        self._qe_name_var.set("")
        self._qe_club_var.set("")
        self._render_bracket(draw)

    def _rename_draw(self):
        if self.current_draw_idx is None:
            messagebox.showwarning("No Selection", "Select a draw."); return
        d = _active_draws()[self.current_draw_idx]
        new = simpledialog.askstring("Rename", f"New name for '{d.category}':",
                                     initialvalue=d.category, parent=self)
        if new and new.strip():
            d.category = new.strip()
            self.edit_category_var.set(d.category)
            self._refresh_list()

    def _delete_draw(self):
        if self.current_draw_idx is None:
            messagebox.showwarning("No Selection", "Select a draw."); return
        d = _active_draws()[self.current_draw_idx]
        if not messagebox.askyesno("Delete", f"Delete '{d.category}'?"): return
        active_folder.draws.pop(self.current_draw_idx)
        self.current_draw_idx = None; self._refresh_list()

    def _add_entries(self):
        if self.current_draw_idx is None:
            messagebox.showwarning("No Draw", "Select a draw."); return
        self.app.show_frame(CreateDrawScreen,
                            draw=_active_draws()[self.current_draw_idx])

    # ── INLINE EDITING METHODS ────────────────────────────────

    def _apply_category_edit(self, event=None):
        """Apply the edited category name from the top bar entry."""
        if self.current_draw_idx is None: return
        new_cat = self.edit_category_var.get().strip()
        if not new_cat: return
        draw = _active_draws()[self.current_draw_idx]
        draw.category = new_cat
        self._refresh_list()

    def _on_left_click_select(self, event):
        """Single left-click on a player box → select it for inline editing."""
        # Only act if this isn't a drag start (drag_start also fires)
        # We detect by checking after a short delay whether a drag happened.
        cx = self.bc.canvasx(event.x)
        cy = self.bc.canvasy(event.y)
        key, data = self._hit_box(cx, cy)
        if key is None: return
        col, m_idx, slot = key
        x1, y1, x2, y2, match, _ = data
        if match is None: return

        pdata = match.player1 if slot == 0 else match.player2
        if not isinstance(pdata, Player): return

        # Populate quick-edit panel
        self._selected_box_key = key
        self._qe_name_var.set(pdata.name)
        self._qe_club_var.set(pdata.club)
        self._qe_name_entry.focus_set()
        self._qe_name_entry.selection_range(0, "end")

        # Highlight the selected box on the canvas
        self._highlight_selected_box(x1, y1, x2, y2)

    def _highlight_selected_box(self, x1, y1, x2, y2):
        """Draw a gold selection outline around the chosen box."""
        self.bc.delete("selection_highlight")
        self.bc.create_rectangle(
            x1 - 2, y1 - 2, x2 + 2, y2 + 2,
            outline=self.app.GOLD, width=2,
            tags="selection_highlight")

    def _apply_quick_edit(self):
        """Apply name + club from the quick-edit panel to the selected box."""
        if self._selected_box_key is None:
            messagebox.showinfo("No Selection",
                "Click on a player box in the bracket first.")
            return
        if self.current_draw_idx is None: return

        key = self._selected_box_key
        col, m_idx, slot = key
        data = self._boxes.get(key)
        if data is None: return
        _, _, _, _, match, _ = data
        if match is None: return

        pdata = match.player1 if slot == 0 else match.player2
        if not isinstance(pdata, Player):
            messagebox.showinfo("Cannot Edit",
                "BYE and TBD slots cannot be edited directly.")
            return

        new_name = self._qe_name_var.get().strip()
        new_club = self._qe_club_var.get().strip()
        if not new_name:
            messagebox.showerror("Error", "Name cannot be empty.")
            return

        pdata.name = new_name
        pdata.club = new_club

        # Re-render to reflect changes
        draw = _active_draws()[self.current_draw_idx]
        self._render_bracket(draw)

        # Re-apply selection highlight after re-render
        new_data = self._boxes.get(key)
        if new_data:
            x1, y1, x2, y2 = new_data[:4]
            self._highlight_selected_box(x1, y1, x2, y2)

    # ── ZOOM ──────────────────────────────────────────────────

    def _zoom_in(self):
        self._set_zoom(self._zoom + self._zoom_step)

    def _zoom_out(self):
        self._set_zoom(self._zoom - self._zoom_step)

    def _zoom_reset(self):
        self._set_zoom(1.0)

    def _zoom_fit(self):
        if self.current_draw_idx is None:
            return
        # Natural (zoom=1) canvas size
        nat_w = 248 * NUM_COLS + 185 + 52 * 2
        nat_h = (33 * 2 + 4 + 20) * 16 + 58 * 2
        cw = self.bc.winfo_width()  or 800
        ch = self.bc.winfo_height() or 500
        fit = min(cw / nat_w, ch / nat_h, self._zoom_max)
        self._set_zoom(max(fit, self._zoom_min))

    def _set_zoom(self, z: float):
        z = round(max(self._zoom_min, min(self._zoom_max, z)), 2)
        self._zoom = z
        if self._zoom_var:
            self._zoom_var.set(f"{int(z * 100)}%")
        if self.current_draw_idx is not None:
            ads = _active_draws()
            if self.current_draw_idx < len(ads):
                self._render_bracket(ads[self.current_draw_idx])

    def _on_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    # ── BRACKET RENDERING ─────────────────────────────────────
    # Base layout constants at zoom = 1.0
    _BOX_W     = 185
    _BOX_H     = 33
    _INNER_GAP = 4
    _PAIR_GAP  = 20
    _COL_W     = 248
    _MARGIN_X  = 52
    _MARGIN_Y  = 58

    def _render_bracket(self, draw: Draw):
        """
        Bracket rendering — v7 fixes:

        1. Connector lines extend from the right edge of each pair all the
           way to the LEFT edge of the matching box in the next column,
           forming a proper H-bracket.  The horizontal stub from the previous
           version (only ~30px) is replaced by a full-width bridge.

        2. The champion box is removed.  The FINAL column already shows the
           two finalists; the layout now uses all 5 columns symmetrically
           with no extra column on the right.

        3. Placeholder boxes (rounds after the real data) show
           "W. Bout X vs W. Bout Y" — both feeding bouts — so each box
           is self-explanatory.

        Col 0  ROUND OF 32    → 16 bouts / 32 boxes
        Col 1  ROUND OF 16    →  8 bouts / 16 boxes
        Col 2  QUARTER-FINALS →  4 bouts /  8 boxes
        Col 3  SEMI-FINALS    →  2 bouts /  4 boxes
        Col 4  FINAL          →  1 bout  /  2 boxes
        """
        z   = self._zoom
        c   = self.bc
        app = self.app
        c.delete("all")
        self._boxes.clear()

        # ── Scaled layout values ──────────────────────────────
        BW  = self._BOX_W     * z
        BH  = self._BOX_H     * z
        IG  = self._INNER_GAP * z
        PG  = self._PAIR_GAP  * z
        # Spread all 5 columns across the available width.
        # We no longer need a champion column, so COL_W is wider.
        CW  = self._COL_W     * z * 1.12   # slightly wider for better spread
        MX  = self._MARGIN_X  * z
        MY  = self._MARGIN_Y  * z

        def fs(base): return max(1, int(base * z))

        start_col  = _BS_TO_COL.get(draw.bracket_size, 0)
        base_sh    = BH * 2 + IG + PG   # slot height at col 0
        N_BOUTS_C0 = 16

        total_h = base_sh * N_BOUTS_C0 + MY * 2
        total_w = CW * NUM_COLS + BW + MX * 2

        self.summary_var.set(
            f"  {draw.category}   │   Entries: {len(draw.players)}"
            f"   │   Bracket: {draw.bracket_size}   │   Byes: {draw.bye_count}"
            f"   │   Bouts: {draw.total_bouts}   │   Zoom: {int(z*100)}%")

        # ── Column headers ────────────────────────────────────
        for col in range(NUM_COLS):
            x      = MX + col * CW
            lbl    = _ROUND_LABELS[col]
            n_b    = N_BOUTS_C0 // (2 ** col)
            colour = "#aaaaaa" if col < start_col else "#111111"
            c.create_text(x + BW / 2, MY - 30 * z,
                          text=lbl, fill=colour,
                          font=("Georgia", fs(9), "bold"))
            c.create_text(x + BW / 2, MY - 16 * z,
                          text=f"{n_b} bouts / {n_b*2} boxes",
                          fill="#aaaaaa", font=("Arial", fs(7)))

        # ── Build a lookup: absolute bout number → Match object ───
        # Used so placeholder boxes can show feeding bout numbers
        # across ghost/real column boundaries.
        bout_map: dict[int, "Match"] = {}
        for rnd in draw.rounds:
            for m in rnd:
                bout_map[m.bout_number] = m

        # Pre-compute which match in the PREVIOUS round feeds each
        # (col, m_idx, slot) in a placeholder column.
        def _feed_label(col: int, m_idx: int, slot: int) -> str:
            """
            Each placeholder BOX shows the TWO bouts whose winners
            will fill this pair.

            The pair at (col, m_idx) is fed by TWO matches in col-1:
              feed_match_0 = draw.rounds[actual_r_prev][m_idx*2]
              feed_match_1 = draw.rounds[actual_r_prev][m_idx*2 + 1]

            slot==0 (RED box)  → winner of feed_match_0
            slot==1 (BLUE box) → winner of feed_match_1
            """
            actual_r_prev = (col - 1) - start_col
            if 0 <= actual_r_prev < len(draw.rounds):
                prev_rnd = draw.rounds[actual_r_prev]
                # Each match in col feeds ONE slot in the next col pair.
                # slot 0 of pair m_idx is fed by prev bout (m_idx*2)
                # slot 1 of pair m_idx is fed by prev bout (m_idx*2 + 1)
                prev_idx = m_idx * 2 + slot
                if prev_idx < len(prev_rnd):
                    m = prev_rnd[prev_idx]
                    return f"Winner of {m.bout_label}"
                # Fewer matches than expected — fall back to pair-level
                elif m_idx < len(prev_rnd):
                    m = prev_rnd[m_idx]
                    return f"Winner of {m.bout_label}"
            # Two levels back — show both bouts that feed into this pair
            actual_r_prev2 = (col - 2) - start_col
            if 0 <= actual_r_prev2 < len(draw.rounds):
                pr2 = draw.rounds[actual_r_prev2]
                base = m_idx * 4 + slot * 2
                labels = []
                for pi in [base, base + 1]:
                    if pi < len(pr2):
                        labels.append(f"B{pr2[pi].bout_number}")
                if labels:
                    return "W. " + " vs W. ".join(labels)
            return "TBD"

        # ── Draw all columns ──────────────────────────────────
        for col in range(NUM_COLS):
            x        = MX + col * CW
            actual_r = col - start_col
            sh       = base_sh * (2 ** col)
            n_bouts  = N_BOUTS_C0 // (2 ** col)
            col_matches = (draw.rounds[actual_r]
                           if 0 <= actual_r < len(draw.rounds) else [])

            for m_idx in range(n_bouts):
                y_ctr = MY + sh * m_idx + sh / 2
                y1    = y_ctr - BH - IG / 2   # top-box top edge
                y2    = y_ctr + IG / 2         # bottom-box top edge

                match = (col_matches[m_idx]
                         if m_idx < len(col_matches) else None)

                for slot, yb in [(0, y1), (1, y2)]:
                    cc = app.RED_CORNER if slot == 0 else app.BLU_CORNER

                    # ── GHOST column ──────────────────────────
                    if col < start_col:
                        c.create_rectangle(x, yb, x+BW, yb+BH,
                                           fill="#f4f4f4", outline="#dddddd")
                        c.create_rectangle(x, yb, x+5*z, yb+BH,
                                           fill="#e0e0e0", outline="#e0e0e0")
                        c.create_line(x+5*z, yb, x+BW, yb+BH,
                                      fill="#e0e0e0", dash=(4, 4))
                        c.create_line(x+5*z, yb+BH, x+BW, yb,
                                      fill="#e0e0e0", dash=(4, 4))

                    # ── PLACEHOLDER column ────────────────────
                    elif actual_r >= len(draw.rounds):
                        c.create_rectangle(x, yb, x+BW, yb+BH,
                                           fill="#fdfdf5", outline="#bbbbbb")
                        c.create_rectangle(x, yb, x+6*z, yb+BH,
                                           fill=cc, outline=cc)
                        feed_txt = _feed_label(col, m_idx, slot)
                        c.create_text(x + BW/2, yb + BH/2,
                                      text=feed_txt, fill="#888888",
                                      font=("Arial", fs(7), "italic"),
                                      width=BW - 12*z)

                    # ── REAL DATA column ──────────────────────
                    else:
                        pdata    = (match.player1 if slot == 0
                                    else match.player2) if match else None
                        is_bye   = (pdata == "BYE")
                        is_empty = (pdata is None)

                        c.create_rectangle(x, yb, x+BW, yb+BH,
                                           fill="#ffffff", outline="#333333")
                        c.create_rectangle(x, yb, x+6*z, yb+BH,
                                           fill=cc, outline=cc)

                        if is_empty:
                            # This slot is undecided — find which previous
                            # bout feeds it and show "Winner of Bout X".
                            # In round actual_r, match m_idx, slot 0 is fed
                            # by round actual_r-1 match (m_idx*2), slot 1
                            # by round actual_r-1 match (m_idx*2 + 1).
                            feed_txt = "TBD"
                            if actual_r > 0 and actual_r - 1 < len(draw.rounds):
                                prev_rnd   = draw.rounds[actual_r - 1]
                                prev_m_idx = m_idx * 2 + slot
                                if prev_m_idx < len(prev_rnd):
                                    feed_txt = f"Winner of {prev_rnd[prev_m_idx].bout_label}"
                                elif m_idx < len(prev_rnd):
                                    feed_txt = f"Winner of {prev_rnd[m_idx].bout_label}"
                            c.create_text(x + BW/2, yb + BH/2,
                                          text=feed_txt,
                                          fill="#999999",
                                          font=("Arial", fs(7), "italic"),
                                          width=BW - 14*z)
                        elif is_bye:
                            c.create_line(x+6*z, yb, x+BW, yb+BH,
                                          fill="#c0392b",
                                          width=max(1, int(2*z)))
                            c.create_line(x+6*z, yb+BH, x+BW, yb,
                                          fill="#c0392b",
                                          width=max(1, int(2*z)))
                            c.create_text(x+BW/2, yb+BH/2, text="BYE",
                                          fill="#c0392b",
                                          font=("Arial", fs(8), "bold"))
                        else:
                            nm = (pdata.name[:24] if isinstance(pdata, Player)
                                  else str(pdata)[:24])
                            cl = pdata.club[:28] if isinstance(pdata, Player) else ""
                            c.create_text(x+10*z, yb+9*z, text=nm,
                                          anchor="nw", fill="#111111",
                                          font=("Arial", fs(9), "bold"))
                            c.create_text(x+10*z, yb+21*z, text=cl,
                                          anchor="nw", fill="#555555",
                                          font=("Arial", fs(7)))

                        if slot == 0 and match:
                            c.create_text(x+BW-4*z, yb+3*z,
                                          text=match.bout_label,
                                          anchor="ne", fill="#777777",
                                          font=("Arial", fs(7)))

                        self._boxes[(col, m_idx, slot)] = (
                            x, yb, x+BW, yb+BH, match, slot)

                        # Gold outline if this box is selected for editing
                        if self._selected_box_key == (col, m_idx, slot):
                            c.create_rectangle(
                                x - 2, yb - 2, x+BW + 2, yb+BH + 2,
                                outline=app.GOLD, width=2,
                                tags="selection_highlight")

                # ── Per-box direct connector ──────────────────
                # Each individual box connects with a horizontal line
                # directly to the specific RED or BLUE box it feeds in
                # the next column.
                #
                # Mapping from current pair to next column:
                #   pair m_idx (col) feeds match m_idx//2 (col+1)
                #   if m_idx is even:  slot 0 → RED (slot 0), slot 1 → RED (slot 0)
                #                      Wait — one pair feeds ONE slot.
                #
                # Correct tournament mapping:
                #   col pair m_idx  →  col+1 match (m_idx // 2)
                #   m_idx even  → that match's slot 0 (RED box in next col)
                #   m_idx odd   → that match's slot 1 (BLUE box in next col)
                #
                # So BOTH boxes of the current pair (slot 0 and slot 1)
                # connect to THE SAME box in the next column:
                #   - their lines converge at a short vertical bar on the right
                #   - then a single horizontal line goes to the target box mid
                #
                # The target box vertical midpoint in next column:
                #   next match index   = m_idx // 2
                #   next slot          = m_idx % 2  (0=RED, 1=BLUE)
                #   next_sh            = base_sh * 2^(col+1)
                #   next y_ctr         = MY + next_sh * next_m + next_sh/2
                #   next y1            = next_y_ctr - BH - IG/2  (RED top edge)
                #   next y2            = next_y_ctr + IG/2        (BLUE top edge)
                #   target_ym          = next_y1 + BH/2  if next_slot==0
                #                      = next_y2 + BH/2  if next_slot==1
                if col < NUM_COLS - 1:
                    x_next   = MX + (col + 1) * CW   # left edge of next col

                    xr       = x + BW                 # right edge of current col
                    ym1      = y1 + BH / 2            # mid of RED box (slot 0)
                    ym2      = y2 + BH / 2            # mid of BLUE box (slot 1)

                    lc = "#cccccc" if col < start_col else "#888888"

                    # Compute the Y midpoint of the TARGET box in col+1
                    next_sh   = base_sh * (2 ** (col + 1))
                    next_m    = m_idx // 2             # which match in col+1
                    next_slot = m_idx % 2              # 0=RED target, 1=BLUE target

                    next_y_ctr = MY + next_sh * next_m + next_sh / 2
                    next_y1    = next_y_ctr - BH - IG / 2
                    next_y2    = next_y_ctr + IG / 2

                    if next_slot == 0:
                        target_ym = next_y1 + BH / 2   # RED box mid
                    else:
                        target_ym = next_y2 + BH / 2   # BLUE box mid

                    # Both boxes of the current pair merge at a short
                    # vertical bar just right of the current column,
                    # then a single line goes straight across to the
                    # target box mid in the next column.
                    bar_x = xr + 14 * z               # vertical junction bar

                    # Stub: each box → vertical bar at their own y
                    c.create_line(xr,    ym1,   bar_x, ym1,
                                  fill=lc, width=1)
                    c.create_line(xr,    ym2,   bar_x, ym2,
                                  fill=lc, width=1)
                    # Vertical bar connecting the two stubs
                    c.create_line(bar_x, ym1,   bar_x, ym2,
                                  fill=lc, width=1)
                    # Long horizontal from bar mid to target box in next col
                    # End point: past the colour strip (x_next + 6*z)
                    bar_mid = (ym1 + ym2) / 2
                    c.create_line(bar_x, bar_mid, x_next + 6 * z, target_ym,
                                  fill=lc, width=1)

        c.configure(scrollregion=(0, 0, total_w, total_h))

    # ── PER-PLAYER DRAG ───────────────────────────────────

    def _hit_box(self, cx, cy):
        """Return (col, m_idx, slot, data) for the box under (cx,cy), or None."""
        for key, data in self._boxes.items():
            x1, y1, x2, y2, match, slot = data
            if x1 <= cx <= x2 and y1 <= cy <= y2:
                return key, data
        return None, None

    def _drag_start(self, event):
        cx = self.bc.canvasx(event.x)
        cy = self.bc.canvasy(event.y)
        key, data = self._hit_box(cx, cy)
        if key is None: self._drag = {}; return
        col, m_idx, slot = key
        x1, y1, x2, y2, match, _ = data
        # Only allow drag in the starting column (real round 0 data)
        draw = (_active_draws()[self.current_draw_idx]
                if self.current_draw_idx is not None else None)
        if draw is None: self._drag = {}; return
        start_col = _BS_TO_COL.get(draw.bracket_size, 0)
        if col != start_col: self._drag = {}; return
        if match is None: self._drag = {}; return
        p = match.player1 if slot == 0 else match.player2
        if not isinstance(p, Player): self._drag = {}; return

        # Ghost rectangle
        self._ghost = self.bc.create_rectangle(
            x1, y1, x2, y2,
            outline=self.app.ACCENT, fill="", width=2,
            dash=(6, 3), tags="ghost")

        self._drag = {
            "src_col":   col,
            "src_m_idx": m_idx,
            "src_slot":  slot,
            "off_y":     cy - y1,
            "box_h":     y2 - y1,
        }

    def _drag_motion(self, event):
        if not self._drag or self._ghost is None: return
        cy = self.bc.canvasy(event.y)
        cx = self.bc.canvasx(event.x)
        bh = self._drag["box_h"]
        x1_src = None
        for key, data in self._boxes.items():
            if (key[0] == self._drag["src_col"] and
                    key[1] == self._drag["src_m_idx"] and
                    key[2] == self._drag["src_slot"]):
                x1_src = data[0]; break
        if x1_src is None: return
        new_y = cy - self._drag["off_y"]
        self.bc.coords(self._ghost, x1_src, new_y,
                       x1_src + self._BOX_W * self._zoom, new_y + bh)

    def _drag_end(self, event):
        self.bc.delete("ghost")
        self._ghost = None
        if not self._drag or self.current_draw_idx is None:
            self._drag = {}; return

        cx = self.bc.canvasx(event.x)
        cy = self.bc.canvasy(event.y)
        draw = _active_draws()[self.current_draw_idx]
        start_col = _BS_TO_COL.get(draw.bracket_size, 0)

        # Find target box in the same column
        tgt_key, tgt_data = self._hit_box(cx, cy)
        if (tgt_key is not None and
                tgt_key[0] == start_col and
                (tgt_key[1] != self._drag["src_m_idx"] or
                 tgt_key[2] != self._drag["src_slot"])):
            _, tgt_m_idx, tgt_slot = tgt_key
            draw.swap_players(0,
                              self._drag["src_m_idx"], self._drag["src_slot"],
                              tgt_m_idx, tgt_slot)
            self._render_bracket(draw)

        self._drag = {}

    # ── RIGHT-CLICK [CHANGE 3] ────────────────────────────────

    def _on_right_click(self, event):
        cx = self.bc.canvasx(event.x)
        cy = self.bc.canvasy(event.y)
        key, data = self._hit_box(cx, cy)
        if key is None: return
        _, _, _, _, match, slot = data
        if match is None: return
        self._ctx_match = match
        self._ctx_slot  = slot
        self.ctx.tk_popup(event.x_root, event.y_root)

    def _ctx_rename_boxer(self):
        """Right-click → Rename Boxer (name + club)."""
        if self._ctx_match is None: return
        m    = self._ctx_match
        slot = self._ctx_slot
        p    = m.player1 if slot == 0 else m.player2
        if not isinstance(p, Player):
            messagebox.showinfo("Info", "Cannot rename a BYE or TBD slot.")
            return
        new_name = simpledialog.askstring(
            "Rename Boxer", "Boxer name:", initialvalue=p.name, parent=self)
        if new_name and new_name.strip():
            p.name = new_name.strip()
        new_club = simpledialog.askstring(
            "Rename Boxer", "Club / Gym:", initialvalue=p.club, parent=self)
        if new_club is not None:
            p.club = new_club.strip()
        if self.current_draw_idx is not None:
            self._render_bracket(_active_draws()[self.current_draw_idx])

    def _ctx_swap(self):
        if self._ctx_match is None: return
        self._ctx_match.swap_corners()
        if self.current_draw_idx is not None:
            self._render_bracket(_active_draws()[self.current_draw_idx])

    def _ctx_rename_bout(self):
        if self._ctx_match is None: return
        new = simpledialog.askstring(
            "Rename Bout", f"Label for '{self._ctx_match.bout_label}':",
            parent=self)
        if new and new.strip():
            self._ctx_match.bout_label = new.strip()
            if self.current_draw_idx is not None:
                self._render_bracket(_active_draws()[self.current_draw_idx])

    # ── TEXT EXPORT ───────────────────────────────────────────

    def _export_txt(self):
        if self.current_draw_idx is None or not _active_draws():
            messagebox.showwarning("No Draw", "Select a draw."); return
        draw = _active_draws()[self.current_draw_idx]
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All", "*.*")],
            initialfile=f"draw_{draw.category.replace(' ','_')}.txt")
        if path:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(_draw_to_text(draw))
            messagebox.showinfo("Exported", f"Saved:\n{path}")

    # ── PDF EXPORT [CHANGE 4] ─────────────────────────────────

    def _export_pdf(self):
        if not PDF_OK:
            messagebox.showerror(
                "reportlab missing",
                "Install reportlab first:\n  pip install reportlab\n"
                "Then restart the app.")
            return
        if self.current_draw_idx is None or not _active_draws():
            messagebox.showwarning("No Draw", "Select a draw."); return

        draw = _active_draws()[self.current_draw_idx]
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
            initialfile=f"draw_{draw.category.replace(' ','_')}.pdf")
        if not path: return

        try:
            _export_draw_pdf(
                draw, path,
                folder_path=(active_folder.path() if active_folder else ""),
                tournament_name=self.tournament_name_var.get().strip()
                               or "Maharashtra Boxing Association")
            messagebox.showinfo("PDF Exported", f"Saved:\n{path}")
        except Exception as exc:
            messagebox.showerror("PDF Error", str(exc))


# ─────────────────────────────────────────────────────────────
# PDF GENERATION  — A4 PORTRAIT
# ─────────────────────────────────────────────────────────────

def _export_draw_pdf(draw: Draw, path: str, folder_path: str = "",
                     tournament_name: str = "Maharashtra Boxing Association"):
    """
    A4-PORTRAIT PDF draw sheet — v9:
      • Portrait orientation (A4: 595 x 841 pt).
      • 5 columns, no champion column.
      • Darker lines and borders for print clarity.
      • Per-box direct connectors matching the on-screen bracket:
          pair m_idx feeds match m_idx//2 in next col,
          slot = m_idx%2 (0=RED, 1=BLUE target box).
      • Undecided slots show "Winner of Bout X".
      • BYE shown with bold red cross.
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm

    PAGE_W, PAGE_H = A4   # portrait: 595.3 x 841.9 pt

    MARGIN_L  = 8  * mm
    MARGIN_R  = 8  * mm
    MARGIN_T  = 10 * mm
    MARGIN_B  = 8  * mm
    HEADER_H  = 24 * mm   # slightly taller for tournament name
    COL_HDR_H = 12 * mm

    avail_w = PAGE_W - MARGIN_L - MARGIN_R
    bracket_top_offset = HEADER_H + COL_HDR_H
    avail_h = PAGE_H - MARGIN_T - MARGIN_B - bracket_top_offset

    COL_W    = avail_w / NUM_COLS
    STRIP_W  = 5

    TOTAL_BOUTS_C0 = 16
    BOX_H    = avail_h / (TOTAL_BOUTS_C0 * 2)
    INNER_GAP = max(BOX_H * 0.08, 0.5)

    RED_C   = colors.HexColor("#a93226")
    BLUE_C  = colors.HexColor("#154360")
    BDR_C   = colors.HexColor("#222222")
    DARK_C  = colors.HexColor("#000000")
    DIM_C   = colors.HexColor("#333333")
    GHOST_C = colors.HexColor("#e0e0e0")
    GHOST_S = colors.HexColor("#999999")
    CONN_C  = colors.HexColor("#444444")

    start_col = _BS_TO_COL.get(draw.bracket_size, 0)
    bracket_y_origin = PAGE_H - MARGIN_T - bracket_top_offset

    def y_pdf(y_down):
        return bracket_y_origin - y_down

    cv = rl_canvas.Canvas(path, pagesize=A4)

    # ── Header bar ───────────────────────────────────────────
    hdr_bottom = PAGE_H - MARGIN_T - HEADER_H
    cv.setFillColor(colors.HexColor("#1a1a2e"))
    cv.rect(MARGIN_L, hdr_bottom, avail_w, HEADER_H, fill=1, stroke=0)
    # Tournament name in gold — user-editable before export
    cv.setFillColor(colors.HexColor("#f5a623"))
    cv.setFont("Helvetica-Bold", 13)
    cv.drawCentredString(PAGE_W / 2, hdr_bottom + HEADER_H * 0.68,
                         tournament_name)
    cv.setFillColor(colors.white)
    cv.setFont("Helvetica-Bold", 8)
    cv.drawCentredString(PAGE_W / 2, hdr_bottom + HEADER_H * 0.40,
                         f"Tournament Draw Sheet  \u2014  {draw.category}")
    cv.setFont("Helvetica", 6)
    cv.drawCentredString(PAGE_W / 2, hdr_bottom + HEADER_H * 0.14,
                         f"Entries: {len(draw.players)}  |  "
                         f"Bracket: {draw.bracket_size}  |  "
                         f"BYEs: {draw.bye_count}  |  "
                         f"Bouts: {draw.total_bouts}  |  "
                         f"Folder: {folder_path}  |  "
                         f"Date: {datetime.datetime.now().strftime('%d %b %Y  %H:%M')}")
    cv.setStrokeColor(colors.HexColor("#e94560"))
    cv.setLineWidth(1.2)
    sep_y = hdr_bottom
    cv.line(MARGIN_L, sep_y, PAGE_W - MARGIN_R, sep_y)

    # ── Column headers ───────────────────────────────────────
    for col in range(NUM_COLS):
        cx      = MARGIN_L + col * COL_W + COL_W / 2
        lbl     = _ROUND_LABELS[col]
        n_bouts = TOTAL_BOUTS_C0 // (2 ** col)
        sub     = f"{n_bouts} bouts / {n_bouts*2} boxes"
        fc = GHOST_S if col < start_col else DARK_C
        cv.setFillColor(fc)
        cv.setFont("Helvetica-Bold", 6)
        cv.drawCentredString(cx, sep_y - 4.5*mm, lbl)
        cv.setFillColor(GHOST_S)
        cv.setFont("Helvetica", 5)
        cv.drawCentredString(cx, sep_y - 8.5*mm, sub)

    cv.setStrokeColor(colors.HexColor("#cccccc"))
    cv.setLineWidth(0.4)
    cv.line(MARGIN_L, bracket_y_origin, PAGE_W - MARGIN_R, bracket_y_origin)

    # ── Draw bracket boxes and connectors ────────────────────
    base_sh = BOX_H * 2 + INNER_GAP

    for col in range(NUM_COLS):
        x_left   = MARGIN_L + col * COL_W
        actual_r = col - start_col
        bw       = COL_W - 1.5
        sh       = base_sh * (2 ** col)
        n_bouts  = TOTAL_BOUTS_C0 // (2 ** col)
        col_matches = (draw.rounds[actual_r]
                       if 0 <= actual_r < len(draw.rounds) else [])

        for m_idx in range(n_bouts):
            y_ctr  = sh * m_idx + sh / 2
            y1_top = y_ctr - BOX_H - INNER_GAP / 2
            y2_top = y_ctr + INNER_GAP / 2
            match  = col_matches[m_idx] if m_idx < len(col_matches) else None

            for slot, y_box_top in [(0, y1_top), (1, y2_top)]:
                pdata = (match.player1 if slot == 0
                         else match.player2) if match else None
                cc  = RED_C if slot == 0 else BLUE_C
                bx  = y_pdf(y_box_top + BOX_H)

                if col < start_col:
                    cv.setFillColor(GHOST_C); cv.setStrokeColor(GHOST_S)
                    cv.setLineWidth(0.3)
                    cv.rect(x_left, bx, bw, BOX_H, fill=1, stroke=1)
                    cv.setFillColor(GHOST_S)
                    cv.rect(x_left, bx, STRIP_W, BOX_H, fill=1, stroke=0)
                    cv.setStrokeColor(colors.HexColor("#cccccc"))
                    cv.setLineWidth(0.3)
                    cv.line(x_left+STRIP_W, bx, x_left+bw, bx+BOX_H)
                    cv.line(x_left+STRIP_W, bx+BOX_H, x_left+bw, bx)
                else:
                    cv.setFillColor(colors.white); cv.setStrokeColor(BDR_C)
                    cv.setLineWidth(0.6)
                    cv.rect(x_left, bx, bw, BOX_H, fill=1, stroke=1)
                    cv.setFillColor(cc)
                    cv.rect(x_left, bx, STRIP_W, BOX_H, fill=1, stroke=0)

                    is_bye   = (pdata == "BYE")
                    is_empty = (pdata is None)

                    if is_bye:
                        cv.setStrokeColor(RED_C); cv.setLineWidth(1.2)
                        cv.line(x_left+STRIP_W, bx, x_left+bw, bx+BOX_H)
                        cv.line(x_left+STRIP_W, bx+BOX_H, x_left+bw, bx)
                        cv.setFillColor(RED_C); cv.setFont("Helvetica-Bold", 5.5)
                        cv.drawCentredString(x_left+bw/2, bx+BOX_H/2-1.5, "BYE")
                    elif is_empty:
                        feed_txt = "TBD"
                        if actual_r > 0 and (actual_r-1) < len(draw.rounds):
                            prev_rnd   = draw.rounds[actual_r - 1]
                            prev_m_idx = m_idx * 2 + slot
                            if prev_m_idx < len(prev_rnd):
                                feed_txt = f"Winner of {prev_rnd[prev_m_idx].bout_label}"
                            elif m_idx < len(prev_rnd):
                                feed_txt = f"Winner of {prev_rnd[m_idx].bout_label}"
                        cv.setFillColor(colors.HexColor("#777777"))
                        cv.setFont("Helvetica-Oblique", 5)
                        cv.drawString(x_left+STRIP_W+2, bx+BOX_H/2-1.5,
                                      feed_txt[:34])
                    else:
                        nm = (pdata.name[:32] if isinstance(pdata, Player)
                              else str(pdata)[:32])
                        cl = pdata.club[:38] if isinstance(pdata, Player) else ""
                        tx = x_left + STRIP_W + 2
                        cv.setFillColor(DARK_C); cv.setFont("Helvetica-Bold", 6)
                        cv.drawString(tx, bx+BOX_H*0.62, nm)
                        cv.setFillColor(DIM_C); cv.setFont("Helvetica", 5)
                        cv.drawString(tx, bx+BOX_H*0.25, cl)

                    if slot == 0 and match is not None:
                        cv.setFillColor(colors.HexColor("#555555"))
                        cv.setFont("Helvetica", 4.5)
                        cv.drawRightString(x_left+bw-1, bx+BOX_H-2.5,
                                           match.bout_label)

        # Per-box direct connectors
        if col < NUM_COLS - 1:
            x_next  = MARGIN_L + (col + 1) * COL_W
            next_sh = base_sh * (2 ** (col + 1))
            lc      = GHOST_S if col < start_col else CONN_C

            for m_idx in range(n_bouts):
                y_ctr  = sh * m_idx + sh / 2
                y1_top = y_ctr - BOX_H - INNER_GAP / 2
                y2_top = y_ctr + INNER_GAP / 2
                ym1 = y_pdf(y1_top + BOX_H / 2)
                ym2 = y_pdf(y2_top + BOX_H / 2)
                next_m    = m_idx // 2
                next_slot = m_idx % 2
                next_y_ctr = next_sh * next_m + next_sh / 2
                next_y1    = next_y_ctr - BOX_H - INNER_GAP / 2
                next_y2    = next_y_ctr + INNER_GAP / 2
                target_ym  = (y_pdf(next_y1 + BOX_H / 2)
                               if next_slot == 0
                               else y_pdf(next_y2 + BOX_H / 2))
                bar_x = x_left + bw + 6
                cv.setStrokeColor(lc); cv.setLineWidth(0.6)
                cv.line(x_left+bw, ym1,  bar_x, ym1)
                cv.line(x_left+bw, ym2,  bar_x, ym2)
                cv.line(bar_x,     ym1,  bar_x, ym2)
                bar_mid = (ym1 + ym2) / 2
                cv.line(bar_x, bar_mid, x_next + STRIP_W, target_ym)

    # ── Footer ───────────────────────────────────────────────
    cv.setFillColor(colors.HexColor("#888888"))
    cv.setFont("Helvetica", 5)
    cv.drawCentredString(PAGE_W / 2, MARGIN_B * 0.5,
                         "Maharashtra Boxing Association  •  "
                         "Official Draw Sheet  •  "
                         "Generated by Boxing Draw Generator")
    cv.save()

def _draw_to_text(draw: Draw) -> str:
    def fmt(p):
        if p is None: return "TBD"
        if p == "BYE": return "BYE"
        return f"{p.name} ({p.club})"

    lines = [
        "=" * 66,
        "  MAHARASHTRA BOXING ASSOCIATION – TOURNAMENT DRAW",
        f"  Generated : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "=" * 66,
        f"  Category  : {draw.category}",
        f"  Entries   : {len(draw.players)}",
        f"  Bracket   : {draw.bracket_size}",
        f"  BYEs      : {draw.bye_count}",
        f"  Bouts     : {draw.total_bouts}",
        "",
    ]
    start_col = _BS_TO_COL.get(draw.bracket_size, 0)
    for r_idx, rnd in enumerate(draw.rounds):
        col   = start_col + r_idx
        rname = _ROUND_LABELS[col] if col < len(_ROUND_LABELS) else f"ROUND {r_idx+1}"
        lines.append(f"{'─'*22}  {rname}  {'─'*22}")
        for m in rnd:
            adv = ""
            if m.winner and not m.is_bye():
                adv = f"  → {fmt(m.winner)}"
            elif m.is_bye():
                adv = f"  → Advances: {fmt(m.winner)}"
            lines.append(
                f"  {m.bout_label:<14}  RED: {fmt(m.player1):<28}"
                f"  BLUE: {fmt(m.player2)}{adv}")
        lines.append("")
    lines.append("=" * 66)
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if "--demo" in sys.argv:
        import random
        random.seed(42)
        players = [Player(f"Fighter {i}", f"Club {chr(64+i)}")
                   for i in range(1, 8)]
        draw = Draw("28–30 kg", players)
        print(_draw_to_text(draw))
    else:
        BoxingApp().mainloop()
