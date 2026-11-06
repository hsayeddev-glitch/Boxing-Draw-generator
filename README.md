# Boxing-Draw-generator
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
