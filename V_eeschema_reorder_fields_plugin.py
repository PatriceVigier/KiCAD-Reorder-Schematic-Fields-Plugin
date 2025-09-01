# -*- coding: utf-8 -*-
"""
Reorder Schematic Fields (GUI) — KiCad 6/7/8/9
ABSOLUTE ORDER — v7.2.0 (per-schematic JSON only)

What this plugin does
---------------------
For every schematic symbol in a .kicad_sch file, it reorders *user* fields
(i.e. `(property "..." "...")` blocks that are not KiCad core/meta fields)
so that their order *matches exactly* the list you define in the GUI.

v7.2.0
------
- Removed project-default JSON entirely (no .reorder_fields.json reading/writing).
- Per-schematic persistence only: <schema>.kicad_sch.reorder.json.
- UI/help strings updated accordingly.

Design goals
------------
- **Safety**: We never touch KiCad core/meta fields (Reference, Value, Footprint,
  Datasheet, Description, and `ki_*` meta properties). Only user properties
  are reordered.
- **Robust parsing**: Simple S-expression reader ignores parentheses inside
  quoted strings and handles `\` escapes. Operates per `(symbol ...)` segment.
- **Atomic writes**: The original .kicad_sch is atomically replaced and a `.bak`
  backup is created next to it.
- **JSON persistence (per-schematic only)**: Your GUI order is saved as
  `<schema>.kicad_sch.reorder.json`. No project-wide fallback file is used.
- **Stable UX**: Single-selection list with Up/Down buttons and Alt+Up/Alt+Down
  shortcuts. The same row keeps moving on repeated clicks and stays focused.

License: MIT
Author: Patrice Vigier + assistant
"""
from __future__ import annotations

import io
import os
import re
import json
from pathlib import Path
from datetime import datetime

import wx
import pcbnew

# ------------------------- Configuration -------------------------
# Field names that must NOT be reordered (KiCad core/meta). Compared
# case-insensitively after trimming whitespace.
INTERNAL_FIELDS = {
    "ki_keywords", "ki_fp_filters", "ki_description",  # KiCad meta
    "name", "reference", "value", "footprint", "datasheet", "description",
}

# Default window sizing (adjust to taste)
DEFAULT_MIN_SIZE = (500, 500)
DEFAULT_WIN_SIZE = (600, 700)  # width, height

# ---------------------- Small helper utilities -------------------
def _norm(s: str) -> str:
    """Normalize a property name for comparisons (trim + lowercase)."""
    return (s or "").strip().lower()

def _is_internal(name: str) -> bool:
    """Return True if the field name is a KiCad core/meta field we must not touch."""
    return _norm(name) in INTERNAL_FIELDS

# Patterns to detect S-expression elements quickly at line starts.
RE_SYM_START  = re.compile(r'^\s*\(symbol\b')
RE_PROP_START = re.compile(r'^\s*\(property\s+"([^"]+)"\s+"([^"]*)"', re.UNICODE)


def confirm_schematic_closed() -> bool:
    """Warn to ensure the schematic is closed in Eeschema (avoid cache confusion)."""
    return wx.MessageBox(
        "\u26a0\ufe0f Make sure the schematic (.kicad_sch) is CLOSED in Eeschema before applying.\n\n"
        "If it's still open: Cancel, close it, then run the plugin again.\n\nContinue?",
        "Reorder schematic fields",
        wx.OK | wx.CANCEL | wx.ICON_WARNING,
    ) == wx.OK

# ---------------------- [Restart helper] ----------------------
def _close_kicad_safely():
    """
    Try to close KiCad gracefully:
      1) Close the Pcbnew frame (preferred).
      2) Close all top-level frames (schematic, footprint, etc.).
      3) Ask the app to exit its main loop (last resort).
    """
    # 1) Preferred: close the Pcbnew frame
    try:
        frame = pcbnew.GetPcbFrame()
        if frame:
            frame.Close(True)
            return
    except Exception:
        pass
    # 2) Close all top-level frames
    try:
        for w in wx.GetTopLevelWindows():
            if isinstance(w, wx.Frame):
                w.Close(True)
        return
    except Exception:
        pass
    # 3) Last resort
    try:
        app = wx.GetApp()
        if hasattr(app, "ExitMainLoop"):
            app.ExitMainLoop()
    except Exception:
        pass

def _prompt_restart(parent=None) -> None:
    """Ask once if the user wants to close KiCad now; close if Yes."""
    dlg = wx.MessageDialog(
        parent,
        "Changes saved.\nKiCad should restart to apply them.\n\nClose KiCad now?",
        "Restart KiCad",
        style=wx.YES_NO | wx.ICON_QUESTION | wx.NO_DEFAULT
    )
    res = dlg.ShowModal()
    dlg.Destroy()
    if res == wx.ID_YES:
        _close_kicad_safely()
# ---------------------------------------------------------------

# ---------------------- Robust S-expression helpers ----------------------
def find_block_end(lines: list[str], start_line: int, limit_line: int | None = None) -> int:
    """Return the line index of the closing ')' for the block starting at `start_line`."""
    depth = 0
    in_str = False
    esc = False
    seen_open = False
    last = limit_line if limit_line is not None else len(lines) - 1
    for i in range(start_line, last + 1):
        for ch in lines[i]:
            if esc:
                esc = False
                continue
            if ch == '\\':
                esc = True
                continue
            if ch == '"':
                in_str = not in_str
                continue
            if not in_str:
                if ch == '(':
                    depth += 1
                    seen_open = True
                elif ch == ')':
                    depth -= 1
                    if seen_open and depth == 0:
                        return i
    return last  # fallback

def find_symbol_bounds(lines: list[str]) -> list[tuple[int, int]]:
    """Return a list of (start_line, end_line) for each `(symbol ...)` block."""
    i = 0
    out: list[tuple[int, int]] = []
    while i < len(lines):
        if RE_SYM_START.match(lines[i]):
            j = find_block_end(lines, i)
            out.append((i, j))
            i = j + 1
        else:
            i += 1
    return out

def extract_properties(lines: list[str], s0: int, s1: int):
    """Extract all `(property "name" "value" ...)` blocks within [s0..s1]."""
    i = s0
    props = []
    while i <= s1:
        m = RE_PROP_START.match(lines[i])
        if m:
            end = find_block_end(lines, i, s1)
            name = m.group(1)
            internal = _is_internal(name)
            block = lines[i:end + 1]
            props.append({"name": name, "start": i, "end": end, "text": block, "internal": internal})
            i = end + 1
        else:
            i += 1
    user_props = [p for p in props if not p["internal"]]
    return props, user_props

# --------------------- Absolute ordering logic ---------------------
def absolute_order_user_props(user_props, gui_list: list[str]):
    """Return [(name, block)] in exact GUI order for present fields; others follow."""
    by = {_norm(p["name"]): p for p in user_props}
    original = [(p["name"], p["text"]) for p in user_props]

    listed = []
    seen = set()
    for name in gui_list:
        p = by.get(_norm(name))
        if p:
            listed.append((p["name"], p["text"]))
            seen.add(_norm(p["name"]))
    tail = [(n, b) for (n, b) in original if _norm(n) not in seen]
    return listed + tail

# --------------------------- Safe file I/O ---------------------------
def write_atomic(path: Path, content: str):
    """Atomically write to `path` and rotate existing file to `.bak`."""
    tmp = path.with_suffix(path.suffix + ".tmp___")
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(content)
    bak = path.with_suffix(path.suffix + ".bak")
    try:
        if bak.exists():
            bak.unlink()
    except Exception:
        pass
    try:
        path.replace(bak)
    except FileNotFoundError:
        pass
    os.replace(tmp, path)

# -------------------- JSON persistence (per-schematic) --------------------
def json_path_for_schematic(sch_path: Path) -> Path:
    """Per-schematic JSON path, e.g. `foo.kicad_sch.reorder.json`."""
    return sch_path.with_suffix(sch_path.suffix + ".reorder.json")

def atomic_write_json(path: Path, obj):
    """Atomically write a JSON object to disk (pretty-printed)."""
    data = json.dumps(obj, ensure_ascii=False, indent=2)
    tmp = path.with_suffix(path.suffix + ".tmp___")
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(data)
    os.replace(tmp, path)

def save_order_json_for_schematic(sch_path: Path, order: list[str]):
    payload = {
        "schema": str(sch_path),
        "order": order,
        "updated": datetime.now().isoformat(timespec="seconds"),
        "note": "Per-schematic saved order for Reorder Schematic Fields plugin",
    }
    atomic_write_json(json_path_for_schematic(sch_path), payload)

def load_order_json(path: Path) -> list[str] | None:
    """Load a JSON file and return an `order` list if available."""
    try:
        if not path.exists():
            return None
        with io.open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and isinstance(data.get("order"), list):
            return data["order"]
        return None
    except Exception:
        return None

def reset_json_for_schematic(sch_path: Path) -> bool:
    p = json_path_for_schematic(sch_path)
    try:
        if p.exists():
            p.unlink()
            return True
    except Exception:
        pass
    return False

def merge_saved_with_detected(saved: list[str] | None, detected: list[str]) -> list[str]:
    """Merge saved order with actually detected fields (saved ∩ detected, then remaining detected)."""
    if not saved:
        return list(detected)
    det_norm = {_norm(x) for x in detected}
    merged = [name for name in saved if _norm(name) in det_norm]
    seen_norm = {_norm(n) for n in merged}
    tail = [n for n in detected if _norm(n) not in seen_norm]
    return merged + tail

# ------------------------- Per-symbol processing -------------------------
def process_symbol_segment(seg: list[str], gui_list: list[str]) -> tuple[bool, list[str]]:
    """Reorder only *user* `(property ...)` blocks within a single `(symbol ...)` segment."""
    props, user_props = extract_properties(seg, 0, len(seg) - 1)
    if not user_props:
        return False, seg

    gui_list = [w for w in gui_list if not _is_internal(w)]
    before = [p["name"] for p in user_props]
    ordered = absolute_order_user_props(user_props, gui_list)
    after = [n for n, _ in ordered]

    if [n.lower() for n in before] == [n.lower() for n in after]:
        return False, seg

    # Remove only user property blocks (descending order keeps indices valid)
    for p in sorted(user_props, key=lambda x: x["start"], reverse=True):
        del seg[p["start"]:p["end"] + 1]

    # Reinsert at earliest user property position
    insert_at = min(p["start"] for p in user_props)
    for _, block in ordered:
        seg[insert_at:insert_at] = block
        insert_at += len(block)

    return True, seg

def process_file(path: Path, gui_list: list[str], dry=False, verbose=False) -> bool:
    """Reorder user fields for all symbols in the schematic file."""
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines(keepends=True)

    bounds = find_symbol_bounds(lines)
    changed_any = False

    # Process from end to start so replacing slices won't move later indices
    for (s0, s1) in reversed(bounds):
        seg = lines[s0:s1+1]
        changed, seg_new = process_symbol_segment(seg, gui_list)
        if changed:
            lines[s0:s1+1] = seg_new
            changed_any = True
            if verbose:
                print(f"{path.name}: reordered symbol at lines {s0}-{s1}")

    if changed_any and not dry:
        write_atomic(path, "".join(lines))
    return changed_any

def collect_field_names_from_file(path: Path) -> list[str]:
    """Scan the file and collect unique *user* property names (sorted)."""
    names = set()
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        m = RE_PROP_START.match(line)
        if m:
            name = m.group(1)
            if not _is_internal(name):
                names.add(name)
    return sorted(names)

# ------------------------------- GUI --------------------------------
class ReorderDialog(wx.Dialog):
    """Main dialog for selecting target order and applying it to a schematic."""

    def __init__(self, parent, prefill_path: Path | None):
        super().__init__(parent, title="Reorder schematic fields",
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        # Window size
        self.SetMinSize(wx.Size(*DEFAULT_MIN_SIZE))

        # Data model and current schematic handle
        self.items: list[str] = []
        self.current_sch: Path | None = None

        # ---- Build UI ----
        pnl = wx.Panel(self)
        v = wx.BoxSizer(wx.VERTICAL)

        # File chooser row
        file_row = wx.BoxSizer(wx.HORIZONTAL)
        file_row.Add(wx.StaticText(pnl, label="Schematic:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self.txt_file = wx.TextCtrl(pnl, style=wx.TE_READONLY)
        self.btn_browse = wx.Button(pnl, label="Browse…")
        self.btn_browse.SetToolTip("By default the current schematic path is selected.\n"
                                   "Choose a different one if you want.")
        file_row.Add(self.txt_file, 1, wx.EXPAND | wx.RIGHT, 6)
        file_row.Add(self.btn_browse, 0)
        v.Add(file_row, 0, wx.ALL | wx.EXPAND, 8)

        # Note / hint
        note = wx.StaticText(
            pnl,
            label=("ABSOLUTE ORDER: user fields present in each symbol will be reordered to\n"
                   "match EXACTLY the list below. Core/meta fields (Reference, Value,\n"
                   "Footprint, Datasheet, Description, ki_*) are untouched.\n\n"
                   "\U0001F4BE Order is saved ONLY per schematic JSON:\n"
                   "<schema>.kicad_sch.reorder.json (no project default)."),
        )
        note.Wrap(720)
        v.Add(note, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # Order list + buttons (single selection)
        v.Add(wx.StaticText(pnl, label="Target order (top = first):"), 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        list_row = wx.BoxSizer(wx.HORIZONTAL)
        self.listbox = wx.ListBox(pnl, choices=[], style=wx.LB_SINGLE)
        col_btns = wx.BoxSizer(wx.VERTICAL)
        self.btn_up = wx.Button(pnl, label="↑ Move Up")
        self.btn_down = wx.Button(pnl, label="↓ Move Down")
        self.btn_reset_file = wx.Button(pnl, label="Reset this schematic order")
        self.btn_reset_file.SetToolTip(
            "Delete the per-schematic saved order JSON of the current schematic.\n"
            "Next time, the detected order will be used (until you Save again)."
        )
        col_btns.Add(self.btn_up, 0, wx.BOTTOM, 6)
        col_btns.Add(self.btn_down, 0, wx.BOTTOM, 12)
        col_btns.Add(self.btn_reset_file, 0)
        list_row.Add(self.listbox, 1, wx.EXPAND | wx.RIGHT, 6)
        list_row.Add(col_btns, 0, wx.ALIGN_TOP)
        v.Add(list_row, 1, wx.ALL | wx.EXPAND, 8)

        # Dry-run option
        self.chk_dry = wx.CheckBox(pnl, label="Dry-run (no write, preview only)")
        self.chk_dry.SetValue(False)
        v.Add(self.chk_dry, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # Action buttons
        btns = wx.StdDialogButtonSizer()
        self.btn_apply = wx.Button(pnl, wx.ID_APPLY, "Save")
        self.btn_close = wx.Button(pnl, wx.ID_CLOSE, "Cancel")
        btns.AddButton(self.btn_apply)
        btns.AddButton(self.btn_close)
        btns.Realize()
        v.Add(btns, 0, wx.ALL | wx.EXPAND, 8)

        # Bind events
        self.btn_browse.Bind(wx.EVT_BUTTON, self.on_browse)
        self.btn_up.Bind(wx.EVT_BUTTON, self.on_up_click)
        self.btn_down.Bind(wx.EVT_BUTTON, self.on_down_click)
        self.btn_reset_file.Bind(wx.EVT_BUTTON, self.on_reset_file)
        self.btn_apply.Bind(wx.EVT_BUTTON, self.on_apply)
        self.btn_close.Bind(wx.EVT_BUTTON, lambda _e: self.EndModal(wx.ID_CLOSE))
        self.listbox.Bind(wx.EVT_KEY_DOWN, self._on_list_key)  # Alt+Up/Down shortcuts

        # Layout and sizing
        pnl.SetSizer(v)
        root = wx.BoxSizer(wx.VERTICAL)
        root.Add(pnl, 1, wx.EXPAND)
        self.SetSizer(root)
        self.Layout()
        root.Fit(self)
        self.SetSize(wx.Size(*DEFAULT_WIN_SIZE))
        self.CentreOnScreen()

        # Prefill from current PCB (convenience)
        if prefill_path and prefill_path.exists():
            self._load_from_path(prefill_path)

    # ------------------ Helpers: loading / list operations ------------------
    def _load_from_path(self, sch_path: Path):
        """Load detected names and apply any saved per-schematic order JSON."""
        self.current_sch = sch_path
        self.txt_file.SetValue(str(sch_path))

        detected = collect_field_names_from_file(sch_path)
        saved_file = load_order_json(json_path_for_schematic(sch_path))
        self.items = merge_saved_with_detected(saved_file, detected)

        self.listbox.Set(self.items)
        if self.listbox.GetCount() > 0:
            self.listbox.SetSelection(0)

    def _swap_item_labels(self, i: int, j: int):
        """Swap two entries both in the model and in the widget without recreating the list."""
        self.items[i], self.items[j] = self.items[j], self.items[i]
        self.listbox.SetString(i, self.items[i])
        self.listbox.SetString(j, self.items[j])

    def _move_selected_up(self):
        idx = self.listbox.GetSelection()
        if idx == wx.NOT_FOUND or idx <= 0:
            return
        self._swap_item_labels(idx - 1, idx)
        self.listbox.SetSelection(idx - 1)
        wx.CallAfter(self.listbox.SetFocus)

    def _move_selected_down(self):
        idx = self.listbox.GetSelection()
        if idx == wx.NOT_FOUND or idx >= self.listbox.GetCount() - 1:
            return
        self._swap_item_labels(idx, idx + 1)
        self.listbox.SetSelection(idx + 1)
        wx.CallAfter(self.listbox.SetFocus)

    # ----------------------------- Event handlers -----------------------------
    def on_browse(self, _evt):
        with wx.FileDialog(self, "Choose a schematic",
                           wildcard="KiCad schematic (*.kicad_sch)|*.kicad_sch",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            p = Path(dlg.GetPath())
            self._load_from_path(p)

    def on_up_click(self, _evt):
        self._move_selected_up()

    def on_down_click(self, _evt):
        self._move_selected_down()

    def _on_list_key(self, evt):
        code = evt.GetKeyCode()
        if evt.AltDown() and code in (wx.WXK_UP, wx.WXK_NUMPAD_UP):
            self._move_selected_up(); return
        if evt.AltDown() and code in (wx.WXK_DOWN, wx.WXK_NUMPAD_DOWN):
            self._move_selected_down(); return
        evt.Skip()

    def on_reset_file(self, _evt):
        if not self.current_sch:
            wx.MessageBox("No schematic loaded.", "Reset", wx.OK | wx.ICON_INFORMATION)
            return
        done = reset_json_for_schematic(self.current_sch)
        wx.MessageBox(
            "Per-schematic order removed." if done else "No per-schematic order file to remove.",
            "Reset (this schematic)", wx.OK | wx.ICON_INFORMATION
        )

    def on_apply(self, _evt):
        path_txt = self.txt_file.GetValue().strip()
        if not path_txt:
            wx.MessageBox("Please choose a .kicad_sch file first.", "Reorder fields", wx.ICON_WARNING)
            return
        sch_path = Path(path_txt)
        if not sch_path.exists():
            wx.MessageBox("File not found.", "Reorder fields", wx.ICON_WARNING)
            return

        gui_list = [self.listbox.GetString(i) for i in range(self.listbox.GetCount())]
        gui_list = [n for n in gui_list if not _is_internal(n)]
        dry = self.chk_dry.GetValue()

        try:
            changed = process_file(sch_path, gui_list, dry=dry, verbose=False)
        except Exception as e:
            wx.MessageBox(f"Error while processing file:\n{e}", "Reorder fields", wx.ICON_ERROR)
            return

        # Save only the per-schematic JSON (no project default)
        try:
            save_order_json_for_schematic(sch_path, gui_list)
        except Exception:
            pass

        per_file = json_path_for_schematic(sch_path).name
        wx.MessageBox(
            f"{'(dry-run) ' if dry else ''}"
            f"{'Modified' if changed else 'No change'}\n\n"
            f"Saved order file:\n  - {per_file}\n"
            f"Backup: {sch_path.with_suffix(sch_path.suffix + '.bak').name}",
            "Reorder schematic fields",
            wx.OK | wx.ICON_INFORMATION,
        )

        # Optional: propose restart (helps KiCad refresh cached data)
        if not dry:
            _prompt_restart(self)

# ---------------------------- Action Plugin ----------------------------
class ReorderSchematicFieldsPlugins(pcbnew.ActionPlugin):
    """KiCad ActionPlugin wrapper to expose the dialog in the UI."""

    def defaults(self):
        self.name = "Reorder schematic fields…"
        self.category = "Schematic utilities"
        self.description = (
            "ABSOLUTE ORDER of user (property ...) fields, per-schematic JSON persistence."
        )
        try:
            self.show_toolbar_button = True
        except Exception:
            pass
        try:
            here = Path(__file__).resolve().parent
            icon = here / 'V_eeschema_reorder_fields_plugin.png'
            self.icon_file_name = str(icon) if icon.exists() else ""
        except Exception:
            self.icon_file_name = ""

    def Run(self):
        if not confirm_schematic_closed():
            return
        # Try to guess a schematic next to the currently open PCB
        sch_guess = None
        try:
            pcb_path = Path(pcbnew.GetBoard().GetFileName())
            if pcb_path and pcb_path.exists():
                cand = pcb_path.with_suffix(".kicad_sch")
                if cand.exists():
                    sch_guess = cand
        except Exception:
            pass

        dlg = ReorderDialog(None, sch_guess)
        dlg.ShowModal()
        dlg.Destroy()

# Register the plugin with KiCad
ReorderSchematicFieldsPlugins().register()
