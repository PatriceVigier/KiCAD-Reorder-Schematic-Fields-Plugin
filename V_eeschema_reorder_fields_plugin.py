# -*- coding: utf-8 -*-
"""
Reorder Schematic Fields (GUI) — KiCad 6/7/8/9
ABSOLUTE ORDER — lean build v7.3.3

• Window title shows the schematic filename. The "Schematic: path" row is removed.
• Uses the schematic of the OPENED PCB automatically:
  - Tries <board>.kicad_sch, else searches the PCB folder for a .kicad_sch (best stem match).
• Case-insensitive matching for order (like v7.2.0 you preferred).
  → "MPN" and "mpn" are treated as the same for ordering,
     but field names in the file are NOT renamed.
• Works ONLY with fields already present in the schematic (no templates, no creation).
• On Save: backup .bak, reorder user fields, persist per-schematic order JSON,
  clean .lck / *.kicad_prl, and offer to close KiCad windows so it reloads cleanly.
• Cleans lock files also on OPEN (cross-platform).
• NEW: The note shows the REAL backup filename (e.g. VLR.kicad_sch.bak).

MIT License — Patrice Vigier + assistant
"""
from __future__ import annotations

import io
import os
import re
import sys
import json
from pathlib import Path
from datetime import datetime

import wx
import pcbnew

# -------------------- constants / helpers --------------------
INTERNAL_FIELDS = {
    "ki_keywords", "ki_fp_filters", "ki_description",
    "name", "reference", "value", "footprint", "datasheet", "description",
}

DEFAULT_MIN_SIZE = (560, 520)
DEFAULT_WIN_SIZE = (760, 740)

def _norm(s: str) -> str:
    return (s or "").strip().lower()

def _is_internal(name: str) -> bool:
    return _norm(name) in INTERNAL_FIELDS

RE_SYM_START  = re.compile(r'^\s*\(symbol\b')
RE_PROP_START = re.compile(r'^\s*\(property\s+"([^"]+)"\s+"([^"]*)"', re.UNICODE)

# -------------------- s-expression helpers --------------------
def find_block_end(lines: list[str], start_line: int, limit_line: int | None = None) -> int:
    depth = 0; in_str = False; esc = False; seen_open = False
    last = limit_line if limit_line is not None else len(lines) - 1
    for i in range(start_line, last + 1):
        for ch in lines[i]:
            if esc: esc = False; continue
            if ch == '\\': esc = True; continue
            if ch == '"': in_str = not in_str; continue
            if not in_str:
                if ch == '(': depth += 1; seen_open = True
                elif ch == ')':
                    depth -= 1
                    if seen_open and depth == 0:
                        return i
    return last

def find_symbol_bounds(lines: list[str]) -> list[tuple[int, int]]:
    i = 0; out = []
    while i < len(lines):
        if RE_SYM_START.match(lines[i]):
            j = find_block_end(lines, i)
            out.append((i, j))
            i = j + 1
        else:
            i += 1
    return out

def extract_properties(lines: list[str], s0: int, s1: int):
    i = s0; props = []
    while i <= s1:
        m = RE_PROP_START.match(lines[i])
        if m:
            end = find_block_end(lines, i, s1)
            name = m.group(1); internal = _is_internal(name)
            block = lines[i:end + 1]
            props.append({"name": name, "start": i, "end": end, "text": block, "internal": internal})
            i = end + 1
        else:
            i += 1
    user_props = [p for p in props if not p["internal"]]
    return props, user_props

# -------------------- lock handling (cross-platform) --------------------
def lock_path_for_schematic(sch_path: Path) -> Path:
    return sch_path.with_suffix(sch_path.suffix + ".lck")

def try_remove_stale_lock(sch_path: Path) -> bool:
    lp = lock_path_for_schematic(sch_path)
    try:
        if lp.exists():
            lp.unlink()
            return True
    except Exception:
        pass
    return False

def try_remove_project_locks(sch_folder: Path) -> list[Path]:
    removed = []
    try:
        for prl in sch_folder.glob("*.kicad_prl"):
            try:
                prl.unlink()
                removed.append(prl)
            except Exception:
                pass
    except Exception:
        pass
    return removed

# -------------------- file I/O --------------------
def write_atomic(path: Path, content: str):
    tmp = path.with_suffix(path.suffix + ".tmp___")
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(content)
    bak = path.with_suffix(path.suffix + ".bak")
    try:
        if bak.exists(): bak.unlink()
    except Exception:
        pass
    try:
        path.replace(bak)
    except FileNotFoundError:
        pass
    os.replace(tmp, path)

def json_path_for_schematic(sch_path: Path) -> Path:
    return sch_path.with_suffix(sch_path.suffix + ".reorder.json")

def atomic_write_json(path: Path, obj):
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
            p.unlink(); return True
    except Exception:
        pass
    return False

# -------------------- discover schematic --------------------
def _guess_schematic_from_board() -> Path | None:
    """
    Try <board>.kicad_sch, else search the board folder for a .kicad_sch.
    Prefer the one whose stem matches the PCB stem.
    """
    try:
        board_path = Path(pcbnew.GetBoard().GetFileName())
    except Exception:
        board_path = None
    if not board_path or not board_path.exists():
        return None
    # 1) <board>.kicad_sch
    cand = board_path.with_suffix(".kicad_sch")
    if cand.exists():
        return cand
    # 2) search folder
    folder = board_path.parent
    sch_files = sorted(folder.glob("*.kicad_sch"))
    if not sch_files:
        return None
    # prefer same stem
    same_stem = [p for p in sch_files if p.stem == board_path.stem]
    return same_stem[0] if same_stem else sch_files[0]

# -------------------- detect present field names (case-insensitive dedup) --------------------
def collect_field_names_present(path: Path) -> list[str]:
    """
    Return DISTINCT user-field names that appear in the schematic, regardless of value.
    Case-insensitive de-duplication: the first spelling encountered becomes the label.
    """
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines(keepends=False)
    seen = set()
    ordered_names = []
    for s0, s1 in find_symbol_bounds(lines):
        i = s0
        while i <= s1:
            m = RE_PROP_START.match(lines[i])
            if m:
                name = m.group(1)
                if not _is_internal(name):
                    key = _norm(name)
                    if key not in seen:
                        seen.add(key)
                        ordered_names.append(name)
                i = find_block_end(lines, i, s1) + 1
            else:
                i += 1
    return ordered_names

def reconcile_gui_with_present(current_gui: list[str], present: list[str]):
    """
    Keep only names that are PRESENT (case-insensitive), preserving current relative order.
    Append any new PRESENT names at the end (using their present label).
    """
    present_map = { _norm(n): n for n in present }  # key->label
    present_keys = set(present_map.keys())

    kept = []
    kept_keys = set()
    for n in current_gui:
        k = _norm(n)
        if k in present_keys and k not in kept_keys:
            kept.append(n)     # keep user's label
            kept_keys.add(k)

    added = []
    for k, label in present_map.items():
        if k not in kept_keys:
            kept.append(label)
            kept_keys.add(k)
            added.append(label)

    removed = [n for n in current_gui if _norm(n) not in present_keys]
    return kept, removed, added

# -------------------- per-symbol reordering --------------------
def absolute_order_user_props_casefold(user_props, gui_list: list[str]):
    """
    Return ordered (name, text) for user_props using case-insensitive matching
    against gui_list, then append remaining (stable).
    """
    wanted_keys = [_norm(x) for x in gui_list]
    selected = [False] * len(user_props)
    ordered = []

    for wkey in wanted_keys:
        for idx, p in enumerate(user_props):
            if not selected[idx] and _norm(p["name"]) == wkey:
                ordered.append((p["name"], p["text"]))
                selected[idx] = True

    for idx, p in enumerate(user_props):
        if not selected[idx]:
            ordered.append((p["name"], p["text"]))

    return ordered

def process_symbol_segment(seg: list[str], gui_list: list[str]) -> tuple[bool, list[str]]:
    props, user_props = extract_properties(seg, 0, len(seg) - 1)
    if not user_props:
        return False, seg
    before = [p["name"] for p in user_props]
    ordered = absolute_order_user_props_casefold(user_props, gui_list)
    after = [n for n, _ in ordered]
    if before == after:
        return False, seg
    for p in sorted(user_props, key=lambda x: x["start"], reverse=True):
        del seg[p["start"]:p["end"] + 1]
    insert_at = min(p["start"] for p in user_props)
    for _, block in ordered:
        seg[insert_at:insert_at] = block
        insert_at += len(block)
    return True, seg

def process_file_reorder_only(path: Path, gui_list: list[str]) -> tuple[bool, dict, int]:
    """
    Reorder only (no creation). Returns: changed_any, stats, symbols_touched
    """
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines(keepends=True)
    bounds = find_symbol_bounds(lines)
    changed_any = False
    stats = {"hits": {name: 0 for name in gui_list}}
    symbols_touched = 0

    for s0, s1 in reversed(bounds):
        seg = lines[s0:s1+1]
        uprops = extract_properties(seg, 0, len(seg) - 1)[1]
        for want in gui_list:
            wkey = _norm(want)
            for p in uprops:
                if _norm(p["name"]) == wkey:
                    stats["hits"][want] += 1
        changed, seg_new = process_symbol_segment(seg, gui_list)
        if changed:
            lines[s0:s1+1] = seg_new
            changed_any = True
            symbols_touched += 1

    if changed_any:
        write_atomic(path, "".join(lines))
    return changed_any, stats, symbols_touched

# -------------------- GUI --------------------
class ReorderDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="Reorder schematic fields",
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.SetMinSize(wx.Size(*DEFAULT_MIN_SIZE))

        self.items: list[str] = []
        self.current_sch: Path | None = None
        self.present_names: list[str] = []

        pnl = wx.Panel(self)
        v = wx.BoxSizer(wx.VERTICAL)

        # (Removed the "Schematic: path" row)

        # NOTE: Create the note now, fill label after schematic is loaded
        self.note = wx.StaticText(pnl, label="")
        self.note.Wrap(760)
        v.Add(self.note, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        self.btn_refresh = wx.Button(pnl, label="Refresh from schematic")
        v.Add(self.btn_refresh, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        v.Add(wx.StaticText(pnl, label="Target order (top = first):"), 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        list_row = wx.BoxSizer(wx.HORIZONTAL)
        self.listbox = wx.ListBox(pnl, choices=[], style=wx.LB_SINGLE)

        col = wx.BoxSizer(wx.VERTICAL)
        self.btn_up = wx.Button(pnl, label="↑ Move Up")
        self.btn_down = wx.Button(pnl, label="↓ Move Down")
        self.btn_reset_file = wx.Button(pnl, label="Reset this schematic order")
        col.Add(self.btn_up, 0, wx.BOTTOM, 6)
        col.Add(self.btn_down, 0, wx.BOTTOM, 12)
        col.Add(self.btn_reset_file, 0)

        list_row.Add(self.listbox, 1, wx.EXPAND | wx.RIGHT, 6)
        list_row.Add(col, 0, wx.ALIGN_TOP)
        v.Add(list_row, 1, wx.ALL | wx.EXPAND, 8)

        row = wx.BoxSizer(wx.HORIZONTAL)
        # Add stretch spacer first → buttons align right
        row.AddStretchSpacer(1)
        self.btn_apply = wx.Button(pnl, wx.ID_APPLY, "Save")
        row.Add(self.btn_apply, 0)
        self.btn_close = wx.Button(pnl, wx.ID_CLOSE, "Cancel")
        row.Add(self.btn_close, 0, wx.RIGHT, 6)
        v.Add(row, 0, wx.ALL | wx.EXPAND, 8)

        # binds
        self.btn_refresh.Bind(wx.EVT_BUTTON, self.on_refresh)
        self.btn_up.Bind(wx.EVT_BUTTON, self.on_up_click)
        self.btn_down.Bind(wx.EVT_BUTTON, self.on_down_click)
        self.btn_reset_file.Bind(wx.EVT_BUTTON, self.on_reset_file)
        self.btn_apply.Bind(wx.EVT_BUTTON, self.on_apply)
        self.btn_close.Bind(wx.EVT_BUTTON, lambda _e: self.EndModal(wx.ID_CLOSE))
        self.listbox.Bind(wx.EVT_KEY_DOWN, self._on_list_key)

        pnl.SetSizer(v)
        root = wx.BoxSizer(wx.VERTICAL)
        root.Add(pnl, 1, wx.EXPAND)
        self.SetSizer(root)
        self.Layout(); root.Fit(self); self.SetSize(wx.Size(*DEFAULT_WIN_SIZE)); self.CentreOnScreen()

        # Auto-load schematic from the opened PCB
        sch = _guess_schematic_from_board()

        # Clean locks on open (cross-platform)
        if sch and sch.exists():
            try_remove_stale_lock(sch)
            try_remove_project_locks(sch.parent)

        # Load and set window title with schematic name, then update note with real .bak name
        if sch and sch.exists():
            self._load_from_schematic(sch)
            self._update_title_with_schematic()
            self._update_note_bak()
        else:
            self.SetTitle("Reorder schematic fields — (no .kicad_sch next to PCB)")
            self._update_note_bak()  # generic message

    # helpers
    def _update_title_with_schematic(self):
        if self.current_sch:
            self.SetTitle(f"Reorder schematic fields — {self.current_sch.name}")

    def _update_note_bak(self):
        if self.current_sch:
           # bak_name = self.current_sch.with_suffix(self.current_sch.suffix + ".bak").name
            bak_name = str(self.current_sch.with_suffix(self.current_sch.suffix + ".bak"))
        else:
            bak_name = ".bak"
        self.note.SetLabel(
            "ABSOLUTE ORDER (case-insensitive match, names not renamed).\n"
            "• Works only on fields already present in the schematic.\n"
            "• Core/meta fields are untouched.\n"
            f"• A {bak_name} (backup) is written before changes."
        )
        self.note.Wrap(760)

    def _render_items(self):
        self.listbox.Set(self.items)
        if self.listbox.GetCount() > 0:
            self.listbox.SetSelection(0)

    def _load_from_schematic(self, sch_path: Path):
        self.current_sch = sch_path
        self.present_names = collect_field_names_present(sch_path)
        saved = load_order_json(json_path_for_schematic(sch_path))
        initial = saved if isinstance(saved, list) else []
        self.items, _, _ = reconcile_gui_with_present(initial, self.present_names)
        self._render_items()
        # keep note in sync if called later
        self._update_note_bak()

    def _refresh_items(self):
        if not self.current_sch:
            return [], [], 0
        self.present_names = collect_field_names_present(self.current_sch)
        final, removed, added = reconcile_gui_with_present(self.items, self.present_names)
        self.items = final
        self._render_items()
        return removed, added, len(self.present_names)

    # moving
    def _swap_item_labels(self, i: int, j: int):
        self.items[i], self.items[j] = self.items[j], self.items[i]
        self.listbox.SetString(i, self.items[i]); self.listbox.SetString(j, self.items[j])

    def _move_selected_up(self):
        idx = self.listbox.GetSelection()
        if idx == wx.NOT_FOUND or idx <= 0: return
        self._swap_item_labels(idx - 1, idx)
        self.listbox.SetSelection(idx - 1); wx.CallAfter(self.listbox.SetFocus)

    def _move_selected_down(self):
        idx = self.listbox.GetSelection()
        if idx == wx.NOT_FOUND or idx >= self.listbox.GetCount() - 1: return
        self._swap_item_labels(idx, idx + 1)
        self.listbox.SetSelection(idx + 1); wx.CallAfter(self.listbox.SetFocus)

    # events
    def on_refresh(self, _evt):
        if not self.current_sch:
            wx.MessageBox("No .kicad_sch was found next to the opened PCB.", "Refresh",
                          wx.OK | wx.ICON_INFORMATION); return
        removed, added, n_present = self._refresh_items()
        try: save_order_json_for_schematic(self.current_sch, self.items)
        except Exception: pass
        msg = [f"Present user fields: {n_present}"]
        if removed: msg.append("\nRemoved (no longer present):\n  - " + "\n  - ".join(removed))
        if added:   msg.append("\nAdded (newly present):\n  - " + "\n  - ".join(added))
        wx.MessageBox("".join(msg), "Refresh", wx.OK | wx.ICON_INFORMATION)

    def on_up_click(self, _evt):   self._move_selected_up()
    def on_down_click(self, _evt): self._move_selected_down()

    def _on_list_key(self, evt):
        code = evt.GetKeyCode()
        if evt.AltDown() and code in (wx.WXK_UP, wx.WXK_NUMPAD_UP):   self._move_selected_up();   return
        if evt.AltDown() and code in (wx.WXK_DOWN, wx.WXK_NUMPAD_DOWN): self._move_selected_down(); return
        evt.Skip()

    def on_reset_file(self, _evt):
        if not self.current_sch:
            wx.MessageBox("No schematic loaded.", "Reset", wx.OK | wx.ICON_INFORMATION); return
        done = reset_json_for_schematic(self.current_sch)
        wx.MessageBox("Per-schematic order removed." if done else "No per-schematic order file to remove.",
                      "Reset (this schematic)", wx.OK | wx.ICON_INFORMATION)
        self.present_names = collect_field_names_present(self.current_sch)
        self.items = list(self.present_names)
        self._render_items()
        try: save_order_json_for_schematic(self.current_sch, self.items)
        except Exception: pass
        # keep the note consistent
        self._update_note_bak()

    def _really_close_kicad(self):
        # Best effort: close top-level windows owned by this process (Pcbnew)
        try:
            frame = pcbnew.GetPcbFrame()
            if frame: frame.Close(True)
        except Exception: pass
        try:
            for w in wx.GetTopLevelWindows():
                try: w.Close(True)
                except Exception: pass
        except Exception: pass
        try:
            app = wx.GetApp()
            if hasattr(app, "ExitMainLoop"): app.ExitMainLoop()
        except Exception: pass

    def on_apply(self, _evt):
        if not self.current_sch:
            wx.MessageBox("No .kicad_sch was found next to the opened PCB.", "Reorder fields",
                          wx.OK | wx.ICON_WARNING); return
        sch_path = self.current_sch

        # Recompute present-only and persist reconciled order immediately
        self.present_names = collect_field_names_present(sch_path)
        final_order, removed, added = reconcile_gui_with_present(self.items, self.present_names)
        self.items = list(final_order); self._render_items()
        try: save_order_json_for_schematic(sch_path, self.items)
        except Exception: pass

        # Apply: REORDER ONLY
        try:
            changed, stats, n_syms = process_file_reorder_only(sch_path, final_order)
        except Exception as e:
            wx.MessageBox(f"Error while processing file:\n{e}", "Reorder fields", wx.ICON_ERROR); return

        # Clean locks after save (cross-platform)
        removed_sch_lock = try_remove_stale_lock(sch_path)
        removed_prls = try_remove_project_locks(sch_path.parent)

        per_file = json_path_for_schematic(sch_path).name
        zero = [n for n, cnt in (stats.get("hits") or {}).items() if cnt == 0]

        parts = []
        parts.append("Modified" if changed else "No change")
        parts.append(f"\n\nSaved order file:\n  - {per_file}")
        parts.append(f"\nBackup: {sch_path.with_suffix(sch_path.suffix + '.bak').name}")
        parts.append(f"\nSymbols reordered: {n_syms}")
        if removed_sch_lock:
            parts.append(f"\nRemoved stale schematic lock: {lock_path_for_schematic(sch_path).name}")
        if removed_prls:
            parts.append("\nRemoved project lock(s):" + "".join(f"\n  - {p.name}" for p in removed_prls))
        if removed:
            parts.append("\n\nRemoved (no longer present):\n  - " + "\n  - ".join(removed))
        if added:
            parts.append("\n\nAdded (newly present):\n  - " + "\n  - ".join(added))
        if zero:
            parts.append("\n\nNames with 0 placements (present but not on some symbols):\n  - " + "\n  - ".join(zero))
        wx.MessageBox("".join(parts), "Reorder schematic fields", wx.OK | wx.ICON_INFORMATION)

        # Offer restart (close this KiCad instance's windows)
        dlg = wx.MessageDialog(
            self,
            "Changes saved.\n"
            "To apply the changes properly, <b>KiCad needs to be closed</b>. \n"
            "Please reopen Eeschema to reload the schematic text.\n\n"
            "Close KiCad windows now?",
            "Close KiCad windows", style=wx.YES_NO | wx.ICON_QUESTION | wx.YES_DEFAULT
        )
        res = dlg.ShowModal(); dlg.Destroy()
        if res == wx.ID_YES:
            self._really_close_kicad()

# -------------------- action plugin --------------------
class ReorderSchematicFieldsPlugins(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "Reorder schematic fields…"
        self.category = "Schematic utilities"
        self.description = "ABSOLUTE ORDER of user fields already present (case-insensitive matching)."
        try: self.show_toolbar_button = True
        except Exception: pass
        try:
            here = Path(__file__).resolve().parent
            ico = here / "V_eeschema_reorder_fields_plugin.png"
            self.icon_file_name = str(ico) if ico.exists() else ""
        except Exception:
            self.icon_file_name = ""

    def Run(self):
        dlg = ReorderDialog(None)
        dlg.ShowModal()
        dlg.Destroy()

ReorderSchematicFieldsPlugins().register()
