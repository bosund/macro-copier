# VBA Macro Copier
# Copyright (c) 2026 Bo Sundgaard — www.uniteapps.dk
# MIT License

__version__ = "1.00"

import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
VBA_CT = "application/vnd.ms-office.vbaProject"
XLSM_WB_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def find_source_file():
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    candidate = base / "source.xlsm"
    return candidate if candidate.exists() else None


def _xml_bytes(root):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="unicode")
    ).encode("utf-8")


def inject_macros(source_path: Path, target_path: Path, output_path: Path):
    with zipfile.ZipFile(source_path, "r") as src:
        if "xl/vbaProject.bin" not in src.namelist():
            raise ValueError("source.xlsm contains no VBA macros (vbaProject.bin not found)")
        vba_data = src.read("xl/vbaProject.bin")

    entries = {}
    with zipfile.ZipFile(target_path, "r") as tgt:
        for name in tgt.namelist():
            entries[name] = tgt.read(name)

    # Patch [Content_Types].xml
    ET.register_namespace("", NS_CT)
    ct_root = ET.fromstring(entries["[Content_Types].xml"])
    vba_ct_present = False
    for child in ct_root:
        pn = child.get("PartName", "")
        if pn == "/xl/workbook.xml":
            child.set("ContentType", XLSM_WB_CT)
        elif pn == "/xl/vbaProject.bin":
            vba_ct_present = True
    if not vba_ct_present:
        ov = ET.SubElement(ct_root, f"{{{NS_CT}}}Override")
        ov.set("PartName", "/xl/vbaProject.bin")
        ov.set("ContentType", VBA_CT)
    entries["[Content_Types].xml"] = _xml_bytes(ct_root)

    # Patch xl/_rels/workbook.xml.rels
    ET.register_namespace("", NS_REL)
    rels_key = "xl/_rels/workbook.xml.rels"
    if rels_key in entries:
        rels_root = ET.fromstring(entries[rels_key])
    else:
        rels_root = ET.Element(f"{{{NS_REL}}}Relationships")

    if not any(child.get("Type") == VBA_REL_TYPE for child in rels_root):
        existing_ids = {child.get("Id", "") for child in rels_root}
        i = 1
        while f"rId{i}" in existing_ids:
            i += 1
        rel = ET.SubElement(rels_root, f"{{{NS_REL}}}Relationship")
        rel.set("Id", f"rId{i}")
        rel.set("Type", VBA_REL_TYPE)
        rel.set("Target", "vbaProject.bin")
    entries[rels_key] = _xml_bytes(rels_root)

    entries["xl/vbaProject.bin"] = vba_data

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in entries.items():
            out.writestr(name, data)


class MacroCopierApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title(f"VBA Macro Copier v{__version__}")
        self.minsize(640, 580)
        self.resizable(True, True)

        self._source_var = tk.StringVar()
        self._target_paths: list[str] = []

        icon = Path(__file__).parent / "icon.ico"
        if icon.exists():
            self.iconbitmap(str(icon))

        self._build_ui()

        src = find_source_file()
        if src:
            self._source_var.set(str(src))

    def _build_ui(self):
        P = 16  # base padding

        # ── Status bar (packed first so it anchors to bottom) ────
        ttk.Separator(self, orient=HORIZONTAL).pack(side=BOTTOM, fill=X)
        status_bar = ttk.Frame(self)
        status_bar.pack(side=BOTTOM, fill=X, padx=P, pady=5)
        ttk.Label(
            status_bar,
            text="Bo Sundgaard 2026  ·  www.uniteapps.dk",
            font=("Segoe UI", 8),
            foreground="#999999",
        ).pack(side=RIGHT)

        # ── Main content ─────────────────────────────────────────
        content = ttk.Frame(self)
        content.pack(fill=BOTH, expand=True, padx=P, pady=P)

        # Section: Source file
        ttk.Label(content, text="Source file", font=("Segoe UI", 9, "bold")).pack(anchor=W)
        src_row = ttk.Frame(content)
        src_row.pack(fill=X, pady=(4, 0))
        ttk.Entry(src_row, textvariable=self._source_var, state="readonly").pack(
            side=LEFT, fill=X, expand=True, padx=(0, 6)
        )
        ttk.Button(src_row, text="Browse…", command=self._browse_source, width=10).pack(side=LEFT)

        ttk.Separator(content, orient=HORIZONTAL).pack(fill=X, pady=(14, 12))

        # Section: Target files
        ttk.Label(content, text="Target files (.xlsx)", font=("Segoe UI", 9, "bold")).pack(anchor=W)

        list_border = ttk.Frame(content, relief="solid", borderwidth=1)
        list_border.pack(fill=BOTH, expand=True, pady=(4, 0))
        list_sb = ttk.Scrollbar(list_border)
        list_sb.pack(side=RIGHT, fill=Y)
        self._listbox = tk.Listbox(
            list_border,
            yscrollcommand=list_sb.set,
            selectmode=tk.EXTENDED,
            font=("Segoe UI", 10),
            activestyle="none",
            relief="flat",
            bd=0,
            highlightthickness=0,
        )
        self._listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=2, pady=2)
        list_sb.config(command=self._listbox.yview)

        tgt_btn_row = ttk.Frame(content)
        tgt_btn_row.pack(fill=X, pady=(6, 0))
        ttk.Button(tgt_btn_row, text="Add files…", command=self._browse_targets).pack(side=LEFT, padx=(0, 4))
        ttk.Button(tgt_btn_row, text="Remove selected", command=self._remove_selected).pack(side=LEFT, padx=(0, 4))
        ttk.Button(tgt_btn_row, text="Clear list", command=self._clear_list, bootstyle="secondary-outline").pack(side=LEFT)

        ttk.Separator(content, orient=HORIZONTAL).pack(fill=X, pady=(14, 12))

        # Section: Log
        ttk.Label(content, text="Output log", font=("Segoe UI", 9, "bold")).pack(anchor=W)

        log_border = ttk.Frame(content, relief="solid", borderwidth=1)
        log_border.pack(fill=BOTH, expand=True, pady=(4, 0))
        log_sb = ttk.Scrollbar(log_border)
        log_sb.pack(side=RIGHT, fill=Y)
        self._log = tk.Text(
            log_border,
            height=7,
            font=("Consolas", 9),
            state="disabled",
            relief="flat",
            bd=0,
            highlightthickness=0,
            yscrollcommand=log_sb.set,
        )
        self._log.pack(fill=BOTH, expand=True, padx=4, pady=4)
        log_sb.config(command=self._log.yview)
        self._log.tag_config("ok",   foreground="#107C10")  # Microsoft green
        self._log.tag_config("err",  foreground="#A4262C")  # Microsoft red
        self._log.tag_config("info", foreground="#0078D4")  # Microsoft blue

        ttk.Separator(content, orient=HORIZONTAL).pack(fill=X, pady=(14, 12))

        # Action row — primary button right-aligned like MS dialogs
        action_row = ttk.Frame(content)
        action_row.pack(fill=X)
        self._copy_btn = ttk.Button(
            action_row,
            text="Copy Macros",
            command=self._run_copy,
            bootstyle="primary",
            width=18,
        )
        self._copy_btn.pack(side=RIGHT)

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select source file",
            filetypes=[("Excel macro file", "*.xlsm")],
        )
        if path:
            self._source_var.set(path)

    def _browse_targets(self):
        paths = filedialog.askopenfilenames(
            title="Select Excel files (hold Ctrl to select multiple)",
            filetypes=[("Excel file", "*.xlsx")],
        )
        for p in paths:
            if p not in self._target_paths:
                self._target_paths.append(p)
                self._listbox.insert(tk.END, Path(p).name)

    def _remove_selected(self):
        for i in reversed(self._listbox.curselection()):
            self._listbox.delete(i)
            self._target_paths.pop(i)

    def _clear_list(self):
        self._listbox.delete(0, tk.END)
        self._target_paths.clear()

    def _log_line(self, text, tag=None):
        self._log.config(state="normal")
        self._log.insert(tk.END, text + "\n", tag or "")
        self._log.see(tk.END)
        self._log.config(state="disabled")

    def _run_copy(self):
        source = self._source_var.get().strip()
        if not source:
            messagebox.showwarning("Missing source file", "Please select a source.xlsm file.")
            return
        if not self._target_paths:
            messagebox.showwarning("Missing target files", "Please add at least one .xlsx file.")
            return

        source_path = Path(source)

        self._log.config(state="normal")
        self._log.delete("1.0", tk.END)
        self._log.config(state="disabled")

        self._copy_btn.config(state="disabled")
        self.update()

        ok = err = 0
        for path_str in self._target_paths:
            target = Path(path_str)
            output = target.parent / f"{target.stem}_new.xlsm"
            try:
                inject_macros(source_path, target, output)
                self._log_line(f"✓ {target.name}  →  {output.name}", "ok")
                ok += 1
            except Exception as e:
                self._log_line(f"✗ {target.name}  –  {e}", "err")
                err += 1

        summary = f"\nDone: {ok} succeeded"
        if err:
            summary += f", {err} failed"
        self._log_line(summary, "info")
        self._copy_btn.config(state="normal")


if __name__ == "__main__":
    app = MacroCopierApp()
    app.mainloop()
