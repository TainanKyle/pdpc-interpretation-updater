# -*- coding: utf-8 -*-
"""
PDPC administrative interpretation updater -- desktop window (thin shell).

This is only the GUI shell around core.py. It lets a non-technical user:
  1. pick the master Excel file,
  2. preview the detected columns,
  3. run a dry-run (compare only) or a real update (write back, with backup).

All scraping/comparison logic lives in core.py and is reused unchanged by the
future Colab notebook. Keeping this file thin is what makes that swap easy.

User-facing labels are in Chinese for the (Chinese-speaking) end users; code
comments and log lines are in English.
"""

import os
import queue
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, scrolledtext, ttk

import core

APP_TITLE = "個資法行政函釋更新工具"
# Bump on every delivered build -- the footer is the only way a user (or you,
# from a screenshot) can tell which build they are actually running.
#   1.0  first release
#   1.1  company-network TLS support (OS certificate store); .xlsx only;
#        atomic write-back; scraped text can no longer become an Excel formula
APP_VERSION = "1.1"
FIELD_LABELS = {           # internal field name -> Chinese label for the preview
    "article": "條號欄",
    "funhao": "函釋欄",
}


class App:
    def __init__(self, root):
        self.root = root
        self.excel_path = None
        self.log_queue = queue.Queue()
        self.running = False

        root.title(APP_TITLE)
        root.geometry("640x520")
        root.minsize(560, 460)

        # UI fonts: buttons + labels enlarged to ~13; emphasis (action buttons
        # and the column-detection status) bold, plain text (the path) regular.
        # The scrolling log keeps its own monospace font and is left untouched.
        family = tkfont.nametofont("TkDefaultFont").actual("family")
        self.ui_font = (family, 13)
        self.ui_font_bold = (family, 13, "bold")
        style = ttk.Style()
        style.configure("TButton", font=self.ui_font_bold)

        pad = {"padx": 12, "pady": 6}

        # --- file picker row ---
        top = ttk.Frame(root)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="選擇 Excel 檔案…", command=self.choose_file).pack(side="left")
        self.path_var = tk.StringVar(value="尚未選擇檔案")
        ttk.Label(top, textvariable=self.path_var, foreground="#444",
                  font=self.ui_font).pack(side="left", padx=10)

        # --- detected columns preview ---
        self.cols_var = tk.StringVar(value="")
        ttk.Label(root, textvariable=self.cols_var, foreground="#0a6",
                  font=self.ui_font_bold).pack(anchor="w", padx=12)

        # --- action buttons ---
        btns = ttk.Frame(root)
        btns.pack(fill="x", **pad)
        self.dry_btn = ttk.Button(btns, text="試跑（不更新檔案）",
                                  command=lambda: self.start(dry_run=True))
        self.dry_btn.pack(side="left")
        self.run_btn = ttk.Button(btns, text="開始更新",
                                  command=lambda: self.start(dry_run=False))
        self.run_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(btns, mode="indeterminate", length=160)
        self.progress.pack(side="right")

        # --- log area ---
        self.log = scrolledtext.ScrolledText(root, height=18,
                                             wrap="word", font=("Menlo", 11))
        # Signature footer pinned to the bottom. Packed BEFORE the log so the
        # log's expand fills the space above it instead of pushing it offscreen.
        ttk.Label(root, text=f"Kyle Huang · v{APP_VERSION} · 2026",
                  foreground="#999", font=(family, 10)).pack(side="bottom",
                                                             pady=(0, 6))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        # A "bold" tag for report section titles (e.g. 新增清單), so they stand
        # out from the body text in the same monospace log.
        self.log.tag_configure(
            "bold", font=tkfont.Font(family="Menlo", size=11, weight="bold"))
        # Read-only but still selectable/copyable: swallow plain typing, but let
        # any modifier combo (Cmd/Ctrl+C copy, Cmd/Ctrl+A select-all) and the
        # navigation keys through to the widget's native handlers — native copy
        # is more reliable in the bundled .app than a hand-rolled one.
        self.log.bind("<Key>", self._block_edit)
        self._build_menu()

        self._set_buttons(enabled=False)   # disabled until a file is picked
        self.root.after(100, self._drain_log)

    # ------------------------------------------------------------------ UI
    def _log(self, msg, bold=False):
        """Thread-safe: push a line to be drained onto the text widget.
        bold=True renders the line with the bold tag (used for report titles)."""
        self.log_queue.put((msg, bold))

    def _build_menu(self):
        """A real Edit menu carrying the Cmd+C / Cmd+A accelerators. On macOS a
        bundled .app only receives those shortcuts when a menu item claims them,
        so without this the log's native copy never fires no matter the bindings.
        The commands act on the log (the only selectable content)."""
        mod = "Cmd" if sys.platform == "darwin" else "Ctrl"
        menubar = tk.Menu(self.root)
        edit = tk.Menu(menubar, tearoff=0)
        edit.add_command(label="複製", accelerator=f"{mod}+C",
                         command=self._menu_copy)
        edit.add_command(label="全選", accelerator=f"{mod}+A",
                         command=self._menu_select_all)
        menubar.add_cascade(label="編輯", menu=edit)
        self.root.config(menu=menubar)

    def _menu_copy(self):
        try:
            text = self.log.get("sel.first", "sel.last")
        except tk.TclError:
            return                      # nothing selected
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def _menu_select_all(self):
        self.log.tag_add("sel", "1.0", "end-1c")

    def _block_edit(self, event):
        """Make the log read-only without disabling it (disabled = no copy).
        Let any modifier combo (Cmd on macOS = state 0x8/0x10, Ctrl = 0x4)
        through — that covers Cmd/Ctrl+C copy and Cmd/Ctrl+A select-all, handled
        natively — plus the navigation keys; swallow plain typing."""
        if event.state & 0x18 or (event.state & 0x4):
            return None
        if event.keysym in ("Left", "Right", "Up", "Down",
                             "Home", "End", "Prior", "Next"):
            return None
        return "break"

    def _drain_log(self):
        try:
            while True:
                msg, bold = self.log_queue.get_nowait()
                self.log.insert("end", msg + "\n", ("bold",) if bold else ())
                self.log.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    def _set_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.dry_btn.configure(state=state)
        self.run_btn.configure(state=state)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="選擇主表 Excel",
            # .xlsx only: openpyxl rebuilds the workbook on save and would drop
            # macros / charts / images from any other format (see core.sync_excel).
            filetypes=[("Excel 活頁簿 (.xlsx)", "*.xlsx")],
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            messagebox.showerror(
                "不支援的檔案格式",
                "只支援 .xlsx 主表。\n\n"
                "本工具寫回時會重建整個活頁簿，巨集（.xlsm）、圖表、圖片等\n"
                "會在存檔時遺失，因此不開放寫入。\n"
                "請先在 Excel 另存成 .xlsx，再用這個工具更新。",
            )
            return
        self.excel_path = path
        self.path_var.set(path)
        self._preview_columns()

    def _preview_columns(self):
        """Load the workbook header and show the detected column mapping."""
        try:
            wb, _ = core.load_workbook_from(self.excel_path)
            cols = core._detect_columns(wb.active)
        except Exception as e:  # noqa: BLE001
            self.cols_var.set(f"無法讀取檔案：{e}")
            self._set_buttons(enabled=False)
            return

        if "article" not in cols or "funhao" not in cols:
            self.cols_var.set(
                "⚠ 找不到「條號」與「最新函釋」欄，請確認第一列標題。"
                f"（偵測到：{cols or '無'}）")
            self._set_buttons(enabled=False)
            return

        desc = "、".join(
            f"{FIELD_LABELS[f]}＝第 {cols[f]} 欄" for f in ("article", "funhao"))
        self.cols_var.set(f"✓ 已辨識欄位：{desc}")
        self._set_buttons(enabled=True)

        # Heads-up if the file is already open: "開始更新" needs to write back, so
        # ask the user to close it. (Dry-run is read-only and stays available.)
        if core.is_excel_locked(self.excel_path):
            self.cols_var.set(f"✓ 已辨識欄位：{desc}　⚠ 檔案使用中，更新前請先關閉")
            messagebox.showwarning(
                "Excel 檔案開啟中",
                "偵測到這個 Excel 檔目前是開啟狀態。\n\n"
                "「試跑」不受影響，但「開始更新」需要寫回檔案，\n"
                "請先在 Excel 關閉它，再按「開始更新」。",
            )

    # -------------------------------------------------------------- worker
    def start(self, dry_run):
        if self.running or not self.excel_path:
            return
        # Re-check right before a real update: catches a file that was opened
        # after selection, or a warning the user clicked past. Dry-run is safe.
        if not dry_run and core.is_excel_locked(self.excel_path):
            messagebox.showwarning(
                "請先關閉 Excel 檔",
                "這個 Excel 檔目前是開啟狀態，無法寫回。\n\n"
                "請先在 Excel 關閉它，再按「開始更新」。",
            )
            return
        self.running = True
        self._set_buttons(enabled=False)
        self.progress.start(12)
        mode = "試跑" if dry_run else "更新"
        self._log("")
        self._log(f"===== 開始{mode} =====")
        threading.Thread(target=self._worker, args=(dry_run,), daemon=True).start()

    def _worker(self, dry_run):
        try:
            rows = core.scrape_all(progress=self._log)
            result = core.sync_excel(self.excel_path, rows,
                                     dry_run=dry_run, progress=self._log)
            self.root.after(0, lambda: self._done(result, dry_run))
        except Exception as e:  # noqa: BLE001
            # Bind e as a default arg: `except ... as e` deletes the name when
            # the block exits, so a plain closure would NameError when the
            # scheduled lambda later runs on the main thread (error dialog
            # never shown, buttons stuck disabled).
            self.root.after(0, lambda err=e: self._error(err))

    def _done(self, result, dry_run):
        self.progress.stop()
        self.running = False
        self._set_buttons(enabled=True)

        # Write the FULL detail (incl. the per-row list) to the scrolling log,
        # which can hold any number of rows without going off-screen.
        # A blank line + header separates the live progress text above from the
        # result report below, so the user can tell them apart.
        self._log("")
        self._log("===== 完成試跑 =====" if dry_run else "===== 完成更新 =====",
                  bold=True)
        self._log(f"網站函釋總數：{result['scraped_count']} 筆", bold=True)
        self._log(f"主表既有：{result['existing_count']} 筆", bold=True)
        self._log(f"本次新增：{result['new_count']} 筆", bold=True)
        if result.get("funhao_total") is not None:
            self._log(f"更新後函釋總數：共 {result['funhao_total']} 條", bold=True)
        if result["new_articles"]:
            self._log(f"新增條文列：{len(result['new_articles'])} 條"
                      f"（{ '、'.join(result['new_articles']) }）", bold=True)
        if result["new_count"]:
            self._log("新增清單：", bold=True)
            for r in result["new_rows"]:
                self._log(f"  [{r['article']}] {core.funhao_line(r)}")
                url = (r.get("article_url") or r.get("detail_url") or "").strip()
                if url:
                    self._log(f"      連結：{url}")
        if dry_run:
            self._log("（試跑模式，未寫入檔案）")
        elif result["written"]:
            self._log("")
            self._log(f"已寫回：{self.excel_path}")
            self._log(f"原檔備份：{result['backup_path']}")

        # Keep the pop-up SHORT (counts only) so it never grows taller than the
        # screen; the full list stays in the log area below.
        summary = [
            f"網站函釋總數：{result['scraped_count']} 筆",
            f"主表既有：{result['existing_count']} 筆",
            f"本次新增：{result['new_count']} 筆",
        ]
        if result.get("funhao_total") is not None:
            summary.append(f"更新後函釋總數：共 {result['funhao_total']} 條")
        if result["new_articles"]:
            summary.append(f"新增條文列：{len(result['new_articles'])} 條")
        if dry_run:
            summary.append("\n（試跑模式，未寫入檔案）")
        elif result["written"]:
            summary.append(f"\n原檔已備份：\n{result['backup_path']}")
        if result["new_count"]:
            summary.append("\n完整新增清單請見視窗下方的紀錄區。")

        title = "試跑結果" if dry_run else "更新完成"
        if not dry_run and result["new_count"] == 0:
            title = "已是最新"
        messagebox.showinfo(title, "\n".join(summary))

    def _error(self, err):
        self.progress.stop()
        self.running = False
        self._set_buttons(enabled=True)
        self._log(f"!! 發生錯誤：{err}")
        messagebox.showerror("發生錯誤", str(err))

    def on_close(self):
        """
        Refuse to close while a run is in progress.

        The worker is a daemon thread, so closing the window kills it wherever
        it happens to be. The write itself is atomic (core.sync_excel swaps in a
        temp file), so the master sheet survives regardless -- but a half-run
        that silently reports nothing is still confusing, and the user rarely
        means it.
        """
        if self.running:
            if not messagebox.askyesno(
                "更新進行中",
                "目前正在執行中，現在關閉會中斷這次更新。\n\n"
                "（主表不會損壞，但這次的更新不會完成。）\n\n"
                "確定要關閉嗎？",
                default="no",
            ):
                return
        self.root.destroy()


def _bring_to_front(root):
    """
    Force the window to the foreground with focus.

    A PyInstaller --windowed tkinter app on macOS otherwise tends to open
    behind other windows without focus, so the user sees "nothing opened".
    """
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.focus_force()
    # Drop the always-on-top flag shortly after, so it behaves normally.
    root.after(600, lambda: root.attributes("-topmost", False))
    # macOS: activate the app process so its window comes to the front.
    if sys.platform == "darwin":
        try:
            from subprocess import run
            run([
                "osascript", "-e",
                'tell app "System Events" to set frontmost of '
                '(first process whose unix id is %d) to true' % os.getpid(),
            ], check=False)
        except Exception:
            pass


def main():
    root = tk.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.after(0, lambda: _bring_to_front(root))
    root.mainloop()


if __name__ == "__main__":
    main()
