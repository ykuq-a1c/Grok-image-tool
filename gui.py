import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

from config import load_settings, save_settings, load_state, save_state
from processor import collect_image_files, run_batch
from thumbnail_panel import ThumbnailPanel


class App(TkinterDnD.Tk if DND_AVAILABLE else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Grok Image Tool")
        self.resizable(False, False)

        self.settings = load_settings()
        self._image_files = []
        self._stop_flag = False
        self._running = False
        self._last_failed = []
        self._round_info = ""

        self._build_ui()
        self._wire_cross_tab_buttons()
        self._restore_settings()
        self._restore_state()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ UI構築

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        tab1 = ttk.Frame(self.notebook)
        tab2 = ttk.Frame(self.notebook)
        tab3 = ttk.Frame(self.notebook)
        self.notebook.add(tab1, text="  処理対象  ")
        self.notebook.add(tab2, text="  保留リスト  ")
        self.notebook.add(tab3, text="  エラーログ  ")

        self._build_tab1(tab1)
        self._build_tab2(tab2)
        self._build_tab3(tab3)

    # ---- Tab1 ----

    def _build_tab1(self, parent):
        pad = {"padx": 8, "pady": 4}

        self.thumb_panel = ThumbnailPanel(parent, count_label="処理対象ファイル", canvas_height=200)
        self.thumb_panel.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 2))

        if DND_AVAILABLE:
            self.thumb_panel.register_dnd(self._on_drop_target)

        frm_btn_in = tk.Frame(parent)
        frm_btn_in.grid(row=1, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 6))
        ttk.Button(frm_btn_in, text="フォルダを選ぶ",
                   command=lambda: self._select_folder_for(self.thumb_panel)).pack(side="left", padx=4)
        ttk.Button(frm_btn_in, text="ファイルを選ぶ",
                   command=lambda: self._select_files_for(self.thumb_panel)).pack(side="left", padx=4)

        frm_out = ttk.LabelFrame(parent, text="出力フォルダ")
        frm_out.grid(row=2, column=0, columnspan=2, sticky="ew", **pad)
        frm_out_row = tk.Frame(frm_out)
        frm_out_row.pack(fill="x", padx=6, pady=6)
        self.var_output_dir = tk.StringVar()
        ttk.Button(frm_out_row, text="フォルダを選ぶ",
                   command=self._select_output_folder).pack(side="right")
        ttk.Entry(frm_out_row, textvariable=self.var_output_dir).pack(
            side="left", fill="x", expand=True, padx=(0, 6))

        frm_prompt = ttk.LabelFrame(parent, text="プロンプト")
        frm_prompt.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.txt_prompt = tk.Text(frm_prompt, height=3, width=55, wrap="word")
        self.txt_prompt.pack(fill="x", padx=6, pady=6)

        frm_name = ttk.LabelFrame(parent, text="ファイル命名")
        frm_name.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
        self.var_naming = tk.StringVar(value="none")
        frm_name_row = tk.Frame(frm_name)
        frm_name_row.pack(fill="x", padx=6, pady=6)
        ttk.Radiobutton(frm_name_row, text="そのまま", variable=self.var_naming,
                        value="none", command=self._update_naming_ui).pack(side="left")
        ttk.Radiobutton(frm_name_row, text="頭に追加:", variable=self.var_naming,
                        value="prefix", command=self._update_naming_ui).pack(side="left", padx=(12, 2))
        ttk.Radiobutton(frm_name_row, text="末尾に追加:", variable=self.var_naming,
                        value="suffix", command=self._update_naming_ui).pack(side="left", padx=(12, 2))
        self.var_naming_text = tk.StringVar(value="edited_")
        self.ent_naming_text = ttk.Entry(frm_name_row, textvariable=self.var_naming_text, width=14)
        self.ent_naming_text.pack(side="left", padx=4)
        self._update_naming_ui()

        frm_ctrl = tk.Frame(parent)
        frm_ctrl.grid(row=5, column=0, columnspan=2, pady=6)
        self.btn_start = ttk.Button(frm_ctrl, text="開始", width=14, command=self._start)
        self.btn_start.pack(side="left", padx=8)
        self.btn_stop = ttk.Button(frm_ctrl, text="中断", width=14, command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=8)

        frm_prog = ttk.LabelFrame(parent, text="進捗")
        frm_prog.grid(row=6, column=0, columnspan=2, sticky="ew", **pad)
        self.progressbar = ttk.Progressbar(frm_prog, length=460, mode="determinate")
        self.progressbar.pack(padx=6, pady=(6, 2))
        self.lbl_progress = ttk.Label(frm_prog, text="待機中")
        self.lbl_progress.pack(anchor="w", padx=8, pady=(0, 6))

    # ---- Tab2 ----

    def _build_tab2(self, parent):
        frm_hold = ttk.LabelFrame(parent, text="保留リスト（処理対象外）")
        frm_hold.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(frm_hold, text="処理したくない画像を一時退避する場所です。",
                  foreground="#555").pack(anchor="w", padx=8, pady=(6, 2))

        self.hold_panel = ThumbnailPanel(frm_hold, count_label="保留ファイル", canvas_height=150)
        self.hold_panel.pack(fill="both", expand=True, padx=4, pady=(2, 2))

        if DND_AVAILABLE:
            self.hold_panel.register_dnd(self._on_drop_hold)

        frm_btn = tk.Frame(frm_hold)
        frm_btn.pack(fill="x", padx=6, pady=(2, 8))
        ttk.Button(frm_btn, text="フォルダを選ぶ",
                   command=lambda: self._select_folder_for(self.hold_panel)).pack(side="left", padx=4)
        ttk.Button(frm_btn, text="ファイルを選ぶ",
                   command=lambda: self._select_files_for(self.hold_panel)).pack(side="left", padx=4)

        # ---- 設定 ----
        frm_cfg = ttk.LabelFrame(parent, text="設定")
        frm_cfg.pack(fill="x", padx=8, pady=(4, 8))

        frm_provider = tk.Frame(frm_cfg)
        frm_provider.pack(fill="x", padx=6, pady=(8, 4))
        ttk.Label(frm_provider, text="使用API:").pack(side="left")
        self.var_provider = tk.StringVar(value="xai")
        ttk.Radiobutton(frm_provider, text="xAI (Grok Imagine)", variable=self.var_provider,
                        value="xai").pack(side="left", padx=(8, 0))
        ttk.Radiobutton(frm_provider, text="Venice AI", variable=self.var_provider,
                        value="venice").pack(side="left", padx=(12, 0))

        frm_xai_key = tk.Frame(frm_cfg)
        frm_xai_key.pack(fill="x", padx=6, pady=(2, 2))
        ttk.Label(frm_xai_key, text="xAI APIキー:     ", foreground="#333").pack(side="left")
        self.var_xai_key = tk.StringVar()
        ttk.Entry(frm_xai_key, textvariable=self.var_xai_key, show="*", width=36).pack(side="left")

        frm_venice_key = tk.Frame(frm_cfg)
        frm_venice_key.pack(fill="x", padx=6, pady=(2, 4))
        ttk.Label(frm_venice_key, text="Venice APIキー:", foreground="#333").pack(side="left")
        self.var_venice_key = tk.StringVar()
        ttk.Entry(frm_venice_key, textvariable=self.var_venice_key, show="*", width=36).pack(side="left")

        frm_save = tk.Frame(frm_cfg)
        frm_save.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(frm_save, text="保存", command=self._save_settings).pack(side="right")

        frm_cfg2 = tk.Frame(frm_cfg)
        frm_cfg2.pack(fill="x", padx=6, pady=(2, 8))
        ttk.Label(frm_cfg2, text="処理間隔(秒):").pack(side="left")
        self.var_interval = tk.StringVar(value="2.0")
        ttk.Entry(frm_cfg2, textvariable=self.var_interval, width=5).pack(side="left", padx=(4, 12))
        ttk.Label(frm_cfg2, text="同時処理数:").pack(side="left")
        self.var_workers = tk.StringVar(value="5")
        ttk.Entry(frm_cfg2, textvariable=self.var_workers, width=4).pack(side="left", padx=(4, 12))
        ttk.Label(frm_cfg2, text="自動リトライ:").pack(side="left")
        self.var_retry_count = tk.StringVar(value="0")
        ttk.Entry(frm_cfg2, textvariable=self.var_retry_count, width=3).pack(side="left", padx=4)
        ttk.Label(frm_cfg2, text="回", foreground="#666").pack(side="left")

    # ---- Tab3 ----

    def _build_tab3(self, parent):
        frm = ttk.LabelFrame(parent, text="最終エラーログ")
        frm.pack(fill="both", expand=True, padx=8, pady=8)

        ttk.Label(frm, text="処理完了後に失敗したファイルの一覧が表示されます。",
                  foreground="#555").pack(anchor="w", padx=8, pady=(6, 2))

        frm_text = tk.Frame(frm)
        frm_text.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        self.txt_error_log = tk.Text(
            frm_text, height=14, width=1, state="disabled", wrap="word",
            font=("Courier New", 9)
        )
        vscroll = ttk.Scrollbar(frm_text, orient="vertical", command=self.txt_error_log.yview)
        self.txt_error_log.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        self.txt_error_log.pack(side="left", fill="both", expand=True)

        frm_btn = tk.Frame(frm)
        frm_btn.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(frm_btn, text="エラーリストをCSVで保存",
                   command=self._save_error_csv).pack(side="left", padx=(0, 8))
        ttk.Button(frm_btn, text="失敗分を処理対象リストへ",
                   command=self._load_failed_to_target).pack(side="left")

    def _wire_cross_tab_buttons(self):
        self.thumb_panel.add_header_button("保留リストへ", self._move_to_hold)
        self.thumb_panel._pack_remove_btn()
        self.thumb_panel.set_preview_move_callback(
            "この画像を保留リストへ", self._preview_move_to_hold)
        self.hold_panel.add_header_button("対象リストへ", self._move_to_target)
        self.hold_panel._pack_remove_btn()
        self.hold_panel.set_preview_move_callback(
            "この画像を処理対象リストへ", self._preview_move_to_target)

    # ------------------------------------------------------------------ 設定

    def _restore_settings(self):
        s = self.settings
        self.var_provider.set(s.get("api_provider", "xai"))
        self.var_xai_key.set(s.get("xai_api_key", ""))
        self.var_venice_key.set(s.get("venice_api_key", ""))
        self.var_interval.set(str(s.get("interval_sec", 0.3)))
        self.var_workers.set(str(s.get("max_workers", 5)))
        self.var_retry_count.set(str(s.get("retry_count", 0)))
        self.var_naming.set(s.get("naming_mode", "none"))
        self.var_naming_text.set(s.get("naming_text", "edited_"))
        self.var_output_dir.set(s.get("last_output_dir", ""))
        if s.get("last_prompt"):
            self.txt_prompt.insert("1.0", s["last_prompt"])
        self._update_naming_ui()

    def _save_settings(self):
        self.settings.update(self._current_settings())
        save_settings(self.settings)
        messagebox.showinfo("保存", "設定を保存しました。")

    def _current_settings(self) -> dict:
        return {
            "api_provider": self.var_provider.get(),
            "xai_api_key": self.var_xai_key.get().strip(),
            "venice_api_key": self.var_venice_key.get().strip(),
            "interval_sec": self._parse_interval(),
            "max_workers": self._parse_workers(),
            "retry_count": self._parse_retry_count(),
            "naming_mode": self.var_naming.get(),
            "naming_text": self.var_naming_text.get(),
            "last_output_dir": self.var_output_dir.get().strip(),
            "last_prompt": self.txt_prompt.get("1.0", "end").strip(),
        }

    def _parse_interval(self) -> float:
        try:
            return max(0.1, float(self.var_interval.get()))
        except ValueError:
            return 0.3

    def _parse_workers(self) -> int:
        try:
            return max(1, min(int(self.var_workers.get()), 20))
        except ValueError:
            return 5

    def _parse_retry_count(self) -> int:
        try:
            return max(0, min(int(self.var_retry_count.get()), 10))
        except ValueError:
            return 0

    # ------------------------------------------------------------------ 入力選択

    def _on_drop_target(self, event):
        self.thumb_panel.add_files(self._parse_drop(event.data))

    def _on_drop_hold(self, event):
        self.hold_panel.add_files(self._parse_drop(event.data))

    def _parse_drop(self, raw: str) -> list:
        files = []
        for p in self.tk.splitlist(raw.strip()):
            if p.lower().endswith(".csv"):
                files.extend(self._load_files_from_csv(p))
            else:
                files.extend(collect_image_files(p))
        return files

    def _select_folder_for(self, panel: ThumbnailPanel):
        folder = filedialog.askdirectory(title="入力フォルダを選択")
        if folder:
            panel.add_files(collect_image_files(folder))

    def _select_files_for(self, panel: ThumbnailPanel):
        paths = filedialog.askopenfilenames(
            title="画像ファイルまたは失敗リスト(CSV)を選択",
            filetypes=[("画像・CSVファイル", "*.jpg *.jpeg *.png *.webp *.csv"),
                       ("すべて", "*.*")]
        )
        if not paths:
            return
        files = []
        for p in paths:
            if p.lower().endswith(".csv"):
                files.extend(self._load_files_from_csv(p))
            else:
                files.extend(collect_image_files(p))
        panel.add_files(files)

    def _load_files_from_csv(self, csv_path: str) -> list:
        files = []
        try:
            with open(csv_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            for line in lines[1:]:
                line = line.strip()
                if not line:
                    continue
                fp = line[1:line.find('"', 1)] if line.startswith('"') else line.split(",")[0]
                if os.path.isfile(fp):
                    files.append(fp)
        except Exception as e:
            messagebox.showerror("エラー", f"CSVの読み込みに失敗しました:\n{e}")
        return files

    def _select_output_folder(self):
        folder = filedialog.askdirectory(title="出力フォルダを選択")
        if folder:
            self.var_output_dir.set(folder)

    # ------------------------------------------------------------------ タブ間移動

    def _move_to_hold(self):
        files = self.thumb_panel.get_selected_files()
        if files:
            self.thumb_panel.remove_selected()
            self.hold_panel.add_files(files)

    def _move_to_target(self):
        files = self.hold_panel.get_selected_files()
        if files:
            self.hold_panel.remove_selected()
            self.thumb_panel.add_files(files)

    def _preview_move_to_hold(self, filepath: str):
        self.thumb_panel.remove_file(filepath)
        self.hold_panel.add_files([filepath])

    def _preview_move_to_target(self, filepath: str):
        self.hold_panel.remove_file(filepath)
        self.thumb_panel.add_files([filepath])

    # ------------------------------------------------------------------ 命名UI

    def _update_naming_ui(self):
        mode = self.var_naming.get()
        self.ent_naming_text.config(state="normal" if mode in ("prefix", "suffix") else "disabled")

    # ------------------------------------------------------------------ 処理開始・中断

    def _start(self):
        self._image_files = self.thumb_panel.get_files()
        if not self._image_files:
            messagebox.showwarning("確認", "処理対象の画像が選択されていません。")
            return
        if not self.var_output_dir.get().strip():
            messagebox.showwarning("確認", "出力フォルダを選択してください。")
            return
        provider = self.var_provider.get()
        active_key = self.var_xai_key.get().strip() if provider == "xai" else self.var_venice_key.get().strip()
        if not active_key:
            label = "xAI" if provider == "xai" else "Venice AI"
            messagebox.showwarning("確認", f"{label} のAPIキーを入力してください。（Tab2 設定で入力）")
            return
        prompt = self.txt_prompt.get("1.0", "end").strip()
        if not prompt:
            messagebox.showwarning("確認", "プロンプトを入力してください。")
            return

        self._stop_flag = False
        self._running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progressbar["maximum"] = len(self._image_files)
        self.progressbar["value"] = 0

        self.settings.update(self._current_settings())
        save_settings(self.settings)

        threading.Thread(target=self._run_worker, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self.btn_stop.config(state="disabled")
        self.lbl_progress.config(text="中断中...")

    def _run_worker(self):
        max_retries = self._parse_retry_count()
        current_files = list(self._image_files)
        last_failed = []
        provider   = self.var_provider.get()
        api_key    = self.var_xai_key.get().strip() if provider == "xai" else self.var_venice_key.get().strip()
        prompt     = self.txt_prompt.get("1.0", "end").strip()
        output_dir = self.var_output_dir.get().strip()
        naming_mode = self.var_naming.get()
        naming_text = self.var_naming_text.get()

        for attempt in range(max_retries + 1):
            if not current_files or self._stop_flag:
                break

            self._round_info = "" if attempt == 0 else f"[再試行 {attempt}/{max_retries}] "
            self.after(0, self._reset_progress, len(current_files))

            last_failed = run_batch(
                api_key=api_key,
                image_files=current_files,
                prompt=prompt,
                output_dir=output_dir,
                naming_mode=naming_mode,
                naming_text=naming_text,
                interval_sec=self._parse_interval(),
                max_workers=self._parse_workers(),
                on_progress=self._on_progress,
                stop_flag=lambda: self._stop_flag,
                api_provider=provider,
            )

            if not last_failed:
                break
            current_files = [fp for fp, _, _ in last_failed]

        self.after(0, self._on_done, last_failed)

    def _reset_progress(self, total: int):
        self.progressbar["maximum"] = total
        self.progressbar["value"] = 0

    def _on_progress(self, done, total, success, fail, skip):
        self.after(0, self._update_progress_ui, done, total, success, fail, skip)

    def _update_progress_ui(self, done, total, success, fail, skip):
        self.progressbar["value"] = done
        retry_info = getattr(self, "_round_info", "")
        text = f"{retry_info}{done} / {total}  ✓成功: {success}  ✗失敗: {fail}  →スキップ: {skip}"
        self.lbl_progress.config(text=text)

    def _on_done(self, failed: list):
        self._running = False
        self._last_failed = failed
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.progressbar["value"] = self.progressbar["maximum"]

        self._update_error_log(failed)

        if self._stop_flag:
            msg = "処理を中断しました。\n同じ設定で「開始」すると続きから再開できます。"
            messagebox.showinfo("中断", msg)
        elif not failed:
            messagebox.showinfo("完了", "処理が完了しました。エラーはありませんでした。")
        else:
            self._show_done_with_failures(failed)

    def _show_done_with_failures(self, failed: list):
        dlg = tk.Toplevel(self)
        dlg.title("完了")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

        msg = f"処理が完了しました。\n失敗: {len(failed)} 件 — エラーログを確認してください。"
        ttk.Label(dlg, text=msg, padding=(24, 16)).pack()

        frm_btn = tk.Frame(dlg)
        frm_btn.pack(pady=(0, 16))

        ttk.Button(frm_btn, text="OK", width=10,
                   command=dlg.destroy).pack(side="left", padx=8)
        ttk.Button(frm_btn, text="失敗分だけをリストに再登録",
                   command=lambda: self._reload_failed_and_close(failed, dlg)).pack(side="left", padx=8)

        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width()  - dlg.winfo_width())  // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")

        self.wait_window(dlg)

    def _on_close(self):
        self._save_state()
        self.destroy()

    def _save_state(self):
        save_state({
            "target_files": self.thumb_panel.get_files(),
            "hold_files":   self.hold_panel.get_files(),
            "last_failed":  self._last_failed,
        })

    def _restore_state(self):
        state = load_state()
        target = state.get("target_files", [])
        hold   = state.get("hold_files", [])
        failed = state.get("last_failed", [])
        if target:
            existing = [fp for fp in target if os.path.isfile(fp)]
            if existing:
                self.thumb_panel.set_files(existing)
        if hold:
            existing = [fp for fp in hold if os.path.isfile(fp)]
            if existing:
                self.hold_panel.set_files(existing)
        if failed:
            self._last_failed = [tuple(item) for item in failed]
            self._update_error_log(self._last_failed)

    def _reload_failed_and_close(self, failed: list, dlg):
        dlg.destroy()
        files = [fp for fp, _, _ in failed]
        self.thumb_panel.set_files(files)
        self.notebook.select(0)

    # ------------------------------------------------------------------ エラーログ

    def _update_error_log(self, failed: list):
        self.txt_error_log.config(state="normal")
        self.txt_error_log.delete("1.0", "end")
        if not failed:
            self.txt_error_log.insert("end", "エラーなし — すべて正常に処理されました。\n")
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.txt_error_log.insert("end", f"処理完了: {now}  /  失敗: {len(failed)} 件\n")
            self.txt_error_log.insert("end", "─" * 70 + "\n\n")
            for fp, error_type, error_msg in failed:
                self.txt_error_log.insert("end", f"ファイル : {fp}\n")
                self.txt_error_log.insert("end", f"  種別   : {error_type}\n")
                self.txt_error_log.insert("end", f"  詳細   : {error_msg}\n\n")
        self.txt_error_log.config(state="disabled")

    def _save_error_csv(self):
        if not self._last_failed:
            messagebox.showinfo("確認", "保存するエラーがありません。")
            return
        path = filedialog.asksaveasfilename(
            title="エラーリストを保存",
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべて", "*.*")],
            initialfile=f"failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("filepath,error_type,error_message\n")
            for fp, et, em in self._last_failed:
                f.write(f'"{fp}",{et},"{em.replace(chr(34), chr(39))}"\n')
        messagebox.showinfo("保存完了", f"保存しました:\n{path}")

    def _load_failed_to_target(self):
        if not self._last_failed:
            messagebox.showinfo("確認", "失敗ファイルがありません。")
            return
        files = [fp for fp, _, _ in self._last_failed]
        self.thumb_panel.set_files(files)
        self.notebook.select(0)
        messagebox.showinfo("読み込み完了", f"{len(files)} 件を処理対象リストに読み込みました。")
