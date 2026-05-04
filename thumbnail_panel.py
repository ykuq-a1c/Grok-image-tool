import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from queue import Queue, Empty

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

THUMB = 88
CELL  = 104
MAX_DISPLAY = 500


class ThumbnailPanel(tk.Frame):
    def __init__(self, master, count_label="対象ファイル", on_count_change=None, canvas_height=230, **kwargs):
        super().__init__(master, **kwargs)
        self._count_label = count_label
        self._on_count_change = on_count_change
        self._canvas_height = canvas_height
        self._files = []
        self._display = []
        self._selected = set()
        self._photos = {}
        self._pending = Queue()
        self._redraw_scheduled = False
        # ドラッグ選択の状態
        self._drag_start = None
        self._drag_rect = None
        self._is_dragging = False
        self._preview_move_label = None
        self._preview_move_cb = None
        self._build()
        threading.Thread(target=self._loader, daemon=True).start()

    # ------------------------------------------------------------------ 構築

    def _build(self):
        self._frm_head = tk.Frame(self)
        self._frm_head.pack(fill="x", padx=4, pady=(4, 2))

        self.lbl_count = ttk.Label(self._frm_head, text=f"{self._count_label}: 0 枚")
        self.lbl_count.pack(side="left")

        self.btn_select_all = ttk.Button(
            self._frm_head, text="全選択",
            command=self._select_all_or_none, state="disabled"
        )
        self.btn_select_all.pack(side="left", padx=(8, 0))

        if not PIL_AVAILABLE:
            ttk.Label(self._frm_head, text="※pip install Pillow でサムネイル表示",
                      foreground="#c00").pack(side="left", padx=8)

        # 右端のボタングループ（左から: クロスタブ / 消去）
        self._frm_right = tk.Frame(self._frm_head)
        self._frm_right.pack(side="right")

        # btn_remove は _pack_remove_btn() で最後に追加される
        self.btn_remove = ttk.Button(
            self._frm_right, text="リストから消去",
            command=self.remove_selected, state="disabled"
        )

        frm_body = tk.Frame(self)
        frm_body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            frm_body, bg="#dcdcdc", height=self._canvas_height,
            highlightthickness=1, highlightbackground="#bbbbbb"
        )
        vscroll = ttk.Scrollbar(frm_body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.bind("<Configure>", lambda e: self._schedule_redraw())
        self.canvas.bind("<ButtonPress-1>",   self._on_press)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<B1-Motion>",       self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Button-3>",        self._on_right_click)
        self.canvas.bind("<MouseWheel>",
            lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self.canvas.bind("<Delete>", lambda e: self.remove_selected() if self._selected else None)

        self._draw_placeholder()

    # ------------------------------------------------------------------ 公開 API

    def set_preview_move_callback(self, label: str, callback):
        """プレビュー右クリックメニューの「他リストへ移動」アクションを設定する"""
        self._preview_move_label = label
        self._preview_move_cb = callback

    def register_dnd(self, callback):
        try:
            from tkinterdnd2 import DND_FILES
            self.canvas.drop_target_register(DND_FILES)
            self.canvas.dnd_bind("<<Drop>>", callback)
        except Exception:
            pass

    def add_header_button(self, label: str, command) -> ttk.Button:
        """ヘッダー右グループにボタンを追加する（全選択の右、除外の左）"""
        btn = ttk.Button(self._frm_right, text=label, command=command, state="disabled")
        btn.pack(side="left", padx=(4, 0))
        return btn

    def _pack_remove_btn(self):
        """クロスタブボタン追加後に呼ぶ。除外ボタンをグループ右端に配置する。"""
        self.btn_remove.pack(side="left", padx=(4, 0))

    def set_files(self, files: list):
        """ファイルリストを置き換える"""
        self._files = list(files)
        self._display = self._files[:MAX_DISPLAY]
        self._selected.clear()
        self._enqueue(self._display)
        self._update_header()
        self._schedule_redraw()

    def add_files(self, files: list):
        """重複しないようにファイルを追記する"""
        existing = set(self._files)
        new = [f for f in files if f not in existing]
        if not new:
            return
        self._files.extend(new)
        self._display = self._files[:MAX_DISPLAY]
        self._enqueue(new)
        self._update_header()
        self._schedule_redraw()

    def get_files(self) -> list:
        return list(self._files)

    def get_selected_files(self) -> list:
        return [self._display[i] for i in sorted(self._selected)
                if i < len(self._display)]

    def _select_all_or_none(self):
        if self._display and len(self._selected) >= len(self._display):
            self._selected.clear()
        else:
            self._selected = set(range(len(self._display)))
        self._update_header()
        self._schedule_redraw()

    def remove_selected(self):
        excluded = {self._display[i] for i in self._selected if i < len(self._display)}
        self._files = [fp for fp in self._files if fp not in excluded]
        self._display = self._files[:MAX_DISPLAY]
        self._selected.clear()
        self._update_header()
        self._schedule_redraw()

    def remove_file(self, filepath: str):
        """1ファイルをリストから削除する（プレビュー操作用）"""
        if filepath not in self._files:
            return
        self._files.remove(filepath)
        self._display = self._files[:MAX_DISPLAY]
        self._selected.clear()
        self._update_header()
        self._schedule_redraw()

    def _enqueue(self, files):
        for fp in files:
            if fp not in self._photos:
                self._pending.put(fp)

    # ------------------------------------------------------------------ ヘッダー

    def _update_header(self):
        n = len(self._files)
        text = f"{self._count_label}: {n} 枚"
        if n > MAX_DISPLAY:
            text += f"  （先頭 {MAX_DISPLAY} 枚を表示）"
        self.lbl_count.config(text=text)

        has_files = bool(self._display)
        has_sel = bool(self._selected)
        all_sel = has_files and len(self._selected) >= len(self._display)

        self.btn_select_all.config(
            text="全解除" if all_sel else "全選択",
            state="normal" if has_files else "disabled"
        )
        sel_state = "normal" if has_sel else "disabled"
        for w in self._frm_right.winfo_children():
            if isinstance(w, ttk.Button) and w is not self.btn_select_all:
                w.config(state=sel_state)

        if self._on_count_change:
            self._on_count_change(n)

    # ------------------------------------------------------------------ 描画

    def _schedule_redraw(self):
        if not self._redraw_scheduled:
            self._redraw_scheduled = True
            self.after(60, self._do_redraw)

    def _do_redraw(self):
        self._redraw_scheduled = False
        self._redraw()

    def _draw_placeholder(self):
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 460
        h = self.canvas.winfo_height() or 230
        self.canvas.create_text(
            w // 2, h // 2,
            text="ここにフォルダ・画像ファイル・エラーリスト(CSV)をドロップ\nまたは下のボタンで選択",
            justify="center", fill="#888888", font=("", 10)
        )
        self.canvas.configure(scrollregion=(0, 0, w, h))

    def _redraw(self):
        self.canvas.delete("all")
        if not self._display:
            self._draw_placeholder()
            return

        cols = self._cols()
        for i, fp in enumerate(self._display):
            row, col = divmod(i, cols)
            cx = col * CELL + CELL // 2
            cy = row * CELL + CELL // 2
            x0, y0 = cx - THUMB // 2, cy - THUMB // 2

            photo = self._photos.get(fp)
            if photo:
                self.canvas.create_image(cx, cy, image=photo, anchor="center", tags=f"c{i}")
            else:
                self.canvas.create_rectangle(
                    x0, y0, x0 + THUMB, y0 + THUMB,
                    fill="#b8b8b8", outline="#999", tags=f"c{i}"
                )
                name = os.path.basename(fp)
                self.canvas.create_text(
                    cx, cy,
                    text=(name[:10] + "…" if len(name) > 12 else name),
                    fill="#444", font=("", 7), tags=f"c{i}"
                )

            if i in self._selected:
                self.canvas.create_rectangle(
                    x0 - 1, y0 - 1, x0 + THUMB + 1, y0 + THUMB + 1,
                    outline="#dd0000", width=3, tags=f"s{i}"
                )

        if len(self._files) > MAX_DISPLAY:
            rows = (len(self._display) + cols - 1) // cols
            self.canvas.create_text(
                (self.canvas.winfo_width() or 460) // 2, rows * CELL + 16,
                text=f"他 {len(self._files) - MAX_DISPLAY} 枚（表示省略）",
                fill="#666", font=("", 9)
            )

        self._update_scrollregion()

    def _update_scrollregion(self):
        cols = self._cols()
        n = len(self._display)
        rows = (n + cols - 1) // cols if n else 1
        extra = 32 if len(self._files) > MAX_DISPLAY else 0
        w = self.canvas.winfo_width() or 460
        h = max(rows * CELL + extra, self.canvas.winfo_height() or 230)
        self.canvas.configure(scrollregion=(0, 0, w, h))

    def _cols(self) -> int:
        w = self.canvas.winfo_width() or 460
        return max(1, w // CELL)

    # ------------------------------------------------------------------ マウス操作

    def _on_press(self, event):
        self.canvas.focus_set()
        self._drag_start = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self._is_dragging = False

    def _on_drag(self, event):
        if self._drag_start is None:
            return
        sx, sy = self._drag_start
        ex = self.canvas.canvasx(event.x)
        ey = self.canvas.canvasy(event.y)
        if not self._is_dragging and (abs(ex - sx) > 5 or abs(ey - sy) > 5):
            self._is_dragging = True
        if self._is_dragging:
            if self._drag_rect:
                self.canvas.delete(self._drag_rect)
            self._drag_rect = self.canvas.create_rectangle(
                sx, sy, ex, ey,
                outline="#0078d7", width=1, dash=(4, 2), fill=""
            )

    def _on_release(self, event):
        ex = self.canvas.canvasx(event.x)
        ey = self.canvas.canvasy(event.y)

        if self._drag_rect:
            self.canvas.delete(self._drag_rect)
            self._drag_rect = None

        if self._is_dragging and self._drag_start:
            # 範囲内のインデックスを収集
            sx, sy = self._drag_start
            x0, x1 = min(sx, ex), max(sx, ex)
            y0, y1 = min(sy, ey), max(sy, ey)
            cols = self._cols()
            in_range = set()
            for i in range(len(self._display)):
                row, col = divmod(i, cols)
                cx = col * CELL + CELL // 2
                cy = row * CELL + CELL // 2
                if x0 <= cx <= x1 and y0 <= cy <= y1:
                    in_range.add(i)
            # 範囲内が全て選択済みなら解除、そうでなければ選択
            if in_range and in_range.issubset(self._selected):
                self._selected -= in_range
            else:
                self._selected |= in_range
            self._update_header()
            self._schedule_redraw()
        else:
            # 通常クリック
            self._click_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

        self._drag_start = None
        self._is_dragging = False

    def _click_at(self, cx: float, cy: float):
        cols = self._cols()
        idx = int(cy // CELL) * cols + int(cx // CELL)
        if 0 <= idx < len(self._display):
            if idx in self._selected:
                self._selected.discard(idx)
            else:
                self._selected.add(idx)
            self._update_header()
            self._schedule_redraw()

    def _on_double_click(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        cols = self._cols()
        idx = int(cy // CELL) * cols + int(cx // CELL)
        if 0 <= idx < len(self._display):
            self._open_preview(idx)

    def _on_right_click(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        cols = self._cols()
        idx = int(cy // CELL) * cols + int(cx // CELL)
        if 0 <= idx < len(self._display):
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="プレビューを表示",
                             command=lambda i=idx: self._open_preview(i))
            menu.tk_popup(event.x_root, event.y_root)

    def _open_preview(self, start_idx: int):
        if not PIL_AVAILABLE:
            messagebox.showinfo("確認", "プレビューには Pillow が必要です。\npip install Pillow")
            return
        from PIL import Image, ImageTk

        win = tk.Toplevel(self)
        win.resizable(True, True)

        current = [start_idx]

        preview_canvas = tk.Canvas(win, bg="#1a1a1a", highlightthickness=0)
        preview_canvas.pack(fill="both", expand=True)
        lbl_info = ttk.Label(win, anchor="w")
        lbl_info.pack(fill="x", padx=8, pady=4)

        def load(idx):
            if idx < 0 or idx >= len(self._display):
                return
            current[0] = idx
            fp = self._display[idx]
            win.title(os.path.basename(fp))
            try:
                img = Image.open(fp)
            except Exception as e:
                messagebox.showerror("エラー", f"画像を開けません:\n{e}", parent=win)
                return
            orig_w, orig_h = img.width, img.height
            img.thumbnail((int(win.winfo_screenwidth() * 0.85),
                           int(win.winfo_screenheight() * 0.85)), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            win._photo_ref = photo
            preview_canvas.config(width=img.width, height=img.height)
            preview_canvas.delete("all")
            preview_canvas.create_image(img.width // 2, img.height // 2,
                                        image=photo, anchor="center")
            info = (f"{os.path.basename(fp)}  |  {orig_w} × {orig_h}"
                    f"  |  {os.path.getsize(fp) // 1024} KB"
                    f"  |  {idx + 1} / {len(self._display)}")
            lbl_info.config(text=info)

        def navigate(delta):
            load(current[0] + delta)

        def remove_and_navigate():
            idx = current[0]
            fp = self._display[idx]
            self.remove_file(fp)
            _navigate_after_removal(idx)

        def move_and_navigate():
            idx = current[0]
            fp = self._display[idx]
            self._preview_move_cb(fp)
            _navigate_after_removal(idx)

        def _navigate_after_removal(idx):
            new_len = len(self._display)
            if new_len == 0:
                win.destroy()
            elif idx < new_len:
                load(idx)
            else:
                load(new_len - 1)

        def show_context_menu(event):
            menu = tk.Menu(win, tearoff=0)
            if self._preview_move_cb:
                menu.add_command(label=self._preview_move_label,
                                 command=move_and_navigate)
            menu.add_command(label="この画像をリストから消去",
                             command=remove_and_navigate)
            menu.tk_popup(event.x_root, event.y_root)

        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Left>",   lambda e: navigate(-1))
        win.bind("<Right>",  lambda e: navigate(1))
        preview_canvas.bind("<Button-1>",    lambda e: navigate(1))
        preview_canvas.bind("<Button-3>",    show_context_menu)
        preview_canvas.bind("<MouseWheel>",
                            lambda e: navigate(-1 if e.delta > 0 else 1))

        load(start_idx)
        win.focus_set()

    # ------------------------------------------------------------------ バックグラウンドローダー

    def _loader(self):
        while True:
            try:
                fp = self._pending.get(timeout=1)
                if fp not in self._photos and PIL_AVAILABLE:
                    photo = self._make_thumb(fp)
                    if photo:
                        self._photos[fp] = photo
                        self._schedule_redraw()
            except Empty:
                continue

    def _make_thumb(self, filepath: str):
        try:
            from PIL import Image, ImageTk
            img = Image.open(filepath)
            img.thumbnail((THUMB, THUMB), Image.LANCZOS)
            bg = Image.new("RGB", (THUMB, THUMB), (185, 185, 185))
            bg.paste(img, ((THUMB - img.width) // 2, (THUMB - img.height) // 2))
            return ImageTk.PhotoImage(bg)
        except Exception:
            return None
