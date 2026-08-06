"""Simple tkinter UI for media-monitor."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import main as mm


class GuiLog:
    """Queue lines from worker threads; drained on the Tk main thread."""

    def __init__(self, root: tk.Tk, widget: scrolledtext.ScrolledText) -> None:
        self.root = root
        self.widget = widget
        self.q: queue.Queue[str] = queue.Queue()

    def write_line(self, line: str) -> None:
        self.q.put(line)

    def pump(self) -> None:
        try:
            while True:
                line = self.q.get_nowait()
                self.widget.configure(state="normal")
                self.widget.insert("end", line)
                self.widget.see("end")
                self.widget.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.pump)


class App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(f"Media Monitor v{mm.VERSION}")
        self.root.geometry("860x560")
        self.root.minsize(640, 420)

        self.worker: threading.Thread | None = None
        self.running = False
        self.last_report: Path | None = None

        self._build()
        self._refresh_counts()

    def _build(self) -> None:
        root = self.root
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(root)
        top.pack(fill="x", **pad)

        self.status = ttk.Label(top, text="Готов к запуску", font=("Segoe UI", 11))
        self.status.pack(side="left")

        self.version_lbl = ttk.Label(top, text=f"v{mm.VERSION}")
        self.version_lbl.pack(side="right")

        info = ttk.Frame(root)
        info.pack(fill="x", **pad)
        self.info_lbl = ttk.Label(info, text="")
        self.info_lbl.pack(anchor="w")

        btns = ttk.Frame(root)
        btns.pack(fill="x", **pad)

        self.btn_start = ttk.Button(btns, text="Старт", command=self.on_start)
        self.btn_start.pack(side="left", padx=(0, 6))

        self.btn_stop = ttk.Button(btns, text="Стоп", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 6))

        ttk.Button(btns, text="Папка программы", command=self.open_app_dir).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(btns, text="Отчёты", command=self.open_reports).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(btns, text="sites.txt", command=lambda: self.open_file("sites.txt")).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(btns, text="words.txt", command=lambda: self.open_file("words.txt")).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(
            btns, text="settings.txt", command=lambda: self.open_file("settings.txt")
        ).pack(side="left", padx=(0, 6))

        self.log = scrolledtext.ScrolledText(
            root,
            wrap="word",
            height=24,
            state="disabled",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.gui_log = GuiLog(root, self.log)
        self.gui_log.pump()

        hint = ttk.Label(
            root,
            text="Конфиги рядом с exe. Консольный режим: --console",
            foreground="#555",
        )
        hint.pack(anchor="w", padx=10, pady=(0, 8))

    def _refresh_counts(self) -> None:
        root = mm.app_dir()
        try:
            sites = len(mm.load_lines(mm.ensure_local_config(root, "sites.txt")))
        except Exception:
            sites = 0
        try:
            words = len(mm.load_lines(mm.ensure_local_config(root, "words.txt")))
        except Exception:
            words = 0
        self.info_lbl.configure(
            text=f"Папка: {root}   |   сайтов: {sites}   |   фраз: {words}"
        )

    def append(self, text: str) -> None:
        self.gui_log.write_line(text if text.endswith("\n") else text + "\n")

    def on_start(self) -> None:
        if self.running:
            return
        self._refresh_counts()
        self.running = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.status.configure(text="Сканирование…")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        mm.clear_cancel()
        mm.set_log_callback(self.gui_log.write_line)

        def work() -> None:
            code = 1
            log_file = None
            mm.ensure_stdio()
            original_stdout = sys.stdout
            original_stderr = sys.stderr
            try:
                root = mm.app_dir()
                log_path = root / mm.LOG_NAME
                mm.rotate_log_if_needed(log_path)
                log_file = open(log_path, "a", encoding="utf-8")
                started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_file.write(
                    f"\n===== media-monitor v{mm.VERSION} start {started} =====\n"
                )
                log_file.flush()
                sys.stdout = mm.Tee(original_stdout, log_file)  # type: ignore[assignment]
                sys.stderr = mm.Tee(original_stderr, log_file)  # type: ignore[assignment]
                code = mm.run(interactive_auth=False)
            except Exception as exc:  # noqa: BLE001
                try:
                    mm.log_print(f"ERROR: {exc}")
                except Exception:
                    self.gui_log.write_line(f"ERROR: {exc}\n")
                code = 1
            finally:
                sys.stdout = original_stdout
                sys.stderr = original_stderr
                if log_file is not None:
                    ended = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_file.write(
                        f"===== media-monitor v{mm.VERSION} end {ended} =====\n"
                    )
                    log_file.close()
                    mm.rotate_log_if_needed(mm.app_dir() / mm.LOG_NAME)
                mm.set_log_callback(None)
                self.root.after(0, lambda: self._on_done(code))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _on_done(self, code: int) -> None:
        self.running = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        if mm.is_cancelled():
            self.status.configure(text="Остановлено")
        elif code == 0:
            self.status.configure(text="Готово")
        else:
            self.status.configure(text="Завершено с ошибкой")
        self._refresh_counts()

    def on_stop(self) -> None:
        if not self.running:
            return
        mm.request_cancel()
        self.status.configure(text="Остановка…")
        self.append("Запрошена остановка…\n")

    def open_app_dir(self) -> None:
        self._open_path(mm.app_dir())

    def open_reports(self) -> None:
        path = mm.app_dir() / "reports"
        path.mkdir(parents=True, exist_ok=True)
        self._open_path(path)

    def open_file(self, name: str) -> None:
        root = mm.app_dir()
        path = root / name
        if not path.is_file():
            try:
                mm.ensure_local_config(root, name)
            except Exception as exc:
                messagebox.showerror("Файл", str(exc))
                return
        self._open_path(path)

    def _open_path(self, path: Path) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:
            messagebox.showerror("Открыть", str(exc))

    def run(self) -> int:
        self.root.mainloop()
        return 0


def run_gui() -> int:
    app = App()
    return app.run()


if __name__ == "__main__":
    raise SystemExit(run_gui())
