"""Simple tkinter UI for media-monitor."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
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
        self.root.geometry("920x600")
        self.root.minsize(700, 460)

        self.worker: threading.Thread | None = None
        self.running = False
        self.last_report: Path | None = None
        self._progress_q: queue.Queue[dict] = queue.Queue()
        self._cached_workers = 0
        self._run_started_mono: float | None = None
        self._elapsed_job: str | None = None

        self._build()
        self._refresh_counts()
        self._pump_progress()

    def _build(self) -> None:
        root = self.root
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(root)
        top.pack(fill="x", **pad)

        self.status = ttk.Label(top, text="Готов к запуску", font=("Segoe UI", 11))
        self.status.pack(side="left")

        self.elapsed_lbl = ttk.Label(top, text="Время: —", font=("Segoe UI", 11))
        self.elapsed_lbl.pack(side="left", padx=(16, 0))

        self.version_lbl = ttk.Label(top, text=f"v{mm.VERSION}")
        self.version_lbl.pack(side="right")

        info = ttk.Frame(root)
        info.pack(fill="x", **pad)
        self.info_lbl = ttk.Label(info, text="")
        self.info_lbl.pack(anchor="w")
        self.dates_lbl = ttk.Label(info, text="", font=("Segoe UI", 10))
        self.dates_lbl.pack(anchor="w", pady=(2, 0))
        self.progress_lbl = ttk.Label(info, text="", font=("Segoe UI", 10))
        self.progress_lbl.pack(anchor="w", pady=(2, 0))

        self.progress_bar = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=(0, 4))

        btns = ttk.Frame(root)
        btns.pack(fill="x", **pad)

        self.btn_start = ttk.Button(btns, text="Старт", command=self.on_start)
        self.btn_start.pack(side="left", padx=(0, 6))

        self.btn_stop = ttk.Button(btns, text="Стоп", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 6))

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
            text=(
                "Стоп: сразу закрывает сеть и не ждёт длинные таймауты; "
                "Excel всё равно сохранится. Консоль: --console"
            ),
            foreground="#555",
            wraplength=880,
        )
        hint.pack(anchor="w", padx=10, pady=(0, 8))

    def _format_elapsed(self, seconds: float) -> str:
        total = max(0, int(seconds))
        h, rem = divmod(total, 3600)
        m, s = divmod(rem, 60)
        if h:
            return f"{h}ч {m:02d}м {s:02d}с"
        if m:
            return f"{m}м {s:02d}с"
        return f"{s}с"

    def _tick_elapsed(self) -> None:
        self._elapsed_job = None
        if self._run_started_mono is None:
            return
        elapsed = time.monotonic() - self._run_started_mono
        self.elapsed_lbl.configure(text=f"Время: {self._format_elapsed(elapsed)}")
        if self.running:
            self._elapsed_job = self.root.after(250, self._tick_elapsed)

    def _start_elapsed(self) -> None:
        self._run_started_mono = time.monotonic()
        self.elapsed_lbl.configure(text="Время: 0с")
        if self._elapsed_job is not None:
            try:
                self.root.after_cancel(self._elapsed_job)
            except Exception:
                pass
        self._elapsed_job = self.root.after(250, self._tick_elapsed)

    def _stop_elapsed(self, *, freeze: bool = True) -> None:
        if self._elapsed_job is not None:
            try:
                self.root.after_cancel(self._elapsed_job)
            except Exception:
                pass
            self._elapsed_job = None
        if freeze and self._run_started_mono is not None:
            elapsed = time.monotonic() - self._run_started_mono
            self.elapsed_lbl.configure(text=f"Время: {self._format_elapsed(elapsed)}")
        elif not freeze:
            self.elapsed_lbl.configure(text="Время: —")
            self._run_started_mono = None

    def _workers_from_settings(self, settings: mm.Settings, site_count: int) -> int:
        if site_count <= 0:
            return 0
        if settings.site_workers > 0:
            return max(1, min(settings.site_workers, site_count))
        return site_count

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

        workers = 0
        try:
            settings_path = mm.resolve_settings_path(root)
            mm.sync_settings_file(settings_path)
            settings = mm.load_settings(settings_path)
            workers = self._workers_from_settings(settings, sites)
            self._cached_workers = workers
            eff_from, eff_to = mm.effective_article_date_range(
                settings.article_date_not_older_than,
                settings.article_date_not_later_than,
                settings.article_date_last_days,
            )
            range_ru = mm.format_article_date_range_ru(eff_from, eff_to)
            self.dates_lbl.configure(text=f"Статьи: {range_ru}")
        except Exception:
            self.dates_lbl.configure(text="Статьи: (не удалось прочитать settings.txt)")
            workers = self._cached_workers

        self.info_lbl.configure(
            text=(
                f"Папка: {root}   |   сайтов: {sites}   |   фраз: {words}   |   "
                f"параллельно: {workers}"
            )
        )
        if not self.running:
            self.progress_lbl.configure(text="")
            self.progress_bar["value"] = 0

    def _on_progress(self, snap: dict) -> None:
        """Called from worker threads — queue for UI thread."""
        self._progress_q.put(snap)

    def _pump_progress(self) -> None:
        latest: dict | None = None
        try:
            while True:
                latest = self._progress_q.get_nowait()
        except queue.Empty:
            pass
        if latest is not None:
            self._apply_progress(latest)
        self.root.after(150, self._pump_progress)

    def _apply_progress(self, snap: dict) -> None:
        done = int(snap.get("done", 0))
        total = max(1, int(snap.get("total", 1)))
        hits = int(snap.get("hits", 0))
        active = int(snap.get("active", 0))
        workers = int(snap.get("workers", self._cached_workers))
        unavailable = int(snap.get("unavailable", 0))
        cancelling = bool(snap.get("cancelling", False))
        names = snap.get("active_names") or []

        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = done

        names_txt = ", ".join(names)
        if active > len(names):
            names_txt += f"… (+{active - len(names)})"

        if cancelling:
            self.status.configure(text="Остановка…")
            self.progress_lbl.configure(
                text=(
                    f"Остановка: готово {done}/{total} | находок {hits} | "
                    f"ещё выходят: {active} | параллельно {workers}. "
                    "Сеть закрыта, Excel сохранится."
                )
            )
        else:
            self.status.configure(text=f"Сканирование… {done}/{total}")
            line = (
                f"Готово {done}/{total} | находок {hits} | "
                f"сейчас качают {active} | параллельно {workers}"
            )
            if unavailable:
                line += f" | недоступных {unavailable}"
            if names_txt:
                line += f" | {names_txt}"
            self.progress_lbl.configure(text=line)

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
        self.progress_lbl.configure(text="Запуск…")
        self.progress_bar["value"] = 0
        self._start_elapsed()
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        mm.clear_cancel()
        mm.set_log_callback(self.gui_log.write_line)
        mm.set_progress_callback(self._on_progress)

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
                mm.set_progress_callback(None)
                self.root.after(0, lambda: self._on_done(code))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _on_done(self, code: int) -> None:
        self.running = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self._stop_elapsed(freeze=True)
        if mm.is_cancelled():
            self.status.configure(text="Остановлено — отчёт сохранён")
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
        self.append(
            "\nОстановка: закрываю сеть, новые страницы не берутся. "
            "Excel сохранится с тем, что уже нашли.\n"
        )

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
