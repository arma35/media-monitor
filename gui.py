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

    def __init__(
        self,
        root: tk.Tk,
        widget: scrolledtext.ScrolledText,
        *,
        max_lines: int = 2000,
        max_queue: int = 1500,
        batch_per_tick: int = 200,
    ) -> None:
        self.root = root
        self.widget = widget
        # Bounded queue: under load oldest lines are dropped, not piled up.
        self.q: queue.Queue[str] = queue.Queue(maxsize=max_queue)
        self.max_lines = max_lines
        self.batch_per_tick = batch_per_tick
        self._dropped = 0

    def write_line(self, line: str) -> None:
        # Bound the queue so a flood of detail lines cannot freeze Tk forever.
        while True:
            try:
                self.q.put_nowait(line)
                return
            except queue.Full:
                try:
                    self.q.get_nowait()
                    self._dropped += 1
                except queue.Empty:
                    return

    def clear(self) -> None:
        self.widget.configure(state="normal")
        self.widget.delete("1.0", "end")
        self.widget.configure(state="disabled")
        self._dropped = 0

    def pump(self) -> None:
        # One Text mutate per tick (join batch) — per-line insert/see freezes Tk.
        batch: list[str] = []
        try:
            while len(batch) < self.batch_per_tick:
                batch.append(self.q.get_nowait())
        except queue.Empty:
            pass
        n = len(batch)
        dropped = self._dropped
        if dropped:
            self._dropped = 0
        if batch or dropped:
            self.widget.configure(state="normal")
            if batch:
                self.widget.insert("end", "".join(batch))
            if dropped:
                self.widget.insert(
                    "end",
                    f"… пропущено строк лога (чтобы окно не зависало): {dropped}\n",
                )
            if self.max_lines > 0:
                last = int(float(self.widget.index("end-1c")))
                if last > self.max_lines:
                    self.widget.delete("1.0", f"{last - self.max_lines}.0")
            self.widget.see("end")
            self.widget.configure(state="disabled")
        # Fast while draining; slow right down when idle so a finished
        # window does not keep burning a few % CPU on empty after() ticks.
        if n >= self.batch_per_tick:
            delay = 40
        elif n or dropped:
            delay = 100
        else:
            delay = 750
        self.root.after(delay, self.pump)


class App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(f"Media Monitor v{mm.VERSION}")
        self.root.geometry("920x600")
        self.root.minsize(700, 460)

        self.worker: threading.Thread | None = None
        self.running = False
        self.paused = False
        self.last_report: Path | None = None
        self._progress_q: queue.Queue[dict] = queue.Queue()
        self._cached_workers = 0
        self._run_started_mono: float | None = None
        self._elapsed_job: str | None = None
        self._pause_started_mono: float | None = None
        self._pause_accumulated = 0.0

        self._build()
        # Show the window first; load configs on the next Tk tick so startup
        # does not look frozen while reading settings / creating examples.
        self.root.after(0, self._refresh_counts)
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
        self.pages_lbl = ttk.Label(info, text="", font=("Segoe UI", 10))
        self.pages_lbl.pack(anchor="w", pady=(2, 0))
        self.cache_lbl = ttk.Label(info, text="", font=("Segoe UI", 10))
        self.cache_lbl.pack(anchor="w", pady=(2, 0))

        self.progress_bar = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=(0, 4))

        btns = ttk.Frame(root)
        btns.pack(fill="x", **pad)

        self.btn_start = ttk.Button(btns, text="Старт", command=self.on_start)
        self.btn_start.pack(side="left", padx=(0, 6))

        self.btn_pause = ttk.Button(
            btns, text="Пауза", command=self.on_pause, state="disabled"
        )
        self.btn_pause.pack(side="left", padx=(0, 6))

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

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.notebook = notebook

        tab_status = ttk.Frame(notebook)
        tab_full = ttk.Frame(notebook)
        notebook.add(tab_status, text="Статус")
        notebook.add(tab_full, text="Полный лог")

        self.log = scrolledtext.ScrolledText(
            tab_status,
            wrap="word",
            height=22,
            state="disabled",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, padx=2, pady=2)

        self.full_log = scrolledtext.ScrolledText(
            tab_full,
            wrap="word",
            height=22,
            state="disabled",
            font=("Consolas", 9),
        )
        self.full_log.pack(fill="both", expand=True, padx=2, pady=2)

        # Status: fewer lines, smaller queue. Detail: larger buffer but still capped.
        self.gui_log = GuiLog(root, self.log, max_lines=1500, max_queue=800, batch_per_tick=80)
        self.gui_log.pump()
        self.gui_detail = GuiLog(
            root, self.full_log, max_lines=4000, max_queue=2000, batch_per_tick=250
        )
        self.gui_detail.pump()

        hint = ttk.Label(
            root,
            text=(
                "«Статус» — кратко по сайтам. «Полный лог» — что качается прямо сейчас. "
                "Пауза — временно остановить набор новых страниц; Стоп — прервать и сохранить Excel. "
                "Консоль: --console"
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
        elapsed = time.monotonic() - self._run_started_mono - self._pause_accumulated
        if self.paused and self._pause_started_mono is not None:
            elapsed -= time.monotonic() - self._pause_started_mono
        self.elapsed_lbl.configure(text=f"Время: {self._format_elapsed(elapsed)}")
        if self.running:
            self._elapsed_job = self.root.after(500, self._tick_elapsed)

    def _start_elapsed(self) -> None:
        self._run_started_mono = time.monotonic()
        self._pause_accumulated = 0.0
        self._pause_started_mono = None
        self.elapsed_lbl.configure(text="Время: 0с")
        if self._elapsed_job is not None:
            try:
                self.root.after_cancel(self._elapsed_job)
            except Exception:
                pass
        self._elapsed_job = self.root.after(500, self._tick_elapsed)

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
            # Keep pages_lbl / cache_lbl from the last run until the next Start.
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
        # While scanning — snappy; after finish — barely tick.
        delay = 150 if self.running else 1000
        self.root.after(delay, self._pump_progress)

    def _apply_progress(self, snap: dict) -> None:
        done = int(snap.get("done", 0))
        total = max(1, int(snap.get("total", 1)))
        hits = int(snap.get("hits", 0))
        pages = int(snap.get("pages", 0))
        active = int(snap.get("active", 0))
        workers = int(snap.get("workers", self._cached_workers))
        unavailable = int(snap.get("unavailable", 0))
        cancelling = bool(snap.get("cancelling", False))
        paused = bool(snap.get("paused", False))
        names = snap.get("active_names") or []
        cache_before = int(snap.get("cache_before", 0))
        cache_skipped = int(snap.get("cache_skipped", 0))
        cache_added = int(snap.get("cache_added", 0))

        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = done

        names_txt = ", ".join(names)
        if active > len(names):
            names_txt += f"… (+{active - len(names)})"

        scanned_now = max(0, pages - cache_skipped)
        if cancelling:
            self.status.configure(text="Остановка…")
            self.progress_lbl.configure(
                text=(
                    f"Остановка: готово {done}/{total} | находок {hits} | "
                    f"ещё выходят: {active} | параллельно {workers}. "
                    "Сеть закрыта, Excel сохранится."
                )
            )
        elif paused:
            self.status.configure(text=f"Пауза… {done}/{total}")
            self.progress_lbl.configure(
                text=(
                    f"Пауза: готово {done}/{total} | находок {hits} | "
                    f"активных сайтов {active}. Нажмите «Продолжить»."
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

        self.pages_lbl.configure(
            text=(
                f"Страниц проверено в этом запуске: {pages} "
                f"(скан {scanned_now} + кэш-пропуск {cache_skipped})"
            )
        )
        self.cache_lbl.configure(
            text=(
                f"БД кэш: было {cache_before} | проверено ранее (пропуск) {cache_skipped} | "
                f"добавлено сейчас {cache_added}"
            )
        )

    def append(self, text: str) -> None:
        self.gui_log.write_line(text if text.endswith("\n") else text + "\n")

    def on_start(self) -> None:
        if self.running:
            return
        # Do not call _refresh_counts here — it can stall the UI briefly;
        # counts were refreshed at idle and will refresh again when done.
        self.running = True
        self.paused = False
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal", text="Пауза")
        self.btn_stop.configure(state="normal")
        self.status.configure(text="Сканирование…")
        self.progress_lbl.configure(text="Запуск…")
        self.pages_lbl.configure(text="Страниц проверено в этом запуске: 0")
        self.cache_lbl.configure(text="БД кэш: …")
        self.progress_bar["value"] = 0
        self._start_elapsed()
        self.gui_log.clear()
        self.gui_detail.clear()
        mm.clear_cancel()
        mm.clear_pause()
        mm.set_log_callback(self.gui_log.write_line)
        mm.set_detail_callback(self.gui_detail.write_line)
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
                mm.set_log_file(log_file)
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
                mm.set_log_file(None)
                if log_file is not None:
                    ended = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_file.write(
                        f"===== media-monitor v{mm.VERSION} end {ended} =====\n"
                    )
                    log_file.close()
                    mm.rotate_log_if_needed(mm.app_dir() / mm.LOG_NAME)
                mm.set_log_callback(None)
                mm.set_detail_callback(None)
                mm.set_progress_callback(None)
                self.root.after(0, lambda: self._on_done(code))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _on_done(self, code: int) -> None:
        self.running = False
        self.paused = False
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled", text="Пауза")
        self.btn_stop.configure(state="disabled")
        self._stop_elapsed(freeze=True)
        if mm.is_cancelled():
            self.status.configure(text="Остановлено — отчёт сохранён")
        elif code == 0:
            self.status.configure(text="Готово")
        else:
            self.status.configure(text="Завершено с ошибкой")
        self._refresh_counts()

    def on_pause(self) -> None:
        if not self.running:
            return
        if self.paused:
            mm.request_resume()
            if self._pause_started_mono is not None:
                self._pause_accumulated += time.monotonic() - self._pause_started_mono
                self._pause_started_mono = None
            self.paused = False
            self.btn_pause.configure(text="Пауза")
            self.status.configure(text="Сканирование…")
            self.append("\nПауза снята — продолжаю.\n")
        else:
            mm.request_pause()
            self._pause_started_mono = time.monotonic()
            self.paused = True
            self.btn_pause.configure(text="Продолжить")
            self.status.configure(text="Пауза…")
            self.append(
                "\nПауза: новые страницы не берутся, текущие запросы докачиваются. "
                "Нажмите «Продолжить».\n"
            )

    def on_stop(self) -> None:
        if not self.running:
            return
        if self.paused:
            # Ensure workers are not stuck in pause wait.
            mm.request_resume()
            self.paused = False
            self.btn_pause.configure(text="Пауза")
        mm.request_cancel()
        self.status.configure(text="Остановка…")
        self.btn_pause.configure(state="disabled")
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
