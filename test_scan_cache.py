"""Offline regression tests for media-monitor 4.x (no GUI, no Windows exe)."""

from __future__ import annotations

import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import main as mm


class PhrasesFingerprintTests(unittest.TestCase):
    def test_order_independent(self) -> None:
        self.assertEqual(
            mm.phrases_fingerprint(["б", "а"]),
            mm.phrases_fingerprint(["а", "б"]),
        )

    def test_changes_when_words_change(self) -> None:
        self.assertNotEqual(
            mm.phrases_fingerprint(["а"]),
            mm.phrases_fingerprint(["а", "в"]),
        )

    def test_case_and_space_normalized(self) -> None:
        self.assertEqual(
            mm.phrases_fingerprint(["  Тройка "]),
            mm.phrases_fingerprint(["тройка"]),
        )


class CacheUrlKeyTests(unittest.TestCase):
    def test_strips_fragment_and_slash(self) -> None:
        self.assertEqual(
            mm.cache_url_key("https://ex.com/a/#x"),
            "https://ex.com/a",
        )

    def test_adds_https(self) -> None:
        self.assertEqual(mm.cache_url_key("ex.com/a"), "https://ex.com/a")


class ScanCacheTests(unittest.TestCase):
    def setUp(self) -> None:
        self._td = tempfile.TemporaryDirectory()
        self.path = Path(self._td.name) / "scan_cache.sqlite"
        self.cache = mm.ScanCache(self.path)
        self.cache.open()
        self.h = mm.phrases_fingerprint(["слово"])

    def tearDown(self) -> None:
        self.cache.close()
        self._td.cleanup()

    def test_skip_after_record(self) -> None:
        url = "https://example.com/news/1/"
        self.assertFalse(self.cache.should_skip(url, self.h, 30))
        self.cache.record(
            url, self.h, published=date(2026, 8, 1), had_hit=True, title="T"
        )
        self.assertEqual(self.cache.count(), 1)
        self.assertTrue(self.cache.should_skip(url, self.h, 30))

    def test_no_skip_when_disabled(self) -> None:
        url = "https://example.com/news/2"
        self.cache.record(url, self.h, published=None, had_hit=False)
        self.assertFalse(self.cache.should_skip(url, self.h, 0))

    def test_no_skip_when_phrases_change(self) -> None:
        url = "https://example.com/news/3"
        self.cache.record(url, self.h, published=None, had_hit=False)
        other = mm.phrases_fingerprint(["другое"])
        self.assertFalse(self.cache.should_skip(url, other, 30))

    def test_expired_entry_not_skipped(self) -> None:
        url = "https://example.com/news/4"
        self.cache.record(url, self.h, published=None, had_hit=False)
        old = (datetime.now() - timedelta(days=40)).isoformat(timespec="seconds")
        assert self.cache._conn is not None
        self.cache._conn.execute(
            "UPDATE scanned_pages SET scanned_at = ? WHERE url = ?",
            (old, mm.cache_url_key(url)),
        )
        self.cache._conn.commit()
        self.assertFalse(self.cache.should_skip(url, self.h, 30))

    def test_upsert_updates_same_url(self) -> None:
        url = "https://example.com/news/5"
        self.cache.record(url, self.h, published=None, had_hit=False, title="A")
        self.cache.record(
            url, self.h, published=date(2026, 7, 1), had_hit=True, title="B"
        )
        self.assertEqual(self.cache.count(), 1)
        assert self.cache._conn is not None
        row = self.cache._conn.execute(
            "SELECT had_hit, title, published_at FROM scanned_pages WHERE url = ?",
            (mm.cache_url_key(url),),
        ).fetchone()
        self.assertEqual(row, (1, "B", "2026-07-01"))

    def test_thread_safe_record_smoke(self) -> None:
        import threading

        errors: list[BaseException] = []

        def worker(i: int) -> None:
            try:
                self.cache.record(
                    f"https://example.com/t/{i}",
                    self.h,
                    published=None,
                    had_hit=False,
                )
            except BaseException as exc:  # noqa: BLE001
                errors.append(exc)

        threads = [threading.Thread(target=worker, args=(i,)) for i in range(40)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self.assertEqual(errors, [])
        self.assertEqual(self.cache.count(), 40)


class SettingsCacheOptionTests(unittest.TestCase):
    def test_example_loads_skip_seen_days(self) -> None:
        root = Path(__file__).resolve().parent
        s = mm.load_settings(root / "settings.example.txt")
        self.assertEqual(s.skip_seen_days, 30)

    def test_sync_appends_skip_seen_days_once(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "settings.txt"
            path.write_text("max_scan_urls=0\n", encoding="utf-8")
            mm.sync_settings_file(path)
            text = path.read_text(encoding="utf-8")
            self.assertEqual(text.count("skip_seen_days="), 1)
            mm.sync_settings_file(path)
            text2 = path.read_text(encoding="utf-8")
            self.assertEqual(text2.count("skip_seen_days="), 1)
            s = mm.load_settings(path)
            self.assertEqual(s.skip_seen_days, 30)

    def test_skip_seen_zero(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "settings.txt"
            path.write_text("skip_seen_days=0\n", encoding="utf-8")
            self.assertEqual(mm.load_settings(path).skip_seen_days, 0)


class DateFilterTests(unittest.TestCase):
    def test_june_outside_last_30_days_in_august(self) -> None:
        with patch(
            "main.effective_article_date_range",
            return_value=(date(2026, 7, 7), date(2026, 8, 6)),
        ):
            ok = mm.passes_date_filter(
                date(2026, 6, 15),
                None,
                date(2026, 8, 6),
                recent_days=30,
            )
            self.assertFalse(ok)

    def test_recent_in_window(self) -> None:
        with patch(
            "main.effective_article_date_range",
            return_value=(date(2026, 7, 7), date(2026, 8, 6)),
        ):
            ok = mm.passes_date_filter(
                date(2026, 7, 20),
                None,
                date(2026, 8, 6),
                recent_days=30,
            )
            self.assertTrue(ok)

    def test_unknown_date_kept(self) -> None:
        self.assertTrue(
            mm.passes_date_filter(None, None, date(2026, 8, 6), recent_days=30)
        )

    def test_effective_range_last_days(self) -> None:
        lo, hi = mm.effective_article_date_range(
            None, date(2026, 1, 1), 30, today=date(2026, 8, 6)
        )
        self.assertEqual(lo, date(2026, 7, 7))
        self.assertEqual(hi, date(2026, 8, 6))


class ScanPageCacheRecordTests(unittest.TestCase):
    """scan_page should record into cache even when date filter rejects."""

    def test_records_when_date_filtered_out(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            cache = mm.ScanCache(Path(td) / "c.sqlite")
            cache.open()
            h = mm.phrases_fingerprint(["тройка"])
            settings = MagicMock()
            settings.ssl_verify = False
            settings.article_date_not_older_than = None
            settings.article_date_not_later_than = date(2026, 8, 6)
            settings.article_date_last_days = 30

            html = """
            <html><head><title>Старая</title></head>
            <body>
              <article>
                <h1>Старая новость</h1>
                <time datetime="2026-06-01">1 июня 2026</time>
                <p>Здесь слово тройка встречается.</p>
              </article>
            </body></html>
            """
            url = "https://example.com/news/old-article-12345"

            with (
                patch("main.fetch_html", return_value=html),
                patch(
                    "main.effective_article_date_range",
                    return_value=(date(2026, 7, 7), date(2026, 8, 6)),
                ),
            ):
                hits = mm.scan_page(
                    url,
                    ["тройка"],
                    session=MagicMock(),
                    scanned_at=datetime(2026, 8, 6, 12, 0, 0),
                    settings=settings,
                    auth=None,
                    cache=cache,
                    phrases_hash=h,
                )
            self.assertEqual(hits, [])
            self.assertTrue(cache.should_skip(url, h, 30))
            cache.close()

    def test_records_hit(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            cache = mm.ScanCache(Path(td) / "c.sqlite")
            cache.open()
            h = mm.phrases_fingerprint(["тройка"])
            settings = MagicMock()
            settings.ssl_verify = False
            settings.article_date_not_older_than = None
            settings.article_date_not_later_than = date(2026, 8, 6)
            settings.article_date_last_days = 30

            html = """
            <html><head><title>Свежая</title></head>
            <body>
              <article>
                <h1>Свежая новость</h1>
                <time datetime="2026-08-01">1 августа 2026</time>
                <p>Здесь слово тройка встречается.</p>
              </article>
            </body></html>
            """
            url = "https://example.com/news/fresh-article-999"

            with (
                patch("main.fetch_html", return_value=html),
                patch(
                    "main.effective_article_date_range",
                    return_value=(date(2026, 7, 7), date(2026, 8, 6)),
                ),
            ):
                hits = mm.scan_page(
                    url,
                    ["тройка"],
                    session=MagicMock(),
                    scanned_at=datetime(2026, 8, 6, 12, 0, 0),
                    settings=settings,
                    auth=None,
                    cache=cache,
                    phrases_hash=h,
                )
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0].phrase, "тройка")
            assert cache._conn is not None
            row = cache._conn.execute(
                "SELECT had_hit FROM scanned_pages WHERE url = ?",
                (mm.cache_url_key(url),),
            ).fetchone()
            self.assertEqual(row[0], 1)
            cache.close()


class ProcessSiteCacheSkipTests(unittest.TestCase):
    def test_cached_url_does_not_call_scan_page(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            cache = mm.ScanCache(Path(td) / "c.sqlite")
            cache.open()
            phrases = ["тройка"]
            h = mm.phrases_fingerprint(phrases)
            url = "https://example.com/news/cached-1"
            cache.record(url, h, published=date(2026, 8, 1), had_hit=False)

            settings = MagicMock()
            settings.ssl_verify = False
            settings.skip_seen_days = 30
            settings.max_scan_urls = 0
            settings.max_expand_links = 10

            progress = mm.RunProgress(total_sites=1, workers=1)

            with (
                patch("main.fetch_html", return_value="<html></html>") as fetch,
                patch(
                    "main.collect_urls_for_site",
                    return_value=[url],
                ),
                patch("main.scan_page") as scan,
            ):
                result = mm.process_site(
                    1,
                    1,
                    "https://example.com",
                    ["https://example.com/"],
                    phrases,
                    settings,
                    None,
                    set(),
                    datetime.now(),
                    progress,
                    cache=cache,
                    phrases_hash=h,
                )
            scan.assert_not_called()
            self.assertEqual(result.pages_skipped_cache, 1)
            self.assertEqual(result.pages_scanned, 0)
            # probe + maybe nothing else for scan
            self.assertGreaterEqual(fetch.call_count, 1)
            cache.close()


class MaxExpandLinksTests(unittest.TestCase):
    def test_zero_clamped_to_one(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "settings.txt"
            path.write_text("max_expand_links=0\n", encoding="utf-8")
            self.assertEqual(mm.load_settings(path).max_expand_links, 1)

    def test_large_value_allowed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "settings.txt"
            path.write_text("max_expand_links=1000\n", encoding="utf-8")
            self.assertEqual(mm.load_settings(path).max_expand_links, 1000)


class VersionTests(unittest.TestCase):
    def test_version_files_match(self) -> None:
        root = Path(__file__).resolve().parent
        self.assertEqual(mm.VERSION, "4.0.0")
        self.assertEqual(root.joinpath("VERSION").read_text().strip(), "4.0.0")


if __name__ == "__main__":
    unittest.main(verbosity=2)
