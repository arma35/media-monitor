"""
media-monitor — portable keyword scan of configured sites.
Configs and reports live next to the executable (USB-friendly).
"""

from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, TextIO
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

VERSION = "0.0.2"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)
REQUEST_TIMEOUT = 25
MAX_PAGE_CHARS = 500_000
LOG_NAME = "media-monitor.txt"


def app_dir() -> Path:
    """Directory of the running exe (frozen) or this script (dev)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class Tee:
    """Write the same text to console and an appendable log file."""

    def __init__(self, *streams: TextIO) -> None:
        self.streams = streams

    def write(self, data: str) -> int:
        for stream in self.streams:
            stream.write(data)
            stream.flush()
        return len(data)

    def flush(self) -> None:
        for stream in self.streams:
            stream.flush()


@dataclass(frozen=True)
class Hit:
    word: str
    url: str
    title: str
    scanned_at: datetime


def ensure_local_config(root: Path, name: str) -> Path:
    """
    Use local sites.txt / words.txt.
    If missing, copy from *.example.txt once — never overwrite an existing file.
    """
    target = root / name
    if target.is_file():
        return target
    example = root / name.replace(".txt", ".example.txt")
    if example.is_file():
        shutil.copy2(example, target)
        print(f"Created {name} from {example.name} (edit it for your list).")
        return target
    raise FileNotFoundError(
        f"Missing {name} and {example.name}. "
        "Add the example file from the repo or create the config manually."
    )


def load_lines(path: Path) -> list[str]:
    if not path.is_file():
        raise FileNotFoundError(f"Config not found: {path.name}")
    lines: list[str] = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        lines.append(line)
    return lines


def normalize_url(url: str) -> str:
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


def fetch_html(url: str, session: requests.Session) -> str:
    resp = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or resp.encoding or "utf-8"
    return resp.text


def extract_title(soup: BeautifulSoup) -> str:
    og = soup.find("meta", property="og:title")
    if og and og.get("content"):
        return " ".join(str(og["content"]).split())

    if soup.title and soup.title.string:
        return " ".join(soup.title.string.split())

    h1 = soup.find("h1")
    if h1:
        return " ".join(h1.get_text(" ", strip=True).split())

    return ""


def extract_page(html: str, base_url: str) -> tuple[str, str, list[str]]:
    """Return (title, text, same-host links)."""
    soup = BeautifulSoup(html, "lxml")
    title = extract_title(soup)

    for tag in soup(["script", "style", "noscript", "svg", "template"]):
        tag.decompose()

    text = soup.get_text(separator=" ", strip=True)
    if len(text) > MAX_PAGE_CHARS:
        text = text[:MAX_PAGE_CHARS]

    host = urlparse(base_url).netloc.lower()
    links: list[str] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue
        absolute = urljoin(base_url, href)
        parsed = urlparse(absolute)
        if parsed.scheme not in ("http", "https"):
            continue
        if parsed.netloc.lower() != host:
            continue
        clean = absolute.split("#", 1)[0]
        if clean not in seen:
            seen.add(clean)
            links.append(clean)
    return title, text, links


def word_present(text: str, word: str) -> bool:
    """Case-insensitive match; whole word for Latin/Cyrillic tokens."""
    pattern = rf"(?<!\w){re.escape(word)}(?!\w)"
    return re.search(pattern, text, flags=re.IGNORECASE | re.UNICODE) is not None


def scan_page(
    url: str,
    words: list[str],
    session: requests.Session,
    scanned_at: datetime,
) -> list[Hit]:
    html = fetch_html(url, session)
    title, text, _ = extract_page(html, url)
    hits: list[Hit] = []
    for word in words:
        if word_present(text, word):
            hits.append(Hit(word=word, url=url, title=title, scanned_at=scanned_at))
    return hits


def collect_urls_to_scan(seed_urls: Iterable[str], session: requests.Session) -> list[str]:
    """
    Scan each configured URL. If a seed looks like a listing/home page,
    also include same-host article links found there (capped).
    """
    result: list[str] = []
    seen: set[str] = set()

    for seed in seed_urls:
        url = normalize_url(seed)
        if url in seen:
            continue
        seen.add(url)
        result.append(url)

        path = urlparse(url).path.rstrip("/")
        if path.count("/") <= 1:
            try:
                html = fetch_html(url, session)
                _, _, links = extract_page(html, url)
                for link in links[:40]:
                    if link not in seen:
                        seen.add(link)
                        result.append(link)
            except requests.RequestException as exc:
                print(f"  [warn] cannot expand {url}: {exc}")

    return result


def write_report(hits: list[Hit], reports_dir: Path, generated_at: datetime) -> Path:
    reports_dir.mkdir(parents=True, exist_ok=True)
    stamp = generated_at.strftime("%Y-%m-%d_%H-%M-%S")
    out_path = reports_dir / f"media-monitor-{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Hits"

    headers = ("слово", "ссылка на страницу", "название статьи", "дата время скана")
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    for hit in hits:
        ws.append(
            (
                hit.word,
                hit.url,
                hit.title,
                hit.scanned_at.strftime("%Y-%m-%d %H:%M:%S"),
            )
        )

    for row in ws.iter_rows(min_row=2, max_col=2, max_row=ws.max_row):
        cell = row[1]
        if cell.value:
            cell.hyperlink = str(cell.value)
            cell.style = "Hyperlink"

    widths = (24, 70, 50, 22)
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    wb.save(out_path)
    return out_path


def run() -> int:
    root = app_dir()
    reports_dir = root / "reports"

    print(f"media-monitor v{VERSION}")
    print(f"Working directory: {root}")

    try:
        sites_path = ensure_local_config(root, "sites.txt")
        words_path = ensure_local_config(root, "words.txt")
        sites = load_lines(sites_path)
        words = load_lines(words_path)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Sites: {sites_path.name}")
    print(f"Words: {words_path.name}")

    if not sites:
        print("ERROR: sites.txt is empty (no URLs).")
        return 1
    if not words:
        print("ERROR: words.txt is empty (no keywords).")
        return 1

    print(f"Loaded {len(sites)} site(s), {len(words)} word(s).")
    scanned_at = datetime.now()
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru,en;q=0.8"})

    print("Collecting pages…")
    urls = collect_urls_to_scan(sites, session)
    print(f"Will scan {len(urls)} page(s).")

    hits: list[Hit] = []
    errors = 0
    for i, url in enumerate(urls, start=1):
        print(f"[{i}/{len(urls)}] {url}")
        try:
            page_hits = scan_page(url, words, session, scanned_at)
            if page_hits:
                print(f"  → {len(page_hits)} hit(s)")
                hits.extend(page_hits)
        except requests.RequestException as exc:
            errors += 1
            print(f"  [error] {exc}")

    report_path = write_report(hits, reports_dir, datetime.now())
    print()
    print(f"Done. Hits: {len(hits)}. Errors: {errors}.")
    print(f"Report: {report_path}")
    return 0


def main() -> int:
    root = app_dir()
    log_path = root / LOG_NAME
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    log_file = open(log_path, "a", encoding="utf-8")
    try:
        started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"\n===== media-monitor v{VERSION} start {started} =====\n")
        log_file.flush()
        sys.stdout = Tee(original_stdout, log_file)  # type: ignore[assignment]
        sys.stderr = Tee(original_stderr, log_file)  # type: ignore[assignment]
        return run()
    finally:
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        ended = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"===== end {ended} =====\n")
        log_file.close()


if __name__ == "__main__":
    raise SystemExit(main())
