"""
media-monitor — portable keyword scan of configured sites.
Configs and reports live next to the executable (USB-friendly).
"""

from __future__ import annotations

import getpass
import json
import re
import shutil
import sys
import threading
from dataclasses import dataclass
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Iterable, TextIO
from urllib.parse import quote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

VERSION = "0.0.9"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)
REQUEST_TIMEOUT = 25
MAX_PAGE_CHARS = 500_000
LOG_NAME = "media-monitor.txt"
DEFAULT_AUTH_TIMEOUT = 180


class AuthRequiredError(Exception):
    """Page requires authentication that was not provided or failed."""


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
class Settings:
    # Articles with publish date on/after this day ("не старше"). Empty = no min.
    article_date_not_older_than: date | None
    # Articles with publish date on/before this day ("дата ДО"). Empty = today.
    article_date_not_later_than: date
    auth_timeout_seconds: int


@dataclass(frozen=True)
class Hit:
    phrase: str
    url: str
    published_at: str
    title: str
    scanned_at: datetime


def ensure_local_config(root: Path, name: str) -> Path:
    """
    Use local config file.
    If missing, copy from *.example.txt once — never overwrite an existing file.
    """
    target = root / name
    if target.is_file():
        return target

    # Backward/forward compatibility:
    # - historically user asked for `setting.txt` (singular)
    # - example file is `settings.example.txt` (plural)
    if name.lower() == "setting.txt":
        example = root / "settings.example.txt"
    else:
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


def load_settings(path: Path) -> Settings:
    raw: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        raw[key.strip().lower()] = value.strip()

    min_date: date | None = None
    older_raw = raw.get("article_date_not_older_than", "")
    if older_raw:
        min_date = date.fromisoformat(older_raw)

    # "дата ДО": empty means today
    later_raw = raw.get("article_date_not_later_than", "")
    max_date = date.fromisoformat(later_raw) if later_raw else date.today()

    timeout = DEFAULT_AUTH_TIMEOUT
    timeout_raw = raw.get("auth_timeout_seconds", "")
    if timeout_raw:
        timeout = max(1, int(timeout_raw))

    return Settings(
        article_date_not_older_than=min_date,
        article_date_not_later_than=max_date,
        auth_timeout_seconds=timeout,
    )


def normalize_url(url: str) -> str:
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def timed_input(prompt: str, timeout_sec: float) -> str | None:
    """Read a line from stdin; return None on timeout."""
    if timeout_sec <= 0:
        print(f"{prompt}(timeout)")
        return None

    result: list[str | None] = [None]
    done = threading.Event()

    def worker() -> None:
        try:
            # Print prompt to real console so password flow stays usable under Tee.
            sys.__stdout__.write(prompt)
            sys.__stdout__.flush()
            result[0] = sys.__stdin__.readline()
        except Exception:
            result[0] = None
        finally:
            done.set()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    if not done.wait(timeout_sec):
        print("(timeout)")
        return None
    value = result[0]
    if value is None:
        return None
    return value.rstrip("\r\n")


def timed_password(prompt: str, timeout_sec: float) -> str | None:
    """Read password without echo; return None on timeout."""
    if timeout_sec <= 0:
        print(f"{prompt}(timeout)")
        return None

    result: list[str | None] = [None]
    done = threading.Event()

    def worker() -> None:
        try:
            result[0] = getpass.getpass(prompt)
        except Exception:
            result[0] = None
        finally:
            done.set()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    if not done.wait(timeout_sec):
        print("(timeout)")
        return None
    return result[0]


def prompt_credentials(timeout_sec: int) -> tuple[str, str] | None:
    """
    Ask for login/password for sites that need auth.
    Wait up to timeout_sec total; on timeout or empty login — skip auth.
    """
    print()
    print(
        "Авторизация для соцсетей (необязательно). "
        f"На ввод логина и пароля — до {timeout_sec} сек. "
        "Пустой логин или таймаут = пропуск."
    )
    deadline = datetime.now().timestamp() + timeout_sec

    login = timed_input("Логин: ", deadline - datetime.now().timestamp())
    if login is None:
        print("Авторизация пропущена (таймаут логина).")
        return None
    login = login.strip()
    if not login:
        print("Авторизация пропущена (пустой логин).")
        return None

    password = timed_password("Пароль: ", deadline - datetime.now().timestamp())
    if password is None:
        print("Авторизация пропущена (таймаут пароля).")
        return None

    print("Учётные данные приняты (HTTP Basic Auth для 401/403).")
    return login, password


def fetch_html(
    url: str,
    session: requests.Session,
    auth: tuple[str, str] | None = None,
) -> str:
    resp = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
    if resp.status_code in (401, 403) and auth:
        resp = session.get(
            url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            auth=auth,
        )
    if resp.status_code in (401, 403):
        raise AuthRequiredError(f"HTTP {resp.status_code} auth required/failed")
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or resp.encoding or "utf-8"
    return resp.text


def extract_title(soup: BeautifulSoup) -> str:
    og = soup.find("meta", property="og:title")
    if og and og.get("content"):
        return normalize_spaces(str(og["content"]))

    if soup.title and soup.title.string:
        return normalize_spaces(soup.title.string)

    h1 = soup.find("h1")
    if h1:
        return normalize_spaces(h1.get_text(" ", strip=True))

    return ""


def parse_date_value(raw: str) -> date | None:
    value = raw.strip()
    if not value:
        return None

    # ISO / HTML5 datetime
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).date()
    except ValueError:
        pass

    head = re.split(r"[T\s]", value, maxsplit=1)[0]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(head, fmt).date()
        except ValueError:
            continue

    # Dates written in text, e.g. "5 августа 2026", "5 Aug 2026".
    months_ru = {
        "января": 1,
        "янв.": 1,
        "янв": 1,
        "февраля": 2,
        "февр.": 2,
        "фев": 2,
        "марта": 3,
        "март.": 3,
        "мар": 3,
        "апреля": 4,
        "апр.": 4,
        "апр": 4,
        "мая": 5,
        "май.": 5,
        "май": 5,
        "июня": 6,
        "июн.": 6,
        "июн": 6,
        "июля": 7,
        "июл.": 7,
        "июл": 7,
        "августа": 8,
        "авг.": 8,
        "авг": 8,
        "сентября": 9,
        "сен.": 9,
        "сен": 9,
        "сент": 9,
        "октября": 10,
        "окт.": 10,
        "окт": 10,
        "ноября": 11,
        "ноя.": 11,
        "ноя": 11,
        "декабря": 12,
        "дек.": 12,
        "дек": 12,
    }
    months_en = {
        "jan": 1,
        "january": 1,
        "feb": 2,
        "february": 2,
        "mar": 3,
        "march": 3,
        "apr": 4,
        "april": 4,
        "may": 5,
        "jun": 6,
        "june": 6,
        "jul": 7,
        "july": 7,
        "aug": 8,
        "august": 8,
        "sep": 9,
        "sept": 9,
        "september": 9,
        "oct": 10,
        "october": 10,
        "nov": 11,
        "november": 11,
        "dec": 12,
        "december": 12,
    }

    m = re.search(r"(?i)(\d{1,2})\s*([a-zа-яё\.]+)\s*(\d{4})", value)
    if m:
        day = int(m.group(1))
        mon_raw = m.group(2).strip().lower()
        year = int(m.group(3))
        mon = months_ru.get(mon_raw) or months_en.get(mon_raw)
        if mon:
            try:
                return date(year, mon, day)
            except ValueError:
                return None

    # RFC 2822
    try:
        return parsedate_to_datetime(value).date()
    except (TypeError, ValueError, IndexError):
        return None


ARTICLE_JSONLD_TYPES = {
    "newsarticle",
    "article",
    "blogposting",
    "reportagenewsarticle",
    "analysisnewsarticle",
    "opinionnewsarticle",
    "reviewnewsarticle",
}


def _jsonld_types(obj: dict) -> set[str]:
    raw = obj.get("@type")
    if isinstance(raw, str):
        return {raw.lower()}
    if isinstance(raw, list):
        return {str(x).lower() for x in raw}
    return set()


def extract_date_from_jsonld(soup: BeautifulSoup) -> date | None:
    """Prefer datePublished only from Article/NewsArticle JSON-LD nodes."""

    def walk(obj: object) -> date | None:
        if isinstance(obj, dict):
            types = _jsonld_types(obj)
            if types & ARTICLE_JSONLD_TYPES:
                for key in ("datePublished", "dateCreated", "uploadDate"):
                    value = obj.get(key)
                    if isinstance(value, str):
                        parsed = parse_date_value(value)
                        if parsed:
                            return parsed
            for value in obj.values():
                nested = walk(value)
                if nested:
                    return nested
        elif isinstance(obj, list):
            for item in obj:
                nested = walk(item)
                if nested:
                    return nested
        return None

    for script in soup.find_all("script", attrs={"type": lambda x: x and "ld+json" in x}):
        raw = script.string or script.get_text(strip=True)
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue
        parsed = walk(data)
        if parsed:
            return parsed
    return None


def page_looks_like_article(soup: BeautifulSoup, url: str) -> bool:
    """
    Company cards / listings often have WordPress article:* meta but are not news.
    """
    path = urlparse(url).path.lower()
    non_article_prefixes = (
        "/kompanii/",
        "/company/",
        "/companies/",
        "/category/",
        "/tag/",
        "/author/",
        "/page/",
        "/about/",
        "/contact",
        "/reklama",
        "/advert",
    )
    if any(path.startswith(p) or f"{p.rstrip('/')}/" in path + "/" for p in non_article_prefixes):
        # /category/ and /kompanii/ are not articles to report
        if path.startswith(("/kompanii/", "/company/", "/companies/", "/category/", "/tag/", "/author/")):
            return False

    og_type = soup.find("meta", property="og:type")
    if og_type and str(og_type.get("content", "")).lower() in {"website", "profile", "business.business"}:
        # Still allow if JSON-LD says NewsArticle
        pass

    for script in soup.find_all("script", attrs={"type": lambda x: x and "ld+json" in x}):
        raw = script.string or script.get_text(strip=True)
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue

        def has_article(obj: object) -> bool:
            if isinstance(obj, dict):
                if _jsonld_types(obj) & ARTICLE_JSONLD_TYPES:
                    return True
                return any(has_article(v) for v in obj.values())
            if isinstance(obj, list):
                return any(has_article(x) for x in obj)
            return False

        if has_article(data):
            return True

    if og_type and str(og_type.get("content", "")).lower() in {"article", "news"}:
        return True

    # Path heuristic: /YYYY/MM/slug or long news slug with date-ish segments
    if re.search(r"/\d{4}/\d{2}/", path):
        return True

    # Explicit non-article sections
    if path.startswith(("/kompanii/", "/company/", "/companies/")):
        return False

    # Default: treat as candidate article if it has an article body node
    for sel in (
        "article .entry-content",
        ".entry-content",
        ".post-content",
        ".td-post-content",
        "article",
    ):
        node = soup.select_one(sel)
        if node and len(node.get_text(" ", strip=True)) >= 200:
            # But reject short company cards wrapped in <article>
            if path.startswith(("/kompanii/", "/company/", "/companies/", "/category/")):
                return False
            return True
    return False


def extract_published_date(soup: BeautifulSoup, url: str = "") -> date | None:
    # Only trust dates on pages that look like articles.
    if url and not page_looks_like_article(soup, url):
        return None

    jsonld_date = extract_date_from_jsonld(soup)
    if jsonld_date:
        return jsonld_date

    meta_candidates: list[str] = []
    for prop in ("article:published_time", "og:published_time"):
        tag = soup.find("meta", property=prop)
        if tag and tag.get("content"):
            meta_candidates.append(str(tag["content"]))

    for name in (
        "pubdate",
        "publish-date",
        "publication_date",
        "DC.date",
        "dc.date",
        "sailthru.date",
    ):
        tag = soup.find("meta", attrs={"name": name})
        if tag and tag.get("content"):
            meta_candidates.append(str(tag["content"]))

    tag = soup.find(attrs={"itemprop": "datePublished"})
    if tag:
        content = tag.get("content") or tag.get("datetime") or tag.get_text(" ", strip=True)
        if content:
            meta_candidates.append(str(content))

    for time_tag in soup.find_all("time"):
        content = time_tag.get("datetime") or time_tag.get_text(" ", strip=True)
        if content:
            meta_candidates.append(str(content))

    for raw in meta_candidates:
        parsed = parse_date_value(raw)
        if parsed:
            return parsed
    return None


def extract_main_text(soup: BeautifulSoup) -> str:
    """
    Search phrases only in main content — not header/nav/footer/sidebar chrome
    (otherwise hits like 'Срочные новости – в нашем Telegram' leak into reports).
    """
    work = BeautifulSoup(str(soup), "lxml")
    for tag in work(
        ["script", "style", "noscript", "svg", "template", "nav", "header", "footer", "aside"]
    ):
        tag.decompose()

    junk_re = re.compile(
        r"(sidebar|widget|menu|nav|footer|header|share|social|telegram|subscribe|banner|advert|rekla)",
        re.I,
    )
    for tag in list(work.find_all(True)):
        if not getattr(tag, "attrs", None):
            continue
        cid = " ".join(tag.get("class") or [])
        tid = str(tag.get("id") or "")
        if junk_re.search(f"{cid} {tid}"):
            tag.decompose()

    for sel in (
        "article .entry-content",
        ".entry-content",
        ".post-content",
        ".td-post-content",
        ".article-body",
        ".post-body",
        "article",
        "main",
        "[role=main]",
        "#content",
    ):
        node = work.select_one(sel)
        if node:
            text = node.get_text(separator=" ", strip=True)
            if len(text) >= 80:
                return text[:MAX_PAGE_CHARS]

    text = work.get_text(separator=" ", strip=True)
    return text[:MAX_PAGE_CHARS]


NON_ARTICLE_PATH_PREFIXES = (
    "/kompanii/",
    "/company/",
    "/companies/",
    "/category/",
    "/tag/",
    "/author/",
    "/page/",
    "/wp-admin/",
    "/wp-login",
    "/about/",
    "/contact",
    "/reklama",
    "/advert",
    "/feed",
    "/search",
    "/login",
    "/register",
)


def is_probable_article_url(url: str) -> bool:
    path = urlparse(url).path.lower().rstrip("/") + "/"
    if path == "/":
        return False
    for prefix in NON_ARTICLE_PATH_PREFIXES:
        if path.startswith(prefix) or prefix.rstrip("/") == path.rstrip("/"):
            return False
    # pagination / query junk
    if re.search(r"/page/\d+/?", path):
        return False
    return True


def extract_page(html: str, base_url: str) -> tuple[str, date | None, str, list[str]]:
    """Return (title, published_date, main_text, same-host article-like links)."""
    soup = BeautifulSoup(html, "lxml")
    title = extract_title(soup)
    published = extract_published_date(soup, base_url)
    text = extract_main_text(soup)

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
        if not is_probable_article_url(clean):
            continue
        if clean not in seen:
            seen.add(clean)
            links.append(clean)
    return title, published, text, links


def phrase_present(text: str, phrase: str) -> bool:
    """
    Exact phrase match (case-insensitive), allowing flexible whitespace.
    One line in words.txt = one phrase, e.g. "audi выпустили новую А6".
    """
    needle = normalize_spaces(phrase).casefold()
    haystack = normalize_spaces(text).casefold()
    if not needle:
        return False
    return needle in haystack


def passes_date_filter(
    published: date | None,
    min_date: date | None,
    max_date: date | None,
) -> bool:
    """
    Keep if min_date <= published <= max_date.
    Unknown publish date — keep (include in report).
    """
    if published is None:
        return True
    if min_date is not None and published < min_date:
        return False
    if max_date is not None and published > max_date:
        return False
    return True


def scan_page(
    url: str,
    phrases: list[str],
    session: requests.Session,
    scanned_at: datetime,
    settings: Settings,
    auth: tuple[str, str] | None,
) -> list[Hit]:
    html = fetch_html(url, session, auth=auth)
    soup = BeautifulSoup(html, "lxml")

    # Seeds may be categories; expanded company cards must not become report rows.
    if not is_probable_article_url(url) and not page_looks_like_article(soup, url):
        return []

    title = extract_title(soup)
    published = extract_published_date(soup, url)
    text = extract_main_text(soup)

    if not passes_date_filter(
        published,
        settings.article_date_not_older_than,
        settings.article_date_not_later_than,
    ):
        return []

    published_str = published.isoformat() if published else ""
    hits: list[Hit] = []
    for phrase in phrases:
        if phrase_present(text, phrase):
            hits.append(
                Hit(
                    phrase=phrase,
                    url=url,
                    published_at=published_str,
                    title=title,
                    scanned_at=scanned_at,
                )
            )
    return hits


MAX_EXPAND_LINKS = 80
MAX_SCAN_URLS = 250


def site_search_urls(origin: str, phrase: str) -> list[str]:
    """Common on-site search endpoints (helps find older articles not on homepage)."""
    q = quote(phrase)
    base = origin.rstrip("/")
    return [
        f"{base}/search/?q={q}",
        f"{base}/?s={q}",
        f"{base}/search?query={q}",
        f"{base}/search?text={q}",
    ]


def collect_urls_to_scan(
    seed_urls: Iterable[str],
    phrases: list[str],
    session: requests.Session,
    auth: tuple[str, str] | None,
) -> list[str]:
    """
    Build scan list from seeds:
    1) seed URLs themselves
    2) article links from listing pages (1 hop)
    3) site search results for each phrase (finds older articles)
    """
    result: list[str] = []
    seen: set[str] = set()

    def add(url: str) -> None:
        if url in seen or len(result) >= MAX_SCAN_URLS:
            return
        seen.add(url)
        result.append(url)

    def add_links_from(page_url: str) -> None:
        try:
            html = fetch_html(page_url, session, auth=auth)
            _, _, _, links = extract_page(html, page_url)
            for link in links[:MAX_EXPAND_LINKS]:
                add(link)
        except AuthRequiredError as exc:
            print(f"  [skip auth] cannot expand {page_url}: {exc}")
        except requests.RequestException as exc:
            print(f"  [warn] cannot expand {page_url}: {exc}")

    origins: set[str] = set()
    for seed in seed_urls:
        url = normalize_url(seed)
        add(url)
        parsed = urlparse(url)
        origins.add(f"{parsed.scheme}://{parsed.netloc}")

        path = parsed.path.rstrip("/")
        if path.count("/") <= 2:
            add_links_from(url)

    for origin in sorted(origins):
        for phrase in phrases:
            for search_url in site_search_urls(origin, phrase):
                if len(result) >= MAX_SCAN_URLS:
                    break
                print(f"  search: {search_url}")
                add_links_from(search_url)

    return result


def write_report(hits: list[Hit], reports_dir: Path, generated_at: datetime) -> Path:
    reports_dir.mkdir(parents=True, exist_ok=True)
    stamp = generated_at.strftime("%Y-%m-%d_%H-%M-%S")
    out_path = reports_dir / f"media-monitor-{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Hits"

    headers = (
        "слово / фраза",
        "ссылка на страницу",
        "дата выхода статьи",
        "название статьи",
        "дата время скана",
    )
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    for hit in hits:
        ws.append(
            (
                hit.phrase,
                hit.url,
                hit.published_at,
                hit.title,
                hit.scanned_at.strftime("%Y-%m-%d %H:%M:%S"),
            )
        )

    for row in ws.iter_rows(min_row=2, max_col=2, max_row=ws.max_row):
        cell = row[1]
        if cell.value:
            cell.hyperlink = str(cell.value)
            cell.style = "Hyperlink"

    widths = (36, 70, 20, 50, 22)
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

        # Prefer user-provided `setting.txt` (singular), but if only `settings.txt` exists,
        # copy it to `setting.txt` so their edits are preserved.
        setting_singular = root / "setting.txt"
        settings_plural = root / "settings.txt"
        if setting_singular.is_file():
            settings_path = setting_singular
        elif settings_plural.is_file():
            shutil.copy2(settings_plural, setting_singular)
            print("Copied settings.txt -> setting.txt (preserve your values).")
            settings_path = setting_singular
        else:
            settings_path = ensure_local_config(root, "setting.txt")
        sites = load_lines(sites_path)
        phrases = load_lines(words_path)
        settings = load_settings(settings_path)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Sites: {sites_path.name}")
    print(f"Words: {words_path.name}")
    print(f"Settings: {settings_path.name}")
    older = settings.article_date_not_older_than
    later = settings.article_date_not_later_than
    print(
        "Filter dates: "
        f"not older than {older.isoformat() if older else '(off)'}, "
        f"not later than {later.isoformat()} "
        "(unknown publish dates are kept)"
    )

    if not sites:
        print("ERROR: sites.txt is empty (no URLs).")
        return 1
    if not phrases:
        print("ERROR: words.txt is empty (no keywords/phrases).")
        return 1

    print(f"Loaded {len(sites)} site(s), {len(phrases)} phrase(s).")
    auth = prompt_credentials(settings.auth_timeout_seconds)

    scanned_at = datetime.now()
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru,en;q=0.8"})

    print("Collecting pages…")
    urls = collect_urls_to_scan(sites, phrases, session, auth)
    print(f"Will scan {len(urls)} page(s).")

    hits: list[Hit] = []
    errors = 0
    skipped_auth = 0
    for i, url in enumerate(urls, start=1):
        print(f"[{i}/{len(urls)}] {url}")
        try:
            page_hits = scan_page(url, phrases, session, scanned_at, settings, auth)
            if page_hits:
                print(f"  → {len(page_hits)} hit(s)")
                hits.extend(page_hits)
        except AuthRequiredError as exc:
            skipped_auth += 1
            print(f"  [skip auth] {exc}")
        except requests.RequestException as exc:
            errors += 1
            print(f"  [error] {exc}")

    report_path = write_report(hits, reports_dir, datetime.now())
    print()
    print(
        f"Done. Hits: {len(hits)}. Errors: {errors}. "
        f"Skipped (auth): {skipped_auth}."
    )
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
