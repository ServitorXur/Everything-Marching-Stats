"""
Music For All - Marching Band Recap PDF Downloader
====================================================
Downloads all recap PDFs from https://marching.musicforall.org
for competition years 1976 through 2025.

Requirements:
    pip install requests beautifulsoup4

Usage:
    python download_musicforall_recaps.py

PDFs are saved to: ./musicforall_recaps/<year>/<competition>/filename.pdf
A log file (download_log.txt) is created with full details.
"""

import re
import time
import logging
import requests
from pathlib import Path
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup

# ── Configuration ────────────────────────────────────────────────────────────
BASE_URL    = "https://marching.musicforall.org"
BASE_DOMAIN = "marching.musicforall.org"
START_YEAR  = 1976
END_YEAR    = 2025
OUTPUT_DIR  = Path("musicforall_recaps")
LOG_FILE    = "download_log.txt"
DELAY       = 1.2        # seconds between requests (be polite)
TIMEOUT     = 30
MAX_RETRIES = 3
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; MusicForAllRecapDownloader/1.0)"
})


# ── Helpers ──────────────────────────────────────────────────────────────────

def safe_dirname(text: str) -> str:
    """Convert arbitrary text to a safe Windows/Mac/Linux directory name."""
    text = re.sub(r'[\\/*?:"<>|.]+', "", text)
    text = re.sub(r"\s+", "_", text.strip()).strip("_")
    return text[:100] or "unnamed"


def is_valid_subpage(url: str) -> bool:
    """
    Return True only if the URL is a real deep sub-page on the site.
    Rejects: external domains, root (/), single-segment paths (/result/).
    """
    parsed = urlparse(url)
    if parsed.netloc != BASE_DOMAIN:
        return False
    path = parsed.path.rstrip("/")
    if path in ("", "/"):
        return False
    segments = [s for s in path.split("/") if s]
    return len(segments) >= 2


def fetch(url: str) -> requests.Response | None:
    """GET a URL with retries; return Response or None on failure."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = SESSION.get(url, timeout=TIMEOUT, allow_redirects=True)
            if resp.status_code == 200:
                return resp
            log.warning("HTTP %s for %s (attempt %d)", resp.status_code, url, attempt)
        except requests.RequestException as exc:
            log.warning("Request error %s (attempt %d): %s", url, attempt, exc)
        if attempt < MAX_RETRIES:
            time.sleep(DELAY * attempt)
    return None


def download_pdf(pdf_url: str, dest_path: Path) -> bool:
    """Download a single PDF. Returns True on success."""
    if dest_path.exists():
        log.info("    SKIP (exists): %s", dest_path.name)
        return True
    try:
        dest_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        log.error("    mkdir failed for %s: %s", dest_path.parent, exc)
        return False

    resp = fetch(pdf_url)
    if resp is None:
        log.error("    FAIL: %s", pdf_url)
        return False

    ct = resp.headers.get("Content-Type", "")
    if "pdf" not in ct.lower() and not pdf_url.lower().endswith(".pdf"):
        log.warning("    Unexpected Content-Type '%s' — %s", ct, pdf_url)

    try:
        dest_path.write_bytes(resp.content)
        log.info("    SAVED (%d KB): %s", len(resp.content) // 1024, dest_path)
        return True
    except OSError as exc:
        log.error("    Write failed %s: %s", dest_path, exc)
        return False


def find_pdf_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    seen, results = set(), []
    for tag in soup.find_all("a", href=True):
        href = tag["href"].strip()
        if href.lower().endswith(".pdf"):
            full = urljoin(page_url, href)
            if full not in seen:
                seen.add(full)
                results.append(full)
    return results


def find_event_links(soup: BeautifulSoup, page_url: str, year: int) -> list[tuple[str, str]]:
    """
    Return (label, url) pairs for internal competition/event links.
    Only links that pass is_valid_subpage() are returned.
    """
    keywords = re.compile(
        r"recap|result|score|standing|caption|placement|event|competition"
        r"|show|championship|boa|grand.national|regional|super.regional",
        re.IGNORECASE,
    )
    seen, links = set(), []

    for tag in soup.find_all("a", href=True):
        href  = tag["href"].strip()
        label = tag.get_text(" ", strip=True)
        full  = urljoin(page_url, href)

        if full in seen or not is_valid_subpage(full):
            continue
        if keywords.search(href) or keywords.search(label) or str(year) in href:
            seen.add(full)
            links.append((label or href, full))

    # Fallback: any valid deep sub-page if nothing keyword-matched
    if not links:
        for tag in soup.find_all("a", href=True):
            href  = tag["href"].strip()
            label = tag.get_text(" ", strip=True)
            full  = urljoin(page_url, href)
            if full not in seen and is_valid_subpage(full):
                seen.add(full)
                links.append((label or href, full))

    return links


def process_competition_page(url: str, dest_dir: Path, visited: set) -> int:
    """
    Download PDFs from a competition page and follow one level of
    recap/result sub-links. Returns number of PDFs saved.
    """
    if url in visited:
        return 0
    visited.add(url)

    log.info("  → %s", url)
    resp = fetch(url)
    if resp is None:
        return 0
    time.sleep(DELAY)

    soup = BeautifulSoup(resp.text, "html.parser")
    downloaded = 0

    # PDFs directly on this page
    for pdf_url in find_pdf_links(soup, url):
        fname = Path(urlparse(pdf_url).path).name
        if fname:
            if download_pdf(pdf_url, dest_dir / fname):
                downloaded += 1
            time.sleep(DELAY)

    # One level of recap sub-links
    sub_kw = re.compile(
        r"recap|result|score|standing|caption|placement", re.IGNORECASE
    )
    for tag in soup.find_all("a", href=True):
        href  = tag["href"].strip()
        label = tag.get_text(" ", strip=True)
        full  = urljoin(url, href)

        if full in visited or not is_valid_subpage(full):
            continue
        if not (sub_kw.search(href) or sub_kw.search(label)):
            continue

        visited.add(full)
        log.info("    Sub [%s]: %s", label[:60], full)
        sub_resp = fetch(full)
        if sub_resp is None:
            continue
        time.sleep(DELAY)
        sub_soup = BeautifulSoup(sub_resp.text, "html.parser")
        for pdf_url in find_pdf_links(sub_soup, full):
            fname = Path(urlparse(pdf_url).path).name
            if fname:
                if download_pdf(pdf_url, dest_dir / fname):
                    downloaded += 1
                time.sleep(DELAY)

    return downloaded


def process_year(year: int, global_visited: set) -> int:
    """Scrape all competitions for a single year. Returns total PDFs downloaded."""
    year_url = f"{BASE_URL}/competition-year/{year}/"
    log.info("=" * 60)
    log.info("YEAR %d  →  %s", year, year_url)

    resp = fetch(year_url)
    if resp is None:
        log.warning("  Year page not found, skipping.")
        return 0
    time.sleep(DELAY)
    global_visited.add(year_url)

    soup     = BeautifulSoup(resp.text, "html.parser")
    year_dir = OUTPUT_DIR / str(year)
    total    = 0

    # PDFs directly on the year index page
    direct_pdfs = find_pdf_links(soup, year_url)
    if direct_pdfs:
        index_dir = year_dir / "_index"
        index_dir.mkdir(parents=True, exist_ok=True)
        for pdf_url in direct_pdfs:
            fname = Path(urlparse(pdf_url).path).name
            if fname and download_pdf(pdf_url, index_dir / fname):
                total += 1
            time.sleep(DELAY)

    # All competition/event pages linked from the year index
    event_links = find_event_links(soup, year_url, year)
    log.info("  Found %d event links for %d", len(event_links), year)

    for label, url in event_links:
        if url in global_visited:
            continue
        dirname = safe_dirname(label) or safe_dirname(
            urlparse(url).path.strip("/").split("/")[-1]
        ) or "unnamed"
        comp_dir = year_dir / dirname
        comp_dir.mkdir(parents=True, exist_ok=True)
        log.info("  [%s]", label[:80])
        total += process_competition_page(url, comp_dir, global_visited)

    return total


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    grand_total    = 0
    global_visited: set[str] = set()

    for year in range(START_YEAR, END_YEAR + 1):
        count        = process_year(year, global_visited)
        grand_total += count
        log.info("  → %d PDF(s) downloaded for %d\n", count, year)

    log.info("=" * 60)
    log.info("DONE.  Total PDFs downloaded: %d", grand_total)
    log.info("Saved to: %s", OUTPUT_DIR.resolve())
    log.info("Log:      %s", Path(LOG_FILE).resolve())


if __name__ == "__main__":
    main()
