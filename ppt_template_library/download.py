from __future__ import annotations

import hashlib
import re
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urljoin

import requests


def is_supported_template_url(url: str) -> bool:
    lowered = url.lower()
    return any(
        lowered.endswith(ext)
        for ext in (".pptx", ".potx", ".ppt", ".odp", ".zip", ".html", ".md", ".qmd")
    ) or "github.com" in lowered or "slides" in lowered or "template" in lowered


def safe_filename_from_url(url: str, fallback: str = "template") -> str:
    digest = hashlib.sha256(url.encode("utf-8")).hexdigest()[:12]
    cleaned = url.split("?")[0].rstrip("/")
    stem = re.sub(r"[^a-zA-Z0-9._-]+", "_", Path(cleaned).name) or fallback
    suffix = Path(cleaned).suffix[:8]
    return f"{stem[:48]}-{digest}{suffix}"


def download_file(url: str, out_dir: str | Path, timeout: int = 30) -> Path:
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    filename = safe_filename_from_url(url)
    target = out_path / filename
    with requests.get(url, stream=True, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"}) as response:
        response.raise_for_status()
        with target.open("wb") as file_obj:
            for chunk in response.iter_content(chunk_size=1024 * 64):
                if chunk:
                    file_obj.write(chunk)
    return target


class _LinkCollector(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        href = dict(attrs).get("href")
        if href:
            self.links.append(href)


def extract_candidate_links(page_url: str, html: str) -> list[str]:
    parser = _LinkCollector()
    parser.feed(html)
    links: list[str] = []
    for href in parser.links:
        lowered = href.lower()
        if any(ext in lowered for ext in [".pptx", ".potx", ".ppt", ".odp", ".zip", ".md", ".qmd", "releases/download", "/raw/"]):
            links.append(href)
    if not links and "github.com" in page_url.lower():
        for href in parser.links:
            if "/blob/" in href or "/tree/" in href:
                links.append(href)
    return sorted(set(links))


def discover_template_links(url: str, timeout: int = 30) -> list[str]:
    response = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    base = response.url
    links: list[str] = []
    for link in extract_candidate_links(base, response.text):
        if link.startswith("http"):
            links.append(link)
        elif link.startswith("/"):
            links.append(urljoin(base, link))
    return sorted(set(links))


def hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def file_hash(path: str | Path) -> str:
    file_path = Path(path)
    sha = hashlib.sha256()
    with file_path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            sha.update(chunk)
    return sha.hexdigest()


def directory_hash(path: str | Path) -> str:
    root = Path(path)
    sha = hashlib.sha256()
    for file_path in sorted(candidate for candidate in root.rglob("*") if candidate.is_file() and ".git" not in candidate.parts):
        sha.update(str(file_path.relative_to(root)).encode("utf-8"))
        sha.update(str(file_path.stat().st_size).encode("utf-8"))
    return sha.hexdigest()
