from __future__ import annotations

import argparse
import html
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote, urlsplit

from bs4 import BeautifulSoup, UnicodeDammit
from pylibmspack import ChmArchive
from pypdf import PdfReader, PdfWriter

VERSION = "1.0.0"
DEFAULT_BROWSER_CANDIDATES = ("chromium-browser", "chromium", "google-chrome")
DEFAULT_PAPER_SIZE = "A4"
IGNORED_URL_SCHEMES = ("javascript:", "mailto:", "data:", "tel:", "about:")


@dataclass
class TocNode:
    rel: str
    title: str | None = None
    children: list["TocNode"] = field(default_factory=list)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="chm2pdf",
        description="Convert CHM files into PDF.",
    )
    parser.add_argument("input", type=Path, help="Input CHM file")
    parser.add_argument("output", nargs="?", type=Path, help="Output PDF file")
    parser.add_argument("--extract-only", action="store_true", help="Extract the CHM and stop")
    parser.add_argument("--dontextract", action="store_true", help="Reuse previously extracted files")
    parser.add_argument("--keep-temp", action="store_true", help="Keep temporary build files")
    parser.add_argument("--work-dir", type=Path, help="Override the temporary work root")
    parser.add_argument("--title", help="Override the document title")
    parser.add_argument("--titlefile", help="Promote a specific extracted HTML file to the front")
    parser.add_argument("--paper-size", default=DEFAULT_PAPER_SIZE, help="CSS page size, e.g. A4 or Letter")
    parser.add_argument("--landscape", action="store_true", help="Render in landscape orientation")
    parser.add_argument("--inline-toc", action="store_true", help="Insert a generated table-of-contents page into the PDF")
    parser.add_argument(
        "--pdf-header-footer",
        action="store_true",
        help="Keep Chromium's printed header/footer block (date, title, URL, page numbers)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Print progress information")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    return parser


def log(enabled: bool, message: str) -> None:
    if enabled:
        print(message)


def normalize_archive_path(value: str) -> str:
    path = value.strip().strip('"\'').replace("\\", "/")
    path = path.replace("\x00", "")
    lower_path = path.lower()
    if lower_path.startswith(IGNORED_URL_SCHEMES):
        return ""
    if lower_path.startswith("mk:@msitstore:"):
        path = path[len("mk:@MSITStore:") :]
    elif lower_path.startswith("ms-its:"):
        path = path[len("ms-its:") :]
    elif lower_path.startswith("its:"):
        path = path[len("its:") :]
    if "::" in path:
        path = path.split("::", 1)[1]
    path = path.split("#", 1)[0].split("?", 1)[0]
    path = unquote(path)
    if not path:
        return ""
    if path.startswith("file://"):
        return ""
    if re.match(r"^[A-Za-z]:/", path):
        return ""
    if path.startswith("/"):
        path = path[1:]
    normalized = posixpath.normpath(path).lower()
    return "" if normalized in {"", "."} else normalized


def resolve_relative_path(current_rel: str, target: str) -> str:
    stripped = target.strip()
    lower_stripped = stripped.lower()
    if lower_stripped.startswith(IGNORED_URL_SCHEMES):
        return ""
    if "::" in stripped or lower_stripped.startswith(("mk:@msitstore:", "ms-its:", "its:")):
        return normalize_archive_path(stripped)
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc:
        return ""
    path = unquote(parsed.path).replace("\\", "/")
    if not path:
        return ""
    if path.startswith("/"):
        return normalize_archive_path(path)
    joined = posixpath.join(posixpath.dirname(current_rel), path)
    return posixpath.normpath(joined).lower()


def safe_rmtree(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)


def ensure_browser(user_supplied: str | None = None) -> str:
    candidates = [user_supplied] if user_supplied else list(DEFAULT_BROWSER_CANDIDATES)
    for candidate in candidates:
        if candidate and shutil.which(candidate):
            return candidate
    raise SystemExit(
        "No Chromium-compatible browser found. Install `chromium-browser` or put one on PATH."
    )


def extract_chm(input_file: Path, orig_dir: Path, verbose: bool) -> None:
    log(verbose, f"Extracting {input_file} to {orig_dir}")
    safe_rmtree(orig_dir)
    orig_dir.mkdir(parents=True, exist_ok=True)
    archive = ChmArchive(str(input_file))
    archive.extract_all(str(orig_dir))


def read_text_with_fallback(path: Path, encodings: tuple[str, ...]) -> str:
    data = path.read_bytes()
    dammit = UnicodeDammit(data, known_definite_encodings=list(encodings), is_html=True)
    if dammit.unicode_markup:
        return dammit.unicode_markup
    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode(encodings[-1], errors="replace")


def read_html_text(path: Path) -> str:
    return read_text_with_fallback(path, ("utf-8-sig", "utf-8", "cp1252", "latin-1"))


def read_hhc_text(path: Path) -> str:
    return read_text_with_fallback(path, ("utf-8-sig", "utf-8", "cp1252", "latin-1"))


def list_archive_html_files(input_file: Path) -> list[str]:
    archive = ChmArchive(str(input_file))
    ordered: list[str] = []
    seen: set[str] = set()
    for file_info in archive.files():
        rel = normalize_archive_path(str(file_info.get("name", "")))
        if not rel or rel in seen:
            continue
        if Path(rel).suffix.lower() not in {".htm", ".html"}:
            continue
        ordered.append(rel)
        seen.add(rel)
    return ordered


def build_file_index(root: Path) -> dict[str, Path]:
    index: dict[str, Path] = {}
    for path in root.rglob("*"):
        if path.is_file():
            rel = path.relative_to(root).as_posix()
            index[normalize_archive_path(rel)] = path
    return index


class HhcParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.root: list[TocNode] = []
        self.current_children: list[list[TocNode]] = [self.root]
        self.last_node_at_level: list[TocNode | None] = [None]
        self.in_object = False
        self.object_is_sitemap = False
        self.current_title: str | None = None
        self.current_local: str | None = None
        self.seen_root_ul = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_map = {key.lower(): value for key, value in attrs if key}
        if tag == "ul":
            if not self.seen_root_ul:
                self.seen_root_ul = True
                return
            parent = self.last_node_at_level[-1]
            if parent is None:
                return
            self.current_children.append(parent.children)
            self.last_node_at_level.append(None)
            return

        if tag == "object":
            self.in_object = True
            self.object_is_sitemap = (attrs_map.get("type") or "").lower() == "text/sitemap"
            self.current_title = None
            self.current_local = None
            return

        if tag == "param" and self.in_object and self.object_is_sitemap:
            name = (attrs_map.get("name") or "").strip().lower()
            value = (attrs_map.get("value") or "").strip()
            if name == "name":
                self.current_title = value or None
            elif name == "local":
                self.current_local = normalize_archive_path(value)

    def handle_endtag(self, tag: str) -> None:
        if tag == "object":
            if self.object_is_sitemap and self.current_local:
                node = TocNode(rel=self.current_local, title=self.current_title)
                self.current_children[-1].append(node)
                self.last_node_at_level[-1] = node
            self.in_object = False
            self.object_is_sitemap = False
            self.current_title = None
            self.current_local = None
            return

        if tag == "ul" and len(self.current_children) > 1:
            self.current_children.pop()
            self.last_node_at_level.pop()


def parse_toc_tree(root: Path) -> list[TocNode]:
    best_tree: list[TocNode] = []
    for toc_file in sorted(root.rglob("*.hhc")):
        parser = HhcParser()
        content = read_hhc_text(toc_file)
        parser.feed(content)
        parsed_tree = parser.root
        if count_toc_nodes(parsed_tree) == 0:
            parsed_tree = parse_toc_tree_fallback(content)
        if count_toc_nodes(parsed_tree) > count_toc_nodes(best_tree):
            best_tree = parsed_tree
    return best_tree


def parse_toc_tree_fallback(content: str) -> list[TocNode]:
    soup = BeautifulSoup(content, "html.parser")
    nodes: list[TocNode] = []
    seen: set[str] = set()
    for obj in soup.find_all("object"):
        if (obj.get("type") or "").lower() != "text/sitemap":
            continue
        local = None
        title = None
        for param in obj.find_all("param"):
            name = (param.get("name") or "").strip().lower()
            value = (param.get("value") or "").strip()
            if name == "local":
                local = normalize_archive_path(value)
            elif name == "name":
                title = value or None
        if local and local not in seen:
            nodes.append(TocNode(rel=local, title=title))
            seen.add(local)
    return nodes


def count_toc_nodes(nodes: Iterable[TocNode]) -> int:
    total = 0
    for node in nodes:
        total += 1 + count_toc_nodes(node.children)
    return total


def flatten_toc_tree(nodes: Iterable[TocNode]) -> list[tuple[str, str | None]]:
    entries: list[tuple[str, str | None]] = []
    for node in nodes:
        entries.append((node.rel, node.title))
        entries.extend(flatten_toc_tree(node.children))
    return entries


def discover_cover_page(file_index: dict[str, Path], input_stem: str) -> tuple[str, str | None] | None:
    preferred_names = [
        f"{input_stem.lower()}_default.htm",
        f"{input_stem.lower()}_default.html",
        "default.htm",
        "default.html",
        "index.htm",
        "index.html",
    ]
    for rel in preferred_names:
        path = file_index.get(rel)
        if path and path.suffix.lower() in {".htm", ".html"}:
            soup = BeautifulSoup(read_html_text(path), "html.parser")
            return rel, choose_document_title(soup, path.stem)

    for rel, path in file_index.items():
        if path.suffix.lower() not in {".htm", ".html"}:
            continue
        body = BeautifulSoup(read_html_text(path), "html.parser").body
        classes = body.get("class", []) if body else []
        if "first" in classes:
            soup = BeautifulSoup(read_html_text(path), "html.parser")
            return rel, choose_document_title(soup, path.stem)

    return None


def discover_html_order(
    root: Path,
    file_index: dict[str, Path],
    input_stem: str,
    archive_html_order: list[str],
) -> tuple[list[tuple[str, str | None]], list[TocNode]]:
    toc_tree = parse_toc_tree(root)
    toc_entries = flatten_toc_tree(toc_tree)
    ordered: list[tuple[str, str | None]] = []
    seen: set[str] = set()

    cover_page = discover_cover_page(file_index, input_stem)
    if cover_page and cover_page[0] not in seen:
        ordered.append(cover_page)
        seen.add(cover_page[0])

    for rel, title in toc_entries:
        if rel in file_index and rel not in seen and file_index[rel].suffix.lower() in {".htm", ".html"}:
            ordered.append((rel, title))
            seen.add(rel)

    for rel in archive_html_order:
        if rel not in seen and rel in file_index and file_index[rel].suffix.lower() in {".htm", ".html"}:
            ordered.append((rel, None))
            seen.add(rel)

    for rel in sorted(file_index):
        if rel not in seen and file_index[rel].suffix.lower() in {".htm", ".html"}:
            ordered.append((rel, None))

    return ordered, toc_tree


def prefix_fragment(section_id: str, fragment: str) -> str:
    return f"{section_id}__{fragment}"


def rewrite_named_targets(soup: BeautifulSoup, section_id: str) -> None:
    for tag in soup.find_all(attrs={"id": True}):
        tag["id"] = prefix_fragment(section_id, str(tag["id"]))
    for tag in soup.find_all(attrs={"name": True}):
        tag["name"] = prefix_fragment(section_id, str(tag["name"]))


def rewrite_asset_links(
    soup: BeautifulSoup,
    current_rel: str,
    file_index: dict[str, Path],
) -> None:
    for tag in soup.find_all(src=True):
        rel = resolve_relative_path(current_rel, str(tag["src"]))
        if rel and rel in file_index:
            tag["src"] = file_index[rel].resolve().as_uri()

    for tag in soup.find_all(href=True):
        if tag.name == "a":
            continue
        href = str(tag["href"])
        parsed = urlsplit(href)
        if parsed.scheme or parsed.netloc or href.startswith("#"):
            continue
        rel = resolve_relative_path(current_rel, href)
        if rel and rel in file_index:
            rebuilt = file_index[rel].resolve().as_uri()
            if parsed.fragment:
                rebuilt = f"{rebuilt}#{parsed.fragment}"
            tag["href"] = rebuilt

    for tag in soup.find_all(style=True):
        tag["style"] = rewrite_css_urls(str(tag["style"]), current_rel, file_index)


def rewrite_anchors(
    soup: BeautifulSoup,
    current_rel: str,
    current_section_id: str,
    page_targets: dict[str, str],
) -> None:
    for tag in soup.find_all("a", href=True):
        href = str(tag["href"])
        parsed = urlsplit(href)
        if parsed.scheme or parsed.netloc:
            continue
        if href.startswith("#"):
            fragment = href[1:]
            tag["href"] = f"#{prefix_fragment(current_section_id, fragment)}" if fragment else f"#{current_section_id}"
            continue

        rel = resolve_relative_path(current_rel, href)
        if not rel or rel not in page_targets:
            continue

        target = page_targets[rel]
        if parsed.fragment:
            tag["href"] = f"#{prefix_fragment(target, parsed.fragment)}"
        else:
            tag["href"] = f"#{target}"


def extract_head_assets(
    soup: BeautifulSoup,
    current_rel: str,
    file_index: dict[str, Path],
) -> tuple[list[str], list[str]]:
    stylesheet_links: list[str] = []
    inline_styles: list[str] = []
    head = soup.head
    if not head:
        return stylesheet_links, inline_styles

    for link in head.find_all("link", href=True):
        if (link.get("rel") or [""])[0].lower() != "stylesheet":
            continue
        rel = resolve_relative_path(current_rel, str(link["href"]))
        if rel and rel in file_index:
            stylesheet_links.append(file_index[rel].resolve().as_uri())

    for style in head.find_all("style"):
        if style.string:
            inline_styles.append(rewrite_css_urls(style.string, current_rel, file_index))

    return stylesheet_links, inline_styles


def extract_body_html(soup: BeautifulSoup) -> str:
    body = soup.body if soup.body else soup
    return "".join(str(child) for child in body.contents)


def rewrite_css_urls(css_text: str, current_rel: str, file_index: dict[str, Path]) -> str:
    def replace(match: re.Match[str]) -> str:
        raw = match.group("url").strip().strip('"\'')
        rel = resolve_relative_path(current_rel, raw)
        if rel and rel in file_index:
            return f'url("{file_index[rel].resolve().as_uri()}")'
        return match.group(0)

    return re.sub(r"url\((?P<url>[^)]+)\)", replace, css_text, flags=re.IGNORECASE)


def cleanup_page_soup(soup: BeautifulSoup) -> None:
    for tag in soup.find_all(["script", "noscript", "base", "iframe", "frame"]):
        tag.decompose()

    for meta in soup.find_all("meta"):
        http_equiv = (meta.get("http-equiv") or "").strip().lower()
        if http_equiv in {"refresh", "content-security-policy"}:
            meta.decompose()

    for tag in soup.find_all(True):
        for attr in list(tag.attrs):
            if attr.lower().startswith("on"):
                del tag.attrs[attr]


def choose_title(page_soup: BeautifulSoup, fallback: str) -> str:
    if page_soup.title and page_soup.title.string:
        title = page_soup.title.string.strip()
        if title:
            return title
    return fallback


def choose_document_title(page_soup: BeautifulSoup, fallback: str) -> str:
    book_name = page_soup.find("meta", attrs={"name": "BookName"})
    if book_name and book_name.get("content"):
        return str(book_name["content"]).strip()

    heading = page_soup.find("h1")
    if heading:
        heading_text = heading.get_text(" ", strip=True)
        if heading_text:
            return heading_text

    title = choose_title(page_soup, fallback)
    if ":" in title:
        left, right = (part.strip() for part in title.split(":", 1))
        if left and right and left == right:
            return left
    return title


def page_has_title_header(page_soup: BeautifulSoup, title: str) -> bool:
    heading = page_soup.find("h1")
    if not heading:
        return False
    heading_text = heading.get_text(" ", strip=True)
    return bool(heading_text) and heading_text == title


def page_has_native_heading(page_soup: BeautifulSoup) -> bool:
    heading = page_soup.find("h1")
    if not heading:
        return False
    return bool(heading.get_text(" ", strip=True))


def make_marker(section_id: str) -> str:
    return f"CHM2PDFMARKER-{section_id}"


def render_toc_nodes(nodes: Iterable[TocNode], page_targets: dict[str, str]) -> str:
    items: list[str] = []
    for node in nodes:
        href = f'#{page_targets[node.rel]}' if node.rel in page_targets else "#"
        label = html.escape(node.title or node.rel)
        children_html = render_toc_nodes(node.children, page_targets)
        if children_html:
            items.append(f'<li><a href="{href}">{label}</a>{children_html}</li>')
        else:
            items.append(f'<li><a href="{href}">{label}</a></li>')
    if not items:
        return ""
    return f"<ul>{''.join(items)}</ul>"


def toc_contains_rel(nodes: Iterable[TocNode], rel: str) -> bool:
    for node in nodes:
        if node.rel == rel or toc_contains_rel(node.children, rel):
            return True
    return False


def build_effective_toc_nodes(
    toc_tree: list[TocNode],
    ordered_pages: list[tuple[str, str | None]],
) -> list[TocNode]:
    toc_nodes = list(toc_tree)
    first_rel, first_title = ordered_pages[0]
    if not toc_contains_rel(toc_nodes, first_rel):
        toc_nodes = [TocNode(rel=first_rel, title=first_title), *toc_nodes]
    return toc_nodes


def build_bundle_html(
    ordered_pages: list[tuple[str, str | None]],
    toc_tree: list[TocNode],
    file_index: dict[str, Path],
    title: str,
    paper_size: str,
    landscape: bool,
    inline_toc: bool,
) -> tuple[str, dict[str, str], list[TocNode]]:
    page_targets = {rel: f"page-{index:04d}" for index, (rel, _) in enumerate(ordered_pages, start=1)}
    stylesheet_links: list[str] = []
    inline_styles: list[str] = []
    section_html: list[str] = []
    suppress_generated_title = False

    for index, (rel, toc_title) in enumerate(ordered_pages):
        section_id = page_targets[rel]
        page_path = file_index[rel]
        soup = BeautifulSoup(read_html_text(page_path), "html.parser")
        cleanup_page_soup(soup)
        if index == 0:
            suppress_generated_title = page_has_title_header(soup, title)
        rewrite_named_targets(soup, section_id)
        rewrite_asset_links(soup, rel, file_index)
        rewrite_anchors(soup, rel, section_id, page_targets)

        links, styles = extract_head_assets(soup, rel, file_index)
        stylesheet_links.extend(links)
        inline_styles.extend(styles)

        section_title = toc_title or choose_title(soup, page_path.stem)
        body_html = extract_body_html(soup)
        generated_heading_html = ""
        if not page_has_native_heading(soup):
            generated_heading_html = f'<h1 class="chapter-title">{html.escape(section_title)}</h1>'
        marker_html = f'<div class="bookmark-marker">{make_marker(section_id)}</div>'
        section_html.append(
            f'<section class="chapter" id="{section_id}">'
            f"{marker_html}"
            f"{generated_heading_html}"
            f"{body_html}"
            f"</section>"
        )

    orientation = " landscape" if landscape else ""
    unique_stylesheets = "\n".join(
        f'<link rel="stylesheet" href="{href}">' for href in dict.fromkeys(stylesheet_links)
    )
    combined_inline_styles = "\n".join(f"<style>{style}</style>" for style in dict.fromkeys(inline_styles))
    toc_nodes = build_effective_toc_nodes(toc_tree, ordered_pages)
    generated_toc_html = ""
    if inline_toc:
        generated_toc_html = render_toc_nodes(toc_nodes, page_targets)
        if not generated_toc_html:
            fallback_nodes = [TocNode(rel=rel, title=title_text) for rel, title_text in ordered_pages]
            generated_toc_html = render_toc_nodes(fallback_nodes, page_targets)

    generated_title_html = ""
    if not suppress_generated_title:
        generated_title_html = (
            '  <header class="generated-title">\n'
            f'    <h1>{html.escape(title)}</h1>\n'
            "  </header>\n"
        )
    generated_toc_block = ""
    if inline_toc and generated_toc_html:
        generated_toc_block = (
            '  <nav class="generated-toc">\n'
            "    <h2>Contents</h2>\n"
            f"    {generated_toc_html}\n"
            "  </nav>\n"
        )

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <title>{html.escape(title)}</title>
  <style>
    @page {{ size: {paper_size}{orientation}; margin: 14mm; }}
    body {{ font-family: sans-serif; color: #111; line-height: 1.45; }}
    img {{ max-width: 100%; height: auto; }}
    pre {{ white-space: pre-wrap; overflow-wrap: anywhere; }}
    .generated-title {{ margin: 0 0 2rem; }}
    .generated-toc {{ break-after: page; }}
    .generated-toc ul {{ padding-left: 1.25rem; }}
    .generated-toc li {{ margin: 0.2rem 0; }}
    .chapter {{ break-before: page; }}
    .chapter-title {{ margin: 0 0 1.5rem; font-size: 1.6rem; }}
    .bookmark-marker {{ color: transparent; font-size: 1px; line-height: 1; height: 1px; overflow: hidden; }}
  </style>
  {unique_stylesheets}
  {combined_inline_styles}
</head>
<body>
{generated_title_html}{generated_toc_block}  {''.join(section_html)}
</body>
</html>
""", page_targets, toc_nodes


def render_pdf(
    bundle_path: Path,
    output_path: Path,
    browser: str,
    verbose: bool,
    pdf_header_footer: bool,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    command = [
        browser,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--allow-file-access-from-files",
        f"--print-to-pdf={output_path}",
        bundle_path.resolve().as_uri(),
    ]
    if not pdf_header_footer:
        command.insert(-2, "--no-pdf-header-footer")
        command.insert(-2, "--print-to-pdf-no-header")
    log(verbose, f"Rendering PDF with {' '.join(command[:-1])} {bundle_path.resolve().as_uri()}")
    subprocess.run(command, check=True)


def locate_section_pages(pdf_path: Path, page_targets: dict[str, str]) -> dict[str, int]:
    reader = PdfReader(str(pdf_path))
    marker_to_rel = {make_marker(section_id): rel for rel, section_id in page_targets.items()}
    section_pages: dict[str, int] = {}

    for page_number, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        for marker, rel in marker_to_rel.items():
            if rel not in section_pages and marker in text:
                section_pages[rel] = page_number

    return section_pages


def add_outline_nodes(
    writer: PdfWriter,
    nodes: Iterable[TocNode],
    section_pages: dict[str, int],
    parent: object | None = None,
) -> None:
    for node in nodes:
        if node.rel not in section_pages:
            continue
        outline = writer.add_outline_item(node.title or node.rel, section_pages[node.rel], parent=parent)
        add_outline_nodes(writer, node.children, section_pages, parent=outline)


def inject_pdf_bookmarks(pdf_path: Path, toc_nodes: list[TocNode], page_targets: dict[str, str], verbose: bool) -> None:
    section_pages = locate_section_pages(pdf_path, page_targets)
    if not section_pages:
        log(verbose, "Skipping PDF bookmark injection: no section markers found in rendered PDF.")
        return

    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    add_outline_nodes(writer, toc_nodes, section_pages)

    temp_output = pdf_path.with_suffix(".bookmarks.pdf")
    with temp_output.open("wb") as file_obj:
        writer.write(file_obj)
    temp_output.replace(pdf_path)


def resolve_output_path(input_path: Path, output_path: Path | None) -> Path:
    if output_path is not None:
        return output_path
    return input_path.with_suffix(".pdf")


def reorder_titlefile(
    ordered_pages: list[tuple[str, str | None]],
    titlefile: str | None,
    file_index: dict[str, Path],
) -> list[tuple[str, str | None]]:
    if not titlefile:
        return ordered_pages
    rel = normalize_archive_path(titlefile)
    if rel not in file_index:
        raise SystemExit(f"titlefile not found in extracted CHM: {titlefile}")
    if file_index[rel].suffix.lower() not in {".htm", ".html"}:
        raise SystemExit(f"titlefile is not an HTML file: {titlefile}")

    reordered = [(rel, None)]
    reordered.extend((page_rel, page_title) for page_rel, page_title in ordered_pages if page_rel != rel)
    return reordered


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.extract_only and args.dontextract:
        raise SystemExit("Only one of --extract-only or --dontextract may be used.")

    input_path = args.input.resolve()
    if not input_path.exists():
        raise SystemExit(f"Input CHM not found: {input_path}")

    output_path = resolve_output_path(input_path, args.output.resolve() if args.output else None)
    base_dir = (args.work_dir.resolve() if args.work_dir else Path(tempfile.gettempdir()) / "chm2pdf")
    orig_dir = base_dir / "orig" / input_path.stem
    work_dir = base_dir / "work" / input_path.stem
    bundle_path = work_dir / "bundle.html"

    if args.dontextract:
        if not orig_dir.exists():
            raise SystemExit(f"Previously extracted directory not found: {orig_dir}")
        work_dir.mkdir(parents=True, exist_ok=True)
    else:
        safe_rmtree(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)
        extract_chm(input_path, orig_dir, args.verbose)

    if args.extract_only:
        print(orig_dir)
        return 0

    file_index = build_file_index(orig_dir)
    archive_html_order = list_archive_html_files(input_path)
    ordered_pages, toc_tree = discover_html_order(orig_dir, file_index, input_path.stem, archive_html_order)
    if not ordered_pages:
        raise SystemExit("No HTML files found in the extracted CHM.")

    ordered_pages = reorder_titlefile(ordered_pages, args.titlefile, file_index)
    document_title = args.title or ordered_pages[0][1] or input_path.stem
    bundle_html, page_targets, toc_nodes = build_bundle_html(
        ordered_pages=ordered_pages,
        toc_tree=toc_tree,
        file_index=file_index,
        title=document_title,
        paper_size=args.paper_size,
        landscape=args.landscape,
        inline_toc=args.inline_toc,
    )
    bundle_path.write_text(bundle_html, encoding="utf-8")

    browser = ensure_browser()
    render_pdf(bundle_path, output_path, browser, args.verbose, args.pdf_header_footer)
    inject_pdf_bookmarks(output_path, toc_nodes, page_targets, args.verbose)
    print(output_path)

    if not args.keep_temp:
        safe_rmtree(work_dir)

    return 0
