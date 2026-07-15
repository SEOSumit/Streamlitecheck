from __future__ import annotations

import csv
import html as html_lib
import re
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass
from html.parser import HTMLParser
from io import BytesIO, StringIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urljoin, urlsplit, urlunsplit
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/150.0.0.0 Safari/537.36"
)


@dataclass
class DataTable:
    name: str
    rows: list[list[object]]


class SitemapAnchorParser(HTMLParser):
    def __init__(self, base_url: str) -> None:
        super().__init__(convert_charrefs=True)
        self.base_url = base_url
        self.anchors: list[dict[str, str]] = []
        self._open_anchor: dict[str, object] | None = None
        self._ignored_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.casefold()
        attrs_dict = dict(attrs)
        if tag in {"script", "style", "noscript"}:
            self._ignored_depth += 1
            return
        if self._ignored_depth:
            return
        if tag == "base" and attrs_dict.get("href"):
            self.base_url = urljoin(self.base_url, str(attrs_dict["href"]))
        elif tag == "a":
            self._open_anchor = {
                "href": str(attrs_dict.get("href") or "").strip(),
                "text": [],
            }
        elif tag == "img" and self._open_anchor is not None and attrs_dict.get("alt"):
            self._open_anchor["text"].append(str(attrs_dict["alt"]))

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if tag in {"script", "style", "noscript"}:
            self._ignored_depth = max(0, self._ignored_depth - 1)
            return
        if self._ignored_depth:
            return
        if tag == "a" and self._open_anchor is not None:
            href = str(self._open_anchor["href"])
            if href and not href.casefold().startswith(("mailto:", "tel:", "javascript:")):
                absolute = urljoin(self.base_url, href)
                if absolute.startswith(("http://", "https://")):
                    parts = urlsplit(absolute)
                    absolute = urlunsplit((parts.scheme, parts.netloc, parts.path, parts.query, ""))
                    self.anchors.append(
                        {
                            "url": absolute,
                            "anchor": normalize_text("".join(self._open_anchor["text"])),
                        }
                    )
            self._open_anchor = None

    def handle_data(self, data: str) -> None:
        if not self._ignored_depth and self._open_anchor is not None:
            self._open_anchor["text"].append(data)


def normalize_text(value: object) -> str:
    text = html_lib.unescape(str(value or ""))
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_compare_url(value: str) -> str:
    parts = urlsplit(str(value or "").strip())
    scheme = parts.scheme.casefold()
    host = (parts.hostname or "").casefold()
    port = parts.port
    if port and not ((scheme == "http" and port == 80) or (scheme == "https" and port == 443)):
        host = f"{host}:{port}"
    path = re.sub(r"/{2,}", "/", unquote(parts.path or "/"))
    if path != "/":
        path = path.rstrip("/")
    return urlunsplit((scheme, host, path, "", ""))


def clean_url_without_parameters(value: str) -> str:
    parts = urlsplit(str(value).strip())
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))


def _is_url(value: object) -> bool:
    return isinstance(value, str) and value.strip().casefold().startswith(("http://", "https://"))


def _decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _csv_table(data: bytes, name: str) -> DataTable:
    text = _decode_text(data)
    try:
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t|")
        rows = [list(row) for row in csv.reader(StringIO(text), dialect)]
    except csv.Error:
        rows = [list(row) for row in csv.reader(StringIO(text))]
    return DataTable(name, rows)


def _xlsx_tables(data: bytes, name: str) -> list[DataTable]:
    wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
    return [
        DataTable(f"{name} — {ws.title}", [list(row) for row in ws.iter_rows(values_only=True)])
        for ws in wb.worksheets
    ]


def load_tables(data: bytes, filename: str) -> list[DataTable]:
    suffix = Path(filename).suffix.casefold()
    if suffix == ".xlsx":
        return _xlsx_tables(data, filename)
    if suffix in {".csv", ".tsv", ".txt"}:
        return [_csv_table(data, filename)]
    if suffix == ".zip":
        tables: list[DataTable] = []
        with zipfile.ZipFile(BytesIO(data)) as archive:
            for member in archive.infolist():
                if member.is_dir() or member.file_size > 50_000_000:
                    continue
                member_data = archive.read(member)
                member_suffix = Path(member.filename).suffix.casefold()
                label = f"{filename}/{member.filename}"
                if member_suffix == ".xlsx":
                    tables.extend(_xlsx_tables(member_data, label))
                elif member_suffix in {".csv", ".tsv", ".txt"}:
                    tables.append(_csv_table(member_data, label))
        if tables:
            return tables
    raise ValueError("Upload an XLSX, CSV, or ZIP containing the URL and status-code list.")


def _clean_status(value: object) -> object:
    if value in (None, ""):
        return ""
    text = str(value).strip()
    match = re.search(r"\b([1-5]\d{2})\b", text)
    return int(match.group(1)) if match else text


def _records_from_table(table: DataTable) -> list[dict[str, object]]:
    records = []
    for row in table.rows[1:]:
        if not row or not _is_url(row[0]):
            continue
        url = str(row[0]).strip()
        records.append(
            {
                "url": url,
                "normalized": normalize_compare_url(url),
                "status": _clean_status(row[1] if len(row) > 1 else ""),
                "final_redirect_url": str(row[2]).strip() if len(row) > 2 and row[2] not in (None, "") else "",
                "has_parameters": bool(urlsplit(url).query),
            }
        )
    return records


def load_url_status_records(data: bytes, filename: str) -> list[dict[str, object]]:
    candidates = [_records_from_table(table) for table in load_tables(data, filename)]
    candidates = [records for records in candidates if records]
    if not candidates:
        raise ValueError("No URLs found from A2 onward. Column A must contain URLs and column B status codes.")
    source_records = max(candidates, key=len)

    # Treat clean and parameterized versions as the same URL. Prefer a clean URL when both exist.
    chosen: dict[str, dict[str, object]] = {}
    order: list[str] = []
    for record in source_records:
        key = str(record["normalized"])
        if key not in chosen:
            chosen[key] = record
            order.append(key)
        elif chosen[key]["has_parameters"] and not record["has_parameters"]:
            previous = chosen[key]
            replacement = dict(record)
            if replacement["status"] in (None, ""):
                replacement["status"] = previous["status"]
            if not replacement.get("final_redirect_url"):
                replacement["final_redirect_url"] = previous.get("final_redirect_url", "")
            chosen[key] = replacement
        else:
            if chosen[key]["status"] in (None, "") and record["status"] not in (None, ""):
                chosen[key]["status"] = record["status"]
            if not chosen[key].get("final_redirect_url") and record.get("final_redirect_url"):
                chosen[key]["final_redirect_url"] = record["final_redirect_url"]
    return [chosen[key] for key in order]


def fetch_sitemap_html(sitemap_url: str, timeout: int = 30) -> tuple[str, str, int]:
    if not sitemap_url.casefold().startswith(("http://", "https://")):
        raise ValueError("Enter a complete HTML sitemap URL beginning with http:// or https://.")
    request = Request(
        sitemap_url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Cache-Control": "no-cache",
        },
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            source = response.read(25_000_000).decode(charset, errors="replace")
            return source, response.geturl(), int(response.status)
    except HTTPError as exc:
        raise ValueError(
            f"The sitemap page returned HTTP {exc.code}. Use the HTML upload or paste fallback."
        ) from exc
    except (URLError, TimeoutError, OSError) as exc:
        raise ValueError(f"The sitemap page could not be fetched: {exc}") from exc


def parse_sitemap_anchors(source: str, sitemap_url: str) -> list[dict[str, str]]:
    parser = SitemapAnchorParser(sitemap_url)
    parser.feed(source)
    sitemap_host = (urlsplit(sitemap_url).hostname or "").casefold()
    anchors = []
    for item in parser.anchors:
        if (urlsplit(item["url"]).hostname or "").casefold() != sitemap_host:
            continue
        item["normalized"] = normalize_compare_url(item["url"])
        anchors.append(item)
    if not anchors:
        raise ValueError("No same-domain HTML links were found in the sitemap source.")
    return anchors


def _in_scope(url: str, mode: str, pattern: str) -> bool:
    if mode == "Complete HTML sitemap audit":
        return True
    return pattern.casefold().strip() in clean_url_without_parameters(url).casefold()


ACRONYMS = {
    "ai": "AI", "hpc": "HPC", "gpu": "GPU", "cpu": "CPU", "nas": "NAS",
    "san": "SAN", "sdi": "SDI", "cio": "CIO", "it": "IT", "vmware": "VMware",
    "truscale": "TruScale", "lenovo": "Lenovo", "thinksystem": "ThinkSystem",
    "thinkagile": "ThinkAgile", "neptune": "Neptune", "amd": "AMD",
}


def suggest_anchor_text(url: str) -> str:
    parts = [part for part in unquote(urlsplit(url).path).split("/") if part]
    ignored = {"in", "en", "us", "uk", "br", "pt", "c", "p", "d", "index", "html"}
    slug = next((part for part in reversed(parts) if part.casefold() not in ignored), "")
    words = [word for word in re.split(r"[-_]+", slug) if word]
    return " ".join(ACRONYMS.get(word.casefold(), word.capitalize()) for word in words)


def _status_group(status: object) -> str:
    try:
        code = int(status)
    except (TypeError, ValueError):
        return "other"
    if code == 200:
        return "200"
    if 300 <= code < 400:
        return "3xx"
    return "other"


def _audit_rows(
    uploaded_records: list[dict[str, object]],
    sitemap_anchors: list[dict[str, str]],
    mode: str,
    pattern: str,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    scoped_uploaded = [r for r in uploaded_records if _in_scope(str(r["url"]), mode, pattern)]
    scoped_anchors = [r for r in sitemap_anchors if _in_scope(r["url"], mode, pattern)]
    if mode != "Complete HTML sitemap audit" and not pattern.strip():
        raise ValueError("Enter a folder or URL pattern, for example /servers-storage/.")

    uploaded_map = {str(record["normalized"]): record for record in scoped_uploaded}
    sitemap_keys = {item["normalized"] for item in scoped_anchors}
    url_counts = Counter(item["normalized"] for item in scoped_anchors)
    anchor_counts = Counter(normalize_text(item["anchor"]).casefold() for item in scoped_anchors if item["anchor"])

    existing = []
    for item in scoped_anchors:
        key = item["normalized"]
        uploaded = uploaded_map.get(key)
        status = uploaded["status"] if uploaded else ""
        notes = []
        suggested_link = ""
        if not uploaded:
            notes.append("Manually verify: URL is in the HTML sitemap but missing from the uploaded URL/status list")
        if url_counts[key] > 1:
            notes.append("Duplicate Link")
        anchor_key = normalize_text(item["anchor"]).casefold()
        if anchor_key and anchor_counts[anchor_key] > 1:
            notes.append("Duplicate Anchor")
        if _status_group(status) == "3xx":
            notes.append("Review redirecting link")
            if uploaded and uploaded.get("final_redirect_url"):
                suggested_link = clean_url_without_parameters(str(uploaded["final_redirect_url"]))
        elif status not in (None, "") and _status_group(status) == "other":
            try:
                if int(status) >= 400:
                    notes.append("Remove or replace broken link")
            except (TypeError, ValueError):
                pass
        if urlsplit(item["url"]).query:
            notes.append("Replace parameterized link with clean URL")
            if not suggested_link:
                suggested_link = clean_url_without_parameters(item["url"])
        existing.append(
            {
                "url": item["url"],
                "anchor": item["anchor"],
                "status": status,
                "in_sitemap": "Yes",
                "in_source": "Yes",
                "source_type": "Vanilla HTML",
                "note": " | ".join(dict.fromkeys(notes)),
                "suggested_link": suggested_link,
            }
        )

    missing = []
    for record in scoped_uploaded:
        if record["normalized"] in sitemap_keys:
            continue
        missing.append(
            {
                "url": record["url"],
                "status": record["status"],
                "anchor": suggest_anchor_text(str(record["url"])),
            }
        )
    return existing, missing


def _apply_table_style(ws, widths: list[float], alignments: list[str]) -> None:
    blue_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = blue_fill
        cell.font = Font(name="Aptos Narrow", size=11, bold=True, color="000000")
        cell.border = border
        cell.alignment = Alignment(
            horizontal="left" if cell.column == 1 else "center",
            vertical="center",
            wrap_text=True,
        )
    for row in ws.iter_rows(min_row=2):
        for index, cell in enumerate(row):
            cell.font = Font(name="Aptos Narrow", size=11, color="000000")
            cell.border = border
            cell.alignment = Alignment(
                horizontal=alignments[index], vertical="center", wrap_text=False
            )
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.auto_filter.ref = f"A1:{chr(64 + ws.max_column)}{max(1, ws.max_row)}"


def _build_audit_workbook(
    sitemap_url: str,
    existing: list[dict[str, object]],
    missing: list[dict[str, object]],
) -> bytes:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    existing_ws = wb.create_sheet("Existing Pages in HTML Sitemap")
    missing_ws = wb.create_sheet("URLs Not in HTML Sitemap")

    summary["A1"] = f"Sitemap URL: {sitemap_url}"
    summary.merge_cells("A1:B1")
    summary.append(["Sheet", "Pages With 200 Status Code", "Pages With 3xx Pages", "Total Number of URLs"])
    summary.append([
        "IN Existing",
        sum(_status_group(row["status"]) == "200" for row in existing),
        sum(_status_group(row["status"]) == "3xx" for row in existing),
        len(existing),
    ])
    summary.append([
        "URLs Not in HTML Sitemap",
        sum(_status_group(row["status"]) == "200" for row in missing),
        sum(_status_group(row["status"]) == "3xx" for row in missing),
        len(missing),
    ])
    blue_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in summary.iter_rows(min_row=1, max_row=4, min_col=1, max_col=4):
        for cell in row:
            cell.font = Font(name="Aptos Narrow", size=11, bold=cell.row in {1, 2})
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "center", vertical="center")
            if cell.row in {1, 2}:
                cell.fill = blue_fill
    summary.column_dimensions["A"].width = 31
    for column in ("B", "C", "D"):
        summary.column_dimensions[column].width = 31

    existing_ws.append([
        "URL", "Anchor Text", "Status Code", "In HTML Sitemap", "In HTML Source",
        "Source Code Type", "Note For Developer", "Suggested Link",
    ])
    for row in existing:
        existing_ws.append([
            row["url"], row["anchor"], row["status"], row["in_sitemap"], row["in_source"],
            row["source_type"], row["note"], row["suggested_link"],
        ])
    _apply_table_style(
        existing_ws,
        [76.43, 48.86, 20.71, 27, 23.71, 37.71, 45.57, 75],
        ["left", "left", "center", "center", "center", "center", "center", "left"],
    )

    missing_ws.append(["URLs Missing in HTML Sitemap", "Status Code", "Suggested Anchor Text"])
    for row in missing:
        missing_ws.append([row["url"], row["status"], row["anchor"]])
    _apply_table_style(missing_ws, [99.57, 28.14, 40.86], ["left", "center", "left"])

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def create_html_sitemap_audit(
    url_status_data: bytes,
    url_status_filename: str,
    sitemap_source: str,
    sitemap_url: str,
    scope_mode: str,
    scope_pattern: str = "",
) -> tuple[bytes, list[dict[str, object]], list[dict[str, object]], dict[str, int]]:
    uploaded = load_url_status_records(url_status_data, url_status_filename)
    anchors = parse_sitemap_anchors(sitemap_source, sitemap_url)
    existing, missing = _audit_rows(uploaded, anchors, scope_mode, scope_pattern)
    output = _build_audit_workbook(sitemap_url, existing, missing)
    stats = {
        "uploaded_urls": sum(_in_scope(str(r["url"]), scope_mode, scope_pattern) for r in uploaded),
        "sitemap_links": len(existing),
        "missing_urls": len(missing),
        "sitemap_only": sum("missing from the uploaded" in str(r["note"]) for r in existing),
    }
    return output, existing, missing, stats


def sitemap_output_filename(original_name: str) -> str:
    return f"{Path(original_name).stem}-HTML-Sitemap-Audit.xlsx"
