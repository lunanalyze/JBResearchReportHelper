from __future__ import annotations

import argparse
import csv
import datetime as dt
import email.message
import email.utils
import html as html_lib
import json
import os
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from lxml import etree, html
from openpyxl import Workbook
try:
    from googlenewsdecoder import gnewsdecoder
except Exception:
    gnewsdecoder = None

sys.dont_write_bytecode = True

import paths

BASE_DIR = paths.APP_DIR
OUT_DIR = paths.APP_DATA_DIR
PDF_DIR = OUT_DIR / "raw_pdfs"
META_CSV = OUT_DIR / "metadata.csv"
META_XLSX = OUT_DIR / "metadata.xlsx"
RUNS_DIR = paths.RUNS_DIR
paths.ensure_app_dirs()
paths.migrate_legacy_data()

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)


@dataclass
class Item:
    category: str
    source_name: str
    title: str
    url: str
    published_date: str = ""
    file_type: str = "pdf"
    download_url: str = ""
    local_path: str = ""
    notes: str = ""
    original_url: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def published_mm_dd(self) -> str:
        parsed = parse_date(self.published_date)
        return parsed.strftime("%m.%d") if parsed else ""

    def row(self) -> dict:
        normalized_date = normalize_date(self.published_date) or self.published_date
        return {
            "category": self.category,
            "source_name": self.source_name,
            "title": self.title,
            "url": self.url,
            "original_url": self.original_url,
            "published_date": normalized_date,
            "published_mm_dd": self.published_mm_dd,
            "file_type": self.file_type,
            "download_url": self.download_url,
            "local_path": self.local_path,
            "notes": self.notes,
            "extra_json": json.dumps(self.extra, ensure_ascii=False),
        }


class Http:
    def __init__(self) -> None:
        ctx = ssl.create_default_context()
        self.opener = urllib.request.build_opener(
            urllib.request.ProxyHandler({}),
            urllib.request.HTTPCookieProcessor(),
            urllib.request.HTTPSHandler(context=ctx),
        )

    def get(self, url: str, referer: str | None = None, timeout: int = 35, attempts: int = 3) -> bytes:
        headers = {"User-Agent": USER_AGENT}
        if referer:
            headers["Referer"] = referer
        req = urllib.request.Request(url, headers=headers)
        return self._open_with_retry(req, timeout=timeout, attempts=attempts).read()

    def post_form(
        self, url: str, data: dict, referer: str | None = None, timeout: int = 35
    ) -> bytes:
        headers = {
            "User-Agent": USER_AGENT,
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
        }
        if referer:
            headers["Referer"] = referer
        body = urllib.parse.urlencode(data).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers=headers)
        return self._open_with_retry(req, timeout=timeout).read()

    def open_response(self, url: str, referer: str | None = None, timeout: int = 50, attempts: int = 3):
        headers = {"User-Agent": USER_AGENT}
        if referer:
            headers["Referer"] = referer
        return self._open_with_retry(urllib.request.Request(url, headers=headers), timeout=timeout, attempts=attempts)

    def _open_with_retry(self, req: urllib.request.Request, timeout: int, attempts: int = 3):
        last_exc: Exception | None = None
        for attempt in range(attempts):
            try:
                return self.opener.open(req, timeout=timeout)
            except (urllib.error.URLError, TimeoutError, ConnectionResetError) as exc:
                last_exc = exc
                time.sleep(0.7 * (attempt + 1))
        assert last_exc is not None
        raise last_exc


def decode_html(data: bytes) -> str:
    head = data[:2000].decode("ascii", "ignore")
    m = re.search(r"charset=['\"]?([-\w]+)", head, re.I)
    enc = m.group(1) if m else "utf-8"
    try:
        return data.decode(enc, "replace")
    except LookupError:
        return data.decode("utf-8", "replace")


def text_content(node) -> str:
    return re.sub(r"\s+", " ", node.text_content()).strip()


def clean_post_title(value: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    text = re.sub(r"\b\d{6}[_\-\s]*", "", text)
    text = re.sub(r"^\(?\s*(보도자료|보도참고|보도설명|별첨|참고자료)\s*\)?\s*", "", text)
    text = re.sub(r"\.pdf$", "", text, flags=re.I).strip()
    return text.strip(" _-")


def title_from_pdf_filename(file_name: str) -> str:
    name = urllib.parse.unquote(file_name or "")
    name = re.sub(r"\.pdf$", "", name, flags=re.I)
    name = re.sub(r"^\s*\d{6}[_\-\s]*", "", name)
    name = re.sub(r"^\(?\s*(보도자료|보도참고|보도설명|별첨|참고자료)\s*\)?\s*", "", name)
    name = re.sub(r"^\[?\s*(보도자료|보도참고|보도설명|별첨|참고자료)\s*\]?\s*", "", name)
    return clean_post_title(name)


def first_non_file_link(node):
    if node is None:
        return None
    for a in node.xpath(".//a[@href]"):
        href = a.get("href") or ""
        if any(marker in href for marker in ["fileDown.do", "/comm/getFile", "fileSrc", "ezpdfwv", "customLayout"]):
            continue
        label = text_content(a)
        if label in {"파일뷰어", "바로보기", "다운로드", "첨부파일"}:
            continue
        if label and not label.lower().endswith(".pdf"):
            return a
    return None


def nearest_container(node):
    current = node
    while current is not None:
        tag = getattr(current, "tag", "")
        if tag in {"tr", "li"}:
            return current
        cls = current.get("class") or ""
        if any(token in cls.lower() for token in ["board", "list", "item", "bbs"]):
            return current
        current = current.getparent()
    return node.getparent() if node is not None else None


def parse_date(value: str | None) -> dt.date | None:
    if not value:
        return None
    s = value.strip()
    patterns = [
        r"(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})",
        r"(\d{2})(\d{2})(\d{2})",
    ]
    for p in patterns:
        m = re.search(p, s)
        if not m:
            continue
        y, mo, d = [int(x) for x in m.groups()]
        if y < 100:
            y += 2000
        try:
            return dt.date(y, mo, d)
        except ValueError:
            return None
    return None


def normalize_date(value: str | None) -> str:
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else ""


def find_date_text(*values: str) -> str:
    text = " ".join(v or "" for v in values)
    match = re.search(r"20\d{2}[.\-/년\s]+\d{1,2}[.\-/월\s]+\d{1,2}|\d{6}", text)
    return match.group(0) if match else ""


def within_days(value: str, days: int, today: dt.date) -> bool:
    parsed = parse_date(value)
    if not parsed:
        return True
    return today - dt.timedelta(days=days) <= parsed <= today


def absolutize(base: str, href: str) -> str:
    return urllib.parse.urljoin(base, html_lib.unescape(href))


def clean_filename(name: str, fallback: str = "download.pdf") -> str:
    name = urllib.parse.unquote(name or fallback)
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip().strip(".")
    if not name:
        name = fallback
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name[:180]


def repair_mojibake(value: str) -> str:
    if not value:
        return value
    value = urllib.parse.unquote(value)
    try:
        repaired = value.encode("latin-1").decode("utf-8")
        if sum("\uac00" <= ch <= "\ud7a3" for ch in repaired) > sum("\uac00" <= ch <= "\ud7a3" for ch in value):
            return repaired
    except UnicodeError:
        pass
    return value


def filename_from_headers(headers: email.message.Message, fallback: str) -> str:
    disp = headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=UTF-8''([^;]+)", disp, re.I)
    if m:
        candidate = repair_mojibake(m.group(1))
        if candidate.lower().endswith(".pdf"):
            return clean_filename(candidate, fallback)
        return clean_filename(fallback)
    m = re.search(r'filename="?([^";]+)"?', disp, re.I)
    if m:
        candidate = repair_mojibake(m.group(1))
        if candidate.lower().endswith(".pdf"):
            return clean_filename(candidate, fallback)
        return clean_filename(fallback)
    return clean_filename(fallback)


def save_pdf(http: Http, item: Item, dry_run: bool, pdf_dir: Path | None = None) -> Item:
    if not item.download_url:
        item.notes = append_note(item.notes, "no download_url")
        return item
    if dry_run:
        item.local_path = ""
        return item
    if pdf_dir is None:
        pdf_dir = PDF_DIR
    safe_source = clean_filename(item.source_name, item.source_name).replace(".pdf", "")
    target_dir = pdf_dir / safe_source
    target_dir.mkdir(parents=True, exist_ok=True)
    try:
        is_hana = "hanaif.re.kr" in item.download_url
        timeout = 5 if is_hana else 50
        attempts = 1 if is_hana else 3
        with http.open_response(item.download_url, referer=item.url, timeout=timeout, attempts=attempts) as resp:
            first = resp.read(8)
            fallback = clean_filename(item.title)
            filename = filename_from_headers(resp.headers, fallback)
            path = target_dir / filename
            if not first.startswith(b"%PDF"):
                item.local_path = ""
                item.notes = append_note(
                    item.notes,
                    f"skipped non-PDF response: content-type={resp.headers.get('Content-Type')} first bytes={first!r}",
                )
                return item
            if path.exists():
                item.local_path = str(path)
                item.notes = append_note(item.notes, "skipped existing file")
                return item
            rest = resp.read()
            path.write_bytes(first + rest)
            item.local_path = str(path)
    except Exception as exc:
        item.notes = append_note(item.notes, f"download failed: {exc}")
    return item


def append_note(notes: str, value: str) -> str:
    return value if not notes else f"{notes}; {value}"


def dedupe(items: Iterable[Item]) -> list[Item]:
    seen: set[str] = set()
    out: list[Item] = []
    for item in items:
        key = item.download_url or item.url or f"{item.source_name}:{item.title}"
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def parse_doc(data: bytes):
    return html.fromstring(decode_html(data))


def collect_fss(http: Http, days: int, today: dt.date) -> list[Item]:
    url = "https://www.fss.or.kr/fss/bbs/B0000188/list.do?menuNo=200218&bbsId=&cl1Cd=&pageIndex=1&sdate=&edate=&searchCnd=1&searchWrd="
    doc = parse_doc(http.get(url))
    items: list[Item] = []
    for tr in doc.xpath("//tr"):
        row_text = text_content(tr)
        date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}|\d{2}\d{2}\d{2}", row_text) or [None])[0]
        if date and not within_days(date, days, today):
            continue
        post_link = first_non_file_link(tr)
        post_title = text_content(post_link) if post_link is not None else ""
        post_url = absolutize(url, post_link.get("href")) if post_link is not None else url
        for a in tr.xpath(".//a[contains(@href,'fileDown.do')]"):
            file_name = text_content(a) or row_text[:100]
            if ".pdf" not in file_name.lower():
                continue
            title = clean_post_title(post_title) or title_from_pdf_filename(file_name) or clean_post_title(row_text)
            items.append(
                Item(
                    category="국가기관",
                    source_name="금융감독원",
                    title=title,
                    url=post_url,
                    published_date=date or "",
                    download_url=absolutize(url, a.get("href")),
                    extra={"file_name": file_name},
                )
            )
    return items


def collect_fsc(http: Http, days: int, today: dt.date) -> list[Item]:
    url = "https://www.fsc.go.kr/no010101?curPage=1"
    doc = parse_doc(http.get(url))
    items: list[Item] = []
    for a in doc.xpath("//a[contains(@href,'/comm/getFile')]"):
        file_name = text_content(a)
        href = absolutize(url, a.get("href"))
        container = nearest_container(a)
        context = text_content(container) if container is not None else file_name
        post_link = first_non_file_link(container) if container is not None else None
        post_title = text_content(post_link) if post_link is not None else ""
        post_url = absolutize(url, post_link.get("href")) if post_link is not None else url
        date = find_date_text(file_name, context)
        if date and not within_days(date, days, today):
            continue
        if ".pdf" not in (file_name + href).lower():
            # The endpoint can still be PDF, but non-PDF attachments are common here.
            continue
        title = clean_post_title(post_title) or title_from_pdf_filename(file_name) or "금융위원회 첨부 PDF"
        items.append(
            Item(
                category="국가기관",
                source_name="금융위원회",
                title=title,
                url=post_url,
                published_date=normalize_date(date),
                download_url=href,
                extra={"file_name": file_name},
            )
        )
    return items


def collect_bok(http: Http, days: int, today: dt.date) -> list[Item]:
    base = "https://www.bok.or.kr"
    list_url = (
        "https://www.bok.or.kr/portal/singl/newsData/listCont.do?pageIndex=1&targetDepth=&"
        "menuNo=201150&syncMenuChekKey=5&depthSubMain=&subMainAt=&searchCnd=1&searchKwd=&"
        "depth2=200038&depth2=201156&date=&sdate=&edate=&sort=1&pageUnit=10"
    )
    doc = parse_doc(http.get(list_url, referer="https://www.bok.or.kr/portal/singl/newsData/list.do?menuNo=201150"))
    items: list[Item] = []
    for a in doc.xpath("//a[contains(@href,'/view.do')]"):
        detail_url = absolutize(base, a.get("href"))
        title = clean_listing_title(text_content(a))
        container = nearest_container(a)
        row_text = text_content(container) if container is not None else text_content(a.getparent()) if a.getparent() is not None else title
        date = find_date_text(row_text)
        if date and not within_days(date, days, today):
            continue
        try:
            detail_bytes = http.get(detail_url, referer=list_url)
            detail = parse_doc(detail_bytes)
        except Exception as exc:
            items.append(Item("국가기관", "한국은행", title, detail_url, normalize_date(date), notes=f"detail failed: {exc}"))
            continue
        date = date or find_date_text(text_content(detail), decode_html(detail_bytes))
        for fa in detail.xpath("//a[contains(@href,'.pdf') or contains(@href,'fileSrc')]"):
            fname = text_content(fa)
            href = absolutize(detail_url, fa.get("href"))
            if "viewer.html" in href or fname == "뷰어":
                continue
            if ".pdf" not in (fname + href).lower():
                continue
            items.append(
                Item(
                    "국가기관",
                    "한국은행",
                    title,
                    detail_url,
                    normalize_date(date),
                    download_url=href,
                    extra={"file_name": fname},
                )
            )
    return items


def collect_kif(http: Http, days: int, today: dt.date) -> list[Item]:
    list_url = "https://www.kif.re.kr/kif4/publication/pub_list?mid=20"
    api_url = "https://www.kif.re.kr/kif4/biz/async_proc"
    items: list[Item] = []
    start_date = today - dt.timedelta(days=days)
    params = {
        "ac": "dataSearch",
        "mid": "20",
        "nid": "0",
        "vid": "0",
        "t1": "",
        "t2": "",
        "df": start_date.isoformat(),
        "dt": today.isoformat(),
        "kw": "",
        "pn": "1",
        "at": "0",
        "sfield": "",
        "pcnt": "50",
        "lang": "0",
    }
    try:
        raw = http.get(f"{api_url}?{urllib.parse.urlencode(params)}", referer=list_url)
        data = json.loads(raw.decode("utf-8-sig"))
    except Exception:
        data = {}

    for row in data.get("datalist", []):
        date = str(row.get("pubdate") or "")
        if date and not within_days(date, days, today):
            continue
        mid = str(row.get("mid") or "20")
        nid = str(row.get("nid") or "0")
        sid = str(row.get("sid") or "")
        vid = str(row.get("vid") or "0")
        cno = str(row.get("cno") or "")
        if not cno:
            continue
        detail_url = (
            "https://www.kif.re.kr/kif4/publication/pub_detail?"
            f"mid={urllib.parse.quote(mid)}&nid={urllib.parse.quote(nid)}"
            f"&sid={urllib.parse.quote(sid)}&vid={urllib.parse.quote(vid)}"
            f"&cno={urllib.parse.quote(cno)}&pn=1"
        )
        title = clean_text(str(row.get("title") or row.get("pubname") or "KIF 금융브리프"))
        source = clean_text(str(row.get("source") or row.get("pubname") or "금융브리프"))
        detail_params = {
            "ac": "getPubDetailInfoV2",
            "mid": mid,
            "nid": "0",
            "vid": vid,
            "cno": cno,
            "lang": "0",
        }
        try:
            detail_raw = http.get(f"{api_url}?{urllib.parse.urlencode(detail_params)}", referer=detail_url)
            detail = json.loads(detail_raw.decode("utf-8-sig"))
        except Exception:
            continue
        for file_info in detail.get("attachfiles", []):
            if str(file_info.get("ext") or "").lower() != ".pdf":
                continue
            fcd = str(file_info.get("fcd") or "")
            if not fcd:
                continue
            fname = clean_text(str(file_info.get("fname") or ""))
            download = (
                "https://www.kif.re.kr/kif4/publication/viewer?"
                f"mid={urllib.parse.quote(mid)}&vid=0&cno={urllib.parse.quote(cno)}"
                f"&fcd={urllib.parse.quote(fcd)}&ft=0"
            )
            items.append(
                Item(
                    category="국가기관",
                    source_name="한국금융연구원",
                    title=title,
                    url=detail_url,
                    published_date=date,
                    download_url=download,
                    extra={"section": source, "article_title": title, "file_name": fname},
                )
            )

    if items:
        return items

    doc = parse_doc(http.get(list_url))
    detail_links = [
        absolutize(list_url, a.get("href"))
        for a in doc.xpath("//a[contains(@href,'pub_detail')]")
    ]
    for detail_url in detail_links[:20]:
        try:
            detail_text = decode_html(http.get(detail_url, referer=list_url))
        except Exception:
            continue
        items.extend(kif_items_from_detail(detail_text, detail_url, days, today))
    return items


def kif_items_from_detail(detail_text: str, detail_url: str, days: int, today: dt.date) -> list[Item]:
    doc = html.fromstring(detail_text)
    page_title = text_content(doc.xpath("//h3|//h4|//title")[0]) if doc.xpath("//h3|//h4|//title") else "KIF 금융브리프"
    page_date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", text_content(doc)) or [None])[0] or ""
    if page_date and not within_days(page_date, days, today):
        return []
    out = []
    for m in re.finditer(r"execDownload\(([^)]*)\)", detail_text):
        args = parse_js_args(m.group(1))
        if len(args) < 7:
            continue
        prefix, mid, vid, cno, fcd, _ext, ft = args[:7]
        if str(prefix):
            base = urllib.parse.urljoin(detail_url, str(prefix))
        else:
            base = "https://www.kif.re.kr/kif4/publication/"
        download = urllib.parse.urljoin(
            base,
            f"viewer?mid={mid}&vid={vid}&cno={cno}&fcd={urllib.parse.quote(str(fcd))}&ft={ft}",
        )
        out.append(
            Item(
                category="국가기관",
                source_name="한국금융연구원",
                title=page_title,
                url=detail_url,
                published_date=page_date,
                download_url=download,
            )
        )
    return out


def parse_js_args(arg_text: str) -> list[str | int]:
    args: list[str | int] = []
    for raw in re.findall(r"'[^']*'|\"[^\"]*\"|[^,]+", arg_text):
        raw = raw.strip()
        if not raw:
            continue
        if (raw.startswith("'") and raw.endswith("'")) or (raw.startswith('"') and raw.endswith('"')):
            args.append(raw[1:-1])
        else:
            try:
                args.append(int(raw))
            except ValueError:
                args.append(raw)
    return args


def collect_hana(http: Http, list_url: str, label: str, days: int, today: dt.date) -> list[Item]:
    data = decode_html(http.get(list_url, timeout=10, attempts=2))
    doc = html.fromstring(data)
    items: list[Item] = []
    for li in doc.xpath("//li[.//*[contains(@onclick,'downloadItem')]]"):
        block_text = text_content(li)
        title_nodes = li.xpath(".//*[contains(@class,'tit')]")
        title = text_content(title_nodes[0]) if title_nodes else ""
        if not title:
            link_nodes = li.xpath(".//a[contains(@href,'boardDetail') or contains(@onclick,'goPage')]")
            title = text_content(link_nodes[0]) if link_nodes else ""
        date_nodes = li.xpath(".//*[contains(@class,'date')]")
        date = text_content(date_nodes[0]) if date_nodes else ""
        if not date:
            date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", block_text) or [None])[0] or ""
        if date and not within_days(date, days, today):
            continue
        for a in li.xpath(".//a[contains(@onclick,'downloadItem')]"):
            onclick = a.get("onclick") or ""
            m = re.search(r"downloadItem\((\d+),\s*(\d+)\)", onclick)
            if not m:
                continue
            hmpe, seq = m.groups()
            file_name = clean_text(text_content(a))
            if ".pdf" not in file_name.lower():
                continue
            detail = (
                "https://www.hanaif.re.kr/boardDetail.do?"
                f"hmpeSeqNo={hmpe}&menuId={urllib.parse.parse_qs(urllib.parse.urlparse(list_url).query).get('menuId', [''])[0]}"
            )
            items.append(
                Item(
                    category="금융연구소",
                    source_name=f"하나금융연구소 {label}",
                    title=title[:140] or f"하나금융연구소 {label} PDF",
                    url=detail,
                    published_date=date,
                    download_url=f"https://www.hanaif.re.kr/dev/hanaifFileDownload.jsp?seq={seq}",
                    notes=f"preflight endpoint: https://www.hanaif.re.kr/download.do?hmpeSeqNo={hmpe}",
                    extra={"file_name": file_name},
                )
            )
    return items


def hana_detail_title_and_date(detail_text: str) -> tuple[str, str]:
    if not detail_text:
        return "", ""
    doc = html.fromstring(detail_text)
    title = ""
    title_nodes = doc.xpath("//*[contains(@class,'subTit02')]")
    if title_nodes:
        title = text_content(title_nodes[0])
    if not title:
        h_nodes = doc.xpath("//h1|//h2|//h3")
        title = text_content(h_nodes[0]) if h_nodes else ""
    source_text = " ".join(text_content(n) for n in doc.xpath("//*[contains(@class,'sourceBox')]"))
    date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", source_text) or [None])[0] or ""
    if not date:
        date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", text_content(doc)) or [None])[0] or ""
    return title, date


def collect_kb(http: Http, days: int, today: dt.date) -> list[Item]:
    url = "https://www.kbfg.com/kbresearch/report/reportList.do"
    doc = parse_doc(http.get(url))
    items: list[Item] = []
    seen_detail_urls: set[str] = set()
    for a in doc.xpath("//a[contains(@href,'reportView.do')]"):
        detail_url = absolutize(url, a.get("href"))
        if detail_url in seen_detail_urls:
            continue
        seen_detail_urls.add(detail_url)
        raw_title = text_content(a)
        title = clean_listing_title(raw_title)
        date = find_date_text(raw_title)
        if date and not within_days(date, days, today):
            continue
        try:
            detail_text = decode_html(http.get(detail_url, referer=url))
        except Exception as exc:
            items.append(Item("금융연구소", "KB경영연구소", title, detail_url, normalize_date(date), notes=f"detail failed: {exc}"))
            continue
        date = date or find_date_text(detail_text)
        for m in re.finditer(r"fn_downFile\('([^']+)'\s*,\s*'([^']+)'\)", detail_text):
            fid, sn = m.groups()
            fname = (re.search(rf"fn_downFile\('{re.escape(fid)}'\s*,\s*'{re.escape(sn)}'\)[^>]*>([^<]+)", detail_text) or [None, title])[1]
            item_date = date or find_date_text(fname)
            items.append(
                Item(
                    category="금융연구소",
                    source_name="KB경영연구소",
                    title=title,
                    url=detail_url,
                    published_date=normalize_date(item_date),
                    download_url=f"https://www.kbfg.com/kbresearch/cmm/fms/FileDown.do?atchFileId={fid}&fileSn={sn}",
                    extra={"file_name": clean_text(fname)},
                )
            )
    return items


def collect_kdb(http: Http, days: int, today: dt.date) -> list[Item]:
    referer = "https://rd.kdb.co.kr/FLMNMN00N01.act"
    http.get(referer)
    payload = {
        "REQ_PAGE_NO": 1,
        "PAGE_ROW_COUNT": 20,
        "NEXT_PAGE_YN": "S",
        "RSARRAY": "STD422,STD423,STD424,STD425,STD426,STD427,STD428,STD429",
        "SUB_REC": [],
    }
    body = http.post_form(
        "https://rd.kdb.co.kr/FLMNMN00R01.jct",
        {"_JSON_": urllib.parse.quote(json.dumps(payload, ensure_ascii=False))},
        referer=referer,
    )
    data = json.loads(body.decode("utf-8"))
    records = {str(r.get("ITR_NAC_ID_MNG_SNO")): r for r in data.get("REC", [])}
    items: list[Item] = []
    for sub in data.get("SUB_REC", []):
        rec = records.get(str(sub.get("ITR_NAC_ID_MNG_SNO")), {})
        date = rec.get("LST_CHG_DTM") or rec.get("FST_ENR_DTM") or sub.get("LST_CHG_DTM", "")
        if date and not within_days(date, days, today):
            continue
        file_name = sub.get("ORC_APG_FL_NM") or ""
        title = rec.get("NAC_CONE_TTL") or file_name or "KDB PDF"
        if str(sub.get("APG_FL_XTN_NM", "")).upper() != "PDF" and ".pdf" not in file_name.lower():
            continue
        group_id = sub.get("FL_MPN_ID")
        file_id = sub.get("APG_FL_MPN_ID")
        items.append(
            Item(
                category="금융연구소",
                source_name="KDB미래전략연구소",
                title=title,
                url=referer,
                published_date=date,
                download_url=f"https://rd.kdb.co.kr/fileView?groupId={group_id}&fileId={file_id}",
                extra={"board": rec.get("BLB_NM") or sub.get("BLB_NM"), "file_name": file_name},
            )
        )
    return items


def collect_wfri(http: Http, list_url: str, label: str, days: int, today: dt.date) -> list[Item]:
    doc = parse_doc(http.get(list_url))
    items: list[Item] = []
    for a in doc.xpath("//a[contains(@href,'mode=view') or contains(@href,'page_type=view')]"):
        detail_url = absolutize(list_url, a.get("href"))
        raw_title = text_content(a)
        title = clean_listing_title(raw_title)
        date = (re.search(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}", raw_title) or [None])[0] or ""
        if date and not within_days(date, days, today):
            continue
        try:
            detail = decode_html(http.get(detail_url, referer=list_url))
        except Exception as exc:
            items.append(Item("금융연구소", f"우리금융경영연구소 {label}", title, detail_url, date, notes=f"detail failed: {exc}"))
            continue
        for m in re.finditer(r"board_file_download\('([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\)", detail):
            idx, board, cnt = m.groups()
            fname = (re.search(rf"board_file_download\('{idx}'\s*,\s*'{board}'\s*,\s*'{cnt}'\)[^>]*>([^<]+)", detail) or [None, title])[1]
            items.append(
                Item(
                    category="금융연구소",
                    source_name=f"우리금융경영연구소 {label}",
                    title=title,
                    url=detail_url,
                    published_date=date,
                    download_url=f"https://www.wfri.re.kr/module/lib/board_file_download.php?idx={idx}&board_code={board}&file_cnt={cnt}",
                    extra={"file_name": clean_text(fname)},
                )
            )
    return items


def clean_text(value: str) -> str:
    return html_lib.unescape(re.sub(r"\s+", " ", value or "")).strip()


def clean_listing_title(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r"\s+[가-힣]{2,4}(?:\s*,\s*[가-힣]{2,4}){0,5}\s+20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}\s+조회수\s*\d+\s*$", "", text)
    text = re.sub(r"\s+20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}\s+조회수\s*\d+\s*$", "", text)
    text = re.sub(r"\s+(작성자|등록일|작성일|조회수)\s*[:：]?\s*.*$", "", text)
    return text.strip()


def best_pdf_name(fragment: str) -> str:
    text = clean_text(re.sub(r"<[^>]+>", " ", fragment))
    matches = re.findall(r"[\w가-힣ㄱ-ㅎㅏ-ㅣ\(\)\[\]\-.,& ]{2,180}\.pdf", text, re.I)
    if not matches:
        return ""
    cleaned = [clean_text(m) for m in matches]
    cleaned.sort(key=len)
    return cleaned[0]


DEFAULT_NEWS_KEYWORDS = {
    "bank": {
        "keywords": ["전북은행", "JB금융"],
        "groups": [["전북은행"], ["JB금융"]],
    },
    "other": {
        "keywords": [
            "금융위원회", "금융감독원", "한국은행", "가계부채", "DSR", "대출규제",
            "부동산 PF", "연체율", "부실채권", "금리", "환율", "금융시장", "예대금리차",
        ],
        "groups": [
            ["금리"], ["환율"], ["금융시장"], ["가계부채"], ["DSR"], ["대출규제"],
            ["부동산 PF"], ["연체율"], ["부실채권"], ["예대금리차"],
            ["한국은행", "금리"], ["금융위원회", "대출규제"], ["금융감독원", "부동산 PF"],
        ],
    },
}


def normalize_keyword_groups(value) -> list[list[str]]:
    return [item["keywords"] for item in normalize_keyword_filters(value)]


def normalize_keyword_filters(value, default_limit: int = 3) -> list[dict]:
    if isinstance(value, dict):
        raw_groups = value.get("groups") if "groups" in value else [[x] for x in value.get("keywords", [])]
    else:
        raw_groups = value
    filters: list[dict] = []
    if not isinstance(raw_groups, list):
        return filters
    for raw_group in raw_groups:
        limit = default_limit
        if isinstance(raw_group, dict):
            raw_items = raw_group.get("keywords") or raw_group.get("items") or []
            try:
                limit = int(raw_group.get("limit") or raw_group.get("max") or default_limit)
            except Exception:
                limit = default_limit
        elif isinstance(raw_group, str):
            raw_items = [raw_group]
        elif isinstance(raw_group, list):
            raw_items = raw_group
        else:
            continue
        group = []
        seen = set()
        for raw in raw_items:
            text = clean_text(str(raw or "")).strip().strip('"').strip("'").strip()
            if text and text not in seen:
                group.append(text)
                seen.add(text)
        if group:
            filters.append({"keywords": group, "limit": max(1, min(100, limit))})
    return filters


def google_news_or_query(keyword_groups) -> str:
    parts = []
    for group in normalize_keyword_groups(keyword_groups):
        quoted = [google_news_quote(keyword) for keyword in group]
        if len(quoted) == 1:
            parts.append(quoted[0])
        elif quoted:
            parts.append("(" + " ".join(quoted) + ")")
    return " OR ".join(parts)


def google_news_quote(keyword: str) -> str:
    escaped = str(keyword).replace('"', '\\"')
    return f'"{escaped}"'


def google_news_and_query(group: list[str]) -> str:
    return " ".join(google_news_quote(keyword) for keyword in group if str(keyword).strip())


def clean_google_news_title(title: str, source_name: str = "") -> str:
    title = clean_text(title)
    source_name = clean_text(source_name)
    if source_name:
        suffix = f" - {source_name}"
        while title.endswith(suffix):
            title = title[: -len(suffix)].strip()
        return title
    return re.sub(r"\s+-\s+[^-]{2,80}$", "", title).strip()


def is_google_news_url(url: str) -> bool:
    try:
        host = urllib.parse.urlparse(url).netloc.lower()
    except Exception:
        return False
    return host.endswith("news.google.com") or host == "news.google.com"


def original_url_from_google_news(http: Http, google_link: str) -> str:
    if not google_link:
        return ""
    if not is_google_news_url(google_link):
        return google_link
    parsed = urllib.parse.urlparse(google_link)
    params = urllib.parse.parse_qs(parsed.query)
    for key in ["url", "u", "q"]:
        value = params.get(key, [""])[0]
        if value and not is_google_news_url(value):
            return value
    if gnewsdecoder is None:
        raise RuntimeError("googlenewsdecoder is not installed")
    result = gnewsdecoder(google_link)
    if not result.get("status"):
        raise RuntimeError(f"Google News decoding failed: {result.get('message') or 'unknown error'}")
    decoded = clean_text(result.get("decoded_url") or "")
    if not is_valid_article_url(decoded):
        raise RuntimeError(f"Google News decoding returned invalid article URL: {decoded or '(empty)'}")
    return decoded


def is_valid_article_url(url: str) -> bool:
    if not url or is_google_news_url(url):
        return False
    try:
        parsed = urllib.parse.urlparse(url)
    except Exception:
        return False
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return False
    path = (parsed.path or "").strip("/")
    if path:
        return True
    return bool(parsed.query)


def is_daum_article_url(url: str) -> bool:
    try:
        parsed = urllib.parse.urlparse(url)
    except Exception:
        return False
    return parsed.netloc.lower() == "v.daum.net" and parsed.path.startswith("/v/")


def extract_meta_content(doc, *names: str) -> str:
    for name in names:
        values = doc.xpath(
            "//meta[translate(@property, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')=$name"
            " or translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')=$name]/@content",
            name=name.lower(),
        )
        for value in values:
            text = clean_text(value)
            if text:
                return text
    return ""


def clean_daum_source_name(value: str) -> str:
    text = clean_text(value)
    if "|" in text:
        text = clean_text(text.split("|")[-1])
    if text.lower() in {"daum", "daum 뉴스", "v.daum.net"}:
        return ""
    return text


def is_daum_source_label(value: str) -> bool:
    text = clean_text(value).lower()
    return text in {"daum", "daum 뉴스", "v.daum.net"} or text.endswith("daum.net")


def daum_source_name(http: Http, url: str) -> str:
    if not is_daum_article_url(url):
        return ""
    data = http.get(url, timeout=10, attempts=1)
    doc = parse_doc(data)
    for value in [
        extract_meta_content(doc, "og:article:author"),
        extract_meta_content(doc, "article:author"),
        extract_meta_content(doc, "og:site_name"),
    ]:
        source = clean_daum_source_name(value)
        if source:
            return source
    html_text = decode_html(data)
    match = re.search(r'cpKorName\s*:\s*["\']([^"\']+)["\']', html_text)
    if match:
        return clean_daum_source_name(html_lib.unescape(match.group(1)))
    return ""


def normalize_article_source_name(http: Http, item: Item) -> None:
    if item.file_type != "article":
        return
    article_url = item.original_url or item.url
    if not is_daum_article_url(article_url):
        return
    source = daum_source_name(http, article_url)
    if source:
        item.source_name = source


def normalize_daum_article_item(http: Http, item: Item) -> None:
    if item.file_type != "article" or not is_daum_source_label(item.source_name):
        return
    google_url = ""
    if isinstance(item.extra, dict):
        google_url = clean_text(item.extra.get("google_news_url") or "")
    google_url = google_url or item.url
    if not item.original_url and google_url:
        item.original_url = original_url_from_google_news(http, google_url)
    normalize_article_source_name(http, item)


def google_news_base64_token(url: str) -> str:
    try:
        parsed = urllib.parse.urlparse(url)
    except Exception:
        return ""
    parts = [part for part in parsed.path.split("/") if part]
    if parsed.netloc.lower().endswith("news.google.com") and len(parts) >= 2 and parts[-2] in {"articles", "read"}:
        return parts[-1]
    return ""


def decode_google_news_url(http: Http, google_link: str) -> str:
    token = google_news_base64_token(google_link)
    if not token:
        return ""
    signature = timestamp = ""
    for prefix in ["articles", "rss/articles"]:
        try:
            page = decode_html(http.get(f"https://news.google.com/{prefix}/{token}", timeout=10, attempts=1))
        except Exception:
            continue
        signature_match = re.search(r'data-n-a-sg="([^"]+)"', page)
        timestamp_match = re.search(r'data-n-a-ts="([^"]+)"', page)
        if signature_match and timestamp_match:
            signature = html_lib.unescape(signature_match.group(1))
            timestamp = html_lib.unescape(timestamp_match.group(1))
            break
    if not signature or not timestamp:
        return ""
    payload = [
        [
            [
                "Fbv4je",
                (
                    '["garturlreq",[["X","X",["X","X"],null,null,1,1,"US:en",null,1,'
                    f'null,null,null,null,null,0,1],"X","X",1,[1,1,1],1,1,null,0,0,null,0],"{token}",{timestamp},"{signature}"]'
                ),
                None,
                "generic",
            ]
        ]
    ]
    raw = http.post_form(
        "https://news.google.com/_/DotsSplashUi/data/batchexecute",
        {"f.req": json.dumps(payload, separators=(",", ":"))},
        referer=f"https://news.google.com/articles/{token}",
        timeout=10,
    )
    text = decode_html(raw)
    parts = text.split("\n\n", 1)
    if len(parts) < 2:
        return ""
    parsed = json.loads(parts[1])
    decoded = json.loads(parsed[0][2])[1]
    return clean_text(decoded)


def collect_google_news(
    http: Http,
    days: int,
    today: dt.date,
    news_kind: str = "bank",
    queries=None,
    start_date: dt.date | None = None,
    end_date: dt.date | None = None,
    max_per_filter: int = 3,
) -> list[Item]:
    if queries is None:
        queries = DEFAULT_NEWS_KEYWORDS["bank" if news_kind == "bank" else "other"]
    category = "은행/지주사" if news_kind == "bank" else "그외"
    items: list[Item] = []
    filters = normalize_keyword_filters(queries, default_limit=max_per_filter)
    if not filters:
        return items

    if start_date is None or end_date is None:
        start_date = today - dt.timedelta(days=days)
        end_date = today
    if end_date < start_date:
        start_date, end_date = end_date, start_date

    actual_today = dt.date.today()
    lookback_days = max((actual_today - start_date).days + 2, (end_date - start_date).days + 1, 1)
    seen_links: set[str] = set()
    for keyword_filter in filters:
        group = keyword_filter["keywords"]
        limit = keyword_filter["limit"]
        query = google_news_and_query(group)
        if not query:
            continue
        if news_kind == "other":
            query = f"{query} -site:fsc.go.kr -site:fss.or.kr -site:bok.or.kr"
        dated_query = (
            f"{query} when:{lookback_days}d "
            f"after:{start_date.isoformat()} before:{(end_date + dt.timedelta(days=1)).isoformat()}"
        )
        rss_url = "https://news.google.com/rss/search?" + urllib.parse.urlencode(
            {"q": dated_query, "hl": "ko", "gl": "KR", "ceid": "KR:ko"}
        )
        try:
            root = etree.fromstring(http.get(rss_url))
            nodes = root.xpath("//item")
        except Exception as exc:
            items.append(Item(category, "Google News RSS", dated_query, rss_url, notes=f"rss failed: {exc}"))
            continue

        collected_for_filter = 0
        for node in nodes:
            raw_title = clean_text("".join(node.xpath("./title/text()")))
            google_link = clean_text("".join(node.xpath("./link/text()")))
            source_name = clean_text("".join(node.xpath("./source/text()"))) or "Google News RSS"
            if not google_link or google_link in seen_links:
                continue
            title = clean_google_news_title(raw_title, source_name)
            pub = clean_text("".join(node.xpath("./pubDate/text()")))
            desc = clean_text("".join(node.xpath("./description/text()")))
            parsed = parse_rfc_date(pub)
            if parsed and not (start_date <= parsed <= end_date):
                continue
            seen_links.add(google_link)
            extra = {
                "rss_description": desc,
                "news_keywords": group,
                "news_filter_limit": limit,
            }
            if google_link:
                extra["google_news_url"] = google_link
            item = Item(
                category=category,
                source_name=source_name,
                title=title,
                url=google_link,
                published_date=parsed.isoformat() if parsed else "",
                file_type="article",
                extra=extra,
            )
            if is_daum_source_label(source_name):
                try:
                    normalize_daum_article_item(http, item)
                except Exception as exc:
                    item.notes = append_note(item.notes, f"daum source normalize failed: {exc}")
            items.append(item)
            collected_for_filter += 1
            if collected_for_filter >= limit:
                break
    items.sort(key=lambda item: item.published_date or "", reverse=True)
    return items

def populate_original_urls(
    items: list[Item],
    progress: Callable[[str, dict], None] | None = None,
) -> None:
    cache: dict[str, str] = {}
    http = Http()
    article_items = [item for item in items if item.file_type == "article"]
    total = len(article_items)
    for index, item in enumerate(article_items, start=1):
        if item.original_url and not is_google_news_url(item.original_url):
            normalize_article_source_name(http, item)
            continue
        google_url = ""
        if isinstance(item.extra, dict):
            google_url = clean_text(item.extra.get("google_news_url") or "")
        google_url = google_url or item.original_url or item.url
        if progress:
            progress("decode_start", {"index": index, "total": total, "title": item.title})
        if google_url in cache:
            decoded = cache[google_url]
        else:
            try:
                decoded = original_url_from_google_news(http, google_url)
            except Exception as exc:
                item.notes = append_note(item.notes, f"Google News decode failed: {exc}")
                if progress:
                    progress("decode_error", {"index": index, "total": total, "title": item.title, "error": str(exc)})
                continue
            cache[google_url] = decoded
        item.original_url = decoded
        normalize_article_source_name(http, item)
        if progress:
            progress("decode_done", {"index": index, "total": total, "title": item.title, "original_url": decoded})


def fetch_article_text(http: Http, url: str, referer: str = "") -> tuple[str, str]:
    if not url:
        return "", "article url missing"
    if is_google_news_url(url):
        return "", "article url is still Google News"
    try:
        data = http.get(url, referer=referer or None, timeout=10, attempts=1)
        text = extract_article_text(data)
        if text:
            return text[:6000], ""
        return "", "article text not found"
    except Exception as exc:
        return "", f"article fetch failed: {exc}"


def extract_article_text(data: bytes) -> str:
    try:
        doc = parse_doc(data)
    except Exception:
        return ""
    for bad in doc.xpath("//script|//style|//noscript|//iframe|//nav|//header|//footer|//aside"):
        parent = bad.getparent()
        if parent is not None:
            parent.remove(bad)
    candidates = []
    selectors = [
        "//article",
        "//*[contains(@class,'article')]",
        "//*[contains(@class,'news')]",
        "//*[contains(@class,'content')]",
        "//*[contains(@id,'article')]",
        "//*[contains(@id,'content')]",
    ]
    seen = set()
    for selector in selectors:
        for node in doc.xpath(selector):
            if id(node) in seen:
                continue
            seen.add(id(node))
            text = clean_article_text(text_content(node))
            if len(text) >= 250:
                candidates.append(text)
    if not candidates:
        paragraphs = [clean_article_text(text_content(p)) for p in doc.xpath("//p")]
        paragraphs = [p for p in paragraphs if len(p) >= 35]
        text = clean_article_text("\n".join(paragraphs))
        if len(text) >= 250:
            candidates.append(text)
    if not candidates:
        body = doc.xpath("//body")
        if body:
            text = clean_article_text(text_content(body[0]))
            if len(text) >= 250:
                candidates.append(text)
    if not candidates:
        for text in extract_structured_article_texts(doc):
            text = clean_article_text(text)
            if len(text) >= 60:
                candidates.append(text)
    if not candidates:
        return ""
    candidates.sort(key=len, reverse=True)
    return candidates[0]


def extract_structured_article_texts(doc) -> list[str]:
    texts: list[str] = []
    for raw in doc.xpath("//script[@type='application/ld+json']/text()"):
        try:
            data = json.loads(raw)
        except Exception:
            continue
        texts.extend(extract_text_fields_from_json(data))
    for xpath in [
        "//meta[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='description']/@content",
        "//meta[translate(@property, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='og:description']/@content",
        "//meta[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='twitter:description']/@content",
    ]:
        texts.extend(str(value or "") for value in doc.xpath(xpath))
    return texts


def extract_text_fields_from_json(value) -> list[str]:
    texts: list[str] = []
    if isinstance(value, list):
        for item in value:
            texts.extend(extract_text_fields_from_json(item))
    elif isinstance(value, dict):
        for key in ["articleBody", "description"]:
            text = value.get(key)
            if isinstance(text, str):
                texts.append(text)
        for item in value.values():
            if isinstance(item, (dict, list)):
                texts.extend(extract_text_fields_from_json(item))
    return texts


def clean_article_text(value: str) -> str:
    text = html_lib.unescape(value or "")
    text = re.sub(r"\b기자\s*[:=]?\s*[\w가-힣.@-]+", " ", text)
    text = re.sub(r"\s*(구독|공유|댓글|좋아요|프린트|메일|전체기사|기사제보)\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_rfc_date(value: str) -> dt.date | None:
    if not value:
        return None
    try:
        parsed = email.utils.parsedate_to_datetime(value)
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone(dt.timezone(dt.timedelta(hours=9)))
        return parsed.date()
    except Exception:
        return parse_date(value)


def collect_all(days: int, max_per_source: int | None, dry_run: bool) -> list[Item]:
    today = dt.date.today()
    http = Http()
    collectors = [
        ("금융감독원", lambda: collect_fss(http, days, today)),
        ("금융위원회", lambda: collect_fsc(http, days, today)),
        ("한국은행", lambda: collect_bok(http, days, today)),
        ("한국금융연구원", lambda: collect_kif(http, days, today)),
        ("하나금융연구소 연구보고서", lambda: collect_hana(http, "https://www.hanaif.re.kr/boardList.do?menuId=MN1000&tabMenuId=N", "연구보고서", days, today)),
        ("하나금융연구소 정기보고서", lambda: collect_hana(http, "https://www.hanaif.re.kr/boardList.do?menuId=MN2000&tabMenuId=MN2100", "정기보고서", days, today)),
        ("KB경영연구소", lambda: collect_kb(http, days, today)),
        ("KDB미래전략연구소", lambda: collect_kdb(http, days, today)),
        ("우리금융경영연구소 연구보고서", lambda: collect_wfri(http, "https://www.wfri.re.kr/ko/web/research_report/research_report.php", "연구보고서", days, today)),
        ("우리금융경영연구소 동남아 Review", lambda: collect_wfri(http, "https://www.wfri.re.kr/ko/web/serial/eastSouthAsia.php", "동남아 Review", days, today)),
        ("은행/지주사 뉴스", lambda: collect_google_news(http, days, today, "bank")),
    ]
    all_items: list[Item] = []
    for name, fn in collectors:
        print(f"[collect] {name}", file=sys.stderr)
        try:
            items = dedupe(fn())
            if max_per_source is not None:
                items = items[:max_per_source]
            all_items.extend(items)
            print(f"  -> {len(items)} items", file=sys.stderr)
        except Exception as exc:
            print(f"  !! failed: {exc}", file=sys.stderr)
            if "Google News decoding failed" in str(exc) or "googlenewsdecoder" in str(exc):
                raise
            all_items.append(Item(category="error", source_name=name, title=name, url="", notes=str(exc)))
        time.sleep(0.15)
    downloaded: list[Item] = []
    for item in all_items:
        if item.file_type == "pdf":
            saved = save_pdf(http, item, dry_run=dry_run)
            if "skipped non-PDF response" in saved.notes:
                continue
            downloaded.append(saved)
        else:
            downloaded.append(item)
    return dedupe(downloaded)


def range_to_days(start: dt.date, end: dt.date) -> tuple[int, dt.date]:
    if end < start:
        start, end = end, start
    return max((end - start).days, 0), end


def collect_by_ranges(
    news_start: dt.date,
    news_end: dt.date,
    agency_start: dt.date,
    agency_end: dt.date,
    research_start: dt.date,
    research_end: dt.date,
    max_per_source: int | None = None,
    dry_run: bool = False,
    news_max: int | None = None,
    news_bank_max: int | None = None,
    news_other_max: int | None = None,
    agency_max: int | None = None,
    research_max: int | None = None,
    output_dir: Path | None = None,
    progress: Callable[[str, dict], None] | None = None,
    news_keywords: dict | None = None,
    include_news: bool = True,
    include_agency: bool = True,
    include_research: bool = True,
) -> list[Item]:
    http = Http()
    news_days, news_today = range_to_days(news_start, news_end)
    agency_days, agency_today = range_to_days(agency_start, agency_end)
    research_days, research_today = range_to_days(research_start, research_end)

    if news_max is None:
        news_max = max_per_source
    if news_bank_max is None:
        news_bank_max = news_max
    if news_other_max is None:
        news_other_max = news_max
    if agency_max is None:
        agency_max = max_per_source
    if research_max is None:
        research_max = max_per_source
    news_keywords = news_keywords or {}
    bank_news_keywords = news_keywords.get("bank") or DEFAULT_NEWS_KEYWORDS["bank"]
    other_news_keywords = news_keywords.get("other") or DEFAULT_NEWS_KEYWORDS["other"]

    collectors = []
    if include_agency:
        collectors.extend([
            ("금융감독원", "agency", lambda: collect_fss(http, agency_days, agency_today)),
            ("금융위원회", "agency", lambda: collect_fsc(http, agency_days, agency_today)),
            ("한국은행", "agency", lambda: collect_bok(http, agency_days, agency_today)),
            ("한국금융연구원", "agency", lambda: collect_kif(http, agency_days, agency_today)),
        ])
    if include_research:
        collectors.extend([
            ("하나금융연구소 연구보고서", "research", lambda: collect_hana(http, "https://www.hanaif.re.kr/boardList.do?menuId=MN1000&tabMenuId=N", "연구보고서", research_days, research_today)),
            ("하나금융연구소 정기보고서", "research", lambda: collect_hana(http, "https://www.hanaif.re.kr/boardList.do?menuId=MN2000&tabMenuId=MN2100", "정기보고서", research_days, research_today)),
            ("KB경영연구소", "research", lambda: collect_kb(http, research_days, research_today)),
            ("KDB미래전략연구소", "research", lambda: collect_kdb(http, research_days, research_today)),
            ("우리금융경영연구소 연구보고서", "research", lambda: collect_wfri(http, "https://www.wfri.re.kr/ko/web/research_report/research_report.php", "연구보고서", research_days, research_today)),
            ("우리금융경영연구소 동남아 Review", "research", lambda: collect_wfri(http, "https://www.wfri.re.kr/ko/web/serial/eastSouthAsia.php", "동남아 Review", research_days, research_today)),
        ])
    if include_news:
        collectors.extend([
            ("은행/지주사 뉴스", "news_bank", lambda: collect_google_news(http, news_days, news_today, "bank", bank_news_keywords, news_start, news_end, max_per_filter=news_bank_max or 3)),
            ("그외 뉴스", "news_other", lambda: collect_google_news(http, news_days, news_today, "other", other_news_keywords, news_start, news_end, max_per_filter=news_other_max or 3)),
        ])

    all_items: list[Item] = []
    limits = {"news_bank": None, "news_other": None, "agency": agency_max, "research": research_max}
    for name, group, fn in collectors:
        print(f"[collect] {name}", file=sys.stderr)
        if progress:
            progress("source_start", {"source": name, "group": group})
        try:
            items = dedupe(fn())
            limit = limits[group]
            if limit is not None:
                items = items[:limit]
            all_items.extend(items)
            print(f"  -> {len(items)} items", file=sys.stderr)
            if progress:
                progress("source_done", {"source": name, "group": group, "count": len(items)})
        except Exception as exc:
            print(f"  !! failed: {exc}", file=sys.stderr)
            if group in {"news_bank", "news_other"}:
                raise
            all_items.append(Item(category="error", source_name=name, title=name, url="", notes=str(exc)))
            if progress:
                progress("source_error", {"source": name, "group": group, "error": str(exc)})
        time.sleep(0.15)

    downloaded: list[Item] = []
    pdf_dir = output_dir / "raw_pdfs" if output_dir is not None else None
    pdf_total = sum(1 for item in all_items if item.file_type == "pdf")
    pdf_index = 0
    for item in all_items:
        if item.file_type == "pdf":
            pdf_index += 1
            if progress:
                progress(
                    "download_start",
                    {
                        "index": pdf_index,
                        "total": pdf_total,
                        "source": item.source_name,
                        "title": item.title,
                    },
                )
            saved = save_pdf(http, item, dry_run=dry_run, pdf_dir=pdf_dir)
            if "skipped non-PDF response" in saved.notes:
                if progress:
                    progress(
                        "download_skip",
                        {
                            "index": pdf_index,
                            "total": pdf_total,
                            "source": item.source_name,
                            "title": item.title,
                            "notes": saved.notes,
                        },
                    )
                continue
            downloaded.append(saved)
            if progress:
                progress(
                    "download_done",
                    {
                        "index": pdf_index,
                        "total": pdf_total,
                        "source": item.source_name,
                        "title": item.title,
                    },
                )
        else:
            downloaded.append(item)
    result = dedupe(downloaded)
    if progress:
        progress("done", {"count": len(result)})
    return result


def write_outputs(
    items: list[Item],
    output_dir: Path | None = None,
    csv_path: Path | None = None,
    xlsx_path: Path | None = None,
) -> None:
    if output_dir is None:
        output_dir = OUT_DIR
    if csv_path is None:
        csv_path = output_dir / "metadata.csv"
    if xlsx_path is None:
        xlsx_path = output_dir / "metadata.xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)
    fieldnames = list(Item("", "", "", "").row().keys())
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            writer.writerow(item.row())

    wb = Workbook()
    ws = wb.active
    ws.title = "metadata"
    ws.append(fieldnames)
    for item in items:
        ws.append([item.row()[k] for k in fieldnames])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col[:100])
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 45)
    wb.save(xlsx_path)


def make_run_dir(base_dir: Path | None = None, when: dt.datetime | None = None) -> Path:
    if base_dir is None:
        base_dir = RUNS_DIR
    if when is None:
        when = dt.datetime.now()
    stamp = when.strftime("%y%m%d_%H%M")
    path = base_dir / stamp
    suffix = 1
    while path.exists():
        path = base_dir / f"{stamp}_{suffix}"
        suffix += 1
    path.mkdir(parents=True, exist_ok=True)
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Weekly research/news PDF collector")
    parser.add_argument("--days", type=int, default=7, help="look back this many days")
    parser.add_argument("--max-per-source", type=int, default=None, help="limit items per source")
    parser.add_argument("--dry-run", action="store_true", help="collect metadata without downloading PDFs")
    args = parser.parse_args()

    items = collect_all(args.days, args.max_per_source, args.dry_run)
    write_outputs(items, OUT_DIR, META_CSV, META_XLSX)
    print(f"items={len(items)}")
    print(f"metadata_csv={META_CSV}")
    print(f"metadata_xlsx={META_XLSX}")
    if not args.dry_run:
        print(f"pdf_dir={PDF_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
