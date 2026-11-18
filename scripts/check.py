"""Check generated HTML, links and both resume formats before publication."""

import argparse
import json
import posixpath
import re
import unicodedata
from html.parser import HTMLParser
from itertools import zip_longest
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from pypdf import PdfReader
from pypdf.errors import PdfReadError

from resume import RESUME_FILENAME, resume_pages


ROOT = Path(__file__).resolve().parents[1]
HOME = "index.html"
CASE = "work/merchant-success-hub/index.html"
RESUME = "resume/index.html"
FILENAME = RESUME_FILENAME
LEGACY_ANCHORS = {"hero", "about", "experience", "projects", "contact"}
CSS_URL = re.compile(r"""url\(\s*(['"]?)(.*?)\1\s*\)""", re.IGNORECASE)
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def clean_text(value):
    value = unicodedata.normalize("NFKC", value)
    return " ".join(value.translate(str.maketrans("", "", "\u2022\u007f")).split())


def check_dashes(value, location):
    if "\u2013" in value or "\u2014" in value:
        raise ValueError(f"{location} contains a long dash")


class Page(HTMLParser):
    def __init__(self, filename, markup):
        super().__init__()
        self.filename = filename
        self.ids = set()
        self.links = set()
        self.anchors = set()
        self.text = []
        self.hidden = 0
        self.feed(markup)
        self.close()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        self.hidden += tag in {"script", "style"}
        for key, value in attrs.items():
            if value is None:
                continue
            check_dashes(value, f"{self.filename} {key}")
            if key == "id":
                if value in self.ids:
                    raise ValueError(f"{self.filename} has duplicate id {value!r}")
                self.ids.add(value)
            elif key in {"href", "src", "poster", "action"}:
                self.links.add(value)
                if tag == "a" and key == "href":
                    self.anchors.add(value)
            elif key in {"srcset", "imagesrcset"}:
                for candidate in value.split(","):
                    fields = candidate.split()
                    if not fields:
                        raise ValueError(f"{self.filename} has an empty {key} candidate")
                    self.links.add(fields[0])
        if tag == "meta" and attrs.get("property") in {"og:url", "og:image"}:
            self.links.add(attrs.get("content") or "")

    def handle_endtag(self, tag):
        if tag in {"script", "style"}:
            self.hidden -= 1

    def handle_data(self, value):
        if not self.hidden and value.strip():
            check_dashes(value, self.filename)
            self.text.append(value)


def check_link(source, value, files, pages, domain):
    def reject(reason):
        raise ValueError(f"{source}: {value!r} {reason}")

    if not value.strip() or "\\" in value or any(ord(char) < 32 for char in value):
        reject("is empty or contains an invalid URL character")
    url = urlsplit(value)
    if url.scheme in {"mailto", "tel"} and url.path:
        return
    if url.scheme or url.netloc:
        if url.scheme not in {"https", "http", ""} or not url.hostname or url.username is not None:
            reject("uses an invalid or unsupported URL")
        if url.hostname.lower() != domain:
            return

    path = unquote(url.path)
    if "\\" in path or "\x00" in path:
        reject("contains an invalid file path")
    if path or url.netloc:
        parent = "" if url.netloc or path.startswith("/") else posixpath.dirname(source)
        target = posixpath.normpath(posixpath.join(parent, path.lstrip("/")))
    else:
        target = source
    if target == ".." or target.startswith("../"):
        reject("escapes the published directory")
    directory_index = posixpath.normpath(posixpath.join(target, "index.html"))
    if path.endswith("/") or target not in files and directory_index in files:
        target = directory_index
    if target not in files:
        reject(f"points to missing file {target!r}")
    if url.fragment and target in pages and unquote(url.fragment) not in pages[target].ids:
        reject("points to a missing HTML anchor")
    return target


def check_site(site):
    entries = list(site.rglob("*"))
    if any(entry.is_symlink() for entry in entries):
        raise ValueError("The published site contains a symbolic link")
    files = {p.relative_to(site).as_posix() for p in entries if p.is_file()}
    required = {HOME, CASE, RESUME, "CNAME", ".nojekyll", FILENAME + ".pdf"}
    if required - files:
        raise ValueError(f"Missing public files: {', '.join(sorted(required - files))}")
    if any(path.lower().endswith(".docx") for path in files):
        raise ValueError("The editable Word document belongs in the separate resume artifact")
    domain = (site / "CNAME").read_text(encoding="utf-8").strip().lower()
    if not domain or domain != (ROOT / "CNAME").read_text(encoding="utf-8").strip().lower():
        raise ValueError("Published CNAME differs from the configured custom domain")
    pages = {}
    for name in sorted(files):
        if name.endswith(".html"):
            markup = (site / name).read_text(encoding="utf-8")
            if "{{" in markup or "{%" in markup:
                raise ValueError(f"{name} contains unrendered template markup")
            pages[name] = Page(name, markup)
    missing = LEGACY_ANCHORS - pages[HOME].ids
    if missing:
        raise ValueError(f"Missing home anchors: {', '.join(sorted(missing))}")
    required_links = {HOME: {CASE, RESUME}, RESUME: {FILENAME + ".pdf"}}
    for name, page in pages.items():
        targets = {link: check_link(name, link, files, pages, domain) for link in sorted(page.links)}
        missing = required_links.get(name, set()) - {targets[link] for link in page.anchors}
        if missing:
            raise ValueError(f"{name} needs a link to {', '.join(sorted(missing))}")
    for name in sorted(path for path in files if path.endswith(".css")):
        css = re.sub(r"/\*.*?\*/", "", (site / name).read_text(encoding="utf-8"), flags=re.S)
        for _, link in CSS_URL.findall(css):
            check_link(name, link, files, pages, domain)
    return {name: clean_text(" ".join(page.text)) for name, page in pages.items()}


def word_text(path):
    with ZipFile(path) as archive:
        document = ElementTree.fromstring(archive.read("word/document.xml"))
    paragraphs = document.findall("./w:body/w:p", WORD_NS)
    return "\n".join("".join(node.text or "" for node in p.findall(".//w:t", WORD_NS)) for p in paragraphs)


def check_same_text(actual, expected, label):
    check_dashes(actual, label)
    actual_words = clean_text(actual).split()
    expected_words = clean_text(expected).split()
    for index, (seen, wanted) in enumerate(zip_longest(actual_words, expected_words), start=1):
        if seen != wanted:
            raise ValueError(f"{label} differs at word {index}. Expected {wanted!r}, found {seen!r}")


def check_order(text, values, label):
    position = 0
    for value in values:
        value = clean_text(value)
        found = text.find(value, position)
        if found < 0:
            raise ValueError(f"{label} is missing {value!r} or presents it out of order")
        position = found + len(value)


def check_outputs(output, profile):
    site = output / "site"
    pages = check_site(site)
    pdf = PdfReader(site / (FILENAME + ".pdf"))
    rows = resume_pages(profile)
    expected = [" ".join(row["text"] for row in page) for page in rows]
    if len(pdf.pages) != 2 or len(expected) != 2:
        raise ValueError("Expected a two-page resume")
    for i, (page, text) in enumerate(zip(pdf.pages, expected), start=1):
        check_same_text(page.extract_text() or "", text, f"PDF page {i}")
    docx = word_text(output / "resume" / (FILENAME + ".docx"))
    check_same_text(docx, " ".join(expected), "Editable Word resume")
    web_rows = []
    for page in rows:
        for row in page:
            if row["kind"] == "contact":
                web_rows.extend(link["label"] for link in row["links"])
            else:
                web_rows.append(row["text"])
    check_order(pages[RESUME], web_rows, "Web resume")

    hub = next((project for project in profile["projects"] if project["id"] == "merchant-success-hub"), None)
    if not hub:
        raise ValueError("Profile is missing the Merchant Success Hub case study")
    case = hub["case"]
    spotlight = case["spotlight"]
    story = []
    for section in case["story"]:
        story.append(section["title"])
        story.extend(section["paragraphs"])
    check_order(pages[HOME], [hub["summary"], case["contribution"], spotlight["title"], *spotlight["paragraphs"]], HOME)
    check_order(pages[CASE], [case["intro"], *story], CASE)
    roles = [entry["role"] for entry in profile["experience"]]
    for label, text in {HOME: pages[HOME], RESUME: pages[RESUME], "Resume documents": clean_text(docx)}.items():
        check_order(text, [hub["name"]], label)
        check_order(text, roles, label)
        for entry in profile["education"]:
            for key in ("credential", "institution", "status"):
                check_order(text, [entry[key]], label)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "dist")
    parser.add_argument("--profile", type=Path, default=ROOT / "content" / "profile.json")
    args = parser.parse_args()
    try:
        check_outputs(args.output, json.loads(args.profile.read_text(encoding="utf-8")))
    except (OSError, ValueError, KeyError, BadZipFile, ElementTree.ParseError, PdfReadError) as error:
        parser.exit(1, f"Portfolio check failed: {error}\n")
    print("Verified static pages, local links, legacy anchors, and the matching two-page PDF and editable Word resume.")


if __name__ == "__main__":
    main()
