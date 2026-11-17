"""Generate editable Word and searchable PDF resumes from the portfolio content."""

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import urlsplit
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt, RGBColor
from docx.text.run import Run
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Paragraph


PAGE_WIDTH, PAGE_HEIGHT = A4
RESUME_FILENAME = "Mohamed Zaitoun Resume"
SIDE_MARGIN = 48
TOP_MARGIN = 46
BOTTOM_MARGIN = 45
INK = "23271F"
OLIVE = "455239"
MUTED = "4F5549"


@dataclass(frozen=True)
class TypeStyle:
    word_style: str
    pdf_font: str
    size: float
    leading: float
    before: float = 0
    after: float = 0
    color: str = INK


STYLES = {
    "name": TypeStyle("Title", "Helvetica-Bold", 32, 38, after=2),
    "headline": TypeStyle("Subtitle", "Helvetica", 15, 20, after=6, color=OLIVE),
    "contact": TypeStyle("Contact links", "Helvetica", 10, 14, after=13, color=OLIVE),
    "summary": TypeStyle("Profile summary", "Helvetica", 11.3, 16, after=3),
    "section": TypeStyle("Heading 1", "Helvetica-Bold", 16, 20, before=13, after=5),
    "subhead": TypeStyle("Heading 2", "Helvetica-Bold", 12, 16, before=6, after=2),
    "meta": TypeStyle("Resume details", "Helvetica", 10, 14, after=4, color=MUTED),
    "bullet": TypeStyle("List Bullet", "Helvetica", 10.6, 14.6, after=3.5),
    "text": TypeStyle("Resume body", "Helvetica", 10.6, 14.6, after=4),
}


def _row(kind, text):
    return {"kind": kind, "text": text}


def _employment(role):
    organization = role["employer"]
    if role["client"]:
        organization += ", supporting " + role["client"]
    period = f'{role["start"]} to {role["end"]}'
    metadata = " | ".join(part for part in (organization, role["location"], period) if part)
    return [
        _row("subhead", role["role"]),
        _row("meta", metadata),
        *(_row("bullet", bullet) for bullet in role["bullets"]),
    ]


def resume_pages(profile):
    """Return ordered visible rows for each page, shared by HTML, Word and PDF."""
    resume = profile["resume"]
    contact = profile["contact"]
    links = [
        {"label": contact["email"], "url": "mailto:" + contact["email"]},
        {"label": contact["phone"], "url": "tel:" + contact["phone"].replace(" ", "")},
        {"label": urlsplit(contact["website"]).netloc.removeprefix("www."), "url": contact["website"]},
        {"label": "LinkedIn", "url": contact["linkedin"]},
        {"label": "GitHub", "url": contact["github"]},
    ]
    projects = {project["id"]: project for project in profile["projects"]}
    featured = projects[resume["featuredProject"]]
    featured_details = [featured.get("case", {}).get("role"), featured["year"], ", ".join(featured["stack"])]
    contact_line = " | ".join(link["label"] for link in links)
    first = [
        _row("name", profile["person"]["name"]),
        _row("headline", resume["headline"]),
        {"kind": "contact", "text": contact_line, "links": links},
        _row("summary", resume["summary"]),
        _row("section", "Featured project"),
        _row("subhead", featured["name"]),
        _row("meta", " | ".join(part for part in featured_details if part)),
        _row("text", featured.get("resumeDescription", featured["summary"])),
        _row("section", "Experience"),
    ]
    for role in profile["experience"][:2]:
        first.extend(_employment(role))
    first.append(_row("section", "Skills"))
    for group in profile["skills"]:
        first.append(_row("text", group["label"] + " | " + ", ".join(group["items"])))

    second = [_row("section", "Earlier experience")]
    for role in profile["experience"][2:]:
        second.extend(_employment(role))
    second.append(_row("section", "Selected projects"))
    for project_id in resume["selectedProjects"]:
        project = projects[project_id]
        heading = _row("subhead", project["name"])
        if project["links"]:
            heading["url"] = project["links"][0]["url"]
        second.extend([heading, _row("text", project["summary"])])
    second.append(_row("section", "Education"))
    for qualification in profile["education"]:
        second.extend([
            _row("subhead", qualification["credential"]),
            _row("meta", qualification["institution"] + " | " + qualification["status"]),
        ])
    second.append(_row("section", "Languages"))
    languages = ", ".join(f'{item["name"]} ({item["level"]})' for item in profile["languages"])
    second.append(_row("text", languages))
    return [first, second]


def _word_styles(document):
    """Use explicit dimensions so the document also opens reliably in other editors."""
    for kind, layout in STYLES.items():
        if layout.word_style in document.styles:
            style = document.styles[layout.word_style]
        else:
            style = document.styles.add_style(layout.word_style, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(layout.size)
        style.font.bold = layout.pdf_font.endswith("Bold")
        style.font.italic = False
        style.font.color.rgb = RGBColor.from_string(layout.color)
        fonts = style.element.get_or_add_rPr().rFonts
        for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + attribute), None)
        paragraph = style.paragraph_format
        paragraph.space_before = Pt(layout.before)
        paragraph.space_after = Pt(layout.after)
        paragraph.line_spacing = Pt(layout.leading)
        paragraph.keep_together = True
        paragraph.keep_with_next = kind in {"section", "subhead", "meta"}
        if kind == "bullet":
            paragraph.left_indent = Pt(14)
            paragraph.first_line_indent = Pt(-14)
        properties = style.element.get_or_add_pPr()
        border = properties.find(qn("w:pBdr"))
        if border is not None:
            properties.remove(border)


def _word_link(paragraph, label, url):
    hyperlink = OxmlElement("w:hyperlink")
    relationship = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink.set(qn("r:id"), relationship)
    run = Run(OxmlElement("w:r"), paragraph)
    run.text = label
    run.underline = True
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)


def _write_word(path, pages, name):
    document = Document()
    _word_styles(document)
    section = document.sections[0]
    section.page_width, section.page_height = Pt(PAGE_WIDTH), Pt(PAGE_HEIGHT)
    section.left_margin = section.right_margin = Pt(SIDE_MARGIN)
    section.top_margin, section.bottom_margin = Pt(TOP_MARGIN), Pt(BOTTOM_MARGIN)
    properties = document.core_properties
    properties.title = name + " Resume"
    properties.author = properties.last_modified_by = name
    properties.comments = ""
    # The bundled blank Word template carries dates unrelated to this document.
    for date_field in ("dcterms:created", "dcterms:modified"):
        element = properties._element.find(qn(date_field))
        if element is not None:
            properties._element.remove(element)

    for number, rows in enumerate(pages):
        for index, row in enumerate(rows):
            paragraph = document.add_paragraph(style=STYLES[row["kind"]].word_style)
            if index == 0:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.page_break_before = bool(number)
            if row["kind"] == "contact":
                for link_index, link in enumerate(row["links"]):
                    if link_index:
                        paragraph.add_run(" | ")
                    _word_link(paragraph, link["label"], link["url"])
            elif "url" in row:
                _word_link(paragraph, row["text"], row["url"])
            else:
                paragraph.add_run(row["text"])
    stream = BytesIO()
    document.save(stream)
    # Preserve byte stability across builds even though ZIP normally records wall time.
    with ZipFile(stream) as source, ZipFile(path, "w") as output:
        for entry in source.infolist():
            data = source.read(entry)
            entry.date_time = (1980, 1, 1, 0, 0, 0)
            output.writestr(entry, data)


def _write_pdf(path, pages, name):
    pdf = Canvas(str(path), pagesize=A4, invariant=True, pageCompression=1)
    pdf.setTitle(name + " Resume")
    pdf.setAuthor(name)
    pdf.setCreator("Portfolio resume generator")
    width = PAGE_WIDTH - 2 * SIDE_MARGIN
    for number, rows in enumerate(pages, start=1):
        cursor = PAGE_HEIGHT - TOP_MARGIN
        for index, row in enumerate(rows):
            kind = row["kind"]
            layout = STYLES[kind]
            style = ParagraphStyle(
                kind, fontName=layout.pdf_font, fontSize=layout.size, leading=layout.leading,
                textColor=HexColor("#" + layout.color), leftIndent=14 if kind == "bullet" else 0,
                bulletFontName="Helvetica", bulletFontSize=layout.size, bulletIndent=1,
            )
            if kind == "contact":
                content = " | ".join(
                    f'<link href={quoteattr(link["url"])}><u>{escape(link["label"])}</u></link>'
                    for link in row["links"]
                )
            elif "url" in row:
                content = f'<link href={quoteattr(row["url"])}><u>{escape(row["text"])}</u></link>'
            else:
                content = escape(row["text"])
            paragraph = Paragraph(content, style, bulletText="•" if kind == "bullet" else None)
            if index:
                cursor -= layout.before
            _, height = paragraph.wrap(width, PAGE_HEIGHT)
            cursor -= height
            if cursor < BOTTOM_MARGIN:
                raise ValueError(f"Resume page {number} overflows at '{row['text']}'. Shorten the selection or adjust pagination.")
            paragraph.drawOn(pdf, SIDE_MARGIN, cursor)
            cursor -= layout.after
        pdf.showPage()
    pdf.save()


def render_resume(profile: dict, output_dir: Path) -> tuple[Path, Path]:
    """Return the editable DOCX and published PDF after both are generated."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    name = profile["person"]["name"]
    word = output_dir / f"{RESUME_FILENAME}.docx"
    pdf = output_dir / f"{RESUME_FILENAME}.pdf"
    pages = resume_pages(profile)
    with TemporaryDirectory(dir=output_dir, prefix=".resume-") as scratch:
        draft_word, draft_pdf = Path(scratch) / word.name, Path(scratch) / pdf.name
        _write_word(draft_word, pages, name)
        _write_pdf(draft_pdf, pages, name)
        draft_word.replace(word)
        draft_pdf.replace(pdf)
    return word, pdf
