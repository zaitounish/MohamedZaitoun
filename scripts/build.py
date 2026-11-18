"""Build the static portfolio, public PDF and editable Word resume."""

import argparse
import json
import os
from pathlib import Path
from shutil import copy2, copytree
from tempfile import NamedTemporaryFile, TemporaryDirectory
from urllib.parse import quote

from jinja2 import Environment, FileSystemLoader, StrictUndefined, select_autoescape

from resume import RESUME_FILENAME, render_resume, resume_pages


ROOT = Path(__file__).resolve().parents[1]
PROFILE = ROOT / "content" / "profile.json"
BUILD_MARKER = ".zaitounish-build"
PDF_NAME = RESUME_FILENAME + ".pdf"


def html_resume_rows(profile):
    """Group adjacent Word/PDF bullet paragraphs into semantic HTML lists."""
    rows = []
    for page in resume_pages(profile):
        for row in page:
            if row["kind"] != "bullet":
                rows.append(row)
                continue
            if not rows or rows[-1]["kind"] != "list":
                rows.append({"kind": "list", "bullets": []})
            rows[-1]["bullets"].append(row["text"])
    return rows


def render_site(profile, site):
    site.mkdir()
    templates = Environment(
        loader=FileSystemLoader([ROOT / "templates", ROOT]),
        autoescape=select_autoescape(["html"]),
        undefined=StrictUndefined,
        trim_blocks=True,
        lstrip_blocks=True,
        keep_trailing_newline=True,
    )
    person = profile["person"]
    given, _, family = person["name"].strip().partition(" ")
    site_url = profile["contact"]["website"].rstrip("/")
    hub = next((p for p in profile["projects"] if p["id"] == "merchant-success-hub"), None)
    if hub is None:
        raise ValueError("The profile needs the merchant-success-hub project for its case study")
    common = {
        **profile,
        "hub": hub,
        "given_name": given,
        "family_name": family,
        "site_url": site_url,
        "pdf_url": quote(PDF_NAME),
        "resume_rows": html_resume_rows(profile),
    }
    pages = (
        ("home", "index.html", "", f"{person['name']} | {person['handle']}", person["intro"]),
        ("case", "case.html", "work/merchant-success-hub/", f"{hub['name']} | {person['name']}", hub["case"]["intro"]),
        ("resume", "resume.html", "resume/", f"Resume | {person['name']}", profile["resume"]["summary"]),
    )
    for page, template, route, title, description in pages:
        base = "../" * len(Path(route).parts)
        html = templates.get_template(template).render(
            **common, page=page, base=base, home=base, title=title,
            description=description, canonical=f"{site_url}/{route}",
        )
        destination = site / route / "index.html"
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(html, encoding="utf-8")
    copytree(ROOT / "assets", site / "assets")
    (site / "css").mkdir()
    copy2(ROOT / "css" / "Zaitounish.css", site / "css")
    for filename in ("CNAME", ".nojekyll"):
        copy2(ROOT / filename, site / filename)


def copy_word(source, destination):
    destination = destination.expanduser().resolve()
    if destination.suffix.lower() != ".docx":
        raise ValueError("--word-copy needs a .docx filename")
    with NamedTemporaryFile(dir=destination.parent, suffix=".docx", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def build(profile_path, output):
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    if output == ROOT or output in ROOT.parents:
        raise ValueError("Choose a dedicated build output directory")
    if output.exists():
        if not output.is_dir() or (any(output.iterdir()) and not (output / BUILD_MARKER).is_file()):
            raise ValueError(f"Choose an empty output directory or a previous portfolio build: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix=f".{output.name}-", dir=output.parent) as temporary:
        staging = Path(temporary) / "ready"
        staging.mkdir()
        site = staging / "site"
        render_site(profile, site)
        resume_dir = staging / "resume"
        resume_dir.mkdir()
        docx, pdf = render_resume(profile, resume_dir)
        pdf.replace(site / PDF_NAME)
        (staging / BUILD_MARKER).touch()

        previous = Path(temporary) / "previous"
        if output.exists():
            output.replace(previous)
        try:
            staging.replace(output)
        except OSError:
            if previous.exists():
                previous.replace(output)
            raise
    return output / "resume" / docx.name


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "dist")
    parser.add_argument("--profile", type=Path, default=PROFILE)
    parser.add_argument("--word-copy", type=Path, help="Also update this editable .docx file.")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        word = build(args.profile.resolve(), output)
        if args.word_copy:
            copy_word(word, args.word_copy)
    except (OSError, ValueError) as error:
        parser.exit(1, f"Build failed: {error}\n")
    print(f"Website and PDF: {output / 'site'}")
    print(f"Editable resume: {word}")


if __name__ == "__main__":
    main()
