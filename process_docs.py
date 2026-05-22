"""Identify and redact sensitive data (companies, people, logos) in aged-care .doc/.docx files."""

from __future__ import annotations

import argparse
import base64
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from anthropic import Anthropic
from docx import Document
from dotenv import load_dotenv

MODEL = "claude-opus-4-7"

# 1x1 white PNG, used to blank out identified company logos in the redacted copy.
BLANK_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
    "+A8AAQUBAScY42YAAAAASUVORK5CYII="
)

IMAGE_MEDIA_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".webp": "image/webp",
}


def convert_doc_to_docx(doc_path: Path, out_dir: Path) -> Path:
    """Convert a legacy .doc to .docx using headless LibreOffice."""
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(out_dir),
            str(doc_path),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed for {doc_path.name}:\n{result.stderr}"
        )
    converted = out_dir / (doc_path.stem + ".docx")
    if not converted.exists():
        raise RuntimeError(f"Expected converted file not found: {converted}")
    return converted


def extract_text(docx_path: Path) -> str:
    doc = Document(str(docx_path))
    parts: list[str] = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        parts.append(para.text)
    # Headers and footers
    for section in doc.sections:
        for container in (section.header, section.footer):
            for para in container.paragraphs:
                if para.text.strip():
                    parts.append(para.text)
    return "\n".join(parts)


def extract_images(docx_path: Path, out_dir: Path) -> list[tuple[str, Path]]:
    """Extract embedded media. Returns list of (zip_internal_name, extracted_path)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[tuple[str, Path]] = []
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                target = out_dir / Path(name).name
                with z.open(name) as src, open(target, "wb") as dst:
                    dst.write(src.read())
                extracted.append((name, target))
    return extracted


def _strip_json_fences(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:]
    return raw.strip()


def identify_text_entities(text: str, client: Anthropic) -> dict:
    """Ask Claude to find company and personal names."""
    prompt = (
        "You are analyzing a document from an aged-care applicant file. "
        "Identify sensitive entities so they can be redacted.\n\n"
        "Return ONLY a JSON object with this shape:\n"
        '{"companies": ["..."], "people": ["..."]}\n\n'
        "Rules:\n"
        "- List each unique entity once, using the exact spelling/case from the text.\n"
        "- 'companies' = organisation, business, provider, agency, or service names "
        "(including aged-care providers, hospitals, banks, employers).\n"
        "- 'people' = full names of individuals (applicants, family, contacts, doctors, "
        "case workers). Do not list role/title alone.\n"
        "- Do NOT include generic terms ('the company', 'the doctor').\n"
        "- Do NOT invent entities; only list what appears.\n\n"
        f"Document text:\n---\n{text}\n---"
    )
    msg = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = _strip_json_fences(msg.content[0].text)
    data = json.loads(raw)
    return {
        "companies": list(dict.fromkeys(data.get("companies", []))),
        "people": list(dict.fromkeys(data.get("people", []))),
    }


def identify_logo(image_path: Path, client: Anthropic) -> dict:
    """Use vision to decide if an image is a company logo, and which company."""
    media_type = IMAGE_MEDIA_TYPES.get(image_path.suffix.lower())
    if media_type is None:
        return {"is_logo": False, "company": None, "description": f"unsupported format {image_path.suffix}"}

    with open(image_path, "rb") as f:
        img_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

    msg = client.messages.create(
        model=MODEL,
        max_tokens=512,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": img_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": (
                            "Is this image a company or organisation logo? "
                            "If yes, identify the company name from the visible mark or wordmark. "
                            'Return ONLY JSON: {"is_logo": bool, "company": string|null, "description": string}.'
                        ),
                    },
                ],
            }
        ],
    )
    raw = _strip_json_fences(msg.content[0].text)
    return json.loads(raw)


def redact_paragraph(paragraph, replacements: dict[str, str]) -> None:
    """Apply text replacements to a paragraph, handling text split across runs."""
    if not paragraph.runs:
        return
    full = "".join(r.text for r in paragraph.runs)
    new_full = full
    for key, value in replacements.items():
        if key:
            new_full = new_full.replace(key, value)
    if new_full == full:
        return
    paragraph.runs[0].text = new_full
    for run in paragraph.runs[1:]:
        run.text = ""


def redact_text(docx_path: Path, findings: dict, out_path: Path) -> None:
    """Write a copy with company/person strings replaced by tokens."""
    doc = Document(str(docx_path))

    replacements: dict[str, str] = {}
    for name in findings.get("people", []):
        replacements[name] = "[REDACTED-PERSON]"
    for company in findings.get("companies", []):
        replacements[company] = "[REDACTED-COMPANY]"
    # Longer strings first so 'Acme Health Ltd' is replaced before 'Acme'.
    replacements = dict(sorted(replacements.items(), key=lambda kv: -len(kv[0])))

    for para in doc.paragraphs:
        redact_paragraph(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    redact_paragraph(para, replacements)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for para in container.paragraphs:
                redact_paragraph(para, replacements)

    doc.save(str(out_path))


def blank_logos_in_docx(docx_path: Path, logo_zip_names: list[str]) -> None:
    """Overwrite identified logo image streams in-place with a 1x1 white PNG."""
    if not logo_zip_names:
        return
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in logo_zip_names:
                data = BLANK_PNG
            zout.writestr(item, data)
    tmp.replace(docx_path)


def process_file(input_path: Path, out_dir: Path, client: Anthropic) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir = out_dir / input_path.stem
    work_dir.mkdir(exist_ok=True)

    if input_path.suffix.lower() == ".doc":
        print(f"[{input_path.name}] converting .doc → .docx")
        docx_path = convert_doc_to_docx(input_path, work_dir / "_converted")
    else:
        docx_path = work_dir / input_path.name
        shutil.copy2(input_path, docx_path)

    print(f"[{input_path.name}] extracting text + images")
    text = extract_text(docx_path)
    images = extract_images(docx_path, work_dir / "images")

    print(f"[{input_path.name}] identifying companies + people (text)")
    entities = identify_text_entities(text, client) if text.strip() else {"companies": [], "people": []}

    logos: list[dict] = []
    logo_zip_names: list[str] = []
    for zip_name, img_path in images:
        print(f"[{input_path.name}] inspecting image {img_path.name}")
        try:
            result = identify_logo(img_path, client)
        except json.JSONDecodeError as e:
            result = {"is_logo": False, "company": None, "description": f"parse error: {e}"}
        result["file"] = str(img_path.relative_to(out_dir))
        result["zip_name"] = zip_name
        logos.append(result)
        if result.get("is_logo"):
            logo_zip_names.append(zip_name)
            if result.get("company") and result["company"] not in entities["companies"]:
                entities["companies"].append(result["company"])

    redacted_path = out_dir / f"{input_path.stem}.redacted.docx"
    redact_text(docx_path, entities, redacted_path)
    blank_logos_in_docx(redacted_path, logo_zip_names)

    report = {
        "source": str(input_path),
        "redacted_copy": str(redacted_path),
        "entities": entities,
        "logos": logos,
    }
    report_path = out_dir / f"{input_path.stem}.report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False))
    print(f"[{input_path.name}] wrote {redacted_path.name} + {report_path.name}")

    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", help="One or more .doc / .docx files (or directories).")
    parser.add_argument("-o", "--out", default="out", help="Output directory (default: ./out)")
    args = parser.parse_args()

    load_dotenv()
    client = Anthropic()  # picks up ANTHROPIC_API_KEY from env

    files: list[Path] = []
    for raw in args.inputs:
        p = Path(raw)
        if p.is_dir():
            files.extend(sorted(p.rglob("*.doc")))
            files.extend(sorted(p.rglob("*.docx")))
        elif p.suffix.lower() in {".doc", ".docx"}:
            files.append(p)
        else:
            print(f"skipping {p}: not a .doc/.docx", file=sys.stderr)

    if not files:
        print("no input files found", file=sys.stderr)
        return 1

    out_dir = Path(args.out)
    for f in files:
        try:
            process_file(f, out_dir, client)
        except Exception as e:
            print(f"[{f.name}] FAILED: {e}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
