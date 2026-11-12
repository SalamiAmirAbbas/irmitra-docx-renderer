
import html
import os
import re
import shutil
import subprocess
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

from pathlib import Path


SOURCE = Path("input.docx")
TEMP_DOCX = Path("input.patched.docx")

FONT_FILES = [
    Path("IRMitra.ttf"),
    Path("IRMitra-Bold.ttf"),
    Path("IRMitra-Italic.ttf"),
]

FONT_DIR = Path.home() / ".local" / "share" / "fonts" / "irmitra"

XML_PARTS = re.compile(
    r"^word/(?:document|styles|header\d+|footer\d+|"
    r"footnotes|endnotes|comments)\.xml$"
)

RPR_RE = re.compile(
    r'<w:rPr(?P<attrs>[^>]*)>(?P<body>.*?)</w:rPr>'
    r'|<w:rPr(?P<selfattrs>[^>]*)/>',
    re.DOTALL,
)

RUN_RE = re.compile(
    r'<w:r(?=[\s>])(?P<attrs>[^>]*)>'
    r'(?P<body>.*?)</w:r>',
    re.DOTALL,
)

RFONTS_RE = re.compile(
    r'<w:rFonts\b[^>]*/\s*>'
    r'|<w:rFonts\b[^>]*>.*?</w:rFonts>',
    re.DOTALL,
)


def run_command(args):
    result = subprocess.run(
        args,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return result.stdout.strip()


def font_family(font_path):
    output = run_command([
        "fc-scan",
        "--format",
        "%{family}\n",
        str(font_path),
    ])

    if not output:
        raise RuntimeError(
            f"Could not read the internal font family: {font_path.name}"
        )

    # Fontconfig may return multiple family aliases.
    family = output.splitlines()[0].split(",")[0].strip()

    if not family:
        raise RuntimeError(
            f"Empty font family name in {font_path.name}"
        )

    return family


def install_and_check_fonts():
    for path in FONT_FILES:
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(f"Missing or empty font file: {path.name}")

    families = {
        path.name: font_family(path)
        for path in FONT_FILES
    }

    normalized = {
        unicodedata.normalize("NFKC", value).casefold()
        for value in families.values()
    }

    if len(normalized) != 1:
        details = "; ".join(
            f"{name}={family}"
            for name, family in families.items()
        )
        raise RuntimeError(
            "The three TTF files do not declare the same font family. "
            "Their metadata must be reconciled before rendering. " + details
        )

    family = families[FONT_FILES[0].name]

    FONT_DIR.mkdir(parents=True, exist_ok=True)

    for source in FONT_FILES:
        destination = FONT_DIR / source.name
        shutil.copy2(source, destination)
        destination.chmod(0o644)

    run_command(["fc-cache", "-f"])

    # Confirm that fontconfig resolves the requested family
    # to the regular font installed for this job.
    matched_file = run_command([
        "fc-match",
        "-f",
        "%{file}\n",
        family,
    ]).splitlines()[0]

    expected_file = str((FONT_DIR / FONT_FILES[0].name).resolve())

    if Path(matched_file).resolve() != Path(expected_file):
        raise RuntimeError(
            "Fontconfig did not select the installed IRMitra Regular font. "
            f"Expected {expected_file}; received {matched_file}"
        )

    print("Detected IRMitra internal family:", family)

    for filename, detected_family in families.items():
        print(f"{filename}: {detected_family}")

    return family


def font_element(family):
    escaped = html.escape(family, quote=True)

    return (
        '<w:rFonts'
        f' w:ascii="{escaped}"'
        f' w:hAnsi="{escaped}"'
        f' w:eastAsia="{escaped}"'
        f' w:cs="{escaped}"'
        '/>'
    )


def patch_run_properties(xml_text, font_tag):
    def replace_rpr(match):
        attrs = match.group("attrs")

        if attrs is None:
            attrs = match.group("selfattrs") or ""

        body = match.group("body") or ""

        # Remove previous explicit and theme-based font declarations.
        body = RFONTS_RE.sub("", body)

        # Place w:rFonts first in the run properties.
        return (
            "<w:rPr" + attrs + ">"
            + font_tag
            + body
            + "</w:rPr>"
        )

    xml_text = RPR_RE.sub(replace_rpr, xml_text)

    def add_missing_rpr(match):
        attrs = match.group("attrs")
        body = match.group("body")

        if re.search(r"<w:rPr(?=[\s>])", body):
            return match.group(0)

        return (
            "<w:r" + attrs + ">"
            + "<w:rPr>" + font_tag + "</w:rPr>"
            + body
            + "</w:r>"
        )

    xml_text = RUN_RE.sub(add_missing_rpr, xml_text)

    return xml_text


def patch_docx(family):
    if not SOURCE.is_file():
        raise RuntimeError("input.docx does not exist.")

    if not zipfile.is_zipfile(SOURCE):
        raise RuntimeError("input.docx is not a valid ZIP archive.")

    font_tag = font_element(family)
    changed_parts = 0

    try:
        with zipfile.ZipFile(SOURCE, "r") as source_zip:
            with zipfile.ZipFile(
                TEMP_DOCX,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            ) as target_zip:

                for item in source_zip.infolist():
                    data = source_zip.read(item.filename)

                    if XML_PARTS.fullmatch(item.filename):
                        try:
                            xml_text = data.decode("utf-8")
                        except UnicodeDecodeError as exc:
                            raise RuntimeError(
                                f"Unexpected XML encoding in {item.filename}"
                            ) from exc

                        patched = patch_run_properties(
                            xml_text,
                            font_tag,
                        )

                        # Validate that the modified XML remains well-formed.
                        ET.fromstring(patched.encode("utf-8"))

                        data = patched.encode("utf-8")
                        changed_parts += 1

                    target_zip.writestr(item, data)

        with zipfile.ZipFile(TEMP_DOCX, "r") as check_zip:
            bad_member = check_zip.testzip()

        if bad_member:
            raise RuntimeError(
                f"Patched DOCX archive integrity failed: {bad_member}"
            )

        os.replace(TEMP_DOCX, SOURCE)

    except Exception:
        if TEMP_DOCX.exists():
            TEMP_DOCX.unlink()
        raise

    print("Patched XML parts:", changed_parts)
    print("DOCX font declarations were updated.")


def main():
    family = install_and_check_fonts()
    patch_docx(family)


if __name__ == "__main__":
    main()
