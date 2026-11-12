
import os
import shutil
import subprocess
import unicodedata
import zipfile
from pathlib import Path

from lxml import etree


SOURCE = Path("input.docx")
TEMP_DOCX = Path("input.patched.docx")

FONT_FILES = [
    Path("IRMitra.ttf"),
    Path("IRMitra-Bold.ttf"),
    Path("IRMitra-Italic.ttf"),
]

FONT_DIR = (
    Path.home()
    / ".local"
    / "share"
    / "fonts"
    / "irmitra"
)

W_NS = (
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
)

A_NS = (
    "http://schemas.openxmlformats.org/drawingml/2006/main"
)

M_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/math"
)

W = "{%s}" % W_NS
A = "{%s}" % A_NS
M = "{%s}" % M_NS

W_R = W + "r"
W_RPR = W + "rPr"
W_RFonts = W + "rFonts"
W_SYM = W + "sym"

A_R = A + "r"
A_RPR = A + "rPr"

A_FONT_ELEMENTS = {
    A + "latin",
    A + "ea",
    A + "cs",
    A + "font",
    A + "buFont",
}


def run_command(args):
    result = subprocess.run(
        args,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return result.stdout.strip()


def normalize_family(value):
    return unicodedata.normalize(
        "NFKC",
        value
    ).casefold().strip()


def detect_font_family(path):
    output = run_command([
        "fc-scan",
        "--format",
        "%{family}\n",
        str(path),
    ])

    if not output:
        raise RuntimeError(
            "Cannot detect font family: " + path.name
        )

    family = output.splitlines()[0].split(",")[0].strip()

    if not family:
        raise RuntimeError(
            "Empty font family: " + path.name
        )

    return family


def install_fonts():
    for path in FONT_FILES:
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(
                "Missing or empty font: " + path.name
            )

    families = {
        path.name: detect_font_family(path)
        for path in FONT_FILES
    }

    normalized = {
        normalize_family(value)
        for value in families.values()
    }

    if len(normalized) != 1:
        raise RuntimeError(
            "Regular, Bold and Italic do not share one "
            "internal font family. Rendering stopped. "
            + str(families)
        )

    family = families[FONT_FILES[0].name]

    FONT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    for source in FONT_FILES:
        target = FONT_DIR / source.name
        shutil.copy2(source, target)
        target.chmod(0o644)

    run_command(["fc-cache", "-f"])

    selected = run_command([
        "fc-match",
        "-f",
        "%{file}\n",
        family,
    ]).splitlines()[0]

    expected = str(
        (FONT_DIR / FONT_FILES[0].name).resolve()
    )

    if Path(selected).resolve() != Path(expected).resolve():
        raise RuntimeError(
            "Fontconfig did not select IRMitra Regular. "
            "Expected: " + expected +
            " Selected: " + selected
        )

    print("Selected font family:", family)

    for filename, detected in families.items():
        print(filename + " -> " + detected)

    return family


def set_word_font(rpr, family):
    """
    Replace every WordprocessingML font reference
    with the same explicit IRMitra family.
    """

    for old in list(rpr.findall(W_RFonts)):
        rpr.remove(old)

    fonts = etree.Element(W_RFonts)

    fonts.set(W + "ascii", family)
    fonts.set(W + "hAnsi", family)
    fonts.set(W + "eastAsia", family)
    fonts.set(W + "cs", family)

    # w:rFonts must be the first child of w:rPr.
    rpr.insert(0, fonts)


def force_word_fonts(root, family):
    run_count = 0
    property_count = 0
    symbol_count = 0

    # Apply direct font formatting to every text run.
    # This is the DOCX-level equivalent of selecting all text
    # and assigning IRMitra explicitly.
    for run in root.iter(W_R):
        rpr = run.find(W_RPR)

        if rpr is None:
            rpr = etree.Element(W_RPR)
            run.insert(0, rpr)

        set_word_font(rpr, family)
        run_count += 1

    # Also replace inherited style definitions:
    # document defaults, paragraph styles, character styles,
    # numbering levels and other run-property definitions.
    for rpr in root.iter(W_RPR):
        set_word_font(rpr, family)
        property_count += 1

    # Explicit symbol fonts can bypass normal rFonts settings.
    for symbol in root.iter(W_SYM):
        symbol.set(W + "font", family)
        symbol_count += 1

    return run_count, property_count, symbol_count


def force_drawing_fonts(root, family):
    changed = 0
    run_count = 0

    # Apply explicit font declarations to DrawingML text runs.
    for run in root.iter(A_R):
        rpr = run.find(A_RPR)

        if rpr is None:
            rpr = etree.Element(A_RPR)
            run.insert(0, rpr)

        set_drawing_font_properties(rpr, family)
        run_count += 1

    # Replace DrawingML font declarations everywhere:
    # text boxes, charts, diagrams, SmartArt and theme fonts.
    for element in root.iter():
        if element.tag in A_FONT_ELEMENTS:
            element.set("typeface", family)
            changed += 1

    # Update existing direct formatting containers too.
    for element in root.iter():
        if element.tag in {
            A + "rPr",
            A + "defRPr",
            A + "endParaRPr",
        }:
            set_drawing_font_properties(element, family)

    return changed, run_count


def set_drawing_font_properties(element, family):
    """
    Replace existing DrawingML latin/east-Asian/complex-script
    font declarations. Theme declarations are also overwritten.
    """

    for tag in (A + "latin", A + "ea", A + "cs"):
        matches = element.findall(tag)

        if matches:
            for item in matches:
                item.set("typeface", family)


def force_math_font(root, family):
    changed = 0

    # Set the declared font for Office Math equations.
    for element in root.iter(M + "mathFont"):
        element.set(M + "val", family)
        changed += 1

    return changed


def patch_xml(data, part_name, family):
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
        strip_cdata=False,
        huge_tree=False,
    )

    try:
        root = etree.fromstring(data, parser)
    except Exception as exc:
        raise RuntimeError(
            "XML parsing failed: " + part_name
        ) from exc

    word_runs = 0
    word_properties = 0
    word_symbols = 0
    drawing_fonts = 0
    drawing_runs = 0
    math_fonts = 0

    if root.tag.startswith(W):
        (
            word_runs,
            word_properties,
            word_symbols,
        ) = force_word_fonts(root, family)

    drawing_fonts, drawing_runs = force_drawing_fonts(
        root,
        family
    )

    math_fonts = force_math_font(root, family)

    total = (
        word_runs
        + word_properties
        + word_symbols
        + drawing_fonts
        + drawing_runs
        + math_fonts
    )

    if not total:
        return data, 0, 0, 0, 0, 0, 0

    has_declaration = data.lstrip().startswith(b"<?xml")

    patched = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=has_declaration,
        pretty_print=False,
    )

    # Validate the serialized XML before writing it to DOCX.
    etree.fromstring(patched, parser)

    return (
        patched,
        word_runs,
        word_properties,
        word_symbols,
        drawing_fonts,
        drawing_runs,
        math_fonts,
    )


def ensure_do_not_expand_shift_return(data):
    """
    Preserve all paragraph and line-break structures.
    Enable the OOXML compatibility flag that prevents
    justification expansion on lines ending in Shift+Enter.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
        strip_cdata=False,
    )

    try:
        root = etree.fromstring(data, parser)
    except Exception as exc:
        raise RuntimeError(
            "Could not parse word/settings.xml."
        ) from exc

    if root.tag != W + "settings":
        raise RuntimeError(
            "Unexpected root element in word/settings.xml."
        )

    compat_nodes = root.findall(W + "compat")

    if len(compat_nodes) > 1:
        raise RuntimeError(
            "Multiple w:compat elements found. "
            "Refusing to modify an ambiguous settings structure."
        )

    # Reuse the existing compatibility container.
    if compat_nodes:
        compat = compat_nodes[0]

    else:
        # Create one w:compat element at its correct
        # position in the Word settings sequence.
        compat = etree.Element(W + "compat")

        math_ns = (
            "{http://schemas.openxmlformats.org/"
            "officeDocument/2006/math}"
        )

        settings_successors = {
            W + name
            for name in (
                "docVars",
                "rsids",
                "uiCompat97To2003",
                "attachedSchema",
                "themeFontLang",
                "clrSchemeMapping",
                "doNotIncludeSubdocsInStats",
                "doNotAutoCompressPictures",
                "forceUpgrade",
                "captions",
                "readModeInkLockDown",
                "smartTagType",
                "shapeDefaults",
                "doNotEmbedSmartTags",
                "decimalSymbol",
                "listSeparator",
            )
        }

        settings_successors.add(math_ns + "mathPr")

        insert_at = len(root)

        for index, child in enumerate(root):
            if child.tag in settings_successors:
                insert_at = index
                break

        root.insert(insert_at, compat)

    flag_tag = W + "doNotExpandShiftReturn"
    existing = compat.findall(flag_tag)

    changed = False

    if existing:
        # Enable the existing flag explicitly.
        flag = existing[0]

        old_value = flag.get(W + "val")

        if old_value not in ("true", "1", "on"):
            flag.set(W + "val", "true")
            changed = True

        # The schema allows only one instance.
        for duplicate in existing[1:]:
            compat.remove(duplicate)
            changed = True

    else:
        flag = etree.Element(flag_tag)
        flag.set(W + "val", "true")

        # Existing CT_Compat sequence:
        # doNotExpandShiftReturn comes after ulTrailSpace
        # and before spacingInWholePoints and subsequent flags.
        preceding = {
            W + name
            for name in (
                "useSingleBorderforContiguousCells",
                "wpJustification",
                "noTabHangInd",
                "noLeading",
                "spaceForUL",
                "noColumnBalance",
                "balanceSingleByteDoubleByteWidth",
                "noExtraLineSpacing",
                "doNotLeaveBackslashAlone",
                "ulTrailSpace",
            )
        }

        following = {
            W + name
            for name in (
                "spacingInWholePoints",
                "lineWrapLikeWord6",
                "printBodyTextBeforeHeader",
                "printColBlack",
                "doNotBreakWrappedTables",
                "adjustLineHeightInTable",
                "useFELayout",
                "useNormalStyleForList",
                "doNotUseIndentAsNumberingTabStop",
                "useAltKinsokuLineBreakRules",
                "allowSpaceOfSameStyleInTable",
                "compatSetting",
            )
        }

        first_following = next(
            (
                i
                for i, child in enumerate(compat)
                if child.tag in following
            ),
            None,
        )

        last_preceding = max(
            (
                i
                for i, child in enumerate(compat)
                if child.tag in preceding
            ),
            default=-1,
        )

        if (
            first_following is not None
            and last_preceding >= first_following
        ):
            raise RuntimeError(
                "Existing w:compat children are not in schema order. "
                "Refusing to reorder existing compatibility settings."
            )

        if first_following is not None:
            insert_at = first_following

        elif last_preceding >= 0:
            insert_at = last_preceding + 1

        else:
            insert_at = len(compat)

        compat.insert(insert_at, flag)
        changed = True

    had_declaration = data.lstrip().startswith(b"<?xml")

    patched = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=had_declaration,
        pretty_print=False,
    )

    # Validate the resulting XML.
    etree.fromstring(patched, parser)

    return patched, changed


def patch_docx(family):
    if not SOURCE.is_file():
        raise RuntimeError("input.docx does not exist.")

    if not zipfile.is_zipfile(SOURCE):
        raise RuntimeError("Invalid DOCX archive.")

    
    totals = {
        "parts": 0,
        "word_runs": 0,
        "word_properties": 0,
        "word_symbols": 0,
        "drawing_fonts": 0,
        "drawing_runs": 0,
        "math_fonts": 0,
        "shift_return_setting": False,
    }

    settings_part_found = False

    try:
        with zipfile.ZipFile(SOURCE, "r") as source_zip:
            with zipfile.ZipFile(
                TEMP_DOCX,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            ) as target_zip:

                for item in source_zip.infolist():
                    data = source_zip.read(item.filename)

                    # Process all XML parts under word/.
                    # No restriction to document.xml alone.
                    is_word_xml = (
                        item.filename.startswith("word/")
                        and item.filename.lower().endswith(".xml")
                    )

                    
                    if is_word_xml:
                        (
                            data,
                            wr,
                            wp,
                            ws,
                            df,
                            dr,
                            mf,
                        ) = patch_xml(
                            data,
                            item.filename,
                            family,
                        )

                        settings_changed = False

                        # Apply the compatibility setting only
                        # to the document settings part.
                        if item.filename == "word/settings.xml":

                            data, settings_changed = (
                                ensure_do_not_expand_shift_return(data)
                            )

                            settings_part_found = True
                            totals["shift_return_setting"] = True

                        totals["word_runs"] += wr
                        totals["word_properties"] += wp
                        totals["word_symbols"] += ws
                        totals["drawing_fonts"] += df
                        totals["drawing_runs"] += dr
                        totals["math_fonts"] += mf

                        if (
                            wr or wp or ws or df or dr or mf
                            or settings_changed
                        ):
                            totals["parts"] += 1

                    target_zip.writestr(item, data)

        if not settings_part_found:
            raise RuntimeError(
                "The exported DOCX does not contain word/settings.xml. "
                "Rendering stopped without modifying paragraph structures."
            )

        with zipfile.ZipFile(TEMP_DOCX, "r") as check_zip:
            bad_member = check_zip.testzip()

        if bad_member:
            raise RuntimeError(
                "DOCX integrity check failed: " + bad_member
            )

        os.replace(TEMP_DOCX, SOURCE)

    except Exception:
        if TEMP_DOCX.exists():
            TEMP_DOCX.unlink()
        raise

    print("Modified XML parts:", totals["parts"])
    print("Word text runs:", totals["word_runs"])
    print("Word run-property definitions:", totals["word_properties"])
    print("Word symbol declarations:", totals["word_symbols"])
    print("DrawingML font declarations:", totals["drawing_fonts"])
    print("DrawingML text runs:", totals["drawing_runs"])
    print("Office Math font declarations:", totals["math_fonts"])
    print("Global IRMitra font enforcement completed.")
    print(
        "Shift+Enter line justification protection:",
        "ENABLED" if totals["shift_return_setting"] else "NOT SET"
    )



def main():
    family = install_fonts()
    patch_docx(family)


if __name__ == "__main__":
    main()
