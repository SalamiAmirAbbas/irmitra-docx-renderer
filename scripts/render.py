
import copy
import json
import os
import posixpath
import re
import shutil
import subprocess
import unicodedata
import zipfile

from pathlib import Path

from lxml import etree


# ============================================================================
# Paths and namespaces
# ============================================================================

SOURCE = Path("input.docx")
TEMP_DOCX = Path("input.patched.docx")
OPTIONS_FILE = Path("job-options.json")

FONT_DIR = (
    Path.home()
    / ".local"
    / "share"
    / "fonts"
    / "irmitra"
)

FONT_PATHS = {
    "irmitraRegular": Path("IRMitra.ttf"),
    "irmitraBold": Path("IRMitra-Bold.ttf"),
    "irmitraItalic": Path("IRMitra-Italic.ttf"),
    "bmitraRegular": Path("BMitra.ttf"),
    "bmitraBold": Path("BMitraBd.ttf"),
}

W_NS = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

A_NS = (
    "http://schemas.openxmlformats.org/"
    "drawingml/2006/main"
)

M_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/math"
)

R_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

PKG_REL_NS = (
    "http://schemas.openxmlformats.org/"
    "package/2006/relationships"
)

CT_NS = (
    "http://schemas.openxmlformats.org/"
    "package/2006/content-types"
)

XML_NS = "http://www.w3.org/XML/1998/namespace"

W = "{" + W_NS + "}"
A = "{" + A_NS + "}"
M = "{" + M_NS + "}"
R = "{" + R_NS + "}"
PKG = "{" + PKG_REL_NS + "}"
CT = "{" + CT_NS + "}"

W_P = W + "p"
W_R = W + "r"
W_RPR = W + "rPr"
W_RFONTS = W + "rFonts"
W_T = W + "t"
W_PPR = W + "pPr"
W_SPACING = W + "spacing"
W_STYLES = W + "styles"
W_STYLE = W + "style"
W_FOOTNOTE = W + "footnote"
W_FOOTNOTE_REF = W + "footnoteRef"
W_FOOTNOTE_REFERENCE = W + "footnoteReference"
W_FLD_SIMPLE = W + "fldSimple"
W_FLD_CHAR = W + "fldChar"
W_INSTR_TEXT = W + "instrText"
W_SECTPR = W + "sectPr"
W_SECTPR_CHANGE = W + "sectPrChange"
W_FOOTER_REF = W + "footerReference"
W_PGNUMTYPE = W + "pgNumType"

A_R = A + "r"
A_RPR = A + "rPr"
A_DEFRPR = A + "defRPr"
A_ENDPARARPR = A + "endParaRPr"
A_PPR = A + "pPr"

R_DIGITS = re.compile(
    r"[0-9\u0660-\u0669\u06F0-\u06F9]+"
)

PAGE_INSTRUCTION = re.compile(
    r"(?:^|\s)PAGE(?:\s|$)",
    re.IGNORECASE
)

WORD_XML_ROOT = "word/"


# ============================================================================
# XML helpers
# ============================================================================

def xml_parser():
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
        strip_cdata=False,
        huge_tree=False,
    )


def parse_xml(data, label):
    try:
        return etree.fromstring(data, xml_parser())
    except Exception as exc:
        raise RuntimeError(
            "XML parsing failed: " + label
        ) from exc


def serialize_xml(root, original_data=None):
    has_declaration = (
        original_data is not None
        and original_data.lstrip().startswith(b"<?xml")
    )

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=has_declaration,
        pretty_print=False,
    )


def local_name(tag):
    if not isinstance(tag, str):
        return ""

    return etree.QName(tag).localname


def run_command(args):
    result = subprocess.run(
        args,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    return result.stdout.strip()


# ============================================================================
# Font metadata and installation
# ============================================================================

def normalize_family(value):
    return unicodedata.normalize(
        "NFKC",
        str(value)
    ).casefold().strip()


def detect_font_metadata(path):
    output = run_command([
        "fc-scan",
        "--format",
        "%{family}\n%{style}\n%{weight}\n",
        str(path),
    ])

    lines = output.splitlines()

    if not lines:
        raise RuntimeError(
            "Cannot detect internal font metadata: " + path.name
        )

    family = lines[0].split(",")[0].strip()
    style = lines[1].strip() if len(lines) > 1 else "Unknown"
    weight = lines[2].strip() if len(lines) > 2 else "Unknown"

    if not family:
        raise RuntimeError(
            "Empty internal font family: " + path.name
        )

    return {
        "family": family,
        "style": style,
        "weight": weight,
    }


def install_and_check_fonts():
    for key, path in FONT_PATHS.items():
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(
                "Missing or empty font file: " + path.name
            )

    metadata = {
        key: detect_font_metadata(path)
        for key, path in FONT_PATHS.items()
    }

    # The three IRMitra faces must belong to one internal family.
    ir_families = {
        normalize_family(metadata[key]["family"])
        for key in (
            "irmitraRegular",
            "irmitraBold",
            "irmitraItalic",
        )
    }

    if len(ir_families) != 1:
        raise RuntimeError(
            "The three IRMitra TTF files do not declare "
            "one common internal font family. "
            + json.dumps(metadata, ensure_ascii=False)
        )

    FONT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    for key, source in FONT_PATHS.items():
        destination = FONT_DIR / source.name
        shutil.copy2(source, destination)
        destination.chmod(0o644)

    run_command(["fc-cache", "-f"])

    # Validate regular IRMitra selection.
    ir_family = metadata["irmitraRegular"]["family"]

    selected_ir = run_command([
        "fc-match",
        "-f",
        "%{file}\n",
        ir_family,
    ]).splitlines()[0]

    expected_ir = str(
        (FONT_DIR / FONT_PATHS["irmitraRegular"].name).resolve()
    )

    if Path(selected_ir).resolve() != Path(expected_ir).resolve():
        raise RuntimeError(
            "Fontconfig did not select IRMitra Regular. "
            f"Expected: {expected_ir}; selected: {selected_ir}"
        )

    # Validate that BMitra Regular is available.
    bm_family = metadata["bmitraRegular"]["family"]

    selected_bm = run_command([
        "fc-match",
        "-f",
        "%{file}\n",
        bm_family,
    ]).splitlines()[0]

    expected_bm = str(
        (FONT_DIR / FONT_PATHS["bmitraRegular"].name).resolve()
    )

    if Path(selected_bm).resolve() != Path(expected_bm).resolve():
        raise RuntimeError(
            "Fontconfig did not select BMitra Regular. "
            f"Expected: {expected_bm}; selected: {selected_bm}"
        )

    print("Font metadata detected:")

    for key, item in metadata.items():
        print(
            key + ": family=" + item["family"]
            + "; style=" + item["style"]
            + "; weight=" + item["weight"]
        )

    return {
        "irmitra": metadata["irmitraRegular"]["family"],
        "bmitra_regular": metadata["bmitraRegular"]["family"],
        "bmitra_bold": metadata["bmitraBold"]["family"],
        "metadata": metadata,
    }


# ============================================================================
# Configuration
# ============================================================================

def load_options():
    if not OPTIONS_FILE.is_file():
        raise RuntimeError(
            "job-options.json is missing."
        )

    try:
        options = json.loads(
            OPTIONS_FILE.read_text(encoding="utf-8")
        )
    except Exception as exc:
        raise RuntimeError(
            "Cannot parse job-options.json."
        ) from exc

    if not isinstance(options, dict):
        raise RuntimeError(
            "Customization options must be a JSON object."
        )

    return options


def get_boolean(options, key):
    return options.get(key) is True


def get_paragraph_config(options, key):
    paragraphs = options.get("paragraphs", {})

    if not isinstance(paragraphs, dict):
        return {}

    value = paragraphs.get(key, {})

    return value if isinstance(value, dict) else {}


def bounded_number(value, minimum, maximum, fallback):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback

    if not (number == number and abs(number) != float("inf")):
        return fallback

    return max(minimum, min(maximum, number))


# ============================================================================
# Style index and style inheritance
# ============================================================================

def normalize_style_name(value):
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(value or "").casefold()
    )


def bucket_from_name(value):
    name = normalize_style_name(value)

    aliases = {
        "heading1": "heading1",
        "heading2": "heading2",
        "heading3": "heading3",
        "heading4": "heading4",
        "normal": "normal",
        "normaltext": "normal",
    }

    return aliases.get(name)


def build_style_index(styles_data):
    if not styles_data:
        return {
            "styles": {},
            "default_paragraph": None,
            "default_rpr": None,
        }

    root = parse_xml(styles_data, "word/styles.xml")

    styles = {}
    default_paragraph = None
    default_rpr = None

    defaults = root.find(W + "docDefaults")

    if defaults is not None:
        rpr_default = defaults.find(W + "rPrDefault")

        if rpr_default is not None:
            default_rpr = rpr_default.find(W_RPR)

    for style in root.findall(W_STYLE):
        style_id = style.get(W + "styleId")

        if not style_id:
            continue

        name_element = style.find(W + "name")
        name = (
            name_element.get(W + "val")
            if name_element is not None
            else style_id
        )

        based_on_element = style.find(W + "basedOn")
        based_on = (
            based_on_element.get(W + "val")
            if based_on_element is not None
            else None
        )

        rpr = style.find(W_RPR)
        ppr = style.find(W_PPR)

        styles[style_id] = {
            "type": style.get(W + "type"),
            "name": name,
            "based_on": based_on,
            "rpr": rpr,
            "ppr": ppr,
            "bucket": (
                bucket_from_name(name)
                or bucket_from_name(style_id)
            ),
        }

        if (
            style.get(W + "type") == "paragraph"
            and style.get(W + "default") in ("1", "true", "on")
        ):
            default_paragraph = style_id

    return {
        "styles": styles,
        "default_paragraph": default_paragraph,
        "default_rpr": default_rpr,
    }


def style_chain(style_index, style_id):
    styles = style_index["styles"]
    visited = set()
    chain = []

    current = style_id

    while current and current not in visited:
        visited.add(current)

        info = styles.get(current)

        if not info:
            break

        chain.append(info)
        current = info.get("based_on")

    return chain


def paragraph_style_id(paragraph, style_index):
    ppr = paragraph.find(W_PPR)

    if ppr is not None:
        pstyle = ppr.find(W + "pStyle")

        if pstyle is not None:
            return pstyle.get(W + "val")

    return style_index.get("default_paragraph")


def paragraph_bucket(paragraph, style_index):
    style_id = paragraph_style_id(
        paragraph,
        style_index
    )

    if not style_id:
        return "normal"

    for info in style_chain(style_index, style_id):
        if info.get("bucket"):
            return info["bucket"]

    # An explicitly named, unknown custom style is not
    # automatically treated as Normal.
    if style_id in style_index["styles"]:
        return None

    return "normal"


def property_from_rpr(rpr, property_tags):
    if rpr is None:
        return False, None

    for tag in property_tags:
        element = rpr.find(tag)

        if element is not None:
            return True, element

    return False, None


def on_off_value(element):
    if element is None:
        return None

    value = element.get(W + "val")

    if value is None:
        return True

    return str(value).casefold() not in (
        "0",
        "false",
        "off",
        "no",
    )


def effective_run_property(
    run,
    paragraph,
    style_index,
    property_tags
):
    # 1. Direct run formatting.
    direct_rpr = run.find(W_RPR)

    found, element = property_from_rpr(
        direct_rpr,
        property_tags
    )

    if found:
        return True, element

    # 2. Character style.
    if direct_rpr is not None:
        rstyle = direct_rpr.find(W + "rStyle")

        if rstyle is not None:
            char_style_id = rstyle.get(W + "val")

            for info in style_chain(
                style_index,
                char_style_id
            ):
                found, element = property_from_rpr(
                    info.get("rpr"),
                    property_tags
                )

                if found:
                    return True, element

    # 3. Paragraph style chain.
    pstyle_id = paragraph_style_id(
        paragraph,
        style_index
    )

    for info in style_chain(
        style_index,
        pstyle_id
    ):
        found, element = property_from_rpr(
            info.get("rpr"),
            property_tags
        )

        if found:
            return True, element

    # 4. Document defaults.
    found, element = property_from_rpr(
        style_index.get("default_rpr"),
        property_tags
    )

    if found:
        return True, element

    return False, None


def run_is_bold(run, paragraph, style_index):
    found, element = effective_run_property(
        run,
        paragraph,
        style_index,
        (W + "b", W + "bCs")
    )

    if not found:
        return False

    return on_off_value(element) is True


def run_effective_size(run, paragraph, style_index):
    found, element = effective_run_property(
        run,
        paragraph,
        style_index,
        (W + "sz", W + "szCs")
    )

    if not found:
        return 22  # Word default fallback: 11 pt.

    try:
        return int(element.get(W + "val"))
    except (TypeError, ValueError):
        return 22


# ============================================================================
# WordprocessingML font helpers
# ============================================================================

def ensure_run_properties(run):
    rpr = run.find(W_RPR)

    if rpr is None:
        rpr = etree.Element(W_RPR)
        run.insert(0, rpr)

    return rpr


def set_run_font(rpr, family):
    old_fonts = rpr.findall(W_RFONTS)

    for old in old_fonts:
        rpr.remove(old)

    fonts = etree.Element(W_RFONTS)

    fonts.set(W + "ascii", family)
    fonts.set(W + "hAnsi", family)
    fonts.set(W + "eastAsia", family)
    fonts.set(W + "cs", family)

    rpr.insert(0, fonts)


def set_run_size(rpr, points):
    half_points = int(round(float(points) * 2))

    for tag in (W + "sz", W + "szCs"):
        element = rpr.find(tag)

        if element is None:
            element = etree.SubElement(rpr, tag)

        element.set(W + "val", str(half_points))


def set_run_color_black(rpr):
    color = rpr.find(W + "color")

    if color is None:
        color = etree.SubElement(rpr, W + "color")

    color.set(W + "val", "000000")

    for attribute in (
        "themeColor",
        "themeTint",
        "themeShade",
    ):
        color.attrib.pop(W + attribute, None)


def force_word_fonts(root, family):
    run_count = 0
    property_count = 0
    symbol_count = 0

    for run in root.iter(W_R):
        rpr = ensure_run_properties(run)
        set_run_font(rpr, family)
        run_count += 1

    # Also replace styles, defaults, numbering definitions,
    # paragraph-mark properties and inherited run properties.
    for rpr in root.iter(W_RPR):
        set_run_font(rpr, family)
        property_count += 1

    # Explicit symbol font declarations can bypass rFonts.
    for symbol in root.iter(W + "sym"):
        symbol.set(W + "font", family)
        symbol_count += 1

    return run_count, property_count, symbol_count


# ============================================================================
# DrawingML font and color helpers
# ============================================================================

def set_drawing_font_properties(element, family):
    for tag in (
        A + "latin",
        A + "ea",
        A + "cs",
    ):
        matches = element.findall(tag)

        if matches:
            for item in matches:
                item.set("typeface", family)
        else:
            child = etree.Element(tag)
            child.set("typeface", family)
            element.insert(0, child)


def force_drawing_fonts(root, family):
    changed = 0

    for run in root.iter(A_R):
        rpr = run.find(A_RPR)

        if rpr is None:
            rpr = etree.Element(A_RPR)
            run.insert(0, rpr)

        set_drawing_font_properties(rpr, family)
        changed += 1

    for element in root.iter():
        if element.tag in (
            A_RPR,
            A_DEFRPR,
            A_ENDPARARPR,
        ):
            set_drawing_font_properties(
                element,
                family
            )
            changed += 1

    return changed


def set_drawing_solid_fill_black(element):
    solid = element.find(A + "solidFill")

    if solid is None:
        solid = etree.Element(A + "solidFill")
        element.insert(0, solid)
    else:
        for child in list(solid):
            solid.remove(child)

    etree.SubElement(
        solid,
        A + "srgbClr",
        val="000000"
    )


def force_black_text(root):
    changed = 0

    # WordprocessingML runs, styles, list markers and defaults.
    for rpr in root.iter(W_RPR):
        set_run_color_black(rpr)
        changed += 1

    # DrawingML text runs and paragraph-level text properties.
    text_properties = (
        A_RPR,
        A_DEFRPR,
        A_ENDPARARPR,
        A_PPR,
    )

    for element in root.iter():
        if element.tag not in text_properties:
            continue

        set_drawing_solid_fill_black(element)

        # Bullet color is separate from the normal text color.
        bullet_color = element.find(A + "buClr")

        if bullet_color is not None:
            for child in list(bullet_color):
                bullet_color.remove(child)

            etree.SubElement(
                bullet_color,
                A + "srgbClr",
                val="000000"
            )

        changed += 1

    return changed


# ============================================================================
# Background shading removal
# ============================================================================

def remove_background_shading(root):
    removed = 0

    for shading in list(root.iter(W + "shd")):
        parent = shading.getparent()

        if parent is not None:
            parent.remove(shading)
            removed += 1

    # Dedicated text highlighting (w:highlight) is deliberately preserved.
    # It is a separate WordprocessingML property from background shading.

    return removed


# ============================================================================
# Paragraph spacing and minimum Normal font size
# ============================================================================

def ensure_paragraph_properties(paragraph):
    ppr = paragraph.find(W_PPR)

    if ppr is None:
        ppr = etree.Element(W_PPR)
        paragraph.insert(0, ppr)

    return ppr

def insert_ppr_child_ordered(ppr, child):
    # ترتیب استاندارد فرزندان w:pPr در WordprocessingML
    order = {
        W + "pStyle": 0,
        W + "keepNext": 1,
        W + "keepLines": 2,
        W + "pageBreakBefore": 3,
        W + "framePr": 4,
        W + "widowControl": 5,
        W + "numPr": 6,
        W + "suppressLineNumbers": 7,
        W + "pBdr": 8,
        W + "shd": 9,
        W + "tabs": 10,
        W + "suppressAutoHyphens": 11,
        W + "kinsoku": 12,
        W + "wordWrap": 13,
        W + "overflowPunct": 14,
        W + "topLinePunct": 15,
        W + "autoSpaceDE": 16,
        W + "autoSpaceDN": 17,
        W + "bidi": 18,
        W + "adjustRightInd": 19,
        W + "snapToGrid": 20,
        W + "spacing": 21,
        W + "ind": 22,
        W + "contextualSpacing": 23,
        W + "mirrorIndents": 24,
        W + "suppressOverlap": 25,
        W + "jc": 26,
        W + "textDirection": 27,
        W + "textAlignment": 28,
        W + "outlineLvl": 29,
        W + "divId": 30,
        W + "cnfStyle": 31,
        W + "rPr": 32,
        W + "sectPr": 33,
        W + "pPrChange": 34,
    }

    child_order = order.get(child.tag)

    if child_order is None:
        ppr.append(child)
        return

    for index, existing in enumerate(ppr):
        existing_order = order.get(existing.tag)

        if (
            existing_order is not None
            and existing_order > child_order
        ):
            ppr.insert(index, child)
            return

    ppr.append(child)


def get_style_paragraph_property(paragraph, style_index, tag):
    style_id = paragraph_style_id(
        paragraph,
        style_index
    )

    for info in style_chain(style_index, style_id):
        ppr = info.get("ppr")

        if ppr is None:
            continue

        element = ppr.find(tag)

        if element is not None:
            return element

    return None


def set_ppr_on_off(paragraph, tag, enabled):
    ppr = ensure_paragraph_properties(paragraph)
    element = ppr.find(tag)

    if element is None:
        element = etree.Element(tag)
        insert_ppr_child_ordered(ppr, element)

    element.set(
        W + "val",
        "1" if enabled else "0"
    )

    return element


def apply_advanced_paragraph_controls(
    document_root,
    style_index
):
    # ویژگی‌های صفحه‌آرایی که باید از Style به پاراگراف منتقل شوند.
    pagination_tags = (
        W + "keepNext",
        W + "keepLines",
        W + "pageBreakBefore",
        W + "widowControl",
    )

    report = {
        "paragraphsChecked": 0,
        "widowControlEnabled": 0,
        "headingKeepNextEnabled": 0,
        "inheritedControlsMaterialized": 0,
        "nonHeadingKeepNextPreserved": 0,
    }

    heading_buckets = (
        "heading2",
        "heading3",
        "heading4",
    )

    for paragraph in document_root.iter(W_P):
        report["paragraphsChecked"] += 1

        ppr = paragraph.find(W_PPR)

        # حفظ ویژگی‌های به‌ارث‌رسیده از Style پاراگراف.
        # ویژگی مستقیم پاراگراف همیشه اولویت دارد.
        for tag in pagination_tags:
            if ppr is not None and ppr.find(tag) is not None:
                continue

            inherited = get_style_paragraph_property(
                paragraph,
                style_index,
                tag
            )

            if inherited is None:
                continue

            ppr = ensure_paragraph_properties(paragraph)

            insert_ppr_child_ordered(
                ppr,
                copy.deepcopy(inherited)
            )

            report["inheritedControlsMaterialized"] += 1

        # معادل avoidWidowAndOrphan = true در Google Docs API.
        # برای همه پاراگراف‌های بدنه، از جمله پاراگراف‌های جدول.
        set_ppr_on_off(
            paragraph,
            W + "widowControl",
            True
        )

        report["widowControlEnabled"] += 1

        # مطابق LineSpacing.gs:
        # Keep with next برای تیترهای سطح ۲ تا ۴.
        bucket = paragraph_bucket(
            paragraph,
            style_index
        )

        if bucket in heading_buckets:
            set_ppr_on_off(
                paragraph,
                W + "keepNext",
                True
            )

            report["headingKeepNextEnabled"] += 1

        # Keep with next سایر پاراگراف‌ها، از جمله مرجع‌های
        # سلسله‌مراتبی فهرست‌ها، فقط در صورت وجود در سند اصلی حفظ می‌شود.
        # این ویژگی به تمام پاراگراف‌های فهرست تعمیم داده نمی‌شود.
        final_ppr = paragraph.find(W_PPR)

        if final_ppr is not None:
            keep_next = final_ppr.find(W + "keepNext")

            if (
                keep_next is not None
                and on_off_value(keep_next)
                and bucket not in heading_buckets
            ):
                report["nonHeadingKeepNextPreserved"] += 1

    return report



def set_paragraph_spacing(paragraph, config):
    ppr = ensure_paragraph_properties(paragraph)

    spacing = ppr.find(W_SPACING)

    if spacing is None:
        spacing = etree.Element(W_SPACING)
    else:
        ppr.remove(spacing)

    insert_ppr_child_ordered(ppr, spacing)

    before = bounded_number(
        config.get("before"),
        0,
        200,
        0
    )

    after = bounded_number(
        config.get("after"),
        0,
        200,
        0
    )

    line_spacing = bounded_number(
        config.get("lineSpacing"),
        0.5,
        3,
        1
    )

    spacing.set(
        W + "before",
        str(int(round(before * 20)))
    )

    spacing.set(
        W + "after",
        str(int(round(after * 20)))
    )

    # Word uses 240 units per line for lineRule="auto".
    spacing.set(
        W + "line",
        str(int(round(line_spacing * 240)))
    )

    spacing.set(
        W + "lineRule",
        "auto"
    )

    # Prevent automatic spacing flags from overriding explicit values.
    for attribute in (
        "beforeLines",
        "afterLines",
        "beforeAutospacing",
        "afterAutospacing",
    ):
        spacing.attrib.pop(W + attribute, None)



def enforce_minimum_normal_font_size(
    paragraph,
    style_index,
    minimum_points
):
    minimum_half_points = int(
        round(float(minimum_points) * 2)
    )

    for run in paragraph.iter(W_R):
        current = run_effective_size(
            run,
            paragraph,
            style_index
        )

        if current < minimum_half_points:
            rpr = ensure_run_properties(run)

            set_run_size(
                rpr,
                minimum_points
            )


def apply_paragraph_formatting(
    document_root,
    style_index,
    options
):
    if not get_boolean(options, "paragraphFormatting"):
        return {
            "paragraphs": 0,
            "normalRunsRaised": 0,
        }

    config_map = options.get("paragraphs", {})

    if not isinstance(config_map, dict):
        config_map = {}

    minimum = bounded_number(
        options.get("minNormalFontSize"),
        6,
        72,
        14
    )

    changed_paragraphs = 0
    raised_runs = 0

    for paragraph in document_root.iter(W_P):
        bucket = paragraph_bucket(
            paragraph,
            style_index
        )

        if bucket not in (
            "heading1",
            "heading2",
            "heading3",
            "heading4",
            "normal",
        ):
            continue

        config = config_map.get(bucket, {})

        if not isinstance(config, dict):
            continue

        set_paragraph_spacing(
            paragraph,
            config
        )

        changed_paragraphs += 1

        if bucket == "normal":
            before = sum(
                1
                for run in paragraph.iter(W_R)
                if run_effective_size(
                    run,
                    paragraph,
                    style_index
                ) < int(round(minimum * 2))
            )

            enforce_minimum_normal_font_size(
                paragraph,
                style_index,
                minimum
            )

            raised_runs += before

    return {
        "paragraphs": changed_paragraphs,
        "normalRunsRaised": raised_runs,
    }


# ============================================================================
# Shift+Enter compatibility
# ============================================================================

def ensure_do_not_expand_shift_return(data):
    root = parse_xml(
        data,
        "word/settings.xml"
    )

    if root.tag != W + "settings":
        raise RuntimeError(
            "Unexpected root element in word/settings.xml."
        )

    compat = root.find(W + "compat")

    if compat is None:
        compat = etree.Element(W + "compat")

        successors = {
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

        insert_at = next(
            (
                index
                for index, child in enumerate(root)
                if child.tag in successors
            ),
            len(root)
        )

        root.insert(insert_at, compat)

    flag_tag = W + "doNotExpandShiftReturn"
    existing = compat.findall(flag_tag)

    if existing:
        flag = existing[0]
        flag.set(W + "val", "true")

        for duplicate in existing[1:]:
            compat.remove(duplicate)

    else:
        flag = etree.Element(flag_tag)
        flag.set(W + "val", "true")

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

        insert_at = next(
            (
                index
                for index, child in enumerate(compat)
                if child.tag in following
            ),
            len(compat)
        )

        compat.insert(insert_at, flag)

    return serialize_xml(root, data)


# ============================================================================
# PAGE field removal
# ============================================================================

def is_page_instruction(instruction):
    return bool(
        PAGE_INSTRUCTION.search(
            str(instruction or "").strip()
        )
    )


def remove_simple_page_fields(root):
    removed = 0

    for field in list(root.iter(W_FLD_SIMPLE)):
        instruction = (
            field.get(W + "instr")
            or field.get("instr")
            or ""
        )

        if is_page_instruction(instruction):
            parent = field.getparent()

            if parent is not None:
                parent.remove(field)
                removed += 1

    return removed


def run_field_char_types(run):
    return [
        element.get(W + "fldCharType")
        for element in run.iter(W_FLD_CHAR)
    ]


def remove_complex_page_fields(root):
    removed = 0

    for paragraph in root.iter(W_P):
        runs = list(paragraph.iter(W_R))
        index = 0

        while index < len(runs):
            current_types = run_field_char_types(
                runs[index]
            )

            if "begin" not in current_types:
                index += 1
                continue

            depth = 0
            field_runs = []
            instruction_parts = []
            started = False
            end_index = index

            while end_index < len(runs):
                run = runs[end_index]
                field_runs.append(run)

                for text_node in run.iter(W_INSTR_TEXT):
                    if text_node.text:
                        instruction_parts.append(
                            text_node.text
                        )

                for field_type in run_field_char_types(run):
                    if field_type == "begin":
                        depth += 1
                        started = True

                    elif field_type == "end" and started:
                        depth -= 1

                end_index += 1

                if started and depth == 0:
                    break

            instruction = " ".join(instruction_parts)

            if started and depth == 0 and is_page_instruction(instruction):
                for run in field_runs:
                    parent = run.getparent()

                    if parent is not None:
                        parent.remove(run)
                        removed += 1

            index = max(end_index, index + 1)

    return removed


def remove_page_fields(root):
    simple = remove_simple_page_fields(root)
    complex_fields = remove_complex_page_fields(root)

    return simple + complex_fields


# ============================================================================
# Footer and PAGE field creation
# ============================================================================

def append_page_field_paragraph(footer_root, bmitra_family):
    paragraph = etree.SubElement(
        footer_root,
        W_P
    )

    ppr = etree.SubElement(
        paragraph,
        W_PPR
    )

    alignment = etree.SubElement(
        ppr,
        W + "jc"
    )

    alignment.set(
        W + "val",
        "center"
    )

    def new_run():
        run = etree.SubElement(
            paragraph,
            W_R
        )

        rpr = etree.SubElement(
            run,
            W_RPR
        )

        set_run_font(
            rpr,
            bmitra_family
        )

        set_run_size(
            rpr,
            12
        )

        # The page number is explicitly regular.
        for tag in (W + "b", W + "bCs"):
            existing = rpr.find(tag)

            if existing is not None:
                rpr.remove(existing)

        bold_off = etree.SubElement(
            rpr,
            W + "b"
        )
        bold_off.set(W + "val", "0")

        bold_cs_off = etree.SubElement(
            rpr,
            W + "bCs"
        )
        bold_cs_off.set(W + "val", "0")

        lang = etree.SubElement(
            rpr,
            W + "lang"
        )

        lang.set(W + "val", "en-US")
        lang.set(W + "bidi", "en-US")
        lang.set(W + "eastAsia", "en-US")

        return run

    begin_run = new_run()

    begin = etree.SubElement(
        begin_run,
        W_FLD_CHAR
    )

    begin.set(
        W + "fldCharType",
        "begin"
    )

    begin.set(
        W + "dirty",
        "true"
    )

    instruction_run = new_run()

    instruction = etree.SubElement(
        instruction_run,
        W_INSTR_TEXT
    )

    instruction.text = " PAGE \\* Arabic "

    instruction.set(
        "{" + XML_NS + "}space",
        "preserve"
    )

    separate_run = new_run()

    separate = etree.SubElement(
        separate_run,
        W_FLD_CHAR
    )

    separate.set(
        W + "fldCharType",
        "separate"
    )

    result_run = new_run()

    result = etree.SubElement(
        result_run,
        W_T
    )

    result.text = "1"

    end_run = new_run()

    end = etree.SubElement(
        end_run,
        W_FLD_CHAR
    )

    end.set(
        W + "fldCharType",
        "end"
    )

    return paragraph


def create_footer_root(bmitra_family):
    root = etree.Element(
        W + "ftr",
        nsmap={"w": W_NS}
    )

    append_page_field_paragraph(
        root,
        bmitra_family
    )

    return root


def add_new_footer_part(
    files,
    rel_root,
    content_types_root,
    footer_type,
    bmitra_family
):
    existing_numbers = []

    for name in files:
        match = re.fullmatch(
            r"word/footer(\d+)\.xml",
            name
        )

        if match:
            existing_numbers.append(
                int(match.group(1))
            )

    next_number = max(
        existing_numbers,
        default=0
    ) + 1

    part_name = f"word/footer{next_number}.xml"
    target = f"footer{next_number}.xml"

    existing_ids = []

    for relationship in rel_root.findall(
        PKG + "Relationship"
    ):
        match = re.fullmatch(
            r"rId(\d+)",
            relationship.get("Id", "")
        )

        if match:
            existing_ids.append(
                int(match.group(1))
            )

    relationship_id = (
        "rId" + str(max(existing_ids, default=0) + 1)
    )

    relationship = etree.SubElement(
        rel_root,
        PKG + "Relationship"
    )

    relationship.set(
        "Id",
        relationship_id
    )

    relationship.set(
        "Type",
        R_NS + "/footer"
    )

    relationship.set(
        "Target",
        target
    )

    override_exists = any(
        element.get("PartName") == "/" + part_name
        for element in content_types_root.findall(
            CT + "Override"
        )
    )

    if not override_exists:
        override = etree.SubElement(
            content_types_root,
            CT + "Override"
        )

        override.set(
            "PartName",
            "/" + part_name
        )

        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
        )

    footer_root = create_footer_root(
        bmitra_family
    )

    files[part_name] = serialize_xml(footer_root)

    return part_name, relationship_id


def relationship_map(rel_root):
    result = {}

    for rel in rel_root.findall(
        PKG + "Relationship"
    ):
        result[rel.get("Id")] = rel

    return result


def footer_part_for_reference(reference, relationships):
    relationship_id = reference.get(R + "id")

    relationship = relationships.get(
        relationship_id
    )

    if relationship is None:
        raise RuntimeError(
            "Footer relationship is missing: " +
            str(relationship_id)
        )

    if not relationship.get("Type", "").endswith("/footer"):
        raise RuntimeError(
            "Section footer reference points to a non-footer relationship."
        )

    if relationship.get("TargetMode") == "External":
        raise RuntimeError(
            "External footer relationships are not supported."
        )

    target = relationship.get("Target", "")

    if target.startswith("/"):
        part_name = target.lstrip("/")
    else:
        part_name = posixpath.normpath(
            posixpath.join("word", target)
        )

    if part_name not in files_placeholder:
        pass

    return part_name


def add_footer_reference(sect_pr, footer_type, relationship_id):
    reference = etree.Element(
        W_FOOTER_REF
    )

    reference.set(
        W + "type",
        footer_type
    )

    reference.set(
        R + "id",
        relationship_id
    )

    # Header/footer references precede section layout properties.
    insert_at = 0

    for index, child in enumerate(sect_pr):
        if child.tag in (
            W + "headerReference",
            W_FOOTER_REF,
        ):
            insert_at = index + 1

    sect_pr.insert(
        insert_at,
        reference
    )


def ensure_decimal_section_numbering(sect_pr):
    page_type = sect_pr.find(W_PGNUMTYPE)

    if page_type is None:
        page_type = etree.Element(
            W_PGNUMTYPE
        )

        successors = {
            W + name
            for name in (
                "cols",
                "formProt",
                "vAlign",
                "noEndnote",
                "titlePg",
                "textDirection",
                "bidi",
                "rtlGutter",
                "docGrid",
                "printerSettings",
                "sectPrChange",
            )
        }

        insert_at = next(
            (
                index
                for index, child in enumerate(sect_pr)
                if child.tag in successors
            ),
            len(sect_pr)
        )

        sect_pr.insert(
            insert_at,
            page_type
        )

    page_type.set(
        W + "fmt",
        "decimal"
    )

    # Remove section-specific restarts so numbering continues.
    page_type.attrib.pop(
        W + "start",
        None
    )


def boolean_setting_exists(settings_data, element_name):
    if not settings_data:
        return False

    root = parse_xml(
        settings_data,
        "word/settings.xml"
    )

    element = root.find(W + element_name)

    if element is None:
        return False

    value = element.get(W + "val")

    return value is None or str(value).casefold() not in (
        "0",
        "false",
        "off",
        "no",
    )


def append_number_to_existing_footer(
    data,
    part_name,
    bmitra_family
):
    root = parse_xml(
        data,
        part_name
    )

    if root.tag != W + "ftr":
        raise RuntimeError(
            "Unexpected root element in " + part_name
        )

    removed = remove_page_fields(root)

    append_page_field_paragraph(
        root,
        bmitra_family
    )

    return serialize_xml(root, data), removed


def ensure_page_numbers(files, bmitra_family):
    if "word/document.xml" not in files:
        raise RuntimeError(
            "word/document.xml is missing."
        )

    if "word/_rels/document.xml.rels" not in files:
        raise RuntimeError(
            "word/document.xml.rels is missing."
        )

    if "[Content_Types].xml" not in files:
        raise RuntimeError(
            "[Content_Types].xml is missing."
        )

    document_data = files["word/document.xml"]

    document_root = parse_xml(
        document_data,
        "word/document.xml"
    )

    rel_root = parse_xml(
        files["word/_rels/document.xml.rels"],
        "word/_rels/document.xml.rels"
    )

    content_types_root = parse_xml(
        files["[Content_Types].xml"],
        "[Content_Types].xml"
    )

    # Remove old PAGE fields from headers and footers.
    removed_fields = 0

    header_footer_names = [
        name for name in list(files)
        if re.fullmatch(
            r"word/(?:header|footer)\d+\.xml",
            name
        )
    ]

    for part_name in header_footer_names:
        root = parse_xml(
            files[part_name],
            part_name
        )

        removed_fields += remove_page_fields(root)

        if part_name.startswith("word/footer"):
            append_page_field_paragraph(
                root,
                bmitra_family
            )

        files[part_name] = serialize_xml(
            root,
            files[part_name]
        )

    # Obtain all current section property elements, excluding
    # historical properties retained inside sectPrChange.
    sections = []

    for section in document_root.iter(W_SECTPR):
        parent = section.getparent()

        if parent is not None and parent.tag == W_SECTPR_CHANGE:
            continue

        sections.append(section)

    if not sections:
        body = document_root.find(".//" + W + "body")

        if body is None:
            raise RuntimeError(
                "Main document body is missing."
            )

        section = etree.SubElement(
            body,
            W_SECTPR
        )

        sections = [section]

    relationships = relationship_map(
        rel_root
    )

    even_and_odd = boolean_setting_exists(
        files.get("word/settings.xml"),
        "evenAndOddHeaders"
    )

    new_footer_parts = {}
    created_references = 0

    def obtain_new_footer(footer_type):
        if footer_type in new_footer_parts:
            return new_footer_parts[footer_type]

        part_name, relationship_id = add_new_footer_part(
            files,
            rel_root,
            content_types_root,
            footer_type,
            bmitra_family
        )

        new_footer_parts[footer_type] = (
            part_name,
            relationship_id
        )

        return part_name, relationship_id

    # Footer references inherit from the preceding section when
    # omitted. Track the effective reference for each type.
    inherited = {
        "default": None,
        "first": None,
        "even": None,
    }

    for section_index, section in enumerate(sections):
        direct_refs = section.findall(W_FOOTER_REF)

        for reference in direct_refs:
            footer_type = (
                reference.get(W + "type")
                or "default"
            )

            if footer_type not in inherited:
                continue

            rid = reference.get(R + "id")
            rel = relationships.get(rid)

            if rel is None:
                raise RuntimeError(
                    "Missing footer relationship: " + str(rid)
                )

            if not rel.get("Type", "").endswith("/footer"):
                raise RuntimeError(
                    "A footer reference points to a non-footer relationship."
                )

            target = rel.get("Target", "")

            if target.startswith("/"):
                part_name = target.lstrip("/")
            else:
                part_name = posixpath.normpath(
                    posixpath.join("word", target)
                )

            if part_name not in files:
                raise RuntimeError(
                    "Referenced footer part does not exist: " +
                    part_name
                )

            inherited[footer_type] = rid

        # Every document must have a default footer on its first
        # section, even when Google Docs exported no footer.
        if inherited["default"] is None:
            _, rid = obtain_new_footer("default")

            add_footer_reference(
                section,
                "default",
                rid
            )

            inherited["default"] = rid
            created_references += 1

        # A separate first-page footer must also receive a page number.
        has_first_page_footer = (
            section.find(W + "titlePg") is not None
        )

        if (
            has_first_page_footer
            and inherited["first"] is None
        ):
            _, rid = obtain_new_footer("first")

            add_footer_reference(
                section,
                "first",
                rid
            )

            inherited["first"] = rid
            created_references += 1

        # If even/odd footers are enabled, make sure the effective
        # even footer exists.
        if (
            even_and_odd
            and inherited["even"] is None
        ):
            _, rid = obtain_new_footer("even")

            add_footer_reference(
                section,
                "even",
                rid
            )

            inherited["even"] = rid
            created_references += 1

        ensure_decimal_section_numbering(
            section
        )

    files["word/document.xml"] = serialize_xml(
        document_root,
        document_data
    )

    files["word/_rels/document.xml.rels"] = serialize_xml(
        rel_root,
        files["word/_rels/document.xml.rels"]
    )

    files["[Content_Types].xml"] = serialize_xml(
        content_types_root,
        files["[Content_Types].xml"]
    )

    return {
        "oldPageFieldRunsRemoved": removed_fields,
        "newFooterPartsCreated": len(new_footer_parts),
        "footerReferencesCreated": created_references,
        "sectionsUpdated": len(sections),
        "evenAndOddFootersEnabled": even_and_odd,
    }


# ============================================================================
# Footnote formatting
# ============================================================================

def make_run_from_child(source_run, child, family, size_points=None):
    new_run = etree.Element(
        W_R
    )

    for name, value in source_run.attrib.items():
        new_run.set(name, value)

    source_rpr = source_run.find(W_RPR)

    if source_rpr is not None:
        rpr = copy.deepcopy(source_rpr)
    else:
        rpr = etree.Element(W_RPR)

    set_run_font(
        rpr,
        family
    )

    if size_points is not None:
        set_run_size(
            rpr,
            size_points
        )

    new_run.append(rpr)

    if child is not None:
        new_child = copy.deepcopy(child)
        new_run.append(new_child)

    return new_run


def split_footnote_run(
    run,
    paragraph,
    style_index,
    font_families
):
    parent = run.getparent()

    if parent is None:
        return 0

    is_bold = run_is_bold(
        run,
        paragraph,
        style_index
    )

    numeric_family = (
        font_families["bmitra_bold"]
        if is_bold
        else font_families["bmitra_regular"]
    )

    normal_family = font_families["irmitra"]

    children = [
        child for child in list(run)
        if child.tag != W_RPR
    ]

    if not children:
        rpr = ensure_run_properties(run)
        set_run_font(rpr, normal_family)
        set_run_size(rpr, 11)
        return 0

    generated = []

    def emit(child, family):
        fragment = make_run_from_child(
            run,
            child,
            family,
            11
        )

        generated.append(fragment)

    for child in children:
        if child.tag == W_T:
            text = child.text or ""

            if not text:
                emit(child, normal_family)
                continue

            position = 0

            for match in R_DIGITS.finditer(text):
                if match.start() > position:
                    non_digit = copy.deepcopy(child)
                    non_digit.text = text[position:match.start()]
                    non_digit.tail = None

                    if (
                        non_digit.text.startswith(" ")
                        or non_digit.text.endswith(" ")
                    ):
                        non_digit.set(
                            "{" + XML_NS + "}space",
                            "preserve"
                        )

                    emit(
                        non_digit,
                        normal_family
                    )

                digit_node = copy.deepcopy(child)
                digit_node.text = match.group(0)
                digit_node.tail = None

                if (
                    digit_node.text.startswith(" ")
                    or digit_node.text.endswith(" ")
                ):
                    digit_node.set(
                        "{" + XML_NS + "}space",
                        "preserve"
                    )

                emit(
                    digit_node,
                    numeric_family
                )

                position = match.end()

            if position < len(text):
                non_digit = copy.deepcopy(child)
                non_digit.text = text[position:]
                non_digit.tail = child.tail

                if (
                    non_digit.text.startswith(" ")
                    or non_digit.text.endswith(" ")
                ):
                    non_digit.set(
                        "{" + XML_NS + "}space",
                        "preserve"
                    )

                emit(
                    non_digit,
                    normal_family
                )

            elif generated:
                generated[-1][-1].tail = child.tail

        elif child.tag == W + "footnoteRef":
            emit(
                child,
                numeric_family
            )

        else:
            emit(
                child,
                normal_family
            )

    if not generated:
        rpr = ensure_run_properties(run)
        set_run_font(rpr, normal_family)
        set_run_size(rpr, 11)
        return 0

    position = parent.index(run)
    parent.remove(run)

    for offset, fragment in enumerate(generated):
        parent.insert(
            position + offset,
            fragment
        )

    return len(generated)


def apply_footnote_formatting(
    files,
    style_index,
    font_families
):
    footnotes_name = "word/footnotes.xml"
    document_name = "word/document.xml"

    result = {
        "footnotesFound": 0,
        "footnoteRunsSplit": 0,
        "mainReferencesFormatted": 0,
        "numericCharactersAssigned": 0,
    }

    if footnotes_name not in files:
        return result

    footnotes_data = files[footnotes_name]

    footnotes_root = parse_xml(
        footnotes_data,
        footnotes_name
    )

    for footnote in footnotes_root.findall(
        W_FOOTNOTE
    ):
        footnote_type = footnote.get(W + "type")
        footnote_id = footnote.get(W + "id")

        if footnote_type in (
            "separator",
            "continuationSeparator",
            "continuationNotice",
        ):
            continue

        try:
            if footnote_id is not None and int(footnote_id) < 0:
                continue
        except ValueError:
            pass

        result["footnotesFound"] += 1

        for paragraph in footnote.iter(W_P):
            runs = list(paragraph.iter(W_R))

            for run in runs:
                if run.getparent() is None:
                    continue

                result["footnoteRunsSplit"] += (
                    split_footnote_run(
                        run,
                        paragraph,
                        style_index,
                        font_families
                    )
                )

    files[footnotes_name] = serialize_xml(
        footnotes_root,
        footnotes_data
    )

    # Apply BMitra only to actual footnote reference marks
    # in the main document, not ordinary numbers in body text.
    if document_name in files:
        document_data = files[document_name]

        document_root = parse_xml(
            document_data,
            document_name
        )

        for paragraph in document_root.iter(W_P):
            for run in list(paragraph.iter(W_R)):
                refs = [
                    child
                    for child in list(run)
                    if child.tag == W_FOOTNOTE_REFERENCE
                ]

                if not refs:
                    continue

                bold = run_is_bold(
                    run,
                    paragraph,
                    style_index
                )

                family = (
                    font_families["bmitra_bold"]
                    if bold
                    else font_families["bmitra_regular"]
                )

                children = [
                    child for child in list(run)
                    if child.tag != W_RPR
                ]

                if len(children) == 1 and children[0].tag == W_FOOTNOTE_REFERENCE:
                    rpr = ensure_run_properties(run)
                    set_run_font(rpr, family)
                    result["mainReferencesFormatted"] += 1
                    continue

                parent = run.getparent()

                if parent is None:
                    continue

                generated = []

                for child in children:
                    target_family = (
                        family
                        if child.tag == W_FOOTNOTE_REFERENCE
                        else font_families["irmitra"]
                    )

                    generated.append(
                        make_run_from_child(
                            run,
                            child,
                            target_family,
                            None
                        )
                    )

                    if child.tag == W_FOOTNOTE_REFERENCE:
                        result["mainReferencesFormatted"] += 1

                position = parent.index(run)
                parent.remove(run)

                for offset, fragment in enumerate(generated):
                    parent.insert(
                        position + offset,
                        fragment
                    )

        files[document_name] = serialize_xml(
            document_root,
            document_data
        )

    return result


# ============================================================================
# DOCX package processing
# ============================================================================

def read_docx_package(path):
    if not zipfile.is_zipfile(path):
        raise RuntimeError(
            "input.docx is not a valid ZIP archive."
        )

    with zipfile.ZipFile(path, "r") as archive:
        infos = archive.infolist()

        files = {
            item.filename: archive.read(item.filename)
            for item in infos
        }

    return infos, files


def write_docx_package(infos, files, destination):
    with zipfile.ZipFile(
        destination,
        "w",
        compression=zipfile.ZIP_DEFLATED
    ) as archive:

        written = set()

        for info in infos:
            name = info.filename

            if name in files:
                archive.writestr(
                    info,
                    files[name]
                )
                written.add(name)

        for name, data in files.items():
            if name in written:
                continue

            archive.writestr(
                name,
                data,
                compress_type=zipfile.ZIP_DEFLATED
            )


def process_docx(font_families, options):
    infos, files = read_docx_package(SOURCE)

    if "word/document.xml" not in files:
        raise RuntimeError(
            "word/document.xml is missing."
        )

    style_index = build_style_index(
        files.get("word/styles.xml")
    )

    report = {
        "xmlPartsProcessed": 0,
        "wordRuns": 0,
        "wordRunProperties": 0,
        "wordSymbols": 0,
        "drawingElements": 0,
        "blackTextElements": 0,
        "backgroundShadingRemoved": 0,
        "paragraphsFormatted": 0,
        "normalRunsRaised": 0,
        "shiftReturnProtection": False,
        "paginationControls": None,
        "pageNumbers": None,
        "footnotes": None,
    }

    irmitra_family = font_families["irmitra"]

    # First pass:
    # enforce IRMitra, optional paragraph formatting,
    # optional black text and optional No Fill.
    for part_name in list(files):
        if not (
            part_name.startswith(WORD_XML_ROOT)
            and part_name.lower().endswith(".xml")
        ):
            continue

        original_data = files[part_name]

        root = parse_xml(
            original_data,
            part_name
        )

        if root.tag.startswith(W):
            run_count, property_count, symbol_count = (
                force_word_fonts(
                    root,
                    irmitra_family
                )
            )

            report["wordRuns"] += run_count
            report["wordRunProperties"] += property_count
            report["wordSymbols"] += symbol_count

        report["drawingElements"] += (
            force_drawing_fonts(
                root,
                irmitra_family
            )
        )

        if get_boolean(options, "blackText"):
            report["blackTextElements"] += (
                force_black_text(root)
            )

        if get_boolean(options, "removeBackground"):
            report["backgroundShadingRemoved"] += (
                remove_background_shading(root)
            )

        
        if part_name == "word/document.xml":
            # ابتدا حفظ و اعمال ویژگی‌های پیشرفته صفحه‌آرایی.
            # این مرحله مستقل از روشن یا خاموش بودن فاصله‌گذاری است.
            report["paginationControls"] = (
                apply_advanced_paragraph_controls(
                    root,
                    style_index
                )
            )

            # سپس فاصله‌گذاری و حداقل اندازه فونت.
            spacing_report = apply_paragraph_formatting(
                root,
                style_index,
                options
            )

            report["paragraphsFormatted"] = (
                spacing_report["paragraphs"]
            )

            report["normalRunsRaised"] = (
                spacing_report["normalRunsRaised"]
            )


        if part_name == "word/settings.xml":
            files[part_name] = ensure_do_not_expand_shift_return(
                serialize_xml(
                    root,
                    original_data
                )
            )

            report["shiftReturnProtection"] = True

        else:
            files[part_name] = serialize_xml(
                root,
                original_data
            )

        report["xmlPartsProcessed"] += 1

    # The DOCX export normally contains settings.xml.
    # Do not silently skip the Shift+Enter protection.
    if not report["shiftReturnProtection"]:
        if "word/settings.xml" not in files:
            raise RuntimeError(
                "word/settings.xml is missing."
            )

        files["word/settings.xml"] = (
            ensure_do_not_expand_shift_return(
                files["word/settings.xml"]
            )
        )

        report["shiftReturnProtection"] = True

    # Page-number replacement must run after the normal
    # IRMitra pass, because newly inserted page fields use BMitra.
    if get_boolean(options, "replacePageNumbers"):
        report["pageNumbers"] = ensure_page_numbers(
            files,
            font_families["bmitra_regular"]
        )

    # Footnote formatting runs last so BMitra digit formatting
    # is not overwritten by the normal IRMitra enforcement.
    if get_boolean(options, "formatFootnotes"):
        report["footnotes"] = apply_footnote_formatting(
            files,
            style_index,
            font_families
        )

    write_docx_package(
        infos,
        files,
        TEMP_DOCX
    )

    with zipfile.ZipFile(TEMP_DOCX, "r") as archive:
        bad_member = archive.testzip()

    if bad_member:
        raise RuntimeError(
            "Patched DOCX integrity check failed: " + bad_member
        )

    os.replace(
        TEMP_DOCX,
        SOURCE
    )

    return report


# ============================================================================
# Main
# ============================================================================

def main():
    if not SOURCE.is_file():
        raise RuntimeError(
            "input.docx does not exist."
        )

    options = load_options()
    font_families = install_and_check_fonts()

    report = process_docx(
        font_families,
        options
    )

    print("")
    print("========== IRMitra rendering report ==========")

    print(
        "Detected IRMitra family:",
        font_families["irmitra"]
    )

    print(
        "Detected BMitra Regular family:",
        font_families["bmitra_regular"]
    )

    print(
        "Detected BMitra Bold family:",
        font_families["bmitra_bold"]
    )

    print(
        "XML parts processed:",
        report["xmlPartsProcessed"]
    )

    print(
        "Word text runs:",
        report["wordRuns"]
    )

    print(
        "Word run-property definitions:",
        report["wordRunProperties"]
    )

    print(
        "Word symbol declarations:",
        report["wordSymbols"]
    )

    print(
        "DrawingML text elements:",
        report["drawingElements"]
    )

    print(
        "Shift+Enter protection:",
        "ENABLED" if report["shiftReturnProtection"] else "NOT SET"
    )
    if report["paginationControls"] is not None:
        print(
            "Advanced pagination controls:",
            json.dumps(
                report["paginationControls"],
                ensure_ascii=False
            )
        )
    print(
        "Paragraph formatting enabled:",
        get_boolean(options, "paragraphFormatting")
    )

    print(
        "Paragraphs formatted:",
        report["paragraphsFormatted"]
    )

    print(
        "Normal runs raised to minimum font size:",
        report["normalRunsRaised"]
    )

    print(
        "Black text enabled:",
        get_boolean(options, "blackText")
    )

    print(
        "Black text elements:",
        report["blackTextElements"]
    )

    print(
        "No Fill enabled:",
        get_boolean(options, "removeBackground")
    )

    print(
        "Background shading elements removed:",
        report["backgroundShadingRemoved"]
    )

    print(
        "Page number replacement enabled:",
        get_boolean(options, "replacePageNumbers")
    )

    if report["pageNumbers"] is not None:
        print(
            "Page numbering report:",
            json.dumps(
                report["pageNumbers"],
                ensure_ascii=False
            )
        )

    print(
        "Footnote formatting enabled:",
        get_boolean(options, "formatFootnotes")
    )

    if report["footnotes"] is not None:
        print(
            "Footnote report:",
            json.dumps(
                report["footnotes"],
                ensure_ascii=False
            )
        )

    print("DOCX customization completed.")
    print("==============================================")


if __name__ == "__main__":
    main()
