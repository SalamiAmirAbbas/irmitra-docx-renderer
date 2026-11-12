
import base64
import io
import json
import os
import re
import zipfile
from pathlib import Path

import requests


MAX_DOCX_BYTES = 10 * 1024 * 1024
MAX_FONT_BYTES = 10 * 1024 * 1024

# Canonical names, destination filenames and accepted bridge keys.
# Old aliases are retained for compatibility with the previous bridge.
FONT_FILES = {
    "irmitraRegular": (
        "IRMitra.ttf",
        ("irmitraRegular", "regular"),
    ),
    "irmitraBold": (
        "IRMitra-Bold.ttf",
        ("irmitraBold", "bold"),
    ),
    "irmitraItalic": (
        "IRMitra-Italic.ttf",
        ("irmitraItalic", "italic"),
    ),
    "bmitraRegular": (
        "BMitra.ttf",
        ("bmitraRegular", "bmitra_regular"),
    ),
    "bmitraBold": (
        "BMitraBd.ttf",
        ("bmitraBold", "bmitra_bold"),
    ),
}


def required_env(name):
    value = os.environ.get(name, "").strip()

    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {name}"
        )

    return value


def decode_payload(value, label, maximum):
    if not isinstance(value, str) or not value:
        raise RuntimeError(
            f"Missing {label} payload."
        )

    try:
        data = base64.b64decode(
            value,
            validate=True
        )
    except Exception as exc:
        raise RuntimeError(
            f"Invalid Base64 data for {label}."
        ) from exc

    if not data:
        raise RuntimeError(
            f"{label} is empty."
        )

    if len(data) > maximum:
        raise RuntimeError(
            f"{label} exceeds the permitted size."
        )

    return data


def validate_docx(data):
    try:
        with zipfile.ZipFile(
            io.BytesIO(data)
        ) as archive:

            names = archive.namelist()

            if "[Content_Types].xml" not in names:
                raise RuntimeError(
                    "DOCX content types are missing."
                )

            if "word/document.xml" not in names:
                raise RuntimeError(
                    "DOCX main document part is missing."
                )

            bad_member = archive.testzip()

            if bad_member:
                raise RuntimeError(
                    f"DOCX archive integrity failed: {bad_member}"
                )

    except zipfile.BadZipFile as exc:
        raise RuntimeError(
            "Downloaded file is not a valid DOCX archive."
        ) from exc


def resolve_font_payload(
    fonts,
    canonical_name,
    aliases
):
    for key in aliases:
        value = fonts.get(key)

        if isinstance(value, str) and value:
            return value, key

    available = ", ".join(
        sorted(str(key) for key in fonts.keys())
    )

    raise RuntimeError(
        f"Missing font payload for {canonical_name}. "
        f"Expected one of: {', '.join(aliases)}. "
        f"Keys received from Apps Script: [{available}]"
    )


def validate_ttf(data, filename):
    # Valid TrueType/OpenType signatures.
    valid_signatures = (
        b"\x00\x01\x00\x00",
        b"OTTO",
        b"true",
        b"typ1",
    )

    if len(data) < 12 or data[:4] not in valid_signatures:
        raise RuntimeError(
            f"{filename} does not appear to be a valid TTF/OTF font."
        )


def main():
    bridge_url = required_env("BRIDGE_URL").rstrip("/")
    bridge_secret = required_env("BRIDGE_SECRET")
    job_id = required_env("JOB_ID")

    if not re.fullmatch(
        r"[A-Za-z0-9-]{30,40}",
        job_id
    ):
        raise RuntimeError(
            "Invalid job ID format."
        )

    response = requests.post(
        bridge_url,
        json={
            "action": "download",
            "jobId": job_id,
            "secret": bridge_secret,
        },
        timeout=(20, 240),
        allow_redirects=True,
    )

    response.raise_for_status()

    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError(
            "Apps Script returned a non-JSON response."
        ) from exc

    if not payload.get("ok"):
        raise RuntimeError(
            str(
                payload.get(
                    "error",
                    "Bridge download failed."
                )
            )
        )

    # Download and validate DOCX.
    docx_data = decode_payload(
        payload.get("docx"),
        "DOCX",
        MAX_DOCX_BYTES,
    )

    validate_docx(docx_data)

    Path("input.docx").write_bytes(
        docx_data
    )

    # Download and validate all five fonts.
    fonts = payload.get("fonts")

    if not isinstance(fonts, dict):
        raise RuntimeError(
            "Font payload is missing or is not a JSON object."
        )

    for canonical_name, (
        filename,
        aliases
    ) in FONT_FILES.items():

        encoded, actual_key = resolve_font_payload(
            fonts,
            canonical_name,
            aliases,
        )

        font_data = decode_payload(
            encoded,
            f"Font {canonical_name}",
            MAX_FONT_BYTES,
        )

        validate_ttf(
            font_data,
            filename
        )

        Path(filename).write_bytes(
            font_data
        )

        print(
            f"Font received and validated: {filename} "
            f"(bridge key: {actual_key}, "
            f"{len(font_data)} bytes)"
        )

    # Read the options associated with this specific job.
    options = payload.get("options")

    if not isinstance(options, dict):
        raise RuntimeError(
            "Customization options are missing or invalid. "
            "The deployed Apps Script may be an older version. "
            "Deploy the latest Code.gs as a new Web App version."
        )

    Path("job-options.json").write_text(
        json.dumps(
            options,
            ensure_ascii=False,
            separators=(",", ":"),
        ),
        encoding="utf-8",
    )

    print("DOCX received and validated.")
    print("All five font files received and validated.")
    print("Job customization options received.")
    print(
        "Private document and font contents "
        "were not printed."
    )


if __name__ == "__main__":
    main()
