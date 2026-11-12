
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

FONT_FILES = {
    "regular": "IRMitra.ttf",
    "bold": "IRMitra-Bold.ttf",
    "italic": "IRMitra-Italic.ttf",
}


def required_env(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def decode_payload(value, label, maximum):
    if not isinstance(value, str) or not value:
        raise RuntimeError(f"Missing {label} payload.")

    try:
        data = base64.b64decode(value, validate=True)
    except Exception as exc:
        raise RuntimeError(f"Invalid Base64 data for {label}.") from exc

    if not data:
        raise RuntimeError(f"{label} is empty.")

    if len(data) > maximum:
        raise RuntimeError(f"{label} exceeds the permitted size.")

    return data


def main():
    bridge_url = required_env("BRIDGE_URL").rstrip("/")
    bridge_secret = required_env("BRIDGE_SECRET")
    job_id = required_env("JOB_ID")

    if not re.fullmatch(r"[A-Za-z0-9-]{30,40}", job_id):
        raise RuntimeError("Invalid job ID format.")

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
            str(payload.get("error", "Bridge download failed."))
        )

    docx_data = decode_payload(
        payload.get("docx"),
        "DOCX",
        MAX_DOCX_BYTES,
    )

    # Validate that the downloaded file is a real DOCX archive.
    try:
        with zipfile.ZipFile(io.BytesIO(docx_data)) as archive:
            if "word/document.xml" not in archive.namelist():
                raise RuntimeError("DOCX document.xml is missing.")

            bad_member = archive.testzip()
            if bad_member:
                raise RuntimeError("DOCX archive integrity check failed.")
    except zipfile.BadZipFile as exc:
        raise RuntimeError("Downloaded DOCX is not a valid ZIP archive.") from exc

    Path("input.docx").write_bytes(docx_data)

    fonts = payload.get("fonts")
    if not isinstance(fonts, dict):
        raise RuntimeError("Font payload is missing.")

    for key, filename in FONT_FILES.items():
        font_data = decode_payload(
            fonts.get(key),
            f"Font {key}",
            MAX_FONT_BYTES,
        )

        Path(filename).write_bytes(font_data)

    print("DOCX and all three font files were received and validated.")
    print("No document or font contents were written to the workflow log.")


if __name__ == "__main__":
    main()
