
import base64
import os
import re
import sys
from pathlib import Path

import requests


MAX_PDF_BYTES = 20 * 1024 * 1024


def required_env(name):
    value = os.environ.get(name, "").strip()

    if not value:
        raise RuntimeError(
            f"Missing environment variable: {name}"
        )

    return value


def main():
    if len(sys.argv) != 2:
        raise RuntimeError(
            "Usage: python3 scripts/callback.py complete|fail"
        )

    action = sys.argv[1].strip().lower()

    if action not in ("complete", "fail"):
        raise RuntimeError("Unsupported callback action.")

    bridge_url = required_env("BRIDGE_URL").rstrip("/")
    bridge_secret = required_env("BRIDGE_SECRET")
    job_id = required_env("JOB_ID")

    if not re.fullmatch(r"[A-Za-z0-9-]{30,40}", job_id):
        raise RuntimeError("Invalid job ID format.")

    payload = {
        "action": action,
        "jobId": job_id,
        "secret": bridge_secret,
    }

    if action == "complete":
        pdf_path = Path("out/input.pdf")

        if not pdf_path.is_file():
            raise RuntimeError("PDF output file does not exist.")

        size = pdf_path.stat().st_size

        if size <= 0 or size > MAX_PDF_BYTES:
            raise RuntimeError(
                "PDF is empty or exceeds the 20 MB limit."
            )

        pdf_data = pdf_path.read_bytes()

        if not pdf_data.startswith(b"%PDF-"):
            raise RuntimeError("The output is not a valid PDF file.")

        payload["pdfBase64"] = base64.b64encode(
            pdf_data
        ).decode("ascii")

    else:
        payload["error"] = (
            "GitHub Actions rendering failed. "
            "Inspect the workflow run for the technical error."
        )

    response = requests.post(
        bridge_url,
        json=payload,
        timeout=(20, 300),
        allow_redirects=True,
    )

    response.raise_for_status()

    try:
        result = response.json()
    except ValueError as exc:
        raise RuntimeError(
            "Apps Script returned a non-JSON callback response."
        ) from exc

    if not result.get("ok"):
        raise RuntimeError(
            str(result.get("error", "Apps Script callback failed."))
        )

    if action == "complete":
        print("PDF was successfully saved to Google Drive.")
    else:
        print("Failure notification was sent to Apps Script.")


if __name__ == "__main__":
    main()
