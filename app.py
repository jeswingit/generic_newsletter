"""
app.py — Aon Newsletter Builder Flask Backend
"""

import base64
import copy
import io
import json
import os
import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from generate_newsletter import IMAGE_MIME, _load_image_part, read_excel_rows
from newsletter_renderer import SECTION_LABELS, SECTION_RENDERERS, render_newsletter
from template_generator import create_excel_template

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit

UPLOAD_DIR = Path(__file__).parent / "static" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    _icons = {
        "header":       "⬛",
        "footer":       "⬜",
        "bullet_list":  "•≡",
        "event_list":   "📅",
        "text_block":   "¶",
        "product_card": "🃏",
        "image_block":  "🖼",
        "divider":      "—",
    }
    section_types = [
        {"type": k, "label": v, "icon": _icons.get(k, "+")} for k, v in SECTION_LABELS.items()
    ]
    return render_template("index.html", section_types=section_types)


@app.route("/api/upload-image", methods=["POST"])
def upload_image():
    if "image" not in request.files:
        return jsonify({"error": "No image file provided"}), 400
    file = request.files["image"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    ext = Path(file.filename).suffix.lower()
    if ext not in IMAGE_MIME:
        return jsonify({"error": f"Unsupported image type: {ext}"}), 400

    filename = f"{uuid.uuid4()}{ext}"
    save_path = UPLOAD_DIR / filename
    file.save(str(save_path))

    return jsonify({"url": f"/static/uploads/{filename}"})


@app.route("/api/preview", methods=["POST"])
def preview():
    config = request.get_json(silent=True)
    if not config:
        return jsonify({"error": "Invalid JSON body"}), 400

    try:
        html = render_newsletter(config)
        return jsonify({"html": html})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    from generate_newsletter import build_eml_message

    config = request.get_json(silent=True)
    if not config:
        return jsonify({"error": "Invalid JSON body"}), 400

    meta = config.get("meta", {})
    fmt = config.get("format", "eml")

    from_addr = meta.get("from", "newsletter@aon.com")
    to_addr = meta.get("to", "")
    subject = meta.get("subject", "Aon Newsletter")

    try:
        if fmt == "html":
            # Embed uploaded images as base64 so the HTML file is self-contained
            html = render_newsletter(_embed_images_as_base64(config))
            buf = io.BytesIO(html.encode("utf-8"))
            safe_name = _safe_filename(meta.get("newsletterName", "newsletter")) + ".html"
            return send_file(
                buf,
                as_attachment=True,
                download_name=safe_name,
                mimetype="text/html",
            )

        # EML: collect images (imageUrl + logoUrl) and embed as CID
        image_cids: dict[str, str] = {}
        image_parts: dict[str, object] = {}

        for section in config.get("sections", []):
            props = section.get("props", {})
            for url_key in ("imageUrl", "logoUrl"):
                url = props.get(url_key, "")
                if url and url not in image_cids:
                    server_path = _url_to_server_path(url)
                    if server_path:
                        part, cid = _load_image_part(str(server_path), UPLOAD_DIR)
                        if part is not None:
                            image_cids[url] = cid
                            image_parts[url] = part

        html = render_newsletter(config, image_cids=image_cids)
        msg = build_eml_message(html, from_addr, to_addr, subject)
        for img_part in image_parts.values():
            msg.attach(img_part)

        buf = io.BytesIO(msg.as_bytes())
        safe_name = _safe_filename(meta.get("newsletterName", "newsletter")) + ".eml"
        return send_file(
            buf,
            as_attachment=True,
            download_name=safe_name,
            mimetype="message/rfc822",
        )

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/template")
def template_download():
    buf = create_excel_template()
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="newsletter_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/import-excel", methods=["POST"])
def import_excel():
    if "xlsx_file" not in request.files:
        return jsonify({"error": "No xlsx_file provided"}), 400
    file = request.files["xlsx_file"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp_path = tmp.name
            file.save(tmp_path)

        grouped = read_excel_rows(Path(tmp_path))
        sections = _excel_to_sections(grouped)
        return jsonify({"sections": sections})

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


# ---------------------------------------------------------------------------
# Error handler
# ---------------------------------------------------------------------------

@app.errorhandler(Exception)
def handle_error(exc):
    return jsonify({"error": str(exc)}), 500


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_filename(name: str) -> str:
    """Strip characters that are unsafe in filenames."""
    return "".join(c for c in name if c.isalnum() or c in " _-").strip() or "newsletter"


def _embed_images_as_base64(config: dict) -> dict:
    """Return a deep copy of config with /static/uploads/ image URLs replaced by base64 data URIs."""
    _mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
             ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp"}
    config = copy.deepcopy(config)
    for section in config.get("sections", []):
        props = section.get("props", {})
        for key in ("imageUrl", "logoUrl"):
            url = props.get(key, "")
            if not url:
                continue
            path = _url_to_server_path(url)
            if path and path.exists():
                mime = _mime.get(path.suffix.lower(), "image/png")
                b64 = base64.b64encode(path.read_bytes()).decode()
                props[key] = f"data:{mime};base64,{b64}"
    return config


def _url_to_server_path(url: str) -> Path | None:
    """Convert a /static/uploads/xxx.ext URL to an absolute server path."""
    prefix = "/static/uploads/"
    if url.startswith(prefix):
        filename = url[len(prefix):]
        candidate = UPLOAD_DIR / filename
        if candidate.exists():
            return candidate
    return None


def _excel_to_sections(grouped: dict) -> list[dict]:
    """Convert read_excel_rows() output to a list of newsletter section dicts."""
    sections = []

    # Header
    sections.append({
        "id": str(uuid.uuid4()),
        "type": "header",
        "props": {"orgName": "Aon", "tagline": "Newsletter", "backgroundColor": "#E31837", "textColor": "#ffffff"},
    })

    # Month News → bullet_list
    month_news = grouped.get("Month News", [])
    if month_news:
        sections.append({
            "id": str(uuid.uuid4()),
            "type": "bullet_list",
            "props": {
                "heading": "What's Going On",
                "items": [r.get("data", "") for r in month_news if r.get("data")],
                "bulletColor": "#E31837",
                "backgroundColor": "#ffffff",
            },
        })

    # Save the Date → event_list
    save_date = grouped.get("Save the Date", [])
    if save_date:
        sections.append({
            "id": str(uuid.uuid4()),
            "type": "event_list",
            "props": {
                "heading": "Save the Date!",
                "items": [r.get("data", "") for r in save_date if r.get("data")],
                "backgroundColor": "#EEF6F7",
            },
        })

    # Product → product_card (one per row)
    for row in grouped.get("Product", []):
        sections.append({
            "id": str(uuid.uuid4()),
            "type": "product_card",
            "props": {
                "title": row.get("title", ""),
                "body": row.get("data", ""),
                "creator": row.get("creator", ""),
                "imageUrl": "",
                "backgroundColor": "#ffffff",
            },
        })

    # General → text_block (one per row)
    for row in grouped.get("General", []):
        sections.append({
            "id": str(uuid.uuid4()),
            "type": "text_block",
            "props": {
                "heading": row.get("title", ""),
                "body": row.get("data", ""),
                "backgroundColor": "#ffffff",
            },
        })

    # Footer
    sections.append({
        "id": str(uuid.uuid4()),
        "type": "footer",
        "props": {
            "orgName": "Aon",
            "year": str(__import__("datetime").datetime.now().year),
            "backgroundColor": "#1A1A1A",
            "textColor": "#aaaaaa",
        },
    })

    return sections


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    url = f"http://127.0.0.1:{port}/"
    # Open once: with debug reloader, only the child process has WERKZEUG_RUN_MAIN=true
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true" or not app.debug:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    app.run(debug=True, port=port)
