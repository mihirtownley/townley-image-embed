import gc
import base64
import io
import logging
import requests
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

app = Flask(__name__)

STYLE_COL       = 1
IMAGE_COL       = "B"
ROW_HEIGHT_PT   = 60
COL_B_WIDTH     = 12
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 6
ROW_HEIGHT_PX   = int(ROW_HEIGHT_PT * 96 / 72)
COL_B_WIDTH_PX  = int(COL_B_WIDTH * 7)
PADDING_PX      = 4

def download_and_resize(style_str):
    """Download image and return resized PNG bytes, or None on failure."""
    try:
        resp = requests.get(
            f"{IMAGE_URL_BASE}{style_str}",
            timeout=REQUEST_TIMEOUT,
            stream=True
        )
        resp.raise_for_status()

        raw = io.BytesIO(resp.content)
        del resp

        with PILImage.open(raw) as img:
            img = img.convert("RGB")
            orig_w, orig_h = img.size
            max_w  = COL_B_WIDTH_PX - (PADDING_PX * 2)
            max_h  = ROW_HEIGHT_PX  - (PADDING_PX * 2)
            scale  = min(max_w / orig_w, max_h / orig_h, 1.0)
            new_w  = max(1, int(orig_w * scale))
            new_h  = max(1, int(orig_h * scale))
            img    = img.resize((new_w, new_h), PILImage.LANCZOS)
            buf    = io.BytesIO()
            img.save(buf, format="PNG", optimize=True, compress_level=6)
            buf.seek(0)
            return buf.getvalue(), new_w, new_h

    except Exception as e:
        logging.warning(f"Image failed for {style_str}: {e}")
        return None, 0, 0
    finally:
        gc.collect()


@app.route("/embed_images", methods=["POST"])
def embed_images():
    try:
        body      = request.get_json(force=True)
        excel_b64 = body.get("file")
        filename  = body.get("filename", "report.xlsx")
    except Exception as e:
        return jsonify({"error": f"Invalid request: {e}"}), 400

    if not excel_b64:
        return jsonify({"error": "Missing 'file' field"}), 400

    try:
        excel_bytes = base64.b64decode(excel_b64)
        wb          = load_workbook(io.BytesIO(excel_bytes))
        ws          = wb.active
        del excel_bytes
        gc.collect()
    except Exception as e:
        return jsonify({"error": f"Failed to load Excel: {e}"}), 422

    ws.column_dimensions[IMAGE_COL].width = COL_B_WIDTH

    processed = 0
    skipped   = 0

    for row_idx in range(2, ws.max_row + 1):
        style_val = ws.cell(row=row_idx, column=STYLE_COL).value
        if style_val is None or str(style_val).strip() == "":
            continue

        style_str          = str(style_val).strip()
        img_bytes, w, h    = download_and_resize(style_str)

        if img_bytes is None:
            skipped += 1
            continue

        try:
            xl_img        = XLImage(io.BytesIO(img_bytes))
            xl_img.width  = w
            xl_img.height = h
            xl_img.anchor = f"B{row_idx}"
            ws.add_image(xl_img)
            ws.row_dimensions[row_idx].height = ROW_HEIGHT_PT
            processed += 1
        except Exception as e:
            logging.warning(f"Embed failed for {style_str}: {e}")
            skipped += 1
        finally:
            del img_bytes
            gc.collect()

    try:
        out_buf    = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)
        result_b64 = base64.b64encode(out_buf.read()).decode("utf-8")
        del wb
        gc.collect()
    except Exception as e:
        return jsonify({"error": f"Failed to save workbook: {e}"}), 500

    return jsonify({
        "file"      : result_b64,
        "filename"  : filename,
        "processed" : processed,
        "skipped"   : skipped
    }), 200


@app.route("/", methods=["GET"])
def health_check():
    return jsonify({
        "status"  : "Townley Image Embedder is Live",
        "endpoint": "/embed_images (POST only)"
    }), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)