from flask import Flask, request, jsonify
from concurrent.futures import ThreadPoolExecutor, as_completed
import base64
import io
import logging
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

app = Flask(__name__)

STYLE_COL       = 1
IMAGE_COL       = "B"
ROW_HEIGHT_PT   = 72
COL_B_WIDTH     = 14
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 8
ROW_HEIGHT_PX   = int(ROW_HEIGHT_PT * 96 / 72)
COL_B_WIDTH_PX  = int(COL_B_WIDTH * 7)
PADDING_PX      = 4


def fetch_image(args):
    row_idx, style_str = args
    img_url = f"{IMAGE_URL_BASE}{style_str}"
    try:
        resp = requests.get(img_url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return row_idx, style_str, resp.content, None
    except Exception as e:
        return row_idx, style_str, None, str(e)


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
    except Exception as e:
        return jsonify({"error": f"Failed to load Excel: {e}"}), 422

    ws.column_dimensions[IMAGE_COL].width = COL_B_WIDTH

    rows_to_process = []
    for row_idx in range(2, ws.max_row + 1):
        style_val = ws.cell(row=row_idx, column=STYLE_COL).value
        if style_val is None or str(style_val).strip() == "":
            continue
        rows_to_process.append((row_idx, str(style_val).strip()))

    processed = 0
    skipped   = 0
    results   = {}

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(fetch_image, row): row
            for row in rows_to_process
        }
        for future in as_completed(futures):
            row_idx, style_str, content, error = future.result()
            results[row_idx] = (style_str, content, error)

    for row_idx, (style_str, content, error) in sorted(results.items()):
        if error or content is None:
            logging.warning(f"Skipped {style_str}: {error}")
            skipped += 1
            continue

        try:
            pil_img        = PILImage.open(io.BytesIO(content)).convert("RGBA")
            orig_w, orig_h = pil_img.size
            max_w          = COL_B_WIDTH_PX - (PADDING_PX * 2)
            max_h          = ROW_HEIGHT_PX  - (PADDING_PX * 2)
            scale          = min(max_w / orig_w, max_h / orig_h, 1.0)
            new_w          = max(1, int(orig_w * scale))
            new_h          = max(1, int(orig_h * scale))
            pil_img        = pil_img.resize((new_w, new_h), PILImage.LANCZOS)
            png_buf        = io.BytesIO()
            pil_img.save(png_buf, format="PNG", optimize=True)
            png_buf.seek(0)
        except Exception as e:
            logging.warning(f"Image processing failed for {style_str}: {e}")
            skipped += 1
            continue

        try:
            xl_img        = XLImage(png_buf)
            xl_img.width  = new_w
            xl_img.height = new_h
            xl_img.anchor = f"B{row_idx}"
            ws.add_image(xl_img)
        except Exception as e:
            logging.warning(f"Failed to embed image for {style_str}: {e}")
            skipped += 1
            continue

        ws.row_dimensions[row_idx].height = ROW_HEIGHT_PT
        processed += 1

    try:
        out_buf    = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)
        result_b64 = base64.b64encode(out_buf.read()).decode("utf-8")
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