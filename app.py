import gc
import base64
import io
import logging
import threading
import requests
import os
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from PIL import Image as PILImage

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

STYLE_COL       = 2
IMAGE_COL       = "A"
ROW_HEIGHT_PT   = 60
COL_A_WIDTH     = 12
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 6
ROW_HEIGHT_PX   = int(ROW_HEIGHT_PT * 96 / 72)
COL_A_WIDTH_PX  = int(COL_A_WIDTH * 7)
PADDING_PX      = 4

def download_and_resize(style_str):
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
            img    = img.convert("RGB")
            ow, oh = img.size
            max_w  = COL_A_WIDTH_PX - (PADDING_PX * 2)
            max_h  = ROW_HEIGHT_PX  - (PADDING_PX * 2)
            scale  = min(max_w / ow, max_h / oh, 1.0)
            nw     = max(1, int(ow * scale))
            nh     = max(1, int(oh * scale))
            img    = img.resize((nw, nh), PILImage.LANCZOS)
            buf    = io.BytesIO()
            img.save(buf, format="PNG", optimize=True, compress_level=6)
            buf.seek(0)
            return buf.getvalue(), nw, nh
    except Exception as e:
        logging.warning(f"Image failed for {style_str}: {e}")
        return None, 0, 0
    finally:
        gc.collect()

def process_and_callback(excel_bytes, filename, callback_url):
    try:
        logging.info(f"Received file size: {len(excel_bytes)} bytes")

        wb = load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        del excel_bytes
        gc.collect()

        # ✅ GUARD: Skip if already processed (Image column already exists)
        # This prevents the output email from being re-processed by Flow 1
        if ws['A1'].value and str(ws['A1'].value).strip().lower() == "image":
            logging.warning("File already has Image column — skipping to prevent duplicate processing")
            return

        # ✅ Insert new Image column at position 1
        ws.insert_cols(1)
        ws['A1'] = "Image"

        # ✅ Bold all header cells in row 1
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # ✅ Freeze header row
        ws.freeze_panes = "A2"

        # ✅ Set Image column (A) to fixed width
        ws.column_dimensions[IMAGE_COL].width = COL_A_WIDTH

        processed = 0
        skipped   = 0

        for row_idx in range(2, ws.max_row + 1):
            style_val = ws.cell(row=row_idx, column=STYLE_COL).value
            if style_val is None or str(style_val).strip() == "":
                continue

            style_str       = str(style_val).strip()
            img_bytes, w, h = download_and_resize(style_str)

            if img_bytes is None:
                skipped += 1
                continue

            try:
                xl_img        = XLImage(io.BytesIO(img_bytes))
                xl_img.width  = w
                xl_img.height = h
                xl_img.anchor = f"A{row_idx}"
                ws.add_image(xl_img)
                ws.row_dimensions[row_idx].height = ROW_HEIGHT_PT
                processed += 1
            except Exception as e:
                logging.warning(f"Embed failed for {style_str}: {e}")
                skipped += 1
            finally:
                del img_bytes
                gc.collect()

        # ✅ Auto-width all columns based on content
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)

            # Column A is the Image column — keep fixed width
            if col_letter == IMAGE_COL:
                ws.column_dimensions[col_letter].width = COL_A_WIDTH
                continue

            # All other columns — auto fit based on max content length
            max_length = 0
            for cell in ws[col_letter]:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            # Add padding, cap at 50 to avoid extremely wide columns
            ws.column_dimensions[col_letter].width = min(max_length + 4, 50)

        # ✅ Apply auto filter across full data range
        last_col           = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"

        # ✅ Save and validate output
        logging.info("Saving workbook...")
        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)
        output_bytes = out_buf.read()

        if output_bytes.startswith(b'PK\x03\x04'):
            logging.info(f"Output xlsx VALID ✅ size: {len(output_bytes)} bytes")
        else:
            logging.error(f"Output xlsx INVALID ❌ starts with: {output_bytes[:4].hex()}")

        result_b64 = base64.b64encode(output_bytes).decode("utf-8")
        logging.info(f"Base64 length: {len(result_b64)}")
        del wb, output_bytes
        gc.collect()

        # ✅ Ensure .xlsx extension
        if not filename.lower().endswith(".xlsx"):
            filename = filename + ".xlsx"

        # ✅ Callback to Flow 2
        payload = {
            "file"      : result_b64,
            "filename"  : filename,
            "processed" : processed,
            "skipped"   : skipped
        }
        resp = requests.post(callback_url, json=payload, timeout=30)
        logging.info(f"Callback status: {resp.status_code}, processed={processed}, skipped={skipped}")

    except Exception as e:
        logging.error(f"Background processing failed: {e}", exc_info=True)

@app.route("/embed_images", methods=["POST"])
def embed_images():
    try:
        body         = request.get_json(force=True)
        excel_b64    = body.get("file")
        filename     = body.get("filename", "report.xlsx")
        callback_url = os.environ.get("CALLBACK_URL")
    except Exception as e:
        return jsonify({"error": f"Invalid request: {e}"}), 400

    if not excel_b64:
        return jsonify({"error": "Missing 'file' field"}), 400
    if not callback_url:
        return jsonify({"error": "CALLBACK_URL not configured"}), 500

    try:
        # ✅ First decode
        excel_bytes = base64.b64decode(excel_b64)
        logging.info(f"Decoded excel_bytes size: {len(excel_bytes)} bytes")

        # ✅ Loop until valid xlsx found
        for attempt in range(4):
            if excel_bytes.startswith(b'PK\x03\x04'):
                logging.info(f"Valid xlsx after {attempt + 1} decode(s), size: {len(excel_bytes)} bytes")
                break
            logging.info(f"Attempt {attempt + 1}: starts with {excel_bytes[:4]}, decoding again...")
            excel_bytes = base64.b64decode(excel_bytes)
        else:
            return jsonify({"error": "Could not decode xlsx after 4 attempts"}), 422

    except Exception as e:
        return jsonify({"error": f"Base64 decode failed: {e}"}), 422

    thread = threading.Thread(
        target=process_and_callback,
        args=(excel_bytes, filename, callback_url),
        daemon=True
    )
    thread.start()

    return jsonify({"status": "accepted", "message": "Processing started"}), 202

@app.route("/", methods=["GET"])
def health_check():
    return jsonify({
        "status"  : "Townley Image Embedder is Live",
        "endpoint": "/embed_images (POST only)"
    }), 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
