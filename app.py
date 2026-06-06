import gc
import base64
import io
import logging
import threading
import requests
import os
from datetime import datetime, date
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import xlsxwriter

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

STYLE_COL_IDX   = 0        # Style is column A (index 0) in original file
IMAGE_COL_WIDTH = 18       # Excel character units
ROW_HEIGHT_PT   = 72       # Points (1 inch)
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 6
IMG_MAX_PX      = 300      # Max image dimension for storage quality

def download_and_resize(style_str):
    try:
        from PIL import Image as PILImage
        resp = requests.get(
            f"{IMAGE_URL_BASE}{style_str}",
            timeout=REQUEST_TIMEOUT,
            stream=True
        )
        resp.raise_for_status()
        raw = io.BytesIO(resp.content)
        del resp
        with PILImage.open(raw) as img:
            img  = img.convert("RGB")
            ow, oh = img.size
            # Resize to max IMG_MAX_PX while keeping aspect ratio
            scale = min(IMG_MAX_PX / ow, IMG_MAX_PX / oh, 1.0)
            nw    = max(1, int(ow * scale))
            nh    = max(1, int(oh * scale))
            img   = img.resize((nw, nh), PILImage.LANCZOS)
            buf   = io.BytesIO()
            img.save(buf, format="PNG", optimize=True, compress_level=6)
            buf.seek(0)
            return buf.getvalue()
    except Exception as e:
        logging.warning(f"Image failed for {style_str}: {e}")
        return None
    finally:
        gc.collect()

def process_and_callback(excel_bytes, filename, callback_url):
    try:
        logging.info(f"Received file size: {len(excel_bytes)} bytes")

        # ── STEP 1: Read source data with openpyxl ──────────────────────────
        wb = load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        del excel_bytes
        gc.collect()

        # Guard: skip already-processed files
        if ws['A1'].value and str(ws['A1'].value).strip().lower() == "image":
            logging.warning("File already processed — skipping")
            return

        # Read headers from row 1
        headers = [cell.value for cell in ws[1]]
        num_cols = len(headers)

        # Read all data rows
        all_rows = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if any(v is not None for v in row):
                all_rows.append(list(row))

        del wb
        gc.collect()

        # Sort A→Z by Style column
        all_rows.sort(key=lambda r: (
            r[STYLE_COL_IDX] is None,
            str(r[STYLE_COL_IDX]).lower() if r[STYLE_COL_IDX] else ""
        ))
        logging.info(f"Read {len(all_rows)} rows, sorted by Style A→Z ✅")

        # ── STEP 2: Download all images ──────────────────────────────────────
        logging.info("Downloading images...")
        images = {}
        for row_data in all_rows:
            style_val = row_data[STYLE_COL_IDX]
            if style_val:
                style_str = str(style_val).strip()
                if style_str and style_str not in images:
                    images[style_str] = download_and_resize(style_str)

        processed = sum(1 for v in images.values() if v is not None)
        skipped   = sum(1 for v in images.values() if v is None)
        logging.info(f"Images: {processed} downloaded, {skipped} failed ✅")

        # ── STEP 3: Calculate column widths ──────────────────────────────────
        col_widths = []
        for col_idx in range(num_cols):
            max_len = len(str(headers[col_idx])) if headers[col_idx] else 0
            for row_data in all_rows:
                if col_idx < len(row_data) and row_data[col_idx] is not None:
                    max_len = max(max_len, len(str(row_data[col_idx])))
            col_widths.append(min(max_len + 4, 50))

        # ── STEP 4: Build output xlsx with XlsxWriter ────────────────────────
        logging.info("Building xlsx with XlsxWriter embed_image()...")
        out_buf  = io.BytesIO()
        workbook = xlsxwriter.Workbook(out_buf, {'in_memory': True})
        sheet    = workbook.add_worksheet()

        # Formats
        bold_fmt = workbook.add_format({
            'bold'      : True,
            'font_name' : 'Calibri',
            'font_size' : 11
        })
        cell_fmt = workbook.add_format({
            'font_name' : 'Calibri',
            'font_size' : 11
        })
        date_fmt = workbook.add_format({
            'font_name'  : 'Calibri',
            'font_size'  : 11,
            'num_format' : 'mm/dd/yyyy'
        })
        num_fmt = workbook.add_format({
            'font_name' : 'Calibri',
            'font_size' : 11
        })

        # ── STEP 5: Write header row ─────────────────────────────────────────
        # Column 0 = Image (new), columns 1..N = original headers
        sheet.write(0, 0, "Image", bold_fmt)
        for col_idx, header in enumerate(headers):
            sheet.write(0, col_idx + 1, header if header is not None else "", bold_fmt)

        # ── STEP 6: Set column widths ─────────────────────────────────────────
        sheet.set_column(0, 0, IMAGE_COL_WIDTH)   # Image column fixed width
        for col_idx, width in enumerate(col_widths):
            sheet.set_column(col_idx + 1, col_idx + 1, width)

        # ── STEP 7: Write data rows with embedded images ──────────────────────
        for row_idx, row_data in enumerate(all_rows, start=1):

            # Set row height
            sheet.set_row(row_idx, ROW_HEIGHT_PT)

            # ✅ Embed image as true in-cell image (sorts + filters with data)
            style_val = row_data[STYLE_COL_IDX]
            if style_val:
                style_str = str(style_val).strip()
                img_bytes = images.get(style_str)
                if img_bytes:
                    try:
                        sheet.embed_image(row_idx, 0, "image.png", {
                            'image_data': io.BytesIO(img_bytes)
                        })
                    except Exception as e:
                        logging.warning(f"embed_image failed for {style_str}: {e}")

            # Write data values (col offset +1 for Image column)
            for col_idx, value in enumerate(row_data):
                dest_col = col_idx + 1
                if value is None:
                    sheet.write_blank(row_idx, dest_col, None, cell_fmt)
                elif isinstance(value, bool):
                    sheet.write_boolean(row_idx, dest_col, value, cell_fmt)
                elif isinstance(value, (datetime, date)):
                    sheet.write_datetime(row_idx, dest_col, value, date_fmt)
                elif isinstance(value, (int, float)):
                    sheet.write_number(row_idx, dest_col, value, num_fmt)
                else:
                    sheet.write_string(row_idx, dest_col, str(value), cell_fmt)

        # ── STEP 8: Freeze header row + auto filter ───────────────────────────
        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(all_rows), num_cols)

        # ── STEP 9: Save ──────────────────────────────────────────────────────
        workbook.close()
        out_buf.seek(0)
        output_bytes = out_buf.read()

        if output_bytes.startswith(b'PK\x03\x04'):
            logging.info(f"Output xlsx VALID ✅ size: {len(output_bytes)} bytes")
        else:
            logging.error("Output xlsx INVALID ❌")

        result_b64 = base64.b64encode(output_bytes).decode("utf-8")
        logging.info(f"Base64 length: {len(result_b64)}")
        del output_bytes
        gc.collect()

        if not filename.lower().endswith(".xlsx"):
            filename = filename + ".xlsx"

        # ── STEP 10: Callback to Flow 2 ───────────────────────────────────────
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
        excel_bytes = base64.b64decode(excel_b64)
        logging.info(f"Decoded excel_bytes size: {len(excel_bytes)} bytes")

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
