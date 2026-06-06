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
from PIL import Image as PILImage
import xlsxwriter

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

STYLE_COL_IDX   = 0
IMAGE_COL_WIDTH = 14
ROW_HEIGHT_PT   = 72
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 6
IMG_SIZE_PX = 120

def download_and_resize(style_str):
    try:
        resp = requests.get(f"{IMAGE_URL_BASE}{style_str}", timeout=REQUEST_TIMEOUT, stream=True)
        resp.raise_for_status()

        raw = io.BytesIO(resp.content)
        del resp

        with PILImage.open(raw) as img:
            img = img.convert("RGB")
            img = img.resize((IMG_SIZE_PX, IMG_SIZE_PX), PILImage.LANCZOS)

            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=60, optimize=True, progressive=True)
            data = buf.getvalue()

            logging.info(f"IMG {style_str}: {len(data)/1024:.1f} KB, JPEG {IMG_SIZE_PX}x{IMG_SIZE_PX}")
            return data

    except Exception as e:
        logging.warning(f"Image failed for {style_str}: {e}")
        return None
    finally:
        gc.collect()

def process_and_callback(excel_bytes, filename, callback_url):
    try:
        logging.info(f"Received file size: {len(excel_bytes)} bytes")

        # ── Read source data with openpyxl ───────────────────────────────────
        wb = load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        del excel_bytes
        gc.collect()

        if ws['A1'].value and str(ws['A1'].value).strip().lower() == "image":
            logging.warning("File already processed — skipping")
            return

        headers  = [cell.value for cell in ws[1]]
        num_cols = len(headers)

        all_rows = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if any(v is not None for v in row):
                all_rows.append(list(row))

        del wb
        gc.collect()

        # Sort A→Z by Style
        all_rows.sort(key=lambda r: (
            r[STYLE_COL_IDX] is None,
            str(r[STYLE_COL_IDX]).lower() if r[STYLE_COL_IDX] else ""
        ))
        logging.info(f"Read {len(all_rows)} rows, sorted A→Z ✅")

        # ── Download images ───────────────────────────────────────────────────
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

        # ── Calculate column widths ───────────────────────────────────────────
        col_widths = []
        for col_idx in range(num_cols):
            max_len = len(str(headers[col_idx])) if headers[col_idx] else 0
            for row_data in all_rows:
                if col_idx < len(row_data) and row_data[col_idx] is not None:
                    max_len = max(max_len, len(str(row_data[col_idx])))
            col_widths.append(min(max_len + 4, 50))

        # ── Build xlsx with XlsxWriter ────────────────────────────────────────
        logging.info("Building xlsx with XlsxWriter embed_image()...")

        # ✅ File-based (not in_memory) — more stable for large files
        tmp_path = "/tmp/xlsxwriter_output.xlsx"
        workbook = xlsxwriter.Workbook(tmp_path)
        sheet    = workbook.add_worksheet()

        bold_fmt = workbook.add_format({'bold': True,  'font_name': 'Calibri', 'font_size': 11})
        cell_fmt = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
        date_fmt = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': 'mm/dd/yyyy'})
        num_fmt  = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})

        # Header row
        sheet.write(0, 0, "Image", bold_fmt)
        for col_idx, header in enumerate(headers):
            sheet.write(0, col_idx + 1, header if header is not None else "", bold_fmt)

        # Column widths
        sheet.set_column(0, 0, IMAGE_COL_WIDTH)
        for col_idx, width in enumerate(col_widths):
            sheet.set_column(col_idx + 1, col_idx + 1, width)

        # Data rows with embedded images
        for row_idx, row_data in enumerate(all_rows, start=1):
            sheet.set_row(row_idx, ROW_HEIGHT_PT)

            style_val = row_data[STYLE_COL_IDX]
            if style_val:
                style_str = str(style_val).strip()
                img_bytes = images.get(style_str)
                if img_bytes:
                    try:
                        # ✅ .jpg extension tells XlsxWriter to treat as JPEG
                        sheet.embed_image(row_idx, 0, "image.jpg", {
                            'image_data': io.BytesIO(img_bytes)
                        })
                    except Exception as e:
                        logging.warning(f"embed_image failed for {style_str}: {e}")

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

        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(all_rows), num_cols)
        workbook.close()

        # ✅ Read from temp file
        with open(tmp_path, 'rb') as f:
            output_bytes = f.read()

        # Clean up temp file
        try:
            os.remove(tmp_path)
        except Exception:
            pass

        if output_bytes.startswith(b'PK\x03\x04'):
            logging.info(f"Output xlsx VALID ✅ size: {len(output_bytes)/1024/1024:.2f} MB")
        else:
            logging.error("Output xlsx INVALID ❌")

        result_b64 = base64.b64encode(output_bytes).decode("utf-8")
        logging.info(f"Base64 length: {len(result_b64)}")
        del output_bytes
        gc.collect()

        if not filename.lower().endswith(".xlsx"):
            filename = filename + ".xlsx"

        payload = {
            "file"      : result_b64,
            "filename"  : filename,
            "processed" : processed,
            "skipped"   : skipped
        }
        resp = requests.post(callback_url, json=payload, timeout=30)
        logging.info(f"Callback: {resp.status_code}, processed={processed}, skipped={skipped}")

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
