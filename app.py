import gc
import re
import base64
import io
import logging
import threading
import zipfile
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
ROW_HEIGHT_PT   = 72
COL_A_WIDTH     = 18
IMAGE_URL_BASE  = "https://app.townleygirl.com/Image/preview/"
REQUEST_TIMEOUT = 6
ROW_HEIGHT_PX   = int(ROW_HEIGHT_PT * 96 / 72)   # 96px
COL_A_WIDTH_PX  = int(COL_A_WIDTH * 7)            # 126px
PADDING_PX      = 2

def fix_image_anchors(xlsx_bytes):
    """
    Converts oneCellAnchor → twoCellAnchor with editAs='twoCell'
    = 'Move and size with cells' in Excel Format Picture dialog.

    KEY FIX: openpyxl uses NO namespace prefix in drawing XML.
    Tags are <oneCellAnchor> NOT <xdr:oneCellAnchor>.
    Previous code was matching wrong prefix — this is now corrected.
    """
    try:
        in_buf  = io.BytesIO(xlsx_bytes)
        out_buf = io.BytesIO()

        with zipfile.ZipFile(in_buf, 'r') as zin, \
             zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                data = zin.read(item.filename)

                if 'drawings/drawing' in item.filename and item.filename.endswith('.xml'):
                    try:
                        xml_str = data.decode('utf-8')

                        def convert_anchor(match):
                            content = match.group(1)

                            # ✅ No prefix — openpyxl uses default namespace
                            from_m = re.search(r'<from>(.*?)</from>', content, re.DOTALL)
                            if not from_m:
                                return match.group(0)

                            from_content = from_m.group(1)
                            col_m = re.search(r'<col>(\d+)</col>', from_content)
                            row_m = re.search(r'<row>(\d+)</row>', from_content)

                            if not col_m or not row_m:
                                return match.group(0)

                            from_col = int(col_m.group(1))
                            from_row = int(row_m.group(1))

                            # Clean from — zero offsets
                            new_from = (
                                f'<from>'
                                f'<col>{from_col}</col><colOff>0</colOff>'
                                f'<row>{from_row}</row><rowOff>0</rowOff>'
                                f'</from>'
                            )

                            # Add to — next col + row = fills 1 cell exactly
                            new_to = (
                                f'<to>'
                                f'<col>{from_col + 1}</col><colOff>0</colOff>'
                                f'<row>{from_row + 1}</row><rowOff>0</rowOff>'
                                f'</to>'
                            )

                            content = (
                                content[:from_m.start()] +
                                new_from + new_to +
                                content[from_m.end():]
                            )

                            # Remove <ext .../> — not used in twoCellAnchor
                            content = re.sub(r'\s*<ext[^>]*/>', '', content)

                            # ✅ editAs="twoCell" = "Move and size with cells"
                            return f'<twoCellAnchor editAs="twoCell">{content}</twoCellAnchor>'

                        before = xml_str.count('<oneCellAnchor>')
                        xml_str = re.sub(
                            r'<oneCellAnchor>(.*?)</oneCellAnchor>',
                            convert_anchor,
                            xml_str,
                            flags=re.DOTALL
                        )
                        after = xml_str.count('<twoCellAnchor')
                        logging.info(f"Anchors fixed: {before} oneCellAnchor → {after} twoCellAnchor ✅")
                        data = xml_str.encode('utf-8')

                    except Exception as e:
                        logging.warning(f"Drawing XML fix skipped: {e}")

                zout.writestr(item, data)

        out_buf.seek(0)
        result = out_buf.read()

        if result.startswith(b'PK\x03\x04'):
            logging.info(f"Anchor fix complete ✅ size: {len(result)} bytes")
            return result
        else:
            logging.warning("Anchor fix produced invalid zip — using original")
            return xlsx_bytes

    except Exception as e:
        logging.warning(f"Anchor fix failed — using original: {e}")
        return xlsx_bytes

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
            scale  = min(max_w / ow, max_h / oh)
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

        # ✅ Guard: skip already-processed files
        if ws['A1'].value and str(ws['A1'].value).strip().lower() == "image":
            logging.warning("File already processed — skipping")
            return

        # ✅ Insert Image column at position 1
        ws.insert_cols(1)
        ws['A1'] = "Image"

        # ✅ Bold all header cells
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # ✅ Freeze header row
        ws.freeze_panes = "A2"

        # ✅ Set Image column fixed width
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
                xl_img.width  = COL_A_WIDTH_PX
                xl_img.height = ROW_HEIGHT_PX
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

        # ✅ Auto-width all columns
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            if col_letter == IMAGE_COL:
                ws.column_dimensions[col_letter].width = COL_A_WIDTH
                continue
            max_length = 0
            for cell in ws[col_letter]:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_length + 4, 50)

        # ✅ Auto filter
        last_col           = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"

        # ✅ Save
        logging.info("Saving workbook...")
        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)
        output_bytes = out_buf.read()
        del wb
        gc.collect()

        if output_bytes.startswith(b'PK\x03\x04'):
            logging.info(f"Output xlsx VALID ✅ size: {len(output_bytes)} bytes")
        else:
            logging.error("Output xlsx INVALID ❌")

        # ✅ Fix anchors — oneCellAnchor → twoCellAnchor (no xdr: prefix)
        output_bytes = fix_image_anchors(output_bytes)

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
