import gc
import base64
import io
import logging
import threading
import requests
import os
import zipfile
from lxml import etree
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

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
            max_w  = COL_B_WIDTH_PX - (PADDING_PX * 2)
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

def fix_image_anchors(xlsx_bytes):
    """
    Post-process the xlsx file to change all oneCellAnchor elements
    to use editAs='oneCell' so images move and hide with rows when filtered.
    """
    try:
        in_buf  = io.BytesIO(xlsx_bytes)
        out_buf = io.BytesIO()

        with zipfile.ZipFile(in_buf, 'r') as zin, \
             zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                data = zin.read(item.filename)

                # Only modify drawing XML files
                if item.filename.startswith('xl/drawings/drawing') and item.filename.endswith('.xml'):
                    try:
                        tree = etree.fromstring(data)
                        ns   = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

                        # Change oneCellAnchor → twoCellAnchor with editAs="oneCell"
                        for anchor in tree.findall(f'{{{ns}}}oneCellAnchor'):
                            # Get existing from marker and image size
                            from_el = anchor.find(f'{{{ns}}}from')
                            ext_el  = anchor.find(f'{{{ns}}}ext')

                            if from_el is not None and ext_el is not None:
                                # Build a to marker from from + ext
                                from_col    = int(from_el.find(f'{{{ns}}}col').text)
                                from_row    = int(from_el.find(f'{{{ns}}}row').text)

                                to_el       = etree.SubElement(anchor, f'{{{ns}}}to')
                                to_col      = etree.SubElement(to_el, f'{{{ns}}}col')
                                to_col.text = str(from_col + 1)
                                to_colOff   = etree.SubElement(to_el, f'{{{ns}}}colOff')
                                to_colOff.text = '0'
                                to_row      = etree.SubElement(to_el, f'{{{ns}}}row')
                                to_row.text = str(from_row + 1)
                                to_rowOff   = etree.SubElement(to_el, f'{{{ns}}}rowOff')
                                to_rowOff.text = '0'

                                # Remove ext element (not used in twoCellAnchor)
                                anchor.remove(ext_el)

                            # Rename tag to twoCellAnchor and add editAs
                            anchor.tag      = f'{{{ns}}}twoCellAnchor'
                            anchor.set('editAs', 'oneCell')

                        data = etree.tostring(tree, xml_declaration=True,
                                              encoding='UTF-8', standalone=True)
                    except Exception as e:
                        logging.warning(f"Drawing XML fix skipped: {e}")

                zout.writestr(item, data)

        out_buf.seek(0)
        return out_buf.read()

    except Exception as e:
        logging.warning(f"Anchor fix failed, returning original: {e}")
        return xlsx_bytes

def process_and_callback(excel_bytes, filename, callback_url):
    try:
        wb = load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        del excel_bytes
        gc.collect()

        # ✅ Freeze header row
        ws.freeze_panes = "A2"

        # ✅ Set column B width
        ws.column_dimensions[IMAGE_COL].width = COL_B_WIDTH

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
                xl_img.anchor = f"B{row_idx}"   # ✅ Safe string anchor
                ws.add_image(xl_img)
                ws.row_dimensions[row_idx].height = ROW_HEIGHT_PT
                processed += 1
            except Exception as e:
                logging.warning(f"Embed failed for {style_str}: {e}")
                skipped += 1
            finally:
                del img_bytes
                gc.collect()

        # ✅ Apply auto filter across full data range
        last_col           = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"

        # ✅ Save workbook to bytes
        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)
        xlsx_bytes = out_buf.read()
        del wb
        gc.collect()

        # ✅ Fix image anchors via XML post-processing
        xlsx_bytes = fix_image_anchors(xlsx_bytes)

        result_b64 = base64.b64encode(xlsx_bytes).decode("utf-8")
        del xlsx_bytes
        gc.collect()

        # ✅ Ensure .xlsx extension
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
        logging.error(f"Background processing failed: {e}")

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