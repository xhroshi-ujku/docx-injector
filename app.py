from flask import Flask, request, send_file, jsonify, abort
import io, os, zipfile, tempfile, traceback
from xml.etree import ElementTree as ET
from copy import deepcopy

app = Flask(__name__)

API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")


@app.before_request
def require_api_key():
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


def get_document_xml(docx_bytes):
    """Extract the document.xml content from a DOCX file."""
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml = z.read("word/document.xml")
    return xml


def replace_placeholder_in_xml(template_xml, placeholder, insert_xml):
    """Replace placeholder text in template_xml with the entire insert_xml content."""
    try:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tree = ET.fromstring(template_xml)
        insert_tree = ET.fromstring(insert_xml)

        # Combine all text in document for search
        text_elements = tree.findall(".//w:t", ns)
        combined_text = "".join([t.text or "" for t in text_elements])

        if placeholder not in combined_text:
            print("⚠️ Placeholder not found in XML.")
            return template_xml  # Return original if not found

        # Iterate through runs and find where placeholder begins
        placeholder_chars = list(placeholder)
        start_index = 0
        found_runs = []

        for t in text_elements:
            if not t.text:
                continue
            for ch in t.text:
                if ch == placeholder_chars[start_index]:
                    start_index += 1
                    found_runs.append(t)
                    if start_index == len(placeholder_chars):
                        break
                else:
                    start_index = 0
                    found_runs = []
            if start_index == len(placeholder_chars):
                break

        # Remove placeholder runs
        for t in found_runs:
            parent = t.getparent()
            parent.remove(t)

        # Insert new content (deep copy of all body elements)
        body = tree.find("w:body", ns)
        insert_body = insert_tree.find("w:body", ns)

        # Append insert_body contents where placeholder was
        for el in list(insert_body):
            body.append(deepcopy(el))

        return ET.tostring(tree, encoding="utf-8", xml_declaration=True)
    except Exception as e:
        print("❌ Error replacing placeholder:", e)
        print(traceback.format_exc())
        return template_xml


def rebuild_docx_with_new_xml(template_bytes, new_xml):
    """Rebuild DOCX with modified document.xml."""
    in_mem = io.BytesIO(template_bytes)
    out_mem = io.BytesIO()

    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(out_mem, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename != "word/document.xml":
                zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr("word/document.xml", new_xml)

    out_mem.seek(0)
    return out_mem


@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector ZIP-based API is running",
        "endpoints": ["/inject-docx"]
    })


@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """Merge one DOCX into another using raw XML manipulation."""
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' files are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        # Extract XML from both DOCX files
        template_xml = get_document_xml(template_bytes)
        source_xml = get_document_xml(source_bytes)

        # Replace placeholder in template with source XML
        new_xml = replace_placeholder_in_xml(template_xml, placeholder, source_xml)

        # Rebuild new DOCX
        merged_docx = rebuild_docx_with_new_xml(template_bytes, new_xml)

        return send_file(
            merged_docx,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
