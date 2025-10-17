from flask import Flask, request, send_file, jsonify, abort
import io, os, zipfile, traceback
from copy import deepcopy
from lxml import etree as ET   # ✅ use lxml instead of xml.etree

app = Flask(__name__)

# ------------------------------------------------------------
# 🔑 API Key
# ------------------------------------------------------------
API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")


@app.before_request
def require_api_key():
    """Validate x-api-key header before every request."""
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# ------------------------------------------------------------
# ⚙️ DOCX Utility Functions
# ------------------------------------------------------------
def get_document_xml(docx_bytes):
    """Extract word/document.xml from a DOCX file."""
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.read("word/document.xml")


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


# ------------------------------------------------------------
# 🧩 Placeholder Replacement Logic (using lxml)
# ------------------------------------------------------------
def replace_placeholder_in_xml(template_xml, placeholder, insert_xml):
    """Robust DOCX injection using lxml to preserve structure."""
    try:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parser = ET.XMLParser(remove_blank_text=True)
        template_tree = ET.fromstring(template_xml, parser)
        insert_tree = ET.fromstring(insert_xml, parser)

        xml_str = ET.tostring(template_tree, encoding="unicode")
        if placeholder not in xml_str:
            print("⚠️ Placeholder not found.")
            return template_xml

        # Iterate through paragraphs
        for p in template_tree.xpath(".//w:p", namespaces=ns):
            p_str = ET.tostring(p, encoding="unicode")
            if placeholder in p_str:
                print("✅ Found placeholder paragraph.")

                parent = p.getparent()
                idx = parent.index(p)
                parent.remove(p)

                insert_body = insert_tree.find("w:body", ns)
                if insert_body is None:
                    print("⚠️ Source DOCX missing <w:body>.")
                    return template_xml

                for child in list(insert_body):
                    parent.insert(idx, deepcopy(child))
                    idx += 1

                print("✅ Injection completed.")
                return ET.tostring(template_tree, encoding="utf-8", xml_declaration=True)

        print("⚠️ Placeholder paragraph not found.")
        return template_xml

    except Exception as e:
        print("❌ Replacement error:", e)
        print(traceback.format_exc())
        return template_xml


# ------------------------------------------------------------
# 🌐 API Endpoints
# ------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API (lxml version) is running",
        "endpoints": ["/inject-docx"]
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-injector", "ok": True})


@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    POST /inject-docx
    Multipart form-data:
      - template: DOCX template file
      - source: DOCX source file
      - placeholder: optional string (default {{Permbajtja}})
    Returns: merged DOCX
    """
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        # Extract XML
        template_xml = get_document_xml(template_bytes)
        source_xml = get_document_xml(source_bytes)

        # Replace placeholder
        new_xml = replace_placeholder_in_xml(template_xml, placeholder, source_xml)

        # Rebuild DOCX
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


# ------------------------------------------------------------
# 🚀 Run
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
