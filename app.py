from flask import Flask, request, send_file, jsonify, abort
import io, os, zipfile, traceback
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
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.read("word/document.xml"), z.namelist()


def rebuild_docx_with_new_xml(template_bytes, new_xml):
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


def merge_split_runs(tree, ns):
    """Merge split text runs so placeholders like {{Permbajtja}} can be detected."""
    for p in tree.findall(".//w:p", ns):
        texts = []
        for r in p.findall(".//w:t", ns):
            texts.append(r.text or "")
        full_text = "".join(texts)
        yield p, full_text


def replace_placeholder_in_xml(template_xml, source_xml, placeholder="{{Permbajtja}}"):
    try:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        template_tree = ET.fromstring(template_xml)
        source_tree = ET.fromstring(source_xml)

        body = template_tree.find(".//w:body", ns)
        if body is None:
            print("⚠️ Template has no body.")
            return template_xml

        insert_body = source_tree.find(".//w:body", ns)
        if insert_body is None:
            print("⚠️ Source has no body.")
            return template_xml

        for p, full_text in merge_split_runs(template_tree, ns):
            if placeholder in full_text:
                print("✅ Found fragmented placeholder, replacing...")
                parent = p.getparent() if hasattr(p, "getparent") else body
                idx = list(body).index(p)
                body.remove(p)
                for el in list(insert_body):
                    body.insert(idx, deepcopy(el))
                    idx += 1
                return ET.tostring(template_tree, encoding="utf-8", xml_declaration=True)

        print("⚠️ Placeholder not found even after merge.")
        return template_xml
    except Exception as e:
        print("❌ Replacement error:", e)
        print(traceback.format_exc())
        return template_xml


@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API (split-run safe) is running",
        "endpoints": ["/inject-docx"]
    })


@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        template_xml, _ = get_document_xml(template_bytes)
        source_xml, _ = get_document_xml(source_bytes)

        new_xml = replace_placeholder_in_xml(template_xml, source_xml, placeholder)
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
