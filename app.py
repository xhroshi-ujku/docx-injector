from flask import Flask, request, send_file, jsonify, abort
import io, os, zipfile, traceback
from xml.etree import ElementTree as ET
from copy import deepcopy

app = Flask(__name__)

# ------------------------------------------------------------
# 🔑 API Key
# ------------------------------------------------------------
API_KEY = os.environ.get(
    "DOCX_API_KEY",
    "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997"
)


@app.before_request
def require_api_key():
    """Check API key for all routes except root and status."""
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# ------------------------------------------------------------
# ⚙️ DOCX Utilities
# ------------------------------------------------------------
def get_document_xml(docx_bytes):
    """Extract document.xml from DOCX and return its bytes and file list."""
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.read("word/document.xml"), z.namelist()


def rebuild_docx_with_new_xml(template_bytes, new_xml):
    """Rebuild a DOCX with modified document.xml content."""
    in_mem = io.BytesIO(template_bytes)
    out_mem = io.BytesIO()

    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(
        out_mem, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename != "word/document.xml":
                zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr("word/document.xml", new_xml)

    out_mem.seek(0)
    return out_mem


def merge_split_runs(tree, ns):
    """Merge text runs so placeholders like {{Permbajtja}} can be detected even when split."""
    for p in tree.findall(".//w:p", ns):
        texts = []
        for r in p.findall(".//w:t", ns):
            texts.append(r.text or "")
        full_text = "".join(texts)
        yield p, full_text


def replace_placeholder_in_xml(template_xml, source_xml, placeholder="{{Permbajtja}}"):
    """
    Finds a split placeholder like {{Permbajtja}} in the template DOCX and
    replaces it with all paragraphs from the source DOCX body — safely.
    Keeps final XML structure valid.
    """
    try:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        ET.register_namespace("w", ns["w"])

        template_tree = ET.fromstring(template_xml)
        source_tree = ET.fromstring(source_xml)

        template_body = template_tree.find(".//w:body", ns)
        source_body = source_tree.find(".//w:body", ns)

        if template_body is None or source_body is None:
            print("⚠️ Missing body in one of the DOCX files.")
            return template_xml

        # Extract source content except sectPr (Word section properties)
        source_elems = [
            deepcopy(el) for el in list(source_body)
            if not el.tag.endswith("sectPr")
        ]

        # Locate the paragraph containing the placeholder (even if split)
        for p, full_text in merge_split_runs(template_tree, ns):
            if placeholder in full_text:
                print("✅ Found placeholder in paragraph; replacing content...")
                idx = list(template_body).index(p)
                template_body.remove(p)

                # Insert each element from source
                for el in source_elems:
                    template_body.insert(idx, el)
                    idx += 1

                print("✅ Injection complete and structure preserved.")
                return ET.tostring(template_tree, encoding="utf-8", xml_declaration=True)

        print("⚠️ Placeholder not found.")
        return template_xml

    except Exception as e:
        print("❌ Replacement error:", e)
        print(traceback.format_exc())
        return template_xml



# ------------------------------------------------------------
# 🌐 Endpoints
# ------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API (split-run safe) is running",
        "endpoints": ["/inject-docx", "/debug-scan"]
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-injector", "ok": True})


@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    POST /inject-docx
    Multipart form-data:
      - template: DOCX file (required)
      - source: DOCX file (required)
      - placeholder: optional string (default {{Permbajtja}})
    Returns: merged DOCX file
    """
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

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


@app.route("/debug-scan", methods=["POST"])
def debug_scan():
    """
    POST /debug-scan
    Multipart form-data:
      - file: DOCX file
    Returns: JSON telling if {{Permbajtja}} exists, even if split across runs.
    """
    try:
        if "file" not in request.files:
            return jsonify({"error": "Missing 'file'"}), 400

        docx_bytes = request.files["file"].read()
        xml_bytes, names = get_document_xml(docx_bytes)
        xml_str = xml_bytes.decode("utf-8", errors="ignore")

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tree = ET.fromstring(xml_bytes)
        merged_texts = ["".join(r.text or "" for r in p.findall(".//w:t", ns))
                        for p in tree.findall(".//w:p", ns)]
        merged_full_text = " ".join(merged_texts)

        found_literal = "{{Permbajtja}}" in xml_str
        found_merged = "{{Permbajtja}}" in merged_full_text

        return jsonify({
            "contains_word_document": "word/document.xml" in names,
            "found_literal": found_literal,
            "found_merged_text": found_merged,
            "xml_length": len(xml_str)
        })

    except Exception as e:
        print("❌ DEBUG error:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500
@app.route("/debug-inject-test", methods=["POST"])
def debug_inject_test():
    """
    Diagnostic route: uploads two DOCX files and shows internal results
    without generating a DOCX — just text logs.
    """
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        from xml.etree import ElementTree as ET

        # Extract XMLs
        with zipfile.ZipFile(io.BytesIO(template_bytes)) as zt:
            template_xml = zt.read("word/document.xml").decode("utf-8")

        with zipfile.ZipFile(io.BytesIO(source_bytes)) as zs:
            source_xml = zs.read("word/document.xml").decode("utf-8")

        # Step A: check if placeholder exists
        found_placeholder = placeholder in template_xml
        found_fragments = "{{" in template_xml or "}}" in template_xml

        # Step B: check if the source DOCX has a <w:body> and paragraphs
        has_body = "<w:body" in source_xml
        para_count = source_xml.count("<w:p")

        # Step C: debug XML structure (only first lines)
        sample_template = "\n".join(template_xml.splitlines()[:10])
        sample_source = "\n".join(source_xml.splitlines()[:10])

        return jsonify({
            "found_placeholder_exact": found_placeholder,
            "found_placeholder_fragments": found_fragments,
            "source_body_present": has_body,
            "source_paragraphs_count": para_count,
            "template_snippet": sample_template,
            "source_snippet": sample_source
        })

    except Exception as e:
        print("❌ Debug error:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500
@app.route("/debug-find-placeholder", methods=["POST"])
def debug_find_placeholder():
    """
    Inspect all paragraphs (<w:p>) and check which ones contain placeholder fragments.
    """
    try:
        if "template" not in request.files:
            return jsonify({"error": "template file required"}), 400

        placeholder = request.form.get("placeholder", "{{Permbajtja}}")
        template_bytes = request.files["template"].read()

        with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
            xml_str = z.read("word/document.xml").decode("utf-8")

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        ET.register_namespace("w", ns["w"])
        tree = ET.fromstring(xml_str)

        paragraphs = []
        for p in tree.findall(".//w:p", ns):
            texts = []
            for r in p.findall(".//w:t", ns):
                texts.append(r.text or "")
            merged = "".join(texts)
            if "Permbajtja" in merged or "{{" in merged or "}}" in merged:
                paragraphs.append({
                    "merged_text": merged,
                    "raw_xml_snippet": ET.tostring(p, encoding="unicode")[:500]
                })

        return jsonify({
            "placeholder_paragraphs_found": len(paragraphs),
            "paragraphs": paragraphs[:3]  # show max 3 for readability
        })

    except Exception as e:
        print("❌ Debug error:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# ------------------------------------------------------------
# 🚀 Run locally
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)




