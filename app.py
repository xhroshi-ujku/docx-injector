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

# ------------------------------------------------------------
# 🧩 NEW Robust Injection Logic
# ------------------------------------------------------------
def inject_docx_content(template_bytes, source_bytes, placeholder="{{Permbajtja}}"):
    """
    Replaces placeholder paragraph with paragraphs from the source DOCX,
    excluding <w:sectPr> to prevent duplicate section properties.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ET.register_namespace("w", ns["w"])

    # Extract XMLs
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zt:
        template_xml = zt.read("word/document.xml").decode("utf-8")
        template_files = {n: zt.read(n) for n in zt.namelist() if n.startswith("word/")}

    with zipfile.ZipFile(io.BytesIO(source_bytes)) as zs:
        source_xml = zs.read("word/document.xml").decode("utf-8")

    template_tree = ET.fromstring(template_xml)
    source_tree = ET.fromstring(source_xml)

    template_body = template_tree.find(".//w:body", ns)
    source_body = source_tree.find(".//w:body", ns)

    if template_body is None or source_body is None:
        raise ValueError("Missing <w:body> in one of the documents")

    # ✅ Exclude <w:sectPr> from source
    source_content = [deepcopy(el) for el in source_body if not el.tag.endswith("sectPr")]

    # ✅ Find paragraph with placeholder
    target_p = None
    for p in template_body.findall(".//w:p", ns):
        merged = "".join([(t.text or "") for t in p.findall(".//w:t", ns)])
        if placeholder in merged:
            target_p = p
            break

    if target_p is None:
        raise ValueError(f"Placeholder '{placeholder}' not found in template")

    # ✅ Replace target paragraph with source content
    children = list(template_body)
    idx = children.index(target_p)
    template_body.remove(target_p)
    for i, el in enumerate(source_content):
        template_body.insert(idx + i, el)

    # ✅ Rebuild the DOCX
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in template_files.items():
            if name == "word/document.xml":
                zout.writestr(
                    name,
                    ET.tostring(template_tree, encoding="utf-8", xml_declaration=True)
                )
            else:
                zout.writestr(name, data)
    output.seek(0)
    return output.getvalue()

# ------------------------------------------------------------
# 🌐 Endpoints
# ------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API (split-run safe) is running",
        "endpoints": ["/inject-docx", "/debug-scan", "/debug-rebuild"]
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
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        injected_bytes = inject_docx_content(template_bytes, source_bytes, placeholder)

        return send_file(
            io.BytesIO(injected_bytes),
            as_attachment=True,
            download_name="injected.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# ------------------------------------------------------------
# 🧪 Diagnostics
# ------------------------------------------------------------
@app.route("/debug-scan", methods=["POST"])
def debug_scan():
    """Scan for placeholder presence in a DOCX."""
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
    """Same as before – diagnostic placeholder test."""
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        # Extract XMLs
        with zipfile.ZipFile(io.BytesIO(template_bytes)) as zt:
            template_xml = zt.read("word/document.xml").decode("utf-8")

        with zipfile.ZipFile(io.BytesIO(source_bytes)) as zs:
            source_xml = zs.read("word/document.xml").decode("utf-8")

        found_placeholder = placeholder in template_xml
        found_fragments = "{{" in template_xml or "}}" in template_xml
        has_body = "<w:body" in source_xml
        para_count = source_xml.count("<w:p")

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
    """Inspect all paragraphs and show merged text if placeholder fragments exist."""
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
            "paragraphs": paragraphs[:3]
        })

    except Exception as e:
        print("❌ Debug error:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# ------------------------------------------------------------
# 🧱 NEW Diagnostic: Debug Rebuild
# ------------------------------------------------------------
@app.route("/debug-rebuild", methods=["POST"])
def debug_rebuild():
    """
    Upload a DOCX -> unzip and rezip -> confirm byte validity.
    This checks if the rebuilding process itself introduces corruption.
    """
    try:
        if "file" not in request.files:
            return jsonify({"error": "Missing 'file'"}), 400

        file_bytes = request.files["file"].read()

        # Extract all
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zin:
            files = {n: zin.read(n) for n in zin.namelist()}

        # Rebuild it
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for n, b in files.items():
                zout.writestr(n, b)

        return jsonify({
            "file_count": len(files),
            "word_document_present": "word/document.xml" in files,
            "rebuilt_size": len(out_buf.getvalue()),
            "status": "ok - rebuild successful"
        })
    except Exception as e:
        print("❌ Rebuild test failed:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


@app.route("/debug-extract-xml", methods=["POST"])
def debug_extract_xml():
    """
    Upload a DOCX and extract both the start and end of word/document.xml
    to detect malformed or duplicated closing tags (common corruption cause).
    """
    try:
        if "file" not in request.files:
            return jsonify({"error": "Missing 'file'"}), 400

        docx_bytes = request.files["file"].read()
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml_bytes = z.read("word/document.xml")
        xml_str = xml_bytes.decode("utf-8", errors="ignore")

        return jsonify({
            "length": len(xml_bytes),
            "start": xml_str[:1500],
            "end": xml_str[-1500:]
        })

    except Exception as e:
        print("❌ Debug extract error:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# ------------------------------------------------------------
# 🚀 Run locally
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)



