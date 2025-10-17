from flask import Flask, request, send_file, jsonify
from docx import Document
import html2docx
import io, zipfile, re
import os
from flask import abort
import traceback
print("ERROR:", e)
print(traceback.format_exc())


API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")


@app.before_request
def require_api_key():
    """
    Runs before every request. Checks for x-api-key header.
    """
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")
# --------------------------


# --------------------------
# 🧩 Helper function
# --------------------------
def replace_placeholder_with_html(doc: Document, placeholder: str, html: str):
    """
    Finds the first occurrence of a placeholder string in the DOCX
    and replaces it with formatted HTML content.
    """
    for i, p in enumerate(doc.paragraphs):
        if placeholder in p.text:
            # Split the paragraph around the placeholder
            before, sep, after = p.text.partition(placeholder)
            p.clear()
            if before:
                p.add_run(before)
                p.add_run().add_break()

            # Create a temporary document with the HTML
            tmp_doc = Document()
            html2docx.add_html_to_document(html, tmp_doc)

            # Insert each element from tmp_doc into the main document
            anchor = p._p
            for block in tmp_doc.element.body:
                anchor.addnext(block)
                anchor = block

            # Add text after placeholder, if any
            if after:
                new_para = doc.add_paragraph(after)
                anchor.addnext(new_para._p)
            return True
    return False
# --------------------------


# --------------------------
# 🌐 Endpoints
# --------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API is running",
        "endpoints": ["/inject", "/inject-multi"]
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-injector", "ok": True})


@app.route("/inject", methods=["POST"])
def inject():
    """
    POST /inject
    Multipart form-data:
      - template: DOCX file
      - placeholder: string (e.g. {{Permbajtja}})
      - html: string (HTML content)
    Returns a modified DOCX file.
    """
    if "template" not in request.files:
        return jsonify({"error": "Missing 'template' file"}), 400

    template_file = request.files["template"]
    placeholder = request.form.get("placeholder", "{{Permbajtja}}")
    html = request.form.get("html", "")

    try:
        doc = Document(io.BytesIO(template_file.read()))
        ok = replace_placeholder_with_html(doc, placeholder, html)
        if not ok:
            return jsonify({"error": f"Placeholder '{placeholder}' not found"}), 400

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="injected.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/inject-multi", methods=["POST"])
def inject_multi():
    """
    POST /inject-multi
    Multipart form-data:
      - template: DOCX file
      - map: JSON array of objects [{"placeholder": "...", "html": "..."}]
    Returns a modified DOCX with multiple replacements.
    """
    import json

    if "template" not in request.files:
        return jsonify({"error": "Missing 'template' file"}), 400

    template_file = request.files["template"]
    mapping_raw = request.form.get("map", "[]")

    try:
        mapping = json.loads(mapping_raw)
        doc = Document(io.BytesIO(template_file.read()))

        for m in mapping:
            ph = m.get("placeholder")
            html = m.get("html", "")
            if ph:
                replace_placeholder_with_html(doc, ph, html)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="injected.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
# --------------------------


# --------------------------
# 🚀 Run app
# --------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
