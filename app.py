from flask import Flask, request, send_file, jsonify, abort
from docx import Document
import html2docx
import io, os, json, traceback, re, html

app = Flask(__name__)

# --------------------------
# 🔑 API Key
# --------------------------
API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")

@app.before_request
def require_api_key():
    """Runs before every request. Checks for x-api-key header."""
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# --------------------------
# 🧩 Helper function
# --------------------------
def clean_html(raw_html: str) -> str:
    """Cleans SharePoint-style HTML for compatibility with html2docx."""
    if not raw_html:
        return ""

    # Decode HTML entities and remove problematic tags/classes
    cleaned = html.unescape(raw_html)
    cleaned = cleaned.replace("\\n", " ").replace("\n", " ")
    cleaned = cleaned.replace("&#160;", " ").replace("&nbsp;", " ")
    cleaned = cleaned.replace("&#58;", ":")
    cleaned = re.sub(r'class="[^"]+"', "", cleaned)
    cleaned = cleaned.replace("ExternalClass", "")
    cleaned = cleaned.replace("<o:p>", "").replace("</o:p>", "")
    cleaned = cleaned.replace("<br>", "<br/>").replace("<br />", "<br/>")
    cleaned = cleaned.replace("<div>", "<p>").replace("</div>", "</p>")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def replace_placeholder_with_html(doc: Document, placeholder: str, html_content: str):
    """
    Finds and replaces placeholder text (even if split across runs)
    with formatted HTML. Cleans the input HTML for compatibility.
    """
    html_content = clean_html(html_content)

    for paragraph in doc.paragraphs:
        # Combine all runs in the paragraph
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder in full_text:
            print(f"✅ Found placeholder: {placeholder}")

            # Clear existing runs
            for run in paragraph.runs:
                run.text = ""

            # Convert HTML to a temporary Word doc
            tmp_doc = Document()
            try:
                html2docx.add_html_to_document(html_content, tmp_doc)
            except Exception as err:
                print("❌ HTML conversion error:", err)
                print(traceback.format_exc())
                tmp_doc.add_paragraph("[HTML conversion failed, inserted as plain text]")
                tmp_doc.add_paragraph(html_content)

            # Insert new elements after the placeholder paragraph
            anchor = paragraph._p
            for block in tmp_doc.element.body:
                anchor.addnext(block)
                anchor = block
            return True

    print(f"⚠️ Placeholder '{placeholder}' not found.")
    return False


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
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files:
            return jsonify({"error": "Missing 'template' file"}), 400

        template_file = request.files["template"]
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")
        html = request.form.get("html", "")

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
        print("ERROR:", e)
        print(traceback.format_exc())
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
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files:
            return jsonify({"error": "Missing 'template' file"}), 400

        template_file = request.files["template"]
        mapping_raw = request.form.get("map", "[]")
        mapping = json.loads(mapping_raw)

        doc = Document(io.BytesIO(template_file.read()))

        for m in mapping:
            ph = m.get("placeholder")
            html_content = m.get("html", "")
            if ph:
                replace_placeholder_with_html(doc, ph, html_content)

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
        print("ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# --------------------------
# 🚀 Run app
# --------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
