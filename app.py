from flask import Flask, request, send_file, jsonify, abort
from docx import Document
from copy import deepcopy
import io, os, traceback

app = Flask(__name__)

# --------------------------
# 🔑 API Key
# --------------------------
API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")


# --------------------------
# 🛡️ Security middleware
# --------------------------
@app.before_request
def require_api_key():
    """
    Require API key for all endpoints except root and status.
    """
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# --------------------------
# ⚙️ Helper function
# --------------------------
def replace_placeholder_with_docx_content(template_doc: Document, placeholder: str, source_doc: Document):
    """
    Safely replaces placeholder text in template_doc with all formatted content from source_doc.
    Uses deepcopy to avoid shared XML references (prevents unreadable content errors).
    """
    for paragraph in template_doc.paragraphs:
        if placeholder in paragraph.text:
            before, sep, after = paragraph.text.partition(placeholder)
            p = paragraph
            p.clear()

            if before:
                p.add_run(before)
                p.add_run().add_break()

            # Deep copy all content to preserve structure safely
            anchor = p._p
            for element in list(source_doc.element.body):
                anchor.addnext(deepcopy(element))
                anchor = element

            if after:
                new_para = template_doc.add_paragraph(after)
                anchor.addnext(new_para._p)

            return True
    return False


# --------------------------
# 🌐 Endpoints
# --------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "status": "ok",
        "message": "DOCX Injector API is running",
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
      - template: DOCX file (destination)
      - source: DOCX file (content to inject)
      - placeholder: string (default: {{Permbajtja}})
    """
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' files are required"}), 400

        template_file = request.files["template"]
        source_file = request.files["source"]
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        # Load DOCX files safely
        template_doc = Document(io.BytesIO(template_file.read()))
        source_doc = Document(io.BytesIO(source_file.read()))

        # Inject content
        ok = replace_placeholder_with_docx_content(template_doc, placeholder, source_doc)
        if not ok:
            return jsonify({"error": f"Placeholder '{placeholder}' not found in template"}), 400

        # Save final DOCX in memory
        output = io.BytesIO()
        template_doc.save(output)
        output.seek(0)

        # Return the merged DOCX
        return send_file(
            output,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# --------------------------
# 🚀 Run app
# --------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
