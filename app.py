from flask import Flask, request, send_file, jsonify, abort
from docx import Document
import io, os, traceback, json

app = Flask(__name__)

# --------------------------
# 🔑 API Key
# --------------------------
API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")

@app.before_request
def require_api_key():
    """Simple x-api-key header check."""
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# --------------------------
# 🧩 Helper functions
# --------------------------
def clone_docx_content(source_doc: Document):
    """Returns all paragraphs and tables from a DOCX body."""
    blocks = []
    for element in source_doc.element.body:
        blocks.append(element)
    return blocks


def replace_placeholder_with_docx_content(target_doc: Document, placeholder: str, blocks):
    """
    Replaces placeholder text with formatted DOCX blocks.
    Works even if the placeholder is split across runs.
    """
    for paragraph in target_doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder in full_text:
            print(f"✅ Found placeholder: {placeholder}")
            for run in paragraph.runs:
                run.text = ""

            # Insert all source blocks after this paragraph
            anchor = paragraph._p
            for block in blocks:
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
        "message": "DOCX Rich Text Injector is running",
        "usage": {
            "endpoint": "/inject-docx",
            "method": "POST",
            "fields": ["template", "source", "placeholder"]
        }
    })


@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    POST /inject-docx
    Multipart form-data:
      - template: DOCX file containing placeholder (e.g. {{Permbajtja}})
      - source: DOCX file containing the rich text content to insert
      - placeholder: string (optional, default = {{Permbajtja}})
    Returns: Modified DOCX file.
    """
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Missing required files: 'template' and/or 'source'"}), 400

        template_file = request.files["template"]
        source_file = request.files["source"]
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        target_doc = Document(io.BytesIO(template_file.read()))
        source_doc = Document(io.BytesIO(source_file.read()))

        # Extract formatted content from source
        content_blocks = clone_docx_content(source_doc)

        # Replace placeholder
        ok = replace_placeholder_with_docx_content(target_doc, placeholder, content_blocks)
        if not ok:
            return jsonify({"error": f"Placeholder '{placeholder}' not found"}), 400

        # Save output
        output = io.BytesIO()
        target_doc.save(output)
        output.seek(0)

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
