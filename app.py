from flask import Flask, request, send_file, jsonify, abort
from docx import Document
from copy import deepcopy
import io, os, traceback

app = Flask(__name__)

API_KEY = os.environ.get("DOCX_API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")

@app.before_request
def require_api_key():
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")

def replace_placeholder_with_docx_content(template_doc: Document, placeholder: str, source_doc: Document):
    """
    Replaces placeholder text (even when split across multiple runs)
    with the full formatted content from source_doc.
    """
    for paragraph in template_doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder in full_text:
            print(f"✅ Found placeholder in paragraph: {full_text}")

            # Split into parts
            before, sep, after = full_text.partition(placeholder)

            # Remove all runs
            for run in paragraph.runs:
                run.text = ""

            # Add text before placeholder
            if before:
                paragraph.add_run(before)

            # Get parent XML
            parent = paragraph._element.getparent()
            idx = parent.index(paragraph._element)

            # Insert each element from source_doc
            for element in list(source_doc.element.body):
                parent.insert(idx + 1, deepcopy(element))
                idx += 1

            # Add text after placeholder
            if after.strip():
                new_para = template_doc.add_paragraph(after)
                parent.insert(idx + 1, new_para._element)

            # Remove original paragraph (placeholder)
            parent.remove(paragraph._element)
            return True

    print("⚠️ Placeholder not found in any paragraph.")
    return False

@app.route("/", methods=["GET"])
def root():
    return jsonify({"status": "ok", "message": "DOCX Injector API is running"})

@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' files are required"}), 400

        template_file = request.files["template"]
        source_file = request.files["source"]
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        # Load both DOCX files
        template_doc = Document(io.BytesIO(template_file.read()))
        source_doc = Document(io.BytesIO(source_file.read()))

        ok = replace_placeholder_with_docx_content(template_doc, placeholder, source_doc)
        if not ok:
            return jsonify({"error": f"Placeholder '{placeholder}' not found"}), 400

        output = io.BytesIO()
        template_doc.save(output)
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
