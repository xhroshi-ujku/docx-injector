from flask import Flask, request, send_file, jsonify, abort
from docxtpl import DocxTemplate
from docx import Document
import io, os, traceback

app = Flask(__name__)

# ------------------------------------------------------------
# API Key Security
# ------------------------------------------------------------
API_KEY = os.environ.get(
    "DOCX_API_KEY",
    "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997"
)

@app.before_request
def require_api_key():
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# ------------------------------------------------------------
# Core injection logic
# ------------------------------------------------------------
def inject_with_template(template_path, source_path, output_path, placeholder="my_content"):
    """
    Injects the content of a source .docx into a Jinja placeholder in the template.
    Template must include: {{ p my_content }} or {{ my_content }}
    """
    try:
        tpl = DocxTemplate(template_path)

        # Create a subdoc from the source document
        subdoc = tpl.new_subdoc()
        subdoc.subdocx = Document(source_path)

        # Inject it into the placeholder context
        context = {placeholder: subdoc}
        tpl.render(context)
        tpl.save(output_path)
        print(f"✅ Injected '{placeholder}' from '{source_path}' into '{template_path}' → '{output_path}'", flush=True)
        return True
    except Exception as e:
        print("❌ Injection error:", e)
        print(traceback.format_exc())
        return False


# ------------------------------------------------------------
# Flask endpoint
# ------------------------------------------------------------
@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    POST /inject-docx
    Multipart form-data:
      - template: DOCX template (must include {{ p my_content }})
      - source: DOCX source file
      - placeholder: optional (defaults to 'my_content')
    """
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_file = request.files["template"]
        source_file = request.files["source"]
        placeholder = request.form.get("placeholder", "my_content")

        template_path = "/tmp/template.docx"
        source_path = "/tmp/source.docx"
        output_path = "/tmp/merged.docx"

        template_file.save(template_path)
        source_file.save(source_path)

        ok = inject_with_template(template_path, source_path, output_path, placeholder)
        if not ok:
            return jsonify({"error": "Injection failed"}), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "service": "docx-jinja-injector",
        "ok": True,
        "usage": "POST /inject-docx with form-data: template, source, [placeholder]"
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-jinja-injector", "ok": True})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

