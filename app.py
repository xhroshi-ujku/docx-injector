from flask import Flask, request, send_file, jsonify, abort
from docxtpl import DocxTemplate
import os, traceback

app = Flask(__name__)

# ------------------------------------------------------------
# 🔑 API Key Security
# ------------------------------------------------------------
API_KEY = os.environ.get(
    "DOCX_API_KEY",
    "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997"
)


@app.before_request
def require_api_key():
    """Protect all endpoints except root and status."""
    if request.path in ["/", "/status"]:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        abort(401, description="Invalid or missing API key")


# ------------------------------------------------------------
# ⚙️ Injection Logic
# ------------------------------------------------------------
def inject_with_template(template_path, source_path, output_path, placeholder="my_content"):
    """
    Inject the content of a source .docx into a Jinja placeholder in the template.
    Template must include: {{ p(my_content) }}
    """
    try:
        print("=== DEBUG: Starting injection ===", flush=True)
        print(f"Template: {template_path}", flush=True)
        print(f"Source:   {source_path}", flush=True)
        print(f"Output:   {output_path}", flush=True)
        print(f"Placeholder: {placeholder}", flush=True)

        tpl = DocxTemplate(template_path)
        tpl.render_jinja_env.globals['p'] = tpl.build_paragraph
        print("DEBUG: Template loaded and p() registered", flush=True)

        subdoc = tpl.new_subdoc(source_path)
        print("DEBUG: Created subdoc successfully", flush=True)

        tpl.render({placeholder: subdoc})
        print("DEBUG: Rendered context successfully", flush=True)

        tpl.save(output_path)
        print("DEBUG: tpl.save() executed", flush=True)

        if os.path.exists(output_path):
            print("✅ File saved successfully:", output_path, flush=True)
            return True
        else:
            print("❌ tpl.save() did not create the file!", flush=True)
            return False

    except Exception as e:
        print("❌ Injection error:", e, flush=True)
        print(traceback.format_exc(), flush=True)
        return False



# ------------------------------------------------------------
# 🌐 API Endpoint
# ------------------------------------------------------------
@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    POST /inject-docx
    Multipart form-data:
      - template: DOCX template (must include {{ p(my_content) }})
      - source: DOCX file to inject
      - placeholder: optional string (default = 'my_content')
    """
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_file = request.files["template"]
        source_file = request.files["source"]
        placeholder = request.form.get("placeholder", "my_content")

        # Save files to temporary paths
        template_path = "/tmp/template.docx"
        source_path = "/tmp/source.docx"
        output_path = "/tmp/merged.docx"

        template_file.save(template_path)
        source_file.save(source_path)

        success = inject_with_template(template_path, source_path, output_path, placeholder)
        if not success:
            return jsonify({"error": "Injection failed"}), 500

        # ✅ Send merged DOCX
        return send_file(
            output_path,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ General error:", e)
        print(traceback.format_exc(), flush=True)
        return jsonify({"error": str(e)}), 500


# ------------------------------------------------------------
# 🧭 Root and Status Endpoints
# ------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "service": "docx-jinja-injector",
        "ok": True,
        "usage": "POST /inject-docx with form-data: template, source, [placeholder]",
        "template_placeholder_example": "{{ p(my_content) }}"
    })


@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-jinja-injector", "ok": True})


# ------------------------------------------------------------
# 🚀 Run locally or on Render
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

