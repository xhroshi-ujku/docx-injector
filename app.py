from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
import tempfile
import base64
import traceback
import os

app = Flask(__name__)

# ✅ Set your secret API key (you’ll also set this in Render dashboard as ENV var)
API_KEY = os.environ.get("API_KEY", "eNdertuamFshatinEBemeQytetPartiNenaNenaJonePerjete1997")

@app.before_request
def verify_api_key():
    # Skip check for the home page
    if request.path == "/":
        return None

    # Check for API key in headers
    client_key = request.headers.get("X-API-Key")
    if not client_key or client_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    try:
        data = request.get_json(force=True)
        print("📦 Incoming JSON:", list(data.keys()))

        # Validate required template
        if "template" not in data:
            return jsonify({"error": "Missing 'template' (Base64 of template.docx)"}), 400

        # Decode template.docx
        tpl_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        tpl_temp.write(base64.b64decode(data["template"]))
        tpl_temp.close()

        # ✅ Decode Permbajtja (preferred) or fallback to source
        src_base64 = data.get("Permbajtja") or data.get("source")
        if not src_base64:
            return jsonify({"error": "Missing 'Permbajtja' or 'source' (Base64 DOCX)"}), 400

        src_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        src_temp.write(base64.b64decode(src_base64))
        src_temp.close()

        # Load template
        tpl = DocxTemplate(tpl_temp.name)
        subdoc = tpl.new_subdoc(src_temp.name)

        # Prepare placeholders
        context = {
            "Number": data.get("Number", ""),
            "Date": data.get("Date", ""),
            "Drejtuar": data.get("Drejtuar", ""),
            "Per_dijeni": data.get("Per_dijeni", ""),
            "Subjekti": data.get("Subjekti", ""),
            "Data_Efektive": data.get("Data_Efektive", ""),
            "Data_e_Publikimit": data.get("Data_e_Publikimit", ""),
            "Permbajtja": subdoc,
            "Pergatiti": data.get("Pergatiti", ""),
            "Aprovoi": data.get("Aprovoi", "")
        }

        tpl.render(context)
        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        tpl.save(output_file.name)

        # Cleanup
        os.remove(tpl_temp.name)
        os.remove(src_temp.name)

        print("🎉 Merged DOCX created successfully!")
        return send_file(
            output_file.name,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok", "message": "DOCX merge API is running securely."})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
