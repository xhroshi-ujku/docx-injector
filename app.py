from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
import traceback
import base64
import tempfile
import os

app = Flask(__name__)

@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "template.docx")

        if not os.path.exists(template_path):
            return jsonify({"error": "template.docx not found"}), 404

        # Parse JSON input
        data = request.get_json(force=True)
        print("📦 Incoming JSON:", data)

        # Optional: handle base64-encoded source.docx
        source_docx_path = None
        if "Permbajtja" in data and isinstance(data["Permbajtja"], str) and data["Permbajtja"].startswith("UEs"):
            # Save base64 string as temporary source.docx
            decoded = base64.b64decode(data["Permbajtja"])
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tmp.write(decoded)
            tmp.close()
            source_docx_path = tmp.name
        else:
            # fallback to local source.docx
            source_docx_path = os.path.join(base_dir, "source.docx")

        tpl = DocxTemplate(template_path)
        subdoc = tpl.new_subdoc(source_docx_path)

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

        output_path = os.path.join(base_dir, "merged.docx")
        tpl.save(output_path)
        print("🎉 Document created successfully")

        return send_file(
            output_path,
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
    return jsonify({"status": "ok", "message": "DOCX merge API is running."})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
