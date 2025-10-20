from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
import tempfile
import base64
import traceback
import os

app = Flask(__name__)

@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    try:
        # Parse incoming JSON
        data = request.get_json(force=True)
        print("📦 Incoming JSON:", data.keys())

        # Validate required fields
        if "template" not in data:
            return jsonify({"error": "Missing 'template' (Base64 of template.docx)"}), 400
        if "source" not in data:
            return jsonify({"error": "Missing 'source' (Base64 of source.docx)"}), 400

        # Decode and write template.docx to temp file
        tpl_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        tpl_temp.write(base64.b64decode(data["template"]))
        tpl_temp.close()

        # Decode and write source.docx to temp file
        src_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        src_temp.write(base64.b64decode(data["source"]))
        src_temp.close()

        # Load template dynamically
        tpl = DocxTemplate(tpl_temp.name)
        subdoc = tpl.new_subdoc(src_temp.name)

        # Build context for placeholders
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

        # Render and save output
        tpl.render(context)
        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        tpl.save(output_file.name)

        # Clean up temp template and source
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
        return jsonify({
            "error": str(e),
            "trace": traceback.format_exc()
        }), 500

@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok", "message": "Dynamic DOCX merge API is running."})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
