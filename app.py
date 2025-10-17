from flask import Flask, request, send_file, jsonify, abort
from docx import Document
import html2docx
import io, os, json, traceback, re, html
from bs4 import BeautifulSoup

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
# 🧩 HTML Cleaning Function
# --------------------------
def simplify_html(raw_html: str) -> str:
    """Cleans and simplifies SharePoint-style HTML for compatibility with html2docx."""
    if not raw_html:
        return ""

    try:
        # Decode entities
        cleaned = html.unescape(raw_html)
        cleaned = cleaned.replace("\\n", " ").replace("\n", " ")
        cleaned = cleaned.replace("&#160;", " ").replace("&nbsp;", " ")
        cleaned = cleaned.replace("&#58;", ":")
        cleaned = cleaned.replace("<br>", "<br/>").replace("<br />", "<br/>")

        # Parse with BeautifulSoup
        soup = BeautifulSoup(cleaned, "html.parser")

        # Remove all style, class, and span attributes
        for tag in soup(True):
            for attr in ["class", "style", "id", "lang", "width", "height", "align"]:
                tag.attrs.pop(attr, None)

        # Convert <div> → <p>
        for div in soup.find_all("div"):
            div.name = "p"

        # Ensure table cells contain plain text
        for td in soup.find_all("td"):
            td.string = td.get_text(strip=True)

        # Remove empty tags
        for tag in soup.find_all():
            if not tag.text.strip() and tag.name not in ["br", "p", "table", "tr", "td"]:
                tag.decompose()

        simplified_html = str(soup)
        simplified_html = re.sub(r"\s+", " ", simplified_html).strip()
        return simplified_html

    except Exception as e:
        print("❌ HTML simplification failed:", e)
        print(traceback.format_exc())
        return raw_html


# --------------------------
# 🧩 Placeholder Replacement
# --------------------------
def replace_placeholder_with_html(doc: Document, placeholder: str, html_content: str):
    """Replaces placeholder text (even across runs) with simplified HTML."""
    html_content = simplify_html(html_content)

    for paragraph in doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder in full_text:
            print(f"✅ Found placeholder: {placeholder}")

            for run in paragraph.runs:
                run.text = ""

            tmp_doc = Document()
            try:
                html2docx.add_html_to_document(html_content, tmp_doc)
            except Exception as err:
                print("❌ HTML conversion error:", err)
                print(traceback.format_exc())
                tmp_doc.add_paragraph("[HTML conversion failed — inserted plain text]")
                tmp_doc.add_paragraph(html_content)

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


@app.route("/inject", methods=["POST"])
def inject():
    """POST /inject"""
    try:
        print("FILES RECEIVED:", list(request.files.keys()))
        print("FORM RECEIVED:", dict(request.form))

        if "template" not in request.files:
            return jsonify({"error": "Missing 'template' file"}), 400

        template_file = request.files["template"]
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")
        html_content = request.form.get("html", "")

        doc = Document(io.BytesIO(template_file.read()))

        ok = replace_placeholder_with_html(doc, placeholder, html_content)
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
    """POST /inject-multi"""
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
