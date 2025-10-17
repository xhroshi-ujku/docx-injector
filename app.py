from flask import Flask, request, send_file, jsonify
from docx import Document
import html2docx
import io, zipfile, re

app = Flask(__name__)

def replace_placeholder_with_html(doc: Document, placeholder: str, html: str):
    """
    Find a paragraph containing the placeholder text and replace that single
    placeholder with a formatted block converted from HTML.
    """
    html2docx.add_html_to_document(html, tmp_doc)

    for i, p in enumerate(doc.paragraphs):
        if placeholder in p.text:
            # Split the paragraph around the placeholder
            before, sep, after = p.text.partition(placeholder)
            # Clear the paragraph and put 'before' back (if any)
            p.clear()
            if before:
                p.add_run(before)

            # Insert the HTML as new content directly after this paragraph
            # Strategy: create a temp doc with the HTML, then import its elements here.
            tmp_doc = Document()
            html2docx.add_html_to_document(html, tmp_doc)

            # Insert each element from tmp_doc after the current paragraph
            # Keep a reference to the current paragraph
            anchor = p._p

            # Add a run break before HTML block if there was preceding text
            if before:
                p.add_run().add_break()

            # The tmp_doc will have paragraphs, possibly tables. We copy them.
            # Paragraph copy:
            for block in tmp_doc.element.body:
                anchor.addnext(block)  # insert after anchor
                anchor = block         # move anchor

            # If there is trailing text after the placeholder, put it into a new paragraph
            if after:
                new_para = doc.add_paragraph(after)
                anchor.addnext(new_para._p)

            return True

    return False

@app.route("/inject", methods=["POST"])
def inject():
    """
    Multipart/form-data:
      - template: file (docx)
      - placeholder: string (e.g., {{Permbajtja}})
      - html: string (HTML content)
    Returns: modified DOCX
    """
    if "template" not in request.files:
        return jsonify({"error": "Missing 'template' file"}), 400
    template_file = request.files["template"]

    placeholder = request.form.get("placeholder", "{{Permbajtja}}")
    html = request.form.get("html", "")

    try:
        doc = Document(template_file)
        ok = replace_placeholder_with_html(doc, placeholder, html)
        if not ok:
            return jsonify({"error": f"Placeholder '{placeholder}' not found"}), 400

        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        return send_file(
            out,
            as_attachment=True,
            download_name="injected.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET"])
def root():
    return jsonify({"status": "ok", "endpoints": ["/inject"]})

if __name__ == "__main__":
    # Run on all interfaces for tunneling tools (ngrok, etc.)
    app.run(host="0.0.0.0", port=5000)
