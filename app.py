from flask import Flask, request, send_file, jsonify, abort
import io, os, zipfile, traceback
from xml.etree import ElementTree as ET
from copy import deepcopy

app = Flask(__name__)

# ------------------------------------------------------------
# API key
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
# Helpers
# ------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "r": R_NS}

def read_xml(zf: zipfile.ZipFile, path: str) -> bytes:
    return zf.read(path)

def try_read_xml(zf: zipfile.ZipFile, path: str) -> bytes | None:
    try:
        return zf.read(path)
    except KeyError:
        return None

def get_document_xml(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        return z.read("word/document.xml"), z.namelist()

def merge_split_runs(tree: ET.Element) -> list[tuple[ET.Element, str]]:
    out = []
    for p in tree.findall(".//w:p", NS):
        texts = []
        for t in p.findall(".//w:t", NS):
            texts.append(t.text or "")
        out.append((p, "".join(texts)))
    return out

def next_rid(existing_ids: set[str]) -> str:
    i = 1
    while True:
        rid = f"rId{i}"
        if rid not in existing_ids:
            return rid
        i += 1

def collect_used_rids(xml_bytes: bytes) -> set[str]:
    # Collect any r:embed / r:link / r:id used in drawing, hyperlinks, etc.
    used = set()
    try:
        root = ET.fromstring(xml_bytes)
        # any attribute in R_NS named 'id' or 'embed' or 'link'
        for el in root.iter():
            for k, v in el.attrib.items():
                if k.endswith("}id") or k.endswith("}embed") or k.endswith("}link"):
                    used.add(v)
    except Exception:
        pass
    return used

def merge_relationships_and_media(template_docx: bytes, source_docx: bytes, merged_doc_xml: bytes) -> tuple[bytes, dict[str, bytes], bytes]:
    """
    Returns:
      - updated document.xml (with rId remapped)
      - extra_files: dict[path -> content] to add from source (e.g., word/media/*)
      - updated_rels_xml for word/_rels/document.xml.rels
    """
    tzip = zipfile.ZipFile(io.BytesIO(template_docx), "r")
    szip = zipfile.ZipFile(io.BytesIO(source_docx), "r")

    # Load rels (create minimal if missing)
    t_rels_path = "word/_rels/document.xml.rels"
    s_rels_path = "word/_rels/document.xml.rels"

    t_rels_bytes = try_read_xml(tzip, t_rels_path)
    if not t_rels_bytes:
        # minimal relationships root
        t_rels_root = ET.Element("{%s}Relationships" % PKG_REL_NS)
    else:
        t_rels_root = ET.fromstring(t_rels_bytes)

    s_rels_bytes = try_read_xml(szip, s_rels_path)
    if not s_rels_bytes:
        # nothing to merge
        return merged_doc_xml, {}, ET.tostring(t_rels_root, encoding="utf-8", xml_declaration=True)

    s_rels_root = ET.fromstring(s_rels_bytes)

    # Existing rIds in template
    existing_ids = set()
    for rel in t_rels_root.findall("{%s}Relationship" % PKG_REL_NS):
        rid = rel.get("Id")
        if rid:
            existing_ids.add(rid)

    # rIds actually referenced in merged document.xml
    used_in_doc = collect_used_rids(merged_doc_xml)

    # Build mapping old_rid -> new_rid for ones used in doc
    rid_map: dict[str, str] = {}
    extra_files: dict[str, bytes] = {}

    for s_rel in s_rels_root.findall("{%s}Relationship" % PKG_REL_NS):
        old_id = s_rel.get("Id")
        if not old_id or old_id not in used_in_doc:
            continue  # skip relationships not used in injected XML

        target = s_rel.get("Target", "")
        rel_type = s_rel.get("Type", "")
        target_mode = s_rel.get("TargetMode", None)  # External links won't have a part to copy

        # Assign a new unique rId in template
        new_id = next_rid(existing_ids)
        existing_ids.add(new_id)
        rid_map[old_id] = new_id

        # Clone the relationship into template rels
        new_rel = ET.Element("{%s}Relationship" % PKG_REL_NS)
        new_rel.set("Id", new_id)
        new_rel.set("Type", rel_type)
        new_rel.set("Target", target)
        if target_mode:
            new_rel.set("TargetMode", target_mode)
        t_rels_root.append(new_rel)

        # Copy the target part if it is an internal target like "media/image1.png"
        if not target_mode:  # internal part
            part_path = "word/" + target
            try:
                extra_files[part_path] = szip.read(part_path)
            except KeyError:
                pass  # ignore missing targets

    # ---- Rewrite rIds inside merged document.xml ----
    try:
        doc_root = ET.fromstring(merged_doc_xml)
        for el in doc_root.iter():
            # adjust r:id, r:embed, r:link if present
            for attr in list(el.attrib.keys()):
                if el.attrib[attr] in rid_map and (
                    attr.endswith("}id") or attr.endswith("}embed") or attr.endswith("}link")
                ):
                    el.set(attr, rid_map[el.attrib[attr]])
        merged_doc_xml = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)
    except Exception:
        for old, new in rid_map.items():
            merged_doc_xml = merged_doc_xml.replace(old.encode("utf-8"), new.encode("utf-8"))

    # ---- Serialize relationships with correct default namespace ----
    updated_rels_xml = ET.tostring(t_rels_root, encoding="utf-8", xml_declaration=True)
    updated_rels_xml = (
        updated_rels_xml
        .replace(b"ns0:", b"")
        .replace(
            b'xmlns:ns0="http://schemas.openxmlformats.org/package/2006/relationships"',
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
        )
    )

    return merged_doc_xml, extra_files, updated_rels_xml


def rebuild_docx(template_bytes: bytes, updated_document_xml: bytes, extra_files: dict[str, bytes], updated_rels_xml: bytes) -> io.BytesIO:
    """
    Writes a new DOCX:
      - all entries from template,
      - document.xml replaced,
      - document.xml.rels replaced (or created),
      - extra files from source (media etc.) added if they don't already exist.
    """
    in_mem = io.BytesIO(template_bytes)
    out_mem = io.BytesIO()

    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(out_mem, "w", zipfile.ZIP_DEFLATED) as zout:
        template_names = set(zin.namelist())

        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr("word/document.xml", updated_document_xml)
            elif item.filename == "word/_rels/document.xml.rels":
                # overwrite with our updated rels
                zout.writestr("word/_rels/document.xml.rels", updated_rels_xml)
            else:
                zout.writestr(item, zin.read(item.filename))

        # Add extra files that were not in template
        for path, content in extra_files.items():
            if path not in template_names:
                zout.writestr(path, content)

    out_mem.seek(0)
    return out_mem


def replace_placeholder_in_xml(template_xml: bytes, source_xml: bytes, placeholder="{{Permbajtja}}") -> bytes:
    """
    Replace the paragraph containing the placeholder with the source body contents (except sectPr).
    """
    try:
        ET.register_namespace("w", W_NS)
        ET.register_namespace("r", R_NS)

        t_root = ET.fromstring(template_xml)
        s_root = ET.fromstring(source_xml)

        t_body = t_root.find(".//w:body", NS)
        s_body = s_root.find(".//w:body", NS)
        if t_body is None or s_body is None:
            print("❌ No <w:body> found in template or source.")
            return template_xml

        source_elems = [deepcopy(el) for el in list(s_body) if not el.tag.endswith("sectPr")]

        found = False
        for p, merged_text in merge_split_runs(t_root):
            if placeholder in merged_text:
                found = True
                print("✅ Placeholder replaced successfully. Inserted", len(source_elems), "elements.")
                idx = list(t_body).index(p)
                t_body.remove(p)
                for el in source_elems:
                    t_body.insert(idx, el)
                    idx += 1
                return ET.tostring(t_root, encoding="utf-8", xml_declaration=True)

        if not found:
            print("⚠️ Placeholder not found in template. Check that '{{Permbajtja}}' exists in word/document.xml.")
        return template_xml

    except Exception as e:
        print("replace_placeholder_in_xml ERROR:", e)
        print(traceback.format_exc())
        return template_xml



# ------------------------------------------------------------
# Endpoints
# ------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    return jsonify({
        "service": "docx-injector",
        "ok": True,
        "endpoints": ["/inject-docx", "/debug-scan", "/debug-find-placeholder", "/debug-rels"]
    })

@app.route("/status", methods=["GET"])
def status():
    return jsonify({"service": "docx-injector", "ok": True})

@app.route("/inject-docx", methods=["POST"])
def inject_docx():
    """
    Multipart form-data:
      - template: DOCX file (required)
      - source: DOCX file (required)
      - placeholder: string (default {{Permbajtja}})
    Returns: merged DOCX
    """
    try:
        if "template" not in request.files or "source" not in request.files:
            return jsonify({"error": "Both 'template' and 'source' are required"}), 400

        template_bytes = request.files["template"].read()
        source_bytes = request.files["source"].read()
        placeholder = request.form.get("placeholder", "{{Permbajtja}}")

        t_xml, _ = get_document_xml(template_bytes)
        s_xml, _ = get_document_xml(source_bytes)

        injected_xml = replace_placeholder_in_xml(t_xml, s_xml, placeholder)

        # Merge rels/media for any rId used in injected_xml
        fixed_doc_xml, extra_files, updated_rels_xml = merge_relationships_and_media(
            template_bytes, source_bytes, injected_xml
        )

        # Build final docx
        out_doc = rebuild_docx(template_bytes, fixed_doc_xml, extra_files, updated_rels_xml)

        return send_file(
            out_doc,
            as_attachment=True,
            download_name="merged.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("inject_docx ERROR:", e)
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# ---------- Debug endpoints you already use ----------
@app.route("/debug-scan", methods=["POST"])
def debug_scan():
    try:
        if "file" not in request.files:
            return jsonify({"error": "Missing 'file'"}), 400
        doc = request.files["file"].read()
        xml, names = get_document_xml(doc)
        text = xml.decode("utf-8", errors="ignore")
        # merged-text scan
        root = ET.fromstring(xml)
        merged = " ".join("".join(t.text or "" for t in p.findall(".//w:t", NS)) for p in root.findall(".//w:p", NS))
        return jsonify({
            "contains_word_document": "word/document.xml" in names,
            "found_literal": "{{Permbajtja}}" in text,
            "found_merged_text": "{{Permbajtja}}" in merged,
            "xml_length": len(text)
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/debug-find-placeholder", methods=["POST"])
def debug_find_placeholder():
    try:
        if "template" not in request.files:
            return jsonify({"error": "template file required"}), 400
        t = request.files["template"].read()
        with zipfile.ZipFile(io.BytesIO(t)) as z:
            xml = z.read("word/document.xml")
        root = ET.fromstring(xml)
        out = []
        for p in root.findall(".//w:p", NS):
            merged = "".join(t.text or "" for t in p.findall(".//w:t", NS))
            if "Permbajtja" in merged or "{{" in merged or "}}" in merged:
                out.append({"merged_text": merged})
        return jsonify({"placeholder_paragraphs_found": len(out), "paragraphs": out[:3]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/debug-rels", methods=["POST"])
def debug_rels():
    try:
        if "file" not in request.files:
            return jsonify({"error": "Missing 'file'"}), 400
        b = request.files["file"].read()
        with zipfile.ZipFile(io.BytesIO(b)) as z:
            names = z.namelist()
            rels = {}
            for n in names:
                if "rels" in n:
                    rels[n] = z.read(n).decode("utf-8", errors="ignore")[:1000]
        return jsonify({"rels_files": list(rels.keys()), "sample": rels})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ------------------------------------------------------------
# Run
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)



