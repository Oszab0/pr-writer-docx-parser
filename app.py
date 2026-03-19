import os
import tempfile
import zipfile
from flask import Flask, request, send_file, jsonify
from lxml import etree
import openai

app = Flask(__name__)

openai.api_key = os.environ.get("OPENAI_API_KEY")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


# -----------------------------
# DOCX extraction
# -----------------------------

def extract_docx_xml(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read("word/document.xml")

    return etree.fromstring(xml_content)


# -----------------------------
# Collect paragraphs
# -----------------------------

def collect_paragraphs(xml_root):

    paragraphs = []

    for p in xml_root.findall(".//w:p", NS):

        texts = p.findall(".//w:t", NS)

        full_text = "".join([t.text for t in texts if t.text])

        if full_text.strip():

            paragraphs.append({
                "element": p,
                "text": full_text
            })

    return paragraphs


# -----------------------------
# GPT rewrite
# -----------------------------

def rewrite_with_gpt(text):

    prompt = f"""
Rewrite this press release paragraph according to PR standards.

Keep the meaning but improve clarity, flow and media style.

Text:
{text}
"""

    response = openai.ChatCompletion.create(
        model="gpt-5",
        messages=[
            {"role": "system", "content": "You are a professional PR editor."},
            {"role": "user", "content": prompt}
        ]
    )

    return response.choices[0].message.content.strip()


# -----------------------------
# Write text into paragraph
# -----------------------------

def write_text_to_paragraph(paragraph_element, new_text):

    for r in paragraph_element.findall("w:r", NS):
        paragraph_element.remove(r)

    run = etree.SubElement(
        paragraph_element,
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"
    )

    text_el = etree.SubElement(
        run,
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    )

    text_el.text = new_text


# -----------------------------
# Apply revisions
# -----------------------------

def rebuild_docx_xml(paragraphs):

    for block in paragraphs:

        revised = rewrite_with_gpt(block["text"])

        write_text_to_paragraph(block["element"], revised)


# -----------------------------
# Save DOCX
# -----------------------------

def rebuild_docx(original_docx, xml_root, output_path):

    with tempfile.TemporaryDirectory() as tmp:

        with zipfile.ZipFile(original_docx) as z:
            z.extractall(tmp)

        xml_path = os.path.join(tmp, "word/document.xml")

        with open(xml_path, "wb") as f:
            f.write(etree.tostring(xml_root))

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as new_docx:

            for folder, _, files in os.walk(tmp):

                for file in files:

                    full_path = os.path.join(folder, file)

                    arcname = os.path.relpath(full_path, tmp)

                    new_docx.write(full_path, arcname)


# -----------------------------
# API endpoint
# -----------------------------

@app.route("/review-document", methods=["POST"])

def review_document():

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded = request.files["file"]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_input:

        uploaded.save(temp_input.name)

        xml_root = extract_docx_xml(temp_input.name)

        paragraphs = collect_paragraphs(xml_root)

        rebuild_docx_xml(paragraphs)

        output_path = temp_input.name.replace(".docx", "_reviewed.docx")

        rebuild_docx(temp_input.name, xml_root, output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name="reviewed.docx"
        )


# -----------------------------
# Health check
# -----------------------------

@app.route("/health")

def health():
    return {"status": "ok"}


# -----------------------------
# Run
# -----------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
