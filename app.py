from fastapi import FastAPI, UploadFile, File
import zipfile
import io
from lxml import etree

app = FastAPI()

@app.get("/")
def root():
    return {"status": "PR Writer DOCX parser running"}


@app.post("/extract-comments")
async def extract_comments(file: UploadFile = File(...)):

    content = await file.read()

    comments = {}
    results = []

    with zipfile.ZipFile(io.BytesIO(content)) as docx:

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # ---- COMMENTS ----
        if "word/comments.xml" in docx.namelist():

            comments_xml = docx.read("word/comments.xml")
            tree = etree.fromstring(comments_xml)

            for c in tree.findall(".//w:comment", ns):

                cid = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")

                texts = c.findall(".//w:t", ns)
                comment_text = " ".join([t.text for t in texts if t.text])

                comments[cid] = comment_text

        # ---- DOCUMENT ----
        document_xml = docx.read("word/document.xml")
        doc_tree = etree.fromstring(document_xml)

        text_ranges = {}

        for start in doc_tree.findall(".//w:commentRangeStart", ns):

            cid = start.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")

            collected = []
            node = start.getnext()

            while node is not None:

                if node.tag.endswith("commentRangeEnd"):
                    break

                texts = node.findall(".//w:t", ns)
                for t in texts:
                    if t.text:
                        collected.append(t.text)

                node = node.getnext()

            text_ranges[cid] = " ".join(collected)

        # ---- MERGE ----
        for cid, comment in comments.items():

            results.append({
                "comment_id": cid,
                "comment": comment,
                "text": text_ranges.get(cid, "")
            })

    return {"comments": results}
