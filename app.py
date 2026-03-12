from fastapi import FastAPI, UploadFile, File
import zipfile
import io
from lxml import etree

app = FastAPI()

@app.get("/")
def root():
    return {"status":"PR Writer DOCX parser running"}


@app.post("/extract-comments")
async def extract_comments(file: UploadFile = File(...)):

    content = await file.read()

    comments_list = []

    with zipfile.ZipFile(io.BytesIO(content)) as docx:

        if "word/comments.xml" not in docx.namelist():
            return {"comments":[]}

        comments_xml = docx.read("word/comments.xml")

        tree = etree.fromstring(comments_xml)

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        for c in tree.findall(".//w:comment", ns):

            comment_id = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")

            texts = c.findall(".//w:t", ns)

            comment_text = " ".join([t.text for t in texts if t.text])

            comments_list.append({
                "comment_id": comment_id,
                "comment": comment_text
            })

    return {
        "comments": comments_list
    }
