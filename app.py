from fastapi import FastAPI, UploadFile, File, HTTPException
from pydantic import BaseModel
from typing import List
import zipfile
import io
from lxml import etree

app = FastAPI(title="PR Writer DOCX Parser")


# ----------------------------
# Health / root
# ----------------------------

@app.get("/")
def root():
    return {"status": "PR Writer DOCX parser running"}


@app.get("/health")
def health():
    return {"ok": True, "service": "pr-writer-docx-parser"}


# ----------------------------
# Existing DOCX comment extractor
# ----------------------------

@app.post("/extract-comments")
async def extract_comments(file: UploadFile = File(...)):
    content = await file.read()

    comments = {}
    results = []

    with zipfile.ZipFile(io.BytesIO(content)) as docx:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # COMMENTS
        if "word/comments.xml" in docx.namelist():
            comments_xml = docx.read("word/comments.xml")
            tree = etree.fromstring(comments_xml)

            for c in tree.findall(".//w:comment", ns):
                cid = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")

                texts = c.findall(".//w:t", ns)
                comment_text = " ".join([t.text for t in texts if t.text])

                comments[cid] = comment_text

        # DOCUMENT
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

        # MERGE
        for cid, comment in comments.items():
            results.append({
                "comment_id": cid,
                "comment": comment,
                "text": text_ranges.get(cid, "")
            })

    return {"comments": results}


# ----------------------------
# Review payload schema
# ----------------------------

class ReviewRequest(BaseModel):
    action: str
    project: str
    release_slug: str
    review_round: int
    status_in: str
    status_out: str
    source_file_name: str
    target_file_name: str


# ----------------------------
# AI Revision schema
# ----------------------------

class RevisionItem(BaseModel):
    comment_id: str
    original_text: str
    editor_comment: str
    revised_text: str
    change_type: str
    review_comment: str


class RebuildRequest(BaseModel):
    revisions: List[RevisionItem]


# ----------------------------
# Stub review endpoint for Make
# ----------------------------

@app.post("/review")
def review_document(payload: ReviewRequest):
    if payload.action != "review_document":
        raise HTTPException(status_code=400, detail="Invalid action")

    expected_input = f"{payload.release_slug}_c{payload.review_round}.docx"
    expected_output = f"{payload.release_slug}_r{payload.review_round}.docx"

    filename_valid = payload.source_file_name == expected_input
    target_valid = payload.target_file_name == expected_output

    return {
        "ok": True,
        "status": "review",
        "message": "Review endpoint reached successfully.",
        "release_slug": payload.release_slug,
        "review_round": payload.review_round,
        "status_in": payload.status_in,
        "status_out": payload.status_out,
        "source_file_name": payload.source_file_name,
        "target_file_name": payload.target_file_name,
        "expected_input_file_name": expected_input,
        "expected_output_file_name": expected_output,
        "filename_valid": filename_valid,
        "target_valid": target_valid
    }
