from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List
import zipfile
import io
import json
from lxml import etree

app = FastAPI(title="PR Writer DOCX Parser")

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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
# Helpers
# ----------------------------

def get_text_from_element(element) -> str:
    texts = element.findall(".//w:t", NS)
    return "".join([t.text for t in texts if t.text]).strip()


def extract_paragraphs(doc_tree) -> list[dict]:
    paragraphs = []
    p_nodes = doc_tree.findall(".//w:p", NS)

    for idx, p in enumerate(p_nodes):
        text = get_text_from_element(p)
        if text:
            paragraphs.append({
                "paragraph_index": idx,
                "text": text,
                "element": p
            })

    return paragraphs


def extract_comment_texts(docx) -> dict:
    comments = {}

    if "word/comments.xml" not in docx.namelist():
        return comments

    comments_xml = docx.read("word/comments.xml")
    tree = etree.fromstring(comments_xml)

    for c in tree.findall(".//w:comment", NS):
        cid = c.get(f"{{{W_NS}}}id")
        comment_text = get_text_from_element(c)
        comments[cid] = {
            "comment_id": cid,
            "comment_text": comment_text
        }

    return comments


def extract_comment_targets(paragraphs, comments) -> tuple[dict, list]:
    mapped = {}
    unmapped = []

    for p in paragraphs:
        starts = p["element"].findall(".//w:commentRangeStart", NS)

        for start in starts:
            cid = start.get(f"{{{W_NS}}}id")

            if cid not in comments:
                continue

            if cid not in mapped:
                mapped[cid] = {
                    "comment_id": cid,
                    "comment_text": comments[cid]["comment_text"],
                    "target_text": p["text"],
                    "paragraph_index": p["paragraph_index"]
                }

    for cid, item in comments.items():
        if cid not in mapped:
            unmapped.append({
                "comment_id": cid,
                "comment_text": item["comment_text"],
                "reason": "target_text_not_found"
            })

    return mapped, unmapped


def classify_blocks(paragraphs: list[dict]) -> list[dict]:
    if not paragraphs:
        return []

    quote_indices = []
    for i, p in enumerate(paragraphs):
        text = p["text"]
        if "„" in text or '"' in text:
            quote_indices.append(i)

    last_idx = len(paragraphs) - 1
    first_quote_idx = quote_indices[0] if quote_indices else None

    blocks = []
    counters = {
        "title": 0,
        "lead": 0,
        "body_before_quote": 0,
        "quote_block": 0,
        "body_after_quote": 0,
        "closing": 0
    }

    for i, p in enumerate(paragraphs):
        if i == 0:
            block_type = "title"
        elif i == 1:
            block_type = "lead"
        elif i == last_idx:
            block_type = "closing"
        elif first_quote_idx is not None and i == first_quote_idx:
            block_type = "quote_block"
        elif first_quote_idx is not None and i < first_quote_idx:
            block_type = "body_before_quote"
        elif first_quote_idx is not None and i > first_quote_idx:
            block_type = "body_after_quote"
        else:
            block_type = "body_before_quote"

        counters[block_type] += 1

        blocks.append({
            "block_id": f"{block_type}_{counters[block_type]}",
            "block_type": block_type,
            "paragraph_index": p["paragraph_index"],
            "original_text": p["text"],
            "comments": [],
            "element": p["element"]
        })

    return blocks


def attach_comments_to_blocks(blocks: list[dict], mapped_comments: dict) -> list[dict]:
    by_paragraph_index = {b["paragraph_index"]: b for b in blocks}

    for _, comment in mapped_comments.items():
        p_idx = comment["paragraph_index"]
        if p_idx in by_paragraph_index:
            by_paragraph_index[p_idx]["comments"].append({
                "comment_id": comment["comment_id"],
                "comment_text": comment["comment_text"],
                "target_text": comment["target_text"]
            })

    return blocks


def write_text_to_paragraph(paragraph_element, new_text: str):
    for r in paragraph_element.findall("w:r", NS):
        paragraph_element.remove(r)

    run = etree.SubElement(paragraph_element, f"{{{W_NS}}}r")
    text_el = etree.SubElement(run, f"{{{W_NS}}}t")
    text_el.text = new_text


def apply_revisions_to_blocks(blocks: list[dict], revisions: list[dict]) -> list[dict]:
    blocks_by_id = {b["block_id"]: b for b in blocks}

    for rev in revisions:
        block_id = rev["block_id"]
        revised_text = rev["revised_text"]

        if block_id in blocks_by_id:
            blocks_by_id[block_id]["revised_text"] = revised_text

    return blocks


def rebuild_docx_xml(blocks: list[dict]):
    for block in blocks:
        if "revised_text" in block:
            write_text_to_paragraph(block["element"], block["revised_text"])


# ----------------------------
# Extract comments v2
# ----------------------------

@app.post("/extract-comments")
async def extract_comments(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="INVALID_FILE_TYPE")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Empty DOCX file")

    try:
        with zipfile.ZipFile(io.BytesIO(content)) as docx:
            if "word/document.xml" not in docx.namelist():
                raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")

            document_xml = docx.read("word/document.xml")
            doc_tree = etree.fromstring(document_xml)

            comments = extract_comment_texts(docx)

            if not comments:
                return {
                    "ok": False,
                    "warning": True,
                    "error_code": "COMMENTS_NOT_FOUND",
                    "message": "A dokumentumban nem található Word komment."
                }

            paragraphs = extract_paragraphs(doc_tree)
            mapped_comments, unmapped_comments = extract_comment_targets(paragraphs, comments)
            blocks = classify_blocks(paragraphs)
            blocks = attach_comments_to_blocks(blocks, mapped_comments)

            for block in blocks:
                block.pop("element", None)

            return {
                "ok": True,
                "status": "review_input",
                "comments_found": len(comments),
                "document_blocks": blocks,
                "unmapped_comments": unmapped_comments
            }

    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")
    except etree.XMLSyntaxError:
        raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"UNEXPECTED_ERROR: {str(e)}")


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
    block_id: str
    comment_id: str
    revised_text: str
    change_type: str
    review_comment: str


class RebuildRequest(BaseModel):
    revisions: List[RevisionItem]


# ----------------------------
# Review validation endpoint
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


# ----------------------------
# Rebuild document v1
# Replaces full paragraph text by block_id and returns rebuilt DOCX
# ----------------------------

@app.post("/rebuild-document")
async def rebuild_document(
    file: UploadFile = File(...),
    revisions_json: str = File(...)
):
    content = await file.read()

    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="INVALID_FILE_TYPE")

    if not content:
        raise HTTPException(status_code=400, detail="Empty DOCX file")

    if not revisions_json:
        raise HTTPException(status_code=400, detail="Missing revisions_json")

    try:
        parsed = json.loads(revisions_json)
        payload = RebuildRequest.model_validate(parsed)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"INVALID_REVISIONS_JSON: {str(e)}")

    try:
        input_zip_buffer = io.BytesIO(content)

        with zipfile.ZipFile(input_zip_buffer, "r") as input_zip:
            if "word/document.xml" not in input_zip.namelist():
                raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")

            document_xml = input_zip.read("word/document.xml")
            doc_tree = etree.fromstring(document_xml)

            paragraphs = extract_paragraphs(doc_tree)
            blocks = classify_blocks(paragraphs)
            blocks = apply_revisions_to_blocks(
                blocks,
                [r.model_dump() for r in payload.revisions]
            )
            rebuild_docx_xml(blocks)

            updated_document_xml = etree.tostring(
                doc_tree,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes"
            )

            output_buffer = io.BytesIO()
            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as output_zip:
                for item in input_zip.infolist():
                    if item.filename == "word/document.xml":
                        output_zip.writestr(item, updated_document_xml)
                    else:
                        output_zip.writestr(item, input_zip.read(item.filename))

        output_buffer.seek(0)

        return StreamingResponse(
            output_buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")
    except etree.XMLSyntaxError:
        raise HTTPException(status_code=400, detail="DOCX_PARSE_ERROR")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"REVIEW_REBUILD_ERROR: {str(e)}")
