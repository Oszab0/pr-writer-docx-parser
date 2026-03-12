from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def root():
    return {"status": "PR Writer DOCX parser running"}
