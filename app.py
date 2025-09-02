# app.py
from fastapi import FastAPI, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import backend  # import your existing backend.py

app = FastAPI()

# Allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/", response_class=HTMLResponse)
def home():
    return "<h2>PPT Generator Backend Running ✅</h2>"

@app.post("/generate_ppt/")
def generate_ppt(prompt: str = Form(...), theme: str = Form(...)): # Add this parameter
    hybrid_output = backend.generate_slides(prompt)
    try:
        json_part = hybrid_output[hybrid_output.index("{"):hybrid_output.rindex("}")+1]
        filename = "presentation.pptx"
        backend.create_ppt_from_json(json_part, filename, theme) # Add the theme variable here
        return FileResponse(
            filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=filename
        )
    except Exception as e:
        return {"error": str(e), "raw_output": hybrid_output}
