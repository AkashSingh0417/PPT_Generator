from fastapi import FastAPI, Form, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import backend

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
def generate_ppt(prompt: str = Form(...), ppt_template: UploadFile = File(None)):
    hybrid_output = backend.generate_slides(prompt)
    try:
        json_part = hybrid_output[hybrid_output.index("{"):hybrid_output.rindex("}")+1]
        filename = "presentation.pptx"

        template_path = None
        if ppt_template:
            template_path = "uploaded_template.pptx"
            with open(template_path, "wb") as f:
                f.write(ppt_template.file.read())

        backend.create_ppt_from_json(json_part, filename, template_path=template_path)

        return FileResponse(
            filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=filename
        )
    except Exception as e:
        return {"error": str(e), "raw_output": hybrid_output}
