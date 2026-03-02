from fastapi import FastAPI, HTTPException, Request, Body
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
import uvicorn
import uuid

try:
    import image
except ImportError:
    import sys, os
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import image


# ---------------------------------------------------------------------------
# Middleware: sanitize raw JSON body BEFORE FastAPI parses it
# ---------------------------------------------------------------------------
class SanitizeJsonBodyMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        content_type = request.headers.get("content-type", "")
        if "application/json" in content_type:
            raw = await request.body()
            try:
                sanitized = self._sanitize(raw.decode("utf-8")).encode("utf-8")
            except Exception:
                sanitized = raw

            async def receive():
                return {"type": "http.request", "body": sanitized, "more_body": False}

            request = Request(request.scope, receive)

        return await call_next(request)

    @staticmethod
    def _sanitize(text: str) -> str:
        result = []
        in_string = False
        escape_next = False
        for ch in text:
            if escape_next:
                result.append(ch)
                escape_next = False
                continue
            if ch == "\\" and in_string:
                result.append(ch)
                escape_next = True
                continue
            if ch == '"':
                in_string = not in_string
                result.append(ch)
                continue
            if in_string:
                if ch == "\n":
                    result.append("\\n")
                elif ch == "\r":
                    result.append("\\r")
                elif ch == "\t":
                    result.append("\\t")
                else:
                    result.append(ch)
            else:
                result.append(ch)
        return "".join(result)


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
app = FastAPI(
    title="AI PowerPoint Generator API",
    description="Generate professional PowerPoint presentations using AWS Bedrock (Claude) and Nova Canvas.",
    version="1.0.0"
)

app.add_middleware(SanitizeJsonBodyMiddleware)

# Allow browser-based frontends (e.g. React dev server) to call this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Shared core logic
# ---------------------------------------------------------------------------
def _run_generation(prompt: str):
    if not prompt.strip():
        raise HTTPException(status_code=400, detail="Prompt cannot be empty")
    print(f"Generating for: {prompt[:80]}...")
    pptx_data, slides = image.run(prompt.strip())
    filename = f"presentation_{uuid.uuid4().hex[:8]}.pptx"
    return Response(
        content=pptx_data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename={filename}",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
    )


# ---------------------------------------------------------------------------
# POST /generate-text  — plain text body (Swagger UI + curl + Postman + Python)
# ---------------------------------------------------------------------------
@app.post(
    "/generate-text",
    summary="Generate PPTX — paste plain text directly",
    tags=["Generate"],
    response_class=Response,
    responses={
        200: {
            "description": "Downloadable .pptx file",
            "content": {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": {}
            }
        }
    }
)
async def generate_pptx_text(
    prompt: str = Body(
        ...,
        media_type="text/plain",
        description="Paste your content as plain text — multiline, emojis, bullets all supported. No JSON escaping needed.",
        example=(
            "🏭 Product Manufacturing Lifecycle\n"
            "1️⃣ Raw Material Procurement\n"
            "* Identify material requirements\n"
            "* Select and evaluate suppliers\n"
            "Output: Approved raw materials\n\n"
            "2️⃣ Production Planning\n"
            "* Analyze demand forecasts\n"
            "* Create production schedule\n"
            "Output: Approved production schedule"
        )
    )
):
    """
    Paste your **plain text** directly — no JSON wrapping, no escaping, no `\\n` needed.

    Works from **Swagger UI**, **curl**, **Postman**, and **Python**:

    ```bash
    # curl
    curl -X POST http://localhost:8000/generate-text \\
      -H "Content-Type: text/plain" \\
      --data-binary @prompt.txt \\
      --output presentation.pptx
    ```

    ```python
    # Python
    import requests
    prompt = open("prompt.txt").read()
    r = requests.post("http://localhost:8000/generate-text",
                      data=prompt.encode("utf-8"),
                      headers={"Content-Type": "text/plain"})
    open("out.pptx", "wb").write(r.content)
    ```
    """
    try:
        return _run_generation(prompt)
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------
@app.get("/health", summary="Health Check", tags=["System"])
async def health_check():
    return {"status": "healthy", "service": "presentation-generator"}


if __name__ == "__main__":
    print("Starting AI PowerPoint Generator API on http://0.0.0.0:8000")
    uvicorn.run("api:app", host="localhost", port=8001)