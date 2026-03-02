import json
import re
import uuid
import subprocess
import os
import sys
import base64
import textwrap
import boto3
import concurrent.futures
import time
from io import BytesIO
from dataclasses import dataclass, field, asdict
from typing import Optional
from enum import Enum


try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("⚠️  Pillow not installed. Run: pip install Pillow")


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
AWS_REGION    = os.getenv("AWS_REGION",        "us-east-1")
BEDROCK_MODEL = os.getenv("BEDROCK_MODEL_ID",  "us.anthropic.claude-3-7-sonnet-20250219-v1:0")
NOVA_CANVAS   = "amazon.nova-canvas-v1:0"
S3_BUCKET     = "data-to-infographic"
OUTPUT_DIR    = os.getenv("OUTPUT_DIR",         "./outputs")
ICONS_DIR     = os.path.join(OUTPUT_DIR, "icons")
PROJECT_DIR   = os.path.dirname(os.path.abspath(__file__))

from botocore.config import Config

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ICONS_DIR,  exist_ok=True)

# Increase timeout for large content generation
bedrock_config = Config(connect_timeout=120, read_timeout=120, retries={'max_attempts': 3})
bedrock = boto3.client("bedrock-runtime", region_name=AWS_REGION, config=bedrock_config)
s3      = boto3.client("s3",              region_name=AWS_REGION)

def sync_to_s3(local_path: str, bucket: str = S3_BUCKET) -> Optional[str]:
    """Uploads a local file to S3 and returns the public URL or S3 path."""
    if not os.path.exists(local_path):
        return None
    
    # Use relative path from project root as S3 key
    key = os.path.relpath(local_path, start=PROJECT_DIR).replace("\\", "/")
    # Remove leading "./" or similar if any
    if key.startswith("./"): key = key[2:]
    
    try:
        content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation" if local_path.endswith(".pptx") else "image/png"
        s3.upload_file(local_path, bucket, key, ExtraArgs={"ContentType": content_type})
        print(f"    ☁️  Synced to S3: s3://{bucket}/{key}")
        return f"https://{bucket}.s3.{AWS_REGION}.amazonaws.com/{key}"
    except Exception as e:
        print(f"    ⚠️  S3 Sync Failed for {key}: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# ENUMS
# ────────────────────────────────────────────────────────────────────────────
class ContentType(str, Enum):
    TITLE        = "title"
    BULLETS      = "bullets"
    TWO_COLUMN   = "two_column"
    STAT_CALLOUT = "stat_callout"
    TIMELINE     = "timeline"
    TABLE        = "table"
    QUOTE        = "quote"
    DIAGRAM      = "diagram"
    THANK_YOU    = "thank_you"


class Theme(str, Enum):
    MIDNIGHT_EXECUTIVE = "midnight_executive"
    CORAL_ENERGY       = "coral_energy"
    OCEAN_GRADIENT     = "ocean_gradient"
    FOREST_MOSS        = "forest_moss"
    CHARCOAL_MINIMAL   = "charcoal_minimal"
    WARM_TERRACOTTA    = "warm_terracotta"


THEME_COLORS = {
    Theme.MIDNIGHT_EXECUTIVE: {"primary": "1E2761", "secondary": "CADCFC", "accent": "FFFFFF"},
    Theme.CORAL_ENERGY:       {"primary": "F96167", "secondary": "F9E795", "accent": "2F3C7E"},
    Theme.OCEAN_GRADIENT:     {"primary": "065A82", "secondary": "1C7293", "accent": "21295C"},
    Theme.FOREST_MOSS:        {"primary": "2C5F2D", "secondary": "97BC62", "accent": "F5F5F5"},
    Theme.CHARCOAL_MINIMAL:   {"primary": "36454F", "secondary": "F2F2F2", "accent": "212121"},
    Theme.WARM_TERRACOTTA:    {"primary": "B85042", "secondary": "E7E8D1", "accent": "A7BEAE"},
}


# ─────────────────────────────────────────────────────────────────────────────
# DATA MODELS
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class SlideOutlineItem:
    slide_number: int
    title: str
    content_type: ContentType
    description: str


@dataclass
class FlowStep:
    step_number: int
    label: str          # Short title (≤ 4 words)
    description: str    # One sentence detail shown on diagram
    icon_prompt: str    # Nova Canvas image-gen prompt
    icon_path: str = "" # Filled after generation


@dataclass
class SlideData:
    slide_number: int
    content_type: ContentType
    content: dict
    layout_override: Optional[str] = None


# ─────────────────────────────────────────────────────────────────────────────
# BEDROCK — CLAUDE
# ─────────────────────────────────────────────────────────────────────────────
def call_bedrock(system: str, user: str, max_tokens: int = 4096) -> str:
    body = {
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": max_tokens,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    try:
        response = bedrock.invoke_model(
            modelId=BEDROCK_MODEL,
            contentType="application/json",
            accept="application/json",
            body=json.dumps(body),
        )
    except Exception as e:
        msg = str(e)
        if "on-demand throughput isn't supported" in msg or "inference profile" in msg.lower():
            raise RuntimeError(
                f"\n  ❌  Invalid Model ID: '{BEDROCK_MODEL}'\n"
                f"  Fix: use a cross-region inference profile, e.g.:\n"
                f"    • us.anthropic.claude-3-7-sonnet-20250219-v1:0\n"
                f"    • us.anthropic.claude-3-5-sonnet-20241022-v2:0\n"
                f"  Docs: https://docs.aws.amazon.com/bedrock/latest/userguide/inference-profiles.html"
            ) from e
        raise RuntimeError(f"\n  ❌  Bedrock call failed: {msg}") from e

    return json.loads(response["body"].read())["content"][0]["text"]


def extract_json(text: str):
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r"```(?:json)?\s*([\s\S]+?)```", text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass
    for opener, closer in [("[", "]"), ("{", "}")]:
        start = text.find(opener)
        if start == -1:
            continue
        end = text.rfind(closer)
        if end == -1 or end <= start:
            continue
        candidate = text[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not extract JSON from response:\n{text[:600]}")


def _ensure_list(data) -> list:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list) and len(v) > 0:
                return v
        if any(k in data for k in ("slide_number", "step_number", "title")):
            return [data]
    raise ValueError(f"Expected a JSON array, got: {type(data).__name__} — {str(data)[:300]}")


# ─────────────────────────────────────────────────────────────────────────────
# NOVA CANVAS — ICON GENERATION
# ─────────────────────────────────────────────────────────────────────────────
_NOVA_CANVAS_SIZES = [
    (1024, 1024), (768, 768), (512, 512),
    (1280, 720),  (1152, 896), (896, 1152),
    (1024, 576),  (576, 1024),
]

def _nearest_nova_size(size: int) -> tuple[int, int]:
    squares = [(w, h) for w, h in _NOVA_CANVAS_SIZES if w == h]
    return min(squares, key=lambda s: abs(s[0] - size))


def generate_icon_image(prompt: str, size: int = 512) -> "Image.Image | None":
    if not PIL_AVAILABLE:
        return None

    w, h = _nearest_nova_size(size)

    full_prompt = (
        f"Flat style business icon: {prompt}. "
        "White background, single centered icon, bold clean shapes, "
        "vivid solid colors, no text, no shadows, minimalist professional style."
    )

    body = {
        "taskType": "TEXT_IMAGE",
        "textToImageParams": {
            "text": full_prompt,
            "negativeText": "text, letters, words, watermark, complex background, photo, realistic"
        },
        "imageGenerationConfig": {
            "numberOfImages": 1,
            "width":    w,
            "height":   h,
            "quality":  "standard",
            "cfgScale": 8.0,
            "seed":     42,
        },
    }

    try:
        resp = bedrock.invoke_model(
            modelId    = NOVA_CANVAS,
            body       = json.dumps(body),
            accept     = "application/json",
            contentType= "application/json",
        )
        result    = json.loads(resp["body"].read())

        if result.get("error"):
            print(f"    ⚠️  Nova Canvas error: {result['error']}")
            return None

        img_bytes = base64.b64decode(result["images"][0])
        img       = Image.open(BytesIO(img_bytes)).convert("RGBA")

        if img.size != (size, size):
            img = img.resize((size, size), Image.LANCZOS)

        return img

    except Exception as e:
        err = str(e)
        if "ValidationException" in err:
            print(f"    ⚠️  Nova Canvas ValidationException for '{prompt[:45]}':")
            print(f"         {err}")
        elif "AccessDeniedException" in err:
            print(f"    ⚠️  Nova Canvas Access Denied — enable model access in AWS Console:")
            print(f"         https://console.aws.amazon.com/bedrock/home#/modelaccess")
        else:
            print(f"    ⚠️  Nova Canvas failed for '{prompt[:45]}': {err}")
        return None


def _draw_fallback_icon(path: str, step_number: int) -> str:
    SIZE   = 256
    COLORS = [
        (30, 100, 200),
        (20, 150,  80),
        (200,  80,  20),
        (120,  20, 160),
        (20,  150, 150),
        (180, 140,  20),
    ]
    c   = COLORS[(step_number - 1) % len(COLORS)]
    img = Image.new("RGBA", (SIZE, SIZE), (255, 255, 255, 255))
    d   = ImageDraw.Draw(img)
    cx, cy = SIZE // 2, SIZE // 2
    m = SIZE // 5

    idx = (step_number - 1) % 6
    if idx == 0:
        d.rounded_rectangle([cx-m*2, cy-m*2, cx+m*2, cy+m*2], radius=12, fill=c)
        for offset in [-m//2, m//4, m]:
            d.rectangle([cx-m+6, cy+offset-4, cx+m-6, cy+offset+4], fill=(255,255,255))
    elif idx == 1:
        pts = [cx, cy-int(m*2.2), cx+int(m*1.8), cy-m, cx+int(m*1.8), cy+m//2,
               cx, cy+int(m*2.2), cx-int(m*1.8), cy+m//2, cx-int(m*1.8), cy-m]
        d.polygon(pts, fill=c)
        d.line([cx-m, cy+m//4, cx-m//4, cy+m], fill=(255,255,255), width=10)
        d.line([cx-m//4, cy+m, cx+m, cy-m//2], fill=(255,255,255), width=10)
    elif idx == 2:
        bars = [(-int(m*1.5), m//2), (-m//4, -m//2), (int(m*1.1), -m)]
        for bx, top in bars:
            d.rectangle([cx+bx, cy+top, cx+bx+m-4, cy+m], fill=c)
        d.line([cx-int(m*1.7), cy+m+6, cx+int(m*1.7), cy+m+6], fill=(180,180,180), width=5)
    elif idx == 3:
        d.rounded_rectangle([cx-int(m*1.6), cy, cx+int(m*1.6), cy+int(m*2.2)], radius=12, fill=c)
        d.arc([cx-m, cy-int(m*1.5), cx+m, cy+m//2], start=0, end=180, fill=c, width=14)
        d.ellipse([cx-m//3, cy+m//2, cx+m//3, cy+int(m*1.4)], fill=(255,255,255))
    elif idx == 4:
        d.ellipse([cx-int(m*1.8), cy-int(m*1.8), cx+int(m*1.8), cy+int(m*1.8)], fill=c)
        d.line([cx, cy-m, cx, cy+m], fill=(255,255,255), width=10)
        d.arc([cx-m+6, cy-m, cx+m-6, cy-m//4], start=180, end=0, fill=(255,255,255), width=8)
        d.arc([cx-m+6, cy-m//4, cx+m-6, cy+m//2], start=0, end=180, fill=(255,255,255), width=8)
    else:
        pts = [cx-int(m*2), cy-m//2, cx+m//2, cy-m//2, cx+m//2, cy-m,
               cx+int(m*2), cy, cx+m//2, cy+m, cx+m//2, cy+m//2, cx-int(m*2), cy+m//2]
        d.polygon(pts, fill=c)

    img.save(path, "PNG")
    return path


def generate_all_icons(steps: list[FlowStep]) -> list[FlowStep]:
    total = len(steps)
    for step in steps:
        fname = f"icon_step{step.step_number:02d}_{uuid.uuid4().hex[:6]}.png"
        path  = os.path.join(ICONS_DIR, fname)

        print(f"    🎨 [{step.step_number}/{total}] Icon: {step.label}")
        icon_img = generate_icon_image(step.icon_prompt, size=256)

        if icon_img:
            icon_img.save(path, "PNG")
            sync_to_s3(path)
            step.icon_path = path
        else:
            step.icon_path = _draw_fallback_icon(path, step.step_number)
            sync_to_s3(path)

    return steps


def regenerate_step_icon(icon_prompt: str, step_number: int, domain: str = "Corporate") -> str:
    fname = f"icon_step{step_number:02d}_{uuid.uuid4().hex[:6]}.png"
    path = os.path.join(ICONS_DIR, fname)
    prompt = build_nova_prompt(icon_prompt, domain, is_icon=True)
    img = call_nova(prompt, is_icon=True)
    if img:
        img.save(path, "PNG")
        sync_to_s3(path)
        return path
    _draw_fallback_icon(path, step_number)
    sync_to_s3(path)
    return path


def reassemble_diagram(steps_data: list[dict], title: str, user_query: str = "") -> dict:
    steps = [
        FlowStep(
            step_number=d["step_number"],
            label=d["label"],
            description=d["description"],
            icon_prompt=d.get("icon_prompt", ""),
            icon_path=d.get("icon_path", ""),
        )
        for d in steps_data
    ]
    palette   = _pick_diagram_palette(user_query or title)
    uuid_str  = uuid.uuid4().hex[:8]
    path_horiz = os.path.join(OUTPUT_DIR, f"flow_horiz_{uuid_str}.png")
    path_grid  = os.path.join(OUTPUT_DIR, f"flow_grid_{uuid_str}.png")
    path_vert  = os.path.join(OUTPUT_DIR, f"flow_vert_{uuid_str}.png")

    assemble_flow_diagram_image(steps, title, path_horiz, palette=palette)
    assemble_flow_diagram_grid(steps, title, path_grid, palette=palette)
    assemble_flow_diagram_vertical(steps, title, path_vert, palette=palette)

    for p in [path_horiz, path_grid, path_vert]:
        sync_to_s3(p)

    return {
        "diagram":          path_horiz.replace("\\", "/"),
        "diagram_grid":     path_grid.replace("\\", "/"),
        "diagram_vertical": path_vert.replace("\\", "/"),
    }


# ─────────────────────────────────────────────────────────────────────────────
# FLOW DIAGRAM — PIL ASSEMBLY (used only for Streamlit preview images)
# ─────────────────────────────────────────────────────────────────────────────

_DIAGRAM_PALETTES = {
    "loan|lending|finance|bank|invest|fund|credit": {
        "bg": (13, 17, 39), "card": (30, 39, 97), "card_border": (202, 220, 252),
        "circle_fill": (255, 255, 255), "circle_outline": (202, 220, 252),
        "badge_bg": (202, 220, 252), "badge_text": (30, 39, 97),
        "label_text": (255, 255, 255), "desc_text": (170, 190, 230),
        "arrow": (202, 220, 252), "title_bg": (20, 28, 75), "title_text": (255, 255, 255),
    },
    "health|medical|hospital|patient|care": {
        "bg": (2, 14, 26), "card": (6, 90, 130), "card_border": (156, 205, 207),
        "circle_fill": (255, 255, 255), "circle_outline": (156, 205, 207),
        "badge_bg": (156, 205, 207), "badge_text": (2, 14, 26),
        "label_text": (255, 255, 255), "desc_text": (156, 205, 207),
        "arrow": (156, 205, 207), "title_bg": (4, 60, 90), "title_text": (255, 255, 255),
    },
    "startup|product|launch|growth|innovation": {
        "bg": (26, 10, 11), "card": (180, 40, 50), "card_border": (249, 231, 149),
        "circle_fill": (255, 255, 255), "circle_outline": (249, 231, 149),
        "badge_bg": (249, 231, 149), "badge_text": (100, 10, 15),
        "label_text": (255, 255, 255), "desc_text": (249, 231, 149),
        "arrow": (249, 231, 149), "title_bg": (120, 25, 30), "title_text": (255, 255, 255),
    },
    "nature|green|forest|eco|environment|climate": {
        "bg": (10, 21, 9), "card": (44, 95, 45), "card_border": (151, 188, 98),
        "circle_fill": (255, 255, 255), "circle_outline": (151, 188, 98),
        "badge_bg": (151, 188, 98), "badge_text": (10, 30, 10),
        "label_text": (255, 255, 255), "desc_text": (151, 188, 98),
        "arrow": (151, 188, 98), "title_bg": (25, 60, 26), "title_text": (255, 255, 255),
    },
}

_DEFAULT_PALETTE = {
    "bg": (13, 17, 39), "card": (30, 39, 97), "card_border": (202, 220, 252),
    "circle_fill": (255, 255, 255),
    "circle_outline": (202, 220, 252),
    "badge_bg": (202, 220, 252), "badge_text": (30, 39, 97),
    "label_text": (255, 255, 255), "desc_text": (170, 190, 230),
    "arrow": (202, 220, 252), "title_bg": (20, 28, 75), "title_text": (255, 255, 255),
}


def _pick_diagram_palette(query: str) -> dict:
    lower = query.lower()
    for keywords, palette in _DIAGRAM_PALETTES.items():
        if any(kw in lower for kw in keywords.split("|")):
            return palette
    return _DEFAULT_PALETTE


def _load_font(size: int, bold: bool = False) -> "ImageFont.FreeTypeFont":
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"    if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _wrap_text(text: str, max_chars: int) -> list[str]:
    return textwrap.wrap(text, width=max_chars) or [""]


def _circle_crop(img: "Image.Image", size: int) -> "Image.Image":
    img  = img.resize((size, size), Image.LANCZOS).convert("RGBA")
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size, size], fill=255)
    out  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(img, (0, 0), mask)
    return out


def assemble_flow_diagram_image(
    steps: list[FlowStep],
    title: str,
    output_path: str,
    palette: dict | None = None,
) -> str:
    """Assemble a preview PNG for Streamlit display (not used in PPTX output)."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow is required: pip install Pillow")

    if palette is None:
        palette = _DEFAULT_PALETTE

    W         = 2400
    TITLE_H   = 110
    PAD_X     = 80
    PAD_TOP   = 55
    ICON_SZ   = 220
    BADGE_R   = 26
    ARROW_W   = 58
    LABEL_H   = 62
    DESC_H    = 88
    PAD_BOT   = 48

    n      = len(steps)
    step_w = (W - 2 * PAD_X - (n - 1) * ARROW_W) // n
    H      = TITLE_H + PAD_TOP + ICON_SZ + 14 + BADGE_R * 2 + 12 + LABEL_H + DESC_H + PAD_BOT

    canvas = Image.new("RGB", (W, H), palette["bg"])
    draw   = ImageDraw.Draw(canvas)

    draw.rectangle([0, 0, W, TITLE_H], fill=palette["title_bg"])
    draw.rectangle([0, TITLE_H - 4, W, TITLE_H], fill=palette["card_border"])

    title_font = _load_font(50, bold=True)
    tb         = draw.textbbox((0, 0), title, font=title_font)
    draw.text(
        ((W - (tb[2] - tb[0])) // 2, (TITLE_H - (tb[3] - tb[1])) // 2),
        title, fill=palette["title_text"], font=title_font
    )

    line_y   = TITLE_H + PAD_TOP + ICON_SZ // 2
    first_cx = PAD_X + step_w // 2
    last_cx  = PAD_X + (n - 1) * (step_w + ARROW_W) + step_w // 2

    conn_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(conn_layer).rectangle(
        [first_cx, line_y - 3, last_cx, line_y + 3],
        fill=(*palette["card_border"][:3], 60)
    )
    canvas = canvas.convert("RGBA")
    canvas = Image.alpha_composite(canvas, conn_layer)
    canvas = canvas.convert("RGB")
    draw   = ImageDraw.Draw(canvas)

    label_font = _load_font(28, bold=True)
    desc_font  = _load_font(22, bold=False)
    badge_font = _load_font(26, bold=True)

    for i, step in enumerate(steps):
        x0 = PAD_X + i * (step_w + ARROW_W)
        cx = x0 + step_w // 2

        glow_r   = ICON_SZ // 2 + 14
        circ_top = TITLE_H + PAD_TOP

        glow_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(glow_layer).ellipse(
            [cx - glow_r, circ_top - 14, cx + glow_r, circ_top + ICON_SZ + 14],
            fill=(*palette["card_border"][:3], 30)
        )
        canvas = canvas.convert("RGBA")
        canvas = Image.alpha_composite(canvas, glow_layer)
        canvas = canvas.convert("RGB")
        draw   = ImageDraw.Draw(canvas)

        circ_left    = cx - ICON_SZ // 2
        circle_fill  = palette.get("circle_fill",   (255, 255, 255))
        circle_edge  = palette.get("circle_outline", palette["card_border"])
        draw.ellipse(
            [circ_left, circ_top, circ_left + ICON_SZ, circ_top + ICON_SZ],
            fill=circle_fill,
            outline=circle_edge, width=5
        )

        if step.icon_path and os.path.exists(step.icon_path):
            try:
                raw_icon = Image.open(step.icon_path).convert("RGB")
                inner_sz = ICON_SZ - 50
                raw_icon = raw_icon.resize((inner_sz, inner_sz), Image.LANCZOS)
                ix = circ_left + (ICON_SZ - inner_sz) // 2
                iy = circ_top  + (ICON_SZ - inner_sz) // 2
                canvas.paste(raw_icon, (ix, iy))
                draw = ImageDraw.Draw(canvas)
                draw.ellipse(
                    [circ_left, circ_top, circ_left + ICON_SZ, circ_top + ICON_SZ],
                    outline=circle_edge, width=5
                )
            except Exception as ex:
                print(f"    ⚠️  Icon paste error step {step.step_number}: {ex}")

        badge_y  = circ_top + ICON_SZ + 14
        draw.ellipse(
            [cx - BADGE_R, badge_y, cx + BADGE_R, badge_y + BADGE_R * 2],
            fill=palette["badge_bg"],
            outline=palette["card"], width=2
        )
        num_text = str(step.step_number)
        nb       = draw.textbbox((0, 0), num_text, font=badge_font)
        draw.text(
            (cx - (nb[2] - nb[0]) // 2, badge_y + BADGE_R - (nb[3] - nb[1]) // 2),
            num_text, fill=palette["badge_text"], font=badge_font
        )

        label_y    = badge_y + BADGE_R * 2 + 12
        label_text = step.label.upper()
        label_lines = _wrap_text(label_text, max_chars=14)
        ly = label_y
        for line in label_lines[:2]:
            lb = draw.textbbox((0, 0), line, font=label_font)
            draw.text(
                (cx - (lb[2] - lb[0]) // 2, ly),
                line, fill=palette["label_text"], font=label_font
            )
            ly += lb[3] - lb[1] + 4

        desc_y     = label_y + LABEL_H + 4
        desc_lines = _wrap_text(step.description, max_chars=22)
        dy = desc_y
        for line in desc_lines[:4]:
            db = draw.textbbox((0, 0), line, font=desc_font)
            draw.text(
                (cx - (db[2] - db[0]) // 2, dy),
                line, fill=palette["desc_text"], font=desc_font
            )
            dy += db[3] - db[1] + 5

        if i < n - 1:
            ax  = x0 + step_w + 4
            ay  = TITLE_H + PAD_TOP + ICON_SZ // 2
            draw.rectangle([ax, ay - 4, ax + ARROW_W - 16, ay + 4], fill=palette["arrow"])
            tip = ax + ARROW_W - 8
            draw.polygon(
                [(ax + ARROW_W - 18, ay - 14), (tip, ay), (ax + ARROW_W - 18, ay + 14)],
                fill=palette["arrow"]
            )

    canvas.save(output_path, "PNG", dpi=(240, 240))
    print(f"    ✅  Flow diagram preview image saved → {output_path}")
    return output_path

def assemble_flow_diagram_grid(steps: list[FlowStep], title: str, output_path: str, palette: dict | None = None) -> str:
    if not PIL_AVAILABLE: raise RuntimeError("Pillow is required")
    if palette is None: palette = _DEFAULT_PALETTE

    W = 2400
    TITLE_H = 110
    PAD_TOP = 80
    PAD_BOT = 120
    ICON_SZ = 200
    cols = min(3, len(steps))
    rows = (len(steps) + cols - 1) // cols
    CELL_W = W // cols
    CELL_H = 480
    H = TITLE_H + PAD_TOP + rows * CELL_H + PAD_BOT

    canvas = Image.new("RGB", (W, H), palette["bg"])
    draw = ImageDraw.Draw(canvas)

    draw.rectangle([0, 0, W, TITLE_H], fill=palette["title_bg"])
    draw.rectangle([0, TITLE_H - 4, W, TITLE_H], fill=palette["card_border"])
    title_font = _load_font(50, bold=True)
    tb = draw.textbbox((0, 0), title, font=title_font)
    draw.text(((W - (tb[2] - tb[0])) // 2, (TITLE_H - (tb[3] - tb[1])) // 2), title, fill=palette["title_text"], font=title_font)

    label_font = _load_font(30, bold=True)
    desc_font  = _load_font(24, bold=False)
    badge_font = _load_font(26, bold=True)

    for i, step in enumerate(steps):
        r = i // cols
        c = i % cols
        cx = c * CELL_W + CELL_W // 2
        cy = TITLE_H + PAD_TOP + r * CELL_H + ICON_SZ // 2

        draw.ellipse([cx - ICON_SZ//2, cy - ICON_SZ//2, cx + ICON_SZ//2, cy + ICON_SZ//2], fill=palette.get("circle_fill", (255,255,255)), outline=palette.get("circle_outline", palette["card_border"]), width=5)

        if step.icon_path and os.path.exists(step.icon_path):
            try:
                raw_icon = Image.open(step.icon_path).convert("RGB")
                raw_icon = raw_icon.resize((ICON_SZ - 40, ICON_SZ - 40), Image.LANCZOS)
                canvas.paste(raw_icon, (cx - (ICON_SZ - 40)//2, cy - (ICON_SZ - 40)//2))
                draw = ImageDraw.Draw(canvas)
            except: pass

        badge_r = 28
        draw.ellipse([cx - badge_r, cy - ICON_SZ//2 - 14, cx + badge_r, cy - ICON_SZ//2 - 14 + badge_r*2], fill=palette["badge_bg"], outline=palette["card"], width=2)
        nd = draw.textbbox((0, 0), str(step.step_number), font=badge_font)
        draw.text((cx - (nd[2]-nd[0])//2, cy - ICON_SZ//2 - 14 + badge_r - (nd[3]-nd[1])//2), str(step.step_number), fill=palette["badge_text"], font=badge_font)

        ty = cy + ICON_SZ//2 + 30
        lb = draw.textbbox((0, 0), step.label.upper(), font=label_font)
        draw.text((cx - (lb[2]-lb[0])//2, ty), step.label.upper(), fill=palette["label_text"], font=label_font)

        desc_lines = _wrap_text(step.description, max_chars=28)[:3]
        dy = ty + 45
        for line in desc_lines:
            db = draw.textbbox((0, 0), line, font=desc_font)
            draw.text((cx - (db[2]-db[0])//2, dy), line, fill=palette["desc_text"], font=desc_font)
            dy += db[3]-db[1] + 5

    canvas.save(output_path, "PNG", dpi=(240, 240))
    return output_path

def assemble_flow_diagram_vertical(steps: list[FlowStep], title: str, output_path: str, palette: dict | None = None) -> str:
    if not PIL_AVAILABLE: raise RuntimeError()
    if palette is None: palette = _DEFAULT_PALETTE

    W = 2400
    TITLE_H = 110
    PAD_TOP = 80
    PAD_BOT = 100
    ROW_H = 280
    ICON_SZ = 180

    H = TITLE_H + PAD_TOP + PAD_BOT + len(steps) * ROW_H

    canvas = Image.new("RGB", (W, H), palette["bg"])
    draw = ImageDraw.Draw(canvas)

    draw.rectangle([0, 0, W, TITLE_H], fill=palette["title_bg"])
    draw.rectangle([0, TITLE_H - 4, W, TITLE_H], fill=palette["card_border"])
    title_font = _load_font(50, bold=True)
    tb = draw.textbbox((0, 0), title, font=title_font)
    draw.text(((W - (tb[2] - tb[0])) // 2, (TITLE_H - (tb[3] - tb[1])) // 2), title, fill=palette["title_text"], font=title_font)

    line_x = 400
    draw.rectangle([line_x - 3, TITLE_H + PAD_TOP, line_x + 3, H - PAD_BOT - ROW_H // 2], fill=(*palette["card_border"][:3], 150))

    label_font = _load_font(36, bold=True)
    desc_font  = _load_font(28, bold=False)
    badge_font = _load_font(28, bold=True)

    for i, step in enumerate(steps):
        circ_top = TITLE_H + PAD_TOP + i * ROW_H
        cx = line_x
        cy = circ_top + ICON_SZ // 2

        draw.ellipse([cx - ICON_SZ//2, circ_top, cx + ICON_SZ//2, circ_top + ICON_SZ], fill=palette.get("circle_fill", (255,255,255)), outline=palette.get("circle_outline", palette["card_border"]), width=5)

        if step.icon_path and os.path.exists(step.icon_path):
            try:
                raw_icon = Image.open(step.icon_path).convert("RGB")
                raw_icon = raw_icon.resize((ICON_SZ - 40, ICON_SZ - 40), Image.LANCZOS)
                canvas.paste(raw_icon, (cx - (ICON_SZ - 40)//2, circ_top + 20))
                draw = ImageDraw.Draw(canvas)
            except: pass

        badge_r = 30
        draw.ellipse([cx - badge_r, circ_top - 10, cx + badge_r, circ_top - 10 + badge_r*2], fill=palette["badge_bg"], outline=palette["card"], width=2)
        nd = draw.textbbox((0, 0), str(step.step_number), font=badge_font)
        draw.text((cx - (nd[2]-nd[0])//2, circ_top - 10 + badge_r - (nd[3]-nd[1])//2), str(step.step_number), fill=palette["badge_text"], font=badge_font)

        tx = cx + ICON_SZ//2 + 80
        ty = circ_top + 40
        draw.text((tx, ty), step.label.upper(), fill=palette["label_text"], font=label_font)

        desc_lines = _wrap_text(step.description, max_chars=70)[:2]
        dy = ty + 60
        for line in desc_lines:
            db = draw.textbbox((0, 0), line, font=desc_font)
            draw.text((tx, dy), line, fill=palette["desc_text"], font=desc_font)
            dy += db[3]-db[1] + 8

    canvas.save(output_path, "PNG", dpi=(240, 240))
    return output_path


def extract_domain(query: str) -> str: return "Corporate"


# ─────────────────────────────────────────────────────────────────────────────
# UI HELPERS (Keyword-based, no LLM prompts)
# ─────────────────────────────────────────────────────────────────────────────
PROCESS_KEYWORDS = [
    "step", "process", "flow", "pipeline", "journey", "lifecycle",
    "stages", "procedure", "workflow", "how", "disbursal", "approval",
    "underwriting", "onboarding", "loan", "hiring", "deployment",
    "involved", "sequence", "phase", "cycle", "steps involved",
]

def is_process_topic(text: str) -> bool:
    lower = text.lower()
    return any(kw in lower for kw in PROCESS_KEYWORDS)

def detect_mode(text: str) -> str:
    lines = [l for l in text.strip().splitlines() if l.strip()]
    words = len(text.split())
    return "topic" if (words <= 40 and len(lines) <= 3) else "content"

# TWO MASTER PROMPTS
# ─────────────────────────────────────────────────────────────────────────────

CLAUDE_SYSTEM = (
    "You are a world-class presentation consultant and industry analyst. "
    "You respond ONLY with valid JSON. No markdown fences, no preamble, no postscript."
)

CLAUDE_USER_TEMPLATE = """
USER REQUEST/CONTENT:
\"\"\"
{source}
\"\"\"

TASK:
1. Industry Detection: Identify the core industry/domain (e.g., Banking, Healthcare).
2. Presentation Structure: Design a complete {num_slides}-slide deck.
3. Flow Analysis: Determine if a process flow diagram is appropriate. If yes, plan 6 sequential steps.
4. Content Generation: Write complete content for every slide.

CONTENT TYPES SCHEMAS:
- title: {{"title": "...", "subtitle": "...", "speaker_notes": "..."}}
- bullets: {{"title": "...", "bullets": ["8-12 word point", "8-12 word point", ...], "image_suggestion": "photorealistic subject...", "speaker_notes": "..."}}
- two_column: {{"title": "...", "left_heading": "...", "left_bullets": ["...", "..."], "right_heading": "...", "right_bullets": ["...", "..."], "image_suggestion": "...", "speaker_notes": "..."}}
- stat_callout: {{"title": "...", "stats": [{{"value": "95%", "label": "Text"}}, ...], "body_text": "...", "speaker_notes": "..."}}
- timeline: {{"title": "...", "events": [{{"year": "2024", "label": "...", "detail": "..."}}, ...], "speaker_notes": "..."}}
- table: {{"title": "...", "headers": ["A", "B", "C"], "rows": [["cell", "cell", "cell"], ...], "speaker_notes": "..."}}
- quote: {{"title": "...", "quote": "...", "attribution": "...", "speaker_notes": "..."}}
- thank_you: {{"title": "Thank You", "message": "...", "contact": "...", "speaker_notes": "..."}}

RULES:
- VARIETY: Use different types across the deck (bullets, two_column, stat_callout, timeline).
- PROFESSIONALISM: Bullets must be punchy (8-12 words).
- CONTEXT: If user provided content, use it. If not, generate high-value insights.
- STRUCTURE: If the user provides a 'Slide Structure Recommendation', strictly follow that order and topic list for the slides.
- VISUALS: Provide 'image_suggestion' for slides (except title/thank_you).

RETURN JSON OBJECT:
{{
  "domain": "1-2 words",
  "has_process": boolean,
  "slides": [
    {{
      "slide_number": 1,
      "content_type": "title",
      "content": {{ ... matching schema ... }}
    }},
    ...
  ],
  "flow_steps": [ // exactamente 6 if has_process=true
    {{
      "step_number": 1,
      "label": "Short title",
      "description": "One specific sentence.",
      "icon_prompt": "simple flat business icon of..."
    }},
    ...
  ]
}}
"""

# NOVA Prompt Builder
def build_nova_prompt(subject: str, domain: str, is_icon: bool = False) -> str:
    if is_icon:
        return f"Flat style business icon for {domain}: {subject}. Strictly minimalist, single centered icon, white background, bold clean shapes, vivid professional colors, no text, no shadows."
    else:
        return f"Photorealistic high-end corporate photography for {domain} industry. Subject: {subject}. Cinematic lighting, minimalist aesthetic, 8k resolution, professional composition."

# ─────────────────────────────────────────────────────────────────────────────
# MASTER GENERATION CALLS
# ─────────────────────────────────────────────────────────────────────────────

def call_nova(prompt: str, size: int = 1024, is_icon: bool = False) -> Optional["Image.Image"]:
    if not PIL_AVAILABLE: return None
    w, h = (512, 512) if is_icon else (1024, 1024)
    
    body = {
        "taskType": "TEXT_IMAGE",
        "textToImageParams": {
            "text": prompt,
            "negativeText": "text, words, letters, watermark, distorted, low quality, messy, shadows for icons"
        },
        "imageGenerationConfig": {
            "numberOfImages": 1, "width": w, "height": h, "quality": "standard", "cfgScale": 8.0, "seed": 42
        }
    }
    try:
        resp = bedrock.invoke_model(modelId=NOVA_CANVAS, body=json.dumps(body), accept="application/json", contentType="application/json")
        result = json.loads(resp["body"].read())
        if result.get("error"): return None
        img_bytes = base64.b64decode(result["images"][0])
        return Image.open(BytesIO(img_bytes)).convert("RGBA")
    except: return None

def generate_all_in_one_presentation_data(source: str, num_slides: int) -> dict:
    # Escape braces in the source content to avoid .format() formatting errors
    safe_source = source.replace("{", "{{").replace("}", "}}")
    user_prompt = CLAUDE_USER_TEMPLATE.format(source=safe_source, num_slides=num_slides)
    max_tokens = 8192 # Set to maximum for detailed content generation

    for attempt in range(3):
        try:
            raw = call_bedrock(CLAUDE_SYSTEM, user_prompt, max_tokens=max_tokens)
            data = extract_json(raw)
            
            # Map slides to SlideData
            final_slides = []
            for s in data.get("slides", []):
                ct_str = s.get("content_type", "bullets")
                try: ct = ContentType(ct_str)
                except: ct = ContentType.BULLETS
                final_slides.append(SlideData(
                    slide_number=s.get("slide_number", len(final_slides)+1),
                    content_type=ct,
                    content=s.get("content", {})
                ))
            data["slides_data"] = final_slides
            
            # Map flow steps
            final_steps = []
            if data.get("has_process"):
                for i, stp in enumerate(data.get("flow_steps", [])):
                    final_steps.append(FlowStep(
                        step_number=stp.get("step_number", i+1),
                        label=stp.get("label", "Step"),
                        description=stp.get("description", ""),
                        icon_prompt=stp.get("icon_prompt", "icon")
                    ))
            data["flow_steps_data"] = final_steps
            return data
        except Exception as e:
            print(f"    ⚠️ Master Prompt Attempt {attempt+1} Failed: {e}")
    return {}

# ─────────────────────────────────────────────────────────────────────────────
# REPLACING LEGACY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def regenerate_step_icon(icon_prompt: str, step_number: int, domain: str = "Corporate") -> str:
    fname = f"icon_step{step_number:02d}_{uuid.uuid4().hex[:6]}.png"
    path = os.path.join(ICONS_DIR, fname)
    prompt = build_nova_prompt(icon_prompt, domain, is_icon=True)
    img = call_nova(prompt, is_icon=True)
    if img:
        img.save(path, "PNG")
        return path
    return _draw_fallback_icon(path, step_number)



# ─────────────────────────────────────────────────────────────────────────────
# PPTXGENJS SCRIPT BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def _build_js(slides: list[SlideData], theme: Theme, output_path: str, domain_bg_path: str = None) -> str:
    colors         = json.dumps(THEME_COLORS[theme])
    slides_js      = json.dumps([asdict(s) for s in slides], indent=2)
    output_path_js = output_path.replace("\\", "/")
    bg_js          = f'"{domain_bg_path}"' if domain_bg_path else 'null'

    return f"""
const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const pres    = new pptxgen();
pres.layout   = "LAYOUT_16x9";

const colors = {colors};
const slides = {slides_js};
const domainBg = {bg_js};

// ── Shared title bar ─────────────────────────────────────────
function titleBar(slide, title) {{
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0, y: 0, w: 10, h: 0.85,
    fill: {{ color: colors.primary }}, line: {{ color: colors.primary }}
  }});
  slide.addText(title, {{
    x: 0.4, y: 0, w: 9.2, h: 0.85,
    fontSize: 22, bold: true, color: "FFFFFF",
    fontFace: "Calibri", valign: "middle", margin: 0
  }});
}}

// ── Slide renderers ──────────────────────────────────────────
function renderTitle(slide, c) {{
  slide.background = domainBg ? {{ path: domainBg }} : {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0, y: 3.8, w: 10, h: 1.825,
    fill: {{ color: colors.secondary, transparency: 70 }}, line: {{ color: colors.primary }}
  }});
  slide.addText(c.title || "", {{
    x: 0.5, y: 1.1, w: 9, h: 1.8,
    fontSize: 44, bold: true, color: colors.accent,
    fontFace: "Calibri", align: "center", valign: "middle"
  }});
  slide.addText(c.subtitle || "", {{
    x: 0.5, y: 3.0, w: 9, h: 0.7,
    fontSize: 20, color: "FFFFFF",
    fontFace: "Calibri", align: "center"
  }});
}}

function renderTitleLeft(slide, c) {{
  slide.background = domainBg ? {{ path: domainBg }} : {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0, y: 0, w: 3, h: 5.625,
    fill: {{ color: colors.secondary, transparency: 20 }}
  }});
  slide.addText(c.title || "", {{
    x: 0.5, y: 1.5, w: 8, h: 1.5,
    fontSize: 44, bold: true, color: colors.accent,
    fontFace: "Calibri", align: "left", valign: "middle"
  }});
  slide.addText(c.subtitle || "", {{
    x: 0.5, y: 3.2, w: 6, h: 1.0,
    fontSize: 20, color: colors.secondary,
    fontFace: "Calibri", align: "left"
  }});
}}

function renderTitleDark(slide, c) {{
  slide.background = {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0.5, y: 0.5, w: 9, h: 4.625,
    fill: {{ color: colors.primary }},
    line: {{ color: colors.secondary, width: 3 }},
    rectRadius: 0.2
  }});
  slide.addText(c.title || "", {{
    x: 1.0, y: 1.5, w: 8, h: 2,
    fontSize: 48, bold: true, color: "FFFFFF",
    fontFace: "Georgia", align: "center", valign: "middle"
  }});
  slide.addText(c.subtitle || "", {{
    x: 1.0, y: 3.5, w: 8, h: 1.0,
    fontSize: 22, color: colors.secondary,
    fontFace: "Calibri", align: "center"
  }});
}}

function renderTitleSplit(slide, c) {{
  slide.background = domainBg ? {{ path: domainBg }} : {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0, y: 0, w: 4, h: 5.625,
    fill: {{ color: colors.primary }}
  }});
  slide.addText(c.title || "", {{
    x: 4.5, y: 1.5, w: 5, h: 2,
    fontSize: 50, bold: true, color: "FFFFFF",
    fontFace: "Calibri", shadow: {{type:'outer', color:'000000', blur:3, offset:2, angle:45}}
  }});
  slide.addText(c.subtitle || "", {{
    x: 0.5, y: 2.0, w: 3, h: 1.5,
    fontSize: 24, color: colors.secondary,
    fontFace: "Calibri", align: "left"
  }});
}}

function renderTitleClean(slide, c) {{
  slide.background = {{ color: "FFFFFF" }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 3, y: 1.5, w: 4, h: 0.05,
    fill: {{ color: colors.primary }}
  }});
  slide.addText(c.title || "", {{
    x: 0.5, y: 1.8, w: 9, h: 1.2,
    fontSize: 44, bold: true, color: colors.primary,
    fontFace: "Calibri", align: "center"
  }});
  slide.addText(c.subtitle || "", {{
    x: 0.5, y: 3.2, w: 9, h: 0.8,
    fontSize: 18, color: "666666",
    fontFace: "Calibri", align: "center", italic: true
  }});
}}

function renderTitleAccent(slide, c) {{
  slide.background = domainBg ? {{ path: domainBg }} : {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 0, w: 10, h: 0.5, fill: {{ color: colors.primary }} }});
  slide.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 5.125, w: 10, h: 0.5, fill: {{ color: colors.primary }} }});
  slide.addText(c.title || "", {{
    x: 1, y: 1.5, w: 8, h: 2,
    fontSize: 52, bold: true, color: "FFFFFF",
    fontFace: "Calibri", align: "center", shadow: {{type:'outer', color:'000000', blur:5}}
  }});
  slide.addText(c.subtitle || "", {{
    x: 1, y: 3.5, w: 8, h: 0.6,
    fontSize: 22, color: colors.secondary,
    fontFace: "Calibri", align: "center", fill: {{color:'000000', transparency:30}}
  }});
}}

function renderBullets(slide, c) {{
  slide.background = {{ color: "F8F9FA" }};
  titleBar(slide, c.title || "");
  const bullets = (c.bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.bullets.length - 1), fontSize: 16, color: "2D3436", fontFace: "Calibri" }}
  }}));

  if (c.generated_image_path) {{
      slide.addText(bullets, {{ x: 0.5, y: 1.1, w: 4.5, h: 4.1 }});
      slide.addImage({{ path: c.generated_image_path, x: 5.3, y: 1.1, w: 4.2, h: 4.2, sizing: {{type: "cover", w: 4.2, h: 4.2}} }});
  }} else {{
      slide.addText(bullets, {{ x: 0.5, y: 1.1, w: 8.5, h: 4.1 }});
      if (c.image_suggestion) {{
        slide.addText("Image: " + c.image_suggestion, {{
          x: 0.5, y: 5.25, w: 9, h: 0.3,
          fontSize: 9, color: "ADB5BD", italic: true, fontFace: "Calibri"
        }});
      }}
  }}
}}

function renderBulletsBox(slide, c) {{
  slide.background = {{ color: colors.secondary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0.5, y: 0.5, w: 9, h: 4.625,
    fill: {{ color: "FFFFFF" }},
    line: {{ color: colors.primary, width: 2 }},
    rectRadius: 0.1
  }});
  slide.addText(c.title || "", {{
    x: 0.8, y: 0.6, w: 8.4, h: 0.8,
    fontSize: 26, bold: true, color: colors.primary,
    fontFace: "Calibri", valign: "middle"
  }});
  const bullets = (c.bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.bullets.length - 1), fontSize: 16, color: "2D3436", fontFace: "Calibri" }}
  }}));
  slide.addText(bullets, {{ x: 0.8, y: 1.5, w: 8.4, h: 3.4 }});
}}

function renderTwoColumn(slide, c) {{
  slide.background = {{ color: "F8F9FA" }};
  titleBar(slide, c.title || "");
  const hasImg = !!c.generated_image_path;
  const colW = hasImg ? 3.0 : 4.4;
  const leftX = 0.3;
  const rightX = hasImg ? 3.5 : 5.3;
  const lineX = hasImg ? 3.4 : 4.9;

  slide.addShape(pres.shapes.RECTANGLE, {{
    x: leftX, y: 1.1, w: colW, h: 0.45,
    fill: {{ color: colors.primary }}, line: {{ color: colors.primary }}
  }});
  slide.addText(c.left_heading || "Left", {{
    x: leftX, y: 1.1, w: colW, h: 0.45,
    fontSize: 13, bold: true, color: "FFFFFF", fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  }});
  const lb = (c.left_bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.left_bullets.length - 1), fontSize: 13, color: "2D3436", fontFace: "Calibri" }}
  }}));
  slide.addText(lb, {{ x: leftX, y: 1.65, w: colW, h: 3.7 }});

  slide.addShape(pres.shapes.RECTANGLE, {{
    x: rightX, y: 1.1, w: colW, h: 0.45,
    fill: {{ color: colors.secondary }}, line: {{ color: colors.secondary }}
  }});
  slide.addText(c.right_heading || "Right", {{
    x: rightX, y: 1.1, w: colW, h: 0.45,
    fontSize: 13, bold: true, color: colors.primary, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  }});
  const rb = (c.right_bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.right_bullets.length - 1), fontSize: 13, color: "2D3436", fontFace: "Calibri" }}
  }}));
  slide.addText(rb, {{ x: rightX, y: 1.65, w: colW, h: 3.7 }});

  slide.addShape(pres.shapes.LINE, {{
    x: lineX, y: 1.1, w: 0, h: 4.2,
    line: {{ color: colors.primary, width: 1.5 }}
  }});

  if (hasImg) {{
      slide.addImage({{ path: c.generated_image_path, x: 6.8, y: 1.1, w: 2.8, h: 4.2, sizing: {{type: "cover", w: 2.8, h: 4.2}} }});
  }}
}}

function renderTwoColumnCards(slide, c) {{
  slide.background = {{ color: "E2E8F0" }};
  titleBar(slide, c.title || "");
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0.5, y: 1.1, w: 4.2, h: 4.0,
    fill: {{ color: "FFFFFF" }},
    line: {{ color: "CBD5E1" }},
    rectRadius: 0.1
  }});
  slide.addText(c.left_heading || "Left", {{
    x: 0.5, y: 1.1, w: 4.2, h: 0.6,
    fontSize: 18, bold: true, color: colors.primary, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0,
    fill: {{ color: colors.secondary, transparency: 80 }}
  }});
  const lb = (c.left_bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.left_bullets.length - 1), fontSize: 13, color: "1E293B", fontFace: "Calibri" }}
  }}));
  slide.addText(lb, {{ x: 0.6, y: 1.8, w: 4.0, h: 3.2 }});

  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 5.1, y: 1.1, w: 4.2, h: 4.0,
    fill: {{ color: "FFFFFF" }},
    line: {{ color: "CBD5E1" }},
    rectRadius: 0.1
  }});
  slide.addText(c.right_heading || "Right", {{
    x: 5.1, y: 1.1, w: 4.2, h: 0.6,
    fontSize: 18, bold: true, color: colors.primary, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0,
    fill: {{ color: colors.secondary, transparency: 80 }}
  }});
  const rb = (c.right_bullets || []).map((b, i) => ({{
    text: b,
    options: {{ bullet: true, breakLine: i < (c.right_bullets.length - 1), fontSize: 13, color: "1E293B", fontFace: "Calibri" }}
  }}));
  slide.addText(rb, {{ x: 5.2, y: 1.8, w: 4.0, h: 3.2 }});
}}

function renderStatCallout(slide, c) {{
  slide.background = {{ color: colors.primary }};
  titleBar(slide, c.title || "");
  const stats = (c.stats || []).slice(0, 4);
  const count = stats.length || 1;
  const colW  = 9 / count;
  stats.forEach((s, i) => {{
    const x = 0.5 + i * colW;
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y: 1.3, w: colW - 0.3, h: 2.8,
      fill: {{ color: "FFFFFF", transparency: 10 }}, line: {{ color: colors.secondary }}
    }});
    slide.addText(s.value || "", {{
      x, y: 1.5, w: colW - 0.3, h: 1.4,
      fontSize: 48, bold: true, color: colors.secondary,
      fontFace: "Calibri", align: "center"
    }});
    slide.addText(s.label || "", {{
      x, y: 3.0, w: colW - 0.3, h: 0.8,
      fontSize: 14, color: "FFFFFF", fontFace: "Calibri", align: "center"
    }});
  }});
  if (c.body_text) {{
    slide.addText(c.body_text, {{
      x: 0.5, y: 4.5, w: 9, h: 0.7,
      fontSize: 14, color: colors.secondary,
      fontFace: "Calibri", align: "center", italic: true
    }});
  }}
}}

function renderTimeline(slide, c) {{
  slide.background = {{ color: "F8F9FA" }};
  titleBar(slide, c.title || "");
  const events = (c.events || []).slice(0, 6);
  const count  = events.length || 1;
  const stepW  = 9 / count;
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0.5, y: 2.8, w: 9, h: 0.12,
    fill: {{ color: colors.primary }}, line: {{ color: colors.primary }}
  }});
  events.forEach((e, i) => {{
    const cx = 0.5 + i * stepW + stepW / 2;
    slide.addShape(pres.shapes.OVAL, {{
      x: cx - 0.18, y: 2.67, w: 0.36, h: 0.36,
      fill: {{ color: colors.primary }}, line: {{ color: "FFFFFF" }}
    }});
    slide.addText(e.year || "", {{
      x: cx - 0.6, y: 1.8, w: 1.2, h: 0.5,
      fontSize: 14, bold: true, color: colors.primary,
      fontFace: "Calibri", align: "center"
    }});
    slide.addText(e.label || "", {{
      x: cx - 0.7, y: 3.15, w: 1.4, h: 0.45,
      fontSize: 12, bold: true, color: "2D3436",
      fontFace: "Calibri", align: "center"
    }});
    slide.addText(e.detail || "", {{
      x: cx - 0.7, y: 3.7, w: 1.4, h: 1.5,
      fontSize: 10, color: "636E72",
      fontFace: "Calibri", align: "center"
    }});
  }});
}}

function renderTable(slide, c) {{
  slide.background = {{ color: "F8F9FA" }};
  titleBar(slide, c.title || "");
  const headers = (c.headers || []).map(h => ({{
    text: h,
    options: {{ bold: true, color: "FFFFFF", fill: {{ color: colors.primary }}, align: "center" }}
  }}));
  const rows = [headers, ...(c.rows || []).map(row =>
    row.map(cell => ({{ text: String(cell), options: {{ fontSize: 13, color: "2D3436" }} }}))
  )];
  slide.addTable(rows, {{
    x: 0.4, y: 1.1, w: 9.2,
    border: {{ pt: 1, color: "DEE2E6" }},
    fontSize: 13, fontFace: "Calibri"
  }});
}}

function renderQuote(slide, c) {{
  slide.background = {{ color: colors.primary }};
  titleBar(slide, c.title || "");
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0.4, y: 1.1, w: 0.12, h: 4.0,
    fill: {{ color: colors.secondary }}, line: {{ color: colors.secondary }}
  }});
  slide.addText(c.quote || "", {{
    x: 0.8, y: 1.5, w: 8.5, h: 2.8,
    fontSize: 22, italic: true, color: "FFFFFF",
    fontFace: "Georgia", valign: "middle"
  }});
  slide.addText(c.attribution || "", {{
    x: 0.8, y: 4.5, w: 8.5, h: 0.6,
    fontSize: 14, color: colors.secondary,
    fontFace: "Calibri", align: "right"
  }});
}}

function renderThankYou(slide, c) {{
  slide.background = domainBg ? {{ path: domainBg }} : {{ color: colors.primary }};
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 2, y: 1.5, w: 6, h: 0.1,
    fill: {{ color: colors.secondary }}, line: {{ color: colors.secondary }}
  }});
  slide.addText(c.title || "Thank You", {{
    x: 0.5, y: 1.8, w: 9, h: 1.5,
    fontSize: 48, bold: true, color: "FFFFFF",
    fontFace: "Calibri", align: "center"
  }});
  slide.addText(c.message || "", {{
    x: 0.5, y: 3.4, w: 9, h: 1.0,
    fontSize: 18, color: colors.secondary,
    fontFace: "Calibri", align: "center"
  }});
  if (c.contact) {{
    slide.addText(c.contact, {{
      x: 0.5, y: 4.5, w: 9, h: 0.6,
      fontSize: 14, color: colors.secondary,
      fontFace: "Calibri", align: "center", italic: true
    }});
  }}
}}

// ── DIAGRAM renderer — NATIVE PPTX SHAPES (fully editable after download) ─────
//
// Everything on this slide is a native PowerPoint object:
//   ✅ Title text box           → double-click to edit in PowerPoint
//   ✅ Icon images (per step)   → right-click → "Change Picture" in PowerPoint
//   ✅ Step label text boxes    → double-click to edit in PowerPoint
//   ✅ Description text boxes   → double-click to edit in PowerPoint
//   ✅ Circle shapes            → click to resize/recolor in PowerPoint
//   ✅ Step number badges       → click to edit in PowerPoint
//   ✅ Arrow connectors         → click to recolor/resize in PowerPoint
//
// No PNG image is embedded — every element is independently editable.
// ─────────────────────────────────────────────────────────────────────────────
function renderDiagram(slide, c) {{
  const steps = (c.steps || []);
  const n = steps.length;
  if (n === 0) return;

  // Dark background
  slide.background = {{ color: "0D1127" }};

  // ── Title bar ──────────────────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {{
    x: 0, y: 0, w: 10, h: 0.78,
    fill: {{ color: colors.primary }},
    line: {{ color: colors.primary }}
  }});
  slide.addText(c.title || "Process Flow", {{
    x: 0.3, y: 0, w: 9.4, h: 0.78,
    fontSize: 22, bold: true, color: "FFFFFF",
    fontFace: "Calibri", valign: "middle"
  }});

  // ── Layout geometry ────────────────────────────────────────────────────────
  // Slide is 10" wide × 5.625" tall (16:9)
  const PAD_X    = 0.22;               // left+right edge padding
  const ICON_SZ  = 1.05;              // diameter of each step circle (inches)
  const CIRCLE_Y = 1.05;              // top-edge of icon circles
  const colW     = (10 - PAD_X * 2) / n;  // width per step column
  const LABEL_Y  = CIRCLE_Y + ICON_SZ + 0.16;  // top of label text
  const DESC_Y   = LABEL_Y + 0.40;    // top of description text
  const lineY    = CIRCLE_Y + ICON_SZ / 2;      // vertical center of circles

  // ── Horizontal backbone connector ──────────────────────────────────────────
  // Drawn first so it appears behind the circles
  slide.addShape(pres.shapes.LINE, {{
    x: PAD_X + colW * 0.5,
    y: lineY,
    w: colW * (n - 1),
    h: 0,
    line: {{ color: "CADCFC", width: 1.2, transparency: 55 }}
  }});

  steps.forEach((step, i) => {{
    const cx         = PAD_X + colW * i + colW / 2;   // center of this column
    const circleLeft = cx - ICON_SZ / 2;               // left edge of circle

    // ── Step circle — white fill so icon colors show clearly ─────────────────
    slide.addShape(pres.shapes.OVAL, {{
      x: circleLeft,
      y: CIRCLE_Y,
      w: ICON_SZ,
      h: ICON_SZ,
      fill: {{ color: "FFFFFF" }},
      line: {{ color: "CADCFC", width: 2.5 }}
    }});

    // ── Icon image — each step gets its own independently replaceable picture ─
    // In PowerPoint: right-click the image → "Change Picture" to swap it.
    const iconPath = step.icon_path || "";
    if (iconPath && fs.existsSync(iconPath)) {{
      const pad = 0.09;
      slide.addImage({{
        path: iconPath,
        x: circleLeft + pad,
        y: CIRCLE_Y + pad,
        w: ICON_SZ - pad * 2,
        h: ICON_SZ - pad * 2,
        sizing: {{ type: "contain", w: ICON_SZ - pad * 2, h: ICON_SZ - pad * 2 }}
      }});
    }}

    // ── Step number badge ─────────────────────────────────────────────────────
    const BR = 0.19;   // badge radius (inches)
    // Badge circle
    slide.addShape(pres.shapes.OVAL, {{
      x: cx - BR,
      y: CIRCLE_Y - BR * 0.55,
      w: BR * 2,
      h: BR * 2,
      fill: {{ color: "CADCFC" }},
      line: {{ color: "0D1127", width: 0.75 }}
    }});
    // Badge number — separate text box, independently editable
    slide.addText(String(step.step_number || i + 1), {{
      x: cx - BR,
      y: CIRCLE_Y - BR * 0.55,
      w: BR * 2,
      h: BR * 2,
      fontSize: 8.5, bold: true, color: "1E2761",
      fontFace: "Calibri", align: "center", valign: "middle"
    }});

    // ── Step label — fully editable text box ──────────────────────────────────
    // In PowerPoint: double-click to edit the label text directly.
    slide.addText((step.label || "").toUpperCase(), {{
      x: circleLeft - 0.06,
      y: LABEL_Y,
      w: ICON_SZ + 0.12,
      h: 0.40,
      fontSize: 8.5, bold: true, color: "FFFFFF",
      fontFace: "Calibri", align: "center", valign: "top",
      wrap: true
    }});

    // ── Step description — fully editable text box ────────────────────────────
    // In PowerPoint: double-click to edit the description text directly.
    slide.addText(step.description || "", {{
      x: circleLeft - 0.12,
      y: DESC_Y,
      w: ICON_SZ + 0.24,
      h: 1.75,
      fontSize: 7.0, color: "AABCD4",
      fontFace: "Calibri", align: "center", valign: "top",
      wrap: true
    }});

    // ── Arrow to next step ────────────────────────────────────────────────────
    // Drawn as a native LINE shape — click to recolor or resize in PowerPoint.
    if (i < n - 1) {{
      const arrowX = cx + ICON_SZ / 2 + 0.03;
      const arrowW = colW - ICON_SZ - 0.06;
      slide.addShape(pres.shapes.LINE, {{
        x: arrowX,
        y: lineY,
        w: arrowW,
        h: 0,
        line: {{ color: "CADCFC", width: 1.8, endArrowType: "arrow" }}
      }});
    }}
  }});
}}

// ── Router ───────────────────────────────────────────────────
const RENDERERS = {{
  title:             renderTitle,
  title_left:        renderTitleLeft,
  title_dark:        renderTitleDark,
  title_split:       renderTitleSplit,
  title_clean:       renderTitleClean,
  title_accent:      renderTitleAccent,
  bullets:           renderBullets,
  bullets_box:       renderBulletsBox,
  two_column:        renderTwoColumn,
  two_column_cards:  renderTwoColumnCards,
  stat_callout:      renderStatCallout,
  timeline:          renderTimeline,
  table:             renderTable,
  quote:             renderQuote,
  diagram:           renderDiagram,
  thank_you:         renderThankYou,
}};

slides.forEach(sd => {{
  const slide = pres.addSlide();
  const fn    = RENDERERS[sd.layout_override] || RENDERERS[sd.content_type] || renderBullets;
  fn(slide, sd.content);
  if (sd.content && sd.content.speaker_notes) {{
    slide.addNotes(sd.content.speaker_notes);
  }}
}});

pres.writeFile({{ fileName: "{output_path_js}" }})
  .then(() => console.log("OK"))
  .catch(e => {{ console.error("ERR:" + e.message); process.exit(1); }});
"""


# ─────────────────────────────────────────────────────────────────────────────
# PPTX EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def export_pptx(slides: list[SlideData], theme: Theme, filename: str | None = None) -> bytes:
    if filename is None:
        filename = f"presentation_{uuid.uuid4().hex[:8]}.pptx"

    domain_bg_path = None
    if slides and isinstance(slides[0].content, dict) and "domain_bg_path" in slides[0].content:
        domain_bg_path = slides[0].content["domain_bg_path"]

    output_path = os.path.abspath(os.path.join(OUTPUT_DIR, filename))
    script      = _build_js(slides, theme, output_path, domain_bg_path)

    script_path = os.path.join(PROJECT_DIR, f"_pptx_tmp_{uuid.uuid4().hex[:8]}.js")
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)

    try:
        result = subprocess.run(
            ["node", script_path],
            capture_output=True, text=True, timeout=180,
            cwd=PROJECT_DIR,
        )
        if result.returncode != 0:
            raise RuntimeError(f"PptxGenJS error:\n{result.stdout}\n{result.stderr}")

        if not os.path.exists(output_path):
            raise RuntimeError(f"Output file not created: {output_path}")

        # Sync to S3 before internal cleanup
        sync_to_s3(output_path)

        with open(output_path, "rb") as f:
            pptx_data = f.read()
    finally:
        # Robust cleanup for Windows
        for p in [script_path, output_path]:
            if os.path.exists(p):
                for attempt in range(5):
                    try:
                        os.unlink(p)
                        break
                    except Exception as e:
                        if attempt == 4:
                            print(f"    ⚠️ Failed to cleanup {p} after 5 attempts: {e}")
                        time.sleep(0.5)

    return pptx_data


# ─────────────────────────────────────────────────────────────────────────────
# THEME AUTO-SELECTOR
# ─────────────────────────────────────────────────────────────────────────────
def pick_theme(text: str) -> Theme:
    lower = text.lower()
    if any(w in lower for w in ["energy", "nature", "green", "environment", "eco", "forest", "climate"]):
        return Theme.FOREST_MOSS
    if any(w in lower for w in ["finance", "invest", "bank", "market", "stock", "fund", "revenue", "loan", "lending", "credit"]):
        return Theme.MIDNIGHT_EXECUTIVE
    if any(w in lower for w in ["health", "medical", "care", "hospital", "doctor", "patient"]):
        return Theme.OCEAN_GRADIENT
    if any(w in lower for w in ["startup", "product", "launch", "growth", "innovation", "brand"]):
        return Theme.CORAL_ENERGY
    if any(w in lower for w in ["history", "culture", "art", "heritage", "tradition"]):
        return Theme.WARM_TERRACOTTA
    return Theme.MIDNIGHT_EXECUTIVE


# Redundant functions kept for internal compatibility if needed, but not used by main pipeline
def extract_domain(query: str) -> str:
    return "Corporate"
# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────
def prepare_slides(user_input: str) -> tuple[list[SlideData], Theme]:
    theme = pick_theme(user_input)
    word_count = len(user_input.split())
    
    # Better scaling for content: More slides for more detail
    if word_count <= 40:
        num_slides = 10
    elif word_count <= 150:
        num_slides = 12
    else:
        # Heavily detailed content like end-to-end lifecycles should target 14-16 slides
        num_slides = min(16, max(12, word_count // 30))

    print(f"\n  [1/4] CALLING CLAUDE MASTER PROMPT...")
    data = generate_all_in_one_presentation_data(user_input, num_slides)
    
    slides = data.get("slides_data", [])
    domain = data.get("domain", "Corporate")
    has_process = data.get("has_process", False)
    flow_steps = data.get("flow_steps_data", [])

    print(f"     Domain Detected: {domain}")
    print(f"     Slides Created : {len(slides)}")
    print(f"     Flow Diagram   : {'YES' if has_process else 'NO'}")

    # ── Step 2: Parallel Visual Generation ───────────────────────────────────
    print(f"\n  [2/4] GENERATING AI VISUALS (PARALLEL)...")
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = {}
        
        # 1. Background (only for first slide usually, or abstracted)
        bg_p = build_nova_prompt("Abstract elegant professional presentation background", domain)
        futures[executor.submit(call_nova, bg_p)] = ("bg", slides[0].content if slides else None)
        
        # 2. Slide Images
        for s in slides:
            suggest = s.content.get("image_suggestion")
            if suggest:
                p = build_nova_prompt(suggest, domain)
                futures[executor.submit(call_nova, p)] = ("slide", s.content)
                
        # 3. Diagram Icons
        if has_process and flow_steps:
            for stp in flow_steps:
                p = build_nova_prompt(stp.icon_prompt, domain, is_icon=True)
                futures[executor.submit(call_nova, p, is_icon=True)] = ("icon", stp)

        # Collect results as they finish
        for future in concurrent.futures.as_completed(futures):
            rtype, target = futures[future]
            try:
                img = future.result()
                if img:
                    if rtype == "bg":
                        bg_path = os.path.join(OUTPUT_DIR, f"bg_{uuid.uuid4().hex[:6]}.png")
                        img.save(bg_path, "PNG")
                        sync_to_s3(bg_path)
                        if target: target["domain_bg_path"] = bg_path.replace("\\", "/")
                    elif rtype == "slide":
                        img_path = os.path.join(OUTPUT_DIR, f"slide_{uuid.uuid4().hex[:6]}.png")
                        img.save(img_path, "PNG")
                        sync_to_s3(img_path)
                        target["generated_image_path"] = img_path.replace("\\", "/")
                    elif rtype == "icon":
                        ipath = os.path.join(ICONS_DIR, f"icon_{uuid.uuid4().hex[:6]}.png")
                        img.save(ipath, "PNG")
                        sync_to_s3(ipath)
                        target.icon_path = ipath
                elif rtype == "icon":
                    ipath = os.path.join(ICONS_DIR, f"icon_fb_{uuid.uuid4().hex[:6]}.png")
                    _draw_fallback_icon(ipath, target.step_number)
                    sync_to_s3(ipath)
                    target.icon_path = ipath
            except Exception as e:
                print(f"    ⚠️ Error in parallel visual gen: {e}")

    # ── Step 3: Diagram Assembly ─────────────────────────────────────────────
    if has_process and flow_steps:
        print(f"  [3/4] ASSEMBLING FLOW DIAGRAM...")
        title = data.get("title", f"{domain} Workflow")
        res = reassemble_diagram([asdict(s) for s in flow_steps], title, user_input)
        
        other = [s for s in slides if s.content_type != ContentType.THANK_YOU]
        thank = [s for s in slides if s.content_type == ContentType.THANK_YOU]
        diag_num = (other[-1].slide_number + 1) if other else 2
        
        diag_slide = SlideData(
            slide_number=diag_num,
            content_type=ContentType.DIAGRAM,
            content={
                "title": title,
                "diagram_image": res["diagram"],
                "diagram_images": res,
                "steps": [asdict(s) for s in flow_steps],
                "speaker_notes": f"Process flow for {domain}."
            }
        )
        for s in thank: s.slide_number = diag_num + 1
        slides = other + [diag_slide] + thank

    return slides, theme

def run(user_input: str) -> tuple[bytes, list]:
    slides, theme = prepare_slides(user_input)
    pid = uuid.uuid4().hex[:8]
    print(f"\n  [4/4] Exporting .pptx...")
    pptx_data = export_pptx(slides, theme, filename=f"{pid}.pptx")
    return pptx_data, slides


# ─────────────────────────────────────────────────────────────────────────────
# ✏️  EDIT YOUR INPUT HERE
# ─────────────────────────────────────────────────────────────────────────────
USER_INPUT = """
give me any query
"""


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":

    print("=" * 57)
    print("    AI PowerPoint Generator  (AWS Bedrock + Nova Canvas)")
    print("=" * 57)

    user_input = USER_INPUT.strip()
    if not user_input:
        print("  No input provided. Edit USER_INPUT in the script.")
        sys.exit(1)

    if not PIL_AVAILABLE:
        print("  ❌  Pillow is required for diagram generation.")
        print("       Run: pip install Pillow")
        sys.exit(1)

    print(f"\n  Input : {user_input[:90]}{'...' if len(user_input) > 90 else ''}")

    try:
        print("Backend script now configured to return raw bytes, saving locally handled via Streamlit UI.")
        print("=" * 57)
    except Exception as e:
        print(f"\n  ❌  Error: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)