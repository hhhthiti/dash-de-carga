from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont


def _font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return ImageFont.load_default()


def gerar_imagem_reporte(
    output: Path,
    data_ref: str,
    grade_count: int,
    zles_count: int,
    material_count: int,
    status: str,
    message: str = "",
) -> Path:
    image = Image.new("RGB", (1200, 630), "#0b1220")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((45, 42, 1155, 588), radius=24, fill="#111827", outline="#334155", width=2)
    draw.rectangle((45, 42, 1155, 52), fill="#38bdf8")
    draw.text((80, 86), "Dashboard de Carga", font=_font(42, True), fill="#f8fafc")
    draw.text((80, 145), f"Automacao Oracle VM - {data_ref}", font=_font(20), fill="#94a3b8")

    metrics = [
        ("DTs na grade", str(grade_count), "#60a5fa"),
        ("DTs na ZLES002", str(zles_count), "#22c55e"),
        ("Materiais", str(material_count), "#f59e0b"),
        ("Execucao", status.upper(), "#a78bfa" if status != "error" else "#ef4444"),
    ]
    for index, (label, value, color) in enumerate(metrics):
        left = 80 + index * 260
        draw.rounded_rectangle((left, 205, left + 230, 325), radius=14, fill="#0f172a", outline=color, width=2)
        draw.text((left + 18, 228), label, font=_font(15, True), fill=color)
        draw.text((left + 18, 270), value, font=_font(30, True), fill="#f8fafc")

    draw.text((80, 380), "Ultima execucao", font=_font(17, True), fill="#cbd5e1")
    draw.multiline_text((80, 420), message[:260] or "Processamento concluido.", font=_font(17), fill="#94a3b8", spacing=8)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, format="PNG", optimize=True)
    return output


def imagem_base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")
