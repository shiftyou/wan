#!/usr/bin/env python3
"""홍보여부·이름·수술명·수술날짜를 입력받아 촬영 시작 QR을 생성하고 기록한다."""

from __future__ import annotations

import csv
import re
import sys
from datetime import date
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = BASE_DIR / "qr_codes"
RECORDS_PATH = BASE_DIR / "records.csv"
KOREAN_FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\malgun.ttf"),  # Windows
    Path("/System/Library/Fonts/AppleSDGothicNeo.ttc"),  # macOS
    Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),  # Linux (나눔고딕 설치 시)
]


def resolve_korean_font_path() -> Path | None:
    for candidate in KOREAN_FONT_CANDIDATES:
        if candidate.exists():
            return candidate
    return None

PROMO_CHOICES = ("HP", "HT", "일반")
DATE_RE = re.compile(r"^\d{6}$")
RECORD_FIELDS = ["생성일시", "홍보여부", "이름", "수술명", "수술날짜"]


def safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\n\r\t]', " ", value).strip()
    return re.sub(r"\s+", " ", value)[:80].rstrip() or "환자"


def clean_field(value: str, field_name: str) -> str:
    value = " ".join(value.split())
    if not value:
        raise SystemExit(f"{field_name}은(는) 비워 둘 수 없습니다.")
    if "_" in value:
        raise SystemExit(f"{field_name}에는 '_' 문자를 사용할 수 없습니다.")
    return value


def normalize_surgery(value: str) -> str:
    """수술명에 입력된 스페이스를 콤마로 통일한다 (예: '쌍꺼풀 코성형' -> '쌍꺼풀,코성형')."""
    value = re.sub(r"\s*,\s*", ",", value)
    value = re.sub(r"\s+", ",", value)
    return value


def prompt_promo() -> str:
    while True:
        value = input(f"홍보여부 ({'/'.join(PROMO_CHOICES)}, 엔터=일반): ").strip().upper()
        if not value:
            return "일반"
        matched = next((choice for choice in PROMO_CHOICES if value == choice.upper()), None)
        if matched is None:
            print(f"  {', '.join(PROMO_CHOICES)} 중 하나를 입력하세요.")
            continue
        confirm = input(f"  -> {matched} 맞습니까? (Y/n): ").strip().upper()
        if confirm in ("", "Y"):
            return matched
        print("  다시 입력해 주세요.")


def prompt_date() -> str:
    while True:
        value = input("수술날짜 (YYMMDD): ").strip()
        if DATE_RE.fullmatch(value):
            return value
        print("  YYMMDD 형식으로 입력하세요. 예: 260808")


def make_qr(payload: str, output_path: Path) -> None:
    try:
        import qrcode
        from PIL import ImageDraw, ImageFont
    except ImportError:
        print("qrcode와 Pillow가 필요합니다. 아래 명령을 한 번 실행하세요:\n"
              "  .venv/bin/python -m pip install qrcode[pil]")
        sys.exit(1)

    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_H,
                       box_size=12, border=4)
    qr.add_data(payload)
    qr.make(fit=True)
    image = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    # 촬영자가 QR 내용을 확인할 수 있도록 페이로드를 그대로 표기한다.
    font_path = resolve_korean_font_path()
    font = (
        ImageFont.truetype(str(font_path), size=26)
        if font_path
        else ImageFont.load_default()
    )
    label = payload
    draw = ImageDraw.Draw(image)
    label_box = draw.textbbox((0, 0), label, font=font)
    label_height = label_box[3] - label_box[1] + 28
    from PIL import Image
    labeled = Image.new("RGB", (image.width, image.height + label_height), "white")
    labeled.paste(image, (0, 0))
    draw = ImageDraw.Draw(labeled)
    label_width = label_box[2] - label_box[0]
    draw.text(((image.width - label_width) // 2, image.height + 12), label,
              fill="black", font=font)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    labeled.save(output_path)


def append_record(promo: str, name: str, surgery: str, surgery_date: str) -> None:
    is_new = not RECORDS_PATH.exists()
    with RECORDS_PATH.open("a", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        if is_new:
            writer.writerow(RECORD_FIELDS)
        writer.writerow([date.today().isoformat(), promo, name, surgery, surgery_date])


def create_session(promo: str, name: str, surgery: str, surgery_date: str, output_dir: Path) -> None:
    payload = (
        f"{name}_{surgery}_{surgery_date}"
        if promo == "일반"
        else f"{promo}_{name}_{surgery}_{surgery_date}"
    )
    filename_base = safe_filename(payload)
    make_qr(payload, output_dir / f"{filename_base}.png")
    append_record(promo, name, surgery, surgery_date)
    print(f"QR: {output_dir / f'{filename_base}.png'}")
    print(f"기록: {RECORDS_PATH}")


def main() -> None:
    print("환자 촬영 QR 생성기")
    promo = prompt_promo()
    name = clean_field(input("이름: "), "이름")
    surgery = normalize_surgery(clean_field(input("수술명: "), "수술명"))
    surgery_date = prompt_date()
    create_session(promo, name, surgery, surgery_date, DEFAULT_OUTPUT_DIR)


if __name__ == "__main__":
    main()
