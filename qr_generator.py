#!/usr/bin/env python3
"""홍보여부·이름·수술명·수술날짜를 입력받아 촬영 시작 QR을 생성하고 기록한다."""

from __future__ import annotations

import csv
import ctypes
import json
import os
import re
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None


# PyInstaller onefile로 얼리면 __file__은 실행할 때마다 생기는 임시 압축 해제
# 폴더를 가리키므로, exe가 실제로 위치한 폴더를 기준으로 삼아야 config/records가
# 다음 실행에도 남아있는다.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
RECORDS_PATH = BASE_DIR / "records.csv"
CONFIG_PATH = BASE_DIR / "qr_generator_config.json"
KOREAN_FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\malgun.ttf"),  # Windows
    Path("/System/Library/Fonts/AppleSDGothicNeo.ttc"),  # macOS
    Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),  # Linux (나눔고딕 설치 시)
]


def load_config() -> dict:
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_config(config: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def resolve_korean_font_path() -> Path | None:
    for candidate in KOREAN_FONT_CANDIDATES:
        if candidate.exists():
            return candidate
    return None

PROMO_CHOICES = ("HP", "HT", "일반")
RECORD_FIELDS = ["생성일시", "홍보여부", "이름", "수술명", "수술날짜"]


class ValidationError(Exception):
    pass


def safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\n\r\t]', " ", value).strip()
    return re.sub(r"\s+", " ", value)[:80].rstrip() or "환자"


def clean_field(value: str, field_name: str) -> str:
    value = " ".join(value.split())
    if not value:
        raise ValidationError(f"{field_name}은(는) 비워 둘 수 없습니다.")
    if "_" in value:
        raise ValidationError(f"{field_name}에는 '_' 문자를 사용할 수 없습니다.")
    return value


def normalize_surgery(value: str) -> str:
    """수술명에 입력된 스페이스를 콤마로 통일한다 (예: '쌍꺼풀 코성형' -> '쌍꺼풀,코성형')."""
    value = re.sub(r"\s*,\s*", ",", value)
    value = re.sub(r"\s+", ",", value)
    return value


def make_qr(payload: str, output_path: Path) -> None:
    try:
        import qrcode
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        raise RuntimeError(
            "qrcode와 Pillow가 필요합니다. 아래 명령을 한 번 실행하세요:\n"
            "  .venv/bin/python -m pip install qrcode[pil]"
        ) from exc

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


def create_session(promo: str, name: str, surgery: str, surgery_date: str, output_dir: Path) -> Path:
    payload = (
        f"{name}_{surgery}_{surgery_date}"
        if promo == "일반"
        else f"{promo}_{name}_{surgery}_{surgery_date}"
    )
    filename_base = safe_filename(payload)
    qr_path = output_dir / f"{filename_base}.png"
    make_qr(payload, qr_path)
    append_record(promo, name, surgery, surgery_date)
    return qr_path


def find_patient_photos(base_dir: Path, promo: str, name: str) -> list[Path]:
    """base_dir 아래에서 파일명이 이름_(일반) 또는 홍보여부_이름_(HP/HT)으로
    시작하는 이미지 파일(QR 생성 시 저장된 png 등)을 재귀적으로 찾아 반환한다."""
    if not base_dir.is_dir():
        return []
    prefix = f"{name}_" if promo == "일반" else f"{promo}_{name}_"
    return sorted(
        path for path in base_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS and path.name.startswith(prefix)
    )


def resize_to_fit(image, box_width: int, box_height: int):
    """이미지를 비율을 유지한 채 box 크기에 꽉 차도록 늘리거나 줄인다.
    PIL의 thumbnail()과 달리 원본보다 확대도 한다 (미리보기 영역이 커지면
    사진도 같이 커지게 하기 위함)."""
    from PIL import Image as PILImage

    src_w, src_h = image.size
    if src_w <= 0 or src_h <= 0 or box_width <= 0 or box_height <= 0:
        return image
    scale = min(box_width / src_w, box_height / src_h)
    new_size = (max(int(src_w * scale), 1), max(int(src_h * scale), 1))
    return image.resize(new_size, PILImage.LANCZOS)


UI_FONT = ("Malgun Gothic", 10)
UI_FONT_BOLD = ("Malgun Gothic", 10, "bold")
TITLE_FONT = ("Malgun Gothic", 15, "bold")
HANGEUL_CHARSET = 0x81

BG_COLOR = "#f4f6fa"
CARD_COLOR = "#ffffff"
BORDER_COLOR = "#d7dbe3"
TEXT_COLOR = "#1f2937"
MUTED_COLOR = "#6b7280"
ACCENT_COLOR = "#2f6fed"
ACCENT_HOVER = "#255ac2"
SUCCESS_COLOR = "#1a7f37"
ERROR_COLOR = "#c0392b"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}
DEFAULT_WORK_DIR = Path(os.environ.get("WORK_DIR", str(BASE_DIR / "qr_codes")))

INITIAL_WINDOW_SIZE = (1500, 1050)

ENTRY_KWARGS = dict(
    font=UI_FONT, relief="flat", highlightthickness=1,
    highlightbackground=BORDER_COLOR, highlightcolor=ACCENT_COLOR,
    bg=CARD_COLOR, fg=TEXT_COLOR, insertbackground=TEXT_COLOR,
)


class _LogFontW(ctypes.Structure):
    _fields_ = [
        ("lfHeight", ctypes.c_long),
        ("lfWidth", ctypes.c_long),
        ("lfEscapement", ctypes.c_long),
        ("lfOrientation", ctypes.c_long),
        ("lfWeight", ctypes.c_long),
        ("lfItalic", ctypes.c_byte),
        ("lfUnderline", ctypes.c_byte),
        ("lfStrikeOut", ctypes.c_byte),
        ("lfCharSet", ctypes.c_byte),
        ("lfOutPrecision", ctypes.c_byte),
        ("lfClipPrecision", ctypes.c_byte),
        ("lfQuality", ctypes.c_byte),
        ("lfPitchAndFamily", ctypes.c_byte),
        ("lfFaceName", ctypes.c_wchar * 32),
    ]


def _sync_ime_composition_font(widget: tk.Entry, family: str, point_size: int) -> None:
    """IME는 위젯 폰트를 자동으로 따라오지 않고 자체 기본 폰트로 조합 중 글자를
    그리므로, 조합 글자와 확정된 글자의 크기가 어긋난다. 포커스를 받을 때마다
    조합 폰트를 위젯 폰트와 같은 크기로 직접 지정해 맞춘다."""
    try:
        imm32 = ctypes.windll.imm32
    except (AttributeError, OSError):
        return

    def apply(_event=None):
        hwnd = widget.winfo_id()
        dpi = widget.winfo_fpixels("1i")
        logfont = _LogFontW()
        logfont.lfHeight = -round(point_size * dpi / 72)
        logfont.lfWeight = 400
        logfont.lfCharSet = HANGEUL_CHARSET
        logfont.lfFaceName = family
        himc = imm32.ImmGetContext(hwnd)
        if himc:
            imm32.ImmSetCompositionFontW(himc, ctypes.byref(logfont))
            imm32.ImmReleaseContext(hwnd, himc)

    widget.bind("<FocusIn>", apply, add="+")
    widget.after(50, apply)


class QRGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("환자 촬영 QR 도구")
        self.resizable(True, True)

        # QR 저장 폴더와 사진 검색 폴더는 실제로 같은 폴더를 쓰므로 하나의 값을
        # 공유하고, 마지막으로 쓴 값을 config 파일에 저장해 다음 실행 때도 불러온다.
        self._config = load_config()
        self._work_dir_save_after_id = None
        self.work_dir_var = tk.StringVar(value=self._config.get("work_dir") or str(DEFAULT_WORK_DIR))
        self.work_dir_var.trace_add("write", self._schedule_save_work_dir)

        self._apply_style()
        self._build_ui()
        self.minsize(*INITIAL_WINDOW_SIZE)
        self._center_window()

    def _schedule_save_work_dir(self, *_args) -> None:
        if self._work_dir_save_after_id:
            self.after_cancel(self._work_dir_save_after_id)
        self._work_dir_save_after_id = self.after(500, self._save_work_dir)

    def _save_work_dir(self) -> None:
        self._work_dir_save_after_id = None
        self._config["work_dir"] = self.work_dir_var.get().strip()
        save_config(self._config)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew")

        generate_tab = ttk.Frame(notebook)
        search_tab = ttk.Frame(notebook)
        notebook.add(generate_tab, text="QR 생성")
        notebook.add(search_tab, text="QR 찾기")

        self._build_generate_tab(generate_tab)
        self._build_search_tab(search_tab)

    def _apply_style(self) -> None:
        # 입력창과 QR 라벨(malgun.ttf)의 한글 렌더링을 통일한다.
        self.configure(bg=BG_COLOR)
        self.option_add("*Font", UI_FONT)
        self.option_add("*TCombobox*Listbox.font", UI_FONT)
        self.option_add("*TCombobox*Listbox.background", CARD_COLOR)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=UI_FONT, background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("TFrame", background=BG_COLOR)
        style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("Header.TLabel", font=TITLE_FONT, background=BG_COLOR)
        style.configure("Muted.TLabel", background=BG_COLOR, foreground=MUTED_COLOR)
        style.configure("Success.TLabel", background=BG_COLOR, foreground=SUCCESS_COLOR)
        style.configure("Error.TLabel", background=BG_COLOR, foreground=ERROR_COLOR)

        style.configure("TLabelframe", background=BG_COLOR, bordercolor=BORDER_COLOR)
        style.configure("TLabelframe.Label", background=BG_COLOR, foreground=MUTED_COLOR,
                         font=UI_FONT_BOLD)

        # 드롭다운/달력 버튼이 흰 배경에 묻혀 안 보이지 않도록 버튼 영역에
        # 대비되는 회색 배경과 큼직한 화살표를 준다.
        style.configure("TCombobox", fieldbackground=CARD_COLOR, background="#e2e6ee",
                         bordercolor=BORDER_COLOR, arrowsize=22, arrowcolor=TEXT_COLOR,
                         padding=(8, 6))
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", CARD_COLOR)],
            # clam 테마 기본값은 readonly+focus 상태에서 글자색을 흰색으로 바꾸는데,
            # 배경도 흰색이라 선택한 값이 안 보이게 된다. 글자색을 고정한다.
            foreground=[("readonly", TEXT_COLOR)],
            background=[("active", "#cdd4e0"), ("pressed", "#cdd4e0")],
            arrowcolor=[("disabled", MUTED_COLOR)],
        )

        style.configure("Accent.TButton", font=UI_FONT_BOLD, background=ACCENT_COLOR,
                         foreground="white", borderwidth=0, padding=(12, 10))
        style.map("Accent.TButton", background=[("active", ACCENT_HOVER), ("pressed", ACCENT_HOVER)])

        style.configure("Secondary.TButton", font=UI_FONT, background=CARD_COLOR,
                         foreground=TEXT_COLOR, bordercolor=BORDER_COLOR, padding=(10, 6))
        style.map("Secondary.TButton", background=[("active", "#eef1f6")])

        style.configure("TNotebook", background=BG_COLOR, bordercolor=BORDER_COLOR)
        style.configure("TNotebook.Tab", background=BG_COLOR, foreground=MUTED_COLOR,
                         font=UI_FONT_BOLD, padding=(18, 10))
        style.map("TNotebook.Tab", background=[("selected", CARD_COLOR)],
                  foreground=[("selected", TEXT_COLOR)])

    def _build_generate_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        outer = ttk.Frame(parent, padding=24)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # 왼쪽: 입력 폼 (위쪽에 붙이고 세로로는 늘리지 않는다)
        left = ttk.Frame(outer)
        left.grid(row=0, column=0, sticky="new", padx=(0, 24))
        left.columnconfigure(0, weight=1)

        ttk.Label(left, text="환자 촬영 QR 생성기", style="Header.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(left, text="환자 정보를 입력하고 촬영 시작 QR을 생성하세요.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 16))

        info_frame = ttk.LabelFrame(left, text="환자 정보", padding=(16, 14))
        info_frame.grid(row=2, column=0, sticky="ew")
        info_frame.columnconfigure(1, weight=1)
        field_pad = {"pady": 6}

        ttk.Label(info_frame, text="홍보여부").grid(row=0, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.promo_var = tk.StringVar(value="일반")
        promo_combo = ttk.Combobox(info_frame, textvariable=self.promo_var, values=PROMO_CHOICES,
                                    state="readonly", width=26)
        promo_combo.grid(row=0, column=1, sticky="ew", ipady=3, **field_pad)

        # 일반 ttk.Entry는 IME 조합(한글 입력) 중 폰트가 위젯 폰트와
        # 동기화되지 않는 Tk/Windows 문제가 있어 classic tk.Entry를 사용한다.
        ttk.Label(info_frame, text="이름").grid(row=1, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.name_entry = tk.Entry(info_frame, width=29, **ENTRY_KWARGS)
        self.name_entry.grid(row=1, column=1, sticky="ew", **field_pad)
        _sync_ime_composition_font(self.name_entry, *UI_FONT)

        ttk.Label(info_frame, text="수술명").grid(row=2, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.surgery_entry = tk.Entry(info_frame, width=29, **ENTRY_KWARGS)
        self.surgery_entry.grid(row=2, column=1, sticky="ew", **field_pad)
        _sync_ime_composition_font(self.surgery_entry, *UI_FONT)

        ttk.Label(info_frame, text="수술날짜").grid(row=3, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.date_entry = DateEntry(
            info_frame, width=27, font=UI_FONT, date_pattern="yyyy-mm-dd",
            state="readonly", borderwidth=1, background=ACCENT_COLOR, foreground="white",
            bordercolor=BORDER_COLOR, headersbackground=ACCENT_COLOR, headersforeground="white",
            normalbackground=CARD_COLOR, normalforeground=TEXT_COLOR,
            weekendbackground=CARD_COLOR, weekendforeground=TEXT_COLOR,
            selectbackground=ACCENT_COLOR, selectforeground="white",
        )
        self.date_entry.grid(row=3, column=1, sticky="ew", ipady=3, **field_pad)
        # 화살표 버튼뿐 아니라 날짜 표시 부분을 눌러도 달력이 뜨게 한다.
        self.date_entry.bind("<Button-1>", self._toggle_calendar, add="+")

        submit_btn = ttk.Button(left, text="QR 생성", style="Accent.TButton", command=self._on_submit)
        submit_btn.grid(row=3, column=0, sticky="ew", pady=(20, 8))

        folder_frame = ttk.LabelFrame(left, text="저장 위치", padding=(16, 14))
        folder_frame.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        folder_frame.columnconfigure(0, weight=1)

        self.output_dir_entry = tk.Entry(folder_frame, textvariable=self.work_dir_var, **ENTRY_KWARGS)
        self.output_dir_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        ttk.Button(folder_frame, text="찾아보기", style="Secondary.TButton",
                   command=self._browse_work_dir).grid(row=0, column=1)

        action_btn_frame = ttk.Frame(folder_frame)
        action_btn_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(action_btn_frame, text="폴더 열기", style="Secondary.TButton",
                   command=self._open_work_dir).pack(side="left", padx=(0, 6))
        ttk.Button(action_btn_frame, text="기록 파일 열기", style="Secondary.TButton",
                   command=self._open_records).pack(side="left")

        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(left, textvariable=self.status_var, style="Muted.TLabel",
                                       anchor="center")
        self.status_label.grid(row=5, column=0, sticky="ew")

        # 오른쪽: 미리보기 (창을 늘리면 이쪽이 늘어난다)
        preview_frame = ttk.LabelFrame(outer, text="QR 미리보기", padding=(16, 14))
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self._preview_source = None
        self._preview_photo = None
        self._preview_resize_after_id = None
        self.preview_label = ttk.Label(preview_frame, text="생성된 QR이 여기에 표시됩니다.",
                                        style="Muted.TLabel", anchor="center", justify="center")
        self.preview_label.grid(row=0, column=0, sticky="nsew")
        self.preview_label.bind("<Configure>", lambda _e: self._schedule_render_qr_preview())

        self.name_entry.focus_set()

    def _toggle_calendar(self, event) -> None:
        # 화살표 버튼 클릭은 tkcalendar가 자체적으로 처리하므로 중복 토글을 피한다.
        if self.date_entry.identify(event.x, event.y) != self.date_entry._downarrow_name:
            self.date_entry.drop_down()

    def _show_preview(self, image_path: Path) -> None:
        try:
            from PIL import Image
        except ImportError:
            return
        image = Image.open(image_path)
        image.load()
        self._preview_source = image
        self._render_qr_preview()

    def _schedule_render_qr_preview(self) -> None:
        if self._preview_resize_after_id:
            self.after_cancel(self._preview_resize_after_id)
        self._preview_resize_after_id = self.after(80, self._render_qr_preview)

    def _render_qr_preview(self) -> None:
        self._preview_resize_after_id = None
        if self._preview_source is None:
            return
        try:
            from PIL import ImageTk
        except ImportError:
            return
        width = max(self.preview_label.winfo_width() - 8, 60)
        height = max(self.preview_label.winfo_height() - 8, 60)
        fitted = resize_to_fit(self._preview_source, width, height)
        photo = ImageTk.PhotoImage(fitted)
        self._preview_photo = photo  # 참조를 유지하지 않으면 가비지 컬렉션으로 사라진다.
        self.preview_label.configure(image=photo, text="")

    def _center_window(self) -> None:
        # 처음 만든 QR이 작아 보이지 않도록, 내용에 맞춘 최소 크기 대신
        # 미리보기 영역이 넉넉하게 확보되는 고정 크기로 창을 띄운다.
        self.update_idletasks()
        width, height = INITIAL_WINDOW_SIZE
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 3
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _set_status(self, text: str, *, tone: str = "muted") -> None:
        self.status_var.set(text)
        style_by_tone = {"muted": "Muted.TLabel", "success": "Success.TLabel", "error": "Error.TLabel"}
        self.status_label.configure(style=style_by_tone[tone])

    def _browse_work_dir(self) -> None:
        initial = self.work_dir_var.get().strip() or str(DEFAULT_WORK_DIR)
        selected = filedialog.askdirectory(initialdir=initial, title="폴더 선택")
        if selected:
            self.work_dir_var.set(selected)

    def _open_work_dir(self) -> None:
        path = Path(self.work_dir_var.get().strip() or DEFAULT_WORK_DIR)
        try:
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
        except OSError as exc:
            self._set_status(str(exc), tone="error")
            messagebox.showerror("폴더 열기 실패", str(exc))

    def _open_records(self) -> None:
        if not RECORDS_PATH.exists():
            self._set_status("아직 생성된 기록이 없습니다.", tone="error")
            return
        try:
            os.startfile(str(RECORDS_PATH))
        except OSError as exc:
            self._set_status(str(exc), tone="error")
            messagebox.showerror("파일 열기 실패", str(exc))

    def _on_submit(self) -> None:
        try:
            promo = self.promo_var.get()
            name = clean_field(self.name_entry.get(), "이름")
            surgery = normalize_surgery(clean_field(self.surgery_entry.get(), "수술명"))
            surgery_date = self.date_entry.get_date().strftime("%y%m%d")
            output_dir_str = self.work_dir_var.get().strip()
            if not output_dir_str:
                raise ValidationError("저장 폴더를 지정하세요.")
            output_dir = Path(output_dir_str)
        except ValidationError as exc:
            self._set_status(str(exc), tone="error")
            messagebox.showerror("입력 오류", str(exc))
            return

        try:
            qr_path = create_session(promo, name, surgery, surgery_date, output_dir)
        except RuntimeError as exc:
            self._set_status(str(exc), tone="error")
            messagebox.showerror("생성 실패", str(exc))
            return

        self._set_status(f"생성 완료: {qr_path.name}", tone="success")
        self._show_preview(qr_path)

        self.name_entry.delete(0, tk.END)
        self.surgery_entry.delete(0, tk.END)
        self.date_entry.set_date(date.today())
        self.name_entry.focus_set()

    def _build_search_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        outer = ttk.Frame(parent, padding=24)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # 왼쪽: 검색 조건 + 결과 목록 (위쪽에 붙이고 남는 세로 공간은 목록이 채운다)
        left = ttk.Frame(outer)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 24))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(3, weight=1)

        ttk.Label(left, text="환자 사진 찾기", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="홍보여부와 이름으로 사진 폴더를 검색합니다.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 16))

        search_frame = ttk.LabelFrame(left, text="검색 조건", padding=(16, 14))
        search_frame.grid(row=2, column=0, sticky="ew")
        search_frame.columnconfigure(1, weight=1)
        field_pad = {"pady": 6}

        ttk.Label(search_frame, text="홍보여부").grid(row=0, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.search_promo_var = tk.StringVar(value="일반")
        ttk.Combobox(search_frame, textvariable=self.search_promo_var, values=PROMO_CHOICES,
                     state="readonly", width=20).grid(row=0, column=1, sticky="ew", ipady=3, **field_pad)

        ttk.Label(search_frame, text="이름").grid(row=1, column=0, sticky="e", padx=(0, 12), **field_pad)
        self.search_name_entry = tk.Entry(search_frame, width=26, **ENTRY_KWARGS)
        self.search_name_entry.grid(row=1, column=1, sticky="ew", **field_pad)
        _sync_ime_composition_font(self.search_name_entry, *UI_FONT)
        self.search_name_entry.bind("<Return>", lambda _e: self._on_search())

        ttk.Button(search_frame, text="검색", style="Accent.TButton", command=self._on_search).grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(6, 12))

        ttk.Label(search_frame, text="사진 폴더").grid(row=3, column=0, sticky="e", padx=(0, 12), **field_pad)
        dir_row = ttk.Frame(search_frame)
        dir_row.grid(row=3, column=1, sticky="ew", **field_pad)
        dir_row.columnconfigure(0, weight=1)
        # QR 저장 폴더와 동일한 폴더를 쓰므로 work_dir_var를 그대로 공유한다.
        tk.Entry(dir_row, textvariable=self.work_dir_var, **ENTRY_KWARGS).grid(
            row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(dir_row, text="찾아보기", style="Secondary.TButton",
                   command=self._browse_work_dir).grid(row=0, column=1)

        list_frame = ttk.LabelFrame(left, text="검색 결과", padding=(10, 10))
        list_frame.grid(row=3, column=0, sticky="nsew", pady=(16, 0))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.result_listbox = tk.Listbox(
            list_frame, width=34, font=UI_FONT, activestyle="none",
            relief="flat", highlightthickness=1, highlightbackground=BORDER_COLOR,
            bg=CARD_COLOR, fg=TEXT_COLOR, selectbackground=ACCENT_COLOR, selectforeground="white",
        )
        self.result_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.result_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_listbox.configure(yscrollcommand=scrollbar.set)
        self.result_listbox.bind("<<ListboxSelect>>", self._on_result_select)

        self.search_status_var = tk.StringVar(value="")
        ttk.Label(left, textvariable=self.search_status_var, style="Muted.TLabel").grid(
            row=4, column=0, sticky="w", pady=(8, 0))

        # 오른쪽: 미리보기 (창을 늘리면 이쪽이 늘어나고, 세로 전체를 차지한다)
        preview_frame = ttk.LabelFrame(outer, text="사진 미리보기", padding=(16, 14))
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self._search_photos: list[Path] = []
        self._search_preview_source = None
        self._search_preview_image = None
        self._search_preview_resize_after_id = None
        self._search_preview_placeholder = "검색 결과에서 사진을 클릭하면 여기에 크게 표시됩니다."
        self.search_preview_label = ttk.Label(
            preview_frame, text=self._search_preview_placeholder,
            style="Muted.TLabel", anchor="center", justify="center",
        )
        self.search_preview_label.grid(row=0, column=0, sticky="nsew")
        self.search_preview_label.bind("<Configure>", lambda _e: self._schedule_render_search_preview())

    def _on_search(self) -> None:
        promo = self.search_promo_var.get()
        name = " ".join(self.search_name_entry.get().split())
        if not name:
            self.search_status_var.set("이름을 입력하세요.")
            return

        base_dir = Path(self.work_dir_var.get().strip() or DEFAULT_WORK_DIR)
        photos = find_patient_photos(base_dir, promo, name)
        self._search_photos = photos

        self.result_listbox.delete(0, tk.END)
        for photo in photos:
            self.result_listbox.insert(tk.END, str(photo.relative_to(base_dir)))

        self._search_preview_source = None
        self.search_preview_label.configure(image="", text=self._search_preview_placeholder)
        self._search_preview_image = None

        if not photos:
            self.search_status_var.set(f"'{promo} {name}'의 사진을 찾지 못했습니다.")
        else:
            self.search_status_var.set(f"{len(photos)}장을 찾았습니다.")

    def _on_result_select(self, _event=None) -> None:
        selection = self.result_listbox.curselection()
        if not selection:
            return
        photo_path = self._search_photos[selection[0]]

        try:
            from PIL import Image
        except ImportError:
            self.search_status_var.set("사진 미리보기에는 Pillow가 필요합니다.")
            return

        try:
            image = Image.open(photo_path)
            image.load()
        except OSError as exc:
            self.search_status_var.set(f"사진을 열지 못했습니다: {exc}")
            return

        self._search_preview_source = image
        self._render_search_preview()

    def _schedule_render_search_preview(self) -> None:
        if self._search_preview_resize_after_id:
            self.after_cancel(self._search_preview_resize_after_id)
        self._search_preview_resize_after_id = self.after(80, self._render_search_preview)

    def _render_search_preview(self) -> None:
        self._search_preview_resize_after_id = None
        if self._search_preview_source is None:
            return
        try:
            from PIL import ImageTk
        except ImportError:
            return
        width = max(self.search_preview_label.winfo_width() - 8, 60)
        height = max(self.search_preview_label.winfo_height() - 8, 60)
        fitted = resize_to_fit(self._search_preview_source, width, height)
        photo = ImageTk.PhotoImage(fitted)
        self._search_preview_image = photo  # 참조를 유지하지 않으면 가비지 컬렉션으로 사라진다.
        self.search_preview_label.configure(image=photo, text="")


def _enable_windows_dpi_awareness() -> None:
    """DPI 인식을 켜지 않으면 Windows가 창을 통째로 비트맵 확대/축소해서
    표시하는데, IME 조합 중인 글자는 이 축소 대상이 아니라서 완성된
    글자보다 훨씬 크게 보인다."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def main() -> None:
    _enable_windows_dpi_awareness()
    if DateEntry is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "실행 불가",
            "tkcalendar가 필요합니다. 아래 명령을 한 번 실행하세요:\n"
            "  .venv\\Scripts\\python.exe -m pip install tkcalendar",
        )
        root.destroy()
        return
    app = QRGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
