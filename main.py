# -*- coding: utf-8 -*-
"""
AppPrintPO — Import Excel (Mã đơn, SKU, Số lượng, Mẫu, Loại, Ngang, Cao),
tên in trên PDF: Loại - Ngang x Cao - Mẫu. Khổ trang = 10×15 cm × 1,2.
"""
from __future__ import annotations

import os
import tkinter as tk
from xml.sax.saxutils import escape
from datetime import date
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# Khổ gốc 10×15 cm, tăng 20 % → 12×18 cm (dọc)
PAGE_SCALE = 1.2
PAGE_W, PAGE_H = 10 * PAGE_SCALE * cm, 15 * PAGE_SCALE * cm
# Cỡ chữ dòng tên sản phẩm (Loại - Ngang x Cao - Mẫu) trong bảng — lớn hơn so với trước
TEN_SP_FONT_MULT = 1.38

REQUIRED_COLS = {
    "ma_don": ["Mã đơn", "Mã Đơn", "MA DON", "Mã đơn hàng", "Order"],
    "sku": ["SKU", "Sku", "sku", "Mã SP"],
    "sl": ["Số lượng", "So luong", "SL", "sl", "Qty", "Quantity"],
    "mau": ["Mẫu", "Mau", "mẫu", "MAU", "Pattern"],
    "loai": ["Loại", "Loai", "loại", "LOAI", "Loiaj", "Type"],
    "ngang": ["Ngang", "ngang", "NGANG", "Rộng", "Width", "W"],
    "cao": ["Cao", "cao", "CAO", "Height", "H"],
}


def _find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    cols = {str(c).strip(): c for c in df.columns}
    for a in aliases:
        if a in cols:
            return cols[a]
    lower = {str(c).strip().lower(): c for c in df.columns}
    for a in aliases:
        if a.lower() in lower:
            return lower[a.lower()]
    return None


def normalize_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, str | None]:
    """Đổi tên cột chuẩn. Trả về (df, lỗi)."""
    mapping = {}
    for key, aliases in REQUIRED_COLS.items():
        col = _find_col(df, aliases)
        if col is None:
            return df, f"Thiếu cột: cần một trong {aliases}"
        mapping[col] = key
    out = df.rename(columns=mapping)
    return out[list(REQUIRED_COLS.keys())], None


def _cell_str(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        f = float(val)
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        return str(f).rstrip("0").rstrip(".")
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def _parse_number(val) -> float | None:
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = str(val).strip().replace(",", ".").replace(" ", "")
    if not s or s.lower() == "nan":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _format_ngang_in_ten(val) -> str:
    """Ngang in trong tên SP: số nguyên 2 → 2m; thập phân 2.5 → 2m5 (m thay dấu phẩy/thập phân)."""
    f = _parse_number(val)
    if f is None:
        return _cell_str(val)
    if abs(f - round(f)) < 1e-9:
        return f"{int(round(f))}m"
    s = f"{f:.10f}".rstrip("0").rstrip(".")
    if "." in s:
        a, b = s.split(".", 1)
        return f"{a}m{b}"
    return s


def _format_cao_in_ten(val) -> str:
    """Kích thước cao: 0.7 → 0.7m, 45 → 45m (chữ m đơn vị cuối)."""
    f = _parse_number(val)
    if f is None:
        t = _cell_str(val)
        return f"{t}m" if t else ""
    if abs(f - round(f)) < 1e-9:
        return f"{int(round(f))}m"
    s = f"{f:.10f}".rstrip("0").rstrip(".")
    return f"{s}m"


def build_ten_hien_thi(r: pd.Series) -> str:
    """Tên SP: Loại - Ngang x Cao - Mẫu; ngang/cao format theo quy ước in."""
    loai = _cell_str(r["loai"])
    ngang_d = _format_ngang_in_ten(r["ngang"])
    cao_d = _format_cao_in_ten(r["cao"])
    mau = _cell_str(r["mau"])
    parts: list[str] = []
    if loai:
        parts.append(loai)
    if ngang_d and cao_d:
        parts.append(f"{ngang_d} x {cao_d}")
    elif ngang_d:
        parts.append(ngang_d)
    elif cao_d:
        parts.append(cao_d)
    if mau:
        parts.append(mau)
    return " - ".join(parts) if parts else "—"


def register_vietnamese_font() -> str:
    """Đăng ký font hỗ trợ tiếng Việt (Arial trên Windows)."""
    candidates = [
        os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "arial.ttf"),
        os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "tahoma.ttf"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            name = "AppPrintPOFont"
            try:
                pdfmetrics.registerFont(TTFont(name, path))
            except Exception:
                continue
            return name
    return "Helvetica"


def build_order_page(
    story: list,
    styles: dict,
    font_name: str,
    ma_don: str,
    ngay_str: str,
    ncc: str,
    dia_chi: str,
    rows: list[tuple[str, str, int]],
) -> None:
    """Một đơn hàng — thêm nội dung vào story."""
    fs = PAGE_SCALE
    title_style = ParagraphStyle(
        "title",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=round(11 * fs),
        alignment=TA_CENTER,
        spaceAfter=round(6 * fs),
        leading=round(13 * fs),
    )
    normal = ParagraphStyle(
        "n",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=round(8 * fs),
        alignment=TA_LEFT,
        leading=round(10 * fs),
    )
    small = ParagraphStyle(
        "s",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=round(7 * fs),
        alignment=TA_LEFT,
        leading=round(9 * fs),
    )
    ten_sp = ParagraphStyle(
        "ten_sp",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=max(10, round(7 * fs * TEN_SP_FONT_MULT)),
        alignment=TA_LEFT,
        leading=max(13, round(9 * fs * TEN_SP_FONT_MULT)),
    )

    # Header trái
    story.append(Paragraph(f"<b>Mã đơn đặt:</b> {escape(ma_don)}", normal))
    story.append(Paragraph(f"<b>Ngày tạo:</b> {escape(ngay_str)}", normal))
    story.append(Spacer(1, 4))
    story.append(Paragraph("<b>Đơn đặt hàng nhập</b>", title_style))
    story.append(Spacer(1, 4))
    if ncc.strip():
        story.append(Paragraph(f"<b>Nhà cung cấp:</b> {escape(ncc)}", normal))
    if dia_chi.strip():
        story.append(Paragraph(f"<b>Địa chỉ NCC:</b> {escape(dia_chi)}", normal))
    story.append(Spacer(1, 6))

    # Bảng: STT | Tên sản phẩm | SL
    table_data: list[list] = [
        [
            Paragraph("<b>STT</b>", normal),
            Paragraph("<b>Tên sản phẩm</b>", normal),
            Paragraph("<b>Số lượng</b>", normal),
        ]
    ]
    total = 0
    for i, (ten, sku, sl) in enumerate(rows, start=1):
        total += int(sl)
        sku_pt = max(8, round(6 * PAGE_SCALE * 1.12))
        name_block = (
            f"{escape(ten)}<br/><font size='{sku_pt}' color='#333333'>SKU: {escape(sku)}</font>"
        )
        table_data.append(
            [
                Paragraph(str(i), normal),
                Paragraph(name_block, ten_sp),
                Paragraph(str(sl), ParagraphStyle("r", parent=normal, alignment=TA_RIGHT)),
            ]
        )

    # Dòng tổng
    table_data.append(
        [
            "",
            Paragraph("<b>Số lượng</b>", ParagraphStyle("lbl", parent=normal, alignment=TA_RIGHT)),
            Paragraph(f"<b>{total}</b>", ParagraphStyle("t", parent=normal, alignment=TA_RIGHT)),
        ]
    )

    u = PAGE_SCALE * cm
    col_widths = [0.9 * u, PAGE_W - 2.5 * u - 1.2 * u, 1.2 * u]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), round(8 * PAGE_SCALE)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (2, 0), (2, -1), "RIGHT"),
                ("GRID", (0, 0), (-1, -2), 0.5, colors.black),
                ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.black),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(t)


def export_pdf(
    df: pd.DataFrame,
    out_path: str,
    ngay_str: str,
    ncc: str,
    dia_chi: str,
) -> None:
    font_name = register_vietnamese_font()
    styles = getSampleStyleSheet()

    def on_page(canv, doc):
        canv.saveState()
        canv.setFont(font_name, round(7 * PAGE_SCALE))
        canv.restoreState()

    doc = SimpleDocTemplate(
        out_path,
        pagesize=portrait((PAGE_W, PAGE_H)),
        leftMargin=0.5 * PAGE_SCALE * cm,
        rightMargin=0.5 * PAGE_SCALE * cm,
        topMargin=0.4 * PAGE_SCALE * cm,
        bottomMargin=0.4 * PAGE_SCALE * cm,
    )
    story: list = []

    for ma_don, g in df.groupby("ma_don", sort=True):
        ma_str = str(ma_don).strip()
        rows_list: list[tuple[str, str, int]] = []
        for _, r in g.iterrows():
            ten = build_ten_hien_thi(r)
            sku = _cell_str(r["sku"])
            try:
                sl = int(float(r["sl"]))
            except (TypeError, ValueError):
                sl = 0
            rows_list.append((ten, sku, sl))

        build_order_page(story, styles, font_name, ma_str, ngay_str, ncc, dia_chi, rows_list)
        story.append(PageBreak())

    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("AppPrintPO — Excel → PDF (12×18 cm)")
        self.geometry("520x320")
        self.excel_path: str | None = None

        f = ttk.Frame(self, padding=12)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            f,
            text="File Excel: Mã đơn, SKU, Số lượng, Mẫu, Loại, Ngang, Cao — tên in: Loại - Ngang×Cao - Mẫu",
        ).grid(row=0, column=0, sticky=tk.W)
        self.lbl_file = ttk.Label(f, text="(chưa chọn)", foreground="#666")
        self.lbl_file.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))
        ttk.Button(f, text="Chọn file .xlsx", command=self.pick_file).grid(row=2, column=0, sticky=tk.W)

        ttk.Label(f, text="Ngày in trên phiếu (dd/mm/yyyy):").grid(row=3, column=0, sticky=tk.W, pady=(12, 0))
        today = date.today().strftime("%d/%m/%Y")
        self.var_ngay = tk.StringVar(value=today)
        ttk.Entry(f, textvariable=self.var_ngay, width=18).grid(row=4, column=0, sticky=tk.W)

        ttk.Label(f, text="Nhà cung cấp (tùy chọn):").grid(row=5, column=0, sticky=tk.W, pady=(8, 0))
        self.var_ncc = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_ncc, width=50).grid(row=6, column=0, columnspan=2, sticky=tk.W)

        ttk.Label(f, text="Địa chỉ NCC (tùy chọn):").grid(row=7, column=0, sticky=tk.W, pady=(8, 0))
        self.var_dc = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_dc, width=50).grid(row=8, column=0, columnspan=2, sticky=tk.W)

        ttk.Button(f, text="Xuất PDF", command=self.export).grid(row=9, column=0, sticky=tk.W, pady=16)

    def pick_file(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            self.excel_path = p
            self.lbl_file.config(text=p)

    def export(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file Excel.")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            title="Lưu PDF",
        )
        if not out:
            return
        try:
            df = pd.read_excel(self.excel_path, engine="openpyxl")
            df = df.dropna(how="all")
            df, err = normalize_columns(df)
            if err:
                messagebox.showerror("Lỗi cột", err)
                return
            df = df[df["ma_don"].notna() & (df["ma_don"].astype(str).str.strip() != "")]
            if df.empty:
                messagebox.showerror("Lỗi", "Không có dòng dữ liệu hợp lệ (thiếu Mã đơn).")
                return
            export_pdf(
                df,
                out,
                self.var_ngay.get().strip(),
                self.var_ncc.get(),
                self.var_dc.get(),
            )
            messagebox.showinfo("Xong", f"Đã tạo PDF:\n{out}")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
