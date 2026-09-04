
import io
import json
import re
from datetime import datetime
from copy import copy

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

APP_TITLE = "AI Quản Lý Phòng Trọ"
ROOMS = [f"A{i}" for i in range(1, 9)] + [f"B{i}" for i in range(1, 9)]

DEFAULT_MODEL = "gemini-2.5-flash"

def parse_month_value(value):
    """Return (year, month) from datetime/date or common legacy text such as 01/2022."""
    if isinstance(value, datetime):
        return value.year, value.month
    if isinstance(value, str):
        text = value.strip()
        m = re.fullmatch(r"(\d{1,2})/(\d{4})", text)
        if m:
            return int(m.group(2)), int(m.group(1))
    return None

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏠",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-family: Arial, sans-serif; }
.block-container { max-width: 900px; padding-top: 2rem; padding-bottom: 3rem; }
h1 { font-size: 2.35rem !important; text-align: center; }
.big-note { text-align:center; font-size:1.15rem; color:#555; margin-bottom:1.5rem; }
.stButton > button {
    width: 100%; min-height: 3.6rem; font-size: 1.25rem; font-weight: 700;
    border-radius: 12px;
}
div[data-testid="stFileUploader"] label { font-size: 1.15rem; font-weight: 700; }
div[data-testid="stDataEditor"] { font-size: 1.05rem; }
.warning-box {
    padding: 1rem 1.1rem; border-radius: 12px; background:#fff3cd;
    border:1px solid #ffda6a; font-size:1.05rem;
}
.success-box {
    padding: 1rem 1.1rem; border-radius: 12px; background:#d1e7dd;
    border:1px solid #75b798; font-size:1.05rem;
}
.error-box {
    padding: 1rem 1.1rem; border-radius: 12px; background:#f8d7da;
    border:1px solid #ea868f; font-size:1.05rem;
}
.small { font-size: .95rem; color:#666; }
</style>
""", unsafe_allow_html=True)


class RoomReading(BaseModel):
    room: str = Field(description="Mã phòng, ví dụ A3")
    electricity_new: int | None = Field(default=None, description="Chỉ số điện mới đọc từ ảnh")
    water_new: int | None = Field(default=None, description="Chỉ số nước mới đọc từ ảnh")
    confidence: str = Field(default="medium", description="high, medium hoặc low")
    note: str = Field(default="", description="Ghi chú nếu ảnh khó đọc")


class Extraction(BaseModel):
    month: int = Field(description="Tháng 1-12")
    year: int = Field(description="Năm 4 chữ số")
    rooms: list[RoomReading]


def get_client():
    key = st.secrets.get("GEMINI_API_KEY", None)
    if not key:
        key = st.session_state.get("manual_api_key", None)
    if not key:
        key = __import__("os").environ.get("GEMINI_API_KEY")
    if not key:
        return None
    return genai.Client(api_key=key)


def extract_with_gemini(images, model_name):
    client = get_client()
    if client is None:
        raise RuntimeError("Chưa cấu hình GEMINI_API_KEY.")

    schema = Extraction.model_json_schema()
    system = """
Bạn là AI đọc ảnh chỉ số điện/nước phòng trọ.
Nhiệm vụ: đọc các mã phòng A1-A8, B1-B8 và chỉ số điện mới, nước mới từ ảnh.
Không được đoán số bị mờ. Nếu không chắc chắn, để giá trị null và ghi chú.
Nếu có nhiều ảnh, gộp chúng thành một kết quả.
Tháng/năm phải lấy từ ảnh hoặc ngữ cảnh người dùng; nếu không thấy rõ thì dùng 0 cho tháng/năm và ghi chú.
Chỉ trả JSON theo schema được cung cấp.
"""

    contents = [system]
    for img in images:
        data = img.getvalue()
        mime = img.type or "image/jpeg"
        contents.append(types.Part.from_bytes(data=data, mime_type=mime))

    response = client.models.generate_content(
        model=model_name,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=Extraction,
            thinking_config=types.ThinkingConfig(thinking_level="low"),
        ),
    )
    return Extraction.model_validate_json(response.text)



def room_month_rows(ws, year, month):
    """Find the row whose column A identifies target year/month."""
    for r in range(1, ws.max_row + 1):
        parsed = parse_month_value(ws.cell(r, 1).value)
        if parsed == (year, month):
            return r
    return None


def last_month_row(ws):
    candidates = []
    for r in range(1, ws.max_row + 1):
        parsed = parse_month_value(ws.cell(r, 1).value)
        if parsed:
            candidates.append((r, parsed[0], parsed[1]))
    if not candidates:
        return None
    r, y, m = max(candidates, key=lambda x: (x[1], x[2]))
    return r, datetime(y, m, 1)


def next_month(year, month):
    return (year + 1, 1) if month == 12 else (year, month + 1)


def copy_cell(src, dst):
    """Copy a cell while preserving formatting and translating relative formulas."""
    if src.data_type == "f" and isinstance(src.value, str):
        try:
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
        except Exception:
            dst.value = src.value
    else:
        dst.value = src.value

    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.protection:
        dst.protection = copy(src.protection)
    if src.hyperlink:
        dst._hyperlink = copy(src.hyperlink)
    if src.comment:
        dst.comment = copy(src.comment)


def copy_row_dimensions(ws, src_row, dst_row):
    """Copy visible row properties to the cloned month block."""
    src_dim = ws.row_dimensions[src_row]
    dst_dim = ws.row_dimensions[dst_row]
    if src_dim.height is not None:
        dst_dim.height = src_dim.height
    dst_dim.hidden = src_dim.hidden
    dst_dim.outlineLevel = src_dim.outlineLevel
    dst_dim.collapsed = src_dim.collapsed


def copy_merged_ranges_for_block(ws, src_start, src_end, dst_start):
    shift = dst_start - src_start
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= src_start and rng.max_row <= src_end:
            new_min = rng.min_row + shift
            new_max = rng.max_row + shift
            new_range = f"{ws.cell(new_min, rng.min_col).coordinate}:{ws.cell(new_max, rng.max_col).coordinate}"
            try:
                ws.merge_cells(new_range)
            except Exception:
                pass


def copy_month_block(ws, src_start, dst_start):
    src_end = src_start + 4
    for offset in range(5):
        sr = src_start + offset
        dr = dst_start + offset
        for c in range(1, ws.max_column + 1):
            copy_cell(ws.cell(sr, c), ws.cell(dr, c))
        copy_row_dimensions(ws, sr, dr)
    copy_merged_ranges_for_block(ws, src_start, src_end, dst_start)


def extend_month_block(ws, year, month):
    """
    Create missing month blocks sequentially from the latest available block.
    This supports legacy sheets that store dates as strings like 01/2022.
    """
    existing = room_month_rows(ws, year, month)
    if existing:
        return existing, False

    last = last_month_row(ws)
    if not last:
        return None, False

    current_row, current_date = last
    if str(ws.cell(current_row, 2).value).strip().lower() != "phòng":
        return None, False

    current_y, current_m = current_date.year, current_date.month
    while (current_y, current_m) != (year, month):
        ny, nm = next_month(current_y, current_m)
        dst_start = ws.max_row + 2
        copy_month_block(ws, current_row, dst_start)

        # Use the modern datetime representation for new blocks.
        ws.cell(dst_start, 1).value = datetime(ny, nm, 1)

        current_row = dst_start
        current_y, current_m = ny, nm

    return current_row, True


def validate_workbook(wb):
    required = set(ROOMS + ["báo cáo tổng hợp"])
    missing = sorted(required - set(wb.sheetnames))
    return missing


def get_old_readings(wb, room, start_row):
    ws = wb[room]
    old_e = ws.cell(start_row + 1, 8).value  # H = previous reading
    old_w = ws.cell(start_row + 2, 8).value
    return old_e, old_w

def previous_numeric_reading(ws, target_row, offset):
    """Find the latest numeric I-value for the same utility in an earlier month block."""
    utility_row = target_row + offset
    # Each room month block in this workbook is 5 rows: Phòng/Điện/Nước/Rác/Tổng.
    # Step by 5 so we never accidentally take a value from another utility row.
    for r in range(utility_row - 5, 0, -5):
        v = ws.cell(r, 9).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return v
    return None

def fill_missing_rooms(wb, extraction, records):
    """For rooms absent from the photo, carry forward the latest known I values.
    This makes the month's consumption 0 because the existing Excel formula calculates I-H.
    If there is no prior numeric reading at all, use 0 as a safe numeric baseline and flag it.
    """
    found = {r["Phòng"] for r in records}
    for room in ROOMS:
        if room in found:
            continue
        ws = wb[room]
        row = room_month_rows(ws, extraction.year, extraction.month)
        if row is None:
            row, _ = extend_month_block(ws, extraction.year, extraction.month)
        if row is None:
            raise ValueError(f"{room}: không tìm được tháng {extraction.month:02d}/{extraction.year}.")

        prev_e = previous_numeric_reading(ws, row, 1)
        prev_w = previous_numeric_reading(ws, row, 2)
        # If no previous numeric reading exists, use 0 rather than leaving cells blank.
        e = prev_e if prev_e is not None else 0
        w = prev_w if prev_w is not None else 0
        records.append({
            "Phòng": room,
            "Điện mới": e,
            "Nước mới": w,
            "Điện cũ": ws.cell(row + 1, 8).value,
            "Nước cũ": ws.cell(row + 2, 8).value,
            "Độ tin cậy": "tự giữ số",
            "Trạng thái": "KHÔNG CÓ SỐ LIỆU → GIỮ SỐ THÁNG TRƯỚC",
            "Ghi chú": (
                "Ảnh không có dữ liệu phòng này. Hệ thống tự copy chỉ số gần nhất "
                "để công thức I-H cho kết quả 0."
                + (" Chưa có số cũ nên dùng 0 làm mốc." if prev_e is None or prev_w is None else "")
            ),
        })

    return pd.DataFrame(records).sort_values("Phòng").reset_index(drop=True)


def prepare_review(wb, extraction):
    records = []
    for rr in extraction.rooms:
        room = rr.room.strip().upper()
        if room not in ROOMS:
            continue
        ws = wb[room]
        row = room_month_rows(ws, extraction.year, extraction.month)
        if row is None:
            row, _ = extend_month_block(ws, extraction.year, extraction.month)
        old_e = ws.cell(row + 1, 8).value if row else None
        old_w = ws.cell(row + 2, 8).value if row else None

        status = "OK"
        notes = rr.note or ""
        if rr.electricity_new is None or rr.water_new is None:
            status = "CẦN XÁC NHẬN"
        if isinstance(old_e, (int, float)) and isinstance(rr.electricity_new, int):
            if rr.electricity_new < old_e:
                status = "KIỂM TRA"
                notes += f" Điện mới ({rr.electricity_new}) < chỉ số cũ ({old_e})."
        if isinstance(old_w, (int, float)) and isinstance(rr.water_new, int):
            if rr.water_new < old_w:
                status = "KIỂM TRA"
                notes += f" Nước mới ({rr.water_new}) < chỉ số cũ ({old_w})."

        records.append({
            "Phòng": room,
            "Điện mới": rr.electricity_new,
            "Nước mới": rr.water_new,
            "Điện cũ": old_e,
            "Nước cũ": old_w,
            "Độ tin cậy": rr.confidence,
            "Trạng thái": status,
            "Ghi chú": notes.strip(),
        })
    return pd.DataFrame(records)



def ensure_summary_year(ws, year):
    """Find or append a 13-row year block in 'báo cáo tổng hợp'."""
    total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == f"TỔNG NĂM {year}":
            total_row = r
            break
    if total_row:
        return total_row - 12, total_row

    # Clone the latest existing year block.
    candidates = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("TỔNG NĂM "):
            try:
                y = int(v.split()[-1])
                candidates.append((y, r))
            except Exception:
                pass
    if not candidates:
        raise ValueError("Không tìm thấy mẫu bảng tổng hợp theo năm.")
    _, src_total = max(candidates)
    src_start = src_total - 12
    dst_start = ws.max_row + 2
    for offset in range(13):
        sr, dr = src_start + offset, dst_start + offset
        for c in range(1, ws.max_column + 1):
            copy_cell(ws.cell(sr, c), ws.cell(dr, c))
        copy_row_dimensions(ws, sr, dr)
    # Replace labels and translate total formulas.
    for m in range(1, 13):
        ws.cell(dst_start + m - 1, 1).value = f"Tháng {m:02d}"
    ws.cell(dst_start + 12, 1).value = f"TỔNG NĂM {year}"
    return dst_start, dst_start + 12


def update_summary(ws, wb, year, month):
    start, total = ensure_summary_year(ws, year)
    month_row = start + month - 1
    ws.cell(month_row, 1).value = f"Tháng {month:02d}"

    # B:Q correspond to A1:A8,B1:B8 in this workbook.
    for idx, room in enumerate(ROOMS, start=2):
        room_ws = wb[room]
        row = room_month_rows(room_ws, year, month)
        if row is None:
            raise ValueError(f"Sheet {room} chưa có block {month:02d}/{year}.")
        ws.cell(month_row, idx).value = f"='{room}'!F{row}"

    ws.cell(month_row, 18).value = f"=SUBTOTAL(9,B{month_row}:Q{month_row})"
    ws.cell(month_row, 20).value = f"=R{month_row}-S{month_row}"

    for c in range(2, 18):
        letter = ws.cell(1, c).column_letter
        ws.cell(total, c).value = f"=SUBTOTAL(9,{letter}{start}:{letter}{total-1})"
    ws.cell(total, 18).value = f"=SUM(R{start}:R{total-1})"
    ws.cell(total, 19).value = f"=SUM(S{start}:S{total-1})"
    ws.cell(total, 20).value = f"=SUM(T{start}:T{total-1})"


def update_workbook(input_bytes, extraction, edited_df):
    wb = load_workbook(io.BytesIO(input_bytes), data_only=False)
    missing = validate_workbook(wb)
    if missing:
        raise ValueError("Thiếu sheet bắt buộc: " + ", ".join(missing))

    updated = []
    problems = []

    for _, rec in edited_df.iterrows():
        room = str(rec["Phòng"]).strip().upper()
        e_new = rec["Điện mới"]
        w_new = rec["Nước mới"]

        if room not in ROOMS:
            problems.append(f"{room}: mã phòng không hợp lệ.")
            continue
        if pd.isna(e_new) or pd.isna(w_new):
            problems.append(f"{room}: còn thiếu điện hoặc nước.")
            continue

        e_new = int(e_new)
        w_new = int(w_new)
        ws = wb[room]

        row = room_month_rows(ws, extraction.year, extraction.month)
        if row is None:
            row, created = extend_month_block(ws, extraction.year, extraction.month)
            if row is None:
                problems.append(f"{room}: không tìm được mẫu tháng để tạo {extraction.month}/{extraction.year}.")
                continue

        # Only change the two input cells: column I for electricity and water.
        ws.cell(row + 1, 9).value = e_new
        ws.cell(row + 2, 9).value = w_new
        updated.append(f"{room}: I{row+1}={e_new}, I{row+2}={w_new}")

    # Update the summary sheet using the same room totals.
    update_summary(wb["báo cáo tổng hợp"], wb, extraction.year, extraction.month)

    # Ask Excel to recalculate formulas on open.
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.calculation.calcOnSave = True

    if problems:
        raise ValueError("\n".join(problems))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), updated


def main():
    st.title("🏠 AI QUẢN LÝ PHÒNG TRỌ")
    st.markdown('<div class="big-note">Cập nhật chỉ số điện, nước vào file Excel tự động</div>', unsafe_allow_html=True)

    with st.expander("⚙️ Cài đặt AI", expanded=False):
        st.write("Nếu đã triển khai trên Streamlit Cloud, API key nên đặt trong Secrets với tên `GEMINI_API_KEY`.")
        manual = st.text_input("Gemini API Key (không bắt buộc nếu đã có Secrets)", type="password")
        if manual:
            st.session_state["manual_api_key"] = manual
        model = st.selectbox("Mô hình", [DEFAULT_MODEL, "gemini-3.7-flash"], index=0)

    images = st.file_uploader(
        "📷 1. Chọn ảnh chỉ số điện / nước",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        help="Có thể chọn nhiều ảnh nếu bảng được chụp thành nhiều phần."
    )
    excel = st.file_uploader(
        "📊 2. Chọn file Excel phòng trọ",
        type=["xlsx"],
        accept_multiple_files=False,
        help="Chọn đúng file Excel mà anh/chị muốn cập nhật."
    )

    if images:
        st.success(f"Đã chọn {len(images)} ảnh.")
    if excel:
        st.success(f"Đã chọn: {excel.name}")

    if st.button("▶️ BẮT ĐẦU ĐỌC ẢNH", type="primary", disabled=not (images and excel)):
        try:
            with st.spinner("Đang đọc ảnh và kiểm tra số liệu..."):
                extraction = extract_with_gemini(images, model)

            if not 1 <= extraction.month <= 12 or extraction.year < 2000:
                st.error("AI chưa xác định chắc chắn tháng/năm. Vui lòng ghi rõ tháng/năm trong ảnh hoặc nhập lại bằng ghi chú.")
                return

            input_bytes = excel.getvalue()
            wb = load_workbook(io.BytesIO(input_bytes), data_only=False)
            missing = validate_workbook(wb)
            if missing:
                st.error("File Excel không đúng mẫu. Thiếu: " + ", ".join(missing))
                return

            review = prepare_review(wb, extraction)
            review = fill_missing_rooms(wb, extraction, review.to_dict("records"))

            expected = set(ROOMS)
            found = set(review["Phòng"]) if not review.empty else set()
            missing_rooms = sorted(expected - found)

            st.session_state["input_bytes"] = input_bytes
            st.session_state["extraction"] = extraction.model_dump()
            st.session_state["review"] = review

            st.success(f"AI đã đọc tháng {extraction.month:02d}/{extraction.year}.")
            if missing_rooms:
                st.info("Các phòng không có trong ảnh sẽ tự giữ chỉ số tháng trước để công thức tính ra 0.")

        except Exception as e:
            st.error(f"Không thể xử lý: {e}")

    if "review" in st.session_state:
        st.subheader("🔍 3. Kiểm tra số liệu")
        st.info("Anh/chị có thể sửa trực tiếp ô Điện mới / Nước mới nếu AI đọc sai. Sau đó bấm Xác nhận.")
        edited = st.data_editor(
            st.session_state["review"],
            use_container_width=True,
            hide_index=True,
            disabled=["Phòng", "Điện cũ", "Nước cũ", "Độ tin cậy", "Trạng thái", "Ghi chú"],
            column_config={
                "Điện mới": st.column_config.NumberColumn("Điện mới", min_value=0, step=1),
                "Nước mới": st.column_config.NumberColumn("Nước mới", min_value=0, step=1),
            },
            key="review_editor",
        )

        bad = edited[(edited["Điện mới"].isna()) | (edited["Nước mới"].isna())]
        found_rooms = set(edited["Phòng"].astype(str).str.upper())
        missing_rooms = sorted(set(ROOMS) - found_rooms)
        if len(bad):
            st.warning("Vẫn còn phòng chưa có đủ chỉ số. Vui lòng bổ sung trước khi xác nhận.")
        if missing_rooms:
            st.info("Các phòng không có số liệu đã được tự điền bằng chỉ số gần nhất của tháng trước; không để trống.")

        can_confirm = len(bad) == 0 and found_rooms == set(ROOMS)
        if st.button("✅ XÁC NHẬN & CẬP NHẬT EXCEL", type="primary", disabled=not can_confirm):
            try:
                with st.spinner("Đang cập nhật Excel, giữ nguyên công thức và định dạng..."):
                    result_bytes, updated = update_workbook(
                        st.session_state["input_bytes"],
                        Extraction.model_validate(st.session_state["extraction"]),
                        edited,
                    )
                st.session_state["result_bytes"] = result_bytes
                st.session_state["result_name"] = re.sub(r"\.xlsx$", "", excel.name, flags=re.I) + "_DA_CAP_NHAT.xlsx"
                st.success("✅ Đã cập nhật Excel thành công.")
                st.write("Các ô được cập nhật:")
                st.code("\n".join(updated))
            except Exception as e:
                st.error(f"Không thể cập nhật Excel: {e}")

    if "result_bytes" in st.session_state:
        st.subheader("🎉 4. Hoàn tất")
        st.markdown('<div class="success-box">File Excel đã được cập nhật. Các công thức cũ được giữ lại; Excel sẽ tính lại khi mở file.</div>', unsafe_allow_html=True)
        st.download_button(
            "⬇️ TẢI FILE EXCEL KẾT QUẢ",
            data=st.session_state["result_bytes"],
            file_name=st.session_state["result_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        if st.button("🔄 LÀM LẠI"):
            for k in ["input_bytes", "extraction", "review", "review_editor", "result_bytes", "result_name"]:
                st.session_state.pop(k, None)
            st.rerun()

    st.markdown("---")
    st.markdown('<div class="small">Dữ liệu chỉ được ghi vào các ô chỉ số mới. Không tự ý thay đổi đơn giá hoặc công thức Excel.</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
