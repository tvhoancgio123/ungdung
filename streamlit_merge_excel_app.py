import streamlit as st
import pandas as pd
import io
from typing import List

st.set_page_config(page_title="Gộp file Excel", layout="wide")
st.title("📎 Ứng dụng gộp file Excel — chọn sheet linh hoạt")

st.markdown(
    "Ứng dụng cho phép tải nhiều file Excel lên, chọn sheet từng file hoặc chọn 1 tên sheet chung để gộp, tuỳ chọn thêm cột nguồn và xuất file Excel/CSV." 
)

uploaded_files = st.file_uploader(
    "Kéo thả hoặc chọn nhiều file Excel (xlsx, xls).", type=["xlsx", "xls"], accept_multiple_files=True
)

# Options
with st.sidebar:
    st.header("Tùy chọn gộp")
    add_source_col = st.checkbox("Thêm cột 'source_file'", value=True)
    add_sheet_col = st.checkbox("Thêm cột 'sheet_name'", value=True)
    drop_duplicates = st.checkbox("Loại bỏ bản ghi trùng (toàn bộ cột)", value=False)
    reset_index = st.checkbox("Reset index sau khi gộp", value=True)
    output_format = st.radio("Định dạng xuất", ("xlsx", "csv"))

if not uploaded_files:
    st.info("Vui lòng tải lên ít nhất 1 file Excel để bắt đầu.")
    st.stop()

# Read sheet names for each file and let user choose
file_selections = {}
st.write("### Chọn sheet cho từng file")
for uploaded in uploaded_files:
    try:
        ef = pd.ExcelFile(uploaded)
        sheets = ef.sheet_names
    except Exception as e:
        st.error(f"Không thể đọc file {uploaded.name}: {e}")
        sheets = []

    with st.expander(f"{uploaded.name} — sheets: {len(sheets)}"):
        st.write("Các sheet tìm thấy:", sheets)
        # default select all
        chosen = st.multiselect(f"Chọn sheet để gộp từ {uploaded.name}", options=sheets, default=sheets)
        file_selections[uploaded.name] = {
            "file_obj": uploaded,
            "chosen_sheets": chosen,
        }

st.write("---")
# Option: merge by common sheet name across files
st.write("### Hoặc: chọn 1 tên sheet chung để gộp từ những file có sheet đó")
all_sheet_names = set()
for uploaded in uploaded_files:
    try:
        ef = pd.ExcelFile(uploaded)
        all_sheet_names.update(ef.sheet_names)
    except Exception:
        pass

common_choice = st.selectbox("Chọn tên sheet chung (hoặc để trống)", options=[""] + sorted(list(all_sheet_names)))
apply_common = False
if common_choice:
    apply_common = st.checkbox("Áp dụng gộp theo tên sheet chung cho tất cả file có sheet này", value=True)

if st.button("Gộp các sheet đã chọn"):
    frames: List[pd.DataFrame] = []
    errors = []
    for uploaded in uploaded_files:
        name = uploaded.name
        chosen = file_selections[name]["chosen_sheets"]
        # If common choice enabled, override chosen
        if apply_common and common_choice:
            chosen = [common_choice] if common_choice in pd.ExcelFile(uploaded).sheet_names else []

        for sheet in chosen:
            try:
                df = pd.read_excel(uploaded, sheet_name=sheet)
                if add_source_col:
                    df["source_file"] = name
                if add_sheet_col:
                    df["sheet_name"] = sheet
                frames.append(df)
            except Exception as e:
                errors.append(f"{name} - {sheet}: {e}")

    if not frames:
        st.warning("Không có sheet nào để gộp (kiểm tra lựa chọn).")
    else:
        try:
            result = pd.concat(frames, ignore_index=True, sort=False)
        except Exception as e:
            st.error(f"Lỗi khi gộp dataframes: {e}")
            st.stop()

        if drop_duplicates:
            before = len(result)
            result = result.drop_duplicates()
            after = len(result)
            st.info(f"Đã loại {before - after} bản ghi trùng.")

        if reset_index:
            result = result.reset_index(drop=True)

        st.success("Gộp thành công!")
        st.write("### Xem trước dữ liệu (10 dòng)")
        st.dataframe(result.head(10))

        # Download
        if output_format == "csv":
            towrite = io.BytesIO()
            result.to_csv(towrite, index=False)
            towrite.seek(0)
            st.download_button(label="Tải về CSV", data=towrite, file_name="merged.csv", mime="text/csv")
        else:
            # excel
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                result.to_excel(writer, index=False, sheet_name="merged")
            towrite.seek(0)
            st.download_button(label="Tải về Excel (.xlsx)", data=towrite, file_name="merged.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if errors:
            st.write("### Một số lỗi khi đọc sheet")
            for e in errors:
                st.write("- ", e)

st.write("\n---\nHướng dẫn: cài `pip install streamlit pandas openpyxl` rồi chạy `streamlit run streamlit_merge_excel_app.py`.")
