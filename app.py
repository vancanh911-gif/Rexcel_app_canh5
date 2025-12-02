import glob
import os

import pandas as pd
import streamlit as st


def find_excel_files(patterns=None):
    """
    Tìm tất cả file Excel theo danh sách pattern.
    Mặc định: ['*.xlsx', '*.xls'] trong thư mục hiện tại.
    """
    if patterns is None:
        patterns = ["*.xlsx", "*.xls"]

    files = []
    for pattern in patterns:
        files.extend(glob.glob(pattern))

    # Loại bỏ file tổng hợp (nếu có) để tránh lặp
    files = [f for f in files if os.path.basename(f).lower() not in {"tong_hop.xlsx", "tong_hop.xls"}]
    return sorted(files)


def read_and_concat_excels(files, sheet_name=0):
    """
    Đọc và gộp nhiều file Excel thành một DataFrame.
    Thêm cột 'Nguon_file' để biết dữ liệu đến từ file nào.
    """
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name=sheet_name)
            df["Nguon_file"] = os.path.basename(f)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Lỗi khi đọc file {f}: {e}")

    if not dfs:
        return pd.DataFrame()

    return pd.concat(dfs, ignore_index=True)


def main():
    st.set_page_config(page_title="Tổng hợp dữ liệu Excel", layout="wide")
    st.title("📊 Tổng hợp dữ liệu Excel bằng Python & Streamlit")

    st.markdown(
        """
        Ứng dụng này sẽ:
        - **Tự động tìm** các file Excel trong thư mục hiện tại (`*.xlsx`, `*.xls`)
        - **Gộp dữ liệu** của tất cả file lại thành một bảng
        - Cho phép **xem, lọc, tải về** dữ liệu đã tổng hợp
        """
    )

    # Chọn các file Excel
    all_files = find_excel_files()
    if not all_files:
        st.error("Không tìm thấy file Excel nào trong thư mục hiện tại.")
        return

    with st.expander("Danh sách file Excel được tìm thấy", expanded=True):
        st.write(all_files)

    selected_files = st.multiselect(
        "Chọn các file muốn tổng hợp:",
        options=all_files,
        default=all_files,
    )

    if not selected_files:
        st.info("Vui lòng chọn ít nhất một file để tổng hợp.")
        return

    sheet_option = st.text_input(
        "Tên sheet (để mặc định là sheet đầu tiên, nhập tên sheet nếu muốn chỉ định):",
        value="",
    )

    sheet_name = 0 if sheet_option.strip() == "" else sheet_option.strip()

    if st.button("📥 Tổng hợp dữ liệu"):
        with st.spinner("Đang đọc và gộp dữ liệu..."):
            df = read_and_concat_excels(selected_files, sheet_name=sheet_name)

        if df.empty:
            st.warning("Không có dữ liệu sau khi tổng hợp. Vui lòng kiểm tra lại các file/sheet.")
            return

        st.success(f"Đã tổng hợp {len(df)} dòng dữ liệu từ {len(selected_files)} file.")

        # Hiển thị dữ liệu
        st.subheader("Dữ liệu đã tổng hợp")
        st.dataframe(df, use_container_width=True)

        # Tải về dưới dạng Excel
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Tải về dạng CSV",
                data=df.to_csv(index=False).encode("utf-8-sig"),
                file_name="tong_hop.csv",
                mime="text/csv",
            )

        with col2:
            # Lưu tạm vào Excel trong bộ nhớ
            from io import BytesIO

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Tong_hop")
            buffer.seek(0)

            st.download_button(
                label="⬇️ Tải về dạng Excel",
                data=buffer,
                file_name="tong_hop.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()


