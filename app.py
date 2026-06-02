# ============================================================
# APP STREAMLIT: DASHBOARD KPI GIA HẠN DỊCH VỤ CNTT/VIỄN THÔNG
# Tác giả: Senior Data Analyst & Python Streamlit Developer
# File chạy: streamlit run app.py
# ============================================================

import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# 1. CẤU HÌNH TRANG STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dashboard KPI Gia hạn DVCNTT",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard KPI Gia hạn DVCNTT - Long Thành")
st.caption("Theo dõi KPI gia hạn theo Ngày / Tuần / Tháng cho từng nhân sự kinh doanh")


# ============================================================
# 2. DANH SÁCH NHÂN SỰ CẦN THEO DÕI
# ============================================================

NHAN_VIEN_THEO_DOI = [
    "Vương Thanh Thuận",
    "Nguyễn Thị Mai Anh",
    "Lê Văn Lạc",
    "Trần Lữ Long Vân",
    "Hoàng Thị Giang",
]


# ============================================================
# 3. HÀM TIỆN ÍCH XỬ LÝ CHỮ TIẾNG VIỆT VÀ TÊN CỘT
# ============================================================

def remove_accents(text: str) -> str:
    """
    Bỏ dấu tiếng Việt để việc so khớp tên cột / trạng thái ổn định hơn.
    Ví dụ: 'Đã gia hạn' -> 'da gia han'
    """
    text = str(text).strip()
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = text.replace("Đ", "D").replace("đ", "d")
    return text.lower().strip()


def normalize_text(text) -> str:
    """
    Chuẩn hóa chuỗi: bỏ dấu, viết thường, xóa khoảng trắng dư.
    Dùng để so sánh tên nhân viên, trạng thái và tên cột.
    """
    if pd.isna(text):
        return ""
    return " ".join(remove_accents(str(text)).split())


def find_header_row(file_path, sheet_name, max_scan_rows=15):
    """
    Một số sheet trong file Excel không có header nằm ngay dòng đầu tiên.
    Hàm này quét vài dòng đầu để tìm dòng chứa cột 'Nhân viên thực hiện'
    hoặc 'Nhân viên phụ trách', sau đó dùng dòng đó làm header.
    """
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=max_scan_rows)

    for idx, row in preview.iterrows():
        row_values = [normalize_text(value) for value in row.tolist()]
        joined_row = " | ".join(row_values)

        if "nhan vien thuc hien" in joined_row or "nhan vien phu trach" in joined_row:
            return idx

    # Nếu không tìm thấy, mặc định dùng dòng đầu tiên làm header
    return 0


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hóa các cột quan trọng về 3 tên thống nhất:
    - nhan_vien
    - ngay_thuc_hien
    - trang_thai

    File thực tế có thể dùng các tên hơi khác:
    - 'Nhân viên thực hiện' hoặc 'Nhân viên phụ trách'
    - 'Trạng thái' hoặc 'Kết quả thực hiện'
    """
    col_map = {}

    for col in df.columns:
        norm_col = normalize_text(col)

        # Cột nhân viên
        if norm_col in ["nhan vien thuc hien", "nhan vien phu trach"]:
            col_map[col] = "nhan_vien"

        # Cột ngày thực hiện
        elif norm_col == "ngay thuc hien":
            col_map[col] = "ngay_thuc_hien"

        # Cột trạng thái xử lý chính.
        # Không lấy các cột như "Trạng thái sản lượng" vì đó không phải trạng thái gia hạn.
        elif norm_col in ["trang thai", "ket qua thuc hien"]:
            col_map[col] = "trang_thai"

    df = df.rename(columns=col_map)

    # Nếu sheet có nhiều cột bị map trùng tên, chỉ giữ cột đầu tiên theo tên chuẩn
    df = df.loc[:, ~df.columns.duplicated()]

    return df


# ============================================================
# 4. HÀM ĐỌC TẤT CẢ SHEET TRONG FILE EXCEL
# ============================================================

@st.cache_data(show_spinner=False)
def read_all_sheets(file_input) -> pd.DataFrame:
    """
    Đọc toàn bộ sheet trong file Excel, tự động tìm header,
    chuẩn hóa cột và gộp thành một DataFrame duy nhất.

    Output tối thiểu gồm:
    - nhan_vien
    - ngay_thuc_hien
    - trang_thai
    - sheet_name
    """
    excel = pd.ExcelFile(file_input)
    all_data = []

    for sheet_name in excel.sheet_names:
        try:
            header_row = find_header_row(file_input, sheet_name)
            df = pd.read_excel(file_input, sheet_name=sheet_name, header=header_row)

            df = standardize_columns(df)
            df["sheet_name"] = sheet_name

            required_cols = {"nhan_vien", "ngay_thuc_hien", "trang_thai"}

            # Chỉ lấy các sheet có đủ 3 cột cần phân tích
            if required_cols.issubset(df.columns):
                df = df[["nhan_vien", "ngay_thuc_hien", "trang_thai", "sheet_name"]].copy()
                all_data.append(df)

        except Exception as e:
            # Nếu một sheet bị lỗi định dạng, bỏ qua sheet đó để app vẫn chạy
            st.warning(f"Không đọc được sheet '{sheet_name}': {e}")

    if not all_data:
        return pd.DataFrame(columns=["nhan_vien", "ngay_thuc_hien", "trang_thai", "sheet_name"])

    final_df = pd.concat(all_data, ignore_index=True)
    return final_df


# ============================================================
# 5. HÀM LÀM SẠCH VÀ LỌC DỮ LIỆU "ĐÃ GIA HẠN"
# ============================================================

def clean_and_filter_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Làm sạch dữ liệu:
    - Chuẩn hóa tên nhân viên
    - Chuyển Ngày thực hiện sang datetime
    - Lọc trạng thái chính xác là 'Đã gia hạn'
    - Chỉ giữ 5 nhân viên đang cần theo dõi
    """
    df = raw_df.copy()

    if df.empty:
        return df

    # Chuyển ngày về kiểu datetime.
    # errors='coerce' giúp dòng sai ngày không làm app bị crash.
    df["ngay_thuc_hien"] = pd.to_datetime(df["ngay_thuc_hien"], errors="coerce", dayfirst=True)

    # Chuẩn hóa trạng thái để lọc ổn định dù dữ liệu có dấu / không dấu / thừa khoảng trắng.
    df["trang_thai_chuan"] = df["trang_thai"].apply(normalize_text)

    # Chuẩn hóa tên nhân viên để so khớp.
    df["nhan_vien_chuan"] = df["nhan_vien"].apply(normalize_text)

    nhan_vien_chuan_dict = {
        normalize_text(name): name for name in NHAN_VIEN_THEO_DOI
    }

    # Lọc các dòng hợp lệ:
    # 1. Trạng thái là "Đã gia hạn"
    # 2. Có ngày thực hiện hợp lệ
    # 3. Thuộc 5 nhân sự cần theo dõi
    df = df[
        (df["trang_thai_chuan"] == "da gia han")
        & (df["ngay_thuc_hien"].notna())
        & (df["nhan_vien_chuan"].isin(nhan_vien_chuan_dict.keys()))
    ].copy()

    # Đưa tên nhân viên về đúng format hiển thị chuẩn
    df["nhan_vien"] = df["nhan_vien_chuan"].map(nhan_vien_chuan_dict)

    return df


# ============================================================
# 6. HÀM TÍNH KPI NGÀY / TUẦN / THÁNG VÀ DELTA
# ============================================================

def calculate_kpi(df: pd.DataFrame, today=None) -> pd.DataFrame:
    """
    Tính KPI cho từng nhân viên:
    - KPI hôm nay
    - KPI tuần này
    - KPI tháng này
    - Delta ngày: hôm nay - hôm qua
    - Delta tuần: tuần này - tuần trước

    Delta dùng cho st.metric để hiển thị mũi tên tăng/giảm.
    """
    if today is None:
        today = pd.Timestamp.today().normalize()
    else:
        today = pd.to_datetime(today).normalize()

    # Mốc thời gian ngày
    yesterday = today - pd.Timedelta(days=1)

    # Mốc thời gian tuần: quy ước tuần bắt đầu từ thứ Hai
    week_start = today - pd.Timedelta(days=today.weekday())
    week_end = week_start + pd.Timedelta(days=7)

    last_week_start = week_start - pd.Timedelta(days=7)
    last_week_end = week_start

    # Mốc thời gian tháng
    month_start = today.replace(day=1)
    next_month_start = month_start + pd.offsets.MonthBegin(1)

    result = []

    for employee in NHAN_VIEN_THEO_DOI:
        employee_df = df[df["nhan_vien"] == employee].copy()

        today_count = employee_df[employee_df["ngay_thuc_hien"].dt.normalize() == today].shape[0]
        yesterday_count = employee_df[employee_df["ngay_thuc_hien"].dt.normalize() == yesterday].shape[0]

        this_week_count = employee_df[
            (employee_df["ngay_thuc_hien"] >= week_start)
            & (employee_df["ngay_thuc_hien"] < week_end)
        ].shape[0]

        last_week_count = employee_df[
            (employee_df["ngay_thuc_hien"] >= last_week_start)
            & (employee_df["ngay_thuc_hien"] < last_week_end)
        ].shape[0]

        this_month_count = employee_df[
            (employee_df["ngay_thuc_hien"] >= month_start)
            & (employee_df["ngay_thuc_hien"] < next_month_start)
        ].shape[0]

        # ------------------------------------------------------------
        # ĐOẠN TÍNH DELTA QUAN TRỌNG
        # ------------------------------------------------------------
        # delta_ngay = số ca hôm nay - số ca hôm qua
        # delta_tuan = số ca tuần này - số ca tuần trước
        #
        # Khi đưa vào st.metric:
        # - Delta dương sẽ hiển thị mũi tên tăng
        # - Delta âm sẽ hiển thị mũi tên giảm
        # - Delta bằng 0 thể hiện không thay đổi
        # ------------------------------------------------------------
        delta_ngay = today_count - yesterday_count
        delta_tuan = this_week_count - last_week_count

        result.append({
            "Nhân viên": employee,
            "KPI hôm nay": today_count,
            "KPI hôm qua": yesterday_count,
            "Delta ngày": delta_ngay,
            "KPI tuần này": this_week_count,
            "KPI tuần trước": last_week_count,
            "Delta tuần": delta_tuan,
            "KPI tháng này": this_month_count,
        })

    return pd.DataFrame(result)


# ============================================================
# 7. SIDEBAR: UPLOAD FILE VÀ BỘ LỌC
# ============================================================

st.sidebar.header("⚙️ Cấu hình dữ liệu")

uploaded_file = st.sidebar.file_uploader(
    "Tải file Excel KPI",
    type=["xlsx", "xls"]
)

default_file = Path("Gia hạn DVCNTT_Long Thành.xlsx")

if uploaded_file is not None:
    file_input = uploaded_file
    st.sidebar.success("Đã dùng file bạn vừa upload.")
elif default_file.exists():
    file_input = default_file
    st.sidebar.info("Đang dùng file Excel mặc định trong cùng thư mục app.py.")
else:
    file_input = None
    st.sidebar.error("Vui lòng upload file Excel để bắt đầu.")


# ============================================================
# 8. MAIN APP
# ============================================================

if file_input is None:
    st.info("Hãy upload file `Gia hạn DVCNTT_Long Thành.xlsx` ở thanh bên trái.")
    st.stop()

with st.spinner("Đang đọc và xử lý dữ liệu Excel..."):
    raw_df = read_all_sheets(file_input)
    df = clean_and_filter_data(raw_df)
    kpi_df = calculate_kpi(df)


# ============================================================
# 9. HIỂN THỊ THÔNG TIN TỔNG QUAN DỮ LIỆU
# ============================================================

st.subheader("1. Tổng quan dữ liệu đã xử lý")

col_a, col_b, col_c = st.columns(3)

with col_a:
    st.metric("Tổng dòng đọc được", f"{len(raw_df):,}")

with col_b:
    st.metric("Tổng ca đã gia hạn hợp lệ", f"{len(df):,}")

with col_c:
    st.metric("Số nhân viên theo dõi", len(NHAN_VIEN_THEO_DOI))

with st.expander("Xem dữ liệu đã lọc"):
    st.dataframe(
        df[["nhan_vien", "ngay_thuc_hien", "trang_thai", "sheet_name"]],
        use_container_width=True
    )


# ============================================================
# 10. METRIC CARDS: KPI TUẦN NÀY THEO TỪNG NHÂN VIÊN
# ============================================================

st.subheader("2. KPI tuần này theo từng nhân viên")

metric_cols = st.columns(len(NHAN_VIEN_THEO_DOI))

for idx, employee in enumerate(NHAN_VIEN_THEO_DOI):
    row = kpi_df[kpi_df["Nhân viên"] == employee].iloc[0]

    with metric_cols[idx]:
        # ------------------------------------------------------------
        # st.metric có tham số delta:
        # - value: KPI tuần này
        # - delta: chênh lệch so với tuần trước
        # Streamlit sẽ tự hiển thị mũi tên tăng/giảm theo giá trị delta.
        # ------------------------------------------------------------
        st.metric(
            label=employee,
            value=int(row["KPI tuần này"]),
            delta=int(row["Delta tuần"])
        )

st.caption("Delta = KPI tuần này - KPI tuần trước. Số dương nghĩa là tăng, số âm nghĩa là giảm.")


# ============================================================
# 11. BẢNG KPI CHI TIẾT NGÀY / TUẦN / THÁNG
# ============================================================

st.subheader("3. Bảng KPI chi tiết")

st.dataframe(
    kpi_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# 12. BIỂU ĐỒ CỘT PLOTLY: SO SÁNH KPI THÁNG
# ============================================================

st.subheader("4. So sánh tổng số ca đã gia hạn trong tháng")

chart_df = kpi_df[["Nhân viên", "KPI tháng này"]].copy()

# ------------------------------------------------------------
# ĐOẠN VẼ BIỂU ĐỒ PLOTLY QUAN TRỌNG
# ------------------------------------------------------------
# px.bar dùng để tạo biểu đồ cột so sánh KPI tháng của 5 nhân viên.
# color_discrete_sequence=["blue"] ép toàn bộ cột sang màu xanh dương.
# template="plotly_white" giúp biểu đồ sạch, hợp với nền trắng.
# Sau đó update_layout và update_xaxes/update_yaxes để ẩn gridline rườm rà,
# giúp biểu đồ nhìn tối giản và vẫn ổn trên nền sáng/tối của Streamlit.
# ------------------------------------------------------------
fig = px.bar(
    chart_df,
    x="Nhân viên",
    y="KPI tháng này",
    text="KPI tháng này",
    color_discrete_sequence=["blue"],
    title="Tổng ca đã gia hạn trong tháng theo nhân viên"
)

fig.update_traces(
    textposition="outside",
    cliponaxis=False
)

fig.update_layout(
    height=500,
    showlegend=False,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=20, r=20, t=60, b=20),
    xaxis_title="Nhân viên",
    yaxis_title="Số ca đã gia hạn",
    font=dict(size=13)
)

fig.update_xaxes(
    showgrid=False,
    zeroline=False
)

fig.update_yaxes(
    showgrid=False,
    zeroline=False,
    rangemode="tozero"
)

st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 13. PHÂN TÍCH NHANH TỰ ĐỘNG
# ============================================================

st.subheader("5. Nhận xét nhanh")

if not kpi_df.empty:
    top_month = kpi_df.sort_values("KPI tháng này", ascending=False).iloc[0]
    top_week = kpi_df.sort_values("KPI tuần này", ascending=False).iloc[0]

    st.success(
        f"Trong tháng này, **{top_month['Nhân viên']}** đang có KPI cao nhất "
        f"với **{int(top_month['KPI tháng này'])}** ca đã gia hạn."
    )

    st.info(
        f"Trong tuần này, **{top_week['Nhân viên']}** đang dẫn đầu "
        f"với **{int(top_week['KPI tuần này'])}** ca đã gia hạn."
    )

    negative_delta = kpi_df[kpi_df["Delta tuần"] < 0]

    if not negative_delta.empty:
        names = ", ".join(negative_delta["Nhân viên"].tolist())
        st.warning(
            f"Cần chú ý các nhân sự có KPI tuần này giảm so với tuần trước: **{names}**."
        )
    else:
        st.caption("Không có nhân sự nào bị giảm KPI tuần này so với tuần trước.")


# ============================================================
# 14. FOOTER
# ============================================================

st.divider()
st.caption(
    "Dashboard được xây dựng bằng Python, Pandas, Streamlit và Plotly. "
    "Dữ liệu được tính theo thời gian thực tại thời điểm mở ứng dụng."
)
