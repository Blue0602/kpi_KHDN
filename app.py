# ============================================================
# APP STREAMLIT: DASHBOARD KPI GIA HẠN DỊCH VỤ CNTT/VIỄN THÔNG
# Vai trò: Senior Data Analyst & Python Streamlit Developer
# File chạy: streamlit run app_kpi_gia_han_v2.py
# ============================================================

from datetime import datetime, timezone, timedelta
from pathlib import Path
import unicodedata

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

# Ép cứng múi giờ Việt Nam GMT+7.
# Không dùng giờ hệ thống mặc định để tránh app chạy trên server nước ngoài bị lệch ngày.
VIETNAM_TZ = timezone(timedelta(hours=7))
now_vn = datetime.now(VIETNAM_TZ)
today_vn = pd.Timestamp(now_vn.date())


def get_week_of_month(date_value) -> int:
    """
    Tính tuần thứ mấy trong tháng theo ngày trong tháng.

    Quy ước đơn giản, dễ hiểu cho báo cáo KPI nội bộ:
    - Ngày 01-07: Tuần 1
    - Ngày 08-14: Tuần 2
    - Ngày 15-21: Tuần 3
    - Ngày 22-28: Tuần 4
    - Ngày 29-31: Tuần 5

    Lưu ý: Không dùng .isocalendar().week vì hàm đó trả về tuần 1-52 của cả năm.
    """
    date_value = pd.Timestamp(date_value)
    return min(((date_value.day - 1) // 7) + 1, 5)


current_week_in_month = get_week_of_month(today_vn)

st.title("📊 Dashboard KPI Gia hạn DVCNTT - Long Thành")
st.markdown(
    f"**⏳ Dữ liệu cập nhật: Ngày {today_vn.strftime('%d/%m/%Y')} | "
    f"Tháng: {today_vn.month} | Tuần thứ: {current_week_in_month} trong tháng**"
)
st.caption("Theo dõi KPI gia hạn theo Ngày / Tuần trong tháng / Tháng cho từng nhân sự kinh doanh")


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


def find_header_row(file_input, sheet_name, max_scan_rows=15):
    """
    Một số sheet trong file Excel không có header nằm ngay dòng đầu tiên.
    Hàm này quét vài dòng đầu để tìm dòng chứa cột 'Nhân viên thực hiện'
    hoặc 'Nhân viên phụ trách', sau đó dùng dòng đó làm header.
    """
    preview = pd.read_excel(file_input, sheet_name=sheet_name, header=None, nrows=max_scan_rows)

    for idx, row in preview.iterrows():
        row_values = [normalize_text(value) for value in row.tolist()]
        joined_row = " | ".join(row_values)

        if "nhan vien thuc hien" in joined_row or "nhan vien phu trach" in joined_row:
            return idx

    return 0


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hóa các cột quan trọng về tên thống nhất:
    - nhan_vien
    - ngay_thuc_hien
    - trang_thai

    File thực tế có thể dùng các tên hơi khác:
    - 'Nhân viên thực hiện' hoặc 'Nhân viên phụ trách'
    - 'Trạng thái', 'Kết quả thực hiện', 'KẾT QUẢ THỰC HIỆN'
    """
    col_map = {}

    for col in df.columns:
        norm_col = normalize_text(col)

        if norm_col in ["nhan vien thuc hien", "nhan vien phu trach"]:
            col_map[col] = "nhan_vien"
        elif norm_col == "ngay thuc hien":
            col_map[col] = "ngay_thuc_hien"
        elif norm_col in ["trang thai", "ket qua thuc hien"]:
            col_map[col] = "trang_thai"

    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def remove_empty_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Xóa các cột Unnamed rỗng để bảng chi tiết gọn hơn.
    Các cột Unnamed có dữ liệu thật vẫn được giữ lại để tránh mất dữ liệu.
    """
    cols_to_drop = []
    for col in df.columns:
        if str(col).startswith("Unnamed") and df[col].isna().all():
            cols_to_drop.append(col)
    return df.drop(columns=cols_to_drop, errors="ignore")


# ============================================================
# 4. HÀM ĐỌC TẤT CẢ SHEET TRONG FILE EXCEL
# ============================================================

@st.cache_data(show_spinner=False)
def read_all_sheets(file_input) -> pd.DataFrame:
    """
    Đọc toàn bộ sheet trong file Excel, tự động tìm header,
    chuẩn hóa cột và gộp thành một DataFrame duy nhất.

    Điểm quan trọng:
    - Không chỉ giữ 3 cột KPI, mà giữ lại các cột gốc để dùng cho bảng chi tiết.
    - Thêm cột 'ten_dich_vu' bằng tên sheet để biết ca gia hạn thuộc dịch vụ nào.
    """
    excel = pd.ExcelFile(file_input)
    all_data = []

    for sheet_name in excel.sheet_names:
        try:
            header_row = find_header_row(file_input, sheet_name)
            df = pd.read_excel(file_input, sheet_name=sheet_name, header=header_row)
            df = remove_empty_unnamed_columns(df)
            df = standardize_columns(df)

            df["sheet_name"] = sheet_name
            df["ten_dich_vu"] = sheet_name

            required_cols = {"nhan_vien", "ngay_thuc_hien", "trang_thai"}

            if required_cols.issubset(df.columns):
                all_data.append(df.copy())

        except Exception as e:
            st.warning(f"Không đọc được sheet '{sheet_name}': {e}")

    if not all_data:
        return pd.DataFrame(columns=["nhan_vien", "ngay_thuc_hien", "trang_thai", "sheet_name", "ten_dich_vu"])

    final_df = pd.concat(all_data, ignore_index=True, sort=False)
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

    df["ngay_thuc_hien"] = pd.to_datetime(df["ngay_thuc_hien"], errors="coerce", dayfirst=True)
    df["trang_thai_chuan"] = df["trang_thai"].apply(normalize_text)
    df["nhan_vien_chuan"] = df["nhan_vien"].apply(normalize_text)

    nhan_vien_chuan_dict = {normalize_text(name): name for name in NHAN_VIEN_THEO_DOI}

    df = df[
        (df["trang_thai_chuan"] == "da gia han")
        & (df["ngay_thuc_hien"].notna())
        & (df["nhan_vien_chuan"].isin(nhan_vien_chuan_dict.keys()))
    ].copy()

    df["nhan_vien"] = df["nhan_vien_chuan"].map(nhan_vien_chuan_dict)
    df["ngay"] = df["ngay_thuc_hien"].dt.date
    df["thang"] = df["ngay_thuc_hien"].dt.month
    df["nam"] = df["ngay_thuc_hien"].dt.year
    df["tuan_trong_thang"] = df["ngay_thuc_hien"].apply(get_week_of_month)

    return df


# ============================================================
# 6. HÀM TẠO MỐC THỜI GIAN CHO TUẦN TRONG THÁNG
# ============================================================

def get_previous_week_reference(today: pd.Timestamp):
    """
    Lấy thông tin tuần trước để tính Delta tuần.

    Nếu hiện tại là tuần 2-5: tuần trước nằm trong cùng tháng.
    Nếu hiện tại là tuần 1: tuần trước lấy tuần cuối của tháng trước.
    """
    today = pd.Timestamp(today)
    current_week = get_week_of_month(today)

    if current_week > 1:
        return today.year, today.month, current_week - 1

    previous_month_last_day = today.replace(day=1) - pd.Timedelta(days=1)
    return (
        previous_month_last_day.year,
        previous_month_last_day.month,
        get_week_of_month(previous_month_last_day)
    )


def filter_current_week_in_month(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """
    Lọc KPI tuần này theo Tuần trong tháng, không dùng tuần ISO của năm.
    """
    return df[
        (df["nam"] == today.year)
        & (df["thang"] == today.month)
        & (df["tuan_trong_thang"] == get_week_of_month(today))
    ].copy()


def filter_current_month(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """Lọc dữ liệu trong tháng hiện tại theo GMT+7."""
    return df[
        (df["nam"] == today.year)
        & (df["thang"] == today.month)
    ].copy()


# ============================================================
# 7. HÀM TÍNH KPI NGÀY / TUẦN TRONG THÁNG / THÁNG VÀ DELTA
# ============================================================

def calculate_kpi(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """
    Tính KPI cho từng nhân viên:
    - KPI hôm nay
    - KPI tuần này theo tuần trong tháng
    - KPI tháng này
    - Delta ngày: hôm nay - hôm qua
    - Delta tuần: tuần này trong tháng - tuần liền trước

    Delta dùng cho st.metric để hiển thị mũi tên tăng/giảm.
    """
    today = pd.Timestamp(today).normalize()
    yesterday = today - pd.Timedelta(days=1)

    current_week = get_week_of_month(today)
    previous_year, previous_month, previous_week = get_previous_week_reference(today)

    result = []

    for employee in NHAN_VIEN_THEO_DOI:
        employee_df = df[df["nhan_vien"] == employee].copy()

        today_count = employee_df[employee_df["ngay_thuc_hien"].dt.normalize() == today].shape[0]
        yesterday_count = employee_df[employee_df["ngay_thuc_hien"].dt.normalize() == yesterday].shape[0]

        this_week_count = employee_df[
            (employee_df["nam"] == today.year)
            & (employee_df["thang"] == today.month)
            & (employee_df["tuan_trong_thang"] == current_week)
        ].shape[0]

        last_week_count = employee_df[
            (employee_df["nam"] == previous_year)
            & (employee_df["thang"] == previous_month)
            & (employee_df["tuan_trong_thang"] == previous_week)
        ].shape[0]

        this_month_count = employee_df[
            (employee_df["nam"] == today.year)
            & (employee_df["thang"] == today.month)
        ].shape[0]

        # ------------------------------------------------------------
        # ĐOẠN TÍNH DELTA QUAN TRỌNG
        # ------------------------------------------------------------
        # delta_ngay = số ca hôm nay - số ca hôm qua
        # delta_tuan = số ca tuần hiện tại trong tháng - số ca tuần liền trước
        #
        # Ví dụ ngày 02/06/2026 thuộc Tuần 1 tháng 6:
        # - KPI tuần này = các ca từ ngày 01-07/06/2026
        # - KPI tuần trước = Tuần cuối của tháng 5/2026
        #
        # Khi đưa vào st.metric:
        # - Delta dương hiển thị mũi tên tăng
        # - Delta âm hiển thị mũi tên giảm
        # - Delta bằng 0 thể hiện không thay đổi
        # ------------------------------------------------------------
        delta_ngay = today_count - yesterday_count
        delta_tuan = this_week_count - last_week_count

        result.append({
            "Nhân viên": employee,
            "KPI hôm nay": today_count,
            "KPI hôm qua": yesterday_count,
            "Delta ngày": delta_ngay,
            f"KPI tuần {current_week}": this_week_count,
            "KPI tuần này": this_week_count,
            "KPI tuần trước": last_week_count,
            "Delta tuần": delta_tuan,
            "KPI tháng này": this_month_count,
        })

    return pd.DataFrame(result)


# ============================================================
# 8. SIDEBAR: UPLOAD FILE VÀ CẤU HÌNH
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

st.sidebar.markdown("---")
st.sidebar.caption(
    f"Mốc thời gian đang dùng: {today_vn.strftime('%d/%m/%Y')} - GMT+7"
)


# ============================================================
# 9. MAIN APP
# ============================================================

if file_input is None:
    st.info("Hãy upload file `Gia hạn DVCNTT_Long Thành.xlsx` ở thanh bên trái.")
    st.stop()

with st.spinner("Đang đọc và xử lý dữ liệu Excel..."):
    raw_df = read_all_sheets(file_input)
    df = clean_and_filter_data(raw_df)
    kpi_df = calculate_kpi(df, today_vn)


# ============================================================
# 10. HIỂN THỊ THÔNG TIN TỔNG QUAN DỮ LIỆU
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
    preview_cols = ["nhan_vien", "ten_dich_vu", "ngay_thuc_hien", "trang_thai", "sheet_name", "tuan_trong_thang"]
    preview_cols = [col for col in preview_cols if col in df.columns]
    st.dataframe(
        df[preview_cols],
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 11. METRIC CARDS: KPI TUẦN NÀY THEO TỪNG NHÂN VIÊN
# ============================================================

st.subheader(f"2. KPI tuần {current_week_in_month} trong tháng theo từng nhân viên")

metric_cols = st.columns(len(NHAN_VIEN_THEO_DOI))

for idx, employee in enumerate(NHAN_VIEN_THEO_DOI):
    row = kpi_df[kpi_df["Nhân viên"] == employee].iloc[0]

    with metric_cols[idx]:
        # ------------------------------------------------------------
        # st.metric có tham số delta:
        # - value: KPI tuần này theo tuần trong tháng
        # - delta: chênh lệch so với tuần liền trước
        # Streamlit sẽ tự hiển thị mũi tên tăng/giảm theo giá trị delta.
        # ------------------------------------------------------------
        st.metric(
            label=employee,
            value=int(row["KPI tuần này"]),
            delta=int(row["Delta tuần"])
        )

st.caption(
    f"Delta = KPI tuần {current_week_in_month} tháng {today_vn.month} "
    "- KPI tuần liền trước. Số dương nghĩa là tăng, số âm nghĩa là giảm."
)


# ============================================================
# 12. XEM CHI TIẾT DỮ LIỆU ĐÃ GIA HẠN THEO NHÂN VIÊN
# ============================================================

st.subheader("3. Xem chi tiết các ca đã gia hạn")

selected_employee = st.selectbox(
    "🔍 Chọn nhân viên để xem chi tiết các ca đã gia hạn",
    NHAN_VIEN_THEO_DOI
)

period_option = st.radio(
    "Chọn phạm vi dữ liệu chi tiết",
    [f"Tuần {current_week_in_month} trong tháng", "Tháng này"],
    horizontal=True
)

employee_detail_df = df[df["nhan_vien"] == selected_employee].copy()

if period_option == "Tháng này":
    employee_detail_df = filter_current_month(employee_detail_df, today_vn)
else:
    employee_detail_df = filter_current_week_in_month(employee_detail_df, today_vn)

# Sắp xếp mới nhất lên trên để dễ kiểm tra ca vừa làm.
employee_detail_df = employee_detail_df.sort_values("ngay_thuc_hien", ascending=False)

# Chọn các cột ưu tiên để bảng chi tiết dễ đọc.
priority_cols = [
    "nhan_vien",
    "ten_dich_vu",
    "ngay_thuc_hien",
    "trang_thai",
    "Tên khách hàng",
    "Tên công ty",
    "TÊN TB",
    "Tên tài khoản khách hàng",
    "Mã số thuế",
    "MÃ TB",
    "MA_TB",
    "Mã thuê bao",
    "Tên gói",
    "LOẠI HÌNH",
    "Loại CA sử dụng",
    "Ghi chú",
    "sheet_name",
]

display_cols = [col for col in priority_cols if col in employee_detail_df.columns]

# Nếu thiếu nhiều cột ưu tiên, app vẫn hiển thị các cột hiện có thay vì báo lỗi.
if not display_cols:
    display_cols = employee_detail_df.columns.tolist()

st.info(
    f"Đang hiển thị **{len(employee_detail_df)}** ca đã gia hạn của **{selected_employee}** "
    f"trong phạm vi: **{period_option}**."
)

if employee_detail_df.empty:
    st.warning("Không có dòng dữ liệu đã gia hạn phù hợp với nhân viên và phạm vi đang chọn.")
else:
    detail_display = employee_detail_df[display_cols].copy()

    if "ngay_thuc_hien" in detail_display.columns:
        detail_display["ngay_thuc_hien"] = detail_display["ngay_thuc_hien"].dt.strftime("%d/%m/%Y")

    st.dataframe(
        detail_display,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 13. BẢNG KPI CHI TIẾT NGÀY / TUẦN / THÁNG
# ============================================================

st.subheader("4. Bảng KPI tổng hợp")

st.dataframe(
    kpi_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# 14. BIỂU ĐỒ CỘT PLOTLY: SO SÁNH KPI THÁNG
# ============================================================

st.subheader("5. So sánh tổng số ca đã gia hạn trong tháng")

chart_df = kpi_df[["Nhân viên", "KPI tháng này"]].copy()

# ------------------------------------------------------------
# ĐOẠN VẼ BIỂU ĐỒ PLOTLY QUAN TRỌNG
# ------------------------------------------------------------
# px.bar dùng để tạo biểu đồ cột so sánh KPI tháng của 5 nhân viên.
# color_discrete_sequence=["#1f77b4"] ép toàn bộ cột sang màu xanh dương.
# plot_bgcolor/paper_bgcolor dùng rgba trong suốt để biểu đồ tương thích
# cả nền trắng và nền đen của Streamlit.
# update_xaxes/update_yaxes ẩn gridline rườm rà để giao diện tối giản hơn.
# ------------------------------------------------------------
fig = px.bar(
    chart_df,
    x="Nhân viên",
    y="KPI tháng này",
    text="KPI tháng này",
    color_discrete_sequence=["#1f77b4"],
    title="Tổng ca đã gia hạn trong tháng theo nhân viên"
)

fig.update_traces(
    textposition="outside",
    cliponaxis=False,
    marker_line_width=0
)

fig.update_layout(
    height=500,
    showlegend=False,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=20, r=20, t=60, b=20),
    xaxis_title="Nhân viên",
    yaxis_title="Số ca đã gia hạn",
    font=dict(size=13),
    title_font=dict(size=18)
)

fig.update_xaxes(
    showgrid=False,
    zeroline=False,
    tickangle=0
)

fig.update_yaxes(
    showgrid=False,
    zeroline=False,
    rangemode="tozero"
)

st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 15. PHÂN TÍCH NHANH TỰ ĐỘNG
# ============================================================

st.subheader("6. Nhận xét nhanh")

if not kpi_df.empty:
    top_month = kpi_df.sort_values("KPI tháng này", ascending=False).iloc[0]
    top_week = kpi_df.sort_values("KPI tuần này", ascending=False).iloc[0]

    st.success(
        f"Trong tháng này, **{top_month['Nhân viên']}** đang có KPI cao nhất "
        f"với **{int(top_month['KPI tháng này'])}** ca đã gia hạn."
    )

    st.info(
        f"Trong tuần {current_week_in_month} của tháng này, **{top_week['Nhân viên']}** đang dẫn đầu "
        f"với **{int(top_week['KPI tuần này'])}** ca đã gia hạn."
    )

    negative_delta = kpi_df[kpi_df["Delta tuần"] < 0]

    if not negative_delta.empty:
        names = ", ".join(negative_delta["Nhân viên"].tolist())
        st.warning(
            f"Cần chú ý các nhân sự có KPI tuần này giảm so với tuần liền trước: **{names}**."
        )
    else:
        st.caption("Không có nhân sự nào bị giảm KPI tuần này so với tuần liền trước.")


# ============================================================
# 16. FOOTER
# ============================================================

st.divider()
st.caption(
    "Dashboard được xây dựng bằng Python, Pandas, Streamlit và Plotly. "
    "Mốc ngày hiện tại được ép cứng theo múi giờ Việt Nam GMT+7."
)
