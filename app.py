# ============================================================
# APP STREAMLIT QUẢN LÝ KPI GIA HẠN DỊCH VỤ CNTT / VIỄN THÔNG
# Tác giả: Senior Python Developer - Streamlit & Data Analysis
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import datetime, timedelta, timezone


# ============================================================
# 1. XỬ LÝ TIMEZONE
# Ưu tiên dùng pytz theo yêu cầu.
# Nếu môi trường Streamlit Cloud chưa cài pytz thì tự fallback về timezone GMT+7.
# Cách này giúp app không bị lỗi ModuleNotFoundError.
# ============================================================

try:
    import pytz
except ModuleNotFoundError:
    pytz = None


# ============================================================
# 2. CẤU HÌNH TRANG STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dashboard KPI Gia Hạn DVCNTT",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# 3. KHAI BÁO BIẾN CỐ ĐỊNH
# ============================================================

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1J0H1GkxM2k74XB4LojPSLan5vtC_HRvT/export?format=xlsx"

EMPLOYEES = [
    "Vương Thanh Thuận",
    "Nguyễn Thị Mai Anh",
    "Lê Văn Lạc",
    "Trần Lữ Long Vân",
    "Hoàng Thị Giang"
]

SKIP_SHEETS = [
    "Điểm tin Quý 1",
    "DANH SÁCH GIA HẠN - GIANG",
    "Sheet1",
    "Sheet2"
]


# ============================================================
# 4. HÀM CHUẨN HÓA CHỮ
# Dùng để:
# - Bỏ dấu tiếng Việt
# - Đưa về chữ thường
# - Xóa khoảng trắng thừa
# - So sánh tên cột, trạng thái, tên nhân viên ổn định hơn
# ============================================================

def normalize_text(value):
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()

    text = unicodedata.normalize("NFD", text)
    text = "".join(
        char for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = " ".join(text.split())
    return text


def clean_column_name(col):
    """
    Chuẩn hóa tên cột an toàn.
    Không dùng data.columns.str.strip() vì một số cột Excel có thể là số hoặc NaN.
    """

    if pd.isna(col):
        return ""

    return str(col).strip().title()


# ============================================================
# 5. HÀM TÌM DÒ DÒNG HEADER TRONG TỪNG SHEET
# Một số sheet Excel có thể không đặt header ở dòng đầu tiên.
# Hàm này sẽ quét vài dòng đầu để tìm dòng có chứa các cột quan trọng.
# ============================================================

def detect_header_row(raw_sheet_df, max_scan_rows=15):
    """
    Tìm dòng header thật trong sheet.
    Nếu không tìm thấy thì mặc định dùng dòng đầu tiên.
    """

    important_keywords = [
        "nhan vien thuc hien",
        "nhan vien phu trach",
        "nhan vien",
        "nguoi thuc hien",
        "ngay thuc hien",
        "ngay gia han",
        "ngay xac thuc",
        "trang thai",
        "ket qua thuc hien",
        "ket qua",
        "tinh trang"
    ]

    scan_limit = min(max_scan_rows, len(raw_sheet_df))

    best_row_index = 0
    best_score = 0

    for row_index in range(scan_limit):
        row_values = raw_sheet_df.iloc[row_index].tolist()
        normalized_values = [normalize_text(value) for value in row_values]

        score = 0

        for cell_value in normalized_values:
            for keyword in important_keywords:
                if keyword == cell_value or keyword in cell_value:
                    score += 1

        if score > best_score:
            best_score = score
            best_row_index = row_index

    return best_row_index


# ============================================================
# 6. ĐỌC TOÀN BỘ SHEET TỪ GOOGLE SHEETS DẠNG EXCEL
# Yêu cầu:
# - Không dùng pd.read_csv
# - Dùng pd.read_excel(..., sheet_name=None)
# - Đọc toàn bộ sheet
# - Bỏ qua sheet báo cáo
# - Thêm cột Loại Dịch Vụ theo tên sheet
# - Gộp tất cả sheet hợp lệ thành một DataFrame tổng
# ============================================================

@st.cache_data(ttl=300)
def load_data_from_google_sheet(excel_url):
    try:
        raw_sheets = pd.read_excel(
            excel_url,
            sheet_name=None,
            header=None,
            engine="openpyxl"
        )

        combined_frames = []

        for sheet_name, raw_sheet_df in raw_sheets.items():
            if sheet_name in SKIP_SHEETS:
                continue

            if raw_sheet_df is None or raw_sheet_df.empty:
                continue

            raw_sheet_df = raw_sheet_df.dropna(axis=0, how="all")
            raw_sheet_df = raw_sheet_df.dropna(axis=1, how="all")

            if raw_sheet_df.empty:
                continue

            header_row = detect_header_row(raw_sheet_df)

            headers = raw_sheet_df.iloc[header_row].tolist()
            data_part = raw_sheet_df.iloc[header_row + 1:].copy()

            data_part.columns = [clean_column_name(col) for col in headers]

            data_part = data_part.dropna(axis=0, how="all")
            data_part = data_part.dropna(axis=1, how="all")

            if data_part.empty:
                continue

            data_part["Loại Dịch Vụ"] = sheet_name

            combined_frames.append(data_part)

        if not combined_frames:
            st.error("❌ Không tìm thấy sheet dữ liệu hợp lệ để xử lý KPI.")
            return pd.DataFrame()

        final_df = pd.concat(
            combined_frames,
            ignore_index=True,
            sort=False
        )

        return final_df

    except Exception as e:
        st.error(
            "❌ Không thể đọc dữ liệu từ Google Sheets dạng Excel. "
            "Vui lòng kiểm tra link chia sẻ, quyền truy cập hoặc kết nối mạng."
        )
        st.error(f"Chi tiết lỗi: {e}")
        return pd.DataFrame()


# ============================================================
# 7. HÀM TÌM CỘT THEO NHIỀU TÊN CÓ THỂ CÓ
# Giúp app chịu được việc mỗi sheet đặt tên cột hơi khác nhau.
# ============================================================

def find_column(df, possible_names):
    normalized_columns = {
        normalize_text(col): col
        for col in df.columns
    }

    normalized_possible_names = [
        normalize_text(name)
        for name in possible_names
    ]

    for possible_name in normalized_possible_names:
        if possible_name in normalized_columns:
            return normalized_columns[possible_name]

    for normalized_col, real_col in normalized_columns.items():
        for possible_name in normalized_possible_names:
            if possible_name in normalized_col or normalized_col in possible_name:
                return real_col

    return None


# ============================================================
# 8. LÀM SẠCH DỮ LIỆU
# Yêu cầu:
# - Chuẩn hóa tên cột
# - Gộp cột Nhân Viên Phụ Trách thành Nhân Viên Thực Hiện
# - Gộp Kết Quả Thực Hiện thành Trạng Thái
# - Ép Ngày Thực Hiện về datetime
# ============================================================

def standardize_dataframe(raw_df):
    if raw_df.empty:
        return pd.DataFrame()

    df = raw_df.copy()

    df.columns = [clean_column_name(col) for col in df.columns]

    employee_col = find_column(
        df,
        [
            "Nhân Viên Thực Hiện",
            "Nhân Viên Phụ Trách",
            "Người Thực Hiện",
            "Nhân Viên",
            "NV Thực Hiện",
            "NV Phụ Trách"
        ]
    )

    date_col = find_column(
        df,
        [
            "Ngày Thực Hiện",
            "Ngày Gia Hạn",
            "Ngày Xác Thực",
            "Ngày Hoàn Thành",
            "Ngày Xử Lý",
            "Thời Gian Thực Hiện"
        ]
    )

    status_col = find_column(
        df,
        [
            "Trạng Thái",
            "Kết Quả Thực Hiện",
            "KẾT QUẢ THỰC HIỆN",
            "Kết Quả",
            "Tình Trạng"
        ]
    )

    service_col = find_column(
        df,
        [
            "Loại Dịch Vụ",
            "Tên Dịch Vụ",
            "Dịch Vụ",
            "Loại Dv",
            "Sản Phẩm",
            "Tên Sản Phẩm"
        ]
    )

    missing_columns = []

    if employee_col is None:
        missing_columns.append("Nhân Viên Thực Hiện / Nhân Viên Phụ Trách")

    if date_col is None:
        missing_columns.append("Ngày Thực Hiện")

    if status_col is None:
        missing_columns.append("Trạng Thái / Kết Quả Thực Hiện")

    if missing_columns:
        st.error(
            "❌ Dữ liệu thiếu cột bắt buộc: "
            + ", ".join(missing_columns)
        )
        st.info("👉 Hãy kiểm tra lại tên cột trong Google Sheets hoặc dòng header của các sheet.")
        return pd.DataFrame()

    clean_df = pd.DataFrame()

    clean_df["Nhân Viên Thực Hiện"] = (
        df[employee_col]
        .astype(str)
        .str.strip()
    )

    clean_df["Ngày Thực Hiện"] = pd.to_datetime(
        df[date_col],
        errors="coerce",
        dayfirst=True
    )

    clean_df["Trạng Thái"] = (
        df[status_col]
        .astype(str)
        .str.strip()
    )

    if service_col is not None:
        clean_df["Loại Dịch Vụ"] = (
            df[service_col]
            .astype(str)
            .str.strip()
        )
    else:
        clean_df["Loại Dịch Vụ"] = "Chưa xác định"

    clean_df = clean_df.dropna(subset=["Ngày Thực Hiện"])

    clean_df["Nhân Viên Chuẩn"] = clean_df["Nhân Viên Thực Hiện"].apply(normalize_text)
    clean_df["Trạng Thái Chuẩn"] = clean_df["Trạng Thái"].apply(normalize_text)

    clean_df["Ngày"] = clean_df["Ngày Thực Hiện"].dt.date
    clean_df["Năm"] = clean_df["Ngày Thực Hiện"].dt.year
    clean_df["Tháng"] = clean_df["Ngày Thực Hiện"].dt.month

    return clean_df


# ============================================================
# 9. XỬ LÝ GIỜ VIỆT NAM GMT+7
# Nếu có pytz: dùng Asia/Ho_Chi_Minh.
# Nếu không có pytz: dùng timezone cố định UTC+7.
# ============================================================

def get_current_vietnam_time():
    if pytz is not None:
        vietnam_tz = pytz.timezone("Asia/Ho_Chi_Minh")
        return datetime.now(vietnam_tz)

    vietnam_tz = timezone(timedelta(hours=7))
    return datetime.now(vietnam_tz)


# ============================================================
# 10. TÍNH TUẦN TRONG THÁNG
# Không dùng .isocalendar().week vì hàm đó trả về tuần trong năm.
#
# Quy ước:
# - Ngày 01 - 07: Tuần 1
# - Ngày 08 - 14: Tuần 2
# - Ngày 15 - 21: Tuần 3
# - Ngày 22 - 28: Tuần 4
# - Ngày 29 - 31: Tuần 5
# ============================================================

def get_week_of_month(date_value):
    day = date_value.day
    week = ((day - 1) // 7) + 1
    return min(week, 5)


def get_previous_month(year, month):
    if month == 1:
        return year - 1, 12

    return year, month - 1


# ============================================================
# 11. LỌC TRẠNG THÁI ĐÃ GIA HẠN
# Điều kiện KPI:
# Cột Trạng Thái chứa chữ "Đã gia hạn".
# Có bỏ dấu và lower để tránh lỗi chữ hoa/thường.
# ============================================================

def filter_success_records(clean_df):
    if clean_df.empty:
        return pd.DataFrame()

    success_df = clean_df[
        clean_df["Trạng Thái Chuẩn"].str.contains(
            "da gia han",
            na=False
        )
    ].copy()

    success_df["Tuần Trong Tháng"] = success_df["Ngày Thực Hiện"].apply(
        get_week_of_month
    )

    return success_df


# ============================================================
# 12. TÍNH KPI VÀ DELTA
# Delta Ngày  = Hôm nay - Hôm qua
# Delta Tuần  = Tuần này - Tuần trước
# Delta Tháng = Tháng này - Tháng trước
# ============================================================

def calculate_employee_kpi(success_df, employee_name, now_vn):
    employee_norm = normalize_text(employee_name)

    employee_df = success_df[
        success_df["Nhân Viên Chuẩn"] == employee_norm
    ].copy()

    today = now_vn.date()
    yesterday = today - timedelta(days=1)

    current_year = now_vn.year
    current_month = now_vn.month
    current_week = get_week_of_month(now_vn)

    previous_month_year, previous_month = get_previous_month(
        current_year,
        current_month
    )

    today_count = employee_df[
        employee_df["Ngày"] == today
    ].shape[0]

    yesterday_count = employee_df[
        employee_df["Ngày"] == yesterday
    ].shape[0]

    day_delta = today_count - yesterday_count

    this_week_count = employee_df[
        (employee_df["Năm"] == current_year)
        & (employee_df["Tháng"] == current_month)
        & (employee_df["Tuần Trong Tháng"] == current_week)
    ].shape[0]

    if current_week == 1:
        previous_week_count = employee_df[
            (employee_df["Năm"] == previous_month_year)
            & (employee_df["Tháng"] == previous_month)
            & (employee_df["Tuần Trong Tháng"] == 5)
        ].shape[0]
    else:
        previous_week_count = employee_df[
            (employee_df["Năm"] == current_year)
            & (employee_df["Tháng"] == current_month)
            & (employee_df["Tuần Trong Tháng"] == current_week - 1)
        ].shape[0]

    week_delta = this_week_count - previous_week_count

    this_month_count = employee_df[
        (employee_df["Năm"] == current_year)
        & (employee_df["Tháng"] == current_month)
    ].shape[0]

    previous_month_count = employee_df[
        (employee_df["Năm"] == previous_month_year)
        & (employee_df["Tháng"] == previous_month)
    ].shape[0]

    month_delta = this_month_count - previous_month_count

    return {
        "Nhân Viên": employee_name,
        "KPI Hôm Nay": today_count,
        "Delta Ngày": day_delta,
        "KPI Tuần Này": this_week_count,
        "Delta Tuần": week_delta,
        "KPI Tháng Này": this_month_count,
        "Delta Tháng": month_delta
    }


def build_kpi_dataframe(success_df, now_vn):
    kpi_rows = []

    for employee in EMPLOYEES:
        kpi_rows.append(
            calculate_employee_kpi(
                success_df=success_df,
                employee_name=employee,
                now_vn=now_vn
            )
        )

    return pd.DataFrame(kpi_rows)


# ============================================================
# 13. HEADER APP
# ============================================================

st.title("📊 Dashboard KPI Gia Hạn Dịch Vụ CNTT / Viễn Thông")

now_vn = get_current_vietnam_time()
current_week = get_week_of_month(now_vn)

st.markdown(
    f"""
    **⏳ Dữ liệu cập nhật: Ngày {now_vn.strftime("%d/%m/%Y")} 
    | Tháng: {now_vn.month} 
    | Tuần thứ: {current_week} trong tháng**
    """
)

st.divider()


# ============================================================
# 14. LOAD VÀ XỬ LÝ DỮ LIỆU
# ============================================================

raw_df = load_data_from_google_sheet(EXCEL_URL)

if raw_df.empty:
    st.warning("⚠️ Không có dữ liệu đầu vào hợp lệ.")
    st.stop()

clean_df = standardize_dataframe(raw_df)

if clean_df.empty:
    st.warning("⚠️ Dữ liệu sau khi làm sạch không hợp lệ.")
    st.stop()

success_df = filter_success_records(clean_df)

if success_df.empty:
    st.warning("⚠️ Không tìm thấy dòng nào có trạng thái 'Đã gia hạn'.")
    st.stop()

kpi_df = build_kpi_dataframe(success_df, now_vn)


# ============================================================
# 15. SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Thông tin hệ thống")

    st.write("**Nguồn dữ liệu:** Google Sheets Excel")
    st.write("**Định dạng đọc:** `.xlsx`")
    st.write("**Cache:** 300 giây")
    st.write("**Timezone:** Asia/Ho_Chi_Minh GMT+7")
    st.write("**Số dòng đã gia hạn:**", len(success_df))

    st.divider()

    metric_mode = st.radio(
        "Chỉ số hiển thị trên thẻ KPI",
        [
            "KPI Tháng Này",
            "KPI Tuần Này",
            "KPI Hôm Nay"
        ],
        index=0
    )


# ============================================================
# 16. KHU VỰC 1 - METRIC CARDS
# ============================================================

st.subheader("📌 Khu vực 1 - Thẻ KPI theo nhân viên")

metric_columns = st.columns(5)

for index, employee in enumerate(EMPLOYEES):
    employee_kpi = kpi_df[
        kpi_df["Nhân Viên"] == employee
    ].iloc[0]

    if metric_mode == "KPI Tháng Này":
        value = employee_kpi["KPI Tháng Này"]
        delta = employee_kpi["Delta Tháng"]
        help_text = "KPI tháng này so với tháng trước"

    elif metric_mode == "KPI Tuần Này":
        value = employee_kpi["KPI Tuần Này"]
        delta = employee_kpi["Delta Tuần"]
        help_text = "KPI tuần này so với tuần trước"

    else:
        value = employee_kpi["KPI Hôm Nay"]
        delta = employee_kpi["Delta Ngày"]
        help_text = "KPI hôm nay so với hôm qua"

    metric_columns[index].metric(
        label=employee,
        value=f"{int(value)} ca",
        delta=f"{int(delta):+d} ca",
        help=help_text
    )

st.caption(
    "Delta = KPI kỳ hiện tại - KPI kỳ liền trước. "
    "Streamlit tự hiển thị mũi tên xanh/đỏ theo giá trị delta."
)

st.divider()


# ============================================================
# 17. KHU VỰC 2 - BIỂU ĐỒ
# ============================================================

st.subheader("📈 Khu vực 2 - Biểu đồ trực quan KPI tháng")

chart_df = kpi_df[["Nhân Viên", "KPI Tháng Này"]].copy()
chart_df = chart_df.rename(
    columns={
        "KPI Tháng Này": "Số Ca Đã Gia Hạn"
    }
)

left_chart, right_chart = st.columns(2)

with left_chart:
    fig_bar = px.bar(
        chart_df,
        x="Nhân Viên",
        y="Số Ca Đã Gia Hạn",
        text="Số Ca Đã Gia Hạn",
        title="So sánh KPI tháng giữa 5 nhân viên"
    )

    fig_bar.update_traces(
        marker_color="blue",
        textposition="outside"
    )

    fig_bar.update_layout(
        xaxis_title=None,
        yaxis_title=None,
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20)
    )

    fig_bar.update_xaxes(
        showgrid=False,
        tickangle=-20
    )

    fig_bar.update_yaxes(
        showgrid=False
    )

    st.plotly_chart(
        fig_bar,
        use_container_width=True
    )

with right_chart:
    fig_donut = px.pie(
        chart_df,
        names="Nhân Viên",
        values="Số Ca Đã Gia Hạn",
        hole=0.55,
        title="Tỷ lệ đóng góp KPI tháng"
    )

    fig_donut.update_traces(
        textposition="inside",
        textinfo="percent+label"
    )

    fig_donut.update_layout(
        legend_title_text="",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20)
    )

    st.plotly_chart(
        fig_donut,
        use_container_width=True
    )

st.divider()


# ============================================================
# 18. KHU VỰC 3 - DRILL-DOWN TABLE
# ============================================================

st.subheader("🔎 Khu vực 3 - Tra cứu chi tiết ca đã gia hạn")

selected_employee = st.selectbox(
    "🔍 Chọn nhân viên để xem chi tiết các ca đã gia hạn trong tháng",
    EMPLOYEES
)

selected_employee_norm = normalize_text(selected_employee)

detail_df = success_df[
    (success_df["Nhân Viên Chuẩn"] == selected_employee_norm)
    & (success_df["Năm"] == now_vn.year)
    & (success_df["Tháng"] == now_vn.month)
].copy()

st.markdown(
    f"**Nhân viên:** {selected_employee} | "
    f"**Số ca đã gia hạn trong tháng {now_vn.month}:** {detail_df.shape[0]} ca"
)

if detail_df.empty:
    st.info(
        f"ℹ️ Chưa có ca 'Đã gia hạn' nào trong tháng {now_vn.month} "
        f"cho nhân viên {selected_employee}."
    )
else:
    table_df = detail_df[
        [
            "Nhân Viên Thực Hiện",
            "Loại Dịch Vụ",
            "Ngày Thực Hiện",
            "Trạng Thái"
        ]
    ].copy()

    table_df["Ngày Thực Hiện"] = table_df["Ngày Thực Hiện"].dt.strftime(
        "%d/%m/%Y"
    )

    table_df = table_df.rename(
        columns={
            "Nhân Viên Thực Hiện": "Nhân Viên"
        }
    )

    st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 19. BẢNG KPI TỔNG HỢP
# ============================================================

with st.expander("📋 Xem bảng KPI tổng hợp ngày / tuần / tháng"):
    st.dataframe(
        kpi_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 20. KIỂM TRA DỮ LIỆU SAU LÀM SẠCH
# Dùng khi cần debug dữ liệu từ Google Sheets.
# ============================================================

with st.expander("🧪 Kiểm tra dữ liệu đã làm sạch"):
    st.write("Số dòng dữ liệu sau làm sạch:", len(clean_df))
    st.write("Số dòng trạng thái Đã gia hạn:", len(success_df))

    st.dataframe(
        clean_df.head(50),
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 21. FOOTER
# ============================================================

st.divider()

st.caption(
    "Dashboard tự động đọc toàn bộ sheet hợp lệ từ Google Sheets Excel, "
    "lọc trạng thái 'Đã gia hạn', tính KPI ngày / tuần / tháng, "
    "và cập nhật dữ liệu theo cache 5 phút."
)
