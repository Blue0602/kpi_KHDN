# ============================================================
# APP STREAMLIT QUẢN LÝ KPI GIA HẠN DỊCH VỤ CNTT / VIỄN THÔNG
# Đọc Google Sheets dạng Excel nhiều sheet
# Xử lý KPI cho 5 nhân sự
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import datetime, timedelta, timezone


# ============================================================
# 1. TIMEZONE GMT+7
# Ưu tiên dùng pytz. Nếu môi trường thiếu pytz thì tự fallback.
# ============================================================

try:
    import pytz
except ModuleNotFoundError:
    pytz = None


# ============================================================
# 2. CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dashboard KPI Gia Hạn DVCNTT",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# 3. THÔNG TIN CỐ ĐỊNH
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
# 4. HÀM CHUẨN HÓA CHUỖI
# Dùng để bỏ dấu, đưa về chữ thường, xóa khoảng trắng thừa.
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


def clean_column_name(value):
    """
    Chuẩn hóa tên cột an toàn.
    Không dùng .str.strip() trực tiếp vì có sheet có cột rỗng hoặc cột dạng số.
    """

    if pd.isna(value):
        return ""

    return str(value).strip().title()


def make_unique_columns(columns):
    """
    Xử lý trường hợp tên cột bị trùng sau khi chuẩn hóa.
    Ví dụ:
    ['Trạng Thái', 'Trạng Thái'] -> ['Trạng Thái', 'Trạng Thái_2']
    """

    seen = {}
    unique_cols = []

    for col in columns:
        col = clean_column_name(col)

        if col == "":
            col = "Cot_Rong"

        if col not in seen:
            seen[col] = 1
            unique_cols.append(col)
        else:
            seen[col] += 1
            unique_cols.append(f"{col}_{seen[col]}")

    return unique_cols


# ============================================================
# 5. TÌM DÒNG HEADER THẬT TRONG SHEET
# Một số sheet không có header ở dòng đầu tiên nên cần tự dò.
# ============================================================

def detect_header_row(raw_sheet_df, max_scan_rows=20):
    important_keywords = [
        "nhan vien thuc hien",
        "nhan vien phu trach",
        "nhan vien",
        "nguoi thuc hien",
        "ngay thuc hien",
        "ngay gia han",
        "ngay xac thuc",
        "ngay hoan thanh",
        "trang thai",
        "ket qua thuc hien",
        "ket qua",
        "tinh trang"
    ]

    best_row_index = 0
    best_score = 0

    scan_limit = min(max_scan_rows, len(raw_sheet_df))

    for row_index in range(scan_limit):
        row_values = raw_sheet_df.iloc[row_index].tolist()
        normalized_values = [normalize_text(value) for value in row_values]

        score = 0

        for cell in normalized_values:
            for keyword in important_keywords:
                if keyword == cell or keyword in cell:
                    score += 1

        if score > best_score:
            best_score = score
            best_row_index = row_index

    return best_row_index


# ============================================================
# 6. TÌM CỘT THEO NHIỀU TÊN CÓ THỂ XUẤT HIỆN
# ============================================================

def find_column(df, possible_names):
    normalized_columns = {
        normalize_text(col): col
        for col in df.columns
    }

    possible_names_norm = [
        normalize_text(name)
        for name in possible_names
    ]

    # Khớp chính xác trước
    for name in possible_names_norm:
        if name in normalized_columns:
            return normalized_columns[name]

    # Nếu không khớp chính xác thì dò gần đúng
    for norm_col, real_col in normalized_columns.items():
        for name in possible_names_norm:
            if name in norm_col or norm_col in name:
                return real_col

    return None


# ============================================================
# 7. CHUẨN HÓA 1 SHEET VỀ 4 CỘT CHÍNH
# Cách này tránh lỗi concat do trùng tên cột.
# ============================================================

def standardize_one_sheet(sheet_df, sheet_name):
    if sheet_df is None or sheet_df.empty:
        return pd.DataFrame()

    raw = sheet_df.copy()

    raw = raw.dropna(axis=0, how="all")
    raw = raw.dropna(axis=1, how="all")

    if raw.empty:
        return pd.DataFrame()

    header_row = detect_header_row(raw)

    headers = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].copy()

    data.columns = make_unique_columns(headers)

    data = data.dropna(axis=0, how="all")
    data = data.dropna(axis=1, how="all")

    if data.empty:
        return pd.DataFrame()

    employee_col = find_column(
        data,
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
        data,
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
        data,
        [
            "Trạng Thái",
            "Kết Quả Thực Hiện",
            "KẾT QUẢ THỰC HIỆN",
            "Kết Quả",
            "Tình Trạng"
        ]
    )

    if employee_col is None or date_col is None or status_col is None:
        return pd.DataFrame()

    clean = pd.DataFrame()

    clean["Nhân Viên Thực Hiện"] = data[employee_col].astype(str).str.strip()

    clean["Ngày Thực Hiện"] = pd.to_datetime(
        data[date_col],
        errors="coerce",
        dayfirst=True
    )

    clean["Trạng Thái"] = data[status_col].astype(str).str.strip()

    clean["Loại Dịch Vụ"] = sheet_name

    clean = clean.dropna(subset=["Ngày Thực Hiện"])

    clean["Nhân Viên Chuẩn"] = clean["Nhân Viên Thực Hiện"].apply(normalize_text)
    clean["Trạng Thái Chuẩn"] = clean["Trạng Thái"].apply(normalize_text)

    clean["Ngày"] = clean["Ngày Thực Hiện"].dt.date
    clean["Năm"] = clean["Ngày Thực Hiện"].dt.year
    clean["Tháng"] = clean["Ngày Thực Hiện"].dt.month

    return clean


# ============================================================
# 8. ĐỌC GOOGLE SHEETS DẠNG EXCEL NHIỀU SHEET
# Không dùng pd.read_csv.
# Có cache 5 phút.
# ============================================================

@st.cache_data(ttl=300)
def load_data_from_google_sheet(excel_url):
    try:
        all_sheets = pd.read_excel(
            excel_url,
            sheet_name=None,
            header=None,
            engine="openpyxl"
        )

        valid_frames = []
        skipped_info = []

        for sheet_name, sheet_df in all_sheets.items():
            if sheet_name in SKIP_SHEETS:
                skipped_info.append(f"Bỏ qua sheet báo cáo: {sheet_name}")
                continue

            clean_sheet = standardize_one_sheet(sheet_df, sheet_name)

            if clean_sheet.empty:
                skipped_info.append(f"Bỏ qua sheet không đủ cột KPI: {sheet_name}")
                continue

            valid_frames.append(clean_sheet)

        if len(valid_frames) == 0:
            return pd.DataFrame(), skipped_info

        final_df = pd.concat(
            valid_frames,
            ignore_index=True
        )

        return final_df, skipped_info

    except Exception as e:
        st.error(
            "❌ Không thể đọc dữ liệu từ Google Sheets dạng Excel. "
            "Vui lòng kiểm tra link chia sẻ, quyền truy cập hoặc kết nối mạng."
        )
        st.error(f"Chi tiết lỗi: {e}")
        return pd.DataFrame(), []


# ============================================================
# 9. LẤY GIỜ VIỆT NAM GMT+7
# ============================================================

def get_current_vietnam_time():
    if pytz is not None:
        vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
        return datetime.now(vn_tz)

    vn_tz = timezone(timedelta(hours=7))
    return datetime.now(vn_tz)


# ============================================================
# 10. TÍNH TUẦN TRONG THÁNG
# Không dùng isocalendar().week.
# Quy ước:
# 01-07: Tuần 1
# 08-14: Tuần 2
# 15-21: Tuần 3
# 22-28: Tuần 4
# 29-31: Tuần 5
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
# 11. LỌC DÒNG ĐÃ GIA HẠN
# Điều kiện: Trạng thái chứa chữ "Đã gia hạn".
# Có xử lý lower + bỏ dấu.
# ============================================================

def filter_success_records(df):
    if df.empty:
        return pd.DataFrame()

    success_df = df[
        df["Trạng Thái Chuẩn"].str.contains(
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
    rows = []

    for employee in EMPLOYEES:
        rows.append(
            calculate_employee_kpi(
                success_df=success_df,
                employee_name=employee,
                now_vn=now_vn
            )
        )

    return pd.DataFrame(rows)


# ============================================================
# 13. HEADER
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
# 14. LOAD DỮ LIỆU
# ============================================================

clean_df, skipped_info = load_data_from_google_sheet(EXCEL_URL)

if clean_df.empty:
    st.warning("⚠️ Không có dữ liệu đầu vào hợp lệ.")
    with st.expander("🧪 Xem thông tin sheet bị bỏ qua"):
        for item in skipped_info:
            st.write("-", item)
    st.stop()

success_df = filter_success_records(clean_df)

if success_df.empty:
    st.warning("⚠️ Không tìm thấy dòng nào có trạng thái 'Đã gia hạn'.")
    with st.expander("🧪 Xem dữ liệu đã đọc được"):
        st.dataframe(clean_df.head(100), use_container_width=True, hide_index=True)
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
    st.write("**Số dòng dữ liệu hợp lệ:**", len(clean_df))
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

    st.plotly_chart(fig_bar, use_container_width=True)

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

    st.plotly_chart(fig_donut, use_container_width=True)

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

    table_df["Ngày Thực Hiện"] = table_df["Ngày Thực Hiện"].dt.strftime("%d/%m/%Y")

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
    st.dataframe(kpi_df, use_container_width=True, hide_index=True)


# ============================================================
# 20. DEBUG DỮ LIỆU
# ============================================================

with st.expander("🧪 Kiểm tra dữ liệu đã đọc và sheet bị bỏ qua"):
    st.write("### Sheet bị bỏ qua hoặc không đủ cột")
    if skipped_info:
        for item in skipped_info:
            st.write("-", item)
    else:
        st.write("Không có sheet nào bị bỏ qua ngoài danh sách mặc định.")

    st.write("### Dữ liệu đã chuẩn hóa")
    st.dataframe(clean_df.head(100), use_container_width=True, hide_index=True)


# ============================================================
# 21. FOOTER
# ============================================================

st.divider()

st.caption(
    "Dashboard tự động đọc nhiều sheet từ Google Sheets Excel, "
    "lọc trạng thái 'Đã gia hạn', tính KPI ngày / tuần / tháng, "
    "và cập nhật dữ liệu theo cache 5 phút."
)
