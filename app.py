# ============================================================
# APP STREAMLIT QUẢN LÝ KPI GIA HẠN DỊCH VỤ CNTT / VIỄN THÔNG
# Công nghệ sử dụng: Streamlit, Pandas, Plotly, pytz
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import pytz
import unicodedata


# ============================================================
# 1. CẤU HÌNH TRANG
# ============================================================

st.set_page_config(
    page_title="Dashboard KPI Gia Hạn Dịch Vụ",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# 2. KHAI BÁO THÔNG TIN CỐ ĐỊNH
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
# 3. HÀM CHUẨN HÓA CHUỖI
# Mục đích:
# - Bỏ dấu tiếng Việt
# - Đưa về chữ thường
# - Xóa khoảng trắng thừa
# - Giúp so sánh trạng thái và tên cột ổn định hơn
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
    Chuẩn hóa tên cột theo yêu cầu:
    - strip khoảng trắng
    - title để đồng bộ kiểu viết
    """

    if pd.isna(col):
        return ""

    return str(col).strip().title()


# ============================================================
# 4. HÀM ĐỌC TOÀN BỘ SHEET TỪ GOOGLE SHEETS DẠNG EXCEL
# Yêu cầu:
# - Không dùng pd.read_csv
# - Dùng pd.read_excel(..., sheet_name=None)
# - Cache dữ liệu trong 5 phút
# - Bỏ qua các sheet báo cáo không dùng để tính KPI
# - Gộp toàn bộ sheet hợp lệ thành một DataFrame tổng
# ============================================================

@st.cache_data(ttl=300)
def load_data_from_google_sheet(excel_url):
    try:
        all_sheets = pd.read_excel(
            excel_url,
            sheet_name=None,
            engine="openpyxl"
        )

        combined_data = []

        for sheet_name, sheet_df in all_sheets.items():
            if sheet_name in SKIP_SHEETS:
                continue

            if sheet_df is None or sheet_df.empty:
                continue

            df = sheet_df.copy()

            df = df.dropna(axis=0, how="all")
            df = df.dropna(axis=1, how="all")

            if df.empty:
                continue

            df.columns = [clean_column_name(col) for col in df.columns]

            df["Loại Dịch Vụ"] = sheet_name

            combined_data.append(df)

        if not combined_data:
            st.error("❌ Không tìm thấy sheet dữ liệu hợp lệ để xử lý KPI.")
            return pd.DataFrame()

        final_df = pd.concat(
            combined_data,
            ignore_index=True,
            sort=False
        )

        return final_df

    except Exception as e:
        st.error(
            "❌ Không thể đọc dữ liệu từ Google Sheets dạng Excel. "
            "Vui lòng kiểm tra lại link chia sẻ, quyền truy cập hoặc kết nối mạng."
        )
        st.exception(e)
        return pd.DataFrame()


# ============================================================
# 5. HÀM LÀM SẠCH VÀ CHUẨN HÓA DỮ LIỆU
# Mục đích:
# - Gộp các cột cùng ý nghĩa
# - Đổi tên về bộ cột chuẩn
# - Ép kiểu ngày thực hiện về datetime
# ============================================================

def standardize_columns(df):
    if df.empty:
        return pd.DataFrame()

    data = df.copy()

    data.columns = [clean_column_name(col) for col in data.columns]

    rename_map = {}

    for col in data.columns:
        normalized_col = normalize_text(col)

        if normalized_col in [
            "nhan vien phu trach",
            "nhan vien thuc hien",
            "nguoi thuc hien",
            "nhan vien",
            "nv phu trach",
            "nv thuc hien"
        ]:
            rename_map[col] = "Nhân Viên Thực Hiện"

        elif normalized_col in [
            "ket qua thuc hien",
            "trang thai",
            "ket qua",
            "tinh trang"
        ]:
            rename_map[col] = "Trạng Thái"

        elif normalized_col in [
            "ngay thuc hien",
            "ngay gia han",
            "ngay xac thuc",
            "ngay hoan thanh",
            "ngay xu ly",
            "thoi gian thuc hien"
        ]:
            rename_map[col] = "Ngày Thực Hiện"

    data = data.rename(columns=rename_map)

    required_columns = [
        "Nhân Viên Thực Hiện",
        "Ngày Thực Hiện",
        "Trạng Thái"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in data.columns
    ]

    if missing_columns:
        st.error(
            "❌ Dữ liệu đang thiếu cột bắt buộc: "
            + ", ".join(missing_columns)
        )
        return pd.DataFrame()

    data["Nhân Viên Thực Hiện"] = (
        data["Nhân Viên Thực Hiện"]
        .astype(str)
        .str.strip()
    )

    data["Trạng Thái"] = (
        data["Trạng Thái"]
        .astype(str)
        .str.strip()
    )

    data["Ngày Thực Hiện"] = pd.to_datetime(
        data["Ngày Thực Hiện"],
        errors="coerce",
        dayfirst=True
    )

    data = data.dropna(subset=["Ngày Thực Hiện"])

    data["Trạng Thái Chuẩn"] = data["Trạng Thái"].apply(normalize_text)
    data["Nhân Viên Chuẩn"] = data["Nhân Viên Thực Hiện"].apply(normalize_text)

    return data


# ============================================================
# 6. XỬ LÝ THỜI GIAN GMT+7
# Dùng pytz để ép cứng timezone Asia/Ho_Chi_Minh.
# Việc này giúp app không bị sai ngày khi deploy lên server nước ngoài.
# ============================================================

def get_current_vietnam_time():
    vietnam_timezone = pytz.timezone("Asia/Ho_Chi_Minh")
    return datetime.now(vietnam_timezone)


# ============================================================
# 7. TÍNH TUẦN TRONG THÁNG
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
# 8. LỌC DỮ LIỆU ĐÃ GIA HẠN
# Điều kiện KPI:
# - Cột Trạng Thái chứa chữ "Đã gia hạn"
# - Có xử lý lower và bỏ dấu để tránh lỗi hoa/thường
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

    success_df["Ngày"] = success_df["Ngày Thực Hiện"].dt.date
    success_df["Năm"] = success_df["Ngày Thực Hiện"].dt.year
    success_df["Tháng"] = success_df["Ngày Thực Hiện"].dt.month
    success_df["Tuần Trong Tháng"] = success_df["Ngày Thực Hiện"].apply(
        get_week_of_month
    )

    return success_df


# ============================================================
# 9. TÍNH KPI VÀ DELTA
# Delta:
# - Delta Ngày = Hôm nay - Hôm qua
# - Delta Tuần = Tuần này - Tuần trước
# - Delta Tháng = Tháng này - Tháng trước
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
    kpi_data = []

    for employee in EMPLOYEES:
        kpi_data.append(
            calculate_employee_kpi(
                success_df=success_df,
                employee_name=employee,
                now_vn=now_vn
            )
        )

    return pd.DataFrame(kpi_data)


# ============================================================
# 10. HEADER APP
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
# 11. LOAD DỮ LIỆU
# ============================================================

raw_df = load_data_from_google_sheet(EXCEL_URL)

if raw_df.empty:
    st.warning("⚠️ Không có dữ liệu đầu vào hợp lệ.")
    st.stop()

clean_df = standardize_columns(raw_df)

if clean_df.empty:
    st.warning("⚠️ Dữ liệu sau khi làm sạch không hợp lệ.")
    st.stop()

success_df = filter_success_records(clean_df)

if success_df.empty:
    st.warning("⚠️ Không tìm thấy dòng nào có trạng thái 'Đã gia hạn'.")
    st.stop()

kpi_df = build_kpi_dataframe(success_df, now_vn)


# ============================================================
# 12. SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Thông tin hệ thống")

    st.write("**Nguồn dữ liệu:** Google Sheets Excel")
    st.write("**Định dạng đọc:** `.xlsx`")
    st.write("**Cache:** 300 giây")
    st.write("**Timezone:** Asia/Ho_Chi_Minh GMT+7")

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

    st.caption(
        "Mặc định dùng KPI Tháng để phù hợp với biểu đồ và bảng chi tiết."
    )


# ============================================================
# 13. KHU VỰC 1 - METRIC CARDS
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
        help_text = "KPI tuần này so với tuần trước trong tháng"

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
    "Delta được tính bằng số ca kỳ hiện tại trừ số ca kỳ liền trước. "
    "Streamlit sẽ tự hiển thị mũi tên xanh hoặc đỏ theo giá trị delta."
)

st.divider()


# ============================================================
# 14. KHU VỰC 2 - BIỂU ĐỒ TRỰC QUAN
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
# 15. KHU VỰC 3 - DRILL-DOWN TABLE
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
    display_columns = [
        "Nhân Viên Thực Hiện",
        "Loại Dịch Vụ",
        "Ngày Thực Hiện",
        "Trạng Thái"
    ]

    available_columns = [
        col for col in display_columns
        if col in detail_df.columns
    ]

    table_df = detail_df[available_columns].copy()

    if "Ngày Thực Hiện" in table_df.columns:
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
# 16. BẢNG KPI TỔNG HỢP
# ============================================================

with st.expander("📋 Xem bảng KPI tổng hợp ngày / tuần / tháng"):
    st.dataframe(
        kpi_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 17. FOOTER
# ============================================================

st.divider()

st.caption(
    "Ứng dụng tự động đọc toàn bộ sheet dữ liệu hợp lệ từ Google Sheets Excel, "
    "lọc các dòng có trạng thái 'Đã gia hạn', tính KPI theo ngày / tuần / tháng, "
    "và cập nhật dữ liệu theo cache 5 phút."
)
