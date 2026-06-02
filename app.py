# ============================================================
# APP STREAMLIT QUẢN LÝ KPI GIA HẠN DỊCH VỤ CNTT / VIỄN THÔNG
# Tác giả: Senior Data Analyst & Streamlit Developer
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import pytz
import unicodedata


# ============================================================
# 1. CẤU HÌNH TRANG STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dashboard KPI Gia Hạn DVCNTT",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# 2. KHAI BÁO NGUỒN DỮ LIỆU GOOGLE SHEETS CSV
# ============================================================

CSV_URL = "https://docs.google.com/spreadsheets/d/1J0H1GkxM2k74XB4LojPSLan5vtC_HRvT/export?format=csv&gid=1903922809"

EMPLOYEES = [
    "Vương Thanh Thuận",
    "Nguyễn Thị Mai Anh",
    "Lê Văn Lạc",
    "Trần Lữ Long Vân",
    "Hoàng Thị Giang"
]


# ============================================================
# 3. HÀM CHUẨN HÓA CHUỖI
# Mục đích:
# - Xử lý khác biệt viết hoa / viết thường
# - Xử lý dấu tiếng Việt
# - Giúp dò tên cột linh hoạt hơn
# ============================================================

def normalize_text(text):
    """
    Chuẩn hóa text về dạng không dấu, chữ thường.
    Ví dụ:
    'KẾT QUẢ THỰC HIỆN' -> 'ket qua thuc hien'
    """
    if pd.isna(text):
        return ""

    text = str(text).strip().lower()

    text = unicodedata.normalize("NFD", text)
    text = "".join(
        char for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = " ".join(text.split())
    return text


def find_column(df, possible_names):
    """
    Tìm tên cột thật trong DataFrame dựa trên danh sách tên cột có thể xuất hiện.
    Hàm này giúp app không bị lỗi nếu Google Sheets đổi tên cột nhẹ.
    """

    normalized_columns = {
        normalize_text(col): col
        for col in df.columns
    }

    possible_names = [normalize_text(name) for name in possible_names]

    # Ưu tiên khớp chính xác
    for name in possible_names:
        if name in normalized_columns:
            return normalized_columns[name]

    # Nếu không khớp chính xác thì dò gần đúng
    for norm_col, real_col in normalized_columns.items():
        for name in possible_names:
            if name in norm_col or norm_col in name:
                return real_col

    return None


# ============================================================
# 4. ĐỌC DỮ LIỆU GOOGLE SHEETS
# Có dùng cache 5 phút để tránh tải lại dữ liệu liên tục
# ============================================================

@st.cache_data(ttl=300)
def load_data_from_google_sheet(csv_url):
    """
    Đọc dữ liệu từ Google Sheets thông qua link CSV.
    Nếu lỗi mạng hoặc lỗi quyền truy cập, trả về DataFrame rỗng.
    """

    try:
        df = pd.read_csv(csv_url)

        # Xóa các cột rỗng hoàn toàn nếu có
        df = df.dropna(axis=1, how="all")

        return df

    except Exception as e:
        st.error(
            f"❌ Không thể đọc dữ liệu từ Google Sheets. "
            f"Vui lòng kiểm tra lại link, quyền chia sẻ hoặc kết nối mạng.\n\nChi tiết lỗi: {e}"
        )
        return pd.DataFrame()


# ============================================================
# 5. XỬ LÝ MÚI GIỜ VIỆT NAM GMT+7
# Yêu cầu:
# - Ép cứng thời gian hiện tại theo Asia/Ho_Chi_Minh
# - Không phụ thuộc timezone của server deploy
# ============================================================

def get_vietnam_now():
    """
    Lấy thời gian hiện tại theo múi giờ Việt Nam GMT+7.
    Dùng pytz để tránh sai lệch khi deploy app lên cloud/server nước ngoài.
    """

    vietnam_tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now_vn = datetime.now(vietnam_tz)
    return now_vn


# ============================================================
# 6. HÀM TÍNH TUẦN TRONG THÁNG
# Không dùng .isocalendar().week vì đó là tuần trong năm.
#
# Quy ước:
# - Ngày 01 - 07: Tuần 1
# - Ngày 08 - 14: Tuần 2
# - Ngày 15 - 21: Tuần 3
# - Ngày 22 - 28: Tuần 4
# - Ngày 29 - 31: Tuần 5
# ============================================================

def get_week_of_month(date_value):
    """
    Tính tuần thứ mấy trong tháng, trả về giá trị từ 1 đến 5.
    """

    day = date_value.day
    week = ((day - 1) // 7) + 1
    return min(week, 5)


# ============================================================
# 7. CHUẨN HÓA DATAFRAME
# Mục đích:
# - Dò đúng cột nhân viên
# - Dò đúng cột ngày thực hiện
# - Dò đúng cột trạng thái / kết quả thực hiện
# - Tạo các cột chuẩn để xử lý KPI
# ============================================================

def prepare_dataframe(raw_df):
    """
    Chuẩn hóa dữ liệu đầu vào từ Google Sheets.
    """

    if raw_df.empty:
        return pd.DataFrame()

    df = raw_df.copy()

    # Dò cột nhân viên
    employee_col = find_column(
        df,
        [
            "Nhân viên thực hiện",
            "Nhân viên phụ trách",
            "Nhân viên",
            "Người thực hiện",
            "NV thực hiện",
            "NV phụ trách"
        ]
    )

    # Dò cột ngày
    date_col = find_column(
        df,
        [
            "Ngày thực hiện",
            "Ngày gia hạn",
            "Ngày xác thực",
            "Ngày hoàn thành",
            "Ngày xử lý",
            "Thời gian thực hiện"
        ]
    )

    # Dò cột trạng thái / kết quả
    status_col = find_column(
        df,
        [
            "Trạng thái",
            "KẾT QUẢ THỰC HIỆN",
            "Kết quả thực hiện",
            "Kết quả",
            "Tình trạng"
        ]
    )

    # Dò cột tên dịch vụ nếu có
    service_col = find_column(
        df,
        [
            "Tên dịch vụ",
            "Dịch vụ",
            "Loại dịch vụ",
            "Sản phẩm",
            "Tên sản phẩm",
            "Gói dịch vụ"
        ]
    )

    missing_cols = []

    if employee_col is None:
        missing_cols.append("Nhân viên thực hiện / Nhân viên phụ trách")

    if date_col is None:
        missing_cols.append("Ngày thực hiện")

    if status_col is None:
        missing_cols.append("Trạng thái / KẾT QUẢ THỰC HIỆN")

    if missing_cols:
        st.error(
            "❌ Dữ liệu đang thiếu các cột bắt buộc: "
            + ", ".join(missing_cols)
        )
        return pd.DataFrame()

    # Tạo cột chuẩn để xử lý
    df["__nhan_vien"] = df[employee_col].astype(str).str.strip()
    df["__trang_thai"] = df[status_col].astype(str).str.strip()
    df["__trang_thai_norm"] = df["__trang_thai"].apply(normalize_text)

    # Chuyển cột ngày về dạng datetime
    # dayfirst=True phù hợp định dạng ngày Việt Nam: dd/mm/yyyy
    df["__ngay_thuc_hien"] = pd.to_datetime(
        df[date_col],
        errors="coerce",
        dayfirst=True
    )

    # Nếu không có cột dịch vụ thì tạo cột mặc định
    if service_col is not None:
        df["__ten_dich_vu"] = df[service_col].astype(str).str.strip()
    else:
        df["__ten_dich_vu"] = "Chưa có thông tin dịch vụ"

    # Loại các dòng không có ngày thực hiện
    df = df.dropna(subset=["__ngay_thuc_hien"])

    # Tạo thêm các cột ngày, tháng, năm, tuần trong tháng để tính KPI
    df["__date"] = df["__ngay_thuc_hien"].dt.date
    df["__year"] = df["__ngay_thuc_hien"].dt.year
    df["__month"] = df["__ngay_thuc_hien"].dt.month

    # Tính tuần trong tháng cho từng dòng dữ liệu
    df["__week_of_month"] = df["__ngay_thuc_hien"].apply(get_week_of_month)

    return df


# ============================================================
# 8. LỌC DỮ LIỆU ĐÃ GIA HẠN
# Trigger chính: "Đã gia hạn"
# Có chuẩn hóa để tránh lỗi do viết hoa / viết thường
# ============================================================

def filter_success_data(df):
    """
    Lọc các dòng có trạng thái là 'Đã gia hạn'.
    """

    if df.empty:
        return pd.DataFrame()

    success_df = df[df["__trang_thai_norm"] == "da gia han"].copy()

    return success_df


# ============================================================
# 9. TÍNH KPI VÀ DELTA
# Delta được tính như sau:
# - Delta ngày = số ca hôm nay - số ca hôm qua
# - Delta tuần = số ca tuần này - số ca tuần trước
# - Delta tháng = số ca tháng này - số ca tháng trước
# ============================================================

def get_previous_month(year, month):
    """
    Trả về năm và tháng liền trước.
    Ví dụ:
    2026, 1 -> 2025, 12
    """

    if month == 1:
        return year - 1, 12

    return year, month - 1


def get_employee_kpi(success_df, employee_name, now_vn):
    """
    Tính KPI ngày, tuần, tháng và Delta cho từng nhân viên.
    """

    today = now_vn.date()
    yesterday = today - timedelta(days=1)

    current_year = now_vn.year
    current_month = now_vn.month
    current_week = get_week_of_month(now_vn)

    previous_month_year, previous_month = get_previous_month(
        current_year,
        current_month
    )

    emp_df = success_df[
        success_df["__nhan_vien"].apply(normalize_text)
        == normalize_text(employee_name)
    ].copy()

    # -----------------------------
    # KPI ngày
    # -----------------------------
    today_count = emp_df[emp_df["__date"] == today].shape[0]
    yesterday_count = emp_df[emp_df["__date"] == yesterday].shape[0]

    # Delta ngày: hôm nay so với hôm qua
    day_delta = today_count - yesterday_count

    # -----------------------------
    # KPI tuần trong tháng
    # -----------------------------
    this_week_count = emp_df[
        (emp_df["__year"] == current_year)
        & (emp_df["__month"] == current_month)
        & (emp_df["__week_of_month"] == current_week)
    ].shape[0]

    # Nếu đang ở tuần 1 thì tuần trước được hiểu là tuần 5 của tháng trước
    if current_week == 1:
        previous_week_count = emp_df[
            (emp_df["__year"] == previous_month_year)
            & (emp_df["__month"] == previous_month)
            & (emp_df["__week_of_month"] == 5)
        ].shape[0]
    else:
        previous_week_count = emp_df[
            (emp_df["__year"] == current_year)
            & (emp_df["__month"] == current_month)
            & (emp_df["__week_of_month"] == current_week - 1)
        ].shape[0]

    # Delta tuần: tuần này so với tuần trước
    week_delta = this_week_count - previous_week_count

    # -----------------------------
    # KPI tháng
    # -----------------------------
    this_month_count = emp_df[
        (emp_df["__year"] == current_year)
        & (emp_df["__month"] == current_month)
    ].shape[0]

    previous_month_count = emp_df[
        (emp_df["__year"] == previous_month_year)
        & (emp_df["__month"] == previous_month)
    ].shape[0]

    # Delta tháng: tháng này so với tháng trước
    month_delta = this_month_count - previous_month_count

    return {
        "employee": employee_name,
        "today_count": today_count,
        "day_delta": day_delta,
        "week_count": this_week_count,
        "week_delta": week_delta,
        "month_count": this_month_count,
        "month_delta": month_delta
    }


def build_kpi_table(success_df, now_vn):
    """
    Tạo bảng KPI tổng hợp cho 5 nhân sự.
    """

    kpi_rows = []

    for employee in EMPLOYEES:
        kpi_rows.append(
            get_employee_kpi(success_df, employee, now_vn)
        )

    return pd.DataFrame(kpi_rows)


# ============================================================
# 10. TẠO GIAO DIỆN APP
# ============================================================

st.title("📊 Dashboard KPI Gia Hạn DVCNTT")

# Lấy thời gian Việt Nam GMT+7
now_vn = get_vietnam_now()
current_week_of_month = get_week_of_month(now_vn)

# In thông báo thời gian dưới tiêu đề
st.markdown(
    f"""
    **⏳ Dữ liệu cập nhật: Ngày {now_vn.strftime('%d/%m/%Y')} 
    | Tháng: {now_vn.month} 
    | Tuần thứ: {current_week_of_month} trong tháng**
    """
)

st.divider()


# ============================================================
# 11. LOAD VÀ XỬ LÝ DỮ LIỆU
# ============================================================

raw_df = load_data_from_google_sheet(CSV_URL)
df = prepare_dataframe(raw_df)
success_df = filter_success_data(df)

if df.empty:
    st.warning("⚠️ Chưa có dữ liệu hợp lệ để hiển thị Dashboard.")
    st.stop()

if success_df.empty:
    st.warning("⚠️ Không tìm thấy dòng nào có trạng thái 'Đã gia hạn'.")
    st.stop()


# Tạo bảng KPI cho 5 nhân viên
kpi_df = build_kpi_table(success_df, now_vn)


# ============================================================
# 12. SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Bộ lọc & Thông tin")

    st.write("**Nguồn dữ liệu:** Google Sheets CSV")
    st.write("**Cache:** 5 phút")
    st.write("**Timezone:** Asia/Ho_Chi_Minh GMT+7")

    st.divider()

    selected_metric = st.radio(
        "Chỉ số hiển thị trên thẻ KPI",
        [
            "KPI Tháng này",
            "KPI Tuần này",
            "KPI Hôm nay"
        ],
        index=0
    )

    st.caption(
        "Mặc định nên dùng KPI Tháng này để theo dõi tiến độ tổng thể."
    )


# ============================================================
# 13. KHU VỰC 1 - METRIC CARDS
# ============================================================

st.subheader("📌 Khu vực 1 - Thẻ chỉ số KPI theo nhân viên")

cols = st.columns(5)

for index, employee in enumerate(EMPLOYEES):
    employee_kpi = kpi_df[kpi_df["employee"] == employee].iloc[0]

    if selected_metric == "KPI Tháng này":
        metric_value = employee_kpi["month_count"]
        metric_delta = employee_kpi["month_delta"]
        metric_label = "Tháng này"

    elif selected_metric == "KPI Tuần này":
        metric_value = employee_kpi["week_count"]
        metric_delta = employee_kpi["week_delta"]
        metric_label = "Tuần này"

    else:
        metric_value = employee_kpi["today_count"]
        metric_delta = employee_kpi["day_delta"]
        metric_label = "Hôm nay"

    # st.metric tự động hiện mũi tên xanh/đỏ dựa trên delta
    cols[index].metric(
        label=f"{employee}",
        value=f"{metric_value} ca",
        delta=f"{metric_delta:+d} ca",
        help=f"KPI {metric_label} so với kỳ liền trước"
    )

st.caption(
    "Delta: số ca kỳ hiện tại trừ số ca kỳ trước. "
    "Ví dụ: Tháng này - Tháng trước, Tuần này - Tuần trước."
)

st.divider()


# ============================================================
# 14. KHU VỰC 2 - BIỂU ĐỒ TRỰC QUAN
# ============================================================

st.subheader("📈 Khu vực 2 - Trực quan hóa KPI tháng")

chart_left, chart_right = st.columns(2)

# Chuẩn bị dữ liệu biểu đồ tháng
chart_df = kpi_df[["employee", "month_count"]].copy()
chart_df = chart_df.rename(
    columns={
        "employee": "Nhân viên",
        "month_count": "Số ca đã gia hạn"
    }
)

# -----------------------------
# Biểu đồ cột Plotly
# -----------------------------
# Màu cột dùng xanh dương.
# Layout tối giản, ẩn gridline rườm rà, phù hợp nền sáng/tối của Streamlit.
# -----------------------------
with chart_left:
    fig_bar = px.bar(
        chart_df,
        x="Nhân viên",
        y="Số ca đã gia hạn",
        text="Số ca đã gia hạn",
        title="So sánh KPI tháng theo nhân viên"
    )

    fig_bar.update_traces(
        marker_color="blue",
        textposition="outside"
    )

    fig_bar.update_layout(
        xaxis_title=None,
        yaxis_title=None,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=False
    )

    fig_bar.update_xaxes(
        showgrid=False,
        tickangle=-20
    )

    fig_bar.update_yaxes(
        showgrid=False
    )

    st.plotly_chart(fig_bar, use_container_width=True)


# -----------------------------
# Biểu đồ Donut Chart
# -----------------------------
# Thể hiện tỷ lệ đóng góp KPI tháng của từng nhân viên.
# -----------------------------
with chart_right:
    fig_donut = px.pie(
        chart_df,
        names="Nhân viên",
        values="Số ca đã gia hạn",
        hole=0.55,
        title="Tỷ lệ đóng góp KPI tháng"
    )

    fig_donut.update_traces(
        textposition="inside",
        textinfo="percent+label"
    )

    fig_donut.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20),
        legend_title_text=""
    )

    st.plotly_chart(fig_donut, use_container_width=True)

st.divider()


# ============================================================
# 15. KHU VỰC 3 - TRA CỨU CHI TIẾT THEO NHÂN VIÊN
# ============================================================

st.subheader("🔎 Khu vực 3 - Tra cứu chi tiết ca đã gia hạn")

selected_employee = st.selectbox(
    "🔍 Chọn nhân viên để xem chi tiết các ca đã gia hạn trong tháng",
    EMPLOYEES
)

# Lọc dữ liệu thành công trong tháng hiện tại của nhân viên được chọn
detail_df = success_df[
    (success_df["__nhan_vien"].apply(normalize_text) == normalize_text(selected_employee))
    & (success_df["__year"] == now_vn.year)
    & (success_df["__month"] == now_vn.month)
].copy()

st.markdown(
    f"**Nhân viên đang xem:** {selected_employee} | "
    f"**Số ca đã gia hạn trong tháng:** {detail_df.shape[0]} ca"
)

if detail_df.empty:
    st.info(
        f"ℹ️ Chưa có ca 'Đã gia hạn' nào trong tháng {now_vn.month} "
        f"cho nhân viên {selected_employee}."
    )
else:
    # Ưu tiên hiển thị các cột dễ đọc.
    # Nếu dữ liệu gốc có nhiều cột khác, vẫn giữ lại bên dưới.
    display_df = detail_df.copy()

    display_df.insert(0, "Nhân viên", display_df["__nhan_vien"])
    display_df.insert(1, "Tên dịch vụ", display_df["__ten_dich_vu"])
    display_df.insert(
        2,
        "Ngày thực hiện",
        display_df["__ngay_thuc_hien"].dt.strftime("%d/%m/%Y")
    )
    display_df.insert(3, "Trạng thái", display_df["__trang_thai"])
    display_df.insert(4, "Tuần trong tháng", display_df["__week_of_month"])

    # Ẩn các cột kỹ thuật bắt đầu bằng "__"
    technical_cols = [
        col for col in display_df.columns
        if str(col).startswith("__")
    ]

    display_df = display_df.drop(columns=technical_cols, errors="ignore")

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 16. BẢNG KPI TỔNG HỢP PHỤ
# ============================================================

with st.expander("📋 Xem bảng KPI tổng hợp ngày / tuần / tháng"):
    summary_df = kpi_df.rename(
        columns={
            "employee": "Nhân viên",
            "today_count": "KPI hôm nay",
            "day_delta": "Delta ngày",
            "week_count": "KPI tuần này",
            "week_delta": "Delta tuần",
            "month_count": "KPI tháng này",
            "month_delta": "Delta tháng"
        }
    )

    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 17. FOOTER
# ============================================================

st.divider()

st.caption(
    "Dashboard tự động đọc dữ liệu từ Google Sheets, "
    "lọc trạng thái 'Đã gia hạn', tính KPI theo ngày / tuần / tháng "
    "và cập nhật cache mỗi 5 phút."
)
