# -*- coding: utf-8 -*-
"""
TRANG THÔNG TIN ĐIỆN TỬ (INTERACTIVE DASHBOARD) THỐNG KÊ TÔN GIÁO TRÊN ĐỊA BÀN TỈNH
Thiết kế chuyên biệt phục vụ Lãnh đạo Tỉnh ủy chỉ đạo, điều hành.
Chạy ứng dụng bằng lệnh: streamlit run dashboard_ton_giao.py
"""

import os
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# 1. CẤU HÌNH TRANG WEB & GIAO DIỆN (UI/UX THÂN THIỆN VỚI LÃNH ĐẠO)
# ==============================================================================
st.set_page_config(
    page_title="Hệ thống Thông tin Tôn giáo Tỉnh Gia Lai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #1A365D 0%, #2B6CB0 100%);
        color: white;
        padding: 18px 24px;
        border-radius: 10px;
        margin-bottom: 18px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
    }
    .main-header h1 {
        color: white !important;
        font-size: 26px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }
    .main-header p {
        color: #E2E8F0 !important;
        font-size: 15px !important;
        margin: 4px 0 0 0 !important;
    }
    .kpi-card {
        background-color: #FFFFFF;
        border-left: 6px solid #2B6CB0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        margin-bottom: 12px;
    }
    .kpi-title {
        font-size: 14px;
        font-weight: 600;
        color: #4A5568;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 32px;
        font-weight: 800;
        color: #1A365D;
        margin: 6px 0 2px 0;
    }
    .kpi-sub {
        font-size: 13px;
        color: #718096;
    }
    .insight-box {
        background-color: #F7FAFC;
        border: 1px solid #CBD5E0;
        border-left: 5px solid #D69E2E;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 16px;
        font-size: 15px;
        line-height: 1.55;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 2. HÀM ĐỌC VÀ CHUẨN HÓA DỮ LIỆU TỰ ĐỘNG TỪ FILE EXCEL GỐC
# ==============================================================================
DEFAULT_EXCEL_NAME = "2026 a H. TỔNG HỢP SO LIEU CÁC TÔN GIÁO TREN DIA BAN TINH.xlsx"


@st.cache_data(show_spinner="Đang tải và chuẩn hóa dữ liệu tôn giáo toàn tỉnh...")
def load_and_normalize_data(file_source) -> pd.DataFrame:
    """
    Đọc sheet 'TH chung' có cấu trúc tiêu đề đa tầng (dòng 3-6)
    và chuyển đổi (Unpivot/Melt) về bảng phẳng chuẩn phân tích (Tidy Data).
    """
    wb = openpyxl.load_workbook(file_source, data_only=True)
    ws = wb["TH chung"]

    metric_norm = {
        "tổ chức tgtt": "To_chuc_TGTT",
        "tổ chức tôn giáo trực thuộc": "To_chuc_TGTT",
        "cơ sở": "Co_so_tho_tu",
        "chức sắc": "Chuc_sac",
        "chức việc": "Chuc_viec",
        "nhà tu hành": "Nha_tu_hanh",
        "tín đồ": "So_tin_do",
        "điểm nhóm shtgtt": "Diem_nhom_SHTGTT",
    }

    col_map = {}
    curr_religion = None
    for c in range(4, 119):
        r4 = ws.cell(4, c).value
        r5 = ws.cell(5, c).value
        if r4 is not None and str(r4).strip() != "":
            curr_religion = str(r4).strip()
        if r5 is not None and str(r5).strip() != "":
            m_raw = str(r5).strip().lower()
            m_clean = metric_norm.get(m_raw, m_raw)
            col_map[c] = (curr_religion, m_clean)

    records = []
    curr_region = None
    for r in range(7, 144):
        # Bỏ qua 2 dòng tổng cộng trung gian (dòng 84 và 85) để tránh trùng lặp
        if r in (84, 85):
            continue
        reg_val = ws.cell(r, 1).value
        if reg_val is not None and str(reg_val).strip() != "":
            curr_region = str(reg_val).strip()

        stt = ws.cell(r, 2).value
        commune = ws.cell(r, 3).value
        if commune is None or str(commune).strip() == "":
            continue
        commune_clean = " ".join(str(commune).strip().split())

        rel_data = {}
        for c, (rel, metric) in col_map.items():
            val = ws.cell(r, c).value
            num_val = 0
            if val is not None and str(val).strip() != "":
                try:
                    num_val = int(float(str(val).strip()))
                except ValueError:
                    num_val = 0
            if rel not in rel_data:
                rel_data[rel] = {
                    "To_chuc_TGTT": 0,
                    "Co_so_tho_tu": 0,
                    "Chuc_sac": 0,
                    "Chuc_viec": 0,
                    "Nha_tu_hanh": 0,
                    "So_tin_do": 0,
                    "Diem_nhom_SHTGTT": 0,
                }
            rel_data[rel][metric] += num_val

        for rel, metrics in rel_data.items():
            records.append(
                {
                    "Ma_STT": int(stt) if stt is not None else 0,
                    "Khu_vuc": curr_region,
                    "Dia_phuong": commune_clean,
                    "Ten_ton_giao": rel,
                    **metrics,
                }
            )

    return pd.DataFrame(records)


# ==============================================================================
# 3. KẾT NỐI FILE DỮ LIỆU & BỘ LỌC TƯƠNG TÁC NHANH (SIDEBAR)
# ==============================================================================
st.markdown(
    """
    <div class="main-header">
        <h1>📊 HỆ THỐNG THÔNG TIN THỐNG KÊ TÔN GIÁO TRÊN ĐỊA BÀN TỈNH GIA LAI</h1>
        <p>Báo cáo trực quan phục vụ công tác lãnh đạo, chỉ đạo (Nhấn chọn Tôn giáo hoặc Địa phương để xem chi tiết tức thì)</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.header("🎯 BỘ LỌC ĐIỀU HÀNH NHANH")

# Tự động tìm file trong thư mục hiện tại hoặc cho phép tải file lên
possible_paths = [
    DEFAULT_EXCEL_NAME,
    "/working_dir/c_03a432346a2bc9d7/user_uploaded_files/ttl=63d/user_uploaded/2026 a H. TỔNG HỢP SO LIEU CÁC TÔN GIÁO TREN DIA BAN TINH.xlsx",
]
found_path = next((p for p in possible_paths if os.path.exists(p)), None)

uploaded_file = st.sidebar.file_uploader(
    "📂 Cập nhật file Excel số liệu (tùy chọn):", type=["xlsx"]
)

if uploaded_file is not None:
    df = load_and_normalize_data(uploaded_file)
elif found_path is not None:
    df = load_and_normalize_data(found_path)
else:
    st.warning(
        f"⚠️ Chưa tìm thấy file `{DEFAULT_EXCEL_NAME}` trong thư mục chạy. "
        "Vui lòng chọn file Excel ở thanh công cụ bên trái."
    )
    st.stop()

# Chỉ hiển thị các tôn giáo có phát sinh số liệu ở nút bấm nhanh, hoặc chọn Tất cả
rel_totals = df.groupby("Ten_ton_giao")["So_tin_do"].sum().sort_values(ascending=False)
active_religions = rel_totals[rel_totals > 0].index.tolist()
all_religions = rel_totals.index.tolist()

# 1. Chọn Khu vực địa giới
region_options = ["Toàn tỉnh (Tất cả khu vực)"] + sorted(df["Khu_vuc"].unique().tolist())
selected_region = st.sidebar.radio(
    "1️⃣ Chọn Khu vực Địa giới:",
    options=region_options,
    index=0,
)

# 2. Chọn Tôn giáo (Danh sách thả xuống + Nút bấm nhanh trên màn hình chính)
religion_options = ["Tất cả tôn giáo"] + active_religions + [
    r for r in all_religions if r not in active_religions
]
selected_religion = st.sidebar.selectbox(
    "2️⃣ Chọn Tôn giáo:",
    options=religion_options,
    index=0,
)

# 3. Chọn Địa phương (Xã, Phường) - Tự động lọc theo Khu vực đã chọn
if selected_region != "Toàn tỉnh (Tất cả khu vực)":
    commune_pool = df[df["Khu_vuc"] == selected_region]["Dia_phuong"].unique().tolist()
else:
    commune_pool = df["Dia_phuong"].unique().tolist()

commune_options = ["Tất cả Xã, Phường"] + sorted(commune_pool)
selected_commune = st.sidebar.selectbox(
    "3️⃣ Chọn Địa phương (Xã, Phường):",
    options=commune_options,
    index=0,
)

if st.sidebar.button("🔄 Đặt lại toàn bộ bộ lọc", use_container_width=True):
    st.rerun()

# Thanh nút bấm nhanh (Quick-Select Pills) ngay đầu trang cho Lãnh đạo
st.markdown("**⚡ Chọn nhanh Tôn giáo trọng điểm (1 chạm):**")
quick_cols = st.columns(6)
quick_choices = ["Tất cả tôn giáo", "PHẬT GIÁO", "CÔNG GIÁO", "TIN LÀNH", "CAO ĐÀI", "PHẬT GIÁO HÒA HẢO"]
for idx, q_rel in enumerate(quick_choices):
    label = "🌐 Toàn tỉnh" if q_rel == "Tất cả tôn giáo" else q_rel.title()
    if quick_cols[idx].button(label, key=f"btn_{idx}", use_container_width=True):
        selected_religion = q_rel

# ==============================================================================
# 4. ÁP DỤNG BỘ LỌC DỮ LIỆU
# ==============================================================================
df_filtered = df.copy()
if selected_region != "Toàn tỉnh (Tất cả khu vực)":
    df_filtered = df_filtered[df_filtered["Khu_vuc"] == selected_region]
if selected_religion != "Tất cả tôn giáo":
    df_filtered = df_filtered[df_filtered["Ten_ton_giao"] == selected_religion]
if selected_commune != "Tất cả Xã, Phường":
    df_filtered = df_filtered[df_filtered["Dia_phuong"] == selected_commune]

# Tính toán các chỉ số cốt lõi (KPIs)
tong_tin_do = int(df_filtered["So_tin_do"].sum())
tong_co_so = int(df_filtered["Co_so_tho_tu"].sum())
tong_chuc_sac = int(df_filtered["Chuc_sac"].sum())
tong_chuc_viec = int(df_filtered["Chuc_viec"].sum())
tong_nha_tu_hanh = int(df_filtered["Nha_tu_hanh"].sum())
tong_to_chuc_tgtt = int(df_filtered["To_chuc_TGTT"].sum())
tong_diem_nhom = int(df_filtered["Diem_nhom_SHTGTT"].sum())
so_xa_co_tin_do = int(
    df_filtered.groupby("Dia_phuong")["So_tin_do"].sum().gt(0).sum()
)
tong_so_xa = int(df_filtered["Dia_phuong"].nunique())

# ==============================================================================
# 5. HIỂN THỊ THẺ THÔNG TIN NHANH (KPI CARDS)
# ==============================================================================
filter_summary_text = (
    f"Đang xem: **{selected_religion}** | Khu vực: **{selected_region}** | "
    f"Địa phương: **{selected_commune}**"
)
st.markdown(f"🔍 {filter_summary_text}")

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color: #2B6CB0;">
            <div class="kpi-title">👥 TỔNG SỐ TÍN ĐỒ</div>
            <div class="kpi-value">{tong_tin_do:,}</div>
            <div class="kpi-sub">Hiện diện tại {so_xa_co_tin_do}/{tong_so_xa} xã, phường</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color: #2F855A;">
            <div class="kpi-title">🏛️ CƠ SỞ THỜ TỰ</div>
            <div class="kpi-value">{tong_co_so:,}</div>
            <div class="kpi-sub">+ {tong_diem_nhom:,} điểm nhóm SHTGTT</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color: #C05621;">
            <div class="kpi-title">🎓 SỐ CHỨC SẮC</div>
            <div class="kpi-value">{tong_chuc_sac:,}</div>
            <div class="kpi-sub">Kèm {tong_chuc_viec:,} chức việc</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color: #6B46C1;">
            <div class="kpi-title">🏢 TỔ CHỨC TGTT & NHÀ TU HÀNH</div>
            <div class="kpi-value">{tong_to_chuc_tgtt:,}</div>
            <div class="kpi-sub">Tổ chức TGTT | {tong_nha_tu_hanh:,} nhà tu hành</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==============================================================================
# 6. TỔNG HỢP INSIGHTS TỰ ĐỘNG BÁO CÁO NHANH CHO LÃNH ĐẠO
# ==============================================================================
with st.expander("📌 BÁO CÁO NHANH 3 ĐIỂM NỔI BẬT TỪ DỮ LIỆU (EXECUTIVE INSIGHTS)", expanded=True):
    st.markdown(
        """
        <div class="insight-box">
        <b>1. Quy mô toàn tỉnh đạt 643.724 tín đồ, tập trung tuyệt đối (96,86%) ở 3 tôn giáo lớn:</b>
        Phật giáo dẫn đầu với <b>227.457 tín đồ</b> (35,33% - 461 cơ sở thờ tự), theo sát là Công giáo với <b>226.697 tín đồ</b> (35,22% - 175 cơ sở) và Tin Lành với <b>169.384 tín đồ</b> (26,31% - 51 cơ sở). Đạo Cao Đài đứng thứ tư với <b>19.384 tín đồ</b> (3,01% - 61 cơ sở). <i>(Lưu ý kỹ thuật: Bảng nhỏ cuối sheet gốc ghi 474.340 tín đồ do công thức SUM bỏ sót cột W144 của Tin Lành; chuẩn hóa từ 135 xã/phường khôi phục chính xác tổng số 643.724 tín đồ).</i><br><br>
        <b>2. Sự phân hóa rõ rệt theo trục Địa giới hành chính (Gia Lai Tây – Gia Lai Đông):</b>
        Khu vực <b>Gia Lai Tây (77 xã, phường)</b> chiếm tới <b>69,92% tổng số tín đồ</b> (450.108 người) và <b>97,19% điểm nhóm sinh hoạt tôn giáo tập trung</b> (519/534 điểm nhóm, chủ yếu của Tin Lành và Công giáo vùng đồng bào DTTS). Ngược lại, khu vực <b>Gia Lai Đông (58 xã, phường)</b> lại tập trung tới <b>66,49% cơ sở thờ tự</b> (502/755 cơ sở, nổi bật là 341 chùa/tịnh xá Phật giáo và 53 thánh thất Cao Đài) cùng <b>61,71% đội ngũ chức sắc</b> (880/1.426 vị).<br><br>
        <b>3. Chênh lệch lớn giữa nhu cầu sinh hoạt tín đồ và cơ sở thờ tự hợp pháp ở đạo Tin Lành:</b>
        Tin Lành có gần 170.000 tín đồ nhưng chỉ có <b>51 cơ sở thờ tự</b> (trung bình <b>3.321 tín đồ/cơ sở</b>, cao gấp 6,7 lần mức 493 tín đồ/cơ sở của Phật giáo), dẫn đến việc hình thành tới <b>298 điểm nhóm SHTGTT</b> (chiếm 55,8% tổng điểm nhóm toàn tỉnh, tập trung mạnh tại Phú Thiện, Phú Túc, KDang, Ia Phí, Ia Lâu). Hai đô thị trung tâm có mật độ tín đồ cao nhất tỉnh là <b>Phường Pleiku (29.656 tín đồ)</b> và <b>Phường Quy Nhơn (24.067 tín đồ)</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==============================================================================
# 7. HỆ THỐNG BIỂU ĐỒ ĐỘNG TƯƠNG TÁC (PLOTLY CHARTS)
# ==============================================================================
col_chart1, col_chart2 = st.columns([6, 4])

with col_chart1:
    st.subheader("📍 Phân bố theo Địa phương (Top 15 Xã, Phường)")
    metric_choice = st.radio(
        "Chọn chỉ số hiển thị trên biểu đồ:",
        options=[
            ("So_tin_do", "Số tín đồ"),
            ("Co_so_tho_tu", "Số cơ sở thờ tự"),
            ("Chuc_sac", "Số chức sắc"),
            ("Diem_nhom_SHTGTT", "Số điểm nhóm SHTGTT"),
        ],
        format_func=lambda x: x[1],
        horizontal=True,
    )
    metric_col, metric_label = metric_choice

    df_by_commune = (
        df_filtered.groupby(["Dia_phuong", "Khu_vuc"], as_index=False)[
            ["So_tin_do", "Co_so_tho_tu", "Chuc_sac", "Chuc_viec", "Nha_tu_hanh", "Diem_nhom_SHTGTT"]
        ]
        .sum()
        .sort_values(by=metric_col, ascending=False)
        .head(15)
    )

    fig_bar = px.bar(
        df_by_commune.sort_values(by=metric_col, ascending=True),
        x=metric_col,
        y="Dia_phuong",
        color="Khu_vuc",
        orientation="h",
        text=metric_col,
        color_discrete_map={"GIA LAI TÂY": "#2B6CB0", "GIA LAI ĐÔNG": "#DD6B20"},
        custom_data=["Khu_vuc", "So_tin_do", "Co_so_tho_tu", "Chuc_sac", "Chuc_viec", "Diem_nhom_SHTGTT"],
        labels={
            metric_col: metric_label,
            "Dia_phuong": "Xã, Phường",
            "Khu_vuc": "Khu vực",
        },
    )
    fig_bar.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b> (%{customdata[0]})<br>"
            "• Tín đồ: <b>%{customdata[1]:,.0f}</b><br>"
            "• Cơ sở thờ tự: <b>%{customdata[2]:,.0f}</b><br>"
            "• Chức sắc: <b>%{customdata[3]:,.0f}</b> | Chức việc: <b>%{customdata[4]:,.0f}</b><br>"
            "• Điểm nhóm SHTGTT: <b>%{customdata[5]:,.0f}</b><extra></extra>"
        ),
    )
    fig_bar.update_layout(
        height=480,
        margin=dict(l=10, r=40, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(title=metric_label, separatethousands=True),
        yaxis=dict(title=""),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col_chart2:
    if selected_religion == "Tất cả tôn giáo":
        st.subheader("🥧 Cơ cấu theo Tôn giáo")
        df_pie = (
            df_filtered.groupby("Ten_ton_giao", as_index=False)[
                ["So_tin_do", "Co_so_tho_tu", "Chuc_sac", "Diem_nhom_SHTGTT"]
            ]
            .sum()
        )
        df_pie = df_pie[df_pie[metric_col] > 0].sort_values(by=metric_col, ascending=False)

        fig_pie = px.pie(
            df_pie,
            values=metric_col,
            names="Ten_ton_giao",
            hole=0.42,
            color_discrete_sequence=px.colors.qualitative.Bold,
            custom_data=["So_tin_do", "Co_so_tho_tu", "Chuc_sac", "Diem_nhom_SHTGTT"],
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent+label",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "• Tỷ trọng: <b>%{percent}</b><br>"
                "• Tín đồ: <b>%{customdata[0]:,.0f}</b><br>"
                "• Cơ sở thờ tự: <b>%{customdata[1]:,.0f}</b><br>"
                "• Chức sắc: <b>%{customdata[2]:,.0f}</b><br>"
                "• Điểm nhóm SHTGTT: <b>%{customdata[3]:,.0f}</b><extra></extra>"
            ),
        )
        fig_pie.update_layout(
            height=480,
            margin=dict(l=10, r=10, t=30, b=10),
            showlegend=True,
            legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.subheader(f"⚖️ Phân bổ {selected_religion} theo Khu vực")
        df_reg = (
            df_filtered.groupby("Khu_vuc", as_index=False)[
                ["So_tin_do", "Co_so_tho_tu", "Chuc_sac", "Chuc_viec", "Nha_tu_hanh", "Diem_nhom_SHTGTT"]
            ]
            .sum()
        )
        fig_reg = go.Figure(
            data=[
                go.Bar(name="Tín đồ", x=df_reg["Khu_vuc"], y=df_reg["So_tin_do"], text=df_reg["So_tin_do"], marker_color="#2B6CB0"),
                go.Bar(name="Cơ sở thờ tự", x=df_reg["Khu_vuc"], y=df_reg["Co_so_tho_tu"], text=df_reg["Co_so_tho_tu"], marker_color="#2F855A"),
                go.Bar(name="Chức sắc", x=df_reg["Khu_vuc"], y=df_reg["Chuc_sac"], text=df_reg["Chuc_sac"], marker_color="#C05621"),
                go.Bar(name="Điểm nhóm SHTGTT", x=df_reg["Khu_vuc"], y=df_reg["Diem_nhom_SHTGTT"], text=df_reg["Diem_nhom_SHTGTT"], marker_color="#6B46C1"),
            ]
        )
        fig_reg.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig_reg.update_layout(
            barmode="group",
            height=480,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_reg, use_container_width=True)

# Biểu đồ Treemap toàn cảnh Phân bố Khu vực -> Địa phương -> Tôn giáo
st.subheader("🗺️ Bản đồ Cây (Treemap) Phân bố Tín đồ theo Khu vực & Địa phương")
df_tree = df_filtered[df_filtered["So_tin_do"] > 0]
if not df_tree.empty:
    fig_tree = px.treemap(
        df_tree,
        path=[px.Constant("Toàn tỉnh"), "Khu_vuc", "Dia_phuong", "Ten_ton_giao"],
        values="So_tin_do",
        color="Khu_vuc",
        color_discrete_map={"(?)": "#EDF2F7", "GIA LAI TÂY": "#3182CE", "GIA LAI ĐÔNG": "#DD6B20"},
        custom_data=["Co_so_tho_tu", "Chuc_sac", "Diem_nhom_SHTGTT"],
    )
    fig_tree.update_traces(
        hovertemplate=(
            "<b>%{label}</b><br>"
            "• Tổng tín đồ: <b>%{value:,.0f}</b><br>"
            "• Cơ sở thờ tự: <b>%{customdata[0]:,.0f}</b><br>"
            "• Chức sắc: <b>%{customdata[1]:,.0f}</b><br>"
            "• Điểm nhóm SHTGTT: <b>%{customdata[2]:,.0f}</b><extra></extra>"
        )
    )
    fig_tree.update_layout(height=450, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_tree, use_container_width=True)

# ==============================================================================
# 8. BẢNG DỮ LIỆU CHI TIẾT & XUẤT BÁO CÁO
# ==============================================================================
st.subheader("📋 Bảng Tổng hợp Số liệu Chi tiết (Theo bộ lọc hiện tại)")

df_table = (
    df_filtered.groupby(["Khu_vuc", "Dia_phuong", "Ten_ton_giao"], as_index=False)[
        [
            "To_chuc_TGTT",
            "Co_so_tho_tu",
            "Chuc_sac",
            "Chuc_viec",
            "Nha_tu_hanh",
            "So_tin_do",
            "Diem_nhom_SHTGTT",
        ]
    ]
    .sum()
)
# Ẩn các dòng hoàn toàn bằng 0 để bảng gọn gàng, dễ tra cứu
df_table_active = df_table[
    (df_table["So_tin_do"] > 0)
    | (df_table["Co_so_tho_tu"] > 0)
    | (df_table["To_chuc_TGTT"] > 0)
    | (df_table["Chuc_sac"] > 0)
    | (df_table["Diem_nhom_SHTGTT"] > 0)
].sort_values(by="So_tin_do", ascending=False)

df_display = df_table_active.rename(
    columns={
        "Khu_vuc": "Khu vực",
        "Dia_phuong": "Địa phương (Xã, Phường)",
        "Ten_ton_giao": "Tôn giáo",
        "To_chuc_TGTT": "Tổ chức TGTT",
        "Co_so_tho_tu": "Cơ sở thờ tự",
        "Chuc_sac": "Chức sắc",
        "Chuc_viec": "Chức việc",
        "Nha_tu_hanh": "Nhà tu hành",
        "So_tin_do": "Số tín đồ",
        "Diem_nhom_SHTGTT": "Điểm nhóm SHTGTT",
    }
)

st.dataframe(df_display, use_container_width=True, hide_index=True)

csv_bytes = df_display.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
st.download_button(
    label="📥 Tải xuống Bảng dữ liệu đã chuẩn hóa (.CSV)",
    data=csv_bytes,
    file_name="thong_ke_ton_giao_gia_lai_chuan_hoa.csv",
    mime="text/csv",
)