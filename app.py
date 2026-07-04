import io
import os
import re
import subprocess
import sys
from io import BytesIO
from pathlib import Path

import matplotlib.image as mpimg
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st
from matplotlib.offsetbox import AnnotationBbox, OffsetImage

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output_charts"
DPI = 150
SHIP_ZOOM = 0.10
DANGER_ZONE = (225, 315)

BASE_COLORS = [
    '#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c',
    '#e67e22', '#34495e', '#e91e63', '#00bcd4'
]
BASE_MARKERS = ['o', 's', '^', 'D', 'v', 'P', '*', 'X', 'h', '8']


st.set_page_config(page_title="Mooring Force Dashboard", page_icon="⚓", layout="wide")
st.title("⚓ Mooring Force Dashboard")
st.caption("Upload file Excel, xem biểu đồ trực tiếp trên web và lưu ảnh kết quả.")


def get_style(n_points):
    colors = (BASE_COLORS * ((n_points // len(BASE_COLORS)) + 1))[:n_points]
    markers = (BASE_MARKERS * ((n_points // len(BASE_MARKERS)) + 1))[:n_points]
    return colors, markers


def sanitize_filename(name: str) -> str:
    if not name:
        return 'mooring_force'
    s = re.sub(r"\s+", "_", name.strip())
    s = re.sub(r"[^A-Za-z0-9_\-\.]+", "", s)
    return s or 'mooring_force'


def get_chart_title(xls) -> tuple:
    title = ""
    xlabel = 'Wind Direction (°)'
    ylabel = 'Mooring Force (kN)'
    try:
        candidate_names = [s for s in xls.sheet_names if 'chart' in s.lower()]
        sheets_to_try = candidate_names + [s for s in xls.sheet_names if s not in candidate_names]
        for sheet in sheets_to_try:
            try:
                meta_df = pd.read_excel(xls, sheet_name=sheet, header=None, usecols="B", nrows=3)
            except Exception:
                continue
            if meta_df.empty:
                continue
            vals = [str(v).strip() for v in meta_df.iloc[:, 0].tolist()]
            if len(vals) > 0 and vals[0].lower() not in {"nan", "none", "", 'nan.0'}:
                title = vals[0]
                if len(vals) > 1 and vals[1].lower() not in {"nan", "none", ""}:
                    xlabel = vals[1]
                if len(vals) > 2 and vals[2].lower() not in {"nan", "none", ""}:
                    ylabel = vals[2]
                break
    except Exception:
        pass
    return title, xlabel, ylabel


def get_ship_config(xls, sheet) -> dict:
    result = {"show": False, "image": None, "heading": 0.0}
    try:
        sheet_name = sheet
        if sheet_name == 0:
            names = xls.sheet_names
            if names and names[0].lower() == "charttittle" and len(names) > 1:
                sheet_name = names[1]
            else:
                sheet_name = names[0]
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, usecols="I", nrows=3)
        if raw.empty or raw.shape[0] < 1:
            return result
        i1_val = str(raw.iloc[0, 0]).strip()
        if i1_val.lower() != "yes":
            return result
        result["show"] = True
        if raw.shape[0] >= 2:
            i2_val = str(raw.iloc[1, 0]).strip()
            if i2_val.lower() not in {"nan", "none", ""}:
                image_path = BASE_DIR / "images" / i2_val
                if image_path.exists():
                    result["image"] = image_path
        if raw.shape[0] >= 3:
            try:
                result["heading"] = float(raw.iloc[2, 0])
            except Exception:
                result["heading"] = 0.0
    except Exception:
        return result
    return result


def load_excel_data(uploaded_file, sheet_name: str | int, wind_col: str):
    bytes_data = uploaded_file.getvalue()
    with pd.ExcelFile(BytesIO(bytes_data)) as xls:
        if sheet_name == 0:
            if xls.sheet_names and xls.sheet_names[0].lower() == "charttittle" and len(xls.sheet_names) > 1:
                sheet_name = xls.sheet_names[1]
            else:
                sheet_name = xls.sheet_names[0]

        df = pd.read_excel(BytesIO(bytes_data), sheet_name=sheet_name, header=4)
        df.columns = [str(c).strip() for c in df.columns]

        if wind_col not in df.columns:
            match = [c for c in df.columns if 'wind' in c.lower()]
            if match:
                wind_col = match[0]
            else:
                raise ValueError(f"Không tìm thấy cột '{wind_col}'. Các cột hiện có: {list(df.columns)}")

        df = df.dropna(subset=[wind_col])
        df = df.sort_values(wind_col).reset_index(drop=True)

        wind_dir = df[wind_col].astype(float).tolist()
        data = {}
        for col in [c for c in df.columns if c != wind_col]:
            vals = pd.to_numeric(df[col], errors='coerce').fillna(0).tolist()
            data[col] = vals

        return sheet_name, wind_dir, data, xls


def load_ship_image(image_path: Path, heading_deg: float, zoom: float):
    if image_path is None or not image_path.exists():
        return None
    try:
        from scipy.ndimage import rotate as scipy_rotate
        img = mpimg.imread(str(image_path))
        rotated = scipy_rotate(img, angle=-heading_deg, reshape=True)
        return OffsetImage(rotated, zoom=zoom)
    except Exception:
        try:
            img = mpimg.imread(str(image_path))
            return OffsetImage(img, zoom=zoom)
        except Exception:
            return None


def create_line_chart(wind_dir, data, title, xlabel, ylabel):
    colors, markers = get_style(len(data))
    fig, ax = plt.subplots(figsize=(14, 7))
    for (label, values), color, marker in zip(data.items(), colors, markers):
        ax.plot(wind_dir, values, label=label, color=color, marker=marker, linewidth=2, markersize=6)
    ax.legend(title='Mooring Point', loc='upper left', fontsize=10)
    ax.set_xlabel(xlabel, fontsize=13)
    ax.set_ylabel(ylabel, fontsize=13)
    ax.set_title(title or 'Mooring Force', fontsize=15, fontweight='bold')
    ax.set_xticks(wind_dir)
    ax.set_xticklabels([f'{d}°' for d in wind_dir], rotation=45)
    ax.grid(True, linestyle='--', alpha=0.5)
    if data:
        ax.set_ylim(0, max(max(v) for v in data.values()) * 1.15)
    return fig


def create_polar_chart(wind_dir, data, title, xlabel, ship_cfg):
    colors, _ = get_style(len(data))
    angles = np.radians(wind_dir + [wind_dir[0]])
    fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(projection='polar'))
    ax.set_theta_zero_location('N')
    ax.set_theta_direction(-1)
    for (label, values), color in zip(data.items(), colors):
        vals = values + [values[0]]
        ax.plot(angles, vals, label=label, color=color, linewidth=2)
        ax.fill(angles, vals, alpha=0.05, color=color)
    ax.set_thetagrids(wind_dir, labels=[f'{d}°' for d in wind_dir], fontsize=9)
    ax.set_title(title or 'Mooring Force', fontsize=14, fontweight='bold', pad=20)
    legend_title = xlabel if xlabel else 'Mooring Point'
    ax.legend(loc='upper right', bbox_to_anchor=(1.35, 1.1), fontsize=10, title=legend_title)
    if ship_cfg.get("show") and ship_cfg.get("image") is not None:
        imagebox = load_ship_image(ship_cfg['image'], ship_cfg.get('heading', 0.0), SHIP_ZOOM)
        if imagebox is not None:
            ab = AnnotationBbox(imagebox, xy=(0, 0), xycoords='data', frameon=False, box_alignment=(0.5, 0.5))
            ax.add_artist(ab)
    return fig


def create_heatmap(wind_dir, data, title, xlabel, ylabel):
    df = pd.DataFrame(data, index=[f'{d}°' for d in wind_dir])
    fig, ax = plt.subplots(figsize=(max(8, len(data) * 1.5), 10))
    sns.heatmap(df, annot=True, fmt='.1f', cmap='YlOrRd', linewidths=0.5, ax=ax, cbar_kws={'label': 'Force (kN)'})
    ax.set_title(title or 'Mooring Force Heatmap', fontsize=14, fontweight='bold')
    ax.set_xlabel(xlabel if xlabel else 'Mooring Point', fontsize=12)
    ax.set_ylabel(ylabel if ylabel else 'Wind Direction', fontsize=12)
    return fig


def save_chart(fig, title, chart_type, output_dir: Path) -> tuple[Path, bytes]:
    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = f"{sanitize_filename(title or 'mooring_force')}_{chart_type.lower()}.png"
    out_path = output_dir / file_name
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=DPI, bbox_inches='tight')
    buf.seek(0)
    image_bytes = buf.getvalue()
    out_path.write_bytes(image_bytes)
    return out_path, image_bytes


with st.sidebar:
    st.header("⚙️ Input")
    uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xlsm', 'xls'], help="Drag and drop file Excel hoặc chọn từ máy.")
    chart_type = st.selectbox("Chart type", ["Polar", "Line", "Heatmap"], index=0)
    show_detail = st.checkbox("Show raw data preview", value=True)

if uploaded_file is None:
    st.info("Chọn file Excel để bắt đầu tạo biểu đồ.")
    st.stop()

try:
    with st.spinner("Đang đọc file Excel..."):
        sheet_name = 0
        wind_col = "Wind Direction"
        sheet_name, wind_dir, data, xls = load_excel_data(uploaded_file, sheet_name, wind_col)
        chart_title, xlabel, ylabel = get_chart_title(xls)
        ship_cfg = get_ship_config(xls, sheet_name)
except Exception as e:
    st.error(f"Không thể đọc file Excel: {e}")
    st.stop()

st.success(f"Đã đọc được {len(wind_dir)} điểm dữ liệu từ sheet '{sheet_name}'.")

col1, col2 = st.columns([1.2, 1])
with col1:
    st.subheader("📋 Thông tin file")
    st.write("- File:", uploaded_file.name)
    st.write("- Sheet đang dùng:", sheet_name)
    st.write("- Cột gió:", wind_col)
    st.write("- Tiêu đề biểu đồ:", chart_title or "(không có tiêu đề)")
with col2:
    st.subheader("🧭 Tùy chọn")
    if st.button("Open output folder"):
        try:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            if os.name == 'nt':
                os.startfile(str(OUTPUT_DIR))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(OUTPUT_DIR)])
            else:
                subprocess.Popen(['xdg-open', str(OUTPUT_DIR)])
            st.toast("Đã mở thư mục output.")
        except Exception as e:
            st.warning(f"Không mở được thư mục tự động: {e}")
            st.code(str(OUTPUT_DIR))

if show_detail:
    preview_df = pd.DataFrame({
        "Wind Direction": wind_dir,
        **{k: [v[i] for i in range(len(wind_dir))] for k, v in data.items()}
    })
    st.subheader("📊 Preview dữ liệu")
    st.dataframe(preview_df.head(20), use_container_width=True)

with st.spinner("Đang tạo biểu đồ..."):
    if chart_type == "Polar":
        fig = create_polar_chart(wind_dir, data, chart_title, xlabel, ship_cfg)
    elif chart_type == "Line":
        fig = create_line_chart(wind_dir, data, chart_title, xlabel, ylabel)
    else:
        fig = create_heatmap(wind_dir, data, chart_title, xlabel, ylabel)

    out_path, image_bytes = save_chart(fig, chart_title, chart_type, OUTPUT_DIR)

st.subheader("🖼️ Biểu đồ")
st.pyplot(fig)

st.download_button(
    label="💾 Save as PNG",
    data=image_bytes,
    file_name=out_path.name,
    mime="image/png"
)

st.info(f"Ảnh đã được lưu tại: {out_path}")
