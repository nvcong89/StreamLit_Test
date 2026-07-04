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
import optimoor_rtf_to_excel as optimoor
from matplotlib.offsetbox import AnnotationBbox, OffsetImage

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output_charts"
TASKS_DIR = BASE_DIR / "tasks"
MOORING_TASK_DIR = TASKS_DIR / "mooring_force"
MOORING_INPUT_DIR = MOORING_TASK_DIR / "input"
MOORING_OUTPUT_DIR = MOORING_TASK_DIR / "output"
OPTI_TASK_DIR = TASKS_DIR / "optimoor_rtf_to_excel"
OPTI_INPUT_DIR = OPTI_TASK_DIR / "input"
OPTI_OUTPUT_DIR = OPTI_TASK_DIR / "output"
TEMPLATE_DIR = BASE_DIR / "template"
TEMPLATE_INPUT_PATH = BASE_DIR / "images" / "template_input_data.xlsx"
OPTIMOOR_TEMPLATE_PATH = TEMPLATE_DIR / "Optimoor Tool-Master_Post-Processing.xlsm"
DPI = 150
SHIP_ZOOM = 0.10
DANGER_ZONE = (225, 315)

BASE_COLORS = [
    '#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c',
    '#e67e22', '#34495e', '#e91e63', '#00bcd4'
]
BASE_MARKERS = ['o', 's', '^', 'D', 'v', 'P', '*', 'X', 'h', '8']


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


def ensure_task_directories() -> None:
    for path in [TASKS_DIR, MOORING_TASK_DIR, MOORING_INPUT_DIR, MOORING_OUTPUT_DIR, OPTI_TASK_DIR, OPTI_INPUT_DIR, OPTI_OUTPUT_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def get_template_file_bytes() -> bytes:
    if TEMPLATE_INPUT_PATH.exists():
        return TEMPLATE_INPUT_PATH.read_bytes()
    return b""


def get_optimoor_template_bytes() -> bytes:
    if OPTIMOOR_TEMPLATE_PATH.exists():
        return OPTIMOOR_TEMPLATE_PATH.read_bytes()
    return b""


def build_task_tree_markdown() -> str:
    def list_files(path: Path) -> list[str]:
        if not path.exists():
            return ["(empty)"]
        files = sorted([p.name for p in path.glob("*") if p.is_file()])
        return files or ["(empty)"]

    task_tree = {
        "Mooring Force Dashboard": {
            "input": list_files(MOORING_INPUT_DIR),
            "output": list_files(MOORING_OUTPUT_DIR),
        },
        "Optimoor RTF → Excel": {
            "input": list_files(OPTI_INPUT_DIR),
            "output": list_files(OPTI_OUTPUT_DIR),
        },
    }

    lines = ["- **Tasks**"]
    for task_name, folders in task_tree.items():
        lines.append(f"  - **{task_name}**")
        for folder_name, files in folders.items():
            lines.append(f"    - {folder_name}: ")
            for filename in files:
                lines.append(f"      - {filename}")
    return "\n".join(lines)


def save_uploaded_files(uploaded_files, target_dir: Path) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    saved_paths = []
    for uploaded_file in uploaded_files:
        dest_path = target_dir / uploaded_file.name
        dest_path.write_bytes(uploaded_file.getvalue())
        saved_paths.append(dest_path)
    return saved_paths


def convert_rtf_to_excel(uploaded_rtf_files) -> list[Path]:
    ensure_task_directories()
    for path in OPTI_OUTPUT_DIR.glob("*.xlsx"):
        path.unlink(missing_ok=True)
    for path in OPTI_INPUT_DIR.glob("*.rtf"):
        path.unlink(missing_ok=True)

    saved_paths = save_uploaded_files(uploaded_rtf_files, OPTI_INPUT_DIR)
    rtf_files = sorted([p for p in OPTI_INPUT_DIR.glob("*.rtf") if p.is_file()])
    if not rtf_files:
        return []

    cases, line_tensions, fenders, hook_bollard_forces = [], [], [], []
    greatest_excursions, greatest_line_tensions = [], []
    greatest_berth_fender_thrusts, greatest_horizontal_bollard_forces = [], []

    for file_path in rtf_files:
        optimoor._process_file(
            file_path,
            cases,
            line_tensions,
            fenders,
            hook_bollard_forces,
            greatest_excursions,
            greatest_line_tensions,
            greatest_berth_fender_thrusts,
            greatest_horizontal_bollard_forces,
        )

        optimoor._build_and_save(
            save_path=OPTI_OUTPUT_DIR / f"{file_path.stem}.xlsx",
            cases=cases,
            line_tensions=line_tensions,
            fenders=fenders,
            hook_bollard_forces=hook_bollard_forces,
            greatest_excursions=greatest_excursions,
            greatest_line_tensions=greatest_line_tensions,
            greatest_berth_fender_thrusts=greatest_berth_fender_thrusts,
            greatest_horizontal_bollard_forces=greatest_horizontal_bollard_forces,
            print_label=f"[{file_path.name}]",
        )

    optimoor._build_and_save(
        save_path=OPTI_OUTPUT_DIR / "Optimoor_batch_results.xlsx",
        cases=cases,
        line_tensions=line_tensions,
        fenders=fenders,
        hook_bollard_forces=hook_bollard_forces,
        greatest_excursions=greatest_excursions,
        greatest_line_tensions=greatest_line_tensions,
        greatest_berth_fender_thrusts=greatest_berth_fender_thrusts,
        greatest_horizontal_bollard_forces=greatest_horizontal_bollard_forces,
        print_label="[Combined: Optimoor_batch_results.xlsx]",
    )

    return sorted([p for p in OPTI_OUTPUT_DIR.glob("*.xlsx") if p.is_file()])


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


def resolve_ship_image_path(image_name: str):
    if not image_name:
        return None

    raw_path = Path(str(image_name).strip())
    if raw_path.is_absolute() and raw_path.exists():
        return raw_path.resolve()

    candidates = []
    name = raw_path.name

    if raw_path.exists():
        candidates.append(raw_path.resolve())

    candidates.extend([
        BASE_DIR / "images" / name,
        BASE_DIR / name,
        BASE_DIR / "images" / raw_path,
        BASE_DIR / raw_path,
        Path.cwd() / "images" / name,
        Path.cwd() / name,
        Path.cwd() / raw_path,
        BASE_DIR / "MooringForcePlotter" / "images" / name,
        BASE_DIR / "MooringForcePlotter" / name,
    ])

    for root in [BASE_DIR, BASE_DIR / "images", Path.cwd()]:
        if root is None or not root.exists():
            continue
        try:
            for match in root.rglob(name):
                if match.is_file():
                    candidates.append(match.resolve())
        except Exception:
            continue

    seen = set()
    for path in candidates:
        try:
            if path.exists():
                resolved = path.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    return resolved
        except Exception:
            continue

    return None


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
                image_path = resolve_ship_image_path(i2_val)
                if image_path is not None:
                    result["image"] = image_path
                    st.session_state["ship_debug"] = f"Resolved: {image_path}"
                else:
                    st.session_state["ship_debug"] = f"Missing image file: {i2_val}"

        if raw.shape[0] >= 3:
            i3_val = raw.iloc[2, 0]
            try:
                result["heading"] = float(i3_val)
            except (ValueError, TypeError):
                result["heading"] = 0.0
    except Exception as exc:
        st.session_state["ship_debug"] = f"Excel config error: {exc}"
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
    except ImportError:
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
        imagebox = load_ship_image(
            ship_cfg["image"],
            ship_cfg.get("heading", 0.0),
            SHIP_ZOOM
        )
        if imagebox is not None:
            ab = AnnotationBbox(
                imagebox,
                xy=(0, 0),
                xycoords='data',
                frameon=False,
                box_alignment=(0.5, 0.5),
            )
            #ab.set_zorder(20)
            ax.add_artist(ab)
            st.session_state["ship_debug"] = f"Loaded: {ship_cfg['image']}"
        else:
            st.session_state["ship_debug"] = f"Failed to load: {ship_cfg['image']}"

    fig.tight_layout()
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

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=DPI, bbox_inches='tight')
    buf.seek(0)
    image_bytes = buf.getvalue()

    out_path.write_bytes(image_bytes)
    return out_path, image_bytes


def main():
    ensure_task_directories()
    template_bytes = get_template_file_bytes()
    optimoor_template_bytes = get_optimoor_template_bytes()

    st.set_page_config(page_title="CN Toolkit", page_icon="⚓", layout="wide")
    st.sidebar.header("About me")
    st.sidebar.markdown("**Author:** Nguyen Van Cong  \n"   
                        "**Location:** Vietnam  \n"    
                        "**Phone:** 0979404641  \n"  
                        "**Email:** nvcong89@live.com")
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ Navigation")
    task = st.sidebar.radio("Select task", ["Mooring Force Dashboard", "Optimoor RTF to Excel"])
    st.sidebar.markdown("---")
    

    if task == "Mooring Force Dashboard":
        st.title("CN Toolkit")
        st.caption("Upload Excel files, view charts live on the web, and save results.")

        uploaded_file = st.sidebar.file_uploader(
            "Upload Excel file",
            type=['xlsx', 'xlsm', 'xls'],
            help="Drag and drop file Excel hoặc chọn từ máy.",
        )
        chart_type = st.sidebar.selectbox("Chart type", ["Polar", "Line", "Heatmap"], index=0)
        show_detail = st.sidebar.checkbox("Show raw data preview", value=True)

        if template_bytes:
            st.sidebar.download_button(
                label="⬇️ Download Excel template",
                data=template_bytes,
                file_name="template_input_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.sidebar.warning("Không tìm thấy file template trong thư mục images.")

        with st.expander("📘 Hướng dẫn sử dụng", expanded=True):
            st.markdown(
                """
                1. Tải file template Excel mẫu bằng nút bên trái.
                2. Điền dữ liệu cho cột Wind Direction và các mooring points.
                3. Nếu muốn hiện ảnh tàu trên biểu đồ, giữ giá trị ở ô I1 = Yes, I2 là tên file ảnh, I3 là góc quay.
                4. Upload file Excel lên và chọn loại biểu đồ muốn xem.
                5. Nhấn Save as PNG để tải ảnh hoặc Open output folder để xem kết quả đã lưu.
                """
            )

        with st.expander("🗂️ Task folder structure", expanded=True):
            st.markdown(build_task_tree_markdown())

        if uploaded_file is not None:
            save_uploaded_files([uploaded_file], MOORING_INPUT_DIR)
        else:
            st.info("Chọn file Excel để bắt đầu tạo biểu đồ.")
            return

        try:
            with st.spinner("Đang đọc file Excel..."):
                sheet_name = 0
                wind_col = "Wind Direction"
                sheet_name, wind_dir, data, xls = load_excel_data(uploaded_file, sheet_name, wind_col)
                chart_title, xlabel, ylabel = get_chart_title(xls)
                ship_cfg = get_ship_config(xls, sheet_name)
        except Exception as e:
            st.error(f"Không thể đọc file Excel: {e}")
            return

        st.success(f"Đã đọc được {len(wind_dir)} điểm dữ liệu từ sheet '{sheet_name}'.")
        if "ship_debug" in st.session_state:
            st.caption(st.session_state["ship_debug"])

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
        st.pyplot(fig, use_container_width=True)
        st.download_button(
            label="💾 Save as PNG",
            data=image_bytes,
            file_name=out_path.name,
            mime="image/png"
        )
        st.info(f"Ảnh đã được lưu tại: {out_path}")

    else:
        st.title("🧾 Optimoor RTF to Excel")
        st.caption("Upload .rtf files để chuyển đổi sang Excel.")

        uploaded_rtf_files = st.sidebar.file_uploader(
            "Upload .rtf files",
            type=['rtf'],
            accept_multiple_files=True,
            help="Drag and drop one or more .rtf files để chuyển sang Excel.",
        )
        if optimoor_template_bytes:
            st.sidebar.download_button(
                label="⬇️ Download Optimoor template",
                data=optimoor_template_bytes,
                file_name="Optimoor Tool-Master_Post-Processing.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            )
        else:
            st.sidebar.warning("Không tìm thấy file Optimoor template.")

        with st.expander("📘 Hướng dẫn sử dụng", expanded=True):
            st.markdown(
                """
                1. Tải file template Optimoor bằng nút bên trái.
                2. Upload một hoặc nhiều file .rtf vào ứng dụng.
                3. Nhấn Convert .rtf to Excel để tạo file Excel.
                4. Tải file Excel đầu ra xuống từ các nút hiện ra.
                """
            )

        with st.expander("🗂️ Task folder structure", expanded=True):
            st.markdown(build_task_tree_markdown())

        if st.button("Convert .rtf to Excel"):
            if uploaded_rtf_files:
                with st.spinner("Đang chuyển đổi RTF sang Excel..."):
                    output_paths = convert_rtf_to_excel(uploaded_rtf_files)
                if output_paths:
                    st.success(f"Đã tạo {len(output_paths)} file Excel")
                    for output_path in output_paths:
                        with open(output_path, 'rb') as f:
                            st.download_button(
                                label=f"⬇️ Download {output_path.name}",
                                data=f.read(),
                                file_name=output_path.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                else:
                    st.warning("Không có file .rtf nào được xử lý.")
            else:
                st.info("Vui lòng upload ít nhất một file .rtf.")


if __name__ == "__main__":
    main()
