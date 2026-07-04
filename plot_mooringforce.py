import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.image as mpimg
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import numpy as np
import pandas as pd
import seaborn as sns
import os
import sys
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# ─────────────────────────────────────────────
# 1. CẤU HÌNH — chỉnh tại đây
# ─────────────────────────────────────────────
EXCEL_FILE   = "input_data.xlsx"
SHEET_NAME   = 0
WIND_COL     = "Wind Direction"
VESSEL_TITLE = ""

DANGER_ZONE  = (225, 315)

BASE_COLORS  = ['#e74c3c','#3498db','#2ecc71','#f39c12','#9b59b6','#1abc9c',
                '#e67e22','#34495e','#e91e63','#00bcd4']
BASE_MARKERS = ['o','s','^','D','v','P','*','X','h','8']

OUTPUT_DIR   = "output_charts"
DPI          = 150

# Tỷ lệ zoom ảnh tàu trong chart (chỉnh tại đây nếu cần)
SHIP_ZOOM    = 0.10


# ─────────────────────────────────────────────
# 2. LOAD DỮ LIỆU TỪ EXCEL
# ─────────────────────────────────────────────
def load_excel(filepath: Path, sheet, wind_col: str):
    if not filepath.exists():
        print(f"[ERROR] Không tìm thấy file: {filepath}")
        sys.exit(1)

    if sheet == 0:
        with pd.ExcelFile(filepath) as xls:
            if xls.sheet_names and xls.sheet_names[0].lower() == "charttittle" and len(xls.sheet_names) > 1:
                sheet = xls.sheet_names[1]

    df = pd.read_excel(filepath, sheet_name=sheet, header=4)
    df.columns = df.columns.str.strip()

    if wind_col not in df.columns:
        match = [c for c in df.columns if 'wind' in c.lower()]
        if match:
            wind_col = match[0]
            print(f"[INFO] Dùng cột hướng gió: '{wind_col}'")
        else:
            print(f"[ERROR] Không tìm thấy cột '{wind_col}'. Các cột hiện có: {list(df.columns)}")
            sys.exit(1)

    df = df.dropna(subset=[wind_col])
    df = df.sort_values(wind_col).reset_index(drop=True)

    wind_dir = df[wind_col].astype(float).tolist()

    mooring_cols = [c for c in df.columns if c != wind_col]
    data = {}
    for col in mooring_cols:
        vals = pd.to_numeric(df[col], errors='coerce').fillna(0).tolist()
        data[col] = vals

    print(f"[OK] Đã load {len(wind_dir)} hướng gió, {len(data)} điểm neo: {list(data.keys())}")
    return wind_dir, data


# ─────────────────────────────────────────────
# 3. TIỆN ÍCH
# ─────────────────────────────────────────────
def get_style(n_points):
    colors  = (BASE_COLORS  * ((n_points // len(BASE_COLORS))  + 1))[:n_points]
    markers = (BASE_MARKERS * ((n_points // len(BASE_MARKERS)) + 1))[:n_points]
    return colors, markers


def ensure_output_dir(output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)


def get_chart_title(filepath: Path) -> str:
    title  = ""
    xlabel = 'Wind Direction (°)'
    ylabel = 'Mooring Force (kN)'
    try:
        with pd.ExcelFile(filepath) as xls:
            candidate_names = [s for s in xls.sheet_names if 'chart' in s.lower()]
            sheets_to_try   = candidate_names + [s for s in xls.sheet_names if s not in candidate_names]

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


def get_ship_config(filepath: Path, sheet) -> dict:
    """
    Đọc cấu hình ảnh tàu từ sheet data:
      - I1 : "Yes" / "No"  — có chèn ảnh không
      - I2 : tên file ảnh  — ví dụ "portside.png" hoặc "starboard.png"
      - I3 : heading angle — góc tàu so với True North (số thực)

    Trả về dict:
      {
        "show"    : bool,
        "image"   : Path | None,
        "heading" : float,
      }
    """
    result = {"show": False, "image": None, "heading": 0.0}

    try:
        with pd.ExcelFile(filepath) as xls:
            # Xác định đúng sheet name (xử lý trường hợp sheet=0)
            sheet_name = sheet
            if sheet_name == 0:
                names = xls.sheet_names
                # Bỏ qua sheet ChartTittle nếu là sheet đầu tiên
                if names and names[0].lower() == "charttittle" and len(names) > 1:
                    sheet_name = names[1]
                else:
                    sheet_name = names[0]

            # Đọc cột I (index 8), 3 hàng đầu (row 0–2), không dùng header
            raw = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                usecols="I",      # cột I
                nrows=3           # chỉ đọc 3 hàng đầu
            )

        if raw.empty or raw.shape[0] < 1:
            print("[WARN] Không đọc được cấu hình ảnh tàu từ cột I.")
            return result

        # ── I1: Yes / No ──────────────────────────────────────────────
        i1_val = str(raw.iloc[0, 0]).strip()
        if i1_val.lower() != "yes":
            print(f"[INFO] Ship image: OFF (I1 = '{i1_val}')")
            return result

        result["show"] = True

        # ── I2: Tên file ảnh ──────────────────────────────────────────
        if raw.shape[0] >= 2:
            i2_val = str(raw.iloc[1, 0]).strip()
            if i2_val.lower() not in {"nan", "none", ""}:
                image_path = BASE_DIR / "images" / i2_val
                if image_path.exists():
                    result["image"] = image_path
                    print(f"[INFO] Ship image: {image_path.name}")
                else:
                    print(f"[WARN] Không tìm thấy file ảnh: {image_path}")
            else:
                print("[WARN] I2 trống — không có tên file ảnh.")
        else:
            print("[WARN] Thiếu giá trị I2 (tên file ảnh).")

        # ── I3: Heading angle ─────────────────────────────────────────
        if raw.shape[0] >= 3:
            i3_val = raw.iloc[2, 0]
            try:
                result["heading"] = float(i3_val)
                print(f"[INFO] Ship heading: {result['heading']}°")
            except (ValueError, TypeError):
                print(f"[WARN] I3 không phải số hợp lệ ('{i3_val}') — dùng 0°")
                result["heading"] = 0.0
        else:
            print("[WARN] Thiếu giá trị I3 (heading) — dùng 0°")

    except Exception as e:
        print(f"[WARN] Lỗi đọc ship config từ Excel: {e}")

    return result


def sanitize_filename(name: str) -> str:
    if not name:
        return 'mooring_force'
    s = re.sub(r"\s+", "_", name.strip())
    s = re.sub(r"[^A-Za-z0-9_\-\.]+", "", s)
    return s or 'mooring_force'


def load_ship_image(image_path: Path, heading_deg: float, zoom: float):
    """
    Load và xoay ảnh tàu theo heading.
    Quy ước: ảnh gốc có mũi tàu hướng LÊN TRÊN (North).
    """
    if image_path is None or not image_path.exists():
        print(f"[WARN] Không tìm thấy ảnh tàu: {image_path}")
        return None

    try:
        from scipy.ndimage import rotate as scipy_rotate
        img = mpimg.imread(str(image_path))
        # scipy rotate: dương = ngược chiều kim đồng hồ
        # → dùng -heading để xoay đúng chiều kim đồng hồ theo hướng tàu
        rotated = scipy_rotate(img, angle=-heading_deg, reshape=True)
        return OffsetImage(rotated, zoom=zoom)
    except ImportError:
        print("[WARN] scipy chưa được cài — ảnh tàu không xoay. Chạy: pip install scipy")
        img = mpimg.imread(str(image_path))
        return OffsetImage(img, zoom=zoom)
    except Exception as e:
        print(f"[WARN] Không thể load ảnh tàu: {e}")
        return None


# ─────────────────────────────────────────────
# 4. BIỂU ĐỒ 1 — LINE CHART
# ─────────────────────────────────────────────
def plot_line_chart(wind_dir, data, title, xlabel, ylabel, output_dir: Path, danger_zone=None):
    colors, markers = get_style(len(data))
    fig, ax = plt.subplots(figsize=(14, 7))

    for (label, values), color, marker in zip(data.items(), colors, markers):
        ax.plot(wind_dir, values, label=label, color=color,
                marker=marker, linewidth=2, markersize=6)

    if danger_zone:
        ax.axvspan(*danger_zone, alpha=0.08, color='red')
        danger_patch = mpatches.Patch(color='red', alpha=0.3,
                                      label=f'Danger Zone ({danger_zone[0]}°–{danger_zone[1]}°)')
        handles, labels_leg = ax.get_legend_handles_labels()
        handles.append(danger_patch)
        ax.legend(handles=handles, title='Mooring Point', loc='upper left', fontsize=10)
    else:
        ax.legend(title='Mooring Point', loc='upper left', fontsize=10)

    for (label, values), color in zip(data.items(), colors):
        max_val = max(values)
        max_dir = wind_dir[values.index(max_val)]
        ax.annotate(f'{max_val}', xy=(max_dir, max_val),
                    xytext=(0, 8), textcoords='offset points',
                    ha='center', fontsize=7, color=color, fontweight='bold')

    ax.set_xlabel(xlabel, fontsize=13)
    ax.set_ylabel(ylabel, fontsize=13)
    ax.set_title(f'{title}' if title else '', fontsize=15, fontweight='bold')
    ax.set_xticks(wind_dir)
    ax.set_xticklabels([f'{d}°' for d in wind_dir], rotation=45)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.set_ylim(0, max(max(v) for v in data.values()) * 1.15)

    plt.tight_layout()
    fname = f"{sanitize_filename(title)}_line.png"
    out = output_dir / fname
    plt.savefig(str(out), dpi=DPI)
    print(f"[SAVED] {out}")
    plt.show()


# ─────────────────────────────────────────────
# 5. BIỂU ĐỒ 2 — POLAR CHART
# ─────────────────────────────────────────────
def plot_polar_chart(wind_dir, data, title, xlabel, output_dir: Path, ship_cfg: dict):
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
    ax.set_title(f'{title}' if title else 'Mooring Force',
                 fontsize=14, fontweight='bold', pad=20)
    legend_title = xlabel if xlabel else 'Mooring Point'
    ax.legend(loc='upper right', bbox_to_anchor=(1.35, 1.1), fontsize=10, title=legend_title)

    # ── Chèn ảnh tàu (nếu được bật từ Excel) ─────────────────────────
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
                box_alignment=(0.5, 0.5)
            )
            ax.add_artist(ab)
            print(f"[OK] Đã chèn ảnh '{ship_cfg['image'].name}' "
                  f"(heading={ship_cfg['heading']}°, zoom={SHIP_ZOOM})")

    plt.tight_layout()
    fname = f"{sanitize_filename(title)}_polar.png"
    out = output_dir / fname
    plt.savefig(str(out), dpi=DPI)
    print(f"[SAVED] {out}")
    plt.show()


# ─────────────────────────────────────────────
# 6. BIỂU ĐỒ 3 — HEATMAP
# ─────────────────────────────────────────────
def plot_heatmap(wind_dir, data, title, xlabel, ylabel, output_dir: Path):
    df = pd.DataFrame(data, index=[f'{d}°' for d in wind_dir])

    fig, ax = plt.subplots(figsize=(max(8, len(data) * 1.5), 10))
    sns.heatmap(df, annot=True, fmt='.1f', cmap='YlOrRd',
                linewidths=0.5, ax=ax, cbar_kws={'label': 'Force (kN)'})

    ax.set_title(f'{title}' if title else 'Mooring Force Heatmap', fontsize=14, fontweight='bold')
    ax.set_xlabel(xlabel if xlabel else 'Mooring Point', fontsize=12)
    ax.set_ylabel(ylabel if ylabel else 'Wind Direction', fontsize=12)

    plt.tight_layout()
    fname = f"{sanitize_filename(title)}_heatmap.png"
    out = output_dir / fname
    plt.savefig(str(out), dpi=DPI)
    print(f"[SAVED] {out}")
    plt.show()


# ─────────────────────────────────────────────
# 7. MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    excel_path = Path(EXCEL_FILE)
    if not excel_path.is_absolute():
        excel_path = BASE_DIR / excel_path

    output_dir_path = (BASE_DIR / OUTPUT_DIR
                       if not Path(OUTPUT_DIR).is_absolute()
                       else Path(OUTPUT_DIR))
    ensure_output_dir(output_dir_path)

    # Load metadata & dữ liệu
    chart_title, xlabel, ylabel = get_chart_title(excel_path)
    if not chart_title:
        chart_title = VESSEL_TITLE

    wind_dir, data = load_excel(excel_path, SHEET_NAME, WIND_COL)

    # Đọc cấu hình ảnh tàu từ Excel (I1, I2, I3)
    ship_cfg = get_ship_config(excel_path, SHEET_NAME)

    # Sinh 3 biểu đồ
    # plot_line_chart(wind_dir, data, chart_title, xlabel, ylabel, output_dir_path, DANGER_ZONE)
    plot_polar_chart(wind_dir, data, chart_title, xlabel, output_dir_path, ship_cfg)
    # plot_heatmap(wind_dir, data, chart_title, xlabel, ylabel, output_dir_path)

    print("\n✅ Hoàn tất! Tất cả biểu đồ đã lưu vào thư mục:", OUTPUT_DIR)
