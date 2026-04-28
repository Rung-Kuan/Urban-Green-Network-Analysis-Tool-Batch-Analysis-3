import io
import math
from pathlib import Path
from importlib.util import find_spec
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

try:
    import pydeck as pdk
except ImportError:
    pdk = None


st.set_page_config(
    page_title="城市空品植生綠網淨化單元分析工具",
    page_icon="🌿",
    layout="wide",
)


# ============================================================
# 一、欄位設定
# ============================================================

REQUIRED_COLUMNS = [
    "代碼",
    "縣市",
    "鄉鎮市區",
    "基地名稱",
    "TWD97X",
    "TWD97Y",
    "距主要道路距離(公尺)",
    "土地權屬",
    "管理機關",
    "基地面積(公頃)",
    "基地長度(公里)",
    "短期推動性",
    "開放可及性",
    "基地內部是否有停留活動空間",
    "是否有步行通學騎行路徑",
    "是否有邊界防護需求",
    "道路周邊是否有敏感受體或社區活動空間",
    "是否有短期事件",
    "短期事件是否影響內部空間",
    "短期事件是否影響通行路徑",
    "短期事件是否位於邊界或鄰近受體",
]

AUTO_CONTEXT_COLUMNS = [
    "鄉鎮市區人口密度(人/平方公里)",
    "500公尺內工廠數",
    "500公尺內敏感受體數",
    "500公尺內生活節點數",
    "最近綠化單元距離(公尺)",
    "500公尺內其他綠化單元數",
    "1000公尺內其他綠化單元數",
    "1000公尺內其他綠化單元總面積(公頃)",
]

NUMERIC_COLUMNS = [
    "TWD97X",
    "TWD97Y",
    "距主要道路距離(公尺)",
    "基地面積(公頃)",
    "基地長度(公里)",
] + AUTO_CONTEXT_COLUMNS

YES_NO_COLUMNS = [
    "基地內部是否有停留活動空間",
    "是否有步行通學騎行路徑",
    "是否有邊界防護需求",
    "道路周邊是否有敏感受體或社區活動空間",
    "是否有短期事件",
    "短期事件是否影響內部空間",
    "短期事件是否影響通行路徑",
    "短期事件是否位於邊界或鄰近受體",
]

LOOKUP_FILES = {
    "人口密度": "town_lookup.csv",
    "工廠點位": "factory_points.csv",
    "空品淨化區": "green_air_sites.csv",
    "綠牆": "green_walls.csv",
    "各級學校": "schools.csv",
    "醫療院所": "medical_facilities.csv",
}

NODE_COLORS: Dict[str, List[int]] = {
    "高潛力": [0, 120, 60, 190],
    "中潛力": [255, 170, 0, 180],
    "基礎潛力": [160, 160, 160, 160],
}

LINK_COLORS: Dict[str, List[int]] = {
    "高潛力": [0, 90, 180, 190],
    "中潛力": [255, 170, 0, 180],
    "基礎潛力": [160, 160, 160, 160],
}

ROLE_COLORS: Dict[str, List[int]] = {
    "A1": [46, 204, 113, 180],
    "A2": [26, 188, 156, 180],
    "A3": [39, 174, 96, 180],
    "B1": [52, 152, 219, 180],
    "B2": [41, 128, 185, 180],
    "B3": [31, 97, 141, 190],
    "C1": [230, 126, 34, 180],
    "C2": [211, 84, 0, 180],
    "C3": [192, 57, 43, 190],
    "未判定": [127, 140, 141, 160],
}

ROLE_INFO = {
    "A1": {
        "情境名稱": "A1 多源逸散背景下之長時間停留暴露情境",
        "主要功能": "攔截與沉降、心理／行為調節",
        "適用性": "可導入",
        "管理方式": "以提升停留空間品質為主，維持植栽覆蓋完整、減少裸露地表，並定期檢查積塵、植栽健康與下層覆蓋情形。若問題主要為區域性 PM2.5 或二次污染，植生僅宜作為補強措施。",
        "追蹤重點": "積塵情形、裸露地變化、葉面與下層植栽狀況、停留空間舒適性。",
    },
    "A2": {
        "情境名稱": "A2 多源逸散背景下之短時間移動暴露情境",
        "主要功能": "擾流與稀釋、心理／行為調節",
        "適用性": "可導入",
        "管理方式": "以降低通行者接觸高暴露點的時間與頻率為主，優先檢討動線調整、遮蔭、帶狀綠化與路徑導引。",
        "追蹤重點": "通行舒適性、路徑使用情形、是否仍有高暴露停留點。",
    },
    "A3": {
        "情境名稱": "A3 多源逸散背景下之邊界暴露情境",
        "主要功能": "阻隔與緩衝、攔截與沉降",
        "適用性": "優先導入",
        "管理方式": "以建立連續、多層次、穩定的邊界植栽帶為主，降低源區污染向受體區直接傳輸。",
        "追蹤重點": "邊界連續性、缺口、植栽密度、下風處粉塵或異味感受。",
    },
    "B1": {
        "情境名稱": "B1 線源擾動背景下之長時間停留暴露情境",
        "主要功能": "阻隔與緩衝、攔截與沉降",
        "適用性": "可導入",
        "管理方式": "以道路與受體之間的邊界防護為主，配置應掌握「可防護而不封閉」原則，避免因過密配置造成污染停滯。",
        "追蹤重點": "植栽是否位於道路與受體之間、通風條件、近道路受體使用情形、粉塵感受變化。",
    },
    "B2": {
        "情境名稱": "B2 線源擾動背景下之短暫通行暴露情境",
        "主要功能": "擾流與稀釋、心理／行為調節",
        "適用性": "可導入",
        "管理方式": "以改善通行舒適性、遮蔭、分隔人流與車流為主，避免讓人流更靠近污染源或阻礙通風。",
        "追蹤重點": "通行動線合理性、植栽是否影響通風、遮蔭與視覺緩衝效果。",
    },
    "B3": {
        "情境名稱": "B3 線源擾動背景下之邊界暴露情境",
        "主要功能": "阻隔與緩衝",
        "適用性": "優先導入",
        "管理方式": "以連續植栽帶降低道路污染直接傳輸，並依道路尺度、受體位置與風向條件調整高度、寬度與孔隙度。",
        "追蹤重點": "邊界完整性、通風條件、受體側粉塵或異味感受、植栽有效防護位置。",
    },
    "C1": {
        "情境名稱": "C1 事件型或作業型污染對長時間停留受體之影響情境",
        "主要功能": "攔截與沉降、阻隔與緩衝",
        "適用性": "補強使用",
        "管理方式": "不建議以植生作為優先管理工具，應優先採取源頭抑制、覆蓋、灑水、工法改善、排放控制及作業管理。植生僅作為視覺遮蔽、局部緩衝或生活圈環境整合的補強措施。",
        "追蹤重點": "事件是否持續、源頭控制是否落實、植生是否被誤用為主要改善手段。",
    },
    "C2": {
        "情境名稱": "C2 事件型污染影響短時間移動暴露之情境",
        "主要功能": "擾流與稀釋、心理／行為調節",
        "適用性": "補強使用",
        "管理方式": "原則上不建議作為植生淨化優先導入類型，應以事件管理、作業時間調整、動線管制、臨時防制設施與源頭控制為優先。植生可輔助遮蔭、視覺緩衝或動線整理。",
        "追蹤重點": "臨時動線是否避開污染熱點、事件結束後是否需調整或移除臨時措施。",
    },
    "C3": {
        "情境名稱": "C3 事件型或間歇型污染源與受體之邊界暴露情境",
        "主要功能": "阻隔與緩衝、心理／行為調節",
        "適用性": "慎用／補強使用",
        "管理方式": "可作為周界防護、視覺遮蔽與感知緩衝，但不可取代源頭污染控制，應與工程防制、作業管理及排放控制併行。",
        "追蹤重點": "邊界完整性、下風處受體感受、陳情頻率、事件高峰期間暴露變化。",
    },
}


# ============================================================
# 二、通用工具函式
# ============================================================

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    data.columns = [str(c).strip() for c in data.columns]
    return data


def check_required_columns(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """批次上傳只強制檢查最小必要欄位。

    其他欄位若缺漏，系統會以預設值補上，避免舊版表單或簡化表單直接失敗。
    """
    core_required = [
        "代碼",
        "縣市",
        "鄉鎮市區",
        "基地名稱",
        "TWD97X",
        "TWD97Y",
    ]
    missing = [col for col in core_required if col not in df.columns]
    return len(missing) == 0, missing


def ensure_input_defaults(df: pd.DataFrame) -> pd.DataFrame:
    """補齊非核心欄位，讓批次上傳更耐受。"""
    data = df.copy()

    defaults = {
        "距主要道路距離(公尺)": 999,
        "土地權屬": "其他／待確認",
        "管理機關": "",
        "基地面積(公頃)": np.nan,
        "基地長度(公里)": np.nan,
        "短期推動性": "待評估",
        "開放可及性": "待確認",
    }

    for col, default_value in defaults.items():
        if col not in data.columns:
            data[col] = default_value

    for col in YES_NO_COLUMNS:
        if col not in data.columns:
            data[col] = "未知"

    return data


def clean_numeric(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if value in ["", "-", "—", "無", "NA", "N/A", "未知", "不明"]:
            return np.nan
    return pd.to_numeric(value, errors="coerce")


def normalize_yes_no(value) -> str:
    if pd.isna(value):
        return "未知"
    text = str(value).strip().lower()
    if text in ["是", "yes", "y", "true", "1", "有", "符合"]:
        return "是"
    if text in ["否", "no", "n", "false", "0", "無", "不符合"]:
        return "否"
    return "未知"


def is_yes(value) -> bool:
    return normalize_yes_no(value) == "是"


def classify_level(score: float, high_cut: float, medium_cut: float) -> str:
    if score >= high_cut:
        return "高潛力"
    if score >= medium_cut:
        return "中潛力"
    return "基礎潛力"


def twd97_to_wgs84(x, y, county="") -> Tuple[float, float]:
    if pd.isna(x) or pd.isna(y):
        return np.nan, np.nan

    lon0_degree = 119 if str(county).strip() in ["金門縣"] else 121
    a = 6378137.0
    b = 6356752.314245
    lon0 = math.radians(lon0_degree)
    k0 = 0.9999
    dx = 250000.0
    e = math.sqrt(1 - (b / a) ** 2)
    x = float(x) - dx
    y = float(y)
    m = y / k0
    mu = m / (a * (1 - e**2 / 4 - 3 * e**4 / 64 - 5 * e**6 / 256))
    e1 = (1 - math.sqrt(1 - e**2)) / (1 + math.sqrt(1 - e**2))
    j1 = 3 * e1 / 2 - 27 * e1**3 / 32
    j2 = 21 * e1**2 / 16 - 55 * e1**4 / 32
    j3 = 151 * e1**3 / 96
    j4 = 1097 * e1**4 / 512
    fp = mu + j1 * math.sin(2 * mu) + j2 * math.sin(4 * mu) + j3 * math.sin(6 * mu) + j4 * math.sin(8 * mu)
    e2 = e**2 / (1 - e**2)
    c1 = e2 * math.cos(fp) ** 2
    t1 = math.tan(fp) ** 2
    r1 = a * (1 - e**2) / ((1 - e**2 * math.sin(fp) ** 2) ** 1.5)
    n1 = a / math.sqrt(1 - e**2 * math.sin(fp) ** 2)
    d = x / (n1 * k0)
    q1 = n1 * math.tan(fp) / r1
    q2 = d**2 / 2
    q3 = (5 + 3 * t1 + 10 * c1 - 4 * c1**2 - 9 * e2) * d**4 / 24
    q4 = (61 + 90 * t1 + 298 * c1 + 45 * t1**2 - 3 * c1**2 - 252 * e2) * d**6 / 720
    lat = fp - q1 * (q2 - q3 + q4)
    q5 = d
    q6 = (1 + 2 * t1 + c1) * d**3 / 6
    q7 = (5 - 2 * c1 + 28 * t1 - 3 * c1**2 + 8 * e2 + 24 * t1**2) * d**5 / 120
    lon = lon0 + (q5 - q6 + q7) / math.cos(fp)
    return math.degrees(lat), math.degrees(lon)


def euclidean_distances_m(x: float, y: float, points: pd.DataFrame) -> pd.Series:
    return np.sqrt((points["TWD97X"] - x) ** 2 + (points["TWD97Y"] - y) ** 2)


def read_optional_csv(filename: str, required_cols: List[str]) -> pd.DataFrame:
    """讀取外部對照 CSV。

    外部資料來源可能有 UTF-8、Big5、CP950、BOM 或少數壞字元。
    這裡用多組編碼依序嘗試；若仍失敗，最後以 replacement 方式讀取，
    避免因單一外部資料檔編碼問題造成整個 App 中斷。
    """
    if not Path(filename).exists():
        return pd.DataFrame(columns=required_cols)

    encodings = ["utf-8-sig", "utf-8", "cp950", "big5", "latin1"]
    last_error = None

    for enc in encodings:
        try:
            df = pd.read_csv(filename, encoding=enc, low_memory=False)
            df = normalize_columns(df)
            for col in required_cols:
                if col not in df.columns:
                    df[col] = np.nan
            return df
        except Exception as e:
            last_error = e

    try:
        # 最後保底：用 Python engine 與 errors='replace' 避免壞字元中斷。
        with open(filename, "r", encoding="utf-8", errors="replace") as f:
            df = pd.read_csv(f, engine="python", on_bad_lines="skip")
        df = normalize_columns(df)
        for col in required_cols:
            if col not in df.columns:
                df[col] = np.nan
        return df
    except Exception:
        # 若真的完全無法讀取，回傳空表，並讓系統顯示自動補值提醒。
        return pd.DataFrame(columns=required_cols)


# ============================================================
# 三、外部資料讀取與自動補算
# ============================================================

@st.cache_data
def load_town_lookup() -> pd.DataFrame:
    df = read_optional_csv("town_lookup.csv", ["縣市", "鄉鎮市區", "鄉鎮市區人口密度(人/平方公里)"])
    if not df.empty:
        df["鄉鎮市區人口密度(人/平方公里)"] = df["鄉鎮市區人口密度(人/平方公里)"].apply(clean_numeric)
    return df


@st.cache_data
def load_factory_points() -> pd.DataFrame:
    df = read_optional_csv("factory_points.csv", ["TWD97X", "TWD97Y"])
    if not df.empty:
        df["TWD97X"] = df["TWD97X"].apply(clean_numeric)
        df["TWD97Y"] = df["TWD97Y"].apply(clean_numeric)
        df = df.dropna(subset=["TWD97X", "TWD97Y"])
    return df


@st.cache_data
def load_life_nodes() -> pd.DataFrame:
    schools = read_optional_csv("schools.csv", ["節點名稱", "TWD97X", "TWD97Y"])
    medical = read_optional_csv("medical_facilities.csv", ["節點名稱", "TWD97X", "TWD97Y"])

    if not schools.empty:
        schools["節點類型"] = schools["節點類型"] if "節點類型" in schools.columns else "學校"
        schools["資料來源"] = "各級學校"
    if not medical.empty:
        medical["節點類型"] = medical["節點類型"] if "節點類型" in medical.columns else "醫療院所"
        medical["資料來源"] = "醫療院所"

    df = pd.concat([schools, medical], ignore_index=True)
    if not df.empty:
        df["TWD97X"] = df["TWD97X"].apply(clean_numeric)
        df["TWD97Y"] = df["TWD97Y"].apply(clean_numeric)
        df = df.dropna(subset=["TWD97X", "TWD97Y"])
    return df


@st.cache_data
def load_green_units() -> pd.DataFrame:
    air = read_optional_csv("green_air_sites.csv", ["綠化單元代碼", "綠化單元名稱", "TWD97X", "TWD97Y", "基地面積(公頃)"])
    walls = read_optional_csv("green_walls.csv", ["綠化單元代碼", "綠化單元名稱", "TWD97X", "TWD97Y", "綠牆面積(平方公尺)"])

    if not air.empty:
        air["資料類型"] = "空品淨化區"
        air["基地面積(公頃)"] = air["基地面積(公頃)"].apply(clean_numeric)

    if not walls.empty:
        walls["資料類型"] = "綠牆"
        walls["綠牆面積(平方公尺)"] = walls["綠牆面積(平方公尺)"].apply(clean_numeric)
        walls["基地面積(公頃)"] = walls["綠牆面積(平方公尺)"] / 10000

    cols = ["綠化單元代碼", "綠化單元名稱", "TWD97X", "TWD97Y", "基地面積(公頃)", "資料類型"]
    for df in [air, walls]:
        for col in cols:
            if col not in df.columns:
                df[col] = np.nan

    green = pd.concat([air[cols], walls[cols]], ignore_index=True)
    if not green.empty:
        green["TWD97X"] = green["TWD97X"].apply(clean_numeric)
        green["TWD97Y"] = green["TWD97Y"].apply(clean_numeric)
        green["基地面積(公頃)"] = green["基地面積(公頃)"].apply(clean_numeric)
        green = green.dropna(subset=["TWD97X", "TWD97Y"])
    return green


def lookup_population_density(county: str, town: str, town_lookup: pd.DataFrame):
    if town_lookup.empty:
        return np.nan
    matched = town_lookup[
        (town_lookup["縣市"].astype(str).str.strip() == str(county).strip())
        & (town_lookup["鄉鎮市區"].astype(str).str.strip() == str(town).strip())
    ]
    if matched.empty:
        return np.nan
    return matched.iloc[0]["鄉鎮市區人口密度(人/平方公里)"]


def compute_factory_count_500(x, y, factories: pd.DataFrame):
    if pd.isna(x) or pd.isna(y) or factories.empty:
        return np.nan
    return int((euclidean_distances_m(float(x), float(y), factories) <= 500).sum())


def compute_life_node_count_500(x, y, life_nodes: pd.DataFrame):
    if pd.isna(x) or pd.isna(y) or life_nodes.empty:
        return np.nan
    return int((euclidean_distances_m(float(x), float(y), life_nodes) <= 500).sum())


def compute_green_metrics(row: pd.Series, green_units: pd.DataFrame) -> Dict[str, float]:
    x, y = row.get("TWD97X", np.nan), row.get("TWD97Y", np.nan)
    site_code = str(row.get("代碼", "")).strip()
    site_name = str(row.get("基地名稱", "")).strip()

    result = {
        "最近綠化單元距離(公尺)": np.nan,
        "500公尺內其他綠化單元數": np.nan,
        "1000公尺內其他綠化單元數": np.nan,
        "1000公尺內其他綠化單元總面積(公頃)": np.nan,
    }

    if pd.isna(x) or pd.isna(y) or green_units.empty:
        return result

    candidates = green_units.copy()

    if site_code:
        candidates = candidates[candidates["綠化單元代碼"].astype(str).str.strip() != site_code]
    if site_name:
        candidates = candidates[candidates["綠化單元名稱"].astype(str).str.strip() != site_name]

    if candidates.empty:
        return result

    distances = euclidean_distances_m(float(x), float(y), candidates)

    # 距離 1 公尺以內視為同一基地，不列入其他綠化單元計算。
    candidates = candidates[distances > 1].copy()
    distances = distances[distances > 1]

    if candidates.empty:
        return result

    within_500 = distances <= 500
    within_1000 = distances <= 1000

    result["最近綠化單元距離(公尺)"] = float(distances.min())
    result["500公尺內其他綠化單元數"] = int(within_500.sum())
    result["1000公尺內其他綠化單元數"] = int(within_1000.sum())
    result["1000公尺內其他綠化單元總面積(公頃)"] = float(
        candidates.loc[within_1000, "基地面積(公頃)"].fillna(0).sum()
    )
    return result


def fill_if_missing(data: pd.DataFrame, col: str, values: pd.Series) -> pd.DataFrame:
    if col not in data.columns:
        data[col] = np.nan
    data[col] = data[col].apply(clean_numeric)
    data[col] = data[col].where(~data[col].isna(), values)
    return data


def auto_enrich_context_fields(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    data = df.copy()
    notes = []

    town_lookup = load_town_lookup()
    factories = load_factory_points()
    life_nodes = load_life_nodes()
    green_units = load_green_units()

    if town_lookup.empty:
        notes.append("未偵測到 town_lookup.csv，鄉鎮市區人口密度無法自動帶入")
    if factories.empty:
        notes.append("未偵測到 factory_points.csv，500公尺內工廠數無法自動計算")
    if life_nodes.empty:
        notes.append("未偵測到 schools.csv 或 medical_facilities.csv，敏感受體與生活節點數無法自動計算")
    if green_units.empty:
        notes.append("未偵測到 green_air_sites.csv 或 green_walls.csv，串聯潛力相關綠化單元指標無法自動計算")

    density_values = data.apply(
        lambda r: lookup_population_density(r.get("縣市", ""), r.get("鄉鎮市區", ""), town_lookup),
        axis=1,
    )
    data = fill_if_missing(data, "鄉鎮市區人口密度(人/平方公里)", density_values)

    factory_values = data.apply(
        lambda r: compute_factory_count_500(r.get("TWD97X", np.nan), r.get("TWD97Y", np.nan), factories),
        axis=1,
    )
    data = fill_if_missing(data, "500公尺內工廠數", factory_values)

    life_values = data.apply(
        lambda r: compute_life_node_count_500(r.get("TWD97X", np.nan), r.get("TWD97Y", np.nan), life_nodes),
        axis=1,
    )
    data = fill_if_missing(data, "500公尺內生活節點數", life_values)
    data = fill_if_missing(data, "500公尺內敏感受體數", life_values)

    green_metrics = pd.DataFrame(list(data.apply(lambda r: compute_green_metrics(r, green_units), axis=1)))
    for col in [
        "最近綠化單元距離(公尺)",
        "500公尺內其他綠化單元數",
        "1000公尺內其他綠化單元數",
        "1000公尺內其他綠化單元總面積(公頃)",
    ]:
        data = fill_if_missing(data, col, green_metrics[col])

    data["自動補值提醒"] = "；".join(notes) if notes else ""
    return data, notes


# ============================================================
# 四、節點與串聯評分
# ============================================================

def score_air_pressure(row: pd.Series) -> int:
    factories = row["500公尺內工廠數"]
    road = row["距主要道路距離(公尺)"]

    high = (not pd.isna(factories) and factories >= 3) or (not pd.isna(road) and road <= 100)
    medium = (
        (not pd.isna(factories) and 1 <= factories <= 2)
        or (not pd.isna(road) and 100 < road <= 300)
    )

    if high:
        return 2
    if medium:
        return 1
    return 0


def score_sensitive_receptors(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 3:
        return 2
    if value >= 1:
        return 1
    return 0


def score_population_density(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 1000:
        return 2
    if value >= 500:
        return 1
    return 0


def score_land_ownership(value) -> int:
    text = "" if pd.isna(value) else str(value).strip()

    if text == "":
        return 0
    if any(k in text for k in ["公有", "國有", "市有", "縣有", "鄉有", "鎮有", "區有", "單一權屬"]):
        return 2
    if any(k in text for k in ["公私混合", "混合", "部分公有", "跨機關", "需協調"]):
        return 1
    if any(k in text for k in ["私有", "權屬複雜", "複雜", "徵收"]):
        return 0
    return 1


def score_management_agency(value) -> int:
    text = "" if pd.isna(value) else str(value).strip()

    if text == "":
        return 0
    if any(k in text for k in ["無", "不明", "未知", "未定", "待確認"]):
        return 0
    if any(k in text for k in ["需確認", "可能", "暫由", "協調"]):
        return 1
    return 2


def get_effective_plantable_area(row: pd.Series) -> float:
    area = row["基地面積(公頃)"]
    length = row["基地長度(公里)"]

    if not pd.isna(area):
        return float(area)
    if not pd.isna(length):
        return float(length) * 0.1
    return np.nan


def score_plantable_space(row: pd.Series) -> int:
    area = get_effective_plantable_area(row)

    if pd.isna(area):
        return 0
    if area > 1:
        return 2
    if area > 0.5:
        return 1
    return 0


def score_short_term_feasibility(value) -> int:
    text = "" if pd.isna(value) else str(value).strip()

    if text == "":
        return 0
    if any(k in text for k in ["已納入計畫", "已有預算", "可立即推動", "可短期推動", "高", "是"]):
        return 2
    if any(k in text for k in ["需部分協調", "需協調", "中", "部分", "待評估"]):
        return 1
    if any(k in text for k in ["短期難以推動", "重大協調", "尚無政策支持", "低", "否", "困難"]):
        return 0
    return 1


def score_openness(value) -> int:
    text = "" if pd.isna(value) else str(value).strip()

    if text == "":
        return 0
    if any(k in text for k in ["完全開放", "自由進入", "全日開放", "開放"]):
        if "不開放" not in text and "部分" not in text:
            return 2
    if any(k in text for k in ["部分開放", "特定時段", "特定區域", "需申請", "半開放"]):
        return 1
    if any(k in text for k in ["不開放", "未開放", "封閉", "內部管理"]):
        return 0
    return 1


def score_life_nodes(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 4:
        return 2
    if value >= 2:
        return 1
    return 0


def score_nearest_green(value) -> int:
    if pd.isna(value):
        return 0
    if value <= 100:
        return 2
    if value <= 300:
        return 1
    return 0


def score_green_units_500(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 4:
        return 2
    if value >= 2:
        return 1
    return 0


def score_green_units_1000(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 8:
        return 2
    if value >= 4:
        return 1
    return 0


def score_green_area_1000(value) -> int:
    if pd.isna(value):
        return 0
    if value >= 10:
        return 2
    if value >= 3:
        return 1
    return 0


# ============================================================
# 五、功能角色與建議
# ============================================================

def classify_function_roles(row: pd.Series) -> List[str]:
    roles = []

    q1 = is_yes(row["基地內部是否有停留活動空間"])
    q2 = is_yes(row["是否有步行通學騎行路徑"])
    q3 = is_yes(row["是否有邊界防護需求"])
    q5 = is_yes(row["道路周邊是否有敏感受體或社區活動空間"])
    q6 = is_yes(row["是否有短期事件"])
    q7 = is_yes(row["短期事件是否影響內部空間"])
    q8 = is_yes(row["短期事件是否影響通行路徑"])
    q9 = is_yes(row["短期事件是否位於邊界或鄰近受體"])

    road_distance = row["距主要道路距離(公尺)"]
    near_major_road = not pd.isna(road_distance) and road_distance <= 300

    if q1:
        roles.append("A1")
    if q2:
        roles.append("A2")
    if q3:
        roles.append("A3")
    if near_major_road and q5:
        roles.append("B1")
    if near_major_road and q2:
        roles.append("B2")
    if near_major_road and q3 and q5:
        roles.append("B3")
    if q6 and q7:
        roles.append("C1")
    if q6 and q8:
        roles.append("C2")
    if q6 and q9:
        roles.append("C3")

    return roles


def role_codes_to_text(roles: List[str]) -> str:
    return "、".join(roles) if roles else "未判定"


def roles_to_names(roles: List[str]) -> str:
    if not roles:
        return "未判定"
    return "；".join(ROLE_INFO[r]["情境名稱"] for r in roles)


def roles_to_functions(roles: List[str]) -> str:
    if not roles:
        return "未判定"
    return "；".join(f"{r}：{ROLE_INFO[r]['主要功能']}" for r in roles)


def roles_to_applicability(roles: List[str]) -> str:
    if not roles:
        return "未判定"
    return "；".join(f"{r}：{ROLE_INFO[r]['適用性']}" for r in roles)


def roles_to_management(roles: List[str]) -> str:
    if not roles:
        return "尚無明確功能情境，建議補充現地判讀欄位後再確認。"
    return "；".join(f"{r}：{ROLE_INFO[r]['管理方式']}" for r in roles)


def roles_to_tracking(roles: List[str]) -> str:
    if not roles:
        return "建議補充功能情境資料。"
    return "；".join(f"{r}：{ROLE_INFO[r]['追蹤重點']}" for r in roles)


def first_role_code(roles_text: str) -> str:
    if not isinstance(roles_text, str) or roles_text == "未判定":
        return "未判定"
    return roles_text.split("、")[0]


def build_priority_recommendation(row: pd.Series) -> str:
    node_level = row.get("節點潛力", "基礎潛力")
    link_level = row.get("串聯潛力", "基礎潛力")
    roles_text = row.get("可能功能情境", "未判定")

    if node_level == "高潛力" and link_level == "高潛力":
        main = "節點與串聯條件皆佳，建議優先推動"
    elif node_level == "高潛力":
        main = "節點條件佳，建議優先強化基地功能"
    elif link_level == "高潛力":
        main = "串聯條件佳，建議作為綠網連接補點"
    elif node_level == "中潛力" and link_level == "中潛力":
        main = "節點與串聯條件中等，建議納入第二階段評估"
    elif node_level == "基礎潛力" and link_level == "中潛力":
        main = "串聯條件尚可，建議作為局部連接或補強場址"
    elif node_level == "中潛力" and link_level == "基礎潛力":
        main = "節點條件尚可，建議視管理可行性再評估"
    else:
        main = "節點與串聯條件皆低，建議暫列低優先序"

    if not isinstance(roles_text, str) or roles_text == "未判定":
        note = "功能情境尚未明確，建議補充現地判讀"
    elif any(code in roles_text for code in ["C1", "C2", "C3"]):
        note = "另具事件型情境，應以源頭或作業管理為優先"
    elif any(code in roles_text for code in ["A3", "B3"]):
        note = "另具邊界情境，應檢討連續植栽緩衝"
    elif any(code in roles_text for code in ["A2", "B2"]):
        note = "另具通行暴露情境，應優化通行動線與遮蔭"
    elif any(code in roles_text for code in ["A1", "B1"]):
        note = "另具停留暴露情境，應強化停留空間防護"
    else:
        note = "功能情境尚未明確，建議補充現地判讀"

    return f"{main}；{note}。"


# ============================================================
# 六、主分析流程
# ============================================================

def analyze_green_network(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    data, _ = auto_enrich_context_fields(data)

    for col in NUMERIC_COLUMNS:
        if col not in data.columns:
            data[col] = np.nan
        data[col] = data[col].apply(clean_numeric)

    for col in YES_NO_COLUMNS:
        if col not in data.columns:
            data[col] = "未知"
        data[col] = data[col].apply(normalize_yes_no)

    coords = data.apply(lambda r: twd97_to_wgs84(r["TWD97X"], r["TWD97Y"], r["縣市"]), axis=1)
    data["緯度"] = coords.apply(lambda x: x[0])
    data["經度"] = coords.apply(lambda x: x[1])

    # 節點潛力：9項，每項0-2分，滿分18分
    data["節點_空品壓力分數"] = data.apply(score_air_pressure, axis=1)
    data["節點_敏感受體分數"] = data["500公尺內敏感受體數"].apply(score_sensitive_receptors)
    data["節點_人口暴露分數"] = data["鄉鎮市區人口密度(人/平方公里)"].apply(score_population_density)
    data["節點_土地權屬分數"] = data["土地權屬"].apply(score_land_ownership)
    data["節點_管理權責分數"] = data["管理機關"].apply(score_management_agency)
    data["估算可植栽面積(公頃)"] = data.apply(get_effective_plantable_area, axis=1)
    data["節點_可植栽空間分數"] = data.apply(score_plantable_space, axis=1)
    data["節點_短期推動分數"] = data["短期推動性"].apply(score_short_term_feasibility)
    data["節點_開放可及分數"] = data["開放可及性"].apply(score_openness)
    data["節點_生活節點分數"] = data["500公尺內生活節點數"].apply(score_life_nodes)

    node_cols = [
        "節點_空品壓力分數",
        "節點_敏感受體分數",
        "節點_人口暴露分數",
        "節點_土地權屬分數",
        "節點_管理權責分數",
        "節點_可植栽空間分數",
        "節點_短期推動分數",
        "節點_開放可及分數",
        "節點_生活節點分數",
    ]
    data["節點潛力分數"] = data[node_cols].sum(axis=1)
    data["節點潛力"] = data["節點潛力分數"].apply(lambda s: classify_level(s, high_cut=14, medium_cut=7))

    # 串聯潛力：4項，每項0-2分，滿分8分
    data["串聯_最近綠化距離分數"] = data["最近綠化單元距離(公尺)"].apply(score_nearest_green)
    data["串聯_500公尺綠化單元分數"] = data["500公尺內其他綠化單元數"].apply(score_green_units_500)
    data["串聯_1000公尺綠化單元分數"] = data["1000公尺內其他綠化單元數"].apply(score_green_units_1000)
    data["串聯_1000公尺綠化面積分數"] = data["1000公尺內其他綠化單元總面積(公頃)"].apply(score_green_area_1000)

    link_cols = [
        "串聯_最近綠化距離分數",
        "串聯_500公尺綠化單元分數",
        "串聯_1000公尺綠化單元分數",
        "串聯_1000公尺綠化面積分數",
    ]
    data["串聯潛力分數"] = data[link_cols].sum(axis=1)
    data["串聯潛力"] = data["串聯潛力分數"].apply(lambda s: classify_level(s, high_cut=6, medium_cut=3))

    role_lists = data.apply(classify_function_roles, axis=1)
    data["可能功能情境"] = role_lists.apply(role_codes_to_text)
    data["功能情境名稱"] = role_lists.apply(roles_to_names)
    data["主要對應功能"] = role_lists.apply(roles_to_functions)
    data["植生措施適用性"] = role_lists.apply(roles_to_applicability)
    data["建議後續管理方式"] = role_lists.apply(roles_to_management)
    data["追蹤重點"] = role_lists.apply(roles_to_tracking)
    data["地圖角色代碼"] = data["可能功能情境"].apply(first_role_code)
    data["優先推動建議"] = data.apply(build_priority_recommendation, axis=1)

    return data


# ============================================================
# 七、讀檔與匯出
# ============================================================

def read_csv_with_fallback(uploaded_file) -> pd.DataFrame:
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5", "latin1"]
    last_error = None
    for encoding in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=encoding)
        except Exception as e:
            last_error = e
    raise ValueError(f"CSV讀取失敗，已嘗試編碼：{', '.join(encodings)}。最後錯誤：{last_error}")


def make_csv_download(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def make_excel_download(df: pd.DataFrame):
    if find_spec("openpyxl") is None:
        return None

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="分析結果")

        pd.DataFrame(
            {
                "資料項目": [
                    "人口密度",
                    "工廠資料",
                    "空品淨化區",
                    "綠牆",
                    "各級學校",
                    "醫療院所",
                ],
                "資料來源": [
                    "114年人口密度資料，1150427下載",
                    "環境保護許可管理系統(暨解除列管)對象基本資料-曾有空列管，1150427下載",
                    "空品淨化區，114年第四季季報",
                    "綠牆，113年第四季季報，面積由平方公尺換算為公頃",
                    "111學年度各級學校名錄",
                    "112年12月醫療院所分布圖",
                ],
            }
        ).to_excel(writer, index=False, sheet_name="資料來源說明")

        pd.DataFrame([{"功能情境": code, **info} for code, info in ROLE_INFO.items()]).to_excel(
            writer, index=False, sheet_name="功能角色規則"
        )

    return output.getvalue()


# ============================================================
# 八、地圖與結果顯示
# ============================================================

def build_map(df: pd.DataFrame, color_by: str):
    if pdk is None:
        st.warning("尚未安裝 pydeck，無法顯示互動地圖。請在 requirements.txt 加入 pydeck。")
        return

    map_df = df.dropna(subset=["緯度", "經度"]).copy()
    map_df = map_df[(map_df["緯度"].between(21, 27)) & (map_df["經度"].between(118, 123))]

    if map_df.empty:
        st.warning("沒有有效座標可顯示。請確認 TWD97X、TWD97Y 欄位是否正確。")
        return

    if color_by == "功能情境":
        map_df["color"] = map_df["地圖角色代碼"].apply(lambda x: ROLE_COLORS.get(x, ROLE_COLORS["未判定"]))
    elif color_by == "串聯潛力":
        map_df["color"] = map_df["串聯潛力"].apply(lambda x: LINK_COLORS.get(x, [127, 140, 141, 160]))
    else:
        map_df["color"] = map_df["節點潛力"].apply(lambda x: NODE_COLORS.get(x, [127, 140, 141, 160]))

    map_df["radius"] = np.sqrt(map_df["估算可植栽面積(公頃)"].fillna(0.2).clip(lower=0.2, upper=60)) * 120

    view_state = pdk.ViewState(
        longitude=float(map_df["經度"].mean()),
        latitude=float(map_df["緯度"].mean()),
        zoom=7,
        pitch=0,
    )

    layer = pdk.Layer(
        "ScatterplotLayer",
        data=map_df,
        get_position="[經度, 緯度]",
        get_radius="radius",
        get_fill_color="color",
        pickable=True,
        auto_highlight=True,
    )

    tooltip = {
        "html": """
        <b>{基地名稱}</b><br/>
        縣市：{縣市}<br/>
        鄉鎮市區：{鄉鎮市區}<br/>
        節點潛力：{節點潛力}（{節點潛力分數}/18）<br/>
        串聯潛力：{串聯潛力}（{串聯潛力分數}/8）<br/>
        可能功能情境：{可能功能情境}<br/>
        建議：{優先推動建議}
        """,
        "style": {"backgroundColor": "white", "color": "black"},
    }

    st.pydeck_chart(
        pdk.Deck(
            map_style=None,
            initial_view_state=view_state,
            layers=[layer],
            tooltip=tooltip,
        ),
        use_container_width=True,
    )



# ============================================================
# 八之一、表單下載、單點摘要與周邊對象
# ============================================================

def make_blank_template_csv() -> bytes:
    """產出批次上傳用空白制式表單 CSV。"""
    return pd.DataFrame(columns=REQUIRED_COLUMNS).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def make_blank_template_excel():
    """產出含欄位說明的空白制式 Excel 表單。"""
    if find_spec("openpyxl") is None:
        return None

    template_df = pd.DataFrame(columns=REQUIRED_COLUMNS)

    field_notes = pd.DataFrame(
        [
            {"欄位名稱": "代碼", "填寫說明": "場址唯一識別碼，可自行編號。", "建議填寫值或選項": "例如：SITE-001"},
            {"欄位名稱": "縣市", "填寫說明": "場址所在縣市，需與 town_lookup.csv 中縣市名稱一致，才能自動帶入人口密度。", "建議填寫值或選項": "例如：臺北市、新北市、桃園市"},
            {"欄位名稱": "鄉鎮市區", "填寫說明": "場址所在鄉鎮市區，需與 town_lookup.csv 中鄉鎮市區名稱一致。", "建議填寫值或選項": "例如：中正區、板橋區、草屯鎮"},
            {"欄位名稱": "基地名稱", "填寫說明": "場址名稱。若該基地已存在於綠化單元資料，會用於排除自己。", "建議填寫值或選項": "文字"},
            {"欄位名稱": "TWD97X", "填寫說明": "TWD97 X 座標，單點模式要求 6 碼數字。", "建議填寫值或選項": "例如：250000"},
            {"欄位名稱": "TWD97Y", "填寫說明": "TWD97 Y 座標，單點模式要求 7 碼數字。", "建議填寫值或選項": "例如：2650000"},
            {"欄位名稱": "距主要道路距離(公尺)", "填寫說明": "目前保留人工填寫，用於空品壓力與道路線源情境判斷。", "建議填寫值或選項": "數值；若未知可先填 999"},
            {"欄位名稱": "土地權屬", "填寫說明": "用於判斷土地權屬單純度。", "建議填寫值或選項": "公有或單一權屬／公私混合或需協調／私有或權屬複雜／其他／待確認"},
            {"欄位名稱": "管理機關", "填寫說明": "用於判斷管理權責明確性。", "建議填寫值或選項": "管理機關明確／需確認或協調／無明確管理單位"},
            {"欄位名稱": "基地面積(公頃)", "填寫說明": "用於可植栽空間評分；若無面積但有長度，系統會用長度估算。", "建議填寫值或選項": "數值，例如 0.8"},
            {"欄位名稱": "基地長度(公里)", "填寫說明": "基地面積缺漏時，以長度 × 0.1 估算面積。", "建議填寫值或選項": "數值，例如 1.2"},
            {"欄位名稱": "短期推動性", "填寫說明": "用於判斷近期是否容易推動。", "建議填寫值或選項": "可短期推動／需部分協調／短期難以推動／待評估"},
            {"欄位名稱": "開放可及性", "填寫說明": "用於判斷公共服務性。", "建議填寫值或選項": "完全開放／部分開放／不開放／待確認"},
            {"欄位名稱": "基地內部是否有停留活動空間", "填寫說明": "是否有人會在基地內休憩、運動、活動或等候。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有步行通學騎行路徑", "填寫說明": "基地內或周邊是否有人會步行、通學、騎車或穿越。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有邊界防護需求", "填寫說明": "基地是否位於道路、工廠、空地、住宅或學校等交界，需要緩衝防護。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "道路周邊是否有敏感受體或社區活動空間", "填寫說明": "道路旁是否有住宅、學校、醫療院所或社區活動空間。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有短期事件", "填寫說明": "基地或周邊是否有施工、整地、土方、堆置、臨時活動或短期作業。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否影響內部空間", "填寫說明": "短期事件是否影響人們停留或活動的空間。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否影響通行路徑", "填寫說明": "短期事件是否影響步行、通學、騎車或通行路線。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否位於邊界或鄰近受體", "填寫說明": "短期事件是否靠近基地邊界、周界、住宅、學校或醫療院所。", "建議填寫值或選項": "是／否／未知"},
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template_df.to_excel(writer, index=False, sheet_name="批次上傳空白表單")
        field_notes.to_excel(writer, index=False, sheet_name="欄位填寫說明")
    return output.getvalue()


def object_name(row: pd.Series, candidates: List[str], default: str = "未命名") -> str:
    for col in candidates:
        if col in row.index and not pd.isna(row[col]) and str(row[col]).strip():
            return str(row[col]).strip()
    return default


def prepare_point_layer_df(df: pd.DataFrame, county: str = "") -> pd.DataFrame:
    """將 TWD97 點位資料轉成 pydeck 可用的經緯度資料。"""
    if df.empty:
        return df.copy()

    data = df.copy()
    coords = data.apply(lambda r: twd97_to_wgs84(r["TWD97X"], r["TWD97Y"], county), axis=1)
    data["緯度"] = coords.apply(lambda x: x[0])
    data["經度"] = coords.apply(lambda x: x[1])
    data = data.dropna(subset=["緯度", "經度"])
    data = data[(data["緯度"].between(21, 27)) & (data["經度"].between(118, 123))]
    return data


def nearby_factories(row: pd.Series, radius: float = 500) -> pd.DataFrame:
    factories = load_factory_points()
    if factories.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = factories.copy()
    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)
    # 距離 1 公尺以內視為同一基地，不列入周邊綠化單元清單。
    data = data[(data["距離(公尺)"] > 1) & (data["距離(公尺)"] <= radius)].copy()
    if data.empty:
        return data

    # 工廠地圖與明細顯示名稱：優先直接使用 factory_points.csv 的「工廠名稱」。
    if "工廠名稱" in data.columns:
        direct_name = data["工廠名稱"].fillna("").astype(str).str.strip()
        direct_name = direct_name.replace("", np.nan)
    else:
        direct_name = pd.Series(np.nan, index=data.index)

    fallback_name = data.apply(
        lambda r: object_name(
            r,
            [
                "工廠名稱",
                "公私場所名稱",
                "事業機構名稱",
                "公司名稱",
                "工廠廠名",
                "機構名稱",
                "事業名稱",
                "管制編號",
                "名稱",
            ],
            "未命名工廠",
        ),
        axis=1,
    )

    data["名稱"] = direct_name.fillna(fallback_name).fillna("未命名工廠")
    data["tooltip_name"] = data["名稱"]
    return data.sort_values("距離(公尺)")


def nearby_life_nodes(row: pd.Series, radius: float = 500) -> pd.DataFrame:
    nodes = load_life_nodes()
    if nodes.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = nodes.copy()
    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)
    # 距離 1 公尺以內視為同一基地，不列入周邊綠化單元清單。
    data = data[(data["距離(公尺)"] > 1) & (data["距離(公尺)"] <= radius)].copy()
    if data.empty:
        return data

    data["名稱"] = data.apply(
        lambda r: object_name(r, ["節點名稱", "學校名稱", "醫療院所名稱", "機構名稱", "名稱"], "生活節點"),
        axis=1,
    )
    if "節點類型" not in data.columns:
        data["節點類型"] = data.get("資料來源", "生活節點")
    return data.sort_values("距離(公尺)")


def nearby_green_units(row: pd.Series, radius: float = 1000) -> pd.DataFrame:
    green = load_green_units()
    if green.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = green.copy()

    site_code = str(row.get("代碼", "")).strip()
    site_name = str(row.get("基地名稱", "")).strip()
    if site_code:
        data = data[data["綠化單元代碼"].astype(str).str.strip() != site_code]
    if site_name:
        data = data[data["綠化單元名稱"].astype(str).str.strip() != site_name]

    if data.empty:
        return data

    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)
    # 距離 1 公尺以內視為同一基地，不列入周邊綠化單元清單。
    data = data[(data["距離(公尺)"] > 1) & (data["距離(公尺)"] <= radius)].copy()
    if data.empty:
        return data

    data["名稱"] = data.apply(
        lambda r: object_name(r, ["綠化單元名稱", "名稱"], "綠化單元"),
        axis=1,
    )
    return data.sort_values("距離(公尺)")


def build_single_score_table(row: pd.Series) -> pd.DataFrame:
    """單點模式用：列出各項評分標準、實際數值與得分。"""
    records = [
        {
            "類別": "節點潛力",
            "評分項目": "空品壓力",
            "評估值": f"500m工廠數={row['500公尺內工廠數']}；距主要道路={row['距主要道路距離(公尺)']}m",
            "得分": row["節點_空品壓力分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "敏感受體密度",
            "評估值": f"500m內敏感受體數={row['500公尺內敏感受體數']}",
            "得分": row["節點_敏感受體分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "鄉鎮人口暴露",
            "評估值": f"{row.get('縣市', '')}{row.get('鄉鎮市區', '')}；人口密度={row['鄉鎮市區人口密度(人/平方公里)']} 人/km²",
            "得分": row["節點_人口暴露分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "土地權屬單純度",
            "評估值": row["土地權屬"],
            "得分": row["節點_土地權屬分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "管理權責明確性",
            "評估值": row["管理機關"],
            "得分": row["節點_管理權責分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "可植栽空間",
            "評估值": f"估算可植栽面積={row['估算可植栽面積(公頃)']} ha",
            "得分": row["節點_可植栽空間分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "短期推動性",
            "評估值": row["短期推動性"],
            "得分": row["節點_短期推動分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "開放可及性",
            "評估值": row["開放可及性"],
            "得分": row["節點_開放可及分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "鄰近生活節點",
            "評估值": f"500m內生活節點數={row['500公尺內生活節點數']}",
            "得分": row["節點_生活節點分數"],
        },
        {
            "類別": "節點潛力",
            "評分項目": "節點潛力總分",
            "評估值": row["節點潛力"],
            "得分": f"{row['節點潛力分數']} / 18",
        },
        {
            "類別": "串聯潛力",
            "評分項目": "最近綠化單元距離",
            "評估值": f"{row['最近綠化單元距離(公尺)']} m",
            "得分": row["串聯_最近綠化距離分數"],
        },
        {
            "類別": "串聯潛力",
            "評分項目": "500m內其他綠化單元數",
            "評估值": row["500公尺內其他綠化單元數"],
            "得分": row["串聯_500公尺綠化單元分數"],
        },
        {
            "類別": "串聯潛力",
            "評分項目": "1000m內其他綠化單元數",
            "評估值": row["1000公尺內其他綠化單元數"],
            "得分": row["串聯_1000公尺綠化單元分數"],
        },
        {
            "類別": "串聯潛力",
            "評分項目": "1000m內其他綠化單元總面積",
            "評估值": f"{row['1000公尺內其他綠化單元總面積(公頃)']} ha",
            "得分": row["串聯_1000公尺綠化面積分數"],
        },
        {
            "類別": "串聯潛力",
            "評分項目": "串聯潛力總分",
            "評估值": row["串聯潛力"],
            "得分": f"{row['串聯潛力分數']} / 8",
        },
    ]
    standards = {
        "空品壓力": "2分：工廠≥3或道路≤100m；1分：工廠1–2或道路101–300m；0分：其他",
        "敏感受體密度": "2分：≥3處；1分：1–2處；0分：0處",
        "鄉鎮人口暴露": "2分：≥1000人/km²；1分：500–999人/km²；0分：<500人/km²",
        "土地權屬單純度": "2分：公有或單一權屬；1分：公私混合或需協調；0分：私有或權屬複雜",
        "管理權責明確性": "2分：管理機關明確；1分：需確認或協調；0分：無明確管理單位",
        "可植栽空間": "2分：>1ha；1分：>0.5–1ha；0分：≤0.5ha",
        "短期推動性": "2分：可短期推動；1分：需部分協調；0分：短期難以推動",
        "開放可及性": "2分：完全開放；1分：部分開放；0分：不開放",
        "鄰近生活節點": "2分：≥4處；1分：2–3處；0分：0–1處",
        "節點潛力總分": "高潛力：14–18；中潛力：7–13；基礎潛力：0–6",
        "最近綠化單元距離": "2分：≤100m；1分：>100–300m；0分：>300m",
        "500m內其他綠化單元數": "2分：≥4處；1分：2–3處；0分：0–1處",
        "1000m內其他綠化單元數": "2分：≥8處；1分：4–7處；0分：0–3處",
        "1000m內其他綠化單元總面積": "2分：≥10ha；1分：3–9.99ha；0分：<3ha",
        "串聯潛力總分": "高潛力：6–8；中潛力：3–5；基礎潛力：0–2",
    }
    out = pd.DataFrame(records)
    out["計分標準"] = out["評分項目"].map(standards).fillna("")
    return out


def render_nearby_object_tables(row: pd.Series):
    factories = nearby_factories(row, radius=500)
    life = nearby_life_nodes(row, radius=500)
    green = nearby_green_units(row, radius=1000)

    st.markdown("#### 周邊對象明細")
    st.caption("敏感受體與生活節點目前皆依各級學校與醫療院所資料計算；綠化單元統計範圍為 1000 公尺，500 公尺數量另納入串聯潛力評分。")

    tab1, tab2, tab3, tab4 = st.tabs(["工廠", "敏感受體", "生活節點", "綠化單元"])

    with tab1:
        if factories.empty:
            st.info("500 公尺內未偵測到工廠，或尚未提供 factory_points.csv。")
        else:
            display_cols = [c for c in ["名稱", "工廠名稱", "emsno", "industryname", "facilityaddress", "距離(公尺)"] if c in factories.columns]
            # 若名稱與工廠名稱重複，只保留名稱，避免欄位太雜。
            if "名稱" in display_cols and "工廠名稱" in display_cols:
                display_cols.remove("工廠名稱")
            st.dataframe(
                factories[display_cols].round({"距離(公尺)": 1}),
                hide_index=True,
                use_container_width=True,
            )

    with tab2:
        if life.empty:
            st.info("500 公尺內未偵測到敏感受體，或尚未提供 schools.csv / medical_facilities.csv。")
        else:
            cols = [c for c in ["名稱", "節點類型", "資料來源", "距離(公尺)"] if c in life.columns]
            st.dataframe(life[cols].round({"距離(公尺)": 1}), hide_index=True, use_container_width=True)

    with tab3:
        if life.empty:
            st.info("500 公尺內未偵測到生活節點，或尚未提供 schools.csv / medical_facilities.csv。")
        else:
            cols = [c for c in ["名稱", "節點類型", "資料來源", "距離(公尺)"] if c in life.columns]
            st.dataframe(life[cols].round({"距離(公尺)": 1}), hide_index=True, use_container_width=True)

    with tab4:
        if green.empty:
            st.info("1000 公尺內未偵測到綠化單元，或尚未提供 green_air_sites.csv / green_walls.csv。")
        else:
            cols = [c for c in ["名稱", "資料類型", "基地面積(公頃)", "距離(公尺)"] if c in green.columns]
            st.dataframe(green[cols].round({"基地面積(公頃)": 3, "距離(公尺)": 1}), hide_index=True, use_container_width=True)


def make_circle_polygon_twd97(x: float, y: float, radius_m: float = 500, n_points: int = 160, county: str = "") -> pd.DataFrame:
    """建立 TWD97 圓形範圍並轉為 WGS84 polygon 座標。"""
    angles = np.linspace(0, 2 * np.pi, n_points, endpoint=False)
    coords = []
    for angle in angles:
        px = x + radius_m * np.cos(angle)
        py = y + radius_m * np.sin(angle)
        lat, lon = twd97_to_wgs84(px, py, county)
        coords.append([lon, lat])
    coords.append(coords[0])
    return pd.DataFrame([{"範圍": f"{int(radius_m)}公尺範圍", "coordinates": coords}])


def all_factories_with_distance(row: pd.Series) -> pd.DataFrame:
    factories = load_factory_points()
    if factories.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = factories.copy()
    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)
    # 工廠地圖與明細顯示名稱：優先直接使用 factory_points.csv 的「工廠名稱」。
    if "工廠名稱" in data.columns:
        direct_name = data["工廠名稱"].fillna("").astype(str).str.strip()
        direct_name = direct_name.replace("", np.nan)
    else:
        direct_name = pd.Series(np.nan, index=data.index)

    fallback_name = data.apply(
        lambda r: object_name(
            r,
            [
                "工廠名稱",
                "公私場所名稱",
                "事業機構名稱",
                "公司名稱",
                "工廠廠名",
                "機構名稱",
                "事業名稱",
                "管制編號",
                "名稱",
            ],
            "未命名工廠",
        ),
        axis=1,
    )

    data["名稱"] = direct_name.fillna(fallback_name).fillna("未命名工廠")
    data["tooltip_name"] = data["名稱"]
    return data.sort_values("距離(公尺)")


def all_life_nodes_with_distance(row: pd.Series) -> pd.DataFrame:
    nodes = load_life_nodes()
    if nodes.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = nodes.copy()
    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)
    data["名稱"] = data.apply(
        lambda r: object_name(r, ["節點名稱", "學校名稱", "醫療院所名稱", "機構名稱", "名稱"], "生活節點"),
        axis=1,
    )
    if "節點類型" not in data.columns:
        data["節點類型"] = data.get("資料來源", "生活節點")
    return data.sort_values("距離(公尺)")


def all_green_units_with_distance(row: pd.Series) -> pd.DataFrame:
    green = load_green_units()
    if green.empty or pd.isna(row["TWD97X"]) or pd.isna(row["TWD97Y"]):
        return pd.DataFrame()

    data = green.copy()

    site_code = str(row.get("代碼", "")).strip()
    site_name = str(row.get("基地名稱", "")).strip()
    if site_code:
        data = data[data["綠化單元代碼"].astype(str).str.strip() != site_code]
    if site_name:
        data = data[data["綠化單元名稱"].astype(str).str.strip() != site_name]

    if data.empty:
        return data

    data["距離(公尺)"] = euclidean_distances_m(float(row["TWD97X"]), float(row["TWD97Y"]), data)

    # 距離 1 公尺以內視為同一基地，不列入完整綠化單元圖層與統計。
    data = data[data["距離(公尺)"] > 1].copy()

    if data.empty:
        return data

    data["名稱"] = data.apply(
        lambda r: object_name(r, ["綠化單元名稱", "名稱"], "綠化單元"),
        axis=1,
    )
    return data.sort_values("距離(公尺)")


def render_single_context_map(row: pd.Series):
    if pdk is None:
        st.warning("尚未安裝 pydeck，無法顯示互動地圖。")
        return

    if pd.isna(row["緯度"]) or pd.isna(row["經度"]):
        st.warning("沒有有效座標可顯示。請確認 TWD97X、TWD97Y。")
        return

    county = row.get("縣市", "")

    base_df = pd.DataFrame([{
        "名稱": row.get("基地名稱", "輸入基地"),
        "tooltip_name": row.get("基地名稱", "輸入基地"),
        "類型": "查詢基地",
        "距離(公尺)": 0,
        "緯度": row["緯度"],
        "經度": row["經度"],
        "color": [0, 0, 0, 245],
        "radius": 180,
    }])

    factories_all = all_factories_with_distance(row)
    life_all = all_life_nodes_with_distance(row)
    green_all = all_green_units_with_distance(row)

    factories_map = prepare_point_layer_df(factories_all, county)
    life_map = prepare_point_layer_df(life_all, county)
    green_map = prepare_point_layer_df(green_all, county)

    if not factories_map.empty:
        # 不共用泛用「名稱」欄位，改用地圖專用 tooltip_name，避免被預設值覆蓋。
        if "工廠名稱" in factories_map.columns:
            factories_map["tooltip_name"] = factories_map["工廠名稱"].fillna("").astype(str).str.strip()
        elif "名稱" in factories_map.columns:
            factories_map["tooltip_name"] = factories_map["名稱"].fillna("").astype(str).str.strip()
        else:
            factories_map["tooltip_name"] = ""

        factories_map["tooltip_name"] = factories_map["tooltip_name"].replace("", "未命名工廠")
        factories_map["名稱"] = factories_map["tooltip_name"]
        factories_map["類型"] = "工廠"
        factories_map["color"] = [[231, 76, 60, 130]] * len(factories_map)
        factories_map["radius"] = 60

    if not life_map.empty:
        life_map["tooltip_name"] = life_map["名稱"] if "名稱" in life_map.columns else "生活節點"
        life_map["類型"] = life_map.apply(
            lambda r: f"{r.get('資料來源', '生活節點')}／{r.get('節點類型', '生活節點')}",
            axis=1,
        )
        life_map["color"] = [[52, 152, 219, 140]] * len(life_map)
        life_map["radius"] = 60

    if not green_map.empty:
        green_map["tooltip_name"] = green_map["名稱"] if "名稱" in green_map.columns else "綠化單元"
        green_map["類型"] = green_map.apply(
            lambda r: f"綠化單元／{r.get('資料類型', '')}",
            axis=1,
        )
        green_map["color"] = [[46, 204, 113, 130]] * len(green_map)
        green_map["radius"] = 65

    layers = []

    # 1000 公尺範圍圈：先畫，避免蓋住 500 公尺圈
    circle_1000 = make_circle_polygon_twd97(
        float(row["TWD97X"]),
        float(row["TWD97Y"]),
        radius_m=1000,
        county=county,
    )
    layers.append(
        pdk.Layer(
            "PolygonLayer",
            data=circle_1000,
            get_polygon="coordinates",
            get_fill_color=[255, 170, 0, 18],
            get_line_color=[255, 170, 0, 190],
            line_width_min_pixels=2,
            pickable=False,
        )
    )

    # 500 公尺範圍圈
    circle_500 = make_circle_polygon_twd97(
        float(row["TWD97X"]),
        float(row["TWD97Y"]),
        radius_m=500,
        county=county,
    )
    layers.append(
        pdk.Layer(
            "PolygonLayer",
            data=circle_500,
            get_polygon="coordinates",
            get_fill_color=[0, 120, 255, 30],
            get_line_color=[0, 120, 255, 220],
            line_width_min_pixels=2,
            pickable=False,
        )
    )

    # 完整資料圖層：綠化單元 → 生活節點／敏感受體 → 工廠 → 查詢基地
    if not green_map.empty:
        layers.append(pdk.Layer(
            "ScatterplotLayer",
            data=green_map,
            get_position="[經度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            auto_highlight=True,
        ))

    if not life_map.empty:
        layers.append(pdk.Layer(
            "ScatterplotLayer",
            data=life_map,
            get_position="[經度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            auto_highlight=True,
        ))

    if not factories_map.empty:
        layers.append(pdk.Layer(
            "ScatterplotLayer",
            data=factories_map,
            get_position="[經度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            auto_highlight=True,
        ))

    layers.append(
        pdk.Layer(
            "ScatterplotLayer",
            data=base_df,
            get_position="[經度, 緯度]",
            get_radius="radius",
            get_fill_color="color",
            pickable=True,
            auto_highlight=True,
        )
    )

    view_state = pdk.ViewState(
        longitude=float(row["經度"]),
        latitude=float(row["緯度"]),
        zoom=12,
        pitch=0,
    )

    tooltip = {
        "html": """
        <b>{tooltip_name}</b><br/>
        類型：{類型}<br/>
        距離：{距離(公尺)} 公尺
        """,
        "style": {"backgroundColor": "white", "color": "black"},
    }

    st.pydeck_chart(
        pdk.Deck(
            map_style=None,
            initial_view_state=view_state,
            layers=layers,
            tooltip=tooltip,
        ),
        use_container_width=True,
    )

    st.caption(
        "圖例：黑色＝查詢基地；藍色透明圈＝500 公尺範圍；橘色透明圈＝1000 公尺範圍；"
        "紅色＝完整工廠資料；藍色＝完整敏感受體／生活節點資料；綠色＝完整綠化單元資料。"
    )

    def count_within(df: pd.DataFrame, distance: float) -> int:
        if df.empty or "距離(公尺)" not in df.columns:
            return 0
        return int((df["距離(公尺)"] <= distance).sum())

    summary = pd.DataFrame(
        [
            {
                "資料類別": "工廠",
                "500公尺內數量": count_within(factories_all, 500),
                "1000公尺內數量": count_within(factories_all, 1000),
                "完整資料筆數": len(factories_all),
            },
            {
                "資料類別": "敏感受體／生活節點",
                "500公尺內數量": count_within(life_all, 500),
                "1000公尺內數量": count_within(life_all, 1000),
                "完整資料筆數": len(life_all),
            },
            {
                "資料類別": "綠化單元",
                "500公尺內數量": count_within(green_all, 500),
                "1000公尺內數量": count_within(green_all, 1000),
                "完整資料筆數": len(green_all),
            },
        ]
    )

    st.markdown("#### 地圖點位統計")
    st.dataframe(summary, hide_index=True, use_container_width=True)


# ============================================================
# 八之二、結果顯示
# ============================================================


def classify_push_order(recommendation: str) -> str:
    """將優先推動建議簡化為批次統計用的推動順序分類。"""
    if not isinstance(recommendation, str):
        return "未分類"
    if "建議優先推動" in recommendation:
        return "建議優先推動"
    if "建議優先強化基地功能" in recommendation:
        return "建議優先強化基地功能"
    if "建議作為綠網連接補點" in recommendation:
        return "建議作為綠網連接補點"
    if "建議納入第二階段評估" in recommendation:
        return "建議納入第二階段評估"
    if "建議作為局部連接或補強場址" in recommendation:
        return "建議作為局部連接或補強場址"
    if "建議視管理可行性再評估" in recommendation:
        return "建議視管理可行性再評估"
    if "建議暫列低優先序" in recommendation:
        return "建議暫列低優先序"
    return "其他"

def render_batch_summary(result_df: pd.DataFrame):
    st.subheader("二、分析摘要")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("總筆數", f"{len(result_df):,}")
    col2.metric("高節點潛力", f"{(result_df['節點潛力'] == '高潛力').sum():,}")
    col3.metric("高串聯潛力", f"{(result_df['串聯潛力'] == '高潛力').sum():,}")
    col4.metric("縣市數", f"{result_df['縣市'].nunique():,}")

    chart_col1, chart_col2, chart_col3 = st.columns(3)
    with chart_col1:
        st.write("節點潛力分布")
        st.bar_chart(result_df["節點潛力"].value_counts().reindex(["高潛力", "中潛力", "基礎潛力"]).fillna(0))
        st.caption("節點潛力依空品壓力、敏感受體、人口暴露、管理條件與公共服務性等 9 項指標加總判定。")

    with chart_col2:
        st.write("串聯潛力分布")
        st.bar_chart(result_df["串聯潛力"].value_counts().reindex(["高潛力", "中潛力", "基礎潛力"]).fillna(0))
        st.caption("串聯潛力依最近綠化距離、周邊綠化單元數及周邊綠化總面積判定。")

    with chart_col3:
        st.write("推動順序統計")
        push_order = result_df["優先推動建議"].apply(classify_push_order)
        order = [
            "建議優先推動",
            "建議優先強化基地功能",
            "建議作為綠網連接補點",
            "建議納入第二階段評估",
            "建議作為局部連接或補強場址",
            "建議視管理可行性再評估",
            "建議暫列低優先序",
            "其他",
        ]
        st.bar_chart(push_order.value_counts().reindex(order).dropna())
        st.caption("依優先推動建議整理為推動順序分類，方便批次場址進行排序與分群。")


def render_single_summary(result_df: pd.DataFrame):
    st.subheader("二、分析摘要")
    row = result_df.iloc[0]

    headline = pd.DataFrame([
        {"項目": "節點潛力", "結果": f"{row['節點潛力']}（{row['節點潛力分數']} / 18）"},
        {"項目": "串聯潛力", "結果": f"{row['串聯潛力']}（{row['串聯潛力分數']} / 8）"},
        {"項目": "可能功能情境", "結果": row["可能功能情境"]},
        {"項目": "優先推動建議", "結果": row["優先推動建議"]},
    ])
    st.dataframe(headline, hide_index=True, use_container_width=True)

    st.markdown("#### 評分明細")
    score_table = build_single_score_table(row)

    def highlight_total_rows(r):
        if "總分" in str(r["評分項目"]):
            return ["background-color: #E8F6EF; font-weight: bold"] * len(r)
        return [""] * len(r)

    st.dataframe(
        score_table.style.apply(highlight_total_rows, axis=1),
        hide_index=True,
        use_container_width=True,
    )

    render_nearby_object_tables(row)


def render_result_table_and_filters(result_df: pd.DataFrame) -> pd.DataFrame:
    st.subheader("三、分析結果表")

    f1, f2, f3 = st.columns(3)
    with f1:
        selected_cities = st.multiselect(
            "篩選縣市",
            options=sorted(result_df["縣市"].dropna().unique()),
            default=sorted(result_df["縣市"].dropna().unique()),
        )
    with f2:
        selected_node_levels = st.multiselect(
            "篩選節點潛力",
            options=["高潛力", "中潛力", "基礎潛力"],
            default=["高潛力", "中潛力", "基礎潛力"],
        )
    with f3:
        selected_link_levels = st.multiselect(
            "篩選串聯潛力",
            options=["高潛力", "中潛力", "基礎潛力"],
            default=["高潛力", "中潛力", "基礎潛力"],
        )

    filtered_df = result_df[
        result_df["縣市"].isin(selected_cities)
        & result_df["節點潛力"].isin(selected_node_levels)
        & result_df["串聯潛力"].isin(selected_link_levels)
    ].copy()

    main_columns = [
        "代碼",
        "縣市",
        "鄉鎮市區",
        "基地名稱",
        "估算可植栽面積(公頃)",
        "鄉鎮市區人口密度(人/平方公里)",
        "500公尺內工廠數",
        "500公尺內敏感受體數",
        "500公尺內生活節點數",
        "最近綠化單元距離(公尺)",
        "500公尺內其他綠化單元數",
        "1000公尺內其他綠化單元數",
        "1000公尺內其他綠化單元總面積(公頃)",
        "節點潛力分數",
        "節點潛力",
        "串聯潛力分數",
        "串聯潛力",
        "可能功能情境",
        "主要對應功能",
        "植生措施適用性",
        "優先推動建議",
    ]

    st.dataframe(filtered_df[main_columns], use_container_width=True, hide_index=True)
    return filtered_df


def render_detail_sections(result_df: pd.DataFrame):
    with st.expander("查看自動補值後資料"):
        auto_cols = [
            "代碼",
            "基地名稱",
            "鄉鎮市區人口密度(人/平方公里)",
            "500公尺內工廠數",
            "500公尺內敏感受體數",
            "500公尺內生活節點數",
            "最近綠化單元距離(公尺)",
            "500公尺內其他綠化單元數",
            "1000公尺內其他綠化單元數",
            "1000公尺內其他綠化單元總面積(公頃)",
            "自動補值提醒",
        ]
        st.dataframe(result_df[auto_cols], use_container_width=True, hide_index=True)

    with st.expander("查看功能角色判釋規則"):
        st.dataframe(pd.DataFrame([{"功能情境": code, **info} for code, info in ROLE_INFO.items()]),
                     use_container_width=True, hide_index=True)


def render_download_section(result_df: pd.DataFrame):
    st.subheader("匯出結果")
    excel_bytes = make_excel_download(result_df)
    if excel_bytes is not None:
        st.download_button(
            label="匯出完整分析結果 Excel",
            data=excel_bytes,
            file_name="城市空品植生綠網淨化單元分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.warning("目前環境沒有 openpyxl，暫時提供 CSV 匯出。若要匯出 Excel，請在 requirements.txt 加入 openpyxl。")

    st.download_button(
        label="匯出完整分析結果 CSV",
        data=make_csv_download(result_df),
        file_name="城市空品植生綠網淨化單元分析結果.csv",
        mime="text/csv",
    )


def render_full_results(result_df: pd.DataFrame, mode: str):
    notes = []
    if "自動補值提醒" in result_df.columns:
        notes = sorted(set([n for n in result_df["自動補值提醒"].dropna().astype(str).unique() if n.strip()]))
    if notes:
        st.warning("；".join(notes))

    if mode == "batch":
        render_batch_summary(result_df)
        render_result_table_and_filters(result_df)
        render_download_section(result_df)
        render_detail_sections(result_df)
    else:
        render_single_summary(result_df)
        st.subheader("三、地圖展示")
        render_single_context_map(result_df.iloc[0])
        render_download_section(result_df)
        render_detail_sections(result_df)


# ============================================================
# 九、單點輸入
# ============================================================

def build_single_site_dataframe() -> pd.DataFrame | None:
    town_lookup = load_town_lookup()
    has_town_lookup = not town_lookup.empty

    st.subheader("單點基地資料輸入")

    # 縣市與鄉鎮市區放在 form 外，才能在選擇縣市後即時更新鄉鎮市區清單。
    st.markdown("#### 1. 基本資料")
    location_col1, location_col2 = st.columns(2)

    with location_col1:
        if has_town_lookup:
            town_lookup_display = town_lookup.copy()
            town_lookup_display["縣市"] = town_lookup_display["縣市"].astype(str).str.strip()
            town_lookup_display["鄉鎮市區"] = town_lookup_display["鄉鎮市區"].astype(str).str.strip()

            county_options = sorted(town_lookup_display["縣市"].dropna().unique())
            county = st.selectbox("縣市", county_options, key="single_county")

            town_options = sorted(
                town_lookup_display.loc[
                    town_lookup_display["縣市"] == str(county).strip(),
                    "鄉鎮市區",
                ]
                .dropna()
                .unique()
            )

            if not town_options:
                town_options = [""]

            town = st.selectbox("鄉鎮市區", town_options, key="single_town")
        else:
            county = st.text_input("縣市", value="", key="single_county_manual")
            town = st.text_input("鄉鎮市區", value="", key="single_town_manual")

    with location_col2:
        st.empty()

    with st.form("single_site_form"):
        c1, c2 = st.columns(2)

        with c1:
            site_code = st.text_input("代碼", value="SITE-001")
            site_name = st.text_input("基地名稱", value="")

        with c2:
            twd97x_text = st.text_input("TWD97X（請輸入6碼數字）", value="", max_chars=6)
            twd97y_text = st.text_input("TWD97Y（請輸入7碼數字）", value="", max_chars=7)

        st.markdown("#### 2. 道路、管理與基地條件")
        n1, n2, n3 = st.columns(3)

        with n1:
            road_distance_text = st.text_input("距主要道路距離(公尺)", value="999")
            land_ownership = st.selectbox("土地權屬", ["公有或單一權屬", "公私混合或需協調", "私有或權屬複雜", "其他／待確認"])

        with n2:
            management_agency = st.selectbox(
                "管理機關",
                ["管理機關明確", "需確認或協調", "無明確管理單位"],
            )
            area_text = st.text_input("基地面積(公頃)", value="")

        with n3:
            length_text = st.text_input("基地長度(公里)", value="")
            short_term = st.selectbox("短期推動性", ["可短期推動", "需部分協調", "短期難以推動", "待評估"])
            openness = st.selectbox("開放可及性", ["完全開放", "部分開放", "不開放", "待確認"])

        st.markdown("#### 3. 功能角色判釋資料")
        st.caption("請依基地現況回答。系統會自動整理可能符合的功能情境。")
        y1, y2 = st.columns(2)
        yes_no_options = ["否", "是", "未知"]

        with y1:
            q1 = st.selectbox("基地內是否有人會停留使用？（如休憩、運動、活動、等候）", yes_no_options, index=0)
            q2 = st.selectbox("基地內或周邊是否有人會經過？（如步行、通學、騎車、穿越動線）", yes_no_options, index=0)
            q3 = st.selectbox("基地是否需要在邊界做防護？（如道路、工廠、空地、住宅或學校交界）", yes_no_options, index=0)
            q5 = st.selectbox("道路旁是否有需要保護的對象？（如住宅、學校、醫療院所、社區活動空間）", yes_no_options, index=0)

        with y2:
            q6 = st.selectbox("基地或周邊是否有臨時污染事件？（如施工、整地、土方、堆置、臨時活動）", yes_no_options, index=0)
            q7 = st.selectbox("上述臨時事件是否影響人們停留或活動的空間？", yes_no_options, index=0)
            q8 = st.selectbox("上述臨時事件是否影響步行、通學、騎車或通行路線？", yes_no_options, index=0)
            q9 = st.selectbox("上述臨時事件是否靠近邊界、周界、住宅、學校或醫療院所？", yes_no_options, index=0)

        submitted = st.form_submit_button("開始分析", type="primary")

    if not submitted:
        return None

    if not (twd97x_text.isdigit() and len(twd97x_text) == 6):
        st.error("TWD97X 必須為 6 碼數字。")
        return None

    if not (twd97y_text.isdigit() and len(twd97y_text) == 7):
        st.error("TWD97Y 必須為 7 碼數字。")
        return None

    def parse_nonnegative_number(text: str, label: str, required: bool = False):
        text = str(text).strip()
        if text == "":
            if required:
                st.error(f"{label} 必須填寫。")
                return None
            return np.nan
        try:
            value = float(text.replace(",", ""))
        except ValueError:
            st.error(f"{label} 必須為數字。")
            return None
        if value < 0:
            st.error(f"{label} 不可為負數。")
            return None
        return value

    road_distance_value = parse_nonnegative_number(road_distance_text, "距主要道路距離(公尺)", required=True)
    area_value = parse_nonnegative_number(area_text, "基地面積(公頃)", required=False)
    length_value = parse_nonnegative_number(length_text, "基地長度(公里)", required=False)

    if road_distance_value is None or area_value is None or length_value is None:
        return None

    row = {
        "代碼": site_code,
        "縣市": county,
        "鄉鎮市區": town,
        "基地名稱": site_name,
        "TWD97X": int(twd97x_text),
        "TWD97Y": int(twd97y_text),
        "距主要道路距離(公尺)": road_distance_value,
        "土地權屬": land_ownership,
        "管理機關": management_agency,
        "基地面積(公頃)": area_value if not pd.isna(area_value) and area_value > 0 else np.nan,
        "基地長度(公里)": length_value if not pd.isna(length_value) and length_value > 0 else np.nan,
        "短期推動性": short_term,
        "開放可及性": openness,
        "基地內部是否有停留活動空間": q1,
        "是否有步行通學騎行路徑": q2,
        "是否有邊界防護需求": q3,
        "道路周邊是否有敏感受體或社區活動空間": q5,
        "是否有短期事件": q6,
        "短期事件是否影響內部空間": q7,
        "短期事件是否影響通行路徑": q8,
        "短期事件是否位於邊界或鄰近受體": q9,
    }

    return pd.DataFrame([row])


# ============================================================
# 十、Streamlit 主介面
# ============================================================

st.title("🌿 城市空品植生綠網淨化單元分析工具")
st.caption("可進行批次上傳分析，也可進行單點手動輸入。系統會依座標與外部資料自動補算周邊指標。")

with st.sidebar:
    st.header("空白制式表單")
    st.write("可下載空白表單後，填寫批次場址資料再上傳分析。")

    template_excel = make_blank_template_excel()
    if template_excel is not None:
        st.download_button(
            label="下載批次上傳空白表單 Excel（含欄位說明）",
            data=template_excel,
            file_name="城市空品植生綠網淨化單元_批次上傳空白表單.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.download_button(
        label="下載批次上傳空白表單 CSV",
        data=make_blank_template_csv(),
        file_name="城市空品植生綠網淨化單元_批次上傳空白表單.csv",
        mime="text/csv",
    )

    st.divider()
    st.header("外部資料")
    st.write("系統會自動補算對應指標。")
    with st.expander("自動帶入資料來源"):
        st.markdown(
            """
            - **人口密度**：114年人口密度資料，1150427下載  
            - **工廠資料**：環境保護許可管理系統（暨解除列管）對象基本資料－曾有空列管，1150427下載  
            - **空品淨化區**：114年第四季季報  
            - **綠牆**：113年第四季季報，面積由平方公尺換算為公頃  
            - **各級學校**：111學年度各級學校名錄  
            - **醫療院所**：112年12月醫療院所分布圖  
            """
        )

    st.divider()
    st.header("計分與名詞說明")
    with st.expander("節點潛力"):
        st.markdown(
            """
            節點潛力滿分 18 分，共 9 項指標，每項 0–2 分。  
            評估基地是否適合作為優先投入據點，包含空品壓力、敏感受體、人口暴露、土地權屬、管理權責、可植栽空間、短期推動性、開放可及性與生活節點。
            """
        )
    with st.expander("串聯潛力"):
        st.markdown(
            """
            串聯潛力滿分 8 分，共 4 項指標，每項 0–2 分。  
            評估基地與周邊綠化單元形成空間連接與網絡支撐的可能性。
            """
        )
    with st.expander("敏感受體與生活節點"):
        st.markdown(
            """
            **敏感受體**指較需要空氣品質防護的對象，例如學校、醫療院所等。  
            這類場所通常聚集兒童、學生、病患、高齡者或其他較易受空氣污染影響的人群。  

            **生活節點**指基地周邊與民眾日常使用、公共服務、健康照護或活動停留相關的設施。  
            目前本工具先以學校與醫療院所作為生活節點的初步判斷基礎。  

            未來生活節點可再擴充納入：  
            - 公園、廣場、活動中心  
            - 市場、商圈、車站、公車轉運站  
            - 圖書館、行政服務據點  
            - 長照機構、托嬰中心、社福設施  
            - 其他具日常使用或公共服務功能之場所  

            因此目前版本中，敏感受體與生活節點暫時採相同基礎資料進行判斷；  
            未來若新增更多生活節點資料，可再將兩者比對邏輯分開。
            """
        )
    with st.expander("綠色基礎設施／綠化單元"):
        st.markdown(
            """
            綠化單元指可提供綠覆、滯塵、緩衝、降溫、景觀或綠網串聯功能的空間。  
            目前系統整合空品淨化區與綠牆資料；綠牆面積會由平方公尺換算為公頃。
            """
        )
    with st.expander("短期事件"):
        st.markdown(
            """
            短期事件指基地或周邊短時間、階段性或臨時性的污染或暴露情境，  
            例如施工、整地、土方、臨時堆置、臨時活動或短期作業。  
            這類情境通常應優先檢討源頭管理與作業管理，植生多作為補強措施。
            """
        )
    with st.expander("優先推動建議"):
        st.markdown(
            """
            系統會整合節點潛力、串聯潛力與功能情境，產出推動建議。主要類別如下：  
            - **建議優先推動**：節點與串聯條件皆佳。  
            - **建議優先強化基地功能**：節點條件佳，但串聯條件尚未達高潛力。  
            - **建議作為綠網連接補點**：串聯條件佳，可支撐綠網連接。  
            - **建議納入第二階段評估**：節點與串聯條件中等。  
            - **建議作為局部連接或補強場址**：串聯條件尚可，但節點條件較基礎。  
            - **建議視管理可行性再評估**：節點條件尚可，但串聯條件較基礎。  
            - **建議暫列低優先序**：節點與串聯條件皆低。  
            """
        )

input_mode = st.radio(
    "選擇分析模式",
    ["批次上傳分析", "單點手動輸入"],
    horizontal=True,
)

if input_mode == "批次上傳分析":
    st.subheader("一、批次上傳")
    uploaded_file = st.file_uploader("請上傳綠化單元檔案，可使用 Excel 或 CSV", type=["xlsx", "xls", "csv"])

    if uploaded_file is None:
        st.info("請上傳 Excel 或 CSV 檔案後開始分析。")
        st.stop()

    try:
        if uploaded_file.name.lower().endswith(".csv"):
            raw_df = read_csv_with_fallback(uploaded_file)
        else:
            raw_df = pd.read_excel(uploaded_file)
    except ImportError:
        st.error("目前環境缺少 openpyxl，無法讀取 Excel。你可以先把 Excel 另存成 CSV 後再上傳。")
        st.stop()
    except Exception as e:
        st.error(f"讀取檔案失敗：{e}")
        st.stop()

    raw_df = normalize_columns(raw_df)

    st.subheader("欄位檢查")
    is_valid, missing = check_required_columns(raw_df)

    if not is_valid:
        st.error("欄位檢查未通過，請補齊以下欄位後重新上傳。")
        st.code("\n".join(missing))
        with st.expander("查看目前讀到的欄位"):
            st.write(list(raw_df.columns))
        st.stop()

    raw_df = ensure_input_defaults(raw_df)

    st.success("欄位檢查通過。系統將依座標與外部資料自動補算人口密度、工廠數、生活節點與綠化單元指標。")

    with st.expander("查看原始資料"):
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    result_df = analyze_green_network(raw_df)
    render_full_results(result_df, mode="batch")

else:
    st.subheader("一、單點手動輸入")
    st.info("請填寫單一基地資料。系統會依座標與外部資料自動補算人口密度、工廠數、生活節點與綠化單元指標。")

    single_df = build_single_site_dataframe()

    if single_df is None:
        st.stop()

    result_df = analyze_green_network(single_df)
    render_full_results(result_df, mode="single")

st.caption(
    "提醒：本工具目前為初版規則引擎，適合做原型測試、場址排序與初步盤點。"
    "正式決策前仍建議加入專家審查、現地查核、資料校正與維護管理檢核。"
)
