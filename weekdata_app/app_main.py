# weekdata_app/app_main.py
from typing import Optional
import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import sys
import json
import datetime
import platform
import difflib
from pathlib import Path

# =========================
# 基础路径与资源定位
# =========================
APP_NAME = "双周回流分析自动处理器"

def resource_path(relative_path: str) -> str: #1定义资源路径函数,程序先查找打包后根目录文件，并定义为基础路径变量名，若无根目录则判定为源码态，使用路径函数获取相对父路径目录，返回组合后完整路径
    """
    兼容打包态/源码态的资源定位（内置 Excel 模板所在位置）：
    - 源码态：weekdata_app/ 目录下
    - 打包态（PyInstaller）：_MEIPASS 根目录下
    （程序在打包时把两个模板 xlsx 加到 _MEIPASS 根目录）
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS              # PyInstaller 解包目录
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  # weekdata_app/
    return os.path.join(base_path, relative_path)


def base_dir_for_user_visible_data() -> Path:  #2定义用户可见的数据目录函数
    """
    data/ 用于让用户看得见、改得动的目录。

    设计：
    - 源码态：使用项目根目录（mac包），即 weekdata_app 的上一层
              mac包/
                ├─ weekdata_app/
                ├─ launcher.py
                ├─ run_stub.py
                └─ data/   ← 这里
    - 打包态（.app）：
        假设可执行文件路径类似：
        /XXX/YYY/双周回流分析自动处理器.app/Contents/MacOS/双周回流分析自动处理器

        则：
        exe.parent           = .../MacOS
        exe.parents[1]       = .../Contents
        exe.parents[2]       = .../双周回流分析自动处理器.app
        exe.parents[3]       = .../XXX/YYY   ← 我们希望 data 出现在这里

        也就是说：无论你把 .app 放桌面还是其它文件夹，
        data/ 都会和 .app 并排出现，用户一眼就能看到。
    """
    exe = Path(sys.executable).resolve()
    if getattr(sys, "frozen", False):
        # 尝试取 .app 所在目录（.app 的父目录）
        try:
            return exe.parents[3]
        except IndexError:
            # 兜底：至少保证和二进制同级
            return exe.parent
    else:
        # 源码态：weekdata_app 的父目录（mac包）
        return Path(__file__).resolve().parent.parent


BASE_DIR = base_dir_for_user_visible_data()  # 用户可见的数据目录
DATA_DIR = BASE_DIR / "data"  #
DATA_DIR.mkdir(parents=True, exist_ok=True)

# 内置模板（随包）—— 这两份是“只读模板”，首次运行会被复制一份到 data/ 里
# 请把原始模板文件拷贝到 weekdata_app/ 目录下（与 app_main.py 同级）：
#   weekdata_app/类型具体名称别名映射表.xlsx
#   weekdata_app/历史周度数据表.xlsx
TEMPLATE_MAPPING_FILE = resource_path("类型具体名称别名映射表.xlsx")  #相对路径，字符串形式
TEMPLATE_HISTORY_FILE = resource_path("历史周度数据表.xlsx")

# 可见的用户数据文件（data/ 下，可被用户随意替换）
USER_MAPPING_FILE = DATA_DIR / "类型具体名称别名映射表.xlsx"
USER_HISTORY_FILE = DATA_DIR / "历史周度数据表.xlsx"  # 当前默认使用的历史表（覆盖指向）
DYNAMIC_MAP_FILE = DATA_DIR / "dynamic_category_maps.json" # 动态映射表json里面为字典套字典的形式，一个为model_map为型号类映射表字典，一个为category_map为类别类映射表字典


# =========================
# data/ 首次初始化
# =========================
def ensure_file_from_template(dst: Path, template_path: str) -> None:  #3确保文件从模板文件创建
    """
    确保目标文件从模板文件创建（仅首次）。

    如果目标文件不存在，则尝试从模板文件创建它。
    优先尝试以Excel格式读取并写入，如果失败则退化为二进制复制。

    Args:
        dst: 目标文件路径（Path对象）
        template_path: 模板文件路径（字符串）

    Returns:
        None
    """
    if dst.exists():  #如果目标文件存在，就直接返回（则会自动跳过该步骤）
        return
    if not os.path.exists(template_path): #如果模板文件不存在，就直接返回
        return
    try:
        data = pd.read_excel(template_path, sheet_name=None, engine="openpyxl")
        with pd.ExcelWriter(dst, engine="openpyxl") as writer:
            for s, df in data.items():
                df.to_excel(writer, sheet_name=s, index=False)
    except Exception:
        # 如果不是合法 Excel，就退化为二进制复制
        try:
            with open(template_path, "rb") as fsrc, open(dst, "wb") as fdst:
                fdst.write(fsrc.read())
        except Exception:
            pass


# 首次运行：把模板复制到 data/
ensure_file_from_template(USER_MAPPING_FILE, TEMPLATE_MAPPING_FILE)
ensure_file_from_template(USER_HISTORY_FILE, TEMPLATE_HISTORY_FILE)

# =========================
# 工具函数
# =========================
def read_excel_bytesio(b):  #4读取Excel文件，返回字典形式的数据

    return pd.read_excel(io.BytesIO(b), sheet_name=None, engine="openpyxl")


def normalize(s):  #5规范化字符串，去除空格和首尾空格，并转换为小写

    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def load_dynamic_maps(): #6加载动态映射表
    if not DYNAMIC_MAP_FILE.exists(): #如果动态映射表不存在，就返回空字典，相当于先创建一个空表

        return {}, {}
    try:
        with open(DYNAMIC_MAP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f) or {}
        model_map = data.get("model_map", {}) #获取映射表中的型号类映射
        cat_map = data.get("category_map", {}) #获取映射表中的类别映射
        return (model_map if isinstance(model_map, dict) else {},
                cat_map if isinstance(cat_map, dict) else {})
    except Exception:
        return {}, {}


def save_dynamic_maps(model_map: dict, cat_map: dict): #7保存动态映射表

    try:
        with open(DYNAMIC_MAP_FILE, "w", encoding="utf-8") as f:
            json.dump({"model_map": model_map, "category_map": cat_map}, f, ensure_ascii=False, indent=2) #将映射表写入json文件,使用utf-8编码，ensure_ascii=False保证中文不转义，indent=2表示缩进2个空格

        return True, None #成功，没有错误
    except Exception as e:
        return False, str(e) #失败，有错误


def load_mapping_df(): #8加载映射表
    # 优先 data/ 下
    if USER_MAPPING_FILE.exists():
        try:
            return pd.read_excel(USER_MAPPING_FILE, engine="openpyxl")
        except Exception as e:
            st.warning(f"无法读取 data/ 下的映射表：{e}，尝试使用内置模板")
    # 回退模板
    if os.path.exists(TEMPLATE_MAPPING_FILE):
        try:
            return pd.read_excel(TEMPLATE_MAPPING_FILE, engine="openpyxl")
        except Exception as e:
            st.warning(f"无法读取内置映射表：{e}")
            return None
    st.warning("未找到映射表，将仅做精确匹配，可能导致匹配不到字段值。")
    return None


def find_latest_history_in_data() -> Path:  #9寻找历史表
    """
    返回 data/ 下最合适的历史文件：
    - 优先 data/历史周度数据表.xlsx
    - 否则寻找 pattern: 历史周度数据表_YYYYMMDD.xlsx 最新的一个
    - 再否则回 None
    """
    if USER_HISTORY_FILE.exists():
        return USER_HISTORY_FILE
    candidates = []
    for p in DATA_DIR.glob("历史周度数据表_*.xlsx"):
        try:
            mtime = p.stat().st_mtime
            candidates.append((mtime, p))
        except Exception:
            continue
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    return None


def load_history_sheets(prefer_path: Optional[Path] = None):  #10加载历史表，这里的prefer_path:Optional[Path]是可选参数，默认为None表示默认用户这里未上传最新历史周度表，则会从后台data文件夹中读取

    """
    加载历史数据表，支持手动指定路径或自动查找最新数据表。

    函数会尝试从指定路径或data目录下加载最新的历史周度数据表，用于计算环比数据。

    Args:
        prefer_path (Optional[Path]): 可选参数，用户指定的历史表路径。
            如果未提供，则自动查找data目录下最新的历史表。

    Returns:
        tuple: 包含两个元素的元组:
            - dict: 加载成功返回包含所有sheet的数据字典，失败返回空字典
            - str/None: 成功时返回文件路径字符串，失败返回None

    Note:
        当找不到可用历史表时，会显示提示信息并返回空字典和None。
    """
    path = prefer_path or find_latest_history_in_data() #定义历史表路径来自于两种情况：1是prefer_path为用户界面前端手动上传的历史周度表路径，2是加载最新日期的历史周度数据表
    if path is None:
        st.info("未找到 data/ 下可用的历史表，将仅处理本周，不计算环比。")
        return {}, None #这里的None是None，而不是空字符串，return最终返回的是两个空字典
    try:
        hist = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        return hist, str(path)
    except Exception as e:
        st.warning(f"读取历史表失败：{e}（文件：{path}）")
        return {}, None


def find_mapping_for_sheet(mapping_df, sheet_name):  #11寻找映射表，这里的mapping_df是映射表的DataFrame，sheet_name是当前处理的sheet名称
    if mapping_df is None:
        return None
    cols = [c.lower() for c in mapping_df.columns] #这里的c.lower()是将映射表的所有列名转换为小写，然后存储在列表cols中


    def pick(col_options): #寻找映射表，这里的col_options是列名称列表
        """
        在映射表中查找与给定列名列表匹配的列名。

        Args:
            col_options (list): 待匹配的列名列表

        Returns:
            str/None: 返回映射表中第一个匹配的列名，若无匹配则返回None
        """
        # 遍历输入的列名选项列表
        for opt in col_options:
            # 检查当前选项是否在预定义的列名列表(cols)中
            if opt in cols:
                # 返回映射表中对应的列名
                return mapping_df.columns[cols.index(opt)]
        # 如果没有找到匹配项，返回None
        return None

    sheet_col = pick(['子表'])
    sourcefield_col = pick(['来源列', 'source列'])
    alias_col = pick(['alias', '别名'])
    sourceval_col = pick(['源值', 'source_value'])
    if sheet_col is None or sourcefield_col is None or alias_col is None or sourceval_col is None:
        try:
            sheet_col = sheet_col or mapping_df.columns[0]
            sourcefield_col = sourcefield_col or mapping_df.columns[1]
            alias_col = alias_col or mapping_df.columns[2]
            sourceval_col = sourceval_col or mapping_df.columns[3]
        except:
            return None
    sub = mapping_df[[sheet_col, sourcefield_col, alias_col, sourceval_col]].copy()
    sub.columns = ['sheet', 'source_field', 'alias', 'source_value']  #这里的sub.columns是将子表、来源列、别名、源值列重命名为sheet、source_field、alias、source_value

    sub['sheet_norm'] = sub['sheet'].apply(lambda x: normalize(x)) #这里的sub['sheet_norm']是将子表列的值转换为小写，然后存储在新的列中

    return sub


def _to_number_series_for_base_metrics(series: pd.Series) -> pd.Series:  #12将基础指标列转换为数值类型
    """
    尽量贴近你旧 app.py 的行为：
    1）优先把 NaN 和 '-' 当成 0，直接 astype(float) 求和（= 旧版逻辑）
    2）如果整列实在转不动，再退回到“谨慎解析”逻辑
    """
    # 如果本来就是数值列，直接返回即可
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    # ① 主路径：完全模拟旧 app.py 的用法
    s = series.copy()
    s = s.replace({np.nan: 0, "-": 0})
    try:
        return s.astype(float) #s.astype(float)是将series转换为浮点数
    except Exception:  #这里的异常处理再次设置了一层兜底S2,先转换为字符串再去除前后空格，再将空字符串和-转换为NaN，同时过滤掉包含百分号的值，移除数字字符串，最后尝试转换为数值类型


        # ② 兜底：再按你现在这版的“安全解析”再试一次
        s2 = series.astype(str).str.strip()
        s2 = s2.replace({"": np.nan, "-": np.nan})
        # 基础计数/金额列里如果出现百分号，一律当作无效
        s2 = s2.where(~s2.str.contains("%"), np.nan)
        # 去掉英文千分位逗号
        s2 = s2.str.replace(",", "", regex=False)
        return pd.to_numeric(s2, errors="coerce")


def _to_float_or_none(x):  #13将x转换为浮点数或None
    if x is None or (isinstance(x, float) and np.isnan(x)):  #这里的isinstance的意思是判断x是否是None或者x是float类型且是NaN

        return None
    s = str(x).strip()
    if s in ('', '-'):
        return None
    s2 = s.replace(',', '')
    if s2.endswith('%'):
        try:
            return float(s2[:-1]) / 100.0
        except Exception:
            return None
    try:
        return float(s2)
    except Exception:
        return None


SOURCE_FIELD_MAP = {
    '展现pv': '展现量', '展现': '展现量', 'impression': '展现量',
    '点击pv': '点击量', '点击': '点击量', 'click': '点击量',
    '总费用': '花费', '费用': '花费', 'cost': '花费',
    '总订单行': '成交量', '成交量': '成交量', 'orders': '成交量', 'total_orders': '成交量',
    '总订单金额': '成交额', '成交额': '成交额', 'gmv': '成交额', 'order_amount': '成交额',
    '总加购数': '加购量', '加购数': '加购量', '加购': '加购量'
}


def map_source_columns(cols): #14映射源数据列，将用户上传的源数据子表中的列名映射与待填写的指标名称精准对照上，比如'展现量PV=展现量'，存入字典中用于后续准确取数


    mapping = {}
    for c in cols:
        nc = normalize(c)
        for k, v in SOURCE_FIELD_MAP.items():
            if k == nc:
                mapping[c] = v
                break
    return mapping


def candidate_source_key(source_df):  #15寻找源数据列
    for key in ['资源位', '品类', '产品名称', 'product_name', 'category', 'resource', '资源位', '资源']:
        for c in source_df.columns:
            if key == c.lower():
                return c
    for c in source_df.columns:
        if source_df[c].dtype == object: 
            return c
    return source_df.columns[0]


def compute_metrics(sum_dict, history_row=None, history_totals=None): #16计算指标,sum_dict是指标汇总字典，history_row是历史数据行，history_totals是历史总量
    res = {}
    for k in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量']:
        val = sum_dict.get(k, None) 
        if val is None or (isinstance(val, float) and np.isnan(val)):
            res[k] = None
        else:
            try:
                res[k] = float(val)
            except:
                res[k] = None

    def safe_div(a, b):
        try:
            a = float(a)
            b = float(b)
            if b == 0:
                return None
            return a / b
        except:
            return None

    roi = safe_div(res.get('成交额'), res.get('花费'))
    res['ROI'] = roi
    res['点击率'] = safe_div(res.get('点击量'), res.get('展现量'))
    res['点击转化率'] = safe_div(res.get('成交量'), res.get('点击量'))
    res['CPC'] = safe_div(res.get('花费'), res.get('点击量'))
    res['加购成本'] = safe_div(res.get('花费'), res.get('加购量'))

    if history_row is not None:
        prev_cost = history_row.get('花费')
        prev_roi = history_row.get('ROI')
        try:
            res['消耗环比'] = None if (res.get('花费') is None or prev_cost in [None, 0, np.nan]) else (
                        res.get('花费') / prev_cost - 1)
        except:
            res['消耗环比'] = None
        try:
            res['ROI环比'] = None if (res.get('ROI') is None or prev_roi in [None, 0, np.nan]) else (
                        res.get('ROI') / prev_roi - 1)
        except:
            res['ROI环比'] = None
    else:
        res['消耗环比'] = None
        res['ROI环比'] = None
    return res


def format_value(key, val): #17格式化值，key是指标名称，val是指标值
    if val is None:
        return "-"
    if key in ['消耗占比', '成交占比', '消耗环比', 'ROI环比', '合计环比']:
        try:
            return f"{int(round(val * 100, 0))}%"
        except:
            return "-"
    if key in ['点击率', '点击转化率']:
        try:
            return f"{val * 100:.1f}%"
        except:
            return "-"
    if key in ['ROI', 'CPC']:
        try:
            return f"{val:.2f}"
        except:
            return "-"
    if key == '加购成本':
        try:
            return f"{val:.1f}"
        except:
            return "-"
    try:
        if abs(val - round(val)) < 1e-9:  #如果差值绝对值小于1e-9，即小数点后9位都为0，则返回字符串形式的整数


            return str(int(round(val)))
        else:
            return f"{val:.2f}"
    except:
        return str(val)


# ========== 运行时映射（内置 + 持久化） ==========
BASE_CATEGORY_MAP = {
    "平板型号": "家庭教育",
    "平板电脑型号": "闺蜜机",
    "有屏型号": "智能屏",
    "无屏型号": "无屏音箱",
    "健身镜型号": "健身镜",
    "IPC型号": "IPC(摄像机)",
}
BASE_CATEGORY_SHEET_TO_CATEGORY = {
    "家庭教育": "家庭教育",
    "闺蜜机": "闺蜜机",
    "智能屏": "智能屏",
    "无屏音箱": "无屏音箱",
    "健身镜": "健身镜",
    "IPC": "IPC(摄像机)",
}

# =========================
# 校验器
# =========================
def validate_upload(sheet_names, source_df, runtime_category_map, runtime_category_sheet_to_category): #18该函数表示校验上传的数据，包括源数据、子表和型号类，并返回校验结果

    rows = []
    suggestions_model = []
    suggestions_cat = []
    has_error = False

    src_cols = list(source_df.columns)
    need_cols_any = [('资源', '资源位')]
    need_cols_all = ['品类', '产品名称']

    mapped = map_source_columns(src_cols)
    has_any_metric = any(k in mapped for k in src_cols) or len(mapped) > 0

    for c in need_cols_all:
        ok = (c in src_cols)
        status = "OK" if ok else "ERROR"
        if not ok:
            has_error = True
        rows.append({"对象": "源数据", "类型": "字段", "名称": c, "状态": status, "说明": "" if ok else f"缺少列：{c}"})

    any_ok = False
    any_hit = None
    for a, b in need_cols_any:
        if a in src_cols or b in src_cols:
            any_ok = True
            any_hit = a if a in src_cols else b
            break
    if not any_ok:
        has_error = True
        rows.append(
            {"对象": "源数据", "类型": "字段", "名称": "资源 / 资源位", "状态": "ERROR", "说明": "需包含“资源”或“资源位”任一列"})
    else:
        rows.append({"对象": "源数据", "类型": "字段", "名称": any_hit, "状态": "OK", "说明": "用于资源分组"})

    rows.append({"对象": "源数据", "类型": "字段", "名称": "指标列映射", "状态": "OK" if has_any_metric else "WARN",
                 "说明": "已识别指标映射" if has_any_metric else "未识别到标准指标列映射，可能无法汇总"})

    categories = set()
    if '品类' in source_df.columns:
        categories = set([str(x).strip() for x in source_df['品类'].dropna().unique().tolist()]) #获取源数据中的品类列，并将该列值转换为字符串并去除空值NaN，然后去重复后将数组转为列表，最终转换为集合

    special = {'双周分资源位', '双周分品类'}
    for s in sheet_names:
        if s in special:
            rows.append({"对象": s, "类型": "子表", "名称": s, "状态": "OK", "说明": "特殊子表（固定口径）"})
            continue
        if s in runtime_category_map or s.endswith("型号"):
            if s in runtime_category_map:
                cat = runtime_category_map[s]
                exists = (cat in categories) if categories else True
                status = "OK" if exists else "WARN"
                rows.append({"对象": s, "类型": "型号类", "名称": f"{s}→{cat}", "状态": status,
                             "说明": "" if exists else f"品类“{cat}”在本次数据源中无记录"})
            else:
                base = s[:-2] if s.endswith("型号") else s
                candidates = list(categories) if categories else []
                suggestion = None
                if base and candidates:
                    if base in candidates:
                        suggestion = base
                    else:
                        m = difflib.get_close_matches(base, candidates, n=1, cutoff=0.6) #获取最接近的匹配项
                        if m:
                            suggestion = m[0] 
                msg = f"未配置映射：建议 “{s},{suggestion}”" if suggestion else "未配置映射：请在映射面板添加"
                rows.append({"对象": s, "类型": "型号类", "名称": s, "状态": "WARN", "说明": msg})
                if suggestion:
                    suggestions_model.append((s, suggestion))
        else:
            if s in runtime_category_sheet_to_category:
                cat = runtime_category_sheet_to_category[s]
                exists = (cat in categories) if categories else True
                status = "OK" if exists else "WARN"
                rows.append({"对象": s, "类型": "类别类", "名称": f"{s}→{cat}", "状态": status,
                             "说明": "" if exists else f"品类“{cat}”在本次数据源中无记录"})
            else:
                base = s
                candidates = list(categories) if categories else []
                suggestion = None
                if base and candidates:
                    if base in candidates:
                        suggestion = base
                    else:
                        m = difflib.get_close_matches(base, candidates, n=1, cutoff=0.6)
                        if m:
                            suggestion = m[0]
                msg = f"未配置映射：建议 “{s},{suggestion}”" if suggestion else "未配置映射：请在映射面板添加"
                rows.append({"对象": s, "类型": "类别类", "名称": s, "状态": "WARN", "说明": msg})
                if suggestion:
                    suggestions_cat.append((s, suggestion))

    report_df = pd.DataFrame(rows, columns=["对象", "类型", "名称", "状态", "说明"])
    return report_df, suggestions_model, suggestions_cat, has_error


# =========================
# 核心处理逻辑（与你版本一致）
# =========================
def process_subtable(template_df, source_df, history_df, mapping_for_sheet, src_group_col, sheet_name,
                     runtime_category_map, runtime_category_sheet_to_category): #19该函数表示处理子表，包括源数据、历史数据、映射、源分组列、子表名称、型号类映射和类别类映射，并返回处理后的结果

    out_rows = []
    src_col_map = map_source_columns(source_df.columns)
    inv_map = {v: k for k, v in src_col_map.items()} #反向映射，即键值对调


    type_col = next(
        (c for c in template_df.columns if '类型' in c or '型号' in c or c.lower() == 'type'), template_df.columns[0]) #获取类型列，如果没有，则取第一个列

    types = template_df[type_col].astype(str).tolist() #获取所有类型，并将其转换为字符串列表


    order_amount_map = {} # 针对模版表中手动填的订单出库金额映射

    if sheet_name == '双周分品类' and '订单出库金额' in template_df.columns:

        def _to_float_user_amount(x): #先定义一个数据格式函：将订单出库金额转换为浮点数，并保留一位小数

            if x is None or (isinstance(x, float) and np.isnan(x)):
                return None
            s = str(x).strip().replace(',', '') #去除逗号，并转换为字符串

            if s in ('', '-'): #如果s为空或'-'，则返回None

                return None
            try:
                return float(s)
            except:
                return None

        for _, r in template_df.iterrows(): #遍历每一行，除 合计 行之外，其他行都需要处理


            tname = str(r.get(type_col, '')).strip() #获取类型名称，并去除空格

            if tname:
                order_amount_map[normalize(tname)] = _to_float_user_amount(r.get('订单出库金额'))

    def _fmt_rate_one_decimal(x): #格式化百分比，保留一位小数

        if x is None:
            return '-'
        try:
            return f"{float(x) * 100:.1f}%"
        except:
            return '-'

    # --- 构建历史行索引：用于行级“消耗环比 / ROI环比” ---
    # === 历史行索引：兼容百分比 & 别名映射 ===
    hist_lookup = {}
    if history_df is not None and len(history_df.columns) > 0:
        hist_type_col = next(
            (c for c in history_df.columns if '类型' in c or '型号' in c or c.lower() == 'type'),
            history_df.columns[0]
        )

        # 如果有 mapping_for_sheet，则构造 alias <-> source_value 的双向映射
        alias_map = {}
        if mapping_for_sheet is not None and not mapping_for_sheet.empty: #mapping_for_sheet 是 pandas.DataFrame
            mf = mapping_for_sheet.copy()
            # 注意：mapping_for_sheet 在外面已经标准化过 sheet_norm，这里只关心 alias/source_value
            mf['alias_norm'] = mf['alias'].apply(lambda x: normalize(x)) # 标准化 alias

            mf['source_norm'] = mf['source_value'].apply(lambda x: normalize(x)) # 标准化 source_value

            for _, m in mf.iterrows(): #以下功能构建了一个双向的映射关系，支持双向名称查找
                a = m.get('alias_norm') or ''
                b = m.get('source_norm') or ''
                if a:
                    alias_map.setdefault(a, set()).add(a)
                    if b:
                        alias_map[a].add(b)
                if b:
                    alias_map.setdefault(b, set()).add(b)
                    if a:
                        alias_map[b].add(a)

        metric_cols_hist = ['展现量', '点击量', '花费', '成交量', '成交额', '加购量', 'ROI']

        def _parse_num_or_percent_hist(x): #解析百分比或数字，返回浮点数
            if x is None or (isinstance(x, float) and np.isnan(x)):
                return None
            s = str(x).strip().replace(',', '') #去除逗号，并转换为字符串
            if s in ('', '-'):
                return None
            if s.endswith('%'):
                try:
                    return float(s[:-1]) / 100.0 # 百分比转换为浮点数
                except:
                    return None
            try:
                return float(s)
            except:
                return None

        for _, r in history_df.iterrows():
            raw_name = r.get(hist_type_col, '')
            key0 = normalize(raw_name)
            if not key0:
                continue

            # 把这一行所有需要的字段都用解析函数做一次解析
            metrics = {}
            for k in metric_cols_hist:
                if k in history_df.columns:
                    metrics[k] = _parse_num_or_percent_hist(r.get(k))
                else:
                    metrics[k] = None

            # 主键：自己
            keys_for_this_row = set([key0])
            # 如果在 alias_map 里有别名 / 源值对应关系，一并挂到同一份 metrics 上
            if key0 in alias_map:
                keys_for_this_row |= alias_map[key0]  #｜=是集合的“合并更新”操作符，意思将两个集合元素合并，且更新到左边集合中


            for k in keys_for_this_row:
                if k and k not in hist_lookup: #如果k存在且不在hist_lookup中，则添加到hist_lookup中，避免重复


                    hist_lookup[k] = metrics

    # —— 利用映射表构建 “别名 ↔ 源值” 的双向索引（用于历史表名称回溯）——
    alias2src = {}
    src2alias = {}
    if mapping_for_sheet is not None:
        try:
            for _, row in mapping_for_sheet.iterrows():
                a = normalize(row.get("alias", ""))
                sname = normalize(row.get("source_value", ""))
                if a:
                    alias2src.setdefault(a, set()).add(sname)
                if sname:
                    src2alias.setdefault(sname, set()).add(a)
        except Exception:
            # 映射表坏掉就当没有，降级为原来的精确匹配逻辑
            alias2src = {}
            src2alias = {}

    weekly_totals = {k: 0.0 for k in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量']} # 周总计

    filtered_source = source_df.copy() # 使用source_df.copy()方法创建一个新的DataFrame，并赋值给filtered_source变量,独立副本避免修改原始数据
    fs = filtered_source #这里的filtered_source变量指向source_df，即原始数据源，filtered_source只是source_df的一个副本，可以对filtered_source进行修改而不影响原始数据源source_df。


    if sheet_name in runtime_category_sheet_to_category:
        cat = runtime_category_sheet_to_category[sheet_name]
        if '品类' in fs.columns:
            fs = fs[fs['品类'] == cat]
        if '资源' in fs.columns:
            group_col = '资源'
        else:
            group_col = '资源位'
    elif sheet_name in runtime_category_map:
        cat = runtime_category_map.get(sheet_name)
        if cat and '品类' in fs.columns:
            fs = fs[fs['品类'] == cat]
        group_col = '产品名称'
    elif sheet_name == '双周分资源位':
        group_col = '资源位'
    elif sheet_name == '双周分品类':
        group_col = '品类'
    else:
        group_col = src_group_col

    filtered_source = fs #这里的filtered_source因为前面的filtered_source = source_df，所以它仍然是指向source_df的，但是通过过滤筛选后重新将筛选后数据复制为fs,后面这个filtered_source = fs，指向了筛选条件后的值


    for t in types: #遍历所有类型

        if normalize(t) in ['合计', '合计环比']: #跳过合计和合计环比
            continue
        normalized_t = normalize(t)
        mapped_values = []
        if mapping_for_sheet is not None:
            try:
                alias_norm_series = mapping_for_sheet['alias'].apply(lambda x: normalize(x)) # 标准化别名
                matches_mask = alias_norm_series == normalized_t # 比较别名是否与标准化后的类型名称相等
                matches = mapping_for_sheet[matches_mask]
            except Exception:
                matches = pd.DataFrame()
            if not matches.empty: #如果匹配到映射表，则将匹配到的源值添加到映射值列表中
                mapped_values = [str(x) for x in matches['source_value'].tolist()]
        if not mapped_values: #如果没有映射值，则将类型名称添加到映射值列表中
            mapped_values = [t]

        filtered_parts = []
        if group_col in filtered_source.columns: #如果有分组列，则根据映射值进行筛选
            for mv in mapped_values:
                mv_norm = normalize(mv)
                if mv_norm == "":
                    continue
                mask_eq = filtered_source[group_col].astype(str).apply(lambda x: normalize(x) == mv_norm)
                part = filtered_source[mask_eq] #根据映射值进行筛选
                if not part.empty: #如果筛选结果不为空，则将结果添加到filtered_parts列表中
                    filtered_parts.append(part)

        if filtered_parts: #如果筛选结果不为空，则将结果合并为一个DataFrame

            filtered = pd.concat(filtered_parts, axis=0, ignore_index=False) #pd.concat()函数用于合并多个DataFrame，axis=0表示按行合并，ignore_index=False表示保留原来的索引

        else:
            filtered = filtered_source.iloc[0:0].copy() #则创建一个与原数据结构相同但内容空的DataFrame，iloc[0:0]表示从filtered_source中选取第0行到第0行的所有列，copy()表示复制一份新的DataFrame



        sumd = {}
        for target in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量']: #遍历待填写子表模板中所有基础指标列

            src_col = inv_map.get(target)
            if src_col is None or src_col not in filtered.columns:
                sumd[target] = None
            else:
                try:
                    ser_num = _to_number_series_for_base_metrics(filtered[src_col]) # 将源列转换为数字序列
                    s = ser_num.sum(skipna=True) #求和
                    sumd[target] = float(s) if not pd.isna(s) else None #如果不是空值，则转换为浮点数
                except Exception:
                    sumd[target] = None

        for k in weekly_totals: #遍历所有6个基础指标，计算周总计指标

            if sumd.get(k) not in [None]:
                weekly_totals[k] += sumd[k] #将当前指标的值累加到周总计中


        # —— 使用别名映射 + 源值映射 来寻找历史行 —— #
        candidate_keys = []
        seen = set() # 集合避免重复

        def _add_key(k: str): #添加一个候选键，定义形参为字符串类型
            k = (k or "").strip().lower()
            if not k:
                return
            if k in seen:
                return
            seen.add(k)
            candidate_keys.append(k)

        # 1）本周当前名字本身
        _add_key(normalized_t)

        # 2）如果当前名字是“别名”，那它在历史里很可能是 source_value
        for sname in alias2src.get(normalized_t, []):
            _add_key(sname)

        # 3）如果当前名字是“源值”，那历史里也可能用 alias
        for aname in src2alias.get(normalized_t, []):
            _add_key(aname)

        # 按顺序尝试上述所有候选键，找到第一个命中的历史行
        hist_row = None
        for key in candidate_keys:
            if key in hist_lookup:
                hist_row = hist_lookup[key]
                break

        computed = compute_metrics(sumd, history_row=hist_row) 

        order_out_val = None #订单出库金额
        rate_val_str = '-' #费率数值格式化后的字符串，默认-
        rate_num = None #费率数值（未格式化）
        if sheet_name == '双周分品类':
            order_out_val = order_amount_map.get(normalize('合计' if t == '合计' else t))
            if (order_out_val not in [None, 0]) and (computed.get('花费') not in [None]):
                try:
                    rate_num = float(computed['花费']) / float(order_out_val)
                except:
                    rate_num = None
            rate_val_str = _fmt_rate_one_decimal(rate_num)

        row_out = {
            '类型': t,
            '展现量': sumd.get('展现量'),
            '点击量': sumd.get('点击量'),
            '花费': sumd.get('花费'),
            '成交量': sumd.get('成交量'),
            '成交额': sumd.get('成交额'),
            '加购量': sumd.get('加购量'),
            'ROI': computed.get('ROI'),
            '点击率': computed.get('点击率'),
            '点击转化率': computed.get('点击转化率'),
            'CPC': computed.get('CPC'),
            '加购成本': computed.get('加购成本'),
            '消耗环比': computed.get('消耗环比'),
            'ROI环比': computed.get('ROI环比')
        }
        if sheet_name == '双周分品类':
            row_out['订单出库金额'] = order_out_val
            row_out['__费率数值__'] = rate_num
        out_rows.append(row_out)  #这里out_rows列表中在前面主函数开头创建为空列表，为后续最终输出的表的数据，这里暂时将每个已经处理完的指标以字典形式存放


    total_row = {'类型': '合计'} #当前周度数据
    for k in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量']:
        total_row[k] = weekly_totals[k]
    total_computed = compute_metrics(weekly_totals, history_row=None)
    total_row['ROI'] = total_computed.get('ROI')
    total_row['点击率'] = total_computed.get('点击率')
    total_row['点击转化率'] = total_computed.get('点击转化率')
    total_row['CPC'] = total_computed.get('CPC')
    total_row['加购成本'] = total_computed.get('加购成本')
    total_row['消耗环比'] = None
    total_row['ROI环比'] = None

    def _parse_num_or_percent(x): # 解析数值或百分比
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return None
        s = str(x).strip().replace(',', '')
        if s in ('', '-'):
            return None
        if s.endswith('%'):
            try:
                return float(s[:-1]) / 100.0
            except:
                return None
        try:
            return float(s)
        except:
            return None

    total_prev = None #历史 合计行 数据
    if history_df is not None and len(history_df) > 0:
        history_type_col = None 
        for c in history_df.columns:
            if '类型' in c or c.lower() == 'type':
                history_type_col = c
                break
        if history_type_col is None:
            history_type_col = history_df.columns[1] if len(history_df.columns) > 1 else history_df.columns[0] #如果没有找到类型列，就取第一个列名

        hist_total_mask = history_df[history_type_col].astype(str).apply(lambda x: normalize(x) == '合计') # 从历史表的列名转字符串后找名为 合计 的列名，并将其规范化处理后返回具体的 合计 对应的行的所有值
        hist_total = history_df[hist_total_mask] #获取类型为 合计 的对应指标列的值
        if not hist_total.empty:
            hist_total = hist_total.iloc[0] # 取第一个合计行

            hist_totals_dict = {}
            fields = ['展现量','点击量','花费','成交量','成交额','加购量','订单出库金额','费率',
                      'ROI','点击率','点击转化率','CPC','加购成本'
                      ]  # 这些字段可能是数值或百分比
            for k in fields:
                if k in hist_total.index: #如果历史表中存在该字段，就将其转换为数值或百分比
                    hist_totals_dict[k] = _parse_num_or_percent(hist_total[k])
                else:
                    hist_totals_dict[k] = None
            total_prev = hist_totals_dict #历史 合计行 的目标指标数据

    total_cost = weekly_totals['花费'] if weekly_totals['花费'] != 0 else None # 总花费，这两个指标单独计算，便于后续计算消耗占比、成交占比这2个指标调用
    total_gmv = weekly_totals['成交额'] if weekly_totals['成交额'] != 0 else None # 总成交额

    final_rows = [] #最终的输出行，不含“合计”
    for r in out_rows:
        cost_ratio = None #消耗占比
        gmv_ratio = None #成交占比
        if total_cost is not None and r['花费'] is not None:
            cost_ratio = r['花费'] / total_cost 
        if total_gmv is not None and r['成交额'] is not None:
            gmv_ratio = r['成交额'] / total_gmv

        row_formatted = { # 格式化输出的每个指标，以字典存放
            '类型': r['类型'],
            '展现量': format_value('展现量', r['展现量']),
            '点击量': format_value('点击量', r['点击量']),
            '花费': format_value('花费', r['花费']),
            '成交量': format_value('成交量', r['成交量']),
            '成交额': format_value('成交额', r['成交额']),
            '加购量': format_value('加购量', r['加购量']),
            'ROI': format_value('ROI', r['ROI']),
            '点击率': format_value('点击率', r['点击率']),
            '点击转化率': format_value('点击转化率', r['点击转化率']),
            'CPC': format_value('CPC', r['CPC']),
            '加购成本': format_value('加购成本', r['加购成本']),
            '消耗占比': format_value('消耗占比', cost_ratio),
            '成交占比': format_value('成交占比', gmv_ratio),
            '消耗环比': format_value('消耗环比', r['消耗环比']),
            'ROI环比': format_value('ROI环比', r['ROI环比'])
        }
        if sheet_name == '双周分品类':
            row_formatted['订单出库金额'] = format_value('成交额', r.get('订单出库金额')) #将订单出库金额按成交额指标的格式输出


            rate_num = r.get('__费率数值__')
            row_formatted['费率'] = "-" if rate_num is None else f"{rate_num * 100:.1f}%"
        final_rows.append(row_formatted)

    total_row_formatted = { # 格式化输出“合计”行所有指标
        '类型': '合计',
        '展现量': format_value('展现量', total_row['展现量']),
        '点击量': format_value('点击量', total_row['点击量']),
        '花费': format_value('花费', total_row['花费']),
        '成交量': format_value('成交量', total_row['成交量']),
        '成交额': format_value('成交额', total_row['成交额']),
        '加购量': format_value('加购量', total_row['加购量']),
        'ROI': format_value('ROI', total_row['ROI']),
        '点击率': format_value('点击率', total_row['点击率']),
        '点击转化率': format_value('点击转化率', total_row['点击转化率']),
        'CPC': format_value('CPC', total_row['CPC']),
        '加购成本': format_value('加购成本', total_row['加购成本']),
        '消耗占比': '100%',
        '成交占比': '100%',
        '消耗环比': '-' if (
                total_prev is None or total_prev.get('花费') in [None, 0] or total_row.get('花费') in [None, 0])
        else format_value('合计环比', (total_row['花费'] / total_prev['花费'] - 1)),
        'ROI环比': '-' if (
                total_prev is None or total_prev.get('ROI') in [None, 0] or total_row.get('ROI') in [None, 0])
        else format_value('合计环比', (total_row['ROI'] / total_prev['ROI'] - 1))
    }

    order_out_total = None #“合计”行的订单出库总金额和费率单独计算
    if sheet_name == '双周分品类':
        order_out_total = order_amount_map.get(normalize('合计'))
        total_row_formatted['订单出库金额'] = format_value('成交额', order_out_total)
        if (order_out_total not in [None, 0]) and (total_row.get('花费') not in [None]):
            try:
                total_rate = float(total_row['花费']) / float(order_out_total)
            except:
                total_rate = None
        else:
            total_rate = None
        total_row_formatted['费率'] = "-" if total_rate is None else f"{total_rate * 100:.1f}%"

    final_rows.append(total_row_formatted)

    if total_prev is not None: #这里开始计算“合计环比”行对应的每个合计指标的环比值
        sum_cols = ['展现量', '点击量', '花费', '成交量', '成交额', '加购量', 'ROI', '点击率', '点击转化率', 'CPC', '加购成本']
        comp_row = {'类型': '合计环比'} #这里创建合计环比行，存放每个指标的环比值
        for col in sum_cols: #遍历逐个计算每个合计指标的环比值，计算合计环比行指标
            prev_val = total_prev.get(col) #历史数据的对应指标
            cur_val = total_row.get(col) #当前周度的对应指标
            if prev_val in [None, 0] or cur_val in [None]:
                comp_row[col] = '-'
            else:
                try:
                    val = (cur_val / prev_val - 1)
                    comp_row[col] = format_value('合计环比', val)
                except:
                    comp_row[col] = '-'
        for col in ['消耗占比', '成交占比', '消耗环比', 'ROI环比']:
            comp_row[col] = '-' 
        if sheet_name == '双周分品类': #双周分品类子表额外多2个合计环比指标处理
            try:
                prev_amt = total_prev.get('订单出库金额')
                cur_amt = order_out_total
                if prev_amt not in [None, 0] and cur_amt not in [None]:
                    comp_row['订单出库金额'] = format_value('合计环比', (float(cur_amt) / float(prev_amt) - 1))
                else:
                    comp_row['订单出库金额'] = '-'
            except:
                comp_row['订单出库金额'] = '-'
            try:
                if (order_out_total not in [None, 0]) and (total_row.get('花费') not in [None]):
                    cur_rate = float(total_row['花费']) / float(order_out_total)
                else:
                    cur_rate = None
                prev_rate = total_prev.get('费率')
                if prev_rate is None:
                    prev_cost = total_prev.get('花费')
                    prev_amt2 = total_prev.get('订单出库金额')
                    if prev_cost not in [None] and prev_amt2 not in [None, 0]:
                        prev_rate = float(prev_cost) / float(prev_amt2) #计算历史数据的费率，（针对历史数据的费率无法直接获得的情况）


                if cur_rate not in [None] and prev_rate not in [None, 0]: #如果当前周期的费率和历史数据的费率都不是None，那么直接计算当前周期的费率环比

                    comp_row['费率'] = format_value('合计环比', (cur_rate / prev_rate - 1))
                else:
                    comp_row['费率'] = '-'
            except:
                comp_row['费率'] = '-'
    else:
        comp_row = {'类型': '合计环比'}
        for col in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量', 'ROI', '点击率', '点击转化率', 'CPC',
                    '加购成本', '消耗占比', '成交占比', '消耗环比', 'ROI环比']:
            comp_row[col] = '-'
        if sheet_name == '双周分品类':
            comp_row['订单出库金额'] = '-'
            comp_row['费率'] = '-'

    columns_order = ['类型', '展现量', '点击量', '花费', '成交量', '成交额', '加购量', 'ROI', '点击率', '点击转化率',
                     'CPC', '加购成本', '消耗占比', '成交占比', '消耗环比', 'ROI环比']
    if sheet_name == '双周分品类':
        columns_order = ['类型', '展现量', '点击量', '花费', '成交量', '成交额', '加购量',
                         '订单出库金额', '费率',
                         'ROI', '点击率', '点击转化率', 'CPC', '加购成本',
                         '消耗占比', '成交占比', '消耗环比', 'ROI环比']

    final_rows.append(comp_row)
    result_df = pd.DataFrame(final_rows)[columns_order]
    return result_df


# =========================
# 展示/导出相关（格式化、着色、导出PNG）
# =========================
def _find_type_col(columns) -> str:
    candidates = ['类型', '型号', '资源位', '品类', '产品名称', 'type', 'resource', 'category', 'product_name']
    lower_map = {str(c).lower(): c for c in columns}
    for cand in candidates:
        if cand in columns:
            return cand
        lc = cand.lower()
        if lc in lower_map:
            return lower_map[lc]
    if len(columns) >= 2:
        second = str(columns[1]).strip()
        if any(k in second for k in ['资源位', '品类', '产品', '型号']):
            return columns[1]
    return None


def _format_existing_final_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    已有历史子表的格式化：
    - 兼容“原始小数”（0.123）和“已经带 % 的字符串”（"12.3%"）
    - 避免二次格式化导致的全部变成 "-"
    """
    if df is None or df.empty:
        return df

    def _as_str(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        return str(x).strip()

    def fmt_percent_int(x):
        """
        用于：消耗占比 / 成交占比 / 消耗环比 / ROI环比 / 合计环比（整型百分比）
        """
        s = _as_str(x)
        if not s:
            return "-"
        # 已经是 "12.3%" 这种形式：只做轻度归一，而不是再 *100
        if s.endswith("%"):
            try:
                v = float(s[:-1])
                return f"{int(round(v, 0))}%"
            except Exception:
                # 实在解析不了就原样返回，至少不变成 "-"
                return s
        try:
            v = float(s) * 100.0
            return f"{int(round(v, 0))}%"
        except Exception:
            return "-"

    def fmt_percent_1(x):
        """
        用于：点击率 / 点击转化率 / 费率（保留 1 位小数的百分比）
        """
        s = _as_str(x)
        if not s:
            return "-"
        if s.endswith("%"):
            try:
                v = float(s[:-1])
                return f"{v:.1f}%"
            except Exception:
                return s
        try:
            v = float(s) * 100.0
            return f"{v:.1f}%"
        except Exception:
            return "-"

    def fmt_float2(x):
        """
        ROI、CPC 这类小数（不带百分号），防御性兼容错误的百分比字符串
        """
        s = _as_str(x)
        if not s:
            return "-"
        # 理论上这里不会有 "%", 但防御一下
        if s.endswith("%"):
            try:
                v = float(s[:-1]) / 100.0
                return f"{v:.2f}"
            except Exception:
                return s
        try:
            v = float(s)
            return f"{v:.2f}"
        except Exception:
            return "-"

    def fmt_float1(x):
        """
        加购成本：一位小数
        """
        s = _as_str(x)
        if not s:
            return "-"
        if s.endswith("%"):
            try:
                v = float(s[:-1]) / 100.0
                return f"{v:.1f}"
            except Exception:
                return s
        try:
            v = float(s)
            return f"{v:.1f}"
        except Exception:
            return "-"

    def fmt_number(x):
        """
        基础计数/金额列：展现量/点击量/花费/成交量/成交额/加购量/订单出库金额 等
        - 如果本来就是 "12345" 或 12345.0 就正常格式化
        - 如果误传进来百分号字符串，就原样返回，避免 34% 被当 34 再二次格式化
        """
        s = _as_str(x)
        if not s:
            return "-"
        if s.endswith("%"):
            # 这里不再做任何数值转换，直接原样展示
            return s
        try:
            v = float(s)
            if abs(v - round(v)) < 1e-9:
                return str(int(round(v)))
            return f"{v:.2f}"
        except Exception:
            return s

    df2 = df.copy().where(pd.notna(df), None)

    type_col = _find_type_col(df2.columns)
    mask_total_cmp = pd.Series(False, index=df2.index)
    if type_col is not None and type_col in df2.columns:
        mask_total_cmp = df2[type_col].astype(str).str.strip().eq('合计环比')

    base_cols = [c for c in ['展现量','点击量','花费','成交量','成交额','加购量','订单出库金额'] if c in df2.columns]
    rate_cols = [c for c in ['点击率','点击转化率','费率'] if c in df2.columns]
    pct_cols  = [c for c in ['消耗占比','成交占比','消耗环比','ROI环比'] if c in df2.columns]
    val2_cols = [c for c in ['ROI','CPC'] if c in df2.columns]
    val1_cols = [c for c in ['加购成本'] if c in df2.columns]

    # 合计环比行：所有数值类列都按“整型百分比”显示
    if mask_total_cmp.any():
        for idx in df2.index[mask_total_cmp]:
            for col in df2.columns:
                if type_col is not None and col == type_col:
                    continue
                df2.at[idx, col] = fmt_percent_int(df2.at[idx, col])

    # 其他行：按列类型分别格式化
    mask_other = ~mask_total_cmp
    for idx in df2.index[mask_other]:
        for col in rate_cols:
            df2.at[idx, col] = fmt_percent_1(df2.at[idx, col])
        for col in pct_cols:
            df2.at[idx, col] = fmt_percent_int(df2.at[idx, col])
        for col in val2_cols:
            df2.at[idx, col] = fmt_float2(df2.at[idx, col])
        for col in val1_cols:
            df2.at[idx, col] = fmt_float1(df2.at[idx, col])
        for col in base_cols:
            df2.at[idx, col] = fmt_number(df2.at[idx, col])

    return df2.fillna('-')


def _filter_week_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return df
    df = df.copy()
    type_col = _find_type_col(df.columns)
    if type_col is None or type_col not in df.columns:
        return df
    base_cols = [c for c in ['展现量', '点击量', '花费', '成交量', '成交额', '加购量', '订单出库金额'] if c in df.columns]
    rate_cols = [c for c in ['点击率', '点击转化率', '费率'] if c in df.columns]
    share_cols = [c for c in ['消耗占比', '成交占比'] if c in df.columns]
    value_cols = [c for c in ['ROI', 'CPC', '加购成本'] if c in df.columns]

    def _to_float_or_none2(x):
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return None
        s = str(x).strip()
        if s in ('-', ''):
            return None
        if s.endswith('%'):
            try:
                return float(s[:-1])
            except:
                return None
        try:
            return float(s)
        except:
            return None

    def _row_has_meaningful_value(row):
        for c in base_cols:
            v = _to_float_or_none2(row.get(c))
            if v is not None and v > 0:
                return True
        for c in (share_cols + rate_cols):
            v = _to_float_or_none2(row.get(c))
            if v is not None and v > 0:
                return True
        for c in value_cols:
            v = _to_float_or_none2(row.get(c))
            if v is not None and v != 0.0:
                return True
        return False

    keep_mask = df[type_col].astype(str).str.strip().isin(['合计', '合计环比'])
    has_value_mask = df.apply(_row_has_meaningful_value, axis=1)
    final_mask = keep_mask | has_value_mask
    return df[final_mask].reset_index(drop=True)


def _style_two_rows_two_cols(df: pd.DataFrame):
    PINK = '#ffd6e7'  # >= 0%
    GREEN = '#d7f5d1'  # < 0%
    type_col = None
    for cand in ['类型', '型号', 'type']:
        if cand in df.columns:
            type_col = cand
            break
    if type_col is None:
        return pd.DataFrame('', index=df.index, columns=df.columns)

    def pct_to_float(s):
        if s is None:
            return None
        ss = str(s).strip()
        if ss == '-' or ss == '': return None
        if ss.endswith('%'):
            try:
                return float(ss[:-1])
            except:
                return None
        return None

    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    mask_total_cmp = df[type_col].astype(str).str.strip().eq('合计环比')
    for i in df.index[mask_total_cmp]:
        for col in df.columns:
            p = pct_to_float(df.at[i, col])
            if p is None: continue
            styles.at[i, col] = f'background-color: {PINK if p >= 0 else GREEN};'

    target_cols = [c for c in ['消耗环比', 'ROI环比'] if c in df.columns]
    mask_other = ~mask_total_cmp
    for col in target_cols:
        for i in df.index[mask_other]:
            p = pct_to_float(df.at[i, col])
            if p is None: continue
            styles.at[i, col] = f'background-color: {PINK if p >= 0 else GREEN};'

    mask_total_or_cmp = df[type_col].astype(str).str.strip().isin(['合计', '合计环比'])
    for i in df.index[mask_total_or_cmp]:
        styles.at[i, type_col] = (styles.at[i, type_col] + 'font-weight:700;').strip()
    return styles


# === 新增：统一为“本周结果/最新历史表”增加『日期』首列 ===
def _add_date_col_for_current_week(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """
    在导出的本周结果 / 写入最新历史表时，为每个子表统一加上第 1 列「日期」：
    - 若原表无「日期」列：插入到首列，整列填 label
    - 若原表已有「日期」列：保留已有非空值，空值填 label，并把「日期」移动到首列
    """
    if df is None or df.empty:
        return df
    df2 = df.copy()
    if '日期' not in df2.columns:
        df2.insert(0, '日期', label)
    else:
        df2['日期'] = df2['日期'].apply(
            lambda x: label if (x is None or str(x).strip() in ('', '-')) else x
        )
        cols = list(df2.columns)
        cols.remove('日期')
        df2 = df2[['日期'] + cols]
    return df2


def _df_to_png_bytes(
        df: pd.DataFrame,
        dpi: int = 200,
        font_px: int = 8,
        cell_w_px: int = 140,
        row_px: int = 36,
        max_w_px: int = 4000,
        center: bool = True,
        pad_frac: float = 0.02,
):
    # —— 延迟导入 matplotlib，并在此处设置中文字体兜底 —— #
    import matplotlib
    import matplotlib.pyplot as plt
    matplotlib.rcParams['font.sans-serif'] = ['PingFang SC', 'Microsoft YaHei', 'Arial Unicode MS', 'Heiti SC', 'SimHei',
                                              'DejaVu Sans']
    matplotlib.rcParams['axes.unicode_minus'] = False

    import io as _io
    if df is None or df.empty:
        return _io.BytesIO()

    styles = _style_two_rows_two_cols(df)
    nrows, ncols = df.shape

    width_px = min(max_w_px, max(800, ncols * cell_w_px))
    height_px = max(400, (nrows + 1) * row_px + 40)
    fig_w_in = width_px / dpi
    fig_h_in = height_px / dpi

    fig = plt.figure(figsize=(fig_w_in, fig_h_in), dpi=dpi)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)

    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis('off')

    table = ax.table(
        cellText=df.astype(str).values,
        colLabels=list(df.columns),
        loc='upper left',
        cellLoc='center',
        colLoc='center',
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font_px)
    try:
        table.auto_set_column_width(list(range(ncols)))
    except Exception:
        pass

    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#e5e7eb')
        if r == 0:
            cell.set_facecolor('#f2f2f2')
            cell.get_text().set_weight('bold')

    try:
        sample_h_axes = table[(1, 0)].get_height()
        current_row_px = sample_h_axes * height_px
        y_scale = max(1.0, row_px / max(1.0, current_row_px))
    except Exception:
        y_scale = max(1.2, font_px / 9.0)
    table.scale(1.0, y_scale)

    if styles is not None and not styles.empty:
        for i in range(nrows):
            for j in range(ncols):
                cell = table.get_celld().get((i + 1, j))
                if cell is None:
                    continue
                sty = styles.iat[i, j]
                if 'background-color:' in sty:
                    try:
                        color = sty.split('background-color:')[1].split(';')[0].strip()
                        cell.set_facecolor(color)
                    except Exception:
                        pass
                if 'font-weight:700' in sty or 'font-weight:bold' in sty:
                    cell.get_text().set_weight('bold')

    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    bbox = table.get_window_extent(renderer=renderer)
    table_px_width = bbox.width
    fig_px_width = fig.get_size_inches()[0] * dpi

    if center:
        desired_left_px = (fig_px_width - table_px_width) / 2.0
        dx_px = desired_left_px - bbox.x0
        dx_axes = dx_px / fig_px_width
        for (r, c), cell in table.get_celld().items():
            cell.set_x(cell.get_x() + dx_axes)
        fig.canvas.draw()

    buf = _io.BytesIO()
    fig.savefig(buf, format='png', facecolor='white', transparent=False, pad_inches=0.05)
    plt.close(fig)
    buf.seek(0)
    return buf


# =========================
# SessionState：控制“只在按钮点击时保存历史”
# =========================
def _init_session():
    if "processed_sheets" not in st.session_state:
        st.session_state.processed_sheets = None
    if "last_history_save_token" not in st.session_state:
        st.session_state.last_history_save_token = None
    if "current_history_path" not in st.session_state:
        st.session_state.current_history_path = None


# =========================
# 主入口
# =========================
def main():
    st.set_page_config(page_title="双周回流数据自动填表器", layout="wide")
    _init_session()

    st.title("周度数据自动填表器（本地）")
    st.markdown(
        f"- 数据目录：`{DATA_DIR}`\n"
        "- 请在 data/ 中手动替换/更新：`类型具体名称别名映射表.xlsx`、`历史周度数据表.xlsx`（或带日期的历史文件）。\n"
        "- 本工具仅在你 **点击确认按钮** 后，才会把本周结果写入 data/ 作为最新历史。\n"
    )

    mapping_df = load_mapping_df()

    # 历史：优先读取 data/ 下（也允许用户本次手动上传覆盖一次）
    history_uploaded = st.file_uploader(
        "（可选）本次临时覆盖：上传【上周历史周度数据表.xlsx】（仅本次处理使用；不会自动写回 data/）",
        type=['xlsx', 'xls'],
        accept_multiple_files=False
    )

    if history_uploaded is not None:
        try:
            history_all = pd.read_excel(history_uploaded, sheet_name=None, engine="openpyxl")
            st.success("已加载你上传的历史表（仅本次会话有效，不会写回 data/）。")
            st.session_state.current_history_path = "(本次临时上传)"
        except Exception as e:
            st.error(f"无法读取你上传的历史表：{e}")
            history_all, used_path = load_history_sheets()
            st.session_state.current_history_path = used_path
    else:
        history_all, used_path = load_history_sheets()
        st.session_state.current_history_path = used_path

    # 动态映射（持久化）
    st.markdown("### 🔧 新增/管理 品类子表映射（一次配置，长期生效，保存在 data/）")
    saved_model_map, saved_cat_map = load_dynamic_maps()
    runtime_category_map = dict(BASE_CATEGORY_MAP); runtime_category_map.update(saved_model_map)
    runtime_category_sheet_to_category = dict(BASE_CATEGORY_SHEET_TO_CATEGORY); runtime_category_sheet_to_category.update(saved_cat_map)

    with st.expander("管理新增子表（型号类 & 类别类）", expanded=False):
        st.caption("每行一个映射，英文逗号分隔：左=子表名，右=数据源中“品类”名称。示例：`眼镜型号,眼镜`；`眼镜,眼镜`")

        def _pairs_to_text(d: dict):
            return "\n".join([f"{k},{v}" for k, v in d.items()])

        col1, col2 = st.columns(2)
        with col1:
            model_map_text = st.text_area("型号类子表映射（子表名,品类名）",
                                          value=_pairs_to_text(saved_model_map),
                                          height=160)
        with col2:
            category_map_text = st.text_area("类别子表映射（子表名,品类名）",
                                             value=_pairs_to_text(saved_cat_map),
                                             height=160)

        def _parse_pairs(text):
            result = {}
            for i, line in enumerate(text.splitlines(), start=1):
                s = line.strip()
                if not s: continue
                if ',' not in s:
                    st.warning(f"第 {i} 行缺少英文逗号，已忽略：{s}")
                    continue
                left, right = s.split(',', 1)
                left, right = left.strip(), right.strip()
                if left and right:
                    result[left] = right
            return result

        b1, b2, _ = st.columns(3)
        with b1:
            if st.button("保存映射（写入 data/）", use_container_width=True):
                new_model_map = _parse_pairs(model_map_text)
                new_cat_map = _parse_pairs(category_map_text)
                ok, msg = save_dynamic_maps(new_model_map, new_cat_map)
                if ok:
                    st.success("已保存，下次默认生效。")
                else:
                    st.error(f"保存失败：{msg}")
        with b2:
            if st.button("清空已保存映射", use_container_width=True):
                ok, msg = save_dynamic_maps({}, {})
                if ok:
                    st.success("已清空；将回退到内置默认映射。")
                else:
                    st.error(f"清空失败：{msg}")

    with st.expander("查看当前运行时映射", expanded=False):
        st.write("**型号类子表 → 品类**")
        st.dataframe(pd.DataFrame(sorted(runtime_category_map.items()), columns=["型号类子表", "品类"]),
                     use_container_width=True)
        st.write("**类别子表 → 品类**")
        st.dataframe(pd.DataFrame(sorted(runtime_category_sheet_to_category.items()), columns=["类别子表", "品类"]),
                     use_container_width=True)

    # 仅展示状态：当前使用的历史文件路径
    st.info(f"当前历史参照文件：{st.session_state.current_history_path or '(未找到，无法计算环比)'}")

    uploaded_file = st.file_uploader(
        "上传 最新周度数据表.xlsx（包含“最新周度数据源”与各子表模板）",
        type=['xlsx', 'xls'],
        accept_multiple_files=False
    )

    validation_report_df = None
    suggest_model_pairs = []
    suggest_cat_pairs = []

    if uploaded_file is not None:
        # ★ 每次处理一个“最新周度数据表”之前，如果没有手动上传历史表，
        # 就强制重新从 data 目录读取最新历史表（包括你刚刚确认保存后的那一版）
        if history_uploaded is None:
            history_all, _ = load_history_sheets()

        try:
            latest_all = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
        except Exception as e:
            st.error(f"无法读取上传文件：{e}")
            st.stop()

        sheet_names = list(latest_all.keys())
        st.write("检测到上传文件的子表：", sheet_names)

        # 定位数据源
        source_sheet_name = None
        for s in sheet_names:
            if '数据源' in s or ('data source' in s.lower()) or '最新周度数据源' in s:
                source_sheet_name = s
                break
        if source_sheet_name is None:
            source_sheet_name = sheet_names[0]
            st.warning(f"未检测到明确命名为“最新周度数据源”的表，默认使用第一个子表：{source_sheet_name} 作为数据源。")

        source_df = latest_all[source_sheet_name].copy()
        source_df.columns = [c.strip() for c in source_df.columns]

        # 仅处理上传里出现的子表（删除的子表自然跳过）
        template_sheets = [s for s in sheet_names if s != source_sheet_name]
        st.write("将处理的子表（模板）：", template_sheets)

        # 校验器
        st.markdown("#### ✅ 子表命名与映射校验报告")
        validation_report_df, suggest_model_pairs, suggest_cat_pairs, has_error = validate_upload(
            template_sheets, source_df, runtime_category_map, runtime_category_sheet_to_category
        )
        ok_cnt = (validation_report_df['状态'] == 'OK').sum()
        warn_cnt = (validation_report_df['状态'] == 'WARN').sum()
        err_cnt = (validation_report_df['状态'] == 'ERROR').sum()
        st.write(f"校验结果：OK {ok_cnt} · WARN {warn_cnt} · ERROR {err_cnt}")
        st.dataframe(validation_report_df, use_container_width=True)

        # 一键采纳建议（若有）
        if suggest_model_pairs or suggest_cat_pairs:
            with st.expander("可采纳的建议映射（点击展开）", expanded=False):
                if suggest_model_pairs:
                    st.write("**建议添加到 型号类 映射：**")
                    st.dataframe(pd.DataFrame(suggest_model_pairs, columns=["子表名", "建议品类"]))
                if suggest_cat_pairs:
                    st.write("**建议添加到 类别类 映射：**")
                    st.dataframe(pd.DataFrame(suggest_cat_pairs, columns=["子表名", "建议品类"]))
                if st.button("一键采纳并保存（写入 data/）", use_container_width=True):
                    new_model_map = dict(saved_model_map); new_model_map.update({k: v for k, v in suggest_model_pairs})
                    new_cat_map = dict(saved_cat_map); new_cat_map.update({k: v for k, v in suggest_cat_pairs})
                    ok, msg = save_dynamic_maps(new_model_map, new_cat_map)
                    if ok:
                        st.success("已保存建议映射。本次运行已生效，下次启动也会默认使用。")
                        runtime_category_map.update({k: v for k, v in suggest_model_pairs})
                        runtime_category_sheet_to_category.update({k: v for k, v in suggest_cat_pairs})
                    else:
                        st.error(f"保存失败：{msg}")

        # 自动识别“类型匹配字段”
        src_group_col = candidate_source_key(source_df)
        st.write("用于匹配类型的源表字段（自动识别）：", src_group_col)

        processed_sheets = {}
        default_label = datetime.datetime.now().strftime("本周(%Y-%m-%d)")
        current_week_label = st.text_input("（可选）本周“日期”列标签", value=default_label)

        # 逐子表处理
        for sheet in template_sheets:
            template_df = latest_all[sheet].copy()
            st.info(f"正在处理子表：{sheet}")

            mapping_for_sheet = None
            if mapping_df is not None:
                mapping_all_for_sheet = find_mapping_for_sheet(mapping_df, sheet)
                if mapping_all_for_sheet is not None:
                    mapping_for_sheet = mapping_all_for_sheet[
                        mapping_all_for_sheet['sheet_norm'] == normalize(sheet)]
                    if mapping_for_sheet.empty:
                        mapping_for_sheet = None

            history_df = history_all.get(sheet) if sheet in history_all else None

            try:
                processed_df = process_subtable(
                    template_df=template_df,
                    source_df=source_df,
                    history_df=history_df,
                    mapping_for_sheet=mapping_for_sheet,
                    src_group_col=src_group_col,
                    sheet_name=sheet,
                    runtime_category_map=runtime_category_map,
                    runtime_category_sheet_to_category=runtime_category_sheet_to_category
                )
                processed_sheets[sheet] = processed_df
                st.success(f"{sheet} 本周完成（行数 {len(processed_df)}）")
                st.dataframe(processed_df.head(10), use_container_width=True)
            except Exception as e:
                st.error(f"{sheet} 本周处理失败：{e}")
                processed_sheets[sheet] = None

        # 导出“本周已处理完成的周度表（Excel）”
        out_bytes = io.BytesIO()
        with pd.ExcelWriter(out_bytes, engine='openpyxl') as writer:
            source_df.to_excel(writer, sheet_name=source_sheet_name, index=False)
            for k, v in processed_sheets.items():
                if v is not None:
                    # === 这里：导出时为本周子表加「日期」首列 ===
                    v_export = _add_date_col_for_current_week(v, current_week_label)
                    v_export.to_excel(writer, sheet_name=k, index=False)
            if validation_report_df is not None:
                validation_report_df.to_excel(writer, sheet_name="校验报告", index=False)
        out_bytes.seek(0)
        st.download_button(
            "⬇️ 下载已处理周度数据表（Excel，含“校验报告”& 本周日期列）",
            data=out_bytes.getvalue(),
            file_name="已处理周度数据表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 双周合并视图（上周 + 本周）
        st.markdown("## 子表双周合并视图（单表展示上周 + 本周）")
        st.caption("预览支持右上角『全屏』查看；每个子表下方提供 PNG/Excel 下载（PNG 含配色）。")

        def _ensure_date_col(df, label):
            if df is None:
                return df
            df2 = df.copy()
            if '日期' not in df2.columns:
                df2.insert(0, '日期', label)
            else:
                df2['日期'] = df2['日期'].apply(
                    lambda x: label if (x is None or str(x).strip() in ('', '-')) else x)
            return df2

        def _normalize_type_col(df):
            if df is None:
                return df
            df2 = df.copy()
            if '类型' not in df2.columns and '型号' in df2.columns:
                df2.rename(columns={'型号': '类型'}, inplace=True)
            return df2

        cols_pref = [
            '日期', '类型', '展现量', '点击量', '花费', '成交量', '成交额', '加购量',
            '订单出库金额', '费率',
            'ROI', '点击率', '点击转化率', 'CPC', '加购成本',
            '消耗占比', '成交占比', '消耗环比', 'ROI环比'
        ]

        def _align_cols(df):
            if df is None:
                return df
            ordered = [c for c in cols_pref if c in df.columns]
            tail = [c for c in df.columns if c not in ordered]
            return df[ordered + tail]

        prev_label = "上周" if st.session_state.current_history_path else "上周"

        for sheet in template_sheets:
            st.subheader(f"📘 {sheet}")
            prev_df_raw = history_all.get(sheet)
            prev_df_fmt = _format_existing_final_table(prev_df_raw) if prev_df_raw is not None else None
            cur_df_raw = processed_sheets.get(sheet)

            prev_df = _filter_week_df(prev_df_fmt) if prev_df_fmt is not None else None
            cur_df = _filter_week_df(cur_df_raw) if cur_df_raw is not None else None

            prev_df = _ensure_date_col(prev_df, prev_label)
            cur_df = _ensure_date_col(cur_df, current_week_label)

            prev_df = _normalize_type_col(prev_df)
            cur_df = _normalize_type_col(cur_df)

            prev_df = _align_cols(prev_df)
            cur_df = _align_cols(cur_df)

            if (prev_df is None or prev_df.empty) and (cur_df is None or cur_df.empty):
                st.info("（上周/本周均无数据）")
                st.markdown("---")
                continue

            combined_parts = []
            if prev_df is not None and not prev_df.empty:
                combined_parts.append(prev_df)
            if cur_df is not None and not cur_df.empty:
                combined_parts.append(cur_df)
            combined_df = pd.concat(combined_parts, ignore_index=True)

            st.dataframe(combined_df, use_container_width=True)

            png_buf = _df_to_png_bytes(combined_df)
            st.download_button(
                label=f"⬇️ 下载《{sheet}》双周对比（PNG图片）",
                data=png_buf.getvalue(),
                file_name=f"{sheet}_双周对比.png",
                mime="image/png",
                key=f"img_{sheet}"
            )

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                combined_df.to_excel(w, sheet_name=f"{sheet}_双周对比", index=False)
            buf.seek(0)
            st.download_button(
                label=f"⬇️ 下载《{sheet}》双周对比（Excel）",
                data=buf.getvalue(),
                file_name=f"{sheet}_双周对比.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{sheet}"
            )

            st.markdown("---")

        # ========== 仅在用户确认后，才把本周结果写入 data/ 作为“最新历史周度数据表” ==========
        st.session_state.processed_sheets = processed_sheets  # 为按钮回调保留
        st.markdown("### ✅ 确认无误后更新历史表")
        st.caption("说明：点击下方按钮后，程序会把本周已处理结果写入 data/，包括：覆盖 `历史周度数据表.xlsx`，并额外保存一份带日期归档。")
        col_save1, col_save2 = st.columns([1, 1])
        with col_save1:
            if st.button("确认无误，更新为最新历史周度表（写入 data/）", type="primary", use_container_width=True):
                token = f"{datetime.datetime.now().timestamp()}"
                if st.session_state.last_history_save_token == token:
                    st.info("本次保存已处理。")
                else:
                    st.session_state.last_history_save_token = token
                    try:
                        # 覆盖当前最新历史：子表中写入「日期」列 = 当前周标签
                        with pd.ExcelWriter(USER_HISTORY_FILE, engine="openpyxl") as writer:
                            for s, df in processed_sheets.items():
                                if df is not None:
                                    df_hist = _add_date_col_for_current_week(df, current_week_label)
                                    df_hist.to_excel(writer, sheet_name=s, index=False)
                        # 额外按日期归档一份
                        stamp = datetime.datetime.now().strftime("%Y%m%d")
                        archive = DATA_DIR / f"历史周度数据表_{stamp}.xlsx"
                        with pd.ExcelWriter(archive, engine="openpyxl") as writer:
                            for s, df in processed_sheets.items():
                                if df is not None:
                                    df_hist = _add_date_col_for_current_week(df, current_week_label)
                                    df_hist.to_excel(writer, sheet_name=s, index=False)
                        st.success(f"已写入：{USER_HISTORY_FILE.name}，并归档为：{archive.name}")
                    except Exception as e:
                        st.error(f"写入 data/ 失败：{e}")
        with col_save2:
            st.caption(f"当前 data/ 下历史参照：{(st.session_state.current_history_path or '无')}")

    st.markdown("---")
    st.markdown(
        "**使用提示**：\n"
        "- 请将分发包中 `data/` 文件夹与 App/EXE 放在一起，或仅提供 App/EXE，由程序首次自动创建 data/。\n"
        "- 用户只需替换 `data/` 里的表格即可完成“后台替换”。\n"
        "- 若本周数据确认无误，再点击『确认无误，更新为最新历史周度表』，否则不会写入 data/（避免误保存）。\n"
    )
