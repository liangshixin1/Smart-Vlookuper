# -*- coding: utf-8 -*-
"""
SMART-VLOOKUPER - Excel 字段匹配与 AI 自动化工具 (PyQt6)

主要功能：
- 模糊字段匹配与 COM 批量写入，自动保留原有单元格格式
- 内置 AI 助手：上传表格并描述需求后，实时预览流式生成的 Python 代码并在沙箱中执行
- 失败重试与可取消的进度提示，确保最终产出可正常打开的 Excel 文件

依赖：pip install pyqt6 pandas openpyxl pywin32 thefuzz
"""

import sys, os, re, warnings, json, subprocess, tempfile, threading, random, string
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QGridLayout, QGroupBox, QLabel, QLineEdit, QPushButton, QComboBox,
    QSpinBox, QHBoxLayout, QVBoxLayout, QListWidget, QListWidgetItem, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QStyledItemDelegate, QRadioButton,
    QTabWidget, QDialog, QPlainTextEdit, QTextBrowser,
    QDialogButtonBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QAction, QTextCursor
import gc
from thefuzz import fuzz
from copy import copy
from typing import List

# —— 静默 openpyxl 的数据验证扩展警告 ——
warnings.filterwarnings(
    "ignore",
    message=".*Data Validation extension is not supported and will be removed.*",
    category=UserWarning,
    module="openpyxl"
)

# ===================== 全局设置管理 =====================

class AppSettings:
    """简单的JSON配置管理器，用于持久化用户偏好"""

    DEFAULTS = {
        "ai_api_key": "",
        "theme": "dark",
        "engine_mode": "auto",
        "last_ai_export_path": ""
    }

    def __init__(self):
        self.config_dir = Path.home() / ".smart_vlookuper"
        self.path = self.config_dir / "settings.json"
        self.data = self.DEFAULTS.copy()
        self.load()

    def load(self):
        if self.path.exists():
            try:
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    self.data.update({k: loaded.get(k, v) for k, v in self.DEFAULTS.items()})
            except Exception:
                # 读取失败时保留默认配置
                pass

    def save(self):
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            # 配置写入失败不阻塞主流程
            pass

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value

    def update(self, **kwargs):
        self.data.update(kwargs)
        self.save()


PINK_ACCENT = "#F083A2"
PINK_ACCENT_HOVER = "#D96B8D"


EXPORT_FORMATS = {
    "excel": {
        "label": "Excel 工作簿 (.xlsx)",
        "extensions": [".xlsx", ".xlsm", ".xls"],
        "default_extension": ".xlsx",
        "dialog_filter": "Excel Files (*.xlsx *.xlsm *.xls)",
        "description": "Excel 工作簿"
    },
    "word": {
        "label": "Word 文档 (.docx)",
        "extensions": [".docx"],
        "default_extension": ".docx",
        "dialog_filter": "Word Documents (*.docx)",
        "description": "Word 文档"
    },
    "text": {
        "label": "文本文件 (.txt)",
        "extensions": [".txt", ".csv"],
        "default_extension": ".txt",
        "dialog_filter": "Text Files (*.txt *.csv)",
        "description": "文本文件"
    }
}

PRIVACY_STATEMENT_TEXT = """最后更新日期：2025年9月26日

尊敬的用户：

感谢您使用本系统的AI助手功能。我们深知数据安全与隐私保护的重要性，尤其是您可能处理的个人敏感信息。我们遵循 “数据最小化”与 “隐私优先” 的核心原则。我们的AI助手不会将您的原始数据发送至任何第三方服务器进行处理。所有敏感信息的保护均在您的本地计算机上完成。当您上传Excel文件前，系统会在本地自动扫描表格的表头（列名称），并根据内置的规则库识别可能包含个人敏感信息的列。对于识别出的敏感列，系统会立即在您的电脑内存中**将该列的所有数据内容替换为计算机随机生成的、符合原始数据格式的虚假样本数据。**之后，系统将使用这份**已经过脱敏处理的、不包含任何真实个人信息的数据样本**来生成提供给AI模型的提示词。您原始的、真实的Excel文件数据在任何情况下都绝不会离开您的本地环境。所有生成的代码均在严格的隔离环境中运行，无法访问您的网络或计算机上的其他无关文件。

您有权知晓上述处理流程。使用AI功能即代表您同意此数据处理方式。如果您不希望进行任何形式的数据处理，请不要使用AI助手功能，软件的核心匹配功能仍可完全离线使用。在AI代码预览窗口中，您可以完整审查即将被执行的代码逻辑。

如果您对本声明或我们的数据安全有任何疑问，请访问我们的Github项目地址，并提交Issue"""


DARK_STYLESHEET = """
    QMainWindow { background: #0f172a; color: #f8fafc; }
    QMenuBar { background: #0f172a; color: #f8fafc; }
    QMenuBar::item { background: transparent; color: #f8fafc; padding: 4px 12px; }
    QMenuBar::item:selected { background: #1e293b; color: #f8fafc; }
    QMenu { background: #111827; color: #f8fafc; border: 1px solid #1f2937; }
    QMenu::item:selected { background: #1f2937; }
    QDialog, QMessageBox { background: #0f172a; color: #f8fafc; }
    QLabel { color: #f1f5f9; font-weight: 600; }
    QGroupBox { border: 1px solid #1f2937; border-radius: 10px; margin-top: 10px; padding: 10px; color: #f8fafc; font-weight: 600; }
    QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 2px 6px; }
    QLineEdit, QComboBox, QSpinBox, QListWidget, QTableWidget { background: #111827; color: #f8fafc; border: 1px solid #374151; border-radius: 8px; padding: 6px; }
    QComboBox::drop-down { border: none; }
    QPushButton { background: %(accent)s; color: white; border: none; border-radius: 8px; padding: 8px 12px; font-weight: 600; }
    QPushButton:hover { background: %(accent_hover)s; }
    QPushButton:disabled { background: #334155; color: #9ca3af; }
    QRadioButton { color: #f8fafc; font-weight: normal; }
    QHeaderView::section { background: #0b1220; color: #f1f5f9; padding: 6px; border: none; }
    QTableWidget { gridline-color: #374151; }
    QTableWidget::item { padding-left: 5px; }
    QTabWidget::pane { border: 1px solid #1f2937; border-radius: 10px; }
    QTabBar::tab { background: #1e293b; color: #f1f5f9; padding: 6px 12px; margin: 2px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
    QTabBar::tab:selected { background: %(accent)s; color: white; }
""" % {"accent": PINK_ACCENT, "accent_hover": PINK_ACCENT_HOVER}


LIGHT_STYLESHEET = """
    QMainWindow { background: #f8fafc; color: #0f172a; }
    QMenuBar { background: #f8fafc; color: #0f172a; }
    QMenuBar::item { background: transparent; color: #0f172a; padding: 4px 12px; }
    QMenuBar::item:selected { background: #e2e8f0; color: #0f172a; }
    QMenu { background: #ffffff; color: #0f172a; border: 1px solid #cbd5e1; }
    QMenu::item:selected { background: #e2e8f0; }
    QDialog, QMessageBox { background: #f8fafc; color: #0f172a; }
    QLabel { color: #0f172a; font-weight: 600; }
    QGroupBox { border: 1px solid #cbd5e1; border-radius: 10px; margin-top: 10px; padding: 10px; color: #0f172a; font-weight: 600; background: #ffffff; }
    QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 2px 6px; }
    QLineEdit, QComboBox, QSpinBox, QListWidget, QTableWidget { background: #ffffff; color: #111827; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px; }
    QComboBox::drop-down { border: none; }
    QPushButton { background: %(accent)s; color: white; border: none; border-radius: 8px; padding: 8px 12px; font-weight: 600; }
    QPushButton:hover { background: %(accent_hover)s; }
    QPushButton:disabled { background: #cbd5e1; color: #94a3b8; }
    QRadioButton { color: #0f172a; font-weight: normal; }
    QHeaderView::section { background: #e2e8f0; color: #1e293b; padding: 6px; border: none; }
    QTableWidget { gridline-color: #cbd5e1; }
    QTableWidget::item { padding-left: 5px; }
    QTabWidget::pane { border: 1px solid #cbd5e1; border-radius: 10px; background: #ffffff; }
    QTabBar::tab { background: #e2e8f0; color: #1e293b; padding: 6px 12px; margin: 2px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
    QTabBar::tab:selected { background: %(accent)s; color: white; }
""" % {"accent": PINK_ACCENT, "accent_hover": PINK_ACCENT_HOVER}

# ===================== 基础工具与模糊匹配 =====================

def norm_str(x):
    """规范化字符串，去除首尾空格"""
    return str(x).strip()

def dedup_columns(names):
    """为重复的列名添加后缀以保证唯一性"""
    seen, out = {}, []
    for n in names:
        base = n
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base} ({seen[base]})")
    return out


SENSITIVE_KEYWORDS = {
    "name": ["姓名", "名字", "name", "联系人", "客户姓名", "学生姓名"],
    "id": ["身份证", "身份证号", "身份证号码", "证件号", "证件号码", "id card", "idcard"],
    "phone": ["手机号", "手机", "电话", "联系电话", "phone", "mobile", "联系方式"],
    "address": ["地址", "住址", "家庭地址", "办公地址", "通讯地址", "联系地址", "address"]
}

FAKE_SURNAMES = list("赵钱孙李周吴郑王冯陈楮卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计成戴谈宋茅庞熊纪舒屈项祝董梁")
FAKE_GIVEN_SINGLE = ["明", "芳", "华", "敏", "伟", "磊", "娜", "强", "霞", "杰", "静", "勇", "婷", "超", "燕", "波", "丽", "凯", "睿", "佳"]
FAKE_GIVEN_DOUBLE = ["子涵", "梓萱", "浩然", "欣怡", "宇轩", "思远", "雅静", "晨曦", "博文", "诗涵", "靖雯", "曜文", "逸晨", "沐阳", "琪瑛", "思琪"]
FAKE_PROVINCES = ["京", "沪", "粤", "浙", "苏", "鲁", "川", "渝", "辽", "湘", "闽", "皖", "鄂", "津"]
FAKE_ADDRESS_SUFFIX = ["幸福路", "创新大道", "人民中路", "解放东街", "和平里", "阳光巷", "滨江道", "文化路", "建业街", "望江路", "星海路"]


def classify_sensitive_column(column_name: str):
    lowered = norm_str(column_name).lower()
    for category, keywords in SENSITIVE_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in lowered:
                return category
    return None


def _fake_chinese_name(length_hint: int = 2) -> str:
    surname = random.choice(FAKE_SURNAMES)
    if length_hint <= 2:
        given = random.choice(FAKE_GIVEN_SINGLE)
    else:
        given = random.choice(FAKE_GIVEN_DOUBLE)
    return surname + given


def _fake_id_like(text: str) -> str:
    stripped = re.sub(r"[^0-9Xx]", "", text)
    length = len(stripped) if stripped else max(len(text), 18)
    length = max(length, 6)
    body = "".join(random.choices(string.digits, k=max(length - 1, 1)))
    if length >= 18:
        return body[: length - 1] + random.choice(string.digits + "X")
    return body + random.choice(string.digits)


def _fake_phone_like(text: str) -> str:
    digits_only = re.sub(r"\D", "", text)
    length = len(digits_only) if digits_only else 11
    length = max(length, 8)
    if length >= 11:
        prefix = random.choice(["13", "15", "17", "18", "19"])
        remainder = "".join(random.choices(string.digits, k=length - len(prefix)))
        return (prefix + remainder)[:length]
    return "".join(random.choices(string.digits, k=length))


def _fake_address() -> str:
    return (
        f"{random.choice(FAKE_PROVINCES)}省"
        f"{random.choice(['城区', '新区', '高新区', '经济开发区', '中心区'])}"
        f"{random.choice(FAKE_ADDRESS_SUFFIX)}{random.randint(1, 299)}号"
    )


def _fake_generic(text: str) -> str:
    length = len(text)
    if length <= 0:
        return ""
    choices = string.ascii_letters + string.digits
    return "".join(random.choices(choices, k=length))


def generate_fake_value(value, category: str) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if category == "name":
        return _fake_chinese_name(len(text))
    if category == "id":
        return _fake_id_like(text)
    if category == "phone":
        return _fake_phone_like(text)
    if category == "address":
        return _fake_address()
    return _fake_generic(text)


def desensitize_dataframe(df: pd.DataFrame):
    sanitized = df.copy()
    sensitive_columns = {}
    for col in sanitized.columns:
        sanitized[col] = sanitized[col].apply(lambda x: "" if pd.isna(x) else str(x))
        category = classify_sensitive_column(col)
        if category:
            sanitized[col] = sanitized[col].apply(lambda v: generate_fake_value(v, category))
            sensitive_columns[col] = category
    return sanitized, sensitive_columns

def find_best_match(target_field, source_fields, threshold=85):
    """
    【新增】使用多层策略为目标字段查找最佳的源字段匹配
    """
    norm_tgt = target_field.strip().lower()
    
    # 预定义的近义词字典
    synonyms = {
        "手机号码": ["手机号", "联系方式", "电话", "联系电话"],
        "毕业院校": ["毕业学校", "学校", "院校"],
        "身份证": ["身份证号", "身份证号码", "id", "证件号码"],
        "工号": ["员工编号", "员工id", "eid"],
        "姓名": ["员工姓名"],
        "学历": ["最高学历"]
    }
    
    best_match = None
    highest_score = 0

    # 策略1: 精确与标准化精确匹配
    for src in source_fields:
        if src == target_field: return src
        if norm_str(src).lower() == norm_tgt:
            best_match = src
            highest_score = 101  # 赋予最高优先级

    if best_match: return best_match

    # 策略2: 近义词匹配
    synonym_group = []
    for key, values in synonyms.items():
        if norm_tgt == key.lower() or norm_tgt in [v.lower() for v in values]:
            synonym_group.extend([key.lower()] + [v.lower() for v in values])
            break
    
    if synonym_group:
        for src in source_fields:
            norm_src = norm_str(src).lower()
            if norm_src in synonym_group:
                score = fuzz.ratio(norm_tgt, norm_src)
                if score > highest_score:
                    highest_score = score
                    best_match = src
    
    if best_match: return best_match

    # 策略3: 模糊相似度匹配
    for src in source_fields:
        score = fuzz.partial_ratio(norm_tgt, norm_str(src).lower())
        if score > highest_score:
            highest_score = score
            best_match = src
    
    if highest_score >= threshold:
        return best_match
    
    return None

def read_excel_dataframe(path: Path, sheet_name: str, header_row: int, data_start_col: int, drop_all_blank_rows: bool=True):
    """
    【关键】安全的Excel文件读取，通过 dtype=str 保证数据格式不变
    """
    try:
        # dtype=str 是保证'00001'和手机号不被转换的关键
        df = pd.read_excel(
            path, 
            sheet_name=sheet_name, 
            header=header_row - 1,
            engine="openpyxl", 
            dtype=str
        ).fillna('')
        
        if data_start_col > 1:
            df = df.iloc[:, data_start_col - 1:]

        if drop_all_blank_rows: 
            df = df.dropna(how="all")
        
        raw_cols = [norm_str(c) or f"(无名列){get_column_letter(i + data_start_col)}" for i, c in enumerate(df.columns)]
        df.columns = dedup_columns(raw_cols)
        
        gc.collect()
        return df
        
    except Exception as e:
        gc.collect()
        raise e

def get_sheet_names_safely(path: Path):
    """安全获取工作表名称，确保文件句柄正确关闭"""
    wb = None
    try:
        wb = load_workbook(path, read_only=True, keep_vba=True)
        return wb.sheetnames.copy()
    finally:
        if wb: wb.close()
        gc.collect()

def suggest_index_choice(columns):
    """根据常见词推荐索引列"""
    prefer = {"姓名", "name", "Name", "Full Name", "姓名（必填）", "工号"}
    for w in prefer:
        for c in columns:
            if norm_str(c) == w: return c
    return columns[0] if columns else None

def auto_detect_header_start(path: Path, sheet_name: str, max_rows: int = 50):
    """自动识别表头行和数据起始列"""
    try:
        df = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=max_rows,
            engine="openpyxl",
            dtype=str
        ).fillna('')
    except Exception:
        return 1, 1

    keywords = ["姓名", "名称", "工号", "号码", "电话", "时间", "地址", "金额"]
    best_row, best_score = 0, -1

    for idx, row in df.iterrows():
        cells = [norm_str(c) for c in row.tolist()]
        non_empty = [c for c in cells if c]
        if not non_empty:
            continue
        non_empty_count = len(non_empty)
        text_cells = [c for c in non_empty if not re.fullmatch(r"-?\d+(?:\.\d+)?", c)]
        text_ratio = len(text_cells) / non_empty_count
        keyword_hits = sum(1 for c in non_empty for kw in keywords if kw in c)
        score = non_empty_count + text_ratio * 5 + keyword_hits * 10
        if score > best_score:
            best_score = score
            best_row = idx

    header_row = best_row + 1
    header_cells = [norm_str(c) for c in df.iloc[best_row].tolist()]
    start_col = 1
    for j, val in enumerate(header_cells):
        if val:
            start_col = j + 1
            break

    return header_row, start_col

class ComboDelegate(QStyledItemDelegate):
    """用于表格内嵌下拉框的代理"""
    def __init__(self, parent, options):
        super().__init__(parent)
        self.options = options
    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.addItems(self.options)
        return combo
    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole)
        i = editor.findText(value) if value else -1
        editor.setCurrentIndex(i if i >= 0 else 0)
    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)

# ===================== Excel 写入与导出 =====================

def excel_com_write_and_save_optimized(tgt_path: Path, tgt_sheet: str, out_path: Path,
                                     df_src: pd.DataFrame, df_tgt: pd.DataFrame, src_map: pd.DataFrame,
                                     mapping: list, tgt_field_to_col: dict, tgt_data_start_row: int,
                                     overwrite_all: bool):
    """【性能核心】使用批量写入提升COM性能，按列批量操作而非逐格"""
    try: 
        import win32com.client as win32
    except Exception as e: 
        raise RuntimeError("未安装或无法加载 pywin32。") from e
    
    excel = None
    wb = None
    total_found, total_write = 0, 0
    
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.Calculation = win32.constants.xlCalculationManual
        
        wb = excel.Workbooks.Open(str(tgt_path))
        ws = wb.Worksheets(tgt_sheet)
        
        idx_to_row = {row.get("_IDX_", ""): tgt_data_start_row + row_offset
                      for row_offset, (_, row) in enumerate(df_tgt.iterrows())
                      if row.get("_IDX_", "") in src_map.index}
        total_found = len(idx_to_row)
        
        for tgt_field, src_field in mapping:
            if not src_field or src_field not in df_src.columns: continue
            tgt_col = tgt_field_to_col.get(tgt_field)
            if not tgt_col: continue
            
            updates = []
            for idx_val, excel_row in idx_to_row.items():
                val = src_map.loc[idx_val].get(src_field)
                if val == '' or pd.isna(val): continue
                
                if not overwrite_all:
                    cell = ws.Cells(excel_row, tgt_col)
                    if cell.Value is not None and str(cell.Value).strip() != "": continue
                
                updates.append((excel_row, val))
            
            if updates:
                updates.sort(key=lambda x: x[0])
                i = 0
                while i < len(updates):
                    start_row, values = updates[i][0], [updates[i][1]]
                    j = i + 1
                    while j < len(updates) and updates[j][0] == updates[j-1][0] + 1:
                        values.append(updates[j][1]); j += 1
                    
                    if len(values) == 1:
                        ws.Cells(start_row, tgt_col).Value = values[0]
                    else:
                        end_row = start_row + len(values) - 1
                        range_addr = f"{get_column_letter(tgt_col)}{start_row}:{get_column_letter(tgt_col)}{end_row}"
                        ws.Range(range_addr).Value = [[v] for v in values]
                    
                    total_write += len(values)
                    i = j
        
        excel.Calculation = win32.constants.xlCalculationAutomatic
        excel.ScreenUpdating = True
        
        ext = out_path.suffix.lower()
        ff = 52 if ext == ".xlsm" else 51 if ext == ".xlsx" else None
        wb.SaveAs(str(out_path), FileFormat=ff) if ff else wb.SaveAs(str(out_path))
        
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        gc.collect()
    
    return total_found, total_write

# ===================== AI 助手 =====================

def summarize_error(msg: str, columns=None) -> str:
    """提取错误信息的关键信息，减少token占用"""
    if not msg:
        return "未知错误"
    line = msg.strip().splitlines()[-1]
    m = re.search(r"KeyError: '([^']+)'", line)
    if m:
        info = f"执行因 KeyError 失败，缺少列 '{m.group(1)}'"
        if columns:
            info += f"。可用列: {list(columns)}"
        return info
    return line


class AIWorker(QThread):
    progress = pyqtSignal(str)
    success = pyqtSignal(str)
    error = pyqtSignal(str)
    code_stream = pyqtSignal(str)
    code_ready = pyqtSignal(str)

    def __init__(self, api_key, tables, instruction, temperature, language, output_path, output_format, conversation_history=None):
        super().__init__()
        self.api_key = api_key
        self.tables = [str(p) for p in tables]
        self.instruction = instruction
        self.temperature = temperature
        self.language = language
        self.output_path = Path(output_path)
        self.output_format = output_format if output_format in EXPORT_FORMATS else "excel"
        self.approval_event = threading.Event()
        self.history = conversation_history or []

    def approve_execution(self):
        self.approval_event.set()

    def _finalize_success(self, all_columns):
        expected_path = self.output_path
        if not expected_path.exists():
            return False, f"未生成指定路径的文件：{expected_path}"

        try:
            if self.output_format == "excel":
                suffix = expected_path.suffix.lower()
                if suffix in {".xlsx", ".xlsm"}:
                    load_workbook(expected_path).close()
                elif suffix == ".xls":
                    if expected_path.stat().st_size <= 0:
                        return False, "生成的Excel文件内容为空。"
                else:
                    return False, f"文件扩展名与期望的Excel格式不符：{suffix}"
            else:
                if expected_path.stat().st_size <= 0:
                    return False, "生成的文件内容为空。"
        except Exception as e:
            return False, summarize_error(str(e), all_columns)

        self.success.emit(str(expected_path))
        return True, None

    def run(self):
        try:
            from openai import OpenAI
        except Exception as e:
            self.error.emit(f"未安装openai库: {e}")
            return

        self.progress.emit("读取表格示例...")
        table_texts = []
        all_columns = set()
        for p in self.tables:
            if self.isInterruptionRequested():
                self.error.emit("已取消")
                return
            try:
                df = pd.read_excel(p, dtype=str).fillna("")
            except Exception as e:
                self.error.emit(f"无法读取表格：{p} - {e}")
                return
            sanitized_df, sensitive_map = desensitize_dataframe(df)
            sample = sanitized_df.head(5).to_csv(sep='\t', index=False).strip()
            cols = ", ".join(map(str, df.columns))
            all_columns.update(df.columns)
            block_lines = [
                f"## {Path(p).name}",
                f"路径: {p}",
                f"列: {cols}",
            ]
            if sensitive_map:
                block_lines.append(f"已自动脱敏列: {', '.join(sensitive_map.keys())}")
            block_lines.append("示例:")
            block_lines.append(sample or "(空表)")
            table_texts.append("\n".join(block_lines))

        try:
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.error.emit(f"无法创建导出目录：{e}")
            return

        output_path_str = str(self.output_path)
        tables_json = json.dumps(self.tables, ensure_ascii=False) if self.tables else "[]"
        table_list_string = "\n".join(self.tables)
        table_info_text = "\n\n".join(table_texts) if table_texts else "（未选择表格）"
        language_key = self.language.lower()

        format_info = EXPORT_FORMATS.get(self.output_format, EXPORT_FORMATS["excel"])
        format_desc = format_info["description"]
        extensions_text = ", ".join(format_info["extensions"])

        history_lines = []
        for msg in self.history:
            role = msg.get("role")
            content = (msg.get("content") or "").strip()
            if not content:
                continue
            prefix = "用户" if role == "user" else "助手"
            history_lines.append(f"{prefix}：{content}")

        conversation_text = "\n".join(history_lines)
        if conversation_text:
            conversation_block = (
                "以下是之前的对话历史，请在继续编写代码时保持上下文一致：\n"
                f"{conversation_text}\n\n"
            )
        else:
            conversation_block = ""

        task_block = (
            f"{conversation_block}"
            f"当前用户最新指令：\n{self.instruction}\n\n"
            f"可用的Excel表格信息：\n{table_info_text}\n\n"
            f"目标输出文件类型：{format_desc}（可用扩展名：{extensions_text}）。\n"
            "所有示例数据均已在本地完成脱敏处理，仅用于展示列结构。\n"
            "以上绝对路径已由系统自动注入运行环境，无需等待用户重复提供。请直接使用这些路径，避免出现“Required Files Not Found”错误。"
        )

        if language_key == "vba":
            env_instructions = (
                "你将获得若干Excel文件的路径、列名以及前5行数据示例。请仅生成VBA代码。\n"
                "生成代码时必须直接使用提供的完整文件路径，避免因路径错误导致“Required Files Not Found”。\n"
                "必须声明一个入口宏：Sub ProcessTables(tableList As String, outputPath As String)。\n"
                "参数 tableList 为使用换行分隔的完整Excel路径字符串；outputPath 为结果Excel文件的完整保存路径。\n"
                f"运行环境会调用 ProcessTables(tableList, outputPath)，并且 outputPath 始终为：{output_path_str}（需生成 {format_desc}，扩展名：{extensions_text}）。\n"
                "请在宏内拆分 tableList，按需打开并处理这些工作簿，最终将结果保存到 outputPath 指定的路径。\n"
                "不要弹出对话框或依赖任何交互，也不要修改除结果文件外的其他文件。\n"
                "交互要求：在构思代码时，请同步准备一段代码执行成功后的反馈话术，先确认任务完成，再礼貌询问用户下一步需求。该话术无需写入代码，将在我们系统中展示给用户。\n"
                "仅返回纯VBA代码，不要包含```标记或额外说明。"
            )
            retry_suffix = "请仅返回修正后的VBA代码。"
        else:
            env_instructions = (
                "你将获得若干Excel文件的路径、列名以及前5行数据示例。请仅生成可直接运行的Python代码以满足用户需求。\n"
                "生成代码时必须直接使用提供的完整文件路径，避免因路径错误导致“Required Files Not Found”。\n"
                "运行环境提供了三个环境变量：AI_TABLE_PATHS（JSON数组，包含所有Excel完整路径）、AI_OUTPUT_PATH（结果文件完整路径）与 AI_OUTPUT_FORMAT（目标文件格式关键字）。\n"
                f"输出文件的目标路径固定为：{output_path_str}，并且期望生成 {format_desc}（扩展名：{extensions_text}）。请务必按此格式保存。\n"
                "代码完成后必须打印单行JSON，例如 print(json.dumps({'status':'success','output_path': output_path}, ensure_ascii=False))。\n"
                "不要输出任何解释或额外文本，也不要包含```代码块标记。\n"
                "交互要求：在规划代码时，请准备一段成功完成任务后的反馈话术，先确认任务完成，再询问用户下一步需求。该话术无需写入代码，将由我们的系统在执行成功后提示用户。"
            )
            retry_suffix = "请仅返回修正后的Python代码。"

        base_prompt = env_instructions + "\n\n" + task_block


        client = OpenAI(api_key=self.api_key, base_url="https://api.deepseek.com")
        attempt, last_err, last_code = 0, None, ""
        while attempt < 3:
            if self.isInterruptionRequested():
                self.error.emit("已取消")
                return

            self.approval_event.clear()
            prompt = base_prompt if not last_err else base_prompt + f"\n\n上次执行错误：{last_err}\n{retry_suffix}"
            self.progress.emit(f"调用模型生成{self.language}代码...尝试{attempt + 1}/3")
            try:
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    temperature=self.temperature,
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant"},
                        {"role": "user", "content": prompt}
                    ],
                    stream=True
                )
            except Exception as e:
                self.error.emit(str(e))
                return

            code = ""
            try:
                for chunk in response:
                    if self.isInterruptionRequested():
                        self.error.emit("已取消")
                        return
                    # Each streamed token arrives as a ChoiceDelta object; get its text safely
                    delta_obj = chunk.choices[0].delta
                    delta = getattr(delta_obj, "content", None)

                    if delta:
                        code += delta
                        self.code_stream.emit(code)
            except Exception as e:
                self.error.emit(str(e))
                return

            code = code.strip()
            if code.startswith("```"):
                code = "\n".join(code.splitlines()[1:-1])
            last_code = code

            self.code_ready.emit(code)
            self.progress.emit("等待用户确认...")
            self.approval_event.wait()
            if self.isInterruptionRequested():
                self.error.emit("已取消")
                return

            if self.output_path.exists():
                try:
                    self.output_path.unlink()
                except Exception as e:
                    last_err = f"无法覆盖现有输出文件：{e}"
                    attempt += 1
                    continue

            self.progress.emit("执行代码...")
            if language_key == "vba":
                try:
                    execute_vba_module(code, table_list_string, self.output_path)
                except Exception as e:
                    last_err = summarize_error(str(e), all_columns)
                    attempt += 1
                    continue

                expected_path = self.output_path
                if expected_path.exists():
                    ok, err_text = self._finalize_success(all_columns)
                    if ok:
                        return
                    last_err = err_text
                else:
                    last_err = f"未生成指定路径的文件：{expected_path}"
            else:
                with tempfile.TemporaryDirectory() as td:
                    script = Path(td) / "script.py"
                    script.write_text(code, encoding="utf-8")
                    env = os.environ.copy()
                    env.setdefault("PYTHONPATH", "")
                    env["AI_TABLE_PATHS"] = tables_json
                    env["AI_TABLE_LIST"] = table_list_string
                    env["AI_OUTPUT_PATH"] = output_path_str
                    env["AI_OUTPUT_FORMAT"] = self.output_format
                    env["AI_OUTPUT_EXTENSIONS"] = json.dumps(format_info["extensions"], ensure_ascii=False)
                    env["AI_INSTRUCTION_TEXT"] = self.instruction
                    try:
                        proc = subprocess.run(
                            [sys.executable, str(script)],
                            capture_output=True,
                            text=True,
                            cwd=td,
                            env=env,
                        )
                    except Exception as e:
                        last_err = str(e)
                        attempt += 1
                        continue

                stdout = proc.stdout.strip()
                stderr = proc.stderr.strip()
                if proc.returncode != 0:
                    last_err = summarize_error(stderr or stdout, all_columns)
                else:
                    expected_path = self.output_path
                    if expected_path.exists():
                        ok, err_text = self._finalize_success(all_columns)
                        if ok:
                            return
                        last_err = err_text
                    else:
                        result_json = None
                        if stdout:
                            try:
                                result_json = json.loads(stdout.splitlines()[-1])
                            except Exception:
                                result_json = None
                        if result_json and result_json.get("output_path"):
                            candidate = Path(result_json.get("output_path"))
                            if candidate.exists():
                                last_err = f"模型在 {candidate} 生成了文件，请将结果保存至指定路径：{expected_path}"
                            else:
                                last_err = f"未能在指定路径生成结果文件：{expected_path}"
                        else:
                            last_err = summarize_error(stdout or stderr, all_columns)


            attempt += 1

        self.error.emit((last_err or "执行失败") + f"\n\n最后的代码:\n{last_code}")


class IntentWorker(QThread):
    """识别上传表格的常见处理意图"""

    results = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, api_key: str, tables: List[str]):
        super().__init__()
        self.api_key = api_key
        self.tables = [str(p) for p in tables]

    def run(self):
        if not self.tables:
            self.results.emit([])
            return

        try:
            from openai import OpenAI
        except Exception as e:
            self.error.emit(f"未安装openai库: {e}")
            return

        table_chunks = []
        for path in self.tables:
            if self.isInterruptionRequested():
                return
            try:
                df = pd.read_excel(path, dtype=str).fillna("")
            except Exception as e:
                self.error.emit(f"无法读取表格：{path} - {e}")
                return
            sanitized_df, sensitive_map = desensitize_dataframe(df)
            sample_csv = sanitized_df.head(5).to_csv(index=False).strip()
            columns = json.dumps([str(c) for c in df.columns], ensure_ascii=False)
            chunk_lines = [
                f"文件名: {Path(path).name}",
                f"列名: {columns}",
            ]
            if sensitive_map:
                chunk_lines.append(f"已自动脱敏列: {', '.join(sensitive_map.keys())}")
            chunk_lines.append("数据样本:")
            chunk_lines.append(sample_csv or "(空表)")
            table_chunks.append("\n".join(chunk_lines))

        prompt_text = "以下是一个或多个Excel表格的结构与数据示例：\n\n" + "\n\n".join(table_chunks)

        client = OpenAI(api_key=self.api_key, base_url="https://api.deepseek.com")
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                temperature=0.0,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "你是一个精通Excel和数据处理的专家助理。你的任务是根据用户上传的表格的列名和少量示例数据，预测用户可能想要执行的几个最常见的操作。"
                            "请仅以JSON数组的格式返回5个操作意图，每个意图为2-5个字且动词开头。"
                        ),
                    },
                    {"role": "user", "content": prompt_text},
                ],
            )
        except Exception as e:
            if not self.isInterruptionRequested():
                self.error.emit(str(e))
            return

        if self.isInterruptionRequested():
            return

        content = ""
        try:
            content = (response.choices[0].message.content or "").strip()
        except Exception:
            self.error.emit("意图识别返回内容为空")
            return

        intents = None
        try:
            intents = json.loads(content)
        except Exception:
            try:
                start = content.find("[")
                end = content.rfind("]")
                if start != -1 and end != -1 and end > start:
                    intents = json.loads(content[start : end + 1])
            except Exception:
                intents = None

        if not isinstance(intents, list):
            self.error.emit("意图识别结果解析失败")
            return

        cleaned = []
        for item in intents:
            text = str(item).strip()
            if text:
                cleaned.append(text)

        self.results.emit(cleaned[:5])


def execute_vba_module(code: str, table_payload: str, output_path: Path):
    """在临时工作簿中插入并执行VBA代码"""
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError("未安装或无法加载 pywin32。") from e

    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Add()
        except Exception as e:
            raise RuntimeError(f"无法创建临时工作簿：{e}") from e

        try:
            module = wb.VBProject.VBComponents.Add(1)
        except Exception as e:
            raise RuntimeError("无法访问Excel VBA项目，请在Excel选项中启用“信任对VBA项目对象模型的访问”。") from e

        module.CodeModule.AddFromString(code)
        try:
            excel.Run("ProcessTables", table_payload, str(output_path))
        except Exception as e:
            raise RuntimeError(f"VBA 执行失败：{e}") from e

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        gc.collect()


class AIHelperDialog(QDialog):
    def __init__(self, parent, settings: AppSettings):
        super().__init__(parent)
        self.setWindowTitle("AI助手")
        self.resize(1200, 650)
        self.settings = settings
        self.tables = []
        self.conversation_messages = []
        self.worker = None
        self.awaiting_execution = False
        self._last_status_text = ""
        self.intent_worker = None
        self.recommend_buttons = []
        self.recommend_placeholder = None

        main_layout = QHBoxLayout(self)
        main_layout.setSpacing(12)

        history_group = QGroupBox("历史对话")
        history_group.setMinimumWidth(220)
        history_layout = QVBoxLayout(history_group)
        self.history_list = QListWidget()
        self.history_list.setAlternatingRowColors(True)
        self.history_list.setWordWrap(True)
        self.history_list.itemDoubleClicked.connect(self.show_history_detail)
        history_layout.addWidget(self.history_list)
        history_btn_row = QHBoxLayout()
        self.btn_new_session = QPushButton("开启新对话")
        self.btn_new_session.clicked.connect(self.start_new_session)
        history_btn_row.addWidget(self.btn_new_session)
        self.btn_clear_history = QPushButton("清空历史")
        self.btn_clear_history.clicked.connect(self.clear_history)
        history_btn_row.addWidget(self.btn_clear_history)
        history_layout.addLayout(history_btn_row)
        main_layout.addWidget(history_group, 1)

        center_group = QGroupBox("对话配置")
        center_layout = QVBoxLayout(center_group)
        center_layout.addWidget(QLabel("使用场景:"))
        self.scenario_combo = QComboBox()
        self.scenario_combo.addItems([
            "代码生成/数学解题",
            "数据抽取/分析",
            "通用对话",
            "翻译",
            "创意类写作/诗歌创作"
        ])
        center_layout.addWidget(self.scenario_combo)

        self.btn_add_table = QPushButton("添加表格")
        self.btn_add_table.clicked.connect(self.add_table)
        center_layout.addWidget(self.btn_add_table)

        self.table_list = QListWidget()
        self.table_list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table_list.setMinimumHeight(120)
        center_layout.addWidget(self.table_list, 1)

        center_layout.addWidget(QLabel("生成代码语言:"))
        self.language_combo = QComboBox()
        self.language_combo.addItems(["Python", "VBA"])
        center_layout.addWidget(self.language_combo)

        center_layout.addWidget(QLabel("导出文件格式:"))
        self.export_format_combo = QComboBox()
        for key, info in EXPORT_FORMATS.items():
            self.export_format_combo.addItem(info["label"], userData=key)
        self.export_format_combo.currentIndexChanged.connect(self.on_export_format_changed)
        center_layout.addWidget(self.export_format_combo)

        center_layout.addWidget(QLabel("导出结果路径:"))
        path_layout = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("请选择AI生成结果的保存路径")
        default_output = self.settings.get("last_ai_export_path", "") or ""
        if default_output:
            self.output_edit.setText(default_output)
        self._sync_format_combo_with_path(default_output)
        btn_browse = QPushButton("浏览…")
        btn_browse.clicked.connect(self.browse_output)
        path_layout.addWidget(self.output_edit, 1)
        path_layout.addWidget(btn_browse)
        center_layout.addLayout(path_layout)

        center_layout.addWidget(QLabel("对话输入:"))
        self.message_edit = QPlainTextEdit()
        self.message_edit.setPlaceholderText("请用自然语言描述下一步操作，例如：合并客户表的电话字段。系统会自动提供所选表格的完整路径。")
        self.message_edit.setMinimumHeight(140)
        center_layout.addWidget(self.message_edit)

        self.privacy_label = QLabel("<a href='#'>用户隐私声明</a>")
        self.privacy_label.setStyleSheet("color: #3b82f6;")
        self.privacy_label.setTextFormat(Qt.TextFormat.RichText)
        self.privacy_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        self.privacy_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.privacy_label.linkActivated.connect(self.show_privacy_statement)
        center_layout.addWidget(self.privacy_label, alignment=Qt.AlignmentFlag.AlignLeft)

        self.recommend_container = QWidget()
        self.recommend_layout = QHBoxLayout(self.recommend_container)
        self.recommend_layout.setContentsMargins(0, 0, 0, 0)
        self.recommend_layout.setSpacing(6)
        center_layout.addWidget(self.recommend_container)
        self._reset_recommend_placeholder("推荐操作将在此显示")

        self.btn_send = QPushButton("发送指令")
        self.btn_send.clicked.connect(self.send_message)
        center_layout.addWidget(self.btn_send)
        center_layout.addStretch()
        main_layout.addWidget(center_group, 2)

        preview_group = QGroupBox("预览")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_tabs = QTabWidget()
        self.code_preview = QPlainTextEdit()
        self.code_preview.setReadOnly(True)
        self.code_preview.setPlaceholderText("AI生成的代码将实时显示在此处。")
        self.preview_tabs.addTab(self.code_preview, "代码预览")
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("执行日志与提示将显示在此处。")
        self.preview_tabs.addTab(self.log_view, "执行日志")
        self.table_preview = QTableWidget()
        self.table_preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_preview.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_preview.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table_preview.horizontalHeader().setStretchLastSection(True)
        self.preview_tabs.addTab(self.table_preview, "表格预览")
        preview_layout.addWidget(self.preview_tabs, 1)
        self.status_label = QLabel("等待指令…")
        preview_layout.addWidget(self.status_label)
        btn_row = QHBoxLayout()
        self.btn_execute = QPushButton("执行生成代码")
        self.btn_execute.setEnabled(False)
        self.btn_execute.clicked.connect(self.exec_generated_code)
        self.btn_cancel = QPushButton("取消当前操作")
        self.btn_cancel.clicked.connect(self.cancel_current)
        btn_row.addWidget(self.btn_execute)
        btn_row.addWidget(self.btn_cancel)
        preview_layout.addLayout(btn_row)
        main_layout.addWidget(preview_group, 2)

    def add_table(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择表格", "", "Excel Files (*.xlsx *.xlsm *.xls)")
        if not paths:
            return
        new_paths = []
        for p in paths:
            if p not in self.tables:
                self.tables.append(p)
                self.table_list.addItem(p)
                new_paths.append(p)
        if new_paths:
            self.log_view.appendPlainText(f"系统：已添加 {len(new_paths)} 个表格。")
            self.log_view.appendPlainText("系统：以下完整路径已同步给AI，您无需在指令中重复填写，供确认参考：")
            for path in new_paths:
                self.log_view.appendPlainText(f"    {path}")
            for path in new_paths:
                self._scan_table_privacy(path)
            self.trigger_intent_recognition()

    def browse_output(self):
        current = self.output_edit.text().strip()
        initial_dir = current
        if current:
            cp = Path(current)
            initial_dir = str(cp.parent if cp.suffix else cp)
        else:
            stored = self.settings.get("last_ai_export_path", "") or ""
            if stored:
                try:
                    initial_dir = str(Path(stored).parent)
                except Exception:
                    initial_dir = str(Path.home())
            else:
                initial_dir = str(Path.home())
        fmt_key = self.export_format_combo.currentData()
        fmt_info = EXPORT_FORMATS.get(fmt_key, EXPORT_FORMATS["excel"])
        path, _ = QFileDialog.getSaveFileName(self, "选择导出文件", initial_dir, fmt_info["dialog_filter"])
        if path:
            p = Path(path)
            if not p.suffix or p.suffix.lower() not in fmt_info["extensions"]:
                p = p.with_suffix(fmt_info["default_extension"])
            self.output_edit.setText(str(p))
            self._sync_format_combo_with_path(str(p))
            self.settings.update(last_ai_export_path=str(p))

    def _get_format_info(self, fmt_key=None):
        key = fmt_key if fmt_key in EXPORT_FORMATS else None
        if key is None:
            current_key = self.export_format_combo.currentData() if hasattr(self, "export_format_combo") else None
            key = current_key if current_key in EXPORT_FORMATS else "excel"
        return EXPORT_FORMATS.get(key, EXPORT_FORMATS["excel"])

    def on_export_format_changed(self, index):
        del index
        fmt_info = self._get_format_info()
        current_path = self.output_edit.text().strip()
        if not current_path:
            return
        p = Path(current_path)
        if p.suffix.lower() not in fmt_info["extensions"]:
            new_path = p.with_suffix(fmt_info["default_extension"])
            self.output_edit.setText(str(new_path))
            self.settings.update(last_ai_export_path=str(new_path))

    def _sync_format_combo_with_path(self, path: str):
        if not hasattr(self, "export_format_combo") or self.export_format_combo is None:
            return
        suffix = Path(path).suffix.lower() if path else ""
        matched_index = None
        for idx in range(self.export_format_combo.count()):
            fmt_key = self.export_format_combo.itemData(idx)
            fmt_info = EXPORT_FORMATS.get(fmt_key, EXPORT_FORMATS["excel"])
            if suffix and suffix in fmt_info["extensions"]:
                matched_index = idx
                break
        self.export_format_combo.blockSignals(True)
        self.export_format_combo.setCurrentIndex(matched_index if matched_index is not None else 0)
        self.export_format_combo.blockSignals(False)

    def show_privacy_statement(self, link=None):
        del link
        dlg = QDialog(self)
        dlg.setWindowTitle("用户隐私声明")
        layout = QVBoxLayout(dlg)
        viewer = QTextBrowser()
        viewer.setMarkdown(PRIVACY_STATEMENT_TEXT)
        viewer.setReadOnly(True)
        viewer.setMinimumSize(560, 360)
        layout.addWidget(viewer)
        btn = QPushButton("关闭")
        btn.clicked.connect(dlg.accept)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.resize(640, 480)
        dlg.exec()

    def _scan_table_privacy(self, path: str):
        try:
            df = pd.read_excel(path, dtype=str, nrows=200).fillna("")
        except Exception as e:
            self.log_view.appendPlainText(f"系统：无法读取 {path} 进行脱敏检查：{e}")
            return
        _, sensitive_map = desensitize_dataframe(df)
        if sensitive_map:
            cols = ", ".join(sensitive_map.keys())
            self.log_view.appendPlainText(f"系统：检测到敏感字段 {cols}，已自动使用随机样本对AI上下文进行脱敏。")
        else:
            self.log_view.appendPlainText("系统：未检测到明显敏感字段，仍会对上下文样本执行脱敏处理。")

    def append_history(self, role: str, content: str):
        content = (content or "").strip()
        if not content:
            return
        prefix = "👤" if role == "user" else "🤖"
        first_line = content.splitlines()[0]
        if len(first_line) > 60:
            first_line = first_line[:60] + "…"
        item = QListWidgetItem(f"{prefix} {first_line}")
        item.setData(Qt.ItemDataRole.UserRole, content)
        self.history_list.addItem(item)
        self.history_list.scrollToBottom()
        self.conversation_messages.append({"role": role, "content": content})

    def show_history_detail(self, item):
        content = item.data(Qt.ItemDataRole.UserRole) or ""
        dlg = QDialog(self)
        dlg.setWindowTitle("对话详情")
        lay = QVBoxLayout(dlg)
        txt = QPlainTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(content)
        lay.addWidget(txt)
        btn = QPushButton("关闭")
        btn.clicked.connect(dlg.accept)
        lay.addWidget(btn)
        dlg.resize(520, 320)
        dlg.exec()

    def clear_history(self, checked=False, confirm_dialog=True):
        del checked
        if confirm_dialog and self.conversation_messages:
            confirm = QMessageBox.question(self, "确认", "确定要清空历史对话吗？")
            if confirm != QMessageBox.StandardButton.Yes:
                return
        self.conversation_messages.clear()
        self.history_list.clear()
        self.log_view.appendPlainText("系统：已清空历史对话。")

    def start_new_session(self, checked=False):
        del checked
        if self.worker and self.worker.isRunning():
            confirm = QMessageBox.question(
                self,
                "确认",
                "当前有任务正在执行，是否取消并开启新对话？",
            )
            if confirm != QMessageBox.StandardButton.Yes:
                return
            self.cancel_current()

        self._stop_intent_worker()
        self.log_view.clear()
        self.clear_history(confirm_dialog=False)
        self.tables.clear()
        self.table_list.clear()
        self.message_edit.clear()
        self.code_preview.clear()
        self.table_preview.clear()
        self.table_preview.setRowCount(0)
        self.table_preview.setColumnCount(0)
        self.preview_tabs.setCurrentWidget(self.code_preview)
        self.status_label.setText("等待指令…")
        self._last_status_text = ""
        self.awaiting_execution = False
        self.btn_execute.setEnabled(False)
        self._reset_recommend_placeholder("推荐操作将在此显示")
        self.log_view.appendPlainText("系统：已开启新对话，所有内容已重置。")

    def _stop_intent_worker(self):
        if self.intent_worker and self.intent_worker.isRunning():
            self.intent_worker.requestInterruption()
            self.intent_worker.wait(200)
        self.intent_worker = None

    def _clear_recommend_layout(self):
        while self.recommend_layout.count():
            item = self.recommend_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _reset_recommend_placeholder(self, text="推荐操作将在此显示"):
        self.recommend_buttons.clear()
        self._clear_recommend_layout()
        label = QLabel(text)
        label.setStyleSheet("color: #94a3b8;")
        self.recommend_placeholder = label
        self.recommend_layout.addWidget(label)
        self.recommend_layout.addStretch()

    def trigger_intent_recognition(self):
        if not getattr(self, "recommend_layout", None):
            return
        if not self.tables:
            self._reset_recommend_placeholder("推荐操作将在此显示")
            return

        api_key = (self.settings.get("ai_api_key", "") or "").strip()
        if not api_key:
            self._reset_recommend_placeholder("请在设置中填写API Key以获取推荐。")
            return

        self._reset_recommend_placeholder("正在分析常见操作…")
        self.log_view.appendPlainText("系统：正在分析表格，推荐常见操作…")
        self._stop_intent_worker()
        self.intent_worker = IntentWorker(api_key, self.tables)
        self.intent_worker.results.connect(self.on_intent_results)
        self.intent_worker.error.connect(self.on_intent_error)
        self.intent_worker.finished.connect(self.on_intent_finished)
        self.intent_worker.start()

    def on_intent_results(self, intents: List[str]):
        self.recommend_buttons.clear()
        self._clear_recommend_layout()
        if not intents:
            self._reset_recommend_placeholder("暂未获取推荐操作")
            return

        for text in intents:
            btn = QPushButton(text)
            btn.clicked.connect(lambda _, t=text: self.on_recommendation_clicked(t))
            self.recommend_layout.addWidget(btn)
            self.recommend_buttons.append(btn)

        self.recommend_layout.addStretch()
        self.recommend_placeholder = None
        self.log_view.appendPlainText("系统：推荐操作 → " + "、".join(intents))

    def on_intent_error(self, message: str):
        first_line = (message or "").splitlines()[0] if message else ""
        self._reset_recommend_placeholder("未能获取推荐操作")
        if first_line:
            self.log_view.appendPlainText(f"系统：推荐操作生成失败：{first_line}")

    def on_intent_finished(self):
        self.intent_worker = None

    def on_recommendation_clicked(self, text: str):
        text = text.strip()
        if not text:
            return
        self.message_edit.setPlainText(text)
        self.message_edit.moveCursor(QTextCursor.MoveOperation.End)
        self.send_message()

    def render_table_preview(self, path_str: str):
        self.table_preview.clear()
        self.table_preview.setRowCount(0)
        self.table_preview.setColumnCount(0)
        if not path_str:
            return
        try:
            df = pd.read_excel(path_str)
        except Exception as e:
            self.log_view.appendPlainText(f"系统：无法加载表格预览：{e}")
            return

        preview_df = df.head(200)
        columns = [str(c) for c in preview_df.columns]
        self.table_preview.setColumnCount(len(columns))
        if columns:
            self.table_preview.setHorizontalHeaderLabels(columns)
        self.table_preview.setRowCount(len(preview_df))
        for row_idx, (_, row) in enumerate(preview_df.iterrows()):
            for col_idx, value in enumerate(row):
                display = "" if pd.isna(value) else str(value)
                self.table_preview.setItem(row_idx, col_idx, QTableWidgetItem(display))

        self.table_preview.resizeColumnsToContents()
        self.preview_tabs.setCurrentWidget(self.table_preview)

    def _adopt_result_table(self, path_str: str):
        if not path_str:
            return
        path_str = str(path_str)
        self.tables = [path_str]
        self.table_list.clear()
        self.table_list.addItem(path_str)
        self.log_view.appendPlainText("系统：合并成功，已将生成结果设定为当前处理表格。")
        self.render_table_preview(path_str)
        self.trigger_intent_recognition()

    def send_message(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "提示", "当前有任务正在执行，请稍候。")
            return

        api_key = (self.settings.get("ai_api_key", "") or "").strip()
        if not api_key:
            QMessageBox.warning(self, "提示", "请先在“设置”中填写API Key。")
            return

        if not self.tables:
            QMessageBox.warning(self, "提示", "请至少添加一个表格")
            return

        message = self.message_edit.toPlainText().strip()
        if not message:
            QMessageBox.warning(self, "提示", "请填写对话指令")
            return

        output_path_text = self.output_edit.text().strip()
        if not output_path_text:
            QMessageBox.warning(self, "提示", "请先选择导出结果路径")
            return
        output_path = Path(output_path_text)
        fmt_key = self.export_format_combo.currentData()
        fmt_info = self._get_format_info(fmt_key)
        if not output_path.suffix or output_path.suffix.lower() not in fmt_info["extensions"]:
            output_path = output_path.with_suffix(fmt_info["default_extension"])
            self.output_edit.setText(str(output_path))
            self._sync_format_combo_with_path(str(output_path))
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.warning(self, "提示", f"无法创建导出目录：{e}")
            return

        self.settings.update(last_ai_export_path=str(output_path))

        temp_map = {
            "代码生成/数学解题": 0.0,
            "数据抽取/分析": 1.0,
            "通用对话": 1.3,
            "翻译": 1.3,
            "创意类写作/诗歌创作": 1.5
        }
        temperature = temp_map.get(self.scenario_combo.currentText(), 0.0)
        language = self.language_combo.currentText()

        history_snapshot = [msg.copy() for msg in self.conversation_messages]
        self.append_history("user", message)
        self.log_view.appendPlainText(f"用户：{message}")
        self.log_view.appendPlainText("")

        self.message_edit.clear()
        self.code_preview.clear()
        self.preview_tabs.setCurrentIndex(0)
        self.status_label.setText("准备中…")
        self._last_status_text = ""
        self.btn_send.setEnabled(False)
        self.btn_execute.setEnabled(False)
        self.awaiting_execution = False

        self.worker = AIWorker(api_key, self.tables, message, temperature, language, str(output_path), fmt_key, history_snapshot)
        self.worker.progress.connect(self.on_worker_progress)
        self.worker.code_stream.connect(self.on_worker_code_stream)
        self.worker.code_ready.connect(self.on_worker_code_ready)
        self.worker.success.connect(self.on_worker_success)
        self.worker.error.connect(self.on_worker_error)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.start()

    def exec_generated_code(self):
        if not (self.worker and self.worker.isRunning() and self.awaiting_execution):
            return
        self.awaiting_execution = False
        self.btn_execute.setEnabled(False)
        self.status_label.setText("执行中…")
        self.log_view.appendPlainText("系统：开始执行生成的代码。")
        self.worker.approve_execution()

    def cancel_current(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption()
            self.worker.approve_execution()
            self.status_label.setText("已请求取消…")
            self.log_view.appendPlainText("系统：已请求取消当前任务。")
            self.awaiting_execution = False
            self.btn_execute.setEnabled(False)
        else:
            self.log_view.appendPlainText("系统：当前没有正在执行的任务。")

    def on_worker_progress(self, text: str):
        self.status_label.setText(text)
        if text and text != self._last_status_text:
            self.log_view.appendPlainText(f"系统：{text}")
            self._last_status_text = text

    def on_worker_code_stream(self, text: str):
        self.code_preview.setPlainText(text)
        sb = self.code_preview.verticalScrollBar()
        sb.setValue(sb.maximum())

    def on_worker_code_ready(self, text: str):
        self.on_worker_code_stream(text)
        self.awaiting_execution = True
        self.btn_execute.setEnabled(True)
        self.status_label.setText("代码生成完成，请确认后执行。")
        self.log_view.appendPlainText("系统：模型已生成代码，等待执行。")
        self._last_status_text = "代码生成完成，请确认后执行。"

    def on_worker_success(self, path_str: str):
        msg = f"表格处理已完成，结果保存在：{path_str}。还需要我继续协助处理这个结果吗？"
        self.append_history("assistant", msg)
        self.log_view.appendPlainText(f"成功：{path_str}")
        self.status_label.setText("执行完成")
        self._last_status_text = "执行完成"
        self._adopt_result_table(path_str)
        QMessageBox.information(self, "执行完成", f"已生成文件：\n{path_str}")

    def on_worker_error(self, err: str):
        err = (err or "").strip()
        first_line = err.splitlines()[0] if err else "未知错误"
        self.append_history("assistant", f"执行失败：{first_line}")
        if err:
            self.log_view.appendPlainText("错误：")
            self.log_view.appendPlainText(err)
        self.status_label.setText("执行失败")
        self._last_status_text = "执行失败"
        if first_line != "已取消":
            dlg = QDialog(self)
            dlg.setWindowTitle("错误")
            lay = QVBoxLayout(dlg)
            lay.addWidget(QLabel("执行失败，以下是错误信息："))
            txt = QPlainTextEdit()
            txt.setReadOnly(True)
            txt.setPlainText(err)
            lay.addWidget(txt)
            btn = QPushButton("关闭")
            btn.clicked.connect(dlg.accept)
            lay.addWidget(btn)
            dlg.resize(520, 320)
            dlg.exec()

    def on_worker_finished(self):
        self.worker = None
        self.awaiting_execution = False
        self.btn_execute.setEnabled(False)
        self.btn_send.setEnabled(True)
        self._last_status_text = ""

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.cancel_current()
        self._stop_intent_worker()
        super().closeEvent(event)


class SettingsDialog(QDialog):
    def __init__(self, parent, settings: AppSettings):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("设置")
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("DeepSeek API Key:"))
        self.api_edit = QLineEdit(self.settings.get("ai_api_key", ""))
        self.api_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.api_edit)

        layout.addWidget(QLabel("主题模式:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["深色", "浅色"])
        current_theme = self.settings.get("theme", "dark")
        self.theme_combo.setCurrentIndex(0 if current_theme == "dark" else 1)
        layout.addWidget(self.theme_combo)

        layout.addWidget(QLabel("Excel 写入引擎:"))
        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["自动选择", "仅使用COM", "仅使用openpyxl"])
        engine_mode = self.settings.get("engine_mode", "auto")
        engine_index = {"auto": 0, "com": 1, "openpyxl": 2}.get(engine_mode, 0)
        self.engine_combo.setCurrentIndex(engine_index)
        layout.addWidget(self.engine_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def accept(self):
        theme = "dark" if self.theme_combo.currentIndex() == 0 else "light"
        engine_mode = {0: "auto", 1: "com", 2: "openpyxl"}[self.engine_combo.currentIndex()]
        self.settings.update(
            ai_api_key=self.api_edit.text().strip(),
            theme=theme,
            engine_mode=engine_mode
        )
        super().accept()


def openpyxl_write_and_save_optimized(tgt_path: Path, tgt_sheet: str, out_path: Path,
                                    df_src: pd.DataFrame, df_tgt: pd.DataFrame, src_map: pd.DataFrame,
                                    mapping: list, tgt_field_to_col: dict, tgt_data_start_row: int,
                                    overwrite_all: bool):
    """【兼容核心】openpyxl 版本的写入，不依赖Windows Office"""
    wb = None
    total_found, total_write = 0, 0
    
    try:
        wb = load_workbook(tgt_path, data_only=False, read_only=False, keep_vba=True)
        ws = wb[tgt_sheet]
        
        updates_by_col = {}
        for row_offset, (_, row) in enumerate(df_tgt.iterrows()):
            idx_val = row.get("_IDX_", "")
            if not idx_val or idx_val not in src_map.index: continue
            
            src_row = src_map.loc[idx_val]
            total_found += 1
            excel_row = tgt_data_start_row + row_offset
            
            for tgt_field, src_field in mapping:
                if not src_field or src_field not in df_src.columns: continue
                tgt_col = tgt_field_to_col.get(tgt_field)
                if not tgt_col: continue
                
                val = src_row.get(src_field)
                if val == '' or pd.isna(val): continue
                
                cell = ws.cell(row=excel_row, column=tgt_col)
                if not overwrite_all and cell.value is not None and str(cell.value).strip() != "": continue
                
                if tgt_col not in updates_by_col: updates_by_col[tgt_col] = []
                updates_by_col[tgt_col].append((excel_row, val))
        
        for tgt_col, updates in updates_by_col.items():
            for excel_row, val in updates:
                cell = ws.cell(row=excel_row, column=tgt_col)
                fmt = {
                    'font': copy(cell.font),
                    'fill': copy(cell.fill),
                    'border': copy(cell.border),
                    'alignment': copy(cell.alignment),
                    'number_format': cell.number_format,
                    'protection': copy(cell.protection)
                }
                cell.value = val
                cell.font = fmt['font']
                cell.fill = fmt['fill']
                cell.border = fmt['border']
                cell.alignment = fmt['alignment']
                cell.number_format = fmt['number_format']
                cell.protection = fmt['protection']
                total_write += 1
        
        wb.save(out_path)
        
    finally:
        if wb:
            try: wb.close()
            except: pass
        gc.collect()
    
    return total_found, total_write

# ===================== 主界面 =====================

class MapperUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = AppSettings()
        self.setWindowTitle("智能Vlookuper 1.1 -梁诗忻开发")
        self.resize(1200, 820)
        # 多表模式下，源表和目标表都可能有多个，通过列表维护
        self.src_groups, self.tgt_groups = [], []
        self.src_headers, self.tgt_headers = [], []
        self.mode = "one2one"  # 默认一对一
        self._init_ui()
        self._create_menus()
        self._init_ai_button()
        self._apply_style()

    def _init_ui(self):
        central = QWidget()
        layout = QGridLayout(central)
        layout.setContentsMargins(14, 12, 14, 12); layout.setSpacing(12)

        # ------ 模式选择 ------
        mode_box = QGroupBox("匹配模式")
        mode_layout = QHBoxLayout(mode_box)
        self.rb_one2one = QRadioButton("一对一"); self.rb_one2one.setChecked(True)
        self.rb_one2many = QRadioButton("一对多")
        self.rb_many2one = QRadioButton("多对一")
        mode_layout.addWidget(self.rb_one2one); mode_layout.addWidget(self.rb_one2many); mode_layout.addWidget(self.rb_many2one)
        self.rb_one2one.toggled.connect(self.on_mode_change)
        self.rb_one2many.toggled.connect(self.on_mode_change)
        self.rb_many2one.toggled.connect(self.on_mode_change)
        layout.addWidget(mode_box, 0, 0, 1, 2)

        # ------ 动态源/目标面板 ------
        self.src_container = QWidget(); self.src_layout = QVBoxLayout(self.src_container)
        self.src_layout.setContentsMargins(0,0,0,0); self.src_layout.setSpacing(6)
        self.tgt_container = QWidget(); self.tgt_layout = QVBoxLayout(self.tgt_container)
        self.tgt_layout.setContentsMargins(0,0,0,0); self.tgt_layout.setSpacing(6)

        self.btn_add_src = QPushButton("添加信息源"); self.btn_add_src.clicked.connect(self.add_source_group)
        self.btn_add_tgt = QPushButton("添加目标表"); self.btn_add_tgt.clicked.connect(self.add_target_group)
        self.src_tabs = QTabWidget(); self.tgt_tabs = QTabWidget()
        self.src_layout.addWidget(self.btn_add_src); self.src_layout.addWidget(self.src_tabs)
        self.tgt_layout.addWidget(self.btn_add_tgt); self.tgt_layout.addWidget(self.tgt_tabs)

        layout.addWidget(self.src_container, 1, 0)
        layout.addWidget(self.tgt_container, 1, 1)

        # 初始各添加一个组
        self.add_source_group()
        self.add_target_group()

        map_grp = QGroupBox("字段映射与执行")
        map_layout = QVBoxLayout(map_grp)
        self.map_table = QTableWidget(0, 2)
        self.map_table.setHorizontalHeaderLabels(["目标字段（写入此列）", "来自信息源（下拉选择/跳过）"])
        self.map_table.horizontalHeader().setStretchLastSection(True)
        self.map_table.verticalHeader().setVisible(False)
        self.map_table.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        
        opts_layout = QHBoxLayout()
        opts_layout.addWidget(QLabel("源索引:")); self.cmb_src_index = QComboBox(); opts_layout.addWidget(self.cmb_src_index, 1)
        opts_layout.addWidget(QLabel("目标索引:")); self.cmb_tgt_index = QComboBox(); opts_layout.addWidget(self.cmb_tgt_index, 1)

        write_mode_box = QGroupBox("写入模式"); write_mode_layout = QHBoxLayout(write_mode_box)
        self.rb_fill_empty = QRadioButton("仅填充空值"); self.rb_fill_empty.setChecked(True)
        self.rb_overwrite = QRadioButton("覆盖所有值")
        write_mode_layout.addWidget(self.rb_fill_empty); write_mode_layout.addWidget(self.rb_overwrite)
        opts_layout.addWidget(write_mode_box, 2)

        btns = QHBoxLayout()
        self.btn_load_config = QPushButton("加载方案"); self.btn_save_config = QPushButton("保存方案")
        self.btn_auto = QPushButton("自动预填"); self.btn_run = QPushButton("执行匹配并导出")
        btns.addWidget(self.btn_load_config); btns.addWidget(self.btn_save_config)
        btns.addStretch(); btns.addWidget(self.btn_auto); btns.addWidget(self.btn_run)
        
        map_layout.addWidget(self.map_table); map_layout.addLayout(opts_layout); map_layout.addLayout(btns)
        layout.addWidget(map_grp, 2, 0, 1, 2)

        self.btn_load_config.clicked.connect(self.load_mapping_config)
        self.btn_save_config.clicked.connect(self.save_mapping_config)
        self.btn_auto.clicked.connect(self.auto_fill_mapping)
        self.btn_run.clicked.connect(self.run_and_export)
        self.on_mode_change()
        self.setCentralWidget(central)

    def _create_menus(self):
        bar = self.menuBar()
        settings_menu = bar.addMenu("设置")
        action_prefs = QAction("首选项...", self)
        action_prefs.triggered.connect(self.open_settings_dialog)
        settings_menu.addAction(action_prefs)

        about_menu = bar.addMenu("关于")
        action_about = QAction("关于 SMART VLOOKUPER", self)
        action_about.triggered.connect(self.show_about_dialog)
        about_menu.addAction(action_about)

    def _update_ai_button_style(self):
        if not hasattr(self, "btn_ai"):
            return
        if self.settings.get("theme", "dark") == "dark":
            style = (
                "QPushButton {background:#f97316; color:white; border:none; border-radius:24px;}"
                "QPushButton:hover {background:#ea580c;}"
            )
        else:
            style = (
                "QPushButton {background:#fb923c; color:#1f2937; border:none; border-radius:24px;}"
                "QPushButton:hover {background:#f97316;}"
            )
        self.btn_ai.setStyleSheet(style)

    # ---- 动态增减源/目标组 ----
    def add_source_group(self):
        idx = len(self.src_groups) + 1
        g = self._build_config_group(f"信息源{idx}", is_source=True)
        self.src_groups.append(g)
        self.src_tabs.addTab(g, f"信息源{idx}")
        self.src_tabs.setCurrentWidget(g)

    def add_target_group(self):
        idx = len(self.tgt_groups) + 1
        g = self._build_config_group(f"目标表{idx}", is_source=False)
        self.tgt_groups.append(g)
        self.tgt_tabs.addTab(g, f"目标表{idx}")
        self.tgt_tabs.setCurrentWidget(g)

    def on_mode_change(self):
        if self.rb_one2one.isChecked():
            self.mode = "one2one"
        elif self.rb_one2many.isChecked():
            self.mode = "one2many"
        else:
            self.mode = "many2one"

        self.btn_add_src.setEnabled(self.mode == "many2one")
        self.btn_add_tgt.setEnabled(self.mode == "one2many")

        if self.mode != "many2one":
            while len(self.src_groups) > 1:
                g = self.src_groups.pop()
                idx = self.src_tabs.count() - 1
                self.src_tabs.removeTab(idx)
                g.deleteLater()
        if self.mode != "one2many":
            while len(self.tgt_groups) > 1:
                g = self.tgt_groups.pop()
                idx = self.tgt_tabs.count() - 1
                self.tgt_tabs.removeTab(idx)
                g.deleteLater()

        self._recalc_headers()

    def _recalc_headers(self):
        self.src_headers = []
        for g in self.src_groups:
            if hasattr(g, "_headers"):
                for h in g._headers:
                    if h not in self.src_headers:
                        self.src_headers.append(h)
        self.cmb_src_index.clear()
        if self.src_headers:
            self.cmb_src_index.addItems(self.src_headers)

        self.tgt_headers = []
        if self.tgt_groups and hasattr(self.tgt_groups[0], "_headers"):
            self.tgt_headers = self.tgt_groups[0]._headers
            self.cmb_tgt_index.clear(); self.cmb_tgt_index.addItems(self.tgt_headers)

        self.rebuild_mapping_table()

    def _build_config_group(self, title, is_source: bool):
        g = QWidget(); grid = QGridLayout(g)
        grid.setContentsMargins(6,6,6,6)
        grid.setHorizontalSpacing(6)
        grid.setVerticalSpacing(4)
        le_path = QLineEdit(); le_path.setReadOnly(True)
        btn_browse = QPushButton("浏览…")
        cmb_sheet = QComboBox()
        sp_header = QSpinBox(); sp_header.setRange(1, 100000); sp_header.setValue(1)
        sp_startcol = QSpinBox(); sp_startcol.setRange(1, 10000); sp_startcol.setValue(1)
        btn_extract = QPushButton("提取字段并预览")
        preview_table = QTableWidget(); preview_table.setRowCount(5); preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        
        grid.addWidget(QLabel("Excel 文件："), 0, 0); grid.addWidget(le_path, 0, 1); grid.addWidget(btn_browse, 0, 2)
        grid.addWidget(QLabel("工作表："), 1, 0); grid.addWidget(cmb_sheet, 1, 1, 1, 2)
        grid.addWidget(QLabel("表头行："), 2, 0); grid.addWidget(sp_header, 2, 1)
        grid.addWidget(QLabel("数据起始列："), 2, 2); grid.addWidget(sp_startcol, 2, 3)
        grid.addWidget(btn_extract, 3, 0, 1, 4); grid.addWidget(preview_table, 4, 0, 1, 4)

        def auto_detect(_=None):
            p = le_path.text().strip()
            sheet = cmb_sheet.currentText().strip()
            if not p or not sheet:
                return
            try:
                h, c = auto_detect_header_start(Path(p), sheet)
                sp_header.setValue(h)
                sp_startcol.setValue(c)
                self._update_fields_and_preview(g, is_source)
            except Exception:
                pass

        def on_browse():
            path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel (*.xlsx *.xlsm)")
            if not path: return
            le_path.setText(path)
            try:
                sheet_names = get_sheet_names_safely(Path(path))
                cmb_sheet.clear(); cmb_sheet.addItems(sheet_names)
                auto_detect()
            except Exception as e: QMessageBox.critical(self, "错误", f"无法读取工作表：\n{e}")
        btn_browse.clicked.connect(on_browse)

        cmb_sheet.currentTextChanged.connect(auto_detect)

        def on_extract(): self._update_fields_and_preview(g, is_source)
        btn_extract.clicked.connect(on_extract)
        
        g._le_path, g._cmb_sheet, g._sp_header, g._sp_startcol, g._preview_table = le_path, cmb_sheet, sp_header, sp_startcol, preview_table
        return g

    def _update_fields_and_preview(self, group_box, is_source):
        p_str = group_box._le_path.text().strip()
        if not p_str: QMessageBox.warning(self, "提示", "请先选择Excel文件。"); return
        path, sheet = Path(p_str), group_box._cmb_sheet.currentText().strip()
        if not sheet: QMessageBox.warning(self, "提示", "请先选择工作表。"); return
        header_row, start_col = group_box._sp_header.value(), group_box._sp_startcol.value()

        try:
            df = read_excel_dataframe(path, sheet, header_row, start_col)
            headers = list(df.columns)
            self._update_preview_table(group_box._preview_table, df)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"提取字段失败：\n{e}"); return

        group_box._path, group_box._sheet = path, sheet
        group_box._header_row, group_box._start_col = header_row, start_col
        group_box._headers = headers

        if (not is_source) and self.tgt_groups and group_box != self.tgt_groups[0]:
            if hasattr(self.tgt_groups[0], "_headers") and headers != self.tgt_groups[0]._headers:
                QMessageBox.warning(self, "警告", "目标表字段结构不一致，可能导致匹配问题。")

        self._recalc_headers()

        if is_source and self.src_headers:
            if (guess := suggest_index_choice(self.src_headers)):
                self.cmb_src_index.setCurrentText(guess)
        if (not is_source) and self.tgt_headers:
            if (guess := suggest_index_choice(self.tgt_headers)):
                self.cmb_tgt_index.setCurrentText(guess)

    def _update_preview_table(self, table: QTableWidget, df: pd.DataFrame):
        preview_df = df.head(5)
        table.clear()
        table.setColumnCount(len(preview_df.columns))
        table.setHorizontalHeaderLabels(preview_df.columns)
        table.setRowCount(len(preview_df))
        for r_idx, row in enumerate(preview_df.itertuples(index=False)):
            for c_idx, val in enumerate(row):
                table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))
        table.resizeColumnsToContents()

    def rebuild_mapping_table(self):
        self.map_table.clearContents()
        self.map_table.setRowCount(len(self.tgt_headers))
        options = ["<跳过>"] + self.src_headers
        delegate = ComboDelegate(self.map_table, options)
        self.map_table.setItemDelegateForColumn(1, delegate)
        
        for r, tgt_name in enumerate(self.tgt_headers):
            item0 = QTableWidgetItem(tgt_name)
            item0.setFlags(item0.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.map_table.setItem(r, 0, item0)
            self.map_table.setItem(r, 1, QTableWidgetItem("<跳过>"))
        self.map_table.resizeColumnsToContents()

    def auto_fill_mapping(self):
        """【核心改进】使用模糊匹配自动填充，并提供视觉反馈"""
        if not self.src_headers: QMessageBox.warning(self, "提示", "请先提取信息源的字段。"); return

        unmatched_color = QBrush(QColor("#111827")) # 默认背景色
        matched_color = QBrush(QColor("#14532d"))   # 匹配成功的背景色

        for r in range(self.map_table.rowCount()):
            tgt_name = self.map_table.item(r, 0).text()
            best_match = find_best_match(tgt_name, self.src_headers)
            
            item_tgt = self.map_table.item(r, 0)
            item_src = self.map_table.item(r, 1)

            if best_match:
                item_src.setText(best_match)
                item_tgt.setBackground(matched_color)
                item_src.setBackground(matched_color)
            else:
                item_src.setText("<跳过>")
                item_tgt.setBackground(unmatched_color)
                item_src.setBackground(unmatched_color)

    def save_mapping_config(self):
        if self.mode != "one2one":
            QMessageBox.information(self, "提示", "仅一对一模式支持保存方案。")
            return

        g_src, g_tgt = self.src_groups[0], self.tgt_groups[0]
        if not hasattr(g_src, "_path") or not hasattr(g_tgt, "_path"):
            QMessageBox.warning(self, "提示", "请先配置并提取信息源和目标表的字段。"); return

        mapping = {self.map_table.item(r, 0).text(): self.map_table.item(r, 1).text() for r in range(self.map_table.rowCount())}
        config = {
            "source": {"path": str(g_src._path), "sheet": g_src._sheet, "header_row": g_src._header_row, "start_col": g_src._start_col},
            "target": {"path": str(g_tgt._path), "sheet": g_tgt._sheet, "header_row": g_tgt._header_row, "start_col": g_tgt._start_col},
            "indices": {"source": self.cmb_src_index.currentText(), "target": self.cmb_tgt_index.currentText()},
            "write_mode": "overwrite" if self.rb_overwrite.isChecked() else "fill_empty", "mapping": mapping
        }
        path, _ = QFileDialog.getSaveFileName(self, "保存映射方案", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "成功", f"映射方案已保存至:\n{path}")
            except Exception as e: QMessageBox.critical(self, "错误", f"保存失败：\n{e}")

    def load_mapping_config(self):
        if self.mode != "one2one":
            QMessageBox.information(self, "提示", "仅在一对一模式下可加载方案。")
            return

        path, _ = QFileDialog.getOpenFileName(self, "加载映射方案", "", "JSON Files (*.json)")
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f: config = json.load(f)

            for g, cfg in [(self.src_groups[0], config["source"]), (self.tgt_groups[0], config["target"])]:
                g._le_path.setText(cfg["path"])
                g._sp_header.setValue(cfg["header_row"])
                g._sp_startcol.setValue(cfg["start_col"])
                try:
                    sheet_names = get_sheet_names_safely(Path(cfg["path"]))
                    g._cmb_sheet.clear(); g._cmb_sheet.addItems(sheet_names)
                    g._cmb_sheet.setCurrentText(cfg["sheet"])
                except Exception as e: QMessageBox.warning(self, "警告", f"无法加载工作表 {cfg['path']}: {e}")

            self._update_fields_and_preview(self.src_groups[0], True)
            self._update_fields_and_preview(self.tgt_groups[0], False)

            self.cmb_src_index.setCurrentText(config["indices"]["source"])
            self.cmb_tgt_index.setCurrentText(config["indices"]["target"])
            self.rb_overwrite.setChecked(config["write_mode"] == "overwrite")

            for r in range(self.map_table.rowCount()):
                tgt_field = self.map_table.item(r, 0).text()
                src_field = config["mapping"].get(tgt_field, "<跳过>")
                self.map_table.item(r, 1).setText(src_field)

        except Exception as e: QMessageBox.critical(self, "错误", f"加载方案失败：\n{e}")

    def run_and_export(self):
        self.setEnabled(False)
        self.btn_run.setText("正在匹配...")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        QApplication.processEvents()
        try: 
            self._execute_matching_logic()
        finally:
            QApplication.restoreOverrideCursor()
            self.btn_run.setText("执行匹配并导出")
            self.setEnabled(True)

    def _execute_matching_logic(self):
        if not (self.src_groups and self.tgt_groups and self.src_headers and self.tgt_headers):
            QMessageBox.warning(self, "提示", "请先提取信息源和目标表的字段。"); return
        src_idx, tgt_idx = self.cmb_src_index.currentText(), self.cmb_tgt_index.currentText()
        if not src_idx or not tgt_idx:
            QMessageBox.warning(self, "提示", "请选择索引字段。"); return

        mapping = [(self.map_table.item(r,0).text(), self.map_table.item(r,1).text()) for r in range(self.map_table.rowCount())]
        mapping = [(t, s) for t, s in mapping if s != "<跳过>"]

        # ---- 构建源数据 ----
        try:
            if self.mode == "many2one":
                dfs = [read_excel_dataframe(g._path, g._sheet, g._header_row, g._start_col, True) for g in self.src_groups]
                df_src = pd.concat(dfs, ignore_index=True)
            else:
                g = self.src_groups[0]
                df_src = read_excel_dataframe(g._path, g._sheet, g._header_row, g._start_col, True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取信息源失败：\n{e}"); return

        if src_idx not in df_src.columns:
            QMessageBox.critical(self, "错误", f"源表无索引：{src_idx}"); return
        df_src["_IDX_"] = df_src[src_idx].apply(norm_str)
        src_map = df_src.drop_duplicates(subset=["_IDX_"], keep='last').set_index("_IDX_")

        engine_mode = self.settings.get("engine_mode", "auto")
        results = []
        for tgt_g in self.tgt_groups:
            try:
                df_tgt = read_excel_dataframe(tgt_g._path, tgt_g._sheet, tgt_g._header_row, tgt_g._start_col, False)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取目标表失败：\n{e}"); return
            if tgt_idx not in df_tgt.columns:
                QMessageBox.critical(self, "错误", f"目标表无索引：{tgt_idx}"); return
            df_tgt["_IDX_"] = df_tgt[tgt_idx].apply(norm_str)
            tgt_field_to_col = {name: i + tgt_g._start_col for i, name in enumerate(self.tgt_headers)}
            overwrite_all = self.rb_overwrite.isChecked()
            out_path = Path(tgt_g._path).with_name(f"{Path(tgt_g._path).stem}_匹配输出{Path(tgt_g._path).suffix}")

            engine = ""
            errors = {}
            if engine_mode in ("auto", "com"):
                try:
                    total_found, total_write = excel_com_write_and_save_optimized(
                        tgt_g._path, tgt_g._sheet, out_path, df_src, df_tgt, src_map,
                        mapping, tgt_field_to_col, tgt_g._header_row + 1, overwrite_all)
                    engine = "Excel COM（批量优化）"
                except Exception as e1:
                    errors["com"] = e1
                    if engine_mode == "com":
                        QMessageBox.critical(self, "错误", f"COM 保存失败：\n{e1}")
                        return
            if not engine and engine_mode in ("auto", "openpyxl"):
                try:
                    total_found, total_write = openpyxl_write_and_save_optimized(
                        tgt_g._path, tgt_g._sheet, out_path, df_src, df_tgt, src_map,
                        mapping, tgt_field_to_col, tgt_g._header_row + 1, overwrite_all)
                    engine = "openpyxl（Office365兼容）"

                    # 【关键修复】使用'replace'模式安全地进行Pandas验证回写
                    try:
                        df_verify = pd.read_excel(out_path, sheet_name=tgt_g._sheet, dtype=str).fillna('')
                        with pd.ExcelWriter(
                            out_path,
                            engine='openpyxl',
                            mode='a',
                            if_sheet_exists='replace'
                        ) as writer:
                            df_verify.to_excel(writer, sheet_name=tgt_g._sheet, index=False)
                        engine = "openpyxl（Pandas兼容性优化）"
                    except Exception:
                        # 验证失败，但原文件可能仍然可用
                        pass
                except Exception as e2:
                    errors["openpyxl"] = e2
                    if engine_mode == "openpyxl":
                        QMessageBox.critical(self, "错误", f"openpyxl 保存失败：\n{e2}")
                        return
            if not engine:
                def _short(err, default):
                    if err is None:
                        return default
                    msg = str(err)
                    return msg if len(msg) <= 200 else msg[:200] + "..."

                com_default = "未尝试（根据设置跳过）" if engine_mode == "openpyxl" else "无可用错误信息"
                op_default = "未尝试（根据设置跳过）" if engine_mode == "com" else "无可用错误信息"
                com_msg = _short(errors.get("com"), com_default)
                op_msg = _short(errors.get("openpyxl"), op_default)
                QMessageBox.critical(
                    self,
                    "错误",
                    f"所有保存方式均失败：\n\nCOM错误：{com_msg}\n\nopenpyxl错误：{op_msg}\n\n建议：\n1. 确保目标Excel文件未被其他程序占用\n2. 检查文件权限\n3. 尝试关闭Excel程序后重试"
                )
                return

            results.append((out_path, engine, total_found, total_write))

        if not results:
            return
        if len(results) == 1:
            out_path, engine, total_found, total_write = results[0]
            QMessageBox.information(self, "完成",
                f"匹配完成（引擎：{engine}）：\n\n"
                f"命中索引记录： {total_found}\n"
                f"共写入单元格： {total_write}\n\n"
                f"结果已导出至：\n{out_path}")
        else:
            msg = "\n\n".join([f"{p}: 命中{f} 写入{w}" for p, _, f, w in results])
            QMessageBox.information(self, "完成", f"已处理{len(results)}个目标表：\n\n{msg}")

    def _init_ai_button(self):
        self.btn_ai = QPushButton("AI", self)
        self.btn_ai.setFixedSize(48, 48)
        self._update_ai_button_style()
        self.btn_ai.clicked.connect(self.show_ai_dialog)
        self._position_ai_button()

    def _position_ai_button(self):
        x = self.width() - self.btn_ai.width() - 20
        y = self.height() - self.btn_ai.height() - 20
        self.btn_ai.move(x, y)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'btn_ai'):
            self._position_ai_button()

    def show_ai_dialog(self):
        dlg = AIHelperDialog(self, self.settings)
        dlg.exec()

    def _apply_style(self):
        theme = self.settings.get("theme", "dark")
        style = DARK_STYLESHEET if theme == "dark" else LIGHT_STYLESHEET
        app = QApplication.instance()
        if app:
            app.setStyleSheet(style)
        self._update_ai_button_style()

    def open_settings_dialog(self):
        dlg = SettingsDialog(self, self.settings)
        if dlg.exec():
            self._apply_style()

    def show_about_dialog(self):
        QMessageBox.information(
            self,
            "关于",
            "SMART VLOOKUPER 1.1\n©梁诗忻 2025. 本项目采用MIT许可证。\n项目地址：https://github.com/liangshixin1/Smart-Vlookuper"
        )

def main():
    app = QApplication(sys.argv)
    w = MapperUI()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
