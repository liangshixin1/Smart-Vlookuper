# -*- coding: utf-8 -*-
"""
通用字段映射与匹配工具（PyQt6） - v2.5
- 修复：修正了Pandas验证回写的逻辑，采用更安全的 'replace' 模式，防止文件损坏
- 新增：引入模糊匹配逻辑，增强“自动预填”的智能化程度
- 优化：COM批量写入，大幅提升性能（按列批量写入，而非逐格）
- 修复：加强异常处理和资源释放，确保文件句柄安全关闭
- 改进：更清晰的视觉反馈，自动匹配行会高亮显示

依赖: pip install pyqt6 pandas openpyxl pywin32 thefuzz
"""

import sys, os, re, warnings, json
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QGridLayout, QGroupBox, QLabel, QLineEdit, QPushButton, QComboBox,
    QSpinBox, QHBoxLayout, QVBoxLayout, QListWidget, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QStyledItemDelegate, QRadioButton,
    QButtonGroup
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QColor
import gc
from thefuzz import fuzz

# —— 静默 openpyxl 的数据验证扩展警告 ——
warnings.filterwarnings(
    "ignore",
    message=".*Data Validation extension is not supported and will be removed.*",
    category=UserWarning,
    module="openpyxl"
)

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
                ws.cell(row=excel_row, column=tgt_col).value = val
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
        self.setWindowTitle("智能Vlookuper 1.0.1 -梁诗忻开发")
        self.resize(1200, 820)
        self.src_path, self.tgt_path = None, None
        self.src_sheet, self.tgt_sheet = None, None
        self.src_header_row, self.tgt_header_row = 1, 1
        self.src_data_start_col, self.tgt_data_start_col = 1, 1
        self.src_headers, self.tgt_headers = [], []
        self._init_ui()
        self._apply_style()

    def _init_ui(self):
        central = QWidget()
        layout = QGridLayout(central)
        layout.setContentsMargins(14, 12, 14, 12); layout.setSpacing(12)

        self.grp_src = self._build_config_group("信息源", is_source=True)
        self.grp_tgt = self._build_config_group("目标表", is_source=False)
        layout.addWidget(self.grp_src, 0, 0); layout.addWidget(self.grp_tgt, 0, 1)

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
        layout.addWidget(map_grp, 1, 0, 1, 2)

        self.btn_load_config.clicked.connect(self.load_mapping_config)
        self.btn_save_config.clicked.connect(self.save_mapping_config)
        self.btn_auto.clicked.connect(self.auto_fill_mapping)
        self.btn_run.clicked.connect(self.run_and_export)
        self.setCentralWidget(central)

    def _build_config_group(self, title, is_source: bool):
        g = QGroupBox(title); grid = QGridLayout(g)
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
        
        if is_source:
            self.src_path, self.src_sheet, self.src_header_row, self.src_data_start_col, self.src_headers = path, sheet, header_row, start_col, headers
            self.cmb_src_index.clear(); self.cmb_src_index.addItems(headers)
            if (guess := suggest_index_choice(headers)): self.cmb_src_index.setCurrentText(guess)
        else:
            self.tgt_path, self.tgt_sheet, self.tgt_header_row, self.tgt_data_start_col, self.tgt_headers = path, sheet, header_row, start_col, headers
            self.cmb_tgt_index.clear(); self.cmb_tgt_index.addItems(headers)
            if (guess := suggest_index_choice(headers)): self.cmb_tgt_index.setCurrentText(guess)
        self.rebuild_mapping_table()

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
        if not self.src_path or not self.tgt_path: QMessageBox.warning(self, "提示", "请先配置并提取信息源和目标表的字段。"); return
        
        mapping = {self.map_table.item(r, 0).text(): self.map_table.item(r, 1).text() for r in range(self.map_table.rowCount())}
        config = {
            "source": {"path": str(self.src_path), "sheet": self.src_sheet, "header_row": self.src_header_row, "start_col": self.src_data_start_col},
            "target": {"path": str(self.tgt_path), "sheet": self.tgt_sheet, "header_row": self.tgt_header_row, "start_col": self.tgt_data_start_col},
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
        path, _ = QFileDialog.getOpenFileName(self, "加载映射方案", "", "JSON Files (*.json)")
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f: config = json.load(f)
            
            for g, cfg in [(self.grp_src, config["source"]), (self.grp_tgt, config["target"])]:
                g._le_path.setText(cfg["path"])
                g._sp_header.setValue(cfg["header_row"])
                g._sp_startcol.setValue(cfg["start_col"])
                try:
                    sheet_names = get_sheet_names_safely(Path(cfg["path"]))
                    g._cmb_sheet.clear(); g._cmb_sheet.addItems(sheet_names)
                    g._cmb_sheet.setCurrentText(cfg["sheet"])
                except Exception as e: QMessageBox.warning(self, "警告", f"无法加载工作表 {cfg['path']}: {e}")
            
            self._update_fields_and_preview(self.grp_src, True)
            self._update_fields_and_preview(self.grp_tgt, False)
            
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
        if not (self.src_path and self.tgt_path and self.src_headers and self.tgt_headers):
            QMessageBox.warning(self, "提示", "请先提取信息源和目标表的字段。"); return
        src_idx, tgt_idx = self.cmb_src_index.currentText(), self.cmb_tgt_index.currentText()
        if not src_idx or not tgt_idx: QMessageBox.warning(self, "提示", "请选择索引字段。"); return
        
        mapping = [(self.map_table.item(r,0).text(), self.map_table.item(r,1).text()) for r in range(self.map_table.rowCount())]
        mapping = [(t, s) for t, s in mapping if s != "<跳过>"]
        
        try:
            df_src = read_excel_dataframe(self.src_path, self.src_sheet, self.src_header_row, self.src_data_start_col, True)
            df_tgt = read_excel_dataframe(self.tgt_path, self.tgt_sheet, self.tgt_header_row, self.tgt_data_start_col, False)
        except Exception as e: QMessageBox.critical(self, "错误", f"读取数据失败：\n{e}"); return
        
        if src_idx not in df_src.columns: QMessageBox.critical(self, "错误", f"源表无索引：{src_idx}"); return
        if tgt_idx not in df_tgt.columns: QMessageBox.critical(self, "错误", f"目标表无索引：{tgt_idx}"); return

        df_src["_IDX_"] = df_src[src_idx].apply(norm_str)
        df_tgt["_IDX_"] = df_tgt[tgt_idx].apply(norm_str)
        src_map = df_src.drop_duplicates(subset=["_IDX_"], keep='last').set_index("_IDX_")
        
        tgt_field_to_col = {name: i + self.tgt_data_start_col for i, name in enumerate(self.tgt_headers)}
        overwrite_all = self.rb_overwrite.isChecked()
        out_path = self.tgt_path.with_name(f"{self.tgt_path.stem}_匹配输出{self.tgt_path.suffix}")

        try:
            total_found, total_write = excel_com_write_and_save_optimized(
                self.tgt_path, self.tgt_sheet, out_path, df_src, df_tgt, src_map, 
                mapping, tgt_field_to_col, self.tgt_header_row + 1, overwrite_all)
            engine = "Excel COM（批量优化）"
        except Exception as e1:
            try:
                total_found, total_write = openpyxl_write_and_save_optimized(
                    self.tgt_path, self.tgt_sheet, out_path, df_src, df_tgt, src_map, 
                    mapping, tgt_field_to_col, self.tgt_header_row + 1, overwrite_all)
                engine = "openpyxl（Office365兼容）"
                
                # 【关键修复】使用'replace'模式安全地进行Pandas验证回写
                try:
                    df_verify = pd.read_excel(out_path, sheet_name=self.tgt_sheet, dtype=str).fillna('')
                    with pd.ExcelWriter(
                        out_path,
                        engine='openpyxl',
                        mode='a',
                        if_sheet_exists='replace'
                    ) as writer:
                        df_verify.to_excel(writer, sheet_name=self.tgt_sheet, index=False)
                    engine = "openpyxl（Pandas兼容性优化）"
                except Exception:
                    # 验证失败，但原文件可能仍然可用
                    pass

            except Exception as e2: 
                QMessageBox.critical(self, "错误", f"所有保存方式均失败：\n\nCOM错误：{str(e1)[:200]}...\n\nopenpyxl错误：{str(e2)[:200]}...\n\n建议：\n1. 确保目标Excel文件未被其他程序占用\n2. 检查文件权限\n3. 尝试关闭Excel程序后重试"); return
        
        QMessageBox.information(self, "完成", 
            f"匹配完成（引擎：{engine}）：\n\n"
            f"命中索引记录： {total_found}\n"
            f"共写入单元格： {total_write}\n\n"
            f"结果已导出至：\n{out_path}")

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow { background: #0f172a; color: #e5e7eb; }
            QLabel { color: #cbd5e1; font-weight: 600; }
            QGroupBox { border: 1px solid #1f2937; border-radius: 10px; margin-top: 10px; padding: 10px; color: #e5e7eb; font-weight: 600; }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 2px 6px; }
            QLineEdit, QComboBox, QSpinBox, QListWidget, QTableWidget { background: #111827; color: #e5e7eb; border: 1px solid #374151; border-radius: 8px; padding: 6px; }
            QComboBox::drop-down { border: none; }
            QPushButton { background: #2563eb; color: white; border: none; border-radius: 8px; padding: 8px 12px; font-weight: 600; }
            QPushButton:hover { background: #1d4ed8; }
            QPushButton:disabled { background: #334155; color: #9ca3af; }
            QRadioButton { color: #e5e7eb; font-weight: normal; }
            QHeaderView::section { background: #0b1220; color: #cbd5e1; padding: 6px; border: none; }
            QTableWidget { gridline-color: #374151; }
            QTableWidget::item { padding-left: 5px; }
        """)

def main():
    app = QApplication(sys.argv)
    w = MapperUI()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
