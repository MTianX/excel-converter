import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit, QWidget, QHBoxLayout
)
from PyQt5.QtCore import Qt

class ExcelAuditTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel内容审核工具")
        self.setGeometry(100, 100, 900, 600)
        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout()

        # 文件夹选择
        folder_layout = QHBoxLayout()
        self.folder_label = QLabel("选择审核目录:")
        self.folder_entry = QLineEdit()
        self.folder_button = QPushButton("选择文件夹")
        folder_layout.addWidget(self.folder_label)
        folder_layout.addWidget(self.folder_entry)
        folder_layout.addWidget(self.folder_button)
        main_layout.addLayout(folder_layout)

        # 审核按钮
        self.audit_button = QPushButton("开始审核")
        main_layout.addWidget(self.audit_button)

        # 日志输出区
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # 信号槽
        self.folder_button.clicked.connect(self.select_folder)
        self.audit_button.clicked.connect(self.start_audit)

    def select_folder(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择审核文件夹")
        if dir_path:
            self.folder_entry.setText(dir_path)

    def start_audit(self):
        folder = self.folder_entry.text().strip()
        self.log_text.clear()
        if not folder or not os.path.isdir(folder):
            self.log_message("请选择有效的文件夹！", "ERROR")
            return
        # 审核ana.csv
        ana_path = os.path.join(folder, "ana.csv")
        if os.path.exists(ana_path):
            self.log_message(f"开始审核: {ana_path}")
            self.audit_ana_csv(ana_path)
        else:
            self.log_message("未找到 ana.csv 文件", "WARNING")
        # 审核dig.csv
        dig_path = os.path.join(folder, "dig.csv")
        if os.path.exists(dig_path):
            self.log_message(f"开始审核: {dig_path}")
            self.audit_dig_csv(dig_path)
        else:
            self.log_message("未找到 dig.csv 文件", "WARNING")

    def audit_ana_csv(self, file_path):
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except Exception:
            df = pd.read_csv(file_path, encoding='gbk')
        has_error = False
        group_cols = ['设备类型', '同类型设备号']
        # 统一空值
        for col in group_cols:
            df[col] = df[col].fillna('').astype(str).str.strip()
        for group_keys, group in df.groupby(group_cols, dropna=False):
            dev_type = group_keys[0] if group_keys[0] else '(空)'
            dev_num = group_keys[1] if group_keys[1] else '(空)'
            # 合并量测类型重复行号，细分量测类型为空的情况
            dup = group.duplicated(subset=['量测类型'], keep=False)
            if dup.any():
                has_error = True
                empty_type = (group['量测类型'].astype(str).str.strip() == '') | (group['量测类型'].astype(str).str.strip() == 'nan')
                dup_empty = dup & empty_type
                if dup_empty.any():
                    idxs = group[dup_empty].index + 2
                    idxs_str = ','.join(map(str, sorted(set(idxs))))
                    self.log_message(f"ana.csv: 设备类型={dev_type}, 同类型设备号={dev_num} 组内量测类型为空，行号: {idxs_str}", "ERROR")
                for t in group.loc[dup & ~empty_type, '量测类型'].unique():
                    idxs = group[(group['量测类型'] == t) & dup].index + 2
                    idxs_str = ','.join(map(str, sorted(set(idxs))))
                    self.log_message(f"ana.csv: 设备类型={dev_type}, 同类型设备号={dev_num} 组内量测类型“{t}”重复，行号: {idxs_str}", "ERROR")
            # 合并描述重复行号
            dup_desc = group['描述'][group.duplicated(subset=['描述'], keep=False)]
            if not dup_desc.empty:
                has_error = True
                for d in dup_desc.unique():
                    idxs = group[group['描述'] == d].index + 2
                    idxs_str = ','.join(map(str, sorted(set(idxs))))
                    self.log_message(f"ana.csv: 设备类型={dev_type}, 同类型设备号={dev_num} 组内描述“{d}”重复，行号: {idxs_str}", "ERROR")
        # 合并点号重复行号
        if '点号' in df.columns:
            dup = df.duplicated(subset=['点号'], keep=False)
            if dup.any():
                has_error = True
                idxs = df[dup].index + 2
                idxs_str = ','.join(map(str, sorted(set(idxs))))
                self.log_message(f"ana.csv: 点号重复，行号: {idxs_str}", "ERROR")
        # 合并是否控制非法行号
        if '是否控制' in df.columns:
            invalid = ~df['是否控制'].astype(str).str.strip().isin(['0', '1'])
            if invalid.any():
                has_error = True
                idxs = df[invalid].index + 2
                idxs_str = ','.join(map(str, sorted(set(idxs))))
                self.log_message(f"dig.csv: 是否控制只能为0或1，行号: {idxs_str}", "ERROR")
        if '命名规则' in df.columns:
            def is_not_zero(val):
                if pd.isna(val):
                    return True
                try:
                    return float(val) != 0
                except Exception:
                    return str(val).strip() != '0'
            wrong = df['命名规则'].apply(is_not_zero)
            if wrong.any():
                has_error = True
                for idx in df[wrong].index:
                    self.log_message(f"ana.csv: 命名规则不为0，行号: {idx+2}", "ERROR")
        if '系数' in df.columns:
            empty = df['系数'].isna() | (df['系数'].astype(str).str.strip() == '')
            if empty.any():
                has_error = True
                for idx in df[empty].index:
                    self.log_message(f"ana.csv: 系数为空，行号: {idx+2}", "ERROR")
        for idx, row in df.iterrows():
            dev, num = str(row.get('设备类型', '')).strip(), str(row.get('同类型设备号', '')).strip()
            if (dev and not num) or (not dev and num):
                has_error = True
                self.log_message(f"ana.csv: 设备类型、同类型设备号必须同时有值或同时为空，行号: {idx+2}", "ERROR")
        if not has_error:
            self.log_message("ana.csv: 审核通过，无错误。", "INFO")

    def audit_dig_csv(self, file_path):
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except Exception:
            df = pd.read_csv(file_path, encoding='gbk')
        has_error = False
        group_cols = ['设备类型', '同类型设备号']
        # 统一空值
        for col in group_cols:
            df[col] = df[col].fillna('').astype(str).str.strip()
        for group_keys, group in df.groupby(group_cols, dropna=False):
            if '分量ID' in group.columns:
                sub = group[group['分量ID'].astype(str).str.strip() == '1']
            else:
                sub = group
            # 检查量测类型重复，仅对分量ID=1的子集
            dup_types = sub['量测类型'][sub.duplicated(subset=['量测类型'], keep=False)]
            if not dup_types.empty:
                has_error = True
                for t in dup_types.unique():
                    idxs = sub[sub['量测类型'] == t].index + 2
                    idxs_str = ','.join(map(str, sorted(set(idxs))))
                    self.log_message(f"dig.csv: 设备类型={group_keys[0]}, 同类型设备号={group_keys[1]} 组内量测类型“{t}”重复，行号: {idxs_str}", "ERROR")
            # 检查描述重复，输出所有重复描述及其行号
            dup_desc = sub['描述'][sub.duplicated(subset=['描述'], keep=False)]
            if not dup_desc.empty:
                has_error = True
                dev_type = group_keys[0] if group_keys[0] else '(空)'
                dev_num = group_keys[1] if group_keys[1] else '(空)'
                for d in dup_desc.unique():
                    idxs = sub[sub['描述'] == d].index + 2
                    idxs_str = ','.join(map(str, idxs))
                    self.log_message(f"dig.csv: 设备类型={dev_type}, 同类型设备号={dev_num} 分量ID=1 组内 描述“{d}”重复，行号: {idxs_str}", "ERROR")

        if '遥信点号' in df.columns:
            dup = df.duplicated(subset=['遥信点号'], keep=False)
            if dup.any():
                has_error = True
                dup_vals = df.loc[dup, '遥信点号']
                for val in dup_vals.unique():
                    idxs = df[df['遥信点号'] == val].index + 2
                    idxs_str = ','.join(map(str, sorted(set(idxs))))
                    self.log_message(f"dig.csv: 遥信点号“{val}”重复，行号: {idxs_str}", "ERROR")
        if '命名规则' in df.columns:
            def is_not_zero(val):
                if pd.isna(val):
                    return True
                try:
                    return float(val) != 0
                except Exception:
                    return str(val).strip() != '0'
            wrong = df['命名规则'].apply(is_not_zero)
            if wrong.any():
                has_error = True
                for idx in df[wrong].index:
                    self.log_message(f"dig.csv: 命名规则不为0，行号: {idx+2}", "ERROR")
        if '告警优先级' in df.columns:
            empty = df['告警优先级'].isna() | (df['告警优先级'].astype(str).str.strip() == '')
            if empty.any():
                has_error = True
                for idx in df[empty].index:
                    self.log_message(f"dig.csv: 告警优先级为空，行号: {idx+2}", "ERROR")
        for idx, row in df.iterrows():
            dev, num = str(row.get('设备类型', '')).strip(), str(row.get('同类型设备号', '')).strip()
            if (dev and not num) or (not dev and num):
                has_error = True
                self.log_message(f"dig.csv: 设备类型、同类型设备号必须同时有值或同时为空，行号: {idx+2}", "ERROR")
        if '是否控制' in df.columns:
            invalid = ~df['是否控制'].astype(str).str.strip().isin(['0', '1'])
            if invalid.any():
                has_error = True
                for idx in df[invalid].index:
                    self.log_message(f"dig.csv: 是否控制只能为0或1，行号: {idx+2}", "ERROR")
        if '分量ID' in df.columns:
            for group_keys, group in df.groupby(group_cols, dropna=False):
                group = group.sort_index()
                id1_row = group[group['分量ID'] == 1]
                id2_row = group[group['分量ID'] == 2]
                if not id1_row.empty and not id2_row.empty:
                    id1 = id1_row.iloc[0]
                    id2 = id2_row.iloc[0]
                    for col in ['量测类型', '是否控制', '控制点号']:
                        if col in id1 and col in id2:
                            if str(id1[col]).strip() != str(id2[col]).strip():
                                has_error = True
                                row_id1 = id1.name + 2
                                row_id2 = id2.name + 2
                                self.log_message(
                                    f"dig.csv: 设备类型={group_keys[0]}, 同类型设备号={group_keys[1]} 分量ID=2的{col}与分量ID=1不一致，行号: {row_id1},{row_id2}",
                                    "ERROR"
                                )
        if '分量ID' in df.columns and '是否控制' in df.columns and '控制点号' in df.columns:
            sub = df[(df['分量ID'] == 1) & (df['是否控制'].astype(str).str.strip() == '1')]
            dup = sub.duplicated(subset=['控制点号'], keep=False)
            if dup.any():
                has_error = True
                for idx in sub[dup].index:
                    self.log_message(f"dig.csv: 分量ID=1且是否控制为1的控制点号重复，行号: {idx+2}", "ERROR")
        if not has_error:
            self.log_message("dig.csv: 审核通过，无错误。", "INFO")

    def log_message(self, message, level="INFO"):
        self.log_text.append(f"[{level}] {message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelAuditTool()
    window.show()
    sys.exit(app.exec_()) 