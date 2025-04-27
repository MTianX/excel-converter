import sys
import os
import configparser
import pandas as pd
import re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, 
    QWidget, QTextEdit, QHBoxLayout
)
from PyQt5.QtCore import Qt
import logging

class ConfigMaintainer(QMainWindow):
    """配置维护工具"""
    
    def __init__(self):
        super().__init__()
        
        # 初始化日志
        self.setup_logging()
        
        # 初始化配置
        self.config = configparser.ConfigParser(strict=False)
        self.config_file = self.get_config_path()
        
        # 初始化UI
        self.setup_ui()
        self.setWindowTitle("配置维护工具")
        self.setGeometry(100, 100, 1000, 800)
        
    def setup_logging(self):
        """初始化日志系统"""
        if not logging.getLogger().hasHandlers():
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler('config_maintainer.log', encoding='utf-8'),
                    logging.StreamHandler(sys.stdout)
                ]
            )
        self.logger = logging.getLogger(__name__)
        
    def get_config_path(self):
        """获取配置文件路径"""
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "config.ini")
        
    def setup_ui(self):
        """初始化用户界面"""
        main_layout = QVBoxLayout()
        
        # 配置文件选择
        config_layout = QHBoxLayout()
        self.config_label = QLabel("配置文件:")
        self.config_entry = QLineEdit(self.config_file)
        self.config_entry.setReadOnly(True)
        self.config_button = QPushButton("选择配置文件")
        config_layout.addWidget(self.config_label)
        config_layout.addWidget(self.config_entry)
        config_layout.addWidget(self.config_button)
        main_layout.addLayout(config_layout)
        
        # 文件夹选择
        folder_layout = QHBoxLayout()
        self.folder_label = QLabel("扫描文件夹:")
        self.folder_entry = QLineEdit()
        self.folder_button = QPushButton("选择文件夹")
        folder_layout.addWidget(self.folder_label)
        folder_layout.addWidget(self.folder_entry)
        folder_layout.addWidget(self.folder_button)
        main_layout.addLayout(folder_layout)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        self.export_button = QPushButton("导出配置表格")
        self.update_button = QPushButton("更新配置")
        self.scan_button = QPushButton("扫描文件夹")
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.update_button)
        button_layout.addWidget(self.scan_button)
        main_layout.addLayout(button_layout)
        
        # 日志显示
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)
        
        # 设置主窗口中心组件
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        
        # 连接信号槽
        self.connect_signals()
        
    def connect_signals(self):
        """连接信号槽"""
        self.config_button.clicked.connect(self.select_config_file)
        self.folder_button.clicked.connect(self.select_folder)
        self.export_button.clicked.connect(self.export_config_table)
        self.update_button.clicked.connect(self.update_config)
        self.scan_button.clicked.connect(self.scan_folder)
        
    def select_config_file(self):
        """选择配置文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择配置文件", "", "INI Files (*.ini)"
        )
        if file_path:
            self.config_file = file_path
            self.config_entry.setText(file_path)
            self.load_config()  # 选择后自动加载配置
            
    def select_folder(self):
        """选择扫描文件夹"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择扫描文件夹")
        if dir_path:
            self.folder_entry.setText(dir_path)
            
    def load_config(self):
        """加载配置文件"""
        try:
            if not os.path.exists(self.config_file):
                self.create_default_config()
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config.read_file(f)
            self.log_message("配置加载成功")
            self.auto_export_config_table()  # 自动导出配置表格
        except Exception as e:
            self.log_message(f"配置加载失败: {str(e)}", "ERROR")
            
    def create_default_config(self):
        """创建默认配置文件"""
        self.config['SheetMapping'] = {
            'Sheet表名1': 'dig.csv',
            'Sheet表名2': 'ana.csv'
        }
        
        self.config['Sheet表名1_ColumnMapping'] = {
            '名称': '描述',
            '优先级': '告警优先级'
        }
        
        self.config['Sheet表名2_ColumnMapping'] = {
            '名称': '描述',
            '系数': '系数'
        }
        
        self.config['Sheet表名1_OutputColumns'] = {
            'columns': '列名1,列名2,列名3'
        }
        
        self.config['Sheet表名2_OutputColumns'] = {
            'columns': '列名1,列名2,列名3'
        }
        
        self.config['DataType'] = {
            '备注': 'str'
        }
        
        self.config['KeywordFuzzyMapping'] = {
            '名称_测试关键字': '输出列名1:值1,输出列名2:值2'
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            self.log_message(f"已创建默认配置文件: {os.path.abspath(self.config_file)}")
        except Exception as e:
            self.log_message(f"创建配置文件失败: {str(e)}", "ERROR")
            raise
            
    def export_config_table(self):
        """导出配置表格"""
        try:
            if 'KeywordFuzzyMapping' not in self.config:
                raise ValueError("配置文件中没有KeywordFuzzyMapping部分")
                
            # 创建DataFrame
            data = []
            for key, value in self.config['KeywordFuzzyMapping'].items():
                if '_' in key:
                    col, pattern = key.split('_', 1)
                    # 解析命名规则字符串
                    rule_dict = {}
                    for rule in value.split(','):
                        if ':' in rule:
                            rule_col, rule_val = rule.split(':', 1)
                            rule_dict[rule_col.strip()] = rule_val.strip()
                    
                    # 提取命名规则中的数字（最后一个值）
                    rules = value.split(',')
                    naming_rule = ''
                    if rules:
                        last_rule = rules[-1]
                        if ':' in last_rule:
                            naming_rule = last_rule.split(':')[-1]
                    
                    row = {
                        '原列名称': col,
                        '原列描述': pattern,
                        '描述': rule_dict.get('描述', ''),
                        '量测类型': rule_dict.get('量测类型', ''),
                        '系数': rule_dict.get('系数', ''),
                        '告警优先级': rule_dict.get('告警优先级', ''),
                        '命名规则': naming_rule
                    }
                    data.append(row)
                    
            df = pd.DataFrame(data)
            df = df.sort_values('描述')  # 按描述排序
            
            # 保存为CSV
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存配置表格", "", "CSV Files (*.csv)"
            )
            if save_path:
                df.to_csv(save_path, index=False, encoding='utf-8-sig')
                self.log_message(f"配置表格已保存到: {save_path}")
                
        except Exception as e:
            self.log_message(f"导出配置表格失败: {str(e)}", "ERROR")
            
    def update_config(self):
        """更新配置文件"""
        try:
            # 读取配置表格.csv，自动兼容utf-8和gbk
            config_table_path = os.path.join(os.path.dirname(__file__), '配置表格.csv')
            if not os.path.exists(config_table_path):
                self.log_message("未找到配置表格.csv，请先导出配置表格！", "ERROR")
                return

            try:
                df = pd.read_csv(config_table_path, encoding='utf-8')
            except Exception:
                df = pd.read_csv(config_table_path, encoding='gbk')

            # 更新配置
            if 'KeywordFuzzyMapping' not in self.config:
                self.config['KeywordFuzzyMapping'] = {}

            self.config['KeywordFuzzyMapping'].clear()
            for _, row in df.iterrows():
                # 构建命名规则字符串
                rules = []
                if row['描述'] and not pd.isna(row['描述']):
                    rules.append(f"描述:{row['描述']}")
                if row['量测类型'] and not pd.isna(row['量测类型']):
                    rules.append(f"量测类型:{row['量测类型']}")
                if row['系数'] and not pd.isna(row['系数']):
                    try:
                        coef = int(float(row['系数']))
                        rules.append(f"系数:{coef}")
                    except Exception:
                        pass
                if row['告警优先级'] and not pd.isna(row['告警优先级']):
                    try:
                        alarm = int(float(row['告警优先级']))
                        rules.append(f"告警优先级:{alarm}")
                    except Exception:
                        pass
                if row['命名规则'] and not pd.isna(row['命名规则']):
                    rules.append(f"命名规则:{row['命名规则']}")

                key = f"{row['原列名称']}_{row['原列描述']}"
                value = ','.join(rules)
                self.config['KeywordFuzzyMapping'][key] = value

            # 保存配置
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)

            self.log_message("配置文件更新成功")

        except Exception as e:
            self.log_message(f"更新配置文件失败: {str(e)}", "ERROR")
            
    def scan_folder(self):
        """扫描文件夹"""
        try:
            folder_path = self.folder_entry.text()
            if not folder_path:
                raise ValueError("请先选择扫描文件夹")

            config_table_path = os.path.join(os.path.dirname(__file__), '配置表格.csv')
            if not os.path.exists(config_table_path):
                self.log_message("未找到配置表格.csv，请先导出配置表格！", "ERROR")
                return

            try:
                df = pd.read_csv(config_table_path, encoding='utf-8')
            except Exception:
                df = pd.read_csv(config_table_path, encoding='gbk')

            existing_descriptions = set(df['描述'].tolist())
            new_rows = []

            # 遍历所有csv文件
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.csv'):
                        file_path = os.path.join(root, file)
                        self.log_message(f"扫描文件: {file_path}")
                        # 优先尝试utf-8，失败后尝试gbk
                        file_df = None
                        try:
                            file_df = pd.read_csv(file_path, encoding='utf-8')
                        except Exception as e_utf8:
                            try:
                                file_df = pd.read_csv(file_path, encoding='gbk')
                                self.log_message(f"文件用gbk编码成功读取: {file_path}")
                            except Exception as e_gbk:
                                self.log_message(f"读取文件失败: {file_path}, 错误: {e_gbk}", "ERROR")
                                continue
                        if file_df is not None and '描述' in file_df.columns:
                            for desc in file_df['描述'].unique():
                                if desc not in existing_descriptions:
                                    self.log_message(f"发现新描述: {desc}")
                                    # 取第一个匹配该描述的行
                                    csv_row = file_df[file_df['描述'] == desc].iloc[0]
                                    def safe_get(col):
                                        return csv_row[col] if col in csv_row and not pd.isna(csv_row[col]) else ''
                                    new_row = {
                                        '原列名称': '名称',
                                        '原列描述': desc,
                                        '描述': desc,
                                        '量测类型': safe_get('量测类型'),
                                        '系数':  safe_get('系数') or '1',
                                        '告警优先级': safe_get('告警优先级'),
                                        '命名规则': safe_get('命名规则') or '0'
                                    }
                                    # 如果csv没有这些列，再用配置表格补全
                                    if not new_row['量测类型'] or not new_row['系数'] or not new_row['告警优先级'] or not new_row['命名规则']:
                                        match_row = df[df['原列描述'] == desc]
                                        if match_row.empty:
                                            match_row = df[df['描述'] == desc]
                                        if not match_row.empty:
                                            row_data = match_row.iloc[0]
                                            if not new_row['量测类型']:
                                                new_row['量测类型'] = row_data.get('量测类型', '')
                                            if not new_row['系数']:
                                                new_row['系数'] = row_data.get('系数', '1')
                                            if not new_row['告警优先级']:
                                                new_row['告警优先级'] = row_data.get('告警优先级', '')
                                            if not new_row['命名规则']:
                                                new_row['命名规则'] = row_data.get('命名规则', '0')
                                    new_rows.append(new_row)
                                    existing_descriptions.add(desc)

            # 合并新行
            if new_rows:
                df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

            # 保存更新后的配置表格
            df = df.sort_values('描述')
            save_path = config_table_path  # 直接覆盖原配置表格
            df.to_csv(save_path, index=False, encoding='utf-8-sig')
            self.log_message(f"更新后的配置表格已保存到: {save_path}")

        except Exception as e:
            self.log_message(f"扫描文件夹失败: {str(e)}", "ERROR")
            
    def auto_export_config_table(self):
        """自动导出配置表格到当前目录，文件名为配置表格.csv"""
        try:
            if 'KeywordFuzzyMapping' not in self.config:
                raise ValueError("配置文件中没有KeywordFuzzyMapping部分")
            # 创建DataFrame
            data = []
            for key, value in self.config['KeywordFuzzyMapping'].items():
                if '_' in key:
                    col, pattern = key.split('_', 1)
                    # 解析命名规则字符串
                    rule_dict = {}
                    for rule in value.split(','):
                        if ':' in rule:
                            rule_col, rule_val = rule.split(':', 1)
                            rule_dict[rule_col.strip()] = rule_val.strip()
                    # 提取命名规则中的数字（最后一个值）
                    rules = value.split(',')
                    naming_rule = ''
                    if rules:
                        last_rule = rules[-1]
                        if ':' in last_rule:
                            naming_rule = last_rule.split(':')[-1]
                    row = {
                        '原列名称': col,
                        '原列描述': pattern,
                        '描述': rule_dict.get('描述', ''),
                        '量测类型': rule_dict.get('量测类型', ''),
                        '系数': rule_dict.get('系数', '1'),
                        '告警优先级': rule_dict.get('告警优先级', ''),
                        '命名规则': naming_rule
                    }
                    data.append(row)
            df = pd.DataFrame(data)
            df = df.sort_values('描述')  # 按描述排序
            save_path = os.path.join(os.path.dirname(__file__), '配置表格.csv')
            df.to_csv(save_path, index=False, encoding='utf-8-sig')
            self.log_message(f"配置表格已自动导出到: {save_path}")
        except Exception as e:
            self.log_message(f"自动导出配置表格失败: {str(e)}", "ERROR")
            
    def log_message(self, message, level="INFO"):
        """记录日志消息"""
        self.log_text.append(f"[{level}] {message}")
        if level == "ERROR":
            self.logger.error(message)
        else:
            self.logger.info(message)
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConfigMaintainer()
    window.show()
    sys.exit(app.exec_()) 