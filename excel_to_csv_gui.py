import sys
import os
import configparser
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, 
    QWidget, QCheckBox, QHBoxLayout, QTextEdit
)
from PyQt5.QtCore import Qt
import pandas as pd
from data_cleaner import DataCleaner
import logging
import traceback

def get_config_path(config_name="config.ini"):
    """动态获取配置文件路径（兼容打包环境）"""
    if getattr(sys, 'frozen', False):
        # 打包后：配置文件在EXE同级目录
        base_dir = os.path.dirname(sys.executable)
    else:
        # 开发环境：配置文件在脚本所在目录
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, config_name)


class ExcelToCSVApp(QMainWindow):
    """带数据清洗和列映射功能的Excel转CSV工具"""
    
    def __init__(self):
        # 1. 父类初始化 (必须第一个调用)
        super().__init__()
        
        # 2. 日志系统初始化
        self.setup_logging()
        
        # 3. 配置系统初始化
        self.config = configparser.ConfigParser(strict=False)
        self.config_file = get_config_path("config.ini")  # 使用动态路径
        self.load_config()
        
        # 4. 业务类初始化
        self.cleaner = DataCleaner(self.config)
        
        # 5. UI初始化
        self.setup_ui()
        self.setWindowTitle("Excel转CSV工具")
        self.setGeometry(100, 100, 800, 600)


    def setup_logging(self):
        """独立的日志初始化方法"""
        if not logging.getLogger().hasHandlers():
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler('excel_converter.log', encoding='utf-8'),
                    logging.StreamHandler(sys.stdout)
                ]
            )
        self.logger = logging.getLogger(__name__)

    def setup_ui(self):
        """初始化用户界面"""
        main_layout = QVBoxLayout()
        
        # 配置文件显示
        self.config_label = QLabel(f"配置文件: {os.path.abspath(self.config_file)}")
        main_layout.addWidget(self.config_label)
        
        # 文件选择部分
        main_layout.addWidget(QLabel("Excel文件路径:"))
        self.excel_entry = QLineEdit()
        self.excel_button = QPushButton("选择Excel文件")
        main_layout.addWidget(self.excel_entry)
        main_layout.addWidget(self.excel_button)
        
        # 输出目录部分
        main_layout.addWidget(QLabel("输出目录(可选，默认为Excel所在目录):"))
        self.output_entry = QLineEdit()
        self.output_button = QPushButton("选择输出目录")
        main_layout.addWidget(self.output_entry)
        main_layout.addWidget(self.output_button)
        
        # 列映射配置显示
        self.column_mapping_label = QLabel("当前列映射配置:")
        main_layout.addWidget(self.column_mapping_label)
        
        self.column_mapping_text = QTextEdit()
        self.column_mapping_text.setReadOnly(True)
        self.column_mapping_text.setMaximumHeight(100)
        main_layout.addWidget(self.column_mapping_text)
        
        # 在 setup_ui() 中添加模糊匹配选项
        self.fuzzy_mapping_check = QCheckBox("启用模糊关键字替换")
        self.fuzzy_mapping_check.setChecked(True)
        main_layout.addWidget(self.fuzzy_mapping_check)
        
        # 数据清洗选项
        main_layout.addWidget(QLabel("\n数据清洗选项:"))
        
        self.remove_duplicates_check = QCheckBox("删除重复行")
        self.remove_duplicates_check.setChecked(False)
        main_layout.addWidget(self.remove_duplicates_check)
        
        self.remove_empty_rows_check = QCheckBox("删除空行")
        self.remove_empty_rows_check.setChecked(False)
        main_layout.addWidget(self.remove_empty_rows_check)
        
        self.trim_spaces_check = QCheckBox("去除首尾空格")
        self.trim_spaces_check.setChecked(True)
        main_layout.addWidget(self.trim_spaces_check)
        
        # 填充空值选项
        fill_na_layout = QHBoxLayout()
        self.fill_na_check = QCheckBox("填充空值:")
        self.fill_na_check.setChecked(True)
        self.fill_na_entry = QLineEdit("NA")
        self.fill_na_entry.setMaximumWidth(100)
        fill_na_layout.addWidget(self.fill_na_check)
        fill_na_layout.addWidget(self.fill_na_entry)
        fill_na_layout.addStretch()
        main_layout.addLayout(fill_na_layout)
        
        # 转换按钮
        self.convert_button = QPushButton("开始转换")
        main_layout.addWidget(self.convert_button, alignment=Qt.AlignCenter)
        
        # 设置主窗口中心组件
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        
        # 连接信号槽
        self.connect_signals()
    
    def connect_signals(self):
        """连接所有信号槽"""
        self.excel_button.clicked.connect(self.select_excel_file)
        self.output_button.clicked.connect(self.select_output_dir)
        self.convert_button.clicked.connect(self.convert_to_csv)
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_entry.setText(file_path)
            self.output_entry.setText(os.path.dirname(file_path))
    
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_entry.setText(dir_path)
    
    def load_config(self):
        try:
            # 使用动态路径
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config.read_file(f)
            
            # 新增：打印所有加载的配置段
            self.logger.info("已加载配置段: %s", self.config.sections())
            
            # 新增：检查关键配置段是否存在
            for sheet in self.config['SheetMapping']:
                column_section = f"{sheet}_ColumnMapping"
                if column_section not in self.config:
                    self.logger.warning(f"缺少配置段: {column_section}")
                    
        except Exception as e:
            self.logger.error(f"配置加载失败: {str(e)}")
            self.create_default_config()
        
    def update_column_mapping_display(self):
        """显示所有映射配置（支持新旧格式）"""
        text = ["=== 配置预览 ==="]
        
        # 显示Sheet映射
        if 'SheetMapping' in self.config:
            text.append("\n[Sheet映射]")
            text.extend(f"  {k:<15} → {v}" for k,v in self.config['SheetMapping'].items())
        
        # 显示列映射
        column_sections = [s for s in self.config.sections() if s.endswith('_ColumnMapping')]
        for section in column_sections:
            text.append(f"\n[{section}]")
            for key in self.config[section]:
                if '_' in key:
                    col, pattern = key.split('_', 1)
                    text.append(f"  {col:<12} [匹配: {pattern:<8}] → {self.config[section][key]}")
                else:
                    text.append(f"  {key:<12} → {self.config[section][key]}")
        
        # 显示关键字映射
        if 'KeywordFuzzyMapping' in self.config:
            text.append("\n[关键字模糊映射]")
            for key in self.config['KeywordFuzzyMapping']:
                if '_' in key:
                    col, pattern = key.split('_', 1)
                    text.append(f"  {col:<12} [匹配: {pattern:<8}] → {self.config['KeywordFuzzyMapping'][key]}")
        
        self.column_mapping_text.setPlainText('\n'.join(text))
        
    def create_default_config(self):
        """创建默认配置文件"""
        self.config['SheetMapping'] = {
            'Sheet1': 'output.csv'
        }
        self.config['KeywordFuzzyMapping'] = {
            '名称_测试': 'test'
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            self.logger.info(f"已创建默认配置文件: {os.path.abspath(self.config_file)}")
        except Exception as e:
            self.logger.error(f"创建配置文件失败: {str(e)}")
            raise
    
    def get_clean_options(self) -> dict:
        """获取清洗选项配置"""
        return {
        'apply_fuzzy_mapping': self.fuzzy_mapping_check.isChecked(),
        'trim_spaces': self.trim_spaces_check.isChecked(),
        'remove_empty_rows': self.remove_empty_rows_check.isChecked(),
        'remove_duplicates': self.remove_duplicates_check.isChecked(),
        'fill_na': self.fill_na_check.isChecked(),
        'fill_na_value': self.fill_na_entry.text()
         }
    
    def convert_to_csv(self):
        """执行转换操作"""
        excel_file = self.excel_entry.text()
        if not excel_file:
            QMessageBox.critical(self, "错误", "请先选择 Excel 文件！")
            return
        
        output_dir = self.output_entry.text() or os.path.dirname(excel_file)
        clean_options = self.get_clean_options()
        
        try:
            success_count = self.process_excel_file(excel_file, output_dir, clean_options)
            self.show_conversion_result(success_count, output_dir)
        except Exception as e:
            self.handle_conversion_error(e)
    
    def process_excel_file(self, excel_file: str, output_dir: str, clean_options: dict = None) -> int:
        if clean_options is None:
            clean_options = {}

        success_count = 0

        with pd.ExcelFile(excel_file, engine='openpyxl') as xls:
            for sheet_name, output_name in self.config['SheetMapping'].items():
                if sheet_name in xls.sheet_names:
                    # 读取为原始数据（不强制类型转换）
                    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)  # 保持为字符串类型

                    # 数据清洗
                    df = self.cleaner.clean_data(df, clean_options)

                    # 按配置文件进行类型转换（清洗完成后）
                    if 'DataType' in self.config:
                        for col, dtype in self.config['DataType'].items():
                            if col in df.columns:
                                try:
                                    df[col] = df[col].astype(dtype.lower())
                                except Exception as e:
                                    self.logger.warning(f"列 '{col}' 类型转换失败: {str(e)}")

                    # 列映射与重组
                    column_section = f"{sheet_name}_ColumnMapping"
                    column_mapping = dict(self.config[column_section]) if column_section in self.config else None
                    df = self.cleaner.clean_and_filter_columns(df, sheet_name, column_mapping)

                    # 保存处理后的数据
                    output_path = os.path.join(output_dir, output_name)
                    df.to_csv(output_path, index=False, encoding='utf-8-sig')
                    success_count += 1

        return success_count
    
    def show_conversion_result(self, success_count: int, output_dir: str):
        """显示转换结果"""
        if success_count > 0:
            QMessageBox.information(
                self, "成功", 
                f"转换完成！成功转换 {success_count} 个Sheet。\n"
                f"输出目录: {output_dir}"
            )
        else:
            QMessageBox.warning(
                self, "警告",
                "没有找到匹配的Sheet名称，请检查配置文件！"
            )
    
    def handle_conversion_error(self, error: Exception):
        """处理转换错误"""
        # 输出详细traceback到日志
        self.logger.error("转换失败详细信息：\n" + traceback.format_exc())
        QMessageBox.critical(
            self, "错误", 
            f"转换失败：{str(error)}\n"
            f"请确保Excel文件未被其他程序占用，且配置正确。\n详细错误信息已写入日志文件。"
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = ExcelToCSVApp()
    window.show()
    sys.exit(app.exec_())