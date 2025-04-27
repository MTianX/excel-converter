import pandas as pd
import re
import logging
from typing import Dict, Any, Optional, List

# 配置日志格式
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data_cleaner.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class DataCleaner:
    def __init__(self, config: dict):  # 接收配置参数
        self.config = config
        self.logger = logging.getLogger(__name__)  # 可选：初始化日志

    def clean_data(self, df: pd.DataFrame, clean_options: dict) -> pd.DataFrame:
        # 步骤1：应用模糊关键字替换
        df = self.apply_fuzzy_mapping(df)
        
        # 步骤2：执行基础清洗（去空格、去重等）
        df = self.basic_cleaning(df, clean_options)
        
        return df

    def basic_cleaning(self, df: pd.DataFrame, options: dict) -> pd.DataFrame:
        if options.get("trim_spaces", True):
            for col in df.columns:
                # 检查列是否为字符串类型（包括经过类型转换后的列）
                if pd.api.types.is_string_dtype(df[col]):
                    try:
                        df[col] = df[col].str.strip()
                    except AttributeError:
                        self.logger.warning(f"列 '{col}' 包含非字符串数据，跳过处理")
        
        # 其他清洗选项
        if options.get("remove_empty_rows", False):
            df = self.remove_empty_rows(df)
        if options.get("remove_duplicates", False):
            df = self.remove_duplicates(df)
        if options.get("fill_na", False):
            fill_value = options.get("fill_na_value", "NA")
            df = df.fillna(fill_value)
        
        return df

    def set_config(self, config: dict):
        """允许后期更新配置"""
        self.config = config
        self.logger.info("配置已更新")

    def clean_and_filter_columns(self, df: pd.DataFrame, sheet_name: str, column_mapping: Dict[str, str] = None) -> pd.DataFrame:
        # 获取输出列配置
        sheet_output_section = f"{sheet_name}_OutputColumns"
        if sheet_output_section in self.config:
            output_columns = [col.strip() for col in self.config[sheet_output_section].get('columns', '').split(',')]
        else:
            output_columns = [...]  # 默认列
        
        # 创建空DataFrame时直接使用配置的列顺序
        output_df = pd.DataFrame(columns=output_columns)
        
        # 修正：对所有列名做str和空值判断，防止None或NaN导致strip报错
        df_columns_clean = [str(col).strip().lower() for col in df.columns if pd.notna(col)]
        
        # 处理列映射
        if column_mapping:
            for src_col, dest_col in column_mapping.items():
                src_col_clean = src_col.strip().lower()
                if src_col_clean in df_columns_clean:
                    actual_src_col = df.columns[df_columns_clean.index(src_col_clean)]
                    output_df[dest_col] = df[actual_src_col]
                else:
                    # 如果原始数据有目标列，直接保留
                    if dest_col in df.columns:
                        output_df[dest_col] = df[dest_col]
                    else:
                        output_df[dest_col] = ""
                    self.logger.warning(f"映射失败：Sheet '{sheet_name}' 中不存在列 '{src_col}'，已保留原始数据（如有）")
        self.logger.info(f"输出列顺序: {output_columns}")
        self.logger.info(f"处理后的数据列: {output_df.columns.tolist()}")
        
        # 按最终列顺序重建DataFrame
        output_df = df.reindex(columns=output_columns)
        
        return output_df

    def apply_fuzzy_mapping(self, df: pd.DataFrame) -> pd.DataFrame:
        if 'KeywordFuzzyMapping' not in self.config:
            return df

        df = df.copy()  # 不再强制转换为字符串

        for full_key, replacements in self.config['KeywordFuzzyMapping'].items():
            try:
                if '_' not in full_key:
                    self.logger.error(f"无效的模糊映射键格式: {full_key}")
                    continue

                # 解析列名和匹配模式
                src_col, pattern = full_key.split('_', 1)
                pattern = pattern.replace('*', '.*')  # 转换通配符

                if src_col not in df.columns:
                    self.logger.error(f"源列不存在: {src_col}")
                    continue

                # 仅将源列转换为字符串进行匹配
                src_series = df[src_col].astype(str)
                mask = src_series.str.contains(pattern, case=False, regex=True, na=False)
                match_count = mask.sum()

                for replacement in replacements.split(','):
                    if ':' not in replacement:
                        self.logger.error(f"无效替换规则: {replacement}")
                        continue

                    dest_col, replace_value = replacement.split(':', 1)
                    dest_col = dest_col.strip()
                    replace_value = replace_value.strip()

                    # 初始化目标列为字符串类型（如果不存在）
                    if dest_col not in df.columns:
                        df[dest_col] = ""  # 默认空字符串
                    else:
                        df[dest_col] = df[dest_col].astype(str)  # 确保目标列为字符串

                    # 应用替换
                    df.loc[mask, dest_col] = replace_value

                self.logger.info(f"✅ [{src_col}] 替换完成，命中 {match_count} 行")

            except Exception as e:
                self.logger.error(f"处理键 {full_key} 时出错: {str(e)}")

        return df

    
    
    @staticmethod
    def trim_whitespace(df: pd.DataFrame) -> pd.DataFrame:
        """去除字符串首尾空格（兼容pandas 2.1+）"""
        return df.map(lambda x: x.strip() if isinstance(x, str) else x)
    
    @staticmethod
    def remove_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
        return df.dropna(how='all')
    
    @staticmethod
    def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
        return df.drop_duplicates()
    
    @staticmethod
    def fill_missing_values(df: pd.DataFrame, fill_value: Any) -> pd.DataFrame:
        return df.fillna(fill_value)