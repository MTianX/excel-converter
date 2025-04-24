# Excel转CSV工具

这是一个功能强大的Excel转CSV转换工具，支持多表格转换、列映射、数据清洗和关键字替换等功能。

## 功能特点

- 图形用户界面，操作简单直观
- 支持多个工作表（Sheet）同时转换
- 支持自定义列名映射
- 支持数据清洗（去重、去空格等）
- 支持关键字模糊匹配和替换
- 完整的日志记录系统

## 安装依赖

```bash
pip install -r requirements.txt
```

## 配置文件说明

配置文件（config.ini）包含以下几个主要部分：

### 1. Sheet映射配置 [SheetMapping]

指定需要转换的Excel工作表及其对应的输出CSV文件名：

```ini
[SheetMapping]
Sheet表名1 = dig.csv
Sheet表名2 = ana.csv
```

### 2. 列映射配置 [Sheet名_ColumnMapping]

为每个工作表配置列名映射规则：

```ini
[Sheet表名_ColumnMapping]
原列名 = 新列名
名称 = 描述
系数 = 系数
```

### 3. 输出列配置 [Sheet名_OutputColumns]

指定每个工作表转换后CSV文件的列顺序：

```ini
[Sheet名_OutputColumns]
columns = 列名1,列名2,列名3
```

### 4. 数据类型配置 [DataType]

强制指定特定列的数据类型，避免类型转换错误：

```ini
[DataType]
备注 = str
```

### 5. 关键字模糊映射 [KeywordFuzzyMapping]

配置关键字搜索和替换规则：

```ini
[KeywordFuzzyMapping]
列名_关键字 = 目标列名1:替换值1,目标列名2:替换值2
```

格式说明：
- 列名_关键字：指定要搜索的列和匹配的关键字
- 目标列名:替换值：指定匹配成功后要在哪个列中填入什么值

## 使用示例

1. 配置文件示例：
```ini
[SheetMapping]
Sheet表名1 = dig.csv

[Sheet表名1_ColumnMapping]
名称 = 描述
优先级 = 告警优先级

[Sheet表名1_OutputColumns]
columns = 列名1,列名2,列名3

[KeywordFuzzyMapping]
名称_关键字1 = 输出列名1:值1,输出列名2:值2
```

2. 启动程序：
```bash
python excel_to_csv_gui.py
```

## 注意事项

1. 配置文件必须使用UTF-8编码
2. 列名映射中的原列名必须与Excel文件中的列名完全匹配
3. 关键字模糊匹配支持正则表达式
4. 确保输出列配置中包含所有需要的列名 