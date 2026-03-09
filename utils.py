#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/JSON转换工具共享模块
包含两个转换工具共用的功能函数
"""

import os
import json
import datetime
import pandas as pd
import warnings

# 忽略pandas的一些警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


# 支持的文件扩展名
EXCEL_EXTENSIONS = ['.xlsx', '.xls', '.xlsm']
CSV_EXTENSIONS = ['.csv']
JSON_EXTENSIONS = ['.json']


def print_header(title):
    """
    打印格式化的标题头
    
    Args:
        title: 标题文本
    """
    print(f"=" * 60)
    print(f"          {title}         ")
    print(f"=" * 60)


def print_success(message):
    """
    打印成功消息
    
    Args:
        message: 消息内容
    """
    print(f"✓ {message}")


def print_warning(message):
    """
    打印警告消息
    
    Args:
        message: 消息内容
    """
    print(f"⚠ 警告：{message}")


def print_error(message):
    """
    打印错误消息
    
    Args:
        message: 消息内容
    """
    print(f"✗ 错误：{message}")


def get_unique_filename(file_path):
    """
    获取不冲突的文件名，如果文件已存在则添加数字后缀
    
    Args:
        file_path: 原始文件路径
        
    Returns:
        str: 不冲突的文件名
    """
    if not os.path.exists(file_path):
        return file_path
    
    # 分离文件名和扩展名
    base_path, ext = os.path.splitext(file_path)
    counter = 1
    
    # 循环添加数字后缀直到找到不存在的文件名
    new_file_path = f"{base_path}({counter}){ext}"
    while os.path.exists(new_file_path):
        counter += 1
        new_file_path = f"{base_path}({counter}){ext}"
    
    return new_file_path


def ensure_output_directory(file_path):
    """
    确保输出目录存在，如果不存在则创建
    
    Args:
        file_path: 文件路径
    """
    output_dir = os.path.dirname(file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        print_success(f"创建输出目录：{output_dir}")


def is_timestamp(value):
    """
    判断一个值是否为有效的时间戳
    
    Args:
        value: 要判断的值
        
    Returns:
        bool: 如果是有效的时间戳返回True，否则返回False
    """
    if isinstance(value, (int, float)):
        # 检查时间戳范围（1970年到2100年之间）
        return 0 < value < 4102444800000  # 2100-01-01的毫秒时间戳
    elif isinstance(value, str):
        # 尝试将字符串转换为数字
        try:
            num_value = float(value)
            return 0 < num_value < 4102444800000
        except ValueError:
            return False
    return False


def convert_timestamp_to_datetime(value, date_format='%Y-%m-%d %H:%M:%S'):
    """
    将时间戳转换为指定格式的日期时间字符串
    
    Args:
        value: 时间戳值（数字或字符串）
        date_format: 日期格式，默认为'YYYY-MM-DD HH:MM:SS'
        
    Returns:
        str: 格式化后的日期时间字符串，转换失败返回原值
    """
    try:
        # 将值转换为数字
        if isinstance(value, str):
            timestamp = float(value)
        else:
            timestamp = float(value)
            
        # 检查时间戳长度，判断是秒还是毫秒
        if 0 < timestamp < 2147483647:  # 小于2^31-1，可能是秒级时间戳
            timestamp *= 1000
            
        # 转换为datetime对象
        dt = datetime.datetime.fromtimestamp(timestamp / 1000.0)
        
        # 返回格式化的日期时间字符串
        return dt.strftime(date_format)
    except (ValueError, TypeError, OSError):
        # 转换失败，返回原值
        return value


def convert_date_to_timestamp(value):
    """
    将日期值转换为时间戳（毫秒级）
    
    Args:
        value: 输入的值，可以是字符串、日期对象或数值
    
    Returns:
        int: 毫秒级时间戳（确保返回整数类型）
    """
    # 空值或None直接返回
    if pd.isna(value) or value is None:
        return ""
    
    # 处理pandas的Timestamp对象
    if isinstance(value, pd.Timestamp):
        return int(value.timestamp() * 1000)
    
    # 处理datetime对象
    if isinstance(value, (datetime.datetime, datetime.date)):
        # 如果是date对象，转换为datetime对象（时间设为00:00:00）
        if isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
            value = datetime.datetime.combine(value, datetime.time())
        # 使用timestamp方法获取时间戳
        return int(value.timestamp() * 1000)
    
    # 处理数值类型
    if isinstance(value, (int, float)):
        # 如果看起来已经是时间戳（10位或13位数字），直接返回整数形式
        if 10**9 <= value <= 10**14:  # 10位到14位数字，覆盖秒级和毫秒级时间戳
            return int(value)
        
        try:
            # 尝试将数值转换为日期（Excel的日期格式）
            # Excel的日期从1899-12-30开始计数
            date = pd.to_datetime(value, unit='D', origin='1899-12-30')
            return int(date.timestamp() * 1000)
        except:
            return int(value)  # 转换为整数而不是保持浮点数
    
    # 处理字符串类型
    if isinstance(value, str):
        # 去除字符串两端的空白字符
        value = value.strip()
        
        # 检查是否为纯数字字符串（可能是时间戳）
        if value.isdigit():
            # 将字符串数字转换为整数
            return int(value)
        
        # 尝试将字符串转换为数字
        try:
            num_value = float(value)
            # 检查是否为整数
            return int(num_value)
        except ValueError:
            pass
        
        # 尝试pandas的自动日期解析（更强大的解析能力）
        try:
            date = pd.to_datetime(value)
            return int(date.timestamp() * 1000)
        except:
            pass
            
        # 尝试多种常见的日期格式
        date_formats = [
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%d/%m/%Y %H:%M:%S',
            '%d/%m/%Y',
            '%m/%d/%Y %H:%M:%S',
            '%m/%d/%Y',
            '%Y年%m月%d日 %H:%M:%S',
            '%Y年%m月%d日',
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%dT%H:%M:%SZ',
            '%Y/%m/%d %H:%M:%S',
            '%Y/%m/%d',
            '%Y-%m-%d %H:%M',
            '%Y/%m/%d %H:%M',
            '%d/%m/%Y %H:%M',
            '%m/%d/%Y %H:%M',
        ]
        
        for fmt in date_formats:
            try:
                date = datetime.datetime.strptime(value, fmt)
                return int(date.timestamp() * 1000)
            except ValueError:
                continue
    
    # 如果无法转换，返回原值
    return value


def detect_timestamp_fields(data):
    """
    自动检测数据中的时间戳字段
    
    Args:
        data: 数据列表
        
    Returns:
        list: 时间戳字段名称列表
    """
    if not data or not isinstance(data, list):
        return []
    
    # 获取所有字段名
    fields = list(data[0].keys()) if data else []
    timestamp_fields = []
    
    # 检查字段名是否暗示时间戳
    timestamp_keywords = ['timestamp', 'time', 'date', 'datetime', 'created', 'updated']
    # 排除可能是普通数值的字段名
    numeric_keywords = ['value', 'count', 'amount', 'price', 'score', 'total']
    
    for field in fields:
        field_lower = field.lower()
        
        # 如果字段名包含数值相关关键词，跳过
        if any(numeric_keyword in field_lower for numeric_keyword in numeric_keywords):
            continue
            
        # 检查字段名是否包含时间戳相关关键词
        if any(keyword in field_lower for keyword in timestamp_keywords):
            timestamp_fields.append(field)
            continue
        
        # 采样检查该字段的值是否为时间戳（更严格的检查）
        sample_values = [item[field] for item in data if field in item and item[field] is not None][:10]
        # 只有当所有值都是时间戳且样本量足够（至少3个）时才认为是时间戳字段
        if sample_values and len(sample_values) >= 3 and all(is_timestamp(val) for val in sample_values):
            # 进一步检查这些值是否看起来像时间戳（例如，值的范围合理）
            if all(1000000000000 <= float(str(val)) <= 2000000000000 for val in sample_values):  # 2001年到2033年的毫秒时间戳
                timestamp_fields.append(field)
    
    return timestamp_fields


def get_file_type(file_path):
    """
    根据文件扩展名确定文件类型
    
    Args:
        file_path: 文件路径
        
    Returns:
        str: 文件类型（'excel', 'csv', 'json' 或 'unknown'）
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext in EXCEL_EXTENSIONS:
        return 'excel'
    elif file_ext in CSV_EXTENSIONS:
        return 'csv'
    elif file_ext in JSON_EXTENSIONS:
        return 'json'
    return 'unknown'


def get_default_output_file(input_file, output_type='json'):
    """
    根据输入文件和输出类型生成默认输出文件路径
    
    Args:
        input_file: 输入文件路径
        output_type: 输出类型（'json', 'excel'）
        
    Returns:
        str: 默认输出文件路径
    """
    base_name = os.path.splitext(input_file)[0]
    if output_type == 'json':
        return f"{base_name}.json"
    else:  # excel
        return f"{base_name}.xlsx"


def parse_field_filter(filter_str):
    """
    解析字段过滤条件，如 {name=Ep,age>18}
    
    Args:
        filter_str: 过滤条件字符串
        
    Returns:
        list: 过滤条件列表，每个元素为(key, operator, value)元组
    """
    # 移除花括号
    filter_str = filter_str.strip('{}')
    
    # 分割多个条件
    conditions = []
    for condition in filter_str.split(','):
        condition = condition.strip()
        if '=' in condition:
            key, value = condition.split('=', 1)
            conditions.append((key.strip(), '=', value.strip().strip('"\'')))
        elif '>' in condition:
            key, value = condition.split('>', 1)
            conditions.append((key.strip(), '>', value.strip()))
        elif '<' in condition:
            key, value = condition.split('<', 1)
            conditions.append((key.strip(), '<', value.strip()))
    
    return conditions


def extract_nested_value(data, path):
    """
    从嵌套数据中提取值，支持列表过滤
    
    Args:
        data: 输入数据
        path: 字段路径，支持嵌套和过滤，如 'attrValues[{name=Ep}].value'
        
    Returns:
        提取的值，如果路径不存在或提取失败返回None
    """
    import re
    
    # 匹配列表访问和过滤条件
    list_pattern = r'([\w.]+)\[(\{[^\}]*\})\]'
    
    # 处理列表访问和过滤
    while '[' in path:
        # 查找第一个列表访问模式
        match = re.search(list_pattern, path)
        if not match:
            break
        
        # 提取路径和过滤条件
        base_path = match.group(1)
        filter_str = match.group(2)
        
        # 提取基础路径的值
        current_data = extract_nested_value(data, base_path)
        if not isinstance(current_data, list):
            return None
        
        # 解析过滤条件
        filters = parse_field_filter(filter_str)
        
        # 应用过滤条件
        filtered_items = []
        for item in current_data:
            if isinstance(item, dict):
                match_all = True
                for key, operator, value in filters:
                    if key not in item:
                        match_all = False
                        break
                    
                    item_value = str(item[key])
                    # 尝试转换为数字进行比较
                    try:
                        item_value_num = float(item_value)
                        value_num = float(value)
                        if operator == '=' and item_value_num != value_num:
                            match_all = False
                            break
                        elif operator == '>' and item_value_num <= value_num:
                            match_all = False
                            break
                        elif operator == '<' and item_value_num >= value_num:
                            match_all = False
                            break
                    except (ValueError, TypeError):
                        # 字符串比较
                        if operator == '=' and item_value != value:
                            match_all = False
                            break
                
                if match_all:
                    filtered_items.append(item)
        
        # 如果找到匹配项，使用第一个匹配项
        if filtered_items:
            # 更新数据为第一个匹配项
            data = filtered_items[0]
            # 更新路径为剩余部分
            path = path[match.end():].lstrip('.')
        else:
            return None
    
    # 处理简单的点号分隔路径
    if path:
        parts = path.split('.')
        for part in parts:
            if isinstance(data, dict) and part in data:
                data = data[part]
            else:
                return None
    
    return data


def parse_field_mapping(field_mapping_str):
    """
    解析字段映射字符串，如 'value:attrValues[{name=Ep}].value'
    
    Args:
        field_mapping_str: 字段映射字符串
        
    Returns:
        tuple: (excel_column_name, json_field_path)
    """
    if ':' in field_mapping_str:
        column_name, field_path = field_mapping_str.split(':', 1)
        return column_name.strip(), field_path.strip()
    else:
        # 如果没有冒号，列名和字段名相同
        field_path = field_mapping_str.strip()
        return field_path, field_path


def filter_fields(data, ignore_fields=None):
    """
    根据ignore参数过滤数据中的字段
    
    Args:
        data: 原始数据列表
        ignore_fields: 要忽略的字段列表，或逗号分隔的字段字符串
        
    Returns:
        list: 过滤后的数据列表，不包含被忽略的字段
    """
    if not data:
        return data
    
    # 处理ignore_fields参数
    ignore_list = []
    if ignore_fields:
        if isinstance(ignore_fields, str):
            # 将逗号分隔的字符串转换为列表
            ignore_list = [field.strip() for field in ignore_fields.split(',')]
        elif isinstance(ignore_fields, list):
            ignore_list = ignore_fields
        print(f"应用字段过滤，忽略以下字段: {ignore_list}")
    else:
        print("未指定需要忽略的字段")
        return data
    
    # 过滤每条数据
    result = []
    for i, item in enumerate(data):
        # 获取数据标识用于日志
        data_id = item.get('id', f'第{i+1}条数据')
        filtered_item = {}
        
        # 保留未被忽略的字段
        ignored_count = 0
        for key, value in item.items():
            if key not in ignore_list:
                filtered_item[key] = value
            else:
                ignored_count += 1
        
        if ignored_count > 0:
            print(f"  数据 {data_id}：已忽略 {ignored_count} 个字段")
        
        result.append(filtered_item)
    
    print(f"字段过滤完成，共处理 {len(result)} 条数据")
    return result

def apply_custom_fields(data, custom_fields):
    """
    根据自定义字段配置处理数据，保留原始字段并添加自定义字段
    
    Args:
        data: 原始数据列表
        custom_fields: 自定义字段映射列表，如 ['value:attrValues[{name=Ep}].value']
        
    Returns:
        list: 处理后的数据列表，包含原始字段和自定义字段
    """
    if not data or not custom_fields:
        print_warning("未提供数据或自定义字段，直接返回原始数据")
        return data
    
    # 解析所有字段映射
    field_mappings = [parse_field_mapping(field_str) for field_str in custom_fields]
    print(f"解析自定义字段映射: {field_mappings}")
    
    # 处理每条数据
    result = []
    print(f"开始处理 {len(data)} 条数据...")
    
    for i, item in enumerate(data):
        # 获取数据标识用于日志（使用id或name字段，如果有的话）
        data_id = item.get('id', f'第{i+1}条数据')
        data_name = item.get('name', '')
        data_info = f"{data_id}"
        if data_name:
            data_info += f" ({data_name})"
        
        print(f"处理 {data_info}...")
        
        # 复制原始数据，保留所有原始字段
        processed_item = item.copy()
        
        # 应用每个字段映射
        for column_name, field_path in field_mappings:
            print(f"  提取字段 '{column_name}' 从路径 '{field_path}'")
            # 提取值
            value = extract_nested_value(item, field_path)
            
            # 记录提取结果
            if value is None:
                print_warning(f"    警告：从路径 '{field_path}' 未提取到值")
                processed_item[column_name] = ''
            else:
                print(f"    成功提取值: {value}")
                processed_item[column_name] = value
        
        result.append(processed_item)
    
    print(f"自定义字段处理完成，共处理 {len(result)} 条数据")
    return result
