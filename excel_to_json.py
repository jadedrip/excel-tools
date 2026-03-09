#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/CSV转JSON工具
将Excel或CSV文件转换为JSON格式，其中文件首行作为字段名，后续每行数据转换为一个JSON对象。
"""

import os
import json
import pandas as pd
import argparse
from typing import List, Dict, Any
from datetime import datetime
import numpy as np
from utils import (print_header, print_success, print_warning, print_error,
                 get_unique_filename, ensure_output_directory, convert_date_to_timestamp,
                 filter_fields)


# convert_date_to_timestamp函数已移至utils.py

def process_timestamp_columns(df, timestamp_columns=None):
    """
    处理DataFrame中的时间戳列
    
    Args:
        df: pandas DataFrame
        timestamp_columns: 要转换为时间戳的列名列表
    
    Returns:
        处理后的DataFrame
    """
    # 创建DataFrame的副本以避免修改原始数据
    processed_df = df.copy()
    
    # 如果没有指定时间戳列，自动检测名为'timestamp'的列
    if timestamp_columns is None:
        timestamp_columns = []
        if 'timestamp' in processed_df.columns:
            timestamp_columns.append('timestamp')
            print_success("自动检测到'timestamp'列")
    else:
        # 确保timestamp_columns是列表类型
        if not isinstance(timestamp_columns, list):
            timestamp_columns = [timestamp_columns]
    
    # 处理每个时间戳列
    for col in timestamp_columns:
        if col in processed_df.columns:
            print_success(f"正在将'{col}'列转换为时间戳")
            # 添加调试信息
            sample_value = processed_df[col].iloc[0] if len(processed_df) > 0 else '空'
            print(f"  示例值：{sample_value} (类型: {type(sample_value).__name__})")
            
            # 应用转换函数到列中的每个元素
            processed_df[col] = processed_df[col].apply(convert_date_to_timestamp)
            
            # 重要：确保转换后的结果是整数类型
            # 强制将列转换为数值类型，非数值转为NaN
            processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce')
            
            # 将NaN值替换为空字符串
            processed_df[col] = processed_df[col].fillna('')
            
            # 再次显示转换后的示例值类型
            if len(processed_df) > 0:
                new_sample = processed_df[col].iloc[0]
                print(f"  转换后类型: {type(new_sample).__name__}")
        else:
            print_warning(f"指定的时间戳列'{col}'不存在于数据中")
    
    print_success("时间戳列处理完成")
    return processed_df

def excel_to_json(excel_file: str, output_file: str = None, sheet_name: str = 0, timestamp_columns: list = None, 
                  csv_delimiter: str = ',', csv_encoding: str = 'utf-8', ignore_fields: str = None, 
                  field_mapping: dict = None) -> str:
    """
    将Excel或CSV文件转换为JSON格式
    
    Args:
        excel_file: Excel或CSV文件路径
        output_file: 输出JSON文件路径，默认为原文件名+.json
        sheet_name: 要读取的工作表名称或索引，默认为第一个工作表（仅适用于Excel文件）
        timestamp_columns: 要转换为时间戳的列名列表
        csv_delimiter: CSV文件的分隔符，默认为逗号（仅适用于CSV文件）
        csv_encoding: CSV文件的编码，默认为utf-8（仅适用于CSV文件）
        field_mapping: 输入字段名到输出字段名的映射字典
    """
    print_header("Excel/CSV转JSON工具")
    
    # 检查输入文件是否存在
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"错误：找不到输入文件 '{excel_file}'")
    
    # 检查文件扩展名是否为支持的格式
    file_ext = os.path.splitext(excel_file)[1].lower()
    if file_ext not in ['.xlsx', '.xls', '.xlsm', '.csv']:
        raise ValueError(f"错误：文件 '{excel_file}' 不是支持的格式。支持的格式：.xlsx, .xls, .xlsm, .csv")
    
    # 根据文件类型选择读取方法
    try:
        print(f"开始处理文件：{excel_file}")
        if file_ext == '.csv':
            # 读取CSV文件，使用指定的分隔符和编码，默认使用第一行作为列名
            print(f"  文件类型：CSV")
            print(f"  分隔符：'{csv_delimiter}'")
            print(f"  编码：{csv_encoding}")
            df = pd.read_csv(excel_file, delimiter=csv_delimiter, encoding=csv_encoding)
        else:
            # 读取Excel文件
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # 检查是否为空数据
        if df.empty:
            print_warning(f"Excel文件 '{excel_file}' 的工作表 '{sheet_name}' 为空")
        
        print_success(f"成功读取Excel文件")
        print(f"  工作表：{sheet_name}")
        print(f"  总行数：{len(df)}")
        print(f"  字段名：{list(df.columns)}")
        
        # 数据清洗：处理NaN值，将其转换为空字符串
        df = df.fillna('')
        print_success("数据清洗完成")
        
        # 处理时间戳列
        df = process_timestamp_columns(df, timestamp_columns)
        print_success("时间戳列处理完成")
        
        # 应用字段映射（重命名字段）
        if field_mapping:
            print_success("应用字段映射...")
            df = df.rename(columns=field_mapping)
            print(f"  映射后的字段名：{list(df.columns)}")
        
        # 将DataFrame转换为字典列表（每行一个字典）
        json_data = df.to_dict('records')
        print_success(f"成功转换为JSON对象，共{len(json_data)}个对象")
        
        # 应用字段过滤
        if ignore_fields:
            print_success("应用字段过滤...")
            json_data = filter_fields(json_data, ignore_fields)
        
        # 如果未指定输出文件路径，则使用Excel文件名+.json
        if output_file is None:
            base_name = os.path.splitext(os.path.basename(excel_file))[0]
            output_file = os.path.join(os.path.dirname(excel_file), f"{base_name}.json")
        
        # 获取不冲突的文件名
        original_output = output_file
        output_file = get_unique_filename(output_file)
        if output_file != original_output:
            print_warning(f"输出文件已存在，将使用新文件名: {output_file}")
        
        # 确保输出目录存在
        ensure_output_directory(output_file)
        
        # 确保时间戳字段是数字类型
        ts_columns = timestamp_columns if isinstance(timestamp_columns, list) else [timestamp_columns]
        if 'timestamp' not in ts_columns:
            ts_columns.append('timestamp')
            
        # 最后一步强制转换时间戳为整数类型
        # 这个转换必须在数据转换为JSON前完成
        for item in json_data:
            for col in ts_columns:
                if col in item:
                    # 强制将任何可转换为数字的时间戳转为整数
                    try:
                        # 先尝试转换为字符串，再转换为整数（处理各种类型）
                        str_value = str(item[col])
                        if str_value.isdigit():
                            item[col] = int(str_value)
                    except (ValueError, TypeError):
                        # 如果无法转换，保留原值
                        pass
        
        # 确保json.dump能够正确序列化
        try:
            # 再次检查所有时间戳字段，确保是整数类型
            for item in json_data:
                for col in ts_columns:
                    if col in item:
                        # 最后的安全保障
                        if isinstance(item[col], str) and item[col].isdigit():
                            item[col] = int(item[col])
                        elif isinstance(item[col], float):
                            # 如果是浮点数，转换为整数
                            item[col] = int(item[col])
            
            with open(output_file, 'w', encoding='utf-8') as f:
                # 无需自定义编码器，因为我们已经确保了数据类型
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            print_success(f"JSON文件已保存至：{output_file}")
        except Exception as e:
            print_error(f"保存JSON文件失败：{str(e)}")
            # 添加调试信息
            print_warning("  调试信息 - 前两个记录的时间戳类型：")
            if len(json_data) > 0:
                for i, item in enumerate(json_data[:2]):
                    if 'timestamp' in item:
                        print(f"    记录{i+1} timestamp值：{item['timestamp']}, 类型：{type(item['timestamp']).__name__}")
            raise
        
    except FileNotFoundError as e:
        print(f"错误：找不到文件 - {e}")
        raise
    except ValueError as e:
        print(f"错误：数据格式问题 - {e}")
        raise
    except PermissionError:
        print(f"错误：权限不足，无法读取文件 '{excel_file}' 或写入文件 '{output_file}'")
        raise
    except Exception as e:
        print(f"处理Excel文件时出错：{e}")
        raise
    
    # 返回最终的输出文件路径
    return output_file


if __name__ == "__main__":
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(
        description='Excel/CSV转JSON工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法：
  # Excel文件示例
  python excel_to_json.py example.xlsx
  python excel_to_json.py example.xlsx -o output.json
  python excel_to_json.py example.xlsx -s Sheet2 -o output.json
  python excel_to_json.py example.xlsx -t create_time update_time
  
  # CSV文件示例
  python excel_to_json.py example.csv
  python excel_to_json.py example.csv -o output.json
  python excel_to_json.py example.csv -d ';' -e 'utf-8-sig'  # 使用分号分隔符和带BOM的UTF-8编码
  python excel_to_json.py example.csv -t create_time -d '\t'  # 使用Tab分隔符
""")
    
    # 添加必需的输入文件参数
    parser.add_argument('input_file', help='输入的Excel文件路径')
    
    # 添加可选的输出文件参数
    parser.add_argument('-o', '--output', help='输出的JSON文件路径，默认为Excel文件名+.json')
    
    # 添加可选的工作表参数
    parser.add_argument('-s', '--sheet', default=0, help='要读取的工作表名称或索引，默认为第一个工作表')
    
    # 添加可选的时间戳列参数
    parser.add_argument('-t', '--timestamp', nargs='+', help='要转换为时间戳的列名列表，默认自动检测名为\'timestamp\'的列')
    
    # 添加CSV相关参数
    parser.add_argument('-d', '--delimiter', default=',', help='CSV文件的分隔符，默认为逗号（仅适用于CSV文件）')
    parser.add_argument('-e', '--encoding', default='utf-8', help='CSV文件的编码，默认为utf-8（仅适用于CSV文件）')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    print("=" * 60)
    print("          Excel转JSON转换工具         ")
    print("=" * 60)
    
    # 调用主函数执行转换
    try:
        # 传递所有参数
        excel_to_json(args.input_file, args.output, args.sheet, args.timestamp, args.delimiter, args.encoding)
        print("\n" + "=" * 60)
        print("🎉 转换成功完成！")
        print("=" * 60)
    except FileNotFoundError as e:
        print(f"\n❌ 转换失败：文件未找到 - {e}")
        print("请检查文件路径是否正确。")
        exit(1)
    except ValueError as e:
        print(f"\n❌ 转换失败：{e}")
        print("请检查文件格式或工作表名称是否正确。")
        exit(1)
    except PermissionError:
        print("\n❌ 转换失败：权限不足")
        print("请检查文件的读写权限。")
        exit(1)
    except Exception as e:
        print(f"\n❌ 转换失败：{e}")
        exit(1)