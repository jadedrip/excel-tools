#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一转换工具
自动检测输入文件类型并进行相应的转换：
- JSON文件转换为Excel
- Excel/CSV文件转换为JSON
"""

import os
import argparse
from utils import (
    print_header, print_success, print_warning, print_error,
    get_file_type, EXCEL_EXTENSIONS, CSV_EXTENSIONS, JSON_EXTENSIONS,
    get_default_output_file, get_unique_filename
)

# 动态导入转换模块，避免循环导入

def import_converter_modules():
    """
    动态导入转换模块
    
    Returns:
        tuple: (excel_to_json函数, json_to_excel函数)
    """
    from excel_to_json import excel_to_json
    from json_to_excel import json_to_excel
    return excel_to_json, json_to_excel

def convert_file(input_file: str, output_file: str = None, **kwargs) -> None:
    """
    自动检测文件类型并进行转换
    
    Args:
        input_file: 输入文件路径
        output_file: 输出文件路径，默认为自动生成
        **kwargs: 其他转换参数
    """
    print_header("统一文件格式转换工具")
    print(f"输入文件：{input_file}")
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"错误：找不到输入文件 '{input_file}'")
    
    # 获取文件类型
    file_type = get_file_type(input_file)
    
    # 确定输出文件路径和转换方向
    if file_type == 'json':
        # JSON -> Excel
        conversion_type = "JSON转Excel"
        if output_file is None:
            output_file = get_default_output_file(input_file, 'excel')
        # 获取不冲突的文件名
        original_output = output_file
        output_file = get_unique_filename(output_file)
        if output_file != original_output:
            print_warning(f"输出文件已存在，将使用新文件名: {output_file}")
        print_success(f"检测到{file_type.upper()}文件，将转换为Excel格式")
    elif file_type in ['excel', 'csv']:
        # Excel/CSV -> JSON
        conversion_type = f"{file_type.upper()}转JSON"
        if output_file is None:
            output_file = get_default_output_file(input_file, 'json')
        # 获取不冲突的文件名
        original_output = output_file
        output_file = get_unique_filename(output_file)
        if output_file != original_output:
            print_warning(f"输出文件已存在，将使用新文件名: {output_file}")
        print_success(f"检测到{file_type.upper()}文件，将转换为JSON格式")
    else:
        raise ValueError(f"错误：不支持的文件类型，文件 '{input_file}'")
    
    print(f"输出文件：{output_file}")
    print(f"转换类型：{conversion_type}")
    print("=" * 60)
    
    # 动态导入转换模块
    excel_to_json_func, json_to_excel_func = import_converter_modules()
    
    # 根据文件类型执行转换
    try:
        if file_type == 'json':
            # 从kwargs中提取JSON转Excel所需的参数
            datetime_fields = kwargs.get('datetime_fields')
            date_format = kwargs.get('date_format', '%Y-%m-%d %H:%M:%S')
            root_field = kwargs.get('root_field')
            custom_fields = kwargs.get('custom_fields')
            
            # 执行JSON转Excel
            json_to_excel_func(
                input_file=input_file,
                output_file=output_file,
                datetime_fields=datetime_fields,
                date_format=date_format,
                root_field=root_field,
                custom_fields=custom_fields
            )
        else:  # excel或csv
            # 从kwargs中提取Excel/CSV转JSON所需的参数
            sheet_name = kwargs.get('sheet_name', 0)  # Excel工作表，CSV不使用
            timestamp_columns = kwargs.get('timestamp_columns')
            csv_delimiter = kwargs.get('csv_delimiter', ',')  # CSV分隔符
            csv_encoding = kwargs.get('csv_encoding', 'utf-8')  # CSV编码
            
            # 执行Excel/CSV转JSON
            excel_to_json_func(input_file, output_file, sheet_name, 
                              timestamp_columns, csv_delimiter, csv_encoding, 
                              ignore_fields=kwargs.get('ignore_fields'))
        
        print("=" * 60)
        print_success(f"🎉 转换完成！")
        print(f"输入文件：{input_file}")
        print(f"输出文件：{output_file}")
    except Exception as e:
        print_error(f"转换失败：{e}")
        raise

def main():
    """
    主函数，处理命令行参数
    """
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(
        description='统一文件格式转换工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法：
  # 基本转换
  python convert.py input.json       # 自动转换为Excel
  python convert.py input.xlsx       # 自动转换为JSON
  python convert.py input.csv        # 自动转换为JSON
  
  # 指定输出文件
  python convert.py input.json -o output.xlsx
  python convert.py input.xlsx -o output.json
  
  # 时间戳处理
  python convert.py input.json -t create_time update_time  # 转换指定字段为日期时间
  python convert.py input.xlsx -t timestamp                # 转换指定字段为时间戳
  
  # Excel工作表指定
  python convert.py input.xlsx -s Sheet2                   # 使用指定的工作表
  
  # JSON嵌套字段提取
  python convert.py input.json -r data.items               # 从嵌套结构中提取指定字段
  python convert.py input.json -r response.result.records  # 支持多级嵌套字段路径
  
  # CSV文件参数
  python convert.py input.csv -d ';' -e 'utf-8-sig'        # 使用分号分隔符和带BOM的UTF-8编码
  """)
    
    # 基础参数
    parser.add_argument('input_file', help='输入文件路径')
    parser.add_argument('output_file', nargs='?', help='输出文件路径（可选，默认为自动生成）')
    parser.add_argument('-o', '--output', dest='output_file', help='输出文件路径（可选，与位置参数效果相同）')
    
    # 时间戳处理参数
    parser.add_argument('-t', '--timestamp', nargs='+', help='时间戳相关字段列表')
    
    # Excel工作表参数
    parser.add_argument('-s', '--sheet', help='Excel工作表名称或索引（仅对Excel文件有效）')
    
    # CSV文件相关参数
    parser.add_argument('-d', '--delimiter', default=',', help='CSV文件分隔符（默认为逗号）')
    parser.add_argument('-e', '--encoding', default='utf-8', help='CSV文件编码（默认为utf-8）')
    
    # JSON嵌套字段参数
    parser.add_argument('-r', '--root-field', help='JSON嵌套字段路径，使用.分隔多级字段，如 "data.items"（仅对JSON转Excel有效）')
    
    # 自定义字段映射参数
    parser.add_argument('--field', action='append', dest='custom_fields', help='自定义字段映射，格式为"列名:字段路径"，例如: "value:attrValues[{name=Ep}].value"，可多次使用（仅用于JSON转Excel）')
    
    # 字段过滤参数
    parser.add_argument('--ignore', help='要忽略的字段列表，使用逗号分隔多个字段，例如: "field1,field2,field3"（对双向转换都有效）')
    
    # 解析参数
    args = parser.parse_args()
    
    # 准备转换参数
    kwargs = {}
    
    # 根据文件类型设置相应的参数
    file_type = get_file_type(args.input_file) if os.path.exists(args.input_file) else None
    
    if file_type == 'json':
          # JSON转Excel参数
          kwargs['datetime_fields'] = getattr(args, 'datetime_fields', None)
          kwargs['date_format'] = getattr(args, 'format', '%Y-%m-%d %H:%M:%S')
          kwargs['root_field'] = getattr(args, 'root_field', None)
          kwargs['custom_fields'] = getattr(args, 'custom_fields', None)
          # 添加ignore参数
          kwargs['ignore_fields'] = args.ignore
    elif file_type in ['excel', 'csv']:
        # Excel/CSV转JSON参数
        kwargs['timestamp_columns'] = args.timestamp
        # 添加ignore参数
        kwargs['ignore_fields'] = args.ignore
        if args.sheet:
            # 尝试将sheet参数转换为整数（如果是数字字符串）
            try:
                kwargs['sheet_name'] = int(args.sheet)
            except ValueError:
                kwargs['sheet_name'] = args.sheet
        else:
            kwargs['sheet_name'] = 0
        
        # CSV特有参数
        if file_type == 'csv':
            kwargs['csv_delimiter'] = args.delimiter
            kwargs['csv_encoding'] = args.encoding
    
    # 执行转换
    try:
        convert_file(args.input_file, args.output_file, **kwargs)
    except FileNotFoundError as e:
        print(f"\n❌ 错误：文件未找到 - {e}")
        print("请检查文件路径是否正确。")
        exit(1)
    except ValueError as e:
        print(f"\n❌ 错误：{e}")
        print("请检查文件格式是否正确。")
        exit(1)
    except PermissionError:
        print(f"\n❌ 错误：权限不足")
        print("请检查文件的读写权限。")
        exit(1)
    except Exception as e:
        print(f"\n❌ 错误：{e}")
        exit(1)

if __name__ == "__main__":
    main()