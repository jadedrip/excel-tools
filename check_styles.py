#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
详细检查Excel文件的styles.xml内容
"""

import os
import zipfile
import re

def extract_styles_xml():
    """提取并分析styles.xml文件内容"""
    file_path = "子站设备调试设备清单.xlsx"
    
    if not os.path.exists(file_path):
        print(f"[错误] 文件不存在: {file_path}")
        return
    
    print("=" * 60)
    print("提取styles.xml文件内容")
    print("=" * 60)
    
    with zipfile.ZipFile(file_path, 'r') as zip_file:
        if 'xl/styles.xml' not in zip_file.namelist():
            print("[错误] styles.xml文件不存在")
            return
        
        content = zip_file.read('xl/styles.xml').decode('utf-8')
        
        print(f"[信息] styles.xml大小: {len(content)} bytes")
        print("\n" + "=" * 60)
        print("文件内容预览 (前5000字符)")
        print("=" * 60)
        print(content[:5000])
        
        print("\n" + "=" * 60)
        print("分析fills元素")
        print("=" * 60)
        
        # 查找fills元素
        fills_pattern = re.compile(r'<fills[^>]*>(.*?)</fills>', re.DOTALL | re.IGNORECASE)
        fills_matches = fills_pattern.findall(content)
        
        print(f"找到 {len(fills_matches)} 个fills元素")
        
        for i, fill_match in enumerate(fills_matches):
            print(f"\n--- fills元素 {i+1} ---")
            print(f"内容长度: {len(fill_match)}")
            print(f"内容预览:\n{fill_match[:500]}")
            
            # 查找fill子元素
            fill_items = re.findall(r'<fill[^>]*>(.*?)</fill>', fill_match, re.DOTALL | re.IGNORECASE)
            print(f"找到 {len(fill_items)} 个fill子元素")
            
            for j, fill_item in enumerate(fill_items[:5]):  # 只显示前5个
                print(f"  fill {j+1}: {fill_item[:100]}...")
        
        print("\n" + "=" * 60)
        print("分析cellXfs元素")
        print("=" * 60)
        
        # 查找cellXfs元素
        xfs_pattern = re.compile(r'<cellXfs[^>]*>(.*?)</cellXfs>', re.DOTALL | re.IGNORECASE)
        xfs_matches = xfs_pattern.findall(content)
        
        print(f"找到 {len(xfs_matches)} 个cellXfs元素")
        
        for i, xf_match in enumerate(xfs_matches):
            print(f"\n--- cellXfs元素 {i+1} ---")
            print(f"内容长度: {len(xf_match)}")
            
            # 查找xf子元素
            xf_items = re.findall(r'<xf[^>]*>', xf_match)
            print(f"找到 {len(xf_items)} 个xf元素")
            
            # 显示前几个xf元素的属性
            for j, xf in enumerate(xf_items[:10]):
                print(f"  xf {j+1}: {xf[:150]}")
        
        print("\n" + "=" * 60)
        print("检查可能的错误格式")
        print("=" * 60)
        
        # 检查是否有空的fill元素
        empty_fills = re.findall(r'<fill\s*/>', content)
        print(f"找到 {len(empty_fills)} 个空fill标签 (self-closing)")
        
        # 检查是否有缺少patternFill的fill
        fills_with_pattern = re.findall(r'<fill[^>]*>.*?<patternFill', content, re.DOTALL)
        fills_without_pattern = len(fills_matches) - len(fills_with_pattern) if fills_matches else 0
        print(f"有patternFill的fill: {len(fills_with_pattern)}")
        print(f"可能没有patternFill的fill: {fills_without_pattern}")


def check_openpyxl_version():
    """检查openpyxl版本"""
    import openpyxl
    print("=" * 60)
    print("openpyxl版本信息")
    print("=" * 60)
    print(f"版本: {openpyxl.__version__}")
    print(f"位置: {openpyxl.__file__}")


def create_fixed_styles():
    """创建一个修复后的styles.xml"""
    file_path = "子站设备调试设备清单.xlsx"
    
    if not os.path.exists(file_path):
        print(f"[错误] 文件不存在: {file_path}")
        return
    
    print("\n" + "=" * 60)
    print("创建修复后的文件")
    print("=" * 60)
    
    # 备份原文件
    backup_path = file_path.replace('.xlsx', '_backup.xlsx')
    if not os.path.exists(backup_path):
        import shutil
        shutil.copy2(file_path, backup_path)
        print(f"[信息] 已备份原文件到: {backup_path}")
    
    with zipfile.ZipFile(file_path, 'r') as zip_file:
        content = zip_file.read('xl/styles.xml').decode('utf-8')
    
    # 修复：添加缺失的fills根元素或修复格式问题
    # 方法1: 如果没有fills元素，添加一个默认的
    if '<fills' not in content.lower():
        print("[信息] 文件中没有fills元素，需要添加")
        
        # 查找styles标签位置
        styles_match = re.search(r'<styleSheet[^>]*>', content)
        if styles_match:
            insert_pos = styles_match.end()
            default_fills = '''<fills count="2"><fill><patternFill patternType="none"/><patternFill><fgColor rgb="FFFFFF"/><bgColor rgb="FFFFFF"/></patternFill></fill><fill><patternFill patternType="gray125"/></fill></fills>'''
            new_content = content[:insert_pos] + default_fills + content[insert_pos:]
            
            # 写回文件
            with zipfile.ZipFile(file_path, 'a') as zip_file_output:
                pass  # 需要重新创建整个zip文件
            
            print("[信息] 已尝试修复fills元素")
    
    print("[注意] 请用Excel重新保存文件以修复样式问题")
    print("或者使用备份文件: " + backup_path)


def try_repair_styles_xml():
    """尝试修复styles.xml"""
    file_path = "子站设备调试设备清单.xlsx"
    
    if not os.path.exists(file_path):
        print(f"[错误] 文件不存在: {file_path}")
        return
    
    print("\n" + "=" * 60)
    print("尝试修复styles.xml")
    print("=" * 60)
    
    import shutil
    from lxml import etree
    
    try:
        # 读取原文件
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            styles_xml = zip_file.read('xl/styles.xml').decode('utf-8')
        
        # 尝试解析XML
        print("[信息] 尝试解析styles.xml...")
        root = etree.fromstring(styles_xml.encode())
        print("[信息] XML解析成功")
        
        # 检查命名空间
        nsmap = root.nsmap
        print(f"[信息] 命名空间: {nsmap}")
        
        # 查找fills元素
        fills = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fills')
        if fills is not None:
            print(f"[信息] fills元素存在，子元素数量: {len(fills)}")
            
            # 检查每个fill元素
            for i, fill in enumerate(fills):
                print(f"  fill {i}: {etree.tostring(fill, encoding='unicode')[:100]}")
        else:
            print("[警告] fills元素不存在")
        
    except etree.XMLSyntaxError as e:
        print(f"[错误] XML语法错误: {e}")
        
        # 尝试找到问题位置
        error_line = str(e).split('\n')[0]
        print(f"[错误信息] {error_line}")
        
    except Exception as e:
        print(f"[错误] 修复失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    check_openpyxl_version()
    extract_styles_xml()
    try_repair_styles_xml()
