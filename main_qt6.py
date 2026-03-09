#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel处理工具主程序
功能：读取Excel文件，可视化配置字段处理规则，生成新的Excel文件
作者：jadedrip

"""

import os
import json
import sys
import subprocess
import pandas as pd
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon

# 导入帮助管理模块
from help_manager import show_split_help

# 导入相关类
from config_tree_widget import ConfigTreeWidget
from draggable_tree_widget import DraggableTreeWidget
from droppable_list_widget import DroppableListWidget
from table_split_worker import TableSplitWorker
from excel_processor_app import ExcelProcessorApp   

def main():
    """主函数"""
    try:
        app = QApplication(sys.argv)
        
        # 设置应用程序图标
        app_icon = QIcon("icon.ico")
        app.setWindowIcon(app_icon)
        
        window = ExcelProcessorApp()
        # 设置窗口图标
        window.setWindowIcon(app_icon)
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"[错误] 应用程序启动失败: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()