#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel处理工具应用类
功能：读取Excel文件，可视化配置字段处理规则，生成新的Excel文件
作者：jadedrip

"""

import os
import json
import sys
import subprocess
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QComboBox, QLineEdit, QPushButton, QTreeWidget,
    QTreeWidgetItem, QHeaderView, QMessageBox, QFileDialog, QFrame,
    QSplitter, QGroupBox, QScrollArea, QMenuBar, QStatusBar, QSizePolicy,
    QTabWidget, QCheckBox, QListWidget, QAbstractItemView, QProgressBar
)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QThread, QTimer, QMimeData, QPoint, QEvent
from PyQt6.QtGui import QFont, QIcon, QDrag, QPixmap, QPainter, QColor, QPen, QFontMetrics

# 导入帮助管理模块
from help_manager import show_split_help

# 导入其他相关类
from config_tree_widget import ConfigTreeWidget
from draggable_tree_widget import DraggableTreeWidget
from droppable_list_widget import DroppableListWidget
from table_split_worker import TableSplitWorker

class ExcelProcessorApp(QMainWindow):
    """Excel处理工具应用类"""
    
    def __init__(self):
        """初始化应用"""
        super().__init__()
        
        # 数据存储
        self.file_path = ""
        self.df = None
        self.sheet_names = []
        self.current_sheet = ""
        self.output_configs = []  # 存储输出配置
        self.current_config_path = ""  # 存储当前打开的配置文件路径
        
        # 工作线程
        self.split_worker = None
        
        # 初始化标志
        self._is_initializing = False
        # 配置改动跟踪
        self._config_modified = False
        
        # 初始化UI
        self.init_ui()
        
        # 设置样式
        self.setup_style()
    
    def init_ui(self):
        """初始化用户界面"""
        # 设置窗口标题和大小
        self.setWindowTitle("Excel处理工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建菜单栏
        self.create_menu()
        
        # 创建主窗口部件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # 创建主布局
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)  # 设置边距为0，确保充满整个窗口
        self.main_layout.setSpacing(0)  # 设置组件间距为0
        
        # 创建顶部控制区域
        self.create_top_controls()
        
        # 创建中间主体区域 - 使用分割器
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # 设置分割器可扩展
        self.main_layout.addWidget(self.splitter)
        
        # 创建左侧原始数据区域
        self.create_left_panel()
        
        # 创建右侧配置区域
        self.create_right_panel()
        
        # 设置分割器比例 - 左侧不超过一半
        self.splitter.setSizes([600, 600])
        self.splitter.setStretchFactor(0, 1)  # 左侧可扩展
        self.splitter.setStretchFactor(1, 1)  # 右侧可扩展
        
        # 创建底部按钮区域
        self.create_bottom_buttons()
        
        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
        
        # 加载默认配置
        self.load_default_config()
    
    def load_default_config(self):
        """加载默认配置文件"""
        default_config_path = "default.json"
        
        try:
            if os.path.exists(default_config_path):
                print(f"[信息] 发现默认配置文件: {default_config_path}")
                
                with open(default_config_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                self.output_configs = config_data.get("output_configs", [])
                self.update_config_display()
                
                file_path = config_data.get("file_path", "")
                current_sheet = config_data.get("current_sheet", "")
                
                if file_path and os.path.exists(file_path):
                    print(f"[信息] 尝试加载Excel文件: {file_path}")
                    
                    try:
                        excel_file = pd.ExcelFile(file_path)
                        self.sheet_names = excel_file.sheet_names
                        self.file_path = file_path
                        
                        self.sheet_combobox.clear()
                        self.sheet_combobox.addItems(self.sheet_names)
                        self.sheet_combobox.setEnabled(True)
                        
                        if current_sheet in self.sheet_names:
                            index = self.sheet_names.index(current_sheet)
                            if index >= 0:
                                self.sheet_combobox.setCurrentIndex(index)
                            self.load_sheet_data(current_sheet)
                        elif self.sheet_names:
                            self.load_sheet_data(self.sheet_names[0])
                            
                        print(f"[信息] 默认配置加载完成")
                        return
                        
                    except Exception as file_e:
                        print(f"[信息] 直接加载失败: {str(file_e)}")
                        
                        # 检查是否是样式错误
                        if "Fill" in str(file_e) or "styles" in str(file_e).lower():
                            print("[信息] 检测到样式错误，尝试自动修复...")
                            
                            fixed_path = self.repair_excel_file(file_path)
                            if fixed_path:
                                try:
                                    excel_file = pd.ExcelFile(fixed_path)
                                    self.sheet_names = excel_file.sheet_names
                                    self.file_path = fixed_path
                                    
                                    self.sheet_combobox.clear()
                                    self.sheet_combobox.addItems(self.sheet_names)
                                    self.sheet_combobox.setEnabled(True)
                                    
                                    if current_sheet in self.sheet_names:
                                        index = self.sheet_names.index(current_sheet)
                                        if index >= 0:
                                            self.sheet_combobox.setCurrentIndex(index)
                                        self.load_sheet_data(current_sheet)
                                    elif self.sheet_names:
                                        self.load_sheet_data(self.sheet_names[0])
                                    
                                    print(f"[信息] 已加载修复后的文件，默认配置加载完成")
                                    return
                                    
                                except Exception as retry_e:
                                    print(f"[错误] 加载修复后的文件失败: {str(retry_e)}")
                        
                        # 加载失败，只加载配置，不加载数据
                        print(f"[警告] 无法加载Excel文件，但配置已加载")
                        
            else:
                print(f"[信息] 未找到默认配置文件: {default_config_path}")
                
        except Exception as e:
            error_msg = f"加载默认配置失败: {str(e)}"
            print(f"[错误] {error_msg}")
    
    def create_menu(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件")
        
        open_action = file_menu.addAction("打开Excel文件")
        open_action.triggered.connect(self.open_excel_file)
        
        file_menu.addSeparator()
        
        save_config_action = file_menu.addAction("保存配置")
        save_config_action.triggered.connect(self.save_config)
        
        load_config_action = file_menu.addAction("加载配置")
        load_config_action.triggered.connect(self.load_config)
        
        file_menu.addSeparator()
        
        exit_action = file_menu.addAction("退出")
        exit_action.triggered.connect(self.close)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        
        split_help_action = help_menu.addAction("表格切分说明")
        split_help_action.triggered.connect(lambda: show_split_help(self))
        
        help_menu.addSeparator()
        
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self.show_about)
    
    def create_top_controls(self):
        """创建顶部控制区域"""
        top_frame = QFrame()
        top_layout = QHBoxLayout(top_frame)
        
        # Sheet选择
        sheet_label = QLabel("Sheet:")
        top_layout.addWidget(sheet_label)
        
        self.sheet_combobox = QComboBox()
        self.sheet_combobox.setEnabled(False)
        self.sheet_combobox.setMinimumWidth(200)  # 设置最小宽度为200像素
        self.sheet_combobox.currentTextChanged.connect(self.on_sheet_selected)
        top_layout.addWidget(self.sheet_combobox)
        
        # 打开文件按钮
        open_button = QPushButton("打开Excel文件")
        open_button.clicked.connect(self.open_excel_file)
        top_layout.addWidget(open_button)
        
        top_layout.addStretch()
        
        self.main_layout.addWidget(top_frame)
    
    def create_left_panel(self):
        """创建左侧原始数据面板"""
        left_frame = QFrame()
        left_frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # 设置左侧面板可扩展
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(5, 5, 5, 5)  # 设置内部边距
        
        # 列选择区域
        columns_group = QGroupBox("列信息")
        columns_layout = QVBoxLayout(columns_group)
        
        self.columns_tree = DraggableTreeWidget()
        self.columns_tree.setHeaderLabels(["序号", "列名", "示例数据"])  # 添加序号列
        self.columns_tree.setColumnWidth(0, 50)  # 设置序号列宽度
        self.columns_tree.setColumnWidth(1, 150)
        self.columns_tree.setColumnWidth(2, 200)
        columns_layout.addWidget(self.columns_tree)
        
        left_layout.addWidget(columns_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout(preview_group)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        self.preview_tree = QTreeWidget()
        self.preview_tree.setHeaderLabels([])
        scroll_area.setWidget(self.preview_tree)
        
        preview_layout.addWidget(scroll_area)
        left_layout.addWidget(preview_group)
        
        # 添加到分割器
        self.splitter.addWidget(left_frame)
    
    def create_right_panel(self):
        """创建右侧配置面板"""
        right_frame = QFrame()
        right_frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # 设置右侧面板可扩展
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(5, 5, 5, 5)  # 设置内部边距
        
        # 创建标签页控件
        self.tab_widget = QTabWidget()
        right_layout.addWidget(self.tab_widget)
        
        # 创建"智能输出"标签页
        self.create_smart_output_tab()
        
        # 创建"表格切分"标签页
        self.create_table_split_tab()
        
        # 创建"转为 json"标签页
        self.create_json_convert_tab()
        
        # 添加到分割器
        self.splitter.addWidget(right_frame)
    
    def create_smart_output_tab(self):
        """创建智能输出标签页"""
        smart_output_widget = QWidget()
        smart_output_layout = QVBoxLayout(smart_output_widget)
        
        # 配置表格区域
        config_group = QGroupBox("输出配置")
        config_layout = QVBoxLayout(config_group)
        
        self.config_tree = ConfigTreeWidget()
        self.config_tree.setHeaderLabels(["输入列", "新列名", "处理规则", "参数"])
        self.config_tree.setColumnWidth(0, 120)
        self.config_tree.setColumnWidth(1, 120)
        self.config_tree.setColumnWidth(2, 120)
        self.config_tree.setColumnWidth(3, 150)
        self.config_tree.itemSelectionChanged.connect(self.on_config_selected)
        # 连接模型变化信号
        self.config_tree.model().layoutChanged.connect(self.on_config_layout_changed)
        config_layout.addWidget(self.config_tree)
        
        smart_output_layout.addWidget(config_group)
        
        # 添加配置区域
        add_config_group = QGroupBox("配置")
        add_config_layout = QGridLayout(add_config_group)
        
        # 输入列选择
        add_config_layout.addWidget(QLabel("选择输入列:"), 0, 0)
        self.input_col_combobox = QComboBox()
        self.input_col_combobox.setEnabled(False)
        self.input_col_combobox.currentTextChanged.connect(self.on_form_field_changed)
        add_config_layout.addWidget(self.input_col_combobox, 0, 1)
        
        # 新列名输入
        add_config_layout.addWidget(QLabel("新列名:"), 0, 2)
        self.new_col_entry = QLineEdit()
        self.new_col_entry.textChanged.connect(self.on_form_field_changed)
        add_config_layout.addWidget(self.new_col_entry, 0, 3)
        
        # 处理规则选择
        add_config_layout.addWidget(QLabel("处理规则:"), 1, 0)
        self.rule_combobox = QComboBox()
        self.rule_combobox.addItems(["直接复制", "前缀添加", "后缀添加", "前后添加", "固定值", "正则替换"])
        self.rule_combobox.setEnabled(False)
        self.rule_combobox.currentTextChanged.connect(self.on_rule_selected)
        add_config_layout.addWidget(self.rule_combobox, 1, 1)
        
        # 参数区域
        self.param_frame = QFrame()
        self.param_layout = QGridLayout(self.param_frame)
        add_config_layout.addWidget(self.param_frame, 1, 2, 1, 2)
        
        # 创建参数控件
        self.create_param_widgets()
        
        # 新建按钮（固定值生成列）
        new_button = QPushButton("新建")
        new_button.clicked.connect(self.create_new_field)
        add_config_layout.addWidget(new_button, 2, 0)
        
        smart_output_layout.addWidget(add_config_group)
        
        # 底部按钮区域
        bottom_buttons_widget = QWidget()
        bottom_buttons_layout = QHBoxLayout(bottom_buttons_widget)
        
        # 快速保存按钮
        quick_save_button = QPushButton("快速保存")
        quick_save_button.clicked.connect(self.quick_save_config)
        bottom_buttons_layout.addWidget(quick_save_button)
        
        # 空一个按钮的距离（使用拉伸因子）
        bottom_buttons_layout.addStretch()
        
        # 删除按钮
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(self.remove_config)
        bottom_buttons_layout.addWidget(delete_button)
        
        # 清空配置按钮
        clear_button = QPushButton("清空配置")
        clear_button.clicked.connect(self.clear_configs)
        bottom_buttons_layout.addWidget(clear_button)
        
        # 生成Excel按钮
        generate_button = QPushButton("生成Excel")
        generate_button.clicked.connect(self.generate_excel)
        bottom_buttons_layout.addWidget(generate_button)
        
        smart_output_layout.addWidget(bottom_buttons_widget)
        
        # 添加到标签页
        self.tab_widget.addTab(smart_output_widget, "智能输出")
    
    def create_table_split_tab(self):
        """创建表格切分标签页"""
        table_split_widget = QWidget()
        table_split_layout = QVBoxLayout(table_split_widget)
        
        # 表格切分配置区域
        split_config_group = QGroupBox("表格切分配置")
        split_config_layout = QGridLayout(split_config_group)
        
        # 输出目录选择 - 移动到最上方
        split_config_layout.addWidget(QLabel("输出目录:"), 0, 0)
        self.split_output_dir_entry = QLineEdit()
        self.split_output_dir_entry.setPlaceholderText("默认使用工作表名")
        self.split_output_dir_entry.setEnabled(False)
        split_config_layout.addWidget(self.split_output_dir_entry, 0, 1)
        
        # 浏览按钮
        self.browse_button = QPushButton("浏览")
        self.browse_button.clicked.connect(self.browse_output_dir)
        split_config_layout.addWidget(self.browse_button, 0, 2)
        
        # 保留行数输入
        split_config_layout.addWidget(QLabel("保留行数:"), 1, 0)
        self.split_rows_entry = QLineEdit()
        self.split_rows_entry.setPlaceholderText("例如: 8")
        self.split_rows_entry.setText("8")  # 默认值
        split_config_layout.addWidget(self.split_rows_entry, 1, 1)
        
        # 拆分列输入 - 使用可拖放的列表
        split_config_layout.addWidget(QLabel("拆分列:"), 2, 0)
        
        # 创建可拖放的列表
        self.split_columns_list = DroppableListWidget()
        self.split_columns_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        split_config_layout.addWidget(self.split_columns_list, 3, 1, 1, 2)
        
        # 添加提示文本
        hint_label = QLabel("提示：从左侧列信息列表拖动列到此处，可上下拖动调整顺序")
        hint_label.setStyleSheet("color: #666666; font-size: 12px;")  # 浅色小字体
        hint_label.setWordWrap(True)
        split_config_layout.addWidget(hint_label, 2, 1, 1, 2)  # 跨列显示
        
        # 添加删除按钮
        self.split_columns_remove_button = QPushButton("删除选中列")
        self.split_columns_remove_button.clicked.connect(self.remove_split_column)
        split_config_layout.addWidget(self.split_columns_remove_button, 4, 0)
        
        # 上下移动按钮
        self.split_columns_up_button = QPushButton("上移")
        self.split_columns_up_button.setMaximumWidth(80)
        self.split_columns_up_button.clicked.connect(self.move_split_column_up)
        split_config_layout.addWidget(self.split_columns_up_button, 4, 1, Qt.AlignmentFlag.AlignLeft)
        
        self.split_columns_down_button = QPushButton("下移")
        self.split_columns_down_button.setMaximumWidth(80)
        self.split_columns_down_button.clicked.connect(self.move_split_column_down)
        split_config_layout.addWidget(self.split_columns_down_button, 4, 2, Qt.AlignmentFlag.AlignLeft)
        
        table_split_layout.addWidget(split_config_group)
        
        # 操作按钮区域
        split_buttons_widget = QWidget()
        split_buttons_layout = QHBoxLayout(split_buttons_widget)
        
        # 帮助按钮（移到最左边）
        self.split_help_button = QPushButton("帮助")
        self.split_help_button.clicked.connect(lambda: show_split_help(self))
        split_buttons_layout.addWidget(self.split_help_button)
        
        split_buttons_layout.addStretch()
        
        # 执行切分按钮
        self.split_execute_button = QPushButton("执行切分")
        self.split_execute_button.clicked.connect(self.execute_table_split)
        self.split_execute_button.setEnabled(False)  # 默认禁用
        split_buttons_layout.addWidget(self.split_execute_button)
        
        # 取消按钮（初始隐藏）
        self.split_cancel_button = QPushButton("取消")
        self.split_cancel_button.clicked.connect(self.cancel_table_split)
        self.split_cancel_button.setVisible(False)  # 初始隐藏
        split_buttons_layout.addWidget(self.split_cancel_button)
        
        table_split_layout.addWidget(split_buttons_widget)
        
        # 进度条区域
        progress_group = QGroupBox("进度")
        progress_layout = QVBoxLayout(progress_group)
        
        # 进度条
        self.split_progress_bar = QProgressBar()
        self.split_progress_bar.setMinimum(0)
        self.split_progress_bar.setMaximum(100)
        self.split_progress_bar.setValue(0)
        self.split_progress_bar.setVisible(True)  # 始终可见
        self.split_progress_bar.setMinimumHeight(20)  # 设置最小高度
        progress_layout.addWidget(self.split_progress_bar)
        
        # 状态标签
        self.split_status_label = QLabel("空闲")
        self.split_status_label.setVisible(True)  # 始终可见
        self.split_status_label.setMinimumHeight(25)  # 设置最小高度
        progress_layout.addWidget(self.split_status_label)
        
        # 设置固定高度的占位符，确保进度显示区高度一致
        progress_group.setMinimumHeight(100)  # 设置最小高度
        
        table_split_layout.addWidget(progress_group)
        
        # 添加到标签页
        self.tab_widget.addTab(table_split_widget, "表格切分")
    
    def create_bottom_buttons(self):
        """创建底部按钮区域 - 已移至智能输出标签页"""
        pass
    
    def remove_split_column(self):
        """删除选中的拆分列"""
        for item in self.split_columns_list.selectedItems():
            self.split_columns_list.takeItem(self.split_columns_list.row(item))
    
    def move_split_column_up(self):
        """将选中的拆分列上移"""
        current_row = self.split_columns_list.currentRow()
        if current_row > 0:
            item = self.split_columns_list.takeItem(current_row)
            self.split_columns_list.insertItem(current_row - 1, item)
            self.split_columns_list.setCurrentItem(item)
    
    def move_split_column_down(self):
        """将选中的拆分列下移"""
        current_row = self.split_columns_list.currentRow()
        if current_row < self.split_columns_list.count() - 1:
            item = self.split_columns_list.takeItem(current_row)
            self.split_columns_list.insertItem(current_row + 1, item)
            self.split_columns_list.setCurrentItem(item)
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.split_output_dir_entry.setText(dir_path)
    
    def create_json_convert_tab(self):
        """创建转为json标签页"""
        json_convert_widget = QWidget()
        json_convert_layout = QVBoxLayout(json_convert_widget)
        
        # JSON转换配置区域
        json_config_group = QGroupBox("JSON转换配置")
        json_config_layout = QGridLayout(json_config_group)
        
        # 输出文件路径
        json_config_layout.addWidget(QLabel("输出文件:"), 0, 0)
        self.json_output_entry = QLineEdit()
        self.json_output_entry.setPlaceholderText("默认使用工作表名+.json")
        json_config_layout.addWidget(self.json_output_entry, 0, 1)
        
        # 浏览按钮
        self.json_browse_button = QPushButton("浏览")
        self.json_browse_button.clicked.connect(self.json_browse_output)
        json_config_layout.addWidget(self.json_browse_button, 0, 2)
        
        json_convert_layout.addWidget(json_config_group)
        
        # 列选择区域
        columns_group = QGroupBox("选择转换列")
        columns_layout = QVBoxLayout(columns_group)
        
        # 创建滚动区域
        self.columns_scroll = QScrollArea()
        self.columns_scroll.setWidgetResizable(True)
        self.columns_widget = QWidget()
        self.columns_checkbox_layout = QVBoxLayout(self.columns_widget)
        self.columns_scroll.setWidget(self.columns_widget)
        
        columns_layout.addWidget(self.columns_scroll)
        json_convert_layout.addWidget(columns_group)
        
        # 转换按钮区域
        convert_button_widget = QWidget()
        convert_button_layout = QHBoxLayout(convert_button_widget)
        
        convert_button_layout.addStretch()
        
        # 转换按钮
        self.json_convert_button = QPushButton("转换为JSON")
        self.json_convert_button.clicked.connect(self.execute_json_convert)
        self.json_convert_button.setEnabled(False)  # 默认禁用
        convert_button_layout.addWidget(self.json_convert_button)
        
        json_convert_layout.addWidget(convert_button_widget)
        
        # 添加到标签页
        self.tab_widget.addTab(json_convert_widget, "转为 json")
    
    def json_browse_output(self):
        """浏览JSON输出文件"""
        # 默认使用工作表名作为文件名
        default_filename = f"{self.current_sheet}.json" if self.current_sheet else "output.json"
        file_path, _ = QFileDialog.getSaveFileName(self, "保存JSON文件", default_filename, "JSON files (*.json)")
        if file_path:
            self.json_output_entry.setText(file_path)
    
    def execute_table_split(self):
        """执行表格切分"""
        try:
            if not self.file_path:
                QMessageBox.warning(self, "警告", "请先打开Excel文件")
                return
            
            if not self.current_sheet:
                QMessageBox.warning(self, "警告", "请先选择工作表")
                return
            
            # 获取配置参数
            rows_text = self.split_rows_entry.text().strip()
            if not rows_text:
                QMessageBox.warning(self, "警告", "请输入保留行数")
                return
            
            header_rows_count = int(rows_text)
            
            # 从列表中获取拆分列
            if self.split_columns_list.count() == 0:
                QMessageBox.warning(self, "警告", "请添加拆分列")
                return
            
            # 解析拆分列
            split_columns = []
            for i in range(self.split_columns_list.count()):
                item_text = self.split_columns_list.item(i).text()
                # 从格式 "A (列名)" 中提取Excel列名
                excel_col = item_text.split()[0]
                
                # 将字母转换为数字索引（从1开始）
                excel_col = excel_col.upper()
                col_idx = 0
                for char in excel_col:
                    col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
                split_columns.append(col_idx)
            
            # 输出目录
            output_dir = self.split_output_dir_entry.text().strip()
            if not output_dir:
                # 默认使用当前工作目录下的工作表名目录
                output_dir = os.path.join(os.getcwd(), self.current_sheet)
            
            # 确保输出目录存在
            if not os.path.exists(output_dir):
                print(f"[信息] 创建输出目录: {output_dir}")
                os.makedirs(output_dir, exist_ok=True)
            
            # 检查是否已有工作线程在运行
            if self.split_worker and self.split_worker.isRunning():
                QMessageBox.warning(self, "警告", "表格切分正在进行中，请稍候...")
                return
            
            # 创建工作线程
            self.split_worker = TableSplitWorker(
                self.file_path, 
                self.current_sheet, 
                header_rows_count, 
                split_columns, 
                output_dir
            )
            
            # 连接信号
            self.split_worker.progress_updated.connect(self.on_split_progress_updated)
            self.split_worker.finished.connect(self.on_split_finished)
            self.split_worker.file_saved.connect(self.on_split_file_saved)
            
            # 更新UI状态
            self.split_execute_button.setEnabled(False)
            self.split_cancel_button.setVisible(True)
            self.split_progress_bar.setVisible(True)
            self.split_status_label.setVisible(True)
            self.split_progress_bar.setValue(0)
            self.split_status_label.setText("准备开始...")
            
            # 禁用其他控件
            self.sheet_combobox.setEnabled(False)
            
            # 启动工作线程
            self.split_worker.start()
            
        except Exception as e:
            error_msg = f"启动表格切分失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def cancel_table_split(self):
        """取消表格切分"""
        if self.split_worker and self.split_worker.isRunning():
            self.split_worker.cancel()
            self.split_status_label.setText("正在取消...")
    
    def on_split_progress_updated(self, current, total, message):
        """更新切分进度"""
        self.split_progress_bar.setMaximum(total)
        self.split_progress_bar.setValue(current)
        self.split_status_label.setText(message)
    
    def on_split_finished(self, success, message):
        """切分完成处理"""
        # 恢复UI状态
        self.split_execute_button.setEnabled(True)
        self.split_cancel_button.setVisible(False)
        self.sheet_combobox.setEnabled(True)
        
        if success:
            # 创建自定义消息框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle("成功")
            msg_box.setText(message)
            
            # 添加"打开输出目录"按钮
            open_dir_button = msg_box.addButton("打开输出目录", QMessageBox.ButtonRole.ActionRole)
            ok_button = msg_box.addButton(QMessageBox.StandardButton.Ok)
            
            # 显示消息框并获取用户选择
            msg_box.exec()
            
            # 如果用户点击了"打开输出目录"按钮
            if msg_box.clickedButton() == open_dir_button:
                output_dir = self.split_output_dir_entry.text().strip()
                if not output_dir:
                    output_dir = self.current_sheet
                
                # 调用统一方法打开输出目录
                self.open_output_directory(output_dir)
        else:
            QMessageBox.warning(self, "警告", message)
    
    def on_split_file_saved(self, file_path):
        """文件保存完成处理"""
        print(f"文件已保存: {file_path}")
    
    def execute_json_convert(self):
        """执行JSON转换"""
        try:
            if not self.file_path:
                QMessageBox.warning(self, "警告", "请先打开Excel文件")
                return
            
            # 获取选中的列、类型和输出字段名
            selected_columns = []
            timestamp_columns = []
            field_mapping = {}  # 存储输入字段名到输出字段名的映射
            
            for i in range(self.columns_checkbox_layout.count()):
                widget = self.columns_checkbox_layout.itemAt(i).widget()
                if widget and hasattr(widget, 'layout'):
                    h_layout = widget.layout()
                    checkbox = h_layout.itemAt(0).widget()
                    type_combo = h_layout.itemAt(2).widget()
                    output_field_edit = h_layout.itemAt(3).widget()
                    
                    if checkbox and type_combo and output_field_edit and checkbox.isChecked():
                        col_name = checkbox.text()
                        col_type = type_combo.currentText()
                        output_field = output_field_edit.text().strip()
                        
                        # 如果输出字段名为空，使用原字段名
                        if not output_field:
                            output_field = col_name
                        
                        selected_columns.append(col_name)
                        field_mapping[col_name] = output_field
                        
                        if col_type == "时间戳":
                            timestamp_columns.append(col_name)
            
            # 如果没有选中任何列，使用所有列
            if not selected_columns:
                selected_columns = list(self.df.columns)
                # 为所有列创建默认映射
                for col in selected_columns:
                    field_mapping[col] = col
            
            # 获取输出文件路径
            output_file = self.json_output_entry.text().strip()
            
            # 如果没有指定输出文件，使用工作表名+.json
            if not output_file:
                output_file = f"{self.current_sheet}.json"
            
            # 调用excel_to_json函数
            from excel_to_json import excel_to_json as convert_func
            
            # 执行转换
            result_file = convert_func(
                self.file_path,
                output_file=output_file,
                sheet_name=self.current_sheet,
                timestamp_columns=timestamp_columns,
                field_mapping=field_mapping
            )
            
            # 创建带"打开输出目录"按钮的消息框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle("成功")
            msg_box.setText(f"JSON转换完成！\n输出文件: {result_file}")
            
            # 添加"打开输出目录"按钮
            open_dir_button = msg_box.addButton("打开输出目录", QMessageBox.ButtonRole.ActionRole)
            ok_button = msg_box.addButton(QMessageBox.StandardButton.Ok)
            
            # 显示消息框并获取用户选择
            msg_box.exec()
            
            # 如果用户点击了"打开输出目录"按钮
            if msg_box.clickedButton() == open_dir_button:
                try:
                    # 获取输出文件所在的目录
                    output_dir = os.path.dirname(result_file)
                    
                    # 调用统一方法打开输出目录
                    self.open_output_directory(output_dir)
                except Exception as e:
                    print(f"[错误] 打开输出目录失败: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    QMessageBox.warning(self, "警告", f"打开输出目录时发生错误: {str(e)}")
        except Exception as e:
            error_msg = f"JSON转换失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def create_param_widgets(self):
        """创建参数控件"""
        # 前缀参数
        self.prefix_label = QLabel("前缀:")
        self.prefix_entry = QLineEdit()
        self.prefix_entry.setMaximumWidth(150)
        self.prefix_entry.textChanged.connect(self.on_form_field_changed)
        
        # 后缀参数
        self.suffix_label = QLabel("后缀:")
        self.suffix_entry = QLineEdit()
        self.suffix_entry.setMaximumWidth(150)
        self.suffix_entry.textChanged.connect(self.on_form_field_changed)
        
        # 固定值参数
        self.fixed_value_label = QLabel("固定值:")
        self.fixed_value_entry = QLineEdit()
        self.fixed_value_entry.setMaximumWidth(300)
        self.fixed_value_entry.textChanged.connect(self.on_form_field_changed)
        
        # 正则替换参数
        self.regex_pattern_label = QLabel("正则表达式:")
        self.regex_pattern_entry = QLineEdit()
        self.regex_pattern_entry.setMaximumWidth(200)
        self.regex_pattern_entry.setPlaceholderText("例如: (\\d+)-(\\d+)")
        self.regex_pattern_entry.textChanged.connect(self.on_form_field_changed)
        
        self.regex_replace_label = QLabel("替换字符串:")
        self.regex_replace_entry = QLineEdit()
        self.regex_replace_entry.setMaximumWidth(200)
        self.regex_replace_entry.setPlaceholderText("例如: $2-$1 (使用$1,$2表示捕获组)")
        self.regex_replace_entry.textChanged.connect(self.on_form_field_changed)
        
        # 初始隐藏所有参数
        self.hide_all_params()
    
    def hide_all_params(self):
        """隐藏所有参数控件"""
        for i in reversed(range(self.param_layout.count())):
            child = self.param_layout.itemAt(i).widget()
            if child is not None:
                child.setVisible(False)
    
    def setup_style(self):
        """设置样式"""
        # 设置应用程序字体
        font = QFont("SimHei", 10)
        QApplication.instance().setFont(font)
        
        # 设置树形控件样式
        tree_style = """
        QTreeWidget {
            font-family: SimHei;
            font-size: 10pt;
        }
        QTreeWidget::header {
            font-family: SimHei;
            font-size: 10pt;
            font-weight: bold;
        }
        """
        self.columns_tree.setStyleSheet(tree_style)
        self.preview_tree.setStyleSheet(tree_style)
        self.config_tree.setStyleSheet(tree_style)
    
    def open_excel_file(self):
        """打开Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel files (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 读取所有sheet名称
                excel_file = pd.ExcelFile(file_path)
                self.sheet_names = excel_file.sheet_names
                self.file_path = file_path
                
                # 更新sheet下拉框
                self.sheet_combobox.clear()
                self.sheet_combobox.addItems(self.sheet_names)
                self.sheet_combobox.setEnabled(True)
                
                # 默认选择第一个sheet
                if self.sheet_names:
                    self.load_sheet_data(self.sheet_names[0])
                
                QMessageBox.information(self, "成功", f"成功打开文件: {os.path.basename(file_path)}")
                self.status_bar.showMessage(f"已打开: {os.path.basename(file_path)}")
            
            except Exception as e:
                error_msg = f"打开文件失败: {str(e)}"
                print(f"[错误] {error_msg}")
                
                # 检查是否是样式错误
                if "Fill" in str(e) or "styles" in str(e).lower():
                    print("[信息] 检测到样式错误，尝试自动修复...")
                    
                    # 尝试修复文件
                    fixed_path = self.repair_excel_file(file_path)
                    if fixed_path:
                        # 询问用户是否使用修复后的文件
                        reply = QMessageBox.question(
                            self, "文件修复成功",
                            f"原文件存在样式问题，已自动修复。\n\n"
                            f"是否使用修复后的文件？\n\n"
                            f"修复文件: {os.path.basename(fixed_path)}",
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                            QMessageBox.StandardButton.Yes
                        )
                        
                        if reply == QMessageBox.StandardButton.Yes:
                            try:
                                excel_file = pd.ExcelFile(fixed_path)
                                self.sheet_names = excel_file.sheet_names
                                self.file_path = fixed_path
                                
                                self.sheet_combobox.clear()
                                self.sheet_combobox.addItems(self.sheet_names)
                                self.sheet_combobox.setEnabled(True)
                                
                                if self.sheet_names:
                                    self.load_sheet_data(self.sheet_names[0])
                                
                                QMessageBox.information(
                                    self, "成功",
                                    f"成功打开修复后的文件:\n{os.path.basename(fixed_path)}"
                                )
                                self.status_bar.showMessage(f"已打开: {os.path.basename(fixed_path)}")
                                return
                            except Exception as retry_e:
                                print(f"[错误] 打开修复后的文件失败: {str(retry_e)}")
                                QMessageBox.critical(self, "错误", f"无法打开修复后的文件: {str(retry_e)}")
                    else:
                        QMessageBox.warning(self, "修复失败", "无法自动修复文件样式问题")
                else:
                    QMessageBox.critical(self, "错误", error_msg)
    
    def repair_excel_file(self, file_path):
        """修复Excel文件中的样式问题
        
        Args:
            file_path: 原文件路径
            
        Returns:
            str: 修复后的文件路径，修复失败返回None
        """
        print(f"[信息] 开始修复文件: {file_path}")
        
        try:
            import zipfile
            import re
            import shutil
            
            # 检查文件是否存在
            if not os.path.exists(file_path):
                print(f"[错误] 文件不存在: {file_path}")
                return None
            
            # 防止死循环：检查是否已经是修复后的文件
            if '_fixed' in file_path or '_repaired' in file_path:
                print("[警告] 文件已是修复后的版本，不再重复修复")
                return None
            
            # 解压文件
            temp_dir = "temp_xlsx_repair"
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    zip_file.extractall(temp_dir)
                
                # 检查并修复styles.xml
                styles_path = os.path.join(temp_dir, 'xl', 'styles.xml')
                if os.path.exists(styles_path):
                    with open(styles_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    original_size = len(content)
                    
                    # 移除空的自闭合fill标签
                    fixed_content = re.sub(r'<fill\s*/>', '', content)
                    
                    # 检查是否有修复
                    if len(fixed_content) < original_size:
                        print(f"[信息] 发现并移除了空fill标签")
                        
                        # 写回修复后的styles.xml
                        with open(styles_path, 'w', encoding='utf-8') as f:
                            f.write(fixed_content)
                        
                        # 创建修复后的xlsx文件
                        fixed_path = file_path.replace('.xlsx', '_fixed.xlsx')
                        if os.path.exists(fixed_path):
                            os.remove(fixed_path)
                        
                        with zipfile.ZipFile(fixed_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for root, dirs, files in os.walk(temp_dir):
                                for file in files:
                                    file_path_abs = os.path.join(root, file)
                                    arcname = os.path.relpath(file_path_abs, temp_dir)
                                    zip_file.write(file_path_abs, arcname)
                        
                        print(f"[成功] 修复后的文件: {fixed_path}")
                        return fixed_path
                    else:
                        print("[信息] styles.xml没有发现问题")
                        return None
                else:
                    print("[警告] styles.xml不存在")
                    return None
                    
            finally:
                # 清理临时目录
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            print(f"[错误] 修复文件失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
        
        return None
    
    def col_index_to_letter(self, index):
        """将列索引转换为字母，如0->A, 1->B, ..., 25->Z, 26->AA"""
        letters = []
        while index >= 0:
            letters.append(chr(65 + index % 26))
            index = index // 26 - 1
        return ''.join(reversed(letters))
    
    def load_sheet_data(self, sheet_name):
        """加载指定sheet的数据"""
        try:
            self.current_sheet = sheet_name
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            
            # 处理列名：如果列名是Unnamed开头，替换为A, B, C...格式
            new_columns = []
            for i, col in enumerate(self.df.columns):
                if col.startswith('Unnamed:'):
                    # 使用字母作为列名，如A, B, C...
                    new_columns.append(self.col_index_to_letter(i))
                else:
                    new_columns.append(col)
            
            # 更新DataFrame的列名
            self.df.columns = new_columns
            
            # 更新左侧数据显示
            self.update_columns_display()
            self.update_preview_display()
            
            # 启用输入列选择
            self.input_col_combobox.clear()
            self.input_col_combobox.addItems(list(self.df.columns))
            self.input_col_combobox.setEnabled(True)
            self.rule_combobox.setEnabled(True)
            
            # 启用表格切分执行按钮
            self.split_execute_button.setEnabled(True)
            
            # 启用JSON转换按钮
            self.json_convert_button.setEnabled(True)
            
            # 生成列选择复选框
            self.generate_columns_checkboxes()
            
        except Exception as e:
            error_msg = f"加载sheet失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def generate_columns_checkboxes(self):
        """生成列选择复选框、类型选择和输出字段名输入"""
        # 清空现有控件
        for i in reversed(range(self.columns_checkbox_layout.count())):
            widget = self.columns_checkbox_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        
        # 为每一列生成复选框、类型选择和输出字段名输入，默认选中
        for col in self.df.columns:
            # 创建水平布局
            h_layout = QHBoxLayout()
            
            # 列选择复选框
            checkbox = QCheckBox(col)
            checkbox.setChecked(True)
            h_layout.addWidget(checkbox)
            h_layout.addStretch()
            
            # 类型选择下拉框
            type_combo = QComboBox()
            type_combo.addItems(["字符串", "数字", "时间戳"])
            type_combo.setCurrentText("字符串")
            type_combo.setMinimumWidth(120)
            h_layout.addWidget(type_combo)
            
            # 输出字段名输入框
            output_field_edit = QLineEdit()
            output_field_edit.setText(col)  # 默认和选择字段一样
            output_field_edit.setPlaceholderText("输出字段名")
            output_field_edit.setMinimumWidth(150)
            h_layout.addWidget(output_field_edit)
            
            # 将布局添加到容器中
            container = QWidget()
            container.setLayout(h_layout)
            self.columns_checkbox_layout.addWidget(container)
        
        # 添加拉伸，确保控件向上对齐
        self.columns_checkbox_layout.addStretch()
    
    def update_columns_display(self):
        """更新列显示"""
        # 清空现有数据
        self.columns_tree.clear()
        
        # 添加列信息，包含Excel序号
        for i, col in enumerate(self.df.columns):
            # 生成Excel列序号，如A, B, C...
            excel_col = self.col_index_to_letter(i)
            
            # 获取前几个非空值作为示例
            sample_values = self.df[col].dropna().head(3).tolist()
            sample_text = ", ".join([str(v)[:20] for v in sample_values]) if sample_values else "无数据"
            
            item = QTreeWidgetItem([excel_col, col, sample_text])
            self.columns_tree.addTopLevelItem(item)
    
    def update_preview_display(self):
        """更新数据预览显示"""
        # 清空现有数据
        self.preview_tree.clear()
        
        # 设置新列 - 在最前面添加序号列
        headers = ['序号']
        for i, col in enumerate(self.df.columns):
            # 生成列名格式：A:列名，处理列名空的情况
            col_letter = self.col_index_to_letter(i)
            if col and col.strip():
                headers.append(f"{col_letter}:{col}")
            else:
                headers.append(col_letter)  # 如果列名是空的，只显示字母
        self.preview_tree.setHeaderLabels(headers)
        
        # 配置列宽
        self.preview_tree.setColumnWidth(0, 50)  # 设置序号列宽度
        for i, col in enumerate(self.df.columns):
            self.preview_tree.setColumnWidth(i + 1, 100)
        
        # 添加预览数据（最多显示10行）
        preview_data = self.df.head(10)
        for idx, (_, row) in enumerate(preview_data.iterrows(), start=2):
            # 处理每一行数据，将nan显示为空
            values = [str(idx)]  # 序号从2开始
            for col in self.df.columns:
                value = row[col]
                if pd.isna(value):
                    values.append("")  # 将nan显示为空
                else:
                    values.append(str(value)[:50])
            item = QTreeWidgetItem(values)
            self.preview_tree.addTopLevelItem(item)
    
    def on_sheet_selected(self, sheet_name):
        """选择sheet时的处理"""
        if sheet_name:
            self.load_sheet_data(sheet_name)
    
    def on_rule_selected(self, rule):
        """选择规则时的处理"""
        self.hide_all_params()
        
        if rule == "前缀添加":
            self.param_layout.addWidget(self.prefix_label, 0, 0)
            self.param_layout.addWidget(self.prefix_entry, 0, 1)
            self.prefix_label.setVisible(True)
            self.prefix_entry.setVisible(True)
        elif rule == "后缀添加":
            self.param_layout.addWidget(self.suffix_label, 0, 0)
            self.param_layout.addWidget(self.suffix_entry, 0, 1)
            self.suffix_label.setVisible(True)
            self.suffix_entry.setVisible(True)
        elif rule == "前后添加":
            self.param_layout.addWidget(self.prefix_label, 0, 0)
            self.param_layout.addWidget(self.prefix_entry, 0, 1)
            self.param_layout.addWidget(self.suffix_label, 0, 2)
            self.param_layout.addWidget(self.suffix_entry, 0, 3)
            self.prefix_label.setVisible(True)
            self.prefix_entry.setVisible(True)
            self.suffix_label.setVisible(True)
            self.suffix_entry.setVisible(True)
        elif rule == "固定值":
            self.param_layout.addWidget(self.fixed_value_label, 0, 0)
            self.param_layout.addWidget(self.fixed_value_entry, 0, 1, 1, 3)
            self.fixed_value_label.setVisible(True)
            self.fixed_value_entry.setVisible(True)
        elif rule == "正则替换":
            self.param_layout.addWidget(self.regex_pattern_label, 0, 0)
            self.param_layout.addWidget(self.regex_pattern_entry, 0, 1)
            self.param_layout.addWidget(self.regex_replace_label, 0, 2)
            self.param_layout.addWidget(self.regex_replace_entry, 0, 3)
            self.regex_pattern_label.setVisible(True)
            self.regex_pattern_entry.setVisible(True)
            self.regex_replace_label.setVisible(True)
            self.regex_replace_entry.setVisible(True)
        
        # 只有在非初始化状态下才触发表单字段更改事件
        if hasattr(self, '_is_initializing') and not self._is_initializing:
            self.on_form_field_changed()
    
    def add_config(self):
        """添加配置"""
        input_col = self.input_col_combobox.currentText()
        new_col = self.new_col_entry.text().strip()
        rule = self.rule_combobox.currentText()
        
        # 验证输入
        if not new_col:
            QMessageBox.warning(self, "警告", "请输入新列名")
            return
        
        # 对于固定值规则，不需要输入列
        if rule != "固定值" and not input_col:
            QMessageBox.warning(self, "警告", "请选择输入列")
            return
        
        # 构建参数
        params = {}
        if rule == "前缀添加":
            params["prefix"] = self.prefix_entry.text()
        elif rule == "后缀添加":
            params["suffix"] = self.suffix_entry.text()
        elif rule == "前后添加":
            params["prefix"] = self.prefix_entry.text()
            params["suffix"] = self.suffix_entry.text()
        elif rule == "固定值":
            params["value"] = self.fixed_value_entry.text()
        elif rule == "正则替换":
            params["pattern"] = self.regex_pattern_entry.text()
            params["replace"] = self.regex_replace_entry.text()
        
        # 添加到配置列表
        config = {
            "input_col": input_col,
            "new_col": new_col,
            "rule": rule,
            "params": params
        }
        self.output_configs.append(config)
        
        # 更新配置表格
        self.update_config_display()
        
        # 清空输入
        self.new_col_entry.clear()
        self.hide_all_params()
    
    def update_config_display(self):
        """更新配置显示"""
        # 清空现有数据
        self.config_tree.clear()
        
        # 添加配置
        for config in self.output_configs:
            # 根据规则类型生成友好的参数显示
            rule = config["rule"]
            params = config["params"]
            
            if rule == "直接复制":
                params_text = "无参数"
            elif rule == "前缀添加":
                params_text = f"前缀: {params.get('prefix', '')}"
            elif rule == "后缀添加":
                params_text = f"后缀: {params.get('suffix', '')}"
            elif rule == "前后添加":
                params_text = f"前缀: {params.get('prefix', '')}, 后缀: {params.get('suffix', '')}"
            elif rule == "固定值":
                params_text = f"值: {params.get('value', '')}"
            elif rule == "正则替换":
                params_text = f"模式: {params.get('pattern', '')} -> {params.get('replace', '')}"
            else:
                params_text = "未知规则"
            
            item = QTreeWidgetItem([
                config["input_col"],
                config["new_col"],
                config["rule"],
                params_text
            ])
            self.config_tree.addTopLevelItem(item)
    
    def remove_config(self):
        """移除选中配置"""
        selected_items = self.config_tree.selectedItems()
        if selected_items:
            # 获取选中项的索引
            selected_item = selected_items[0]
            index = self.config_tree.indexOfTopLevelItem(selected_item)
            
            # 从列表中删除对应配置
            if 0 <= index < len(self.output_configs):
                del self.output_configs[index]
                
                # 标记配置已修改
                self._config_modified = True
                
            # 更新显示
            self.update_config_display()
    
    def clear_configs(self):
        """清空所有配置"""
        reply = QMessageBox.question(
            self, "确认", "确定要清空所有配置吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.output_configs = []
            self.update_config_display()
            # 标记配置已修改
            self._config_modified = True
    
    def create_new_field(self):
        """创建新字段"""
        # 检查是否有选中的列
        selected_columns = self.columns_tree.selectedItems()
        
        if selected_columns:
            # 有选中的列，设置为直接复制规则
            selected_column = selected_columns[0].text(0)
            self.input_col_combobox.setCurrentText(selected_column)
            self.rule_combobox.setCurrentText("直接复制")
            self.on_rule_selected("直接复制")
            
            # 设置新列名
            self.new_col_entry.setText(f"新_{selected_column}")
            self.new_col_entry.selectAll()
        else:
            # 没有选中的列，设置为固定值规则
            self.rule_combobox.setCurrentText("固定值")
            self.on_rule_selected("固定值")
            
            # 固定值规则不需要输入列，清空输入列选择
            if self.input_col_combobox.count() > 0:
                self.input_col_combobox.setCurrentIndex(-1)
            
            # 设置新列名
            self.new_col_entry.setText("新列")
            self.new_col_entry.selectAll()
        
        # 添加新配置到列表
        self.add_config_from_form()
    
    def add_config_from_form(self):
        """从表单添加配置"""
        input_col = self.input_col_combobox.currentText()
        new_col = self.new_col_entry.text().strip()
        rule = self.rule_combobox.currentText()
        
        # 验证输入
        if not new_col:
            QMessageBox.warning(self, "警告", "请输入新列名")
            return
        
        # 对于固定值规则，不需要输入列
        if rule != "固定值" and not input_col:
            QMessageBox.warning(self, "警告", "请选择输入列")
            return
        
        # 检查新列名是否已存在
        for config in self.output_configs:
            if config["new_col"] == new_col:
                QMessageBox.warning(self, "警告", f"列名 '{new_col}' 已存在")
                return
        
        # 构建参数
        params = {}
        if rule == "前缀添加":
            params["prefix"] = self.prefix_entry.text()
        elif rule == "后缀添加":
            params["suffix"] = self.suffix_entry.text()
        elif rule == "前后添加":
            params["prefix"] = self.prefix_entry.text()
            params["suffix"] = self.suffix_entry.text()
        elif rule == "固定值":
            params["value"] = self.fixed_value_entry.text()
        
        # 添加到配置列表
        config = {
            "input_col": input_col,
            "new_col": new_col,
            "rule": rule,
            "params": params
        }
        self.output_configs.append(config)
        
        # 标记配置已修改
        self._config_modified = True
        
        # 更新配置表格
        self.update_config_display()
        
        # 选中新增的配置行
        self.select_config_item(len(self.output_configs) - 1)
    
    def select_config_item(self, index):
        """选中指定索引的配置项"""
        if 0 <= index < self.config_tree.topLevelItemCount():
            item = self.config_tree.topLevelItem(index)
            self.config_tree.setCurrentItem(item)
    
    def on_config_selected(self):
        """配置项选择变化时的处理"""
        selected_items = self.config_tree.selectedItems()
        if not selected_items:
            return
        
        # 获取选中项的索引
        selected_item = selected_items[0]
        index = self.config_tree.indexOfTopLevelItem(selected_item)
        
        if 0 <= index < len(self.output_configs):
            # 获取配置数据
            config = self.output_configs[index]
            
            # 设置初始化标志，防止递归更新
            self._is_initializing = True
            
            # 更新表单控件
            self.input_col_combobox.setCurrentText(config["input_col"])
            self.new_col_entry.setText(config["new_col"])
            self.rule_combobox.setCurrentText(config["rule"])
            
            # 触发规则选择事件，显示对应的参数控件
            self.on_rule_selected(config["rule"])
            
            # 设置参数值
            params = config["params"]
            if config["rule"] == "前缀添加":
                self.prefix_entry.setText(params.get("prefix", ""))
            elif config["rule"] == "后缀添加":
                self.suffix_entry.setText(params.get("suffix", ""))
            elif config["rule"] == "前后添加":
                self.prefix_entry.setText(params.get("prefix", ""))
                self.suffix_entry.setText(params.get("suffix", ""))
            elif config["rule"] == "固定值":
                self.fixed_value_entry.setText(params.get("value", ""))
            elif config["rule"] == "正则替换":
                self.regex_pattern_entry.setText(params.get("pattern", ""))
                self.regex_replace_entry.setText(params.get("replace", ""))
            
            # 清除初始化标志
            self._is_initializing = False
    
    def on_config_layout_changed(self):
        """配置布局变化时的处理"""
        # 根据配置树的顺序更新output_configs列表
        new_order = []
        for i in range(self.config_tree.topLevelItemCount()):
            item = self.config_tree.topLevelItem(i)
            input_col = item.text(0)
            new_col = item.text(1)
            rule = item.text(2)
            
            # 查找对应的配置
            for config in self.output_configs:
                if config["new_col"] == new_col:
                    new_order.append(config)
                    break
        
        # 更新配置列表
        if new_order:
            self.output_configs = new_order
            # 标记配置已修改
            self._config_modified = True
            # 重新更新配置树的显示，确保内容与output_configs一致
            self.update_config_display()
    
    def closeEvent(self, event):
        """窗口关闭事件处理"""
        # 检查配置是否有改动
        if self._config_modified:
            reply = QMessageBox.question(
                self, "确认退出",
                "配置已修改，是否保存？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Yes
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                # 快速保存配置
                self.quick_save_config()
                event.accept()
            elif reply == QMessageBox.StandardButton.No:
                # 不保存，直接退出
                event.accept()
            else:
                # 取消退出
                event.ignore()
        else:
            # 配置未修改，直接退出
            event.accept()
    
    def on_form_field_changed(self):
        """表单字段更改时的处理"""
        # 如果正在初始化，不执行更新
        if hasattr(self, '_is_initializing') and self._is_initializing:
            return
        
        # 检查是否有选中的配置项
        selected_items = self.config_tree.selectedItems()
        if not selected_items:
            return
        
        # 获取选中项的索引
        selected_item = selected_items[0]
        index = self.config_tree.indexOfTopLevelItem(selected_item)
        
        if 0 <= index < len(self.output_configs):
            # 更新配置数据
            config = self.output_configs[index]
            config["input_col"] = self.input_col_combobox.currentText()
            config["new_col"] = self.new_col_entry.text().strip()
            config["rule"] = self.rule_combobox.currentText()
            
            # 更新参数
            params = {}
            rule = config["rule"]
            if rule == "前缀添加":
                params["prefix"] = self.prefix_entry.text()
            elif rule == "后缀添加":
                params["suffix"] = self.suffix_entry.text()
            elif rule == "前后添加":
                params["prefix"] = self.prefix_entry.text()
                params["suffix"] = self.suffix_entry.text()
            elif rule == "固定值":
                params["value"] = self.fixed_value_entry.text()
            elif rule == "正则替换":
                params["pattern"] = self.regex_pattern_entry.text()
                params["replace"] = self.regex_replace_entry.text()
            
            config["params"] = params
            
            # 更新配置显示
            # 只更新当前选中项的显示，而不是重新构建整个配置树
            selected_item = self.config_tree.currentItem()
            if selected_item:
                # 根据规则类型生成友好的参数显示
                rule = config["rule"]
                params = config["params"]
                
                if rule == "直接复制":
                    params_text = "无参数"
                elif rule == "前缀添加":
                    params_text = f"前缀: {params.get('prefix', '')}"
                elif rule == "后缀添加":
                    params_text = f"后缀: {params.get('suffix', '')}"
                elif rule == "前后添加":
                    params_text = f"前缀: {params.get('prefix', '')}, 后缀: {params.get('suffix', '')}"
                elif rule == "固定值":
                    params_text = f"值: {params.get('value', '')}"
                elif rule == "正则替换":
                    params_text = f"模式: {params.get('pattern', '')} -> {params.get('replace', '')}"
                else:
                    params_text = "未知规则"
                
                # 更新选中项的显示
                selected_item.setText(0, config["input_col"])
                selected_item.setText(1, config["new_col"])
                selected_item.setText(2, config["rule"])
                selected_item.setText(3, params_text)
    
    def replace_variables(self, text, row_data=None):
        """替换文本中的变量，如${sheet}、${列名}、${列编号}
        
        Args:
            text: 包含变量的文本字符串
            row_data: 当前行的数据（用于替换列变量）
            
        Returns:
            str: 替换变量后的文本
        """
        if not isinstance(text, str):
            return text
        
        # 替换工作表名称变量
        text = text.replace("${sheet}", self.current_sheet)
        
        # 替换列变量（如${列名}或${A}、${B}等）
        if row_data is not None:
            import re
            # 匹配 ${列名} 或 ${A} 格式的变量
            variable_pattern = re.compile(r'\$\{([^}]+)\}')
            
            def replace_match(match):
                variable_name = match.group(1)
                # 检查是否是列名
                if variable_name in row_data:
                    value = row_data[variable_name]
                    return str(value) if value is not None else ""
                # 检查是否是列编号（如A、B、C）
                elif variable_name.isalpha() and len(variable_name) == 1:
                    # 转换列字母到列索引（A=0, B=1, 等）
                    col_index = ord(variable_name.upper()) - ord('A')
                    if 0 <= col_index < len(self.df.columns):
                        col_name = self.df.columns[col_index]
                        value = row_data[col_name]
                        return str(value) if value is not None else ""
                # 变量不存在，返回原始变量
                return match.group(0)
            
            text = variable_pattern.sub(replace_match, text)
        
        return text
    
    def generate_excel(self):
        """生成Excel文件"""
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "警告", "没有数据可以处理")
            return
        
        if not self.output_configs:
            QMessageBox.warning(self, "警告", "请先添加输出配置")
            return
        
        try:
            # 创建输出文件
            output_file, _ = QFileDialog.getSaveFileName(
                self, "保存输出文件", "", "Excel files (*.xlsx)"
            )
            
            if not output_file:
                return
            
            # 确保文件扩展名
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            # 创建结果DataFrame - 只包含配置中指定的列
            result_df = pd.DataFrame()
            
            # 应用所有配置
            for config in self.output_configs:
                input_col = config.get("input_col", "")
                new_col = config["new_col"]
                rule = config["rule"]
                params = config["params"]
                
                # 应用处理规则
                if rule == "直接复制":
                    result_df[new_col] = result_df.get(new_col, self.df[input_col])
                elif rule == "前缀添加":
                    prefix = params.get("prefix", "")
                    result_df[new_col] = prefix + self.df[input_col].astype(str)
                elif rule == "后缀添加":
                    suffix = params.get("suffix", "")
                    result_df[new_col] = self.df[input_col].astype(str) + suffix
                elif rule == "前后添加":
                    prefix = params.get("prefix", "")
                    suffix = params.get("suffix", "")
                    result_df[new_col] = prefix + self.df[input_col].astype(str) + suffix
                elif rule == "固定值":
                    value = params.get("value", "")
                    # 逐行处理变量替换
                    def process_fixed_value(row):
                        return self.replace_variables(value, row)
                    result_df[new_col] = self.df.apply(process_fixed_value, axis=1)
                elif rule == "正则替换":
                    import re
                    pattern = params.get("pattern", "")
                    replace = params.get("replace", "")
                    
                    def apply_regex_replace(text):
                        if pd.isna(text) or text == "":
                            return text
                        text_str = str(text)
                        try:
                            return re.sub(pattern, replace, text_str)
                        except:
                            return text_str
                    
                    result_df[new_col] = self.df[input_col].apply(apply_regex_replace)
            
            # 保存结果
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name=self.current_sheet, index=False)
            
            # 创建带"打开输出目录"按钮的消息框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle("成功")
            msg_box.setText(f"Excel文件已成功生成: {output_file}")
            
            # 添加"打开输出目录"按钮
            open_dir_button = msg_box.addButton("打开输出目录", QMessageBox.ButtonRole.ActionRole)
            ok_button = msg_box.addButton(QMessageBox.StandardButton.Ok)
            
            # 显示消息框并获取用户选择
            msg_box.exec()
            
            # 如果用户点击了"打开输出目录"按钮
            if msg_box.clickedButton() == open_dir_button:
                try:
                    # 获取输出文件所在的目录
                    output_dir = os.path.dirname(output_file)
                    
                    # 调用统一方法打开输出目录
                    self.open_output_directory(output_dir)
                except Exception as e:
                    print(f"[错误] 打开输出目录失败: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    QMessageBox.warning(self, "警告", f"打开输出目录时发生错误: {str(e)}")
            
            self.status_bar.showMessage(f"已生成: {os.path.basename(output_file)}")
            
        except Exception as e:
            error_msg = f"生成Excel失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def quick_save_config(self):
        """快速保存配置到当前打开的配置文件或默认文件"""
        if not self.output_configs:
            QMessageBox.warning(self, "警告", "没有配置可以保存")
            return
        
        try:
            # 确定保存路径
            if self.current_config_path:
                # 如果有当前打开的配置文件，保存到该文件
                config_file = self.current_config_path
            else:
                # 否则保存到default.json
                config_file = "default.json"
            
            # 保存配置
            config_data = {
                "output_configs": self.output_configs,
                "file_path": self.file_path,
                "current_sheet": self.current_sheet
            }
            
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(self, "成功", f"配置已快速保存到: {config_file}")
            # 重置配置修改标记
            self._config_modified = False
            self.status_bar.showMessage(f"配置已快速保存到: {config_file}")
            
        except Exception as e:
            error_msg = f"快速保存配置失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def save_config(self):
        """保存配置到文件"""
        if not self.output_configs:
            QMessageBox.warning(self, "警告", "没有配置可以保存")
            return
        
        try:
            config_file, _ = QFileDialog.getSaveFileName(
                self, "保存配置文件", "", "配置文件 (*.json)"
            )
            
            if not config_file:
                return
            
            # 确保文件扩展名
            if not config_file.endswith('.json'):
                config_file += '.json'
            
            # 保存配置
            config_data = {
                "output_configs": self.output_configs,
                "file_path": self.file_path,
                "current_sheet": self.current_sheet
            }
            
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            
            # 更新当前配置文件路径
            self.current_config_path = config_file
            
            # 重置配置修改标记
            self._config_modified = False
            
            QMessageBox.information(self, "成功", f"配置已成功保存: {config_file}")
            
        except Exception as e:
            error_msg = f"保存配置失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def load_config(self):
        """从文件加载配置"""
        try:
            config_file, _ = QFileDialog.getOpenFileName(
                self, "加载配置文件", "", "配置文件 (*.json)"
            )
            
            if not config_file:
                return
            
            # 保存当前配置文件路径
            self.current_config_path = config_file
            
            # 读取配置
            with open(config_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # 加载配置数据
            self.output_configs = config_data.get("output_configs", [])
            self.update_config_display()
            
            # 重置配置修改标记
            self._config_modified = False
            
            # 如果配置中包含文件路径，尝试加载文件
            file_path = config_data.get("file_path", "")
            current_sheet = config_data.get("current_sheet", "")
            
            if file_path and os.path.exists(file_path):
                # 检查是否需要加载文件
                if self.file_path != file_path or not self.df:
                    self.file_path = file_path
                    
                    # 读取所有sheet名称
                    excel_file = pd.ExcelFile(file_path)
                    self.sheet_names = excel_file.sheet_names
                    
                    # 更新sheet下拉框
                    self.sheet_combobox.clear()
                    self.sheet_combobox.addItems(self.sheet_names)
                    self.sheet_combobox.setEnabled(True)
                    
                    # 加载指定的sheet或默认第一个
                    if current_sheet in self.sheet_names:
                        index = self.sheet_names.index(current_sheet)
                        if index >= 0:
                            self.sheet_combobox.setCurrentIndex(index)
                        self.load_sheet_data(current_sheet)
                    elif self.sheet_names:
                        self.load_sheet_data(self.sheet_names[0])        
        except Exception as e:
            error_msg = f"加载配置失败: {str(e)}"
            print(f"[错误] {error_msg}")  # 打印到控制台
            QMessageBox.critical(self, "错误", error_msg)
    
    def show_about(self):
        """显示关于信息"""
        QMessageBox.about(
            self, "关于Excel处理工具",
            "Excel处理工具 v0.9\n\n" +
            "功能：读取Excel文件，可视化配置字段处理规则，生成新的Excel文件\n\n" +
            "作者：jadedrip\n" +
            ""
        )
    
    def open_output_directory(self, directory):
        """打开指定的输出目录
        
        Args:
            directory: 要打开的目录路径
        """
        try:
            # 转换为绝对路径
            output_dir_abs = os.path.abspath(directory)
            print(f"[调试] 尝试打开输出目录: {output_dir_abs}")
            print(f"[调试] 目录是否存在: {os.path.exists(output_dir_abs)}")
            print(f"[调试] 系统平台: {sys.platform}")
            
            if os.path.exists(output_dir_abs):
                # Windows系统打开资源管理器
                if sys.platform.startswith('win'):
                    # 使用start命令确保路径中的空格被正确处理
                    cmd = f'explorer "{output_dir_abs}"'
                    print(f"[调试] 执行命令: {cmd}")
                    
                    # 使用shell=True，捕获输出和错误
                    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
                    print(f"[调试] 命令返回码: {result.returncode}")
                    print(f"[调试] 命令输出: {result.stdout}")
                    print(f"[调试] 命令错误: {result.stderr}")
                else:
                    # 其他系统（Linux/Mac）
                    subprocess.run(['xdg-open', output_dir_abs])
            else:
                print(f"[错误] 输出目录不存在: {output_dir_abs}")
                # 尝试创建目录
                try:
                    os.makedirs(output_dir_abs, exist_ok=True)
                    print(f"[信息] 已自动创建输出目录: {output_dir_abs}")
                    # 再次尝试打开
                    if sys.platform.startswith('win'):
                        cmd = f'explorer "{output_dir_abs}"'
                        subprocess.run(cmd, shell=True)
                except Exception as create_e:
                    print(f"[错误] 创建输出目录失败: {str(create_e)}")
                    QMessageBox.warning(self, "警告", f"输出目录不存在，且无法创建: {output_dir_abs}")
        except Exception as e:
            error_msg = f"打开输出目录失败: {str(e)}"
            print(f"[错误] {error_msg}")
            import traceback
            traceback.print_exc()
            QMessageBox.warning(self, "警告", error_msg)