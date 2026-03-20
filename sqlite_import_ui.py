#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
SQLite导入UI模块
功能：提供Excel导入SQLite数据库的用户界面
作者：jadedrip

"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, 
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
    QComboBox, QCheckBox, QProgressBar, QFileDialog
)
from PyQt6.QtCore import Qt


class SQLiteImportUI(QWidget):
    """SQLite导入UI类"""
    
    def __init__(self):
        """初始化UI"""
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        """初始化用户界面"""
        # 创建主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # 创建数据库选择区域
        self.create_database_section(main_layout)
        
        # 创建字段映射与类型设置区域
        self.create_field_mapping_section(main_layout)
        
        # 创建操作按钮区域
        self.create_operation_buttons_section(main_layout)
        
        # 创建导入进度区域
        self.create_progress_section(main_layout)
    
    def create_database_section(self, parent_layout):
        """创建数据库选择区域"""
        database_group = QGroupBox("数据库选择")
        database_layout = QVBoxLayout(database_group)
        
        # 数据库文件路径
        db_path_layout = QHBoxLayout()
        db_path_layout.addWidget(QLabel("SQLite数据库:"))
        self.database_path_entry = QLineEdit()
        self.database_path_entry.setReadOnly(True)
        self.database_path_entry.setMinimumWidth(300)
        db_path_layout.addWidget(self.database_path_entry)
        
        # 选择数据库按钮
        self.select_db_button = QPushButton("选择数据库")
        db_path_layout.addWidget(self.select_db_button)
        
        # 新建数据库按钮
        self.create_db_button = QPushButton("新建数据库")
        db_path_layout.addWidget(self.create_db_button)
        
        database_layout.addLayout(db_path_layout)
        
        # 表名输入框
        table_name_layout = QHBoxLayout()
        table_name_layout.addWidget(QLabel("表名:"))
        self.table_name_entry = QLineEdit()
        self.table_name_entry.setPlaceholderText("默认使用工作表名")
        table_name_layout.addWidget(self.table_name_entry)
        
        database_layout.addLayout(table_name_layout)
        
        parent_layout.addWidget(database_group)
    
    def create_field_mapping_section(self, parent_layout):
        """创建字段映射与类型设置区域"""
        mapping_group = QGroupBox("字段映射与类型设置")
        mapping_layout = QVBoxLayout(mapping_group)
        
        # 字段映射表格
        self.field_table = QTableWidget()
        self.field_table.setColumnCount(5)
        self.field_table.setHorizontalHeaderLabels(["原始列", "目标字段", "数据类型", "导入", "索引"])
        self.field_table.setColumnWidth(0, 120)
        self.field_table.setColumnWidth(1, 120)
        self.field_table.setColumnWidth(2, 100)
        self.field_table.setColumnWidth(3, 60)
        self.field_table.setColumnWidth(4, 100)
        mapping_layout.addWidget(self.field_table)
        
        # 按钮区域（移除了自动映射和清空映射按钮）
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        mapping_layout.addLayout(buttons_layout)
        
        parent_layout.addWidget(mapping_group)
    
    def create_operation_buttons_section(self, parent_layout):
        """创建操作按钮区域"""
        buttons_widget = QWidget()
        buttons_layout = QHBoxLayout(buttons_widget)
        
        buttons_layout.addStretch()
        
        # 开始导入按钮
        self.import_button = QPushButton("开始导入")
        self.import_button.setEnabled(False)
        buttons_layout.addWidget(self.import_button)
        
        parent_layout.addWidget(buttons_widget)
    
    def create_progress_section(self, parent_layout):
        """创建导入进度区域"""
        progress_group = QGroupBox("导入进度")
        progress_layout = QVBoxLayout(progress_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        # 状态信息
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.status_label)
        
        parent_layout.addWidget(progress_group)
    
    def add_field_row(self, original_column, target_field, data_type, import_flag=True, index_type="无"):
        """添加字段行
        
        Args:
            original_column: 原始列名
            target_field: 目标字段名
            data_type: 数据类型
            import_flag: 是否导入
            index_type: 索引类型（无、主键、唯一索引、普通索引）
        """
        row_position = self.field_table.rowCount()
        self.field_table.insertRow(row_position)
        
        # 原始列（设置为只读）
        original_item = QTableWidgetItem(original_column)
        original_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
        self.field_table.setItem(row_position, 0, original_item)
        
        # 目标字段
        self.field_table.setItem(row_position, 1, QTableWidgetItem(target_field))
        
        # 数据类型下拉框
        type_combo = QComboBox()
        type_combo.addItems(["TEXT", "INTEGER", "REAL", "BLOB"])
        type_combo.setCurrentText(data_type)
        self.field_table.setCellWidget(row_position, 2, type_combo)
        
        # 导入复选框
        import_checkbox = QCheckBox()
        import_checkbox.setChecked(import_flag)
        # QCheckBox没有setAlignment方法，移除该调用
        self.field_table.setCellWidget(row_position, 3, import_checkbox)
        
        # 索引类型下拉框
        index_combo = QComboBox()
        index_combo.addItems(["无", "主键", "唯一索引", "普通索引"])
        index_combo.setCurrentText(index_type)
        self.field_table.setCellWidget(row_position, 4, index_combo)
    
    def clear_field_table(self):
        """清空字段表格"""
        self.field_table.setRowCount(0)
    
    def get_field_mappings(self):
        """获取字段映射
        
        Returns:
            list: 字段映射列表
        """
        mappings = []
        for row in range(self.field_table.rowCount()):
            original_column = self.field_table.item(row, 0).text()
            target_field = self.field_table.item(row, 1).text()
            data_type = self.field_table.cellWidget(row, 2).currentText()
            import_flag = self.field_table.cellWidget(row, 3).isChecked()
            index_type = self.field_table.cellWidget(row, 4).currentText()
            
            mappings.append({
                "original_column": original_column,
                "target_field": target_field,
                "data_type": data_type,
                "import_flag": import_flag,
                "index_type": index_type
            })
        return mappings
    
    def set_database_path(self, path):
        """设置数据库路径
        
        Args:
            path: 数据库文件路径
        """
        self.database_path_entry.setText(path)
    
    def get_database_path(self):
        """获取数据库路径
        
        Returns:
            str: 数据库文件路径
        """
        return self.database_path_entry.text()
    
    def set_table_name(self, name):
        """设置表名
        
        Args:
            name: 表名
        """
        self.table_name_entry.setText(name)
    
    def get_table_name(self):
        """获取表名
        
        Returns:
            str: 表名
        """
        return self.table_name_entry.text()
    
    def update_progress(self, value, status):
        """更新进度
        
        Args:
            value: 进度值 (0-100)
            status: 状态信息
        """
        self.progress_bar.setValue(value)
        self.status_label.setText(status)
    
    def reset_progress(self):
        """重置进度"""
        self.progress_bar.setValue(0)
        self.status_label.setText("就绪")
    
    def enable_import_button(self, enabled):
        """启用/禁用导入按钮
        
        Args:
            enabled: 是否启用
        """
        self.import_button.setEnabled(enabled)