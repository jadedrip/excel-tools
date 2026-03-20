#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
SQLite导入功能模块
功能：实现Excel数据导入SQLite数据库的核心功能
作者：jadedrip

"""

import os
import sqlite3
import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal
from sqlite_import_ui import SQLiteImportUI


class SQLiteImportWorker(QThread):
    """SQLite导入工作线程"""
    
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, excel_file, sheet_name, database_path, field_mappings, table_name):
        """初始化工作线程
        
        Args:
            excel_file: Excel文件路径
            sheet_name: 工作表名称
            database_path: 数据库文件路径
            field_mappings: 字段映射列表
            table_name: 表名
        """
        super().__init__()
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.database_path = database_path
        self.field_mappings = field_mappings
        self.table_name = table_name
        self.is_cancelled = False
    
    def run(self):
        """执行导入操作"""
        try:
            # 读取Excel数据
            self.progress_updated.emit(10, "正在读取Excel数据...")
            df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
            
            # 过滤需要导入的字段
            import_fields = [m for m in self.field_mappings if m["import_flag"]]
            if not import_fields:
                self.finished.emit(False, "没有选择要导入的字段")
                return
            
            # 准备导入数据
            self.progress_updated.emit(20, "正在准备导入数据...")
            import_data = []
            column_mapping = {}
            
            for mapping in import_fields:
                original_col = mapping["original_column"]
                target_field = mapping["target_field"]
                if original_col in df.columns:
                    column_mapping[original_col] = target_field
            
            # 连接数据库
            self.progress_updated.emit(30, "正在连接数据库...")
            conn = sqlite3.connect(self.database_path)
            cursor = conn.cursor()
            
            # 检查表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (self.table_name,))
            table_exists = cursor.fetchone() is not None
            
            if table_exists:
                # 表已存在，更新表结构
                self.progress_updated.emit(40, "表已存在，正在更新表结构...")
                self._update_table_structure(conn, cursor, self.table_name, import_fields)
            else:
                # 表不存在，创建新表
                self.progress_updated.emit(40, "正在创建数据库表...")
                create_table_sql = self._generate_create_table_sql(self.table_name, import_fields)
                cursor.execute(create_table_sql)
                
                # 创建索引
                index_sql_list = self._generate_index_sql(self.table_name, import_fields)
                for index_sql in index_sql_list:
                    cursor.execute(index_sql)
            
            # 批量导入数据
            self.progress_updated.emit(50, "正在导入数据...")
            total_rows = len(df)
            batch_size = 1000
            
            for i in range(0, total_rows, batch_size):
                if self.is_cancelled:
                    conn.close()
                    self.finished.emit(False, "导入操作已取消")
                    return
                
                batch_end = min(i + batch_size, total_rows)
                batch_df = df.iloc[i:batch_end]
                
                # 准备插入数据
                insert_sql = self._generate_insert_sql(self.table_name, import_fields)
                values = []
                
                for _, row in batch_df.iterrows():
                    row_values = []
                    for mapping in import_fields:
                        original_col = mapping["original_column"]
                        if original_col in row:
                            value = row[original_col]
                            # 处理空值
                            if pd.isna(value):
                                row_values.append(None)
                            else:
                                row_values.append(value)
                        else:
                            row_values.append(None)
                    values.append(tuple(row_values))
                
                # 执行批量插入
                cursor.executemany(insert_sql, values)
                conn.commit()
                
                # 更新进度
                progress = 50 + int((batch_end / total_rows) * 50)
                self.progress_updated.emit(progress, f"已导入 {batch_end}/{total_rows} 条记录")
            
            # 关闭连接
            conn.close()
            
            self.finished.emit(True, f"导入成功！共导入 {total_rows} 条记录")
            
        except Exception as e:
            self.finished.emit(False, f"导入失败: {str(e)}")
    
    def _update_table_structure(self, conn, cursor, table_name, field_mappings):
        """更新表结构
        
        Args:
            conn: 数据库连接
            cursor: 游标
            table_name: 表名
            field_mappings: 字段映射列表
        """
        # 获取现有表结构
        cursor.execute(f"PRAGMA table_info({table_name})")
        existing_columns = {row[1]: row[2] for row in cursor.fetchall()}
        
        # 添加缺失的字段
        for mapping in field_mappings:
            field_name = mapping["target_field"]
            data_type = mapping["data_type"]
            
            if field_name not in existing_columns:
                # 添加新字段
                alter_sql = f"ALTER TABLE `{table_name}` ADD COLUMN `{field_name}` {data_type}"
                cursor.execute(alter_sql)
        
        # 创建索引
        index_sql_list = self._generate_index_sql(table_name, field_mappings)
        for index_sql in index_sql_list:
            try:
                cursor.execute(index_sql)
            except:
                # 索引可能已存在，忽略错误
                pass
    
    def cancel(self):
        """取消导入操作"""
        self.is_cancelled = True
    
    def _generate_create_table_sql(self, table_name, field_mappings):
        """生成创建表的SQL语句
        
        Args:
            table_name: 表名
            field_mappings: 字段映射列表
            
        Returns:
            str: 创建表的SQL语句
        """
        columns = []
        primary_keys = []
        
        for mapping in field_mappings:
            field_name = mapping["target_field"]
            data_type = mapping["data_type"]
            index_type = mapping["index_type"]
            
            column_def = f"`{field_name}` {data_type}"
            
            # 处理主键
            if index_type == "主键":
                column_def += " PRIMARY KEY"
                primary_keys.append(field_name)
            # 处理唯一索引
            elif index_type == "唯一索引":
                column_def += " UNIQUE"
            
            columns.append(column_def)
        
        columns_sql = ", ".join(columns)
        return f"CREATE TABLE IF NOT EXISTS `{table_name}` ({columns_sql})"
    
    def _generate_index_sql(self, table_name, field_mappings):
        """生成创建索引的SQL语句
        
        Args:
            table_name: 表名
            field_mappings: 字段映射列表
            
        Returns:
            list: 创建索引的SQL语句列表
        """
        index_sql_list = []
        
        for mapping in field_mappings:
            field_name = mapping["target_field"]
            index_type = mapping["index_type"]
            
            if index_type == "普通索引":
                index_name = f"idx_{table_name}_{field_name}"
                index_sql = f"CREATE INDEX IF NOT EXISTS `{index_name}` ON `{table_name}` (`{field_name}`)"
                index_sql_list.append(index_sql)
            elif index_type == "唯一索引":
                # 唯一索引已经在字段定义中处理
                pass
            elif index_type == "主键":
                # 主键已经在字段定义中处理
                pass
        
        return index_sql_list
    
    def _generate_insert_sql(self, table_name, field_mappings):
        """生成插入数据的SQL语句
        
        Args:
            table_name: 表名
            field_mappings: 字段映射列表
            
        Returns:
            str: 插入数据的SQL语句
        """
        field_names = [f"`{m['target_field']}`" for m in field_mappings]
        placeholders = ["?" for _ in field_mappings]
        
        fields_sql = ", ".join(field_names)
        placeholders_sql = ", ".join(placeholders)
        
        return f"INSERT INTO `{table_name}` ({fields_sql}) VALUES ({placeholders_sql})"


class SQLiteImportManager:
    """SQLite导入管理器"""
    
    def __init__(self, main_window):
        """初始化导入管理器
        
        Args:
            main_window: 主窗口对象
        """
        self.main_window = main_window
        self.ui = SQLiteImportUI()
        self.worker = None
        self.excel_file = ""
        self.sheet_name = ""
        self.setup_connections()
    
    def setup_connections(self):
        """设置信号连接"""
        # 数据库选择按钮
        self.ui.select_db_button.clicked.connect(self.select_database)
        self.ui.create_db_button.clicked.connect(self.create_database)
        
        # 操作按钮
        self.ui.import_button.clicked.connect(self.start_import)
    
    def select_database(self):
        """选择数据库文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self.main_window, "选择SQLite数据库文件", "", "SQLite文件 (*.db *.sqlite)")
        if file_path:
            self.ui.set_database_path(file_path)
            self.check_import_ready()
    
    def create_database(self):
        """创建数据库文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self.main_window, "创建SQLite数据库文件", "", "SQLite文件 (*.db *.sqlite)")
        if file_path:
            # 确保文件扩展名正确
            if not (file_path.endswith('.db') or file_path.endswith('.sqlite')):
                file_path += '.db'
            
            # 创建空数据库文件
            try:
                conn = sqlite3.connect(file_path)
                conn.close()
                self.ui.set_database_path(file_path)
                self.check_import_ready()
            except Exception as e:
                QMessageBox.warning(self.main_window, "错误", f"创建数据库失败: {str(e)}")
    
    def update_excel_info(self, excel_file, sheet_name, columns):
        """更新Excel信息
        
        Args:
            excel_file: Excel文件路径
            sheet_name: 工作表名称
            columns: 列名列表
        """
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        
        # 设置表名（默认使用工作表名）
        table_name = sheet_name.replace(' ', '_').replace('-', '_')
        self.ui.set_table_name(table_name)
        
        # 清空字段表格
        self.ui.clear_field_table()
        
        # 添加字段行
        for column in columns:
            # 自动生成目标字段名（去除特殊字符，转换为小写）
            target_field = column.lower().replace(' ', '_').replace('-', '_')
            # 自动检测数据类型
            data_type = self._detect_data_type(column)
            self.ui.add_field_row(column, target_field, data_type)
        
        self.check_import_ready()
    
    def _detect_data_type(self, column_name):
        """检测数据类型
        
        Args:
            column_name: 列名
            
        Returns:
            str: 数据类型
        """
        # 简单的类型检测逻辑
        column_lower = column_name.lower()
        if any(keyword in column_lower for keyword in ['id', 'number', 'count', 'amount']):
            return "INTEGER"
        elif any(keyword in column_lower for keyword in ['price', 'value', 'rate', 'score']):
            return "REAL"
        else:
            return "TEXT"
    
    def start_import(self):
        """开始导入"""
        if not self.excel_file:
            QMessageBox.warning(self.main_window, "警告", "请先选择Excel文件")
            return
        
        if not self.sheet_name:
            QMessageBox.warning(self.main_window, "警告", "请先选择工作表")
            return
        
        database_path = self.ui.get_database_path()
        if not database_path:
            QMessageBox.warning(self.main_window, "警告", "请选择或创建SQLite数据库")
            return
        
        # 获取字段映射
        field_mappings = self.ui.get_field_mappings()
        if not any(m["import_flag"] for m in field_mappings):
            QMessageBox.warning(self.main_window, "警告", "请至少选择一个要导入的字段")
            return
        
        # 获取表名
        table_name = self.ui.get_table_name()
        if not table_name:
            # 如果表名为空，使用工作表名
            table_name = self.sheet_name.replace(' ', '_').replace('-', '_')
        
        # 创建并启动工作线程
        self.worker = SQLiteImportWorker(
            self.excel_file,
            self.sheet_name,
            database_path,
            field_mappings,
            table_name
        )
        
        # 连接信号
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.finished.connect(self.import_finished)
        
        # 禁用按钮
        self.ui.import_button.setEnabled(False)
        
        # 启动线程
        self.worker.start()
    
    def update_progress(self, value, status):
        """更新进度"""
        self.ui.update_progress(value, status)
    
    def import_finished(self, success, message):
        """导入完成
        
        Args:
            success: 是否成功
            message: 消息
        """
        if success:
            QMessageBox.information(self.main_window, "成功", message)
        else:
            QMessageBox.warning(self.main_window, "失败", message)
        
        # 重置UI
        self.ui.reset_progress()
        self.ui.import_button.setEnabled(True)
    
    def check_import_ready(self):
        """检查是否可以开始导入"""
        database_path = self.ui.get_database_path()
        has_excel = bool(self.excel_file)
        has_database = bool(database_path)
        has_fields = self.ui.field_table.rowCount() > 0
        
        self.ui.enable_import_button(has_excel and has_database and has_fields)