#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
可接收拖拽的列表控件类
功能：用于右侧切分列选择，支持拖拽功能
作者：jadedrip

"""

from PyQt6.QtWidgets import QListWidget, QAbstractItemView
from PyQt6.QtCore import QMimeData

class DroppableListWidget(QListWidget):
    """可接收拖拽的列表控件，用于右侧切分列选择"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
    
    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.source() == self:
            # 内部拖动，接受事件
            event.acceptProposedAction()
        elif event.mimeData().hasText():
            # 外部拖拽，检查数据格式
            text = event.mimeData().text()
            if ":" in text:
                event.acceptProposedAction()
    
    def dragMoveEvent(self, event):
        """拖拽移动事件"""
        if event.source() == self:
            # 内部拖动，接受事件
            event.acceptProposedAction()
        elif event.mimeData().hasText():
            # 外部拖拽，接受事件
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """拖拽释放事件"""
        if event.source() == self:
            # 内部移动，使用默认行为
            super().dropEvent(event)
        else:
            # 外部拖拽，处理数据
            text = event.mimeData().text()
            if ":" in text:
                # 格式：Excel列名:实际列名
                excel_col, actual_col = text.split(":", 1)
                
                # 创建列表项，显示为"Excel列名(实际列名)"格式
                item_text = f"{excel_col} ({actual_col})"
                
                # 检查是否已存在相同的项
                for i in range(self.count()):
                    if self.item(i).text() == item_text:
                        return  # 已存在，不重复添加
                
                # 添加新项
                self.addItem(item_text)
                event.acceptProposedAction()