#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置树控件类
功能：确保拖放只调整顺序而不创建层次结构
作者：jadedrip
"""

from PyQt6.QtWidgets import QTreeWidget, QAbstractItemView
from PyQt6.QtCore import QPoint

class ConfigTreeWidget(QTreeWidget):
    """配置树控件，确保拖放只调整顺序而不创建层次结构"""
    def __init__(self, parent=None):
        super().__init__(parent)
        # 启用拖放功能
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        # 设置为内部移动模式
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
    
    def dropEvent(self, event):
        """处理放置事件，确保只调整顺序"""
        # 获取拖动的项目
        source_item = self.currentItem()
        if not source_item:
            event.ignore()
            return
        
        # 获取目标位置
        # 在PyQt6中使用position()而不是pos()
        pos = event.position().toPoint()
        target_item = self.itemAt(pos)
        
        # 先计算插入位置
        if target_item:
            # 获取目标项的索引
            target_index = self.indexOfTopLevelItem(target_item)
            # 如果拖动到目标项下方，插入位置为target_index + 1
            rect = self.visualItemRect(target_item)
            if pos.y() > rect.y() + rect.height() / 2:
                insert_index = target_index + 1
            else:
                insert_index = target_index
        else:
            # 拖到空白处，插入到末尾
            insert_index = self.topLevelItemCount()
        
        # 移除原项目
        source_index = self.indexOfTopLevelItem(source_item)
        if source_index == -1:
            event.ignore()
            return
        
        # 从模型中移除
        self.takeTopLevelItem(source_index)
        
        # 调整插入位置（如果源项目在目标项目之前，目标索引需要减1）
        if source_index < insert_index:
            insert_index -= 1
        
        # 确保插入位置有效
        if insert_index < 0:
            insert_index = 0
        elif insert_index > self.topLevelItemCount():
            insert_index = self.topLevelItemCount()
        
        # 插入到新位置
        self.insertTopLevelItem(insert_index, source_item)
        
        # 触发rowsMoved信号
        model = self.model()
        if model:
            # 发送信号通知顺序变化
            model.layoutChanged.emit()
        
        event.accept()