#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
可拖拽的树形控件类
功能：用于左侧列信息展示，支持拖拽功能
作者：jadedrip

"""

from PyQt6.QtWidgets import QTreeWidget, QAbstractItemView
from PyQt6.QtCore import Qt, QMimeData, QPoint
from PyQt6.QtGui import QDrag, QPixmap, QPainter, QColor, QPen, QFont, QFontMetrics

class DraggableTreeWidget(QTreeWidget):
    """可拖拽的树形控件，用于左侧列信息展示"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
    
    def startDrag(self, supportedActions):
        """开始拖拽操作"""
        item = self.currentItem()
        if not item:
            return
        
        # 获取要拖拽的数据：Excel列名（A, B, C...）和实际列名
        excel_col = item.text(0)
        actual_col = item.text(1)
        
        # 创建MIME数据
        mime_data = QMimeData()
        mime_data.setText(f"{excel_col}:{actual_col}")
        
        # 创建拖拽对象
        drag = QDrag(self)
        drag.setMimeData(mime_data)
        
        # 设置拖拽时的图标 - 美化显示
        
        # 要显示的文本
        display_text = f"{excel_col}: {actual_col}"
        
        # 创建字体
        font = QFont("SimHei", 10)
        
        # 计算文本宽度和高度
        font_metrics = QFontMetrics(font)
        text_width = font_metrics.horizontalAdvance(display_text) + 20  # 左右各留10px边距
        text_height = font_metrics.height() + 10  # 上下各留5px边距
        
        # 创建合适大小的Pixmap
        pixmap = QPixmap(text_width, text_height)
        pixmap.fill(QColor(0, 0, 0, 0))  # 完全透明背景
        
        # 创建画家对象
        painter = QPainter(pixmap)
        painter.setFont(font)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)  # 抗锯齿
        
        # 绘制带圆角的淡蓝色填充
        fill_color = QColor(220, 230, 255, 200)  # 淡蓝色填充，透明度200/255
        painter.setBrush(fill_color)
        
        # 绘制带圆角的边框
        border_color = QColor(150, 180, 255, 220)  # 淡蓝色边框，透明度220/255
        pen = QPen(border_color, 1)
        painter.setPen(pen)
        
        # 绘制圆角矩形，圆角半径为6px（使用fillRect和drawRoundedRect组合）
        painter.drawRoundedRect(0, 0, text_width - 1, text_height - 1, 6, 6)
        
        # 绘制文本
        text_color = QColor(0, 0, 0, 230)  # 黑色文本，透明度230/255
        painter.setPen(text_color)
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, display_text)
        
        # 结束绘制
        painter.end()
        
        # 设置拖拽图标
        drag.setPixmap(pixmap)
        drag.setHotSpot(QPoint(text_width // 2, text_height // 2))  # 设置热点为中心
        
        # 执行拖拽
        drag.exec(supportedActions)