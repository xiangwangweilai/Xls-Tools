#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格分析与选择性分类工具 - PyQt6 版
功能：打开表格 -> 分析类目 -> 勾选需要的列 -> 另存为新表格
"""

import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea,
    QCheckBox, QGroupBox, QSplitter, QComboBox, QStatusBar,
    QDialog, QTextBrowser
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("表格分析与选择性分类工具")
        self.resize(1100, 700)

        self.df = None
        self.file_path = None
        self.column_checkboxes = []

        self._build_ui()
        self.statusBar().showMessage("就绪 - 请打开一个表格文件")

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        main_v = QVBoxLayout(central)
        main_v.setSpacing(8)
        main_v.setContentsMargins(10, 10, 10, 10)

        # 左上角工具栏：文件操作 + 保存设置
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)

        self.btn_open = QPushButton("打开表格")
        self.btn_open.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 4px 12px; }"
        )
        self.btn_open.clicked.connect(self.open_file)
        top_bar.addWidget(self.btn_open)

        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: #888; font-size: 12px;")
        top_bar.addWidget(self.file_label)

        top_bar.addWidget(QLabel("保存格式:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Excel (.xlsx)", "CSV (.csv)"])
        top_bar.addWidget(self.format_combo)

        self.btn_save = QPushButton("保存选中的列到新文件")
        self.btn_save.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 4px 12px; }"
        )
        self.btn_save.clicked.connect(self.save_selected)
        self.btn_save.setEnabled(False)
        top_bar.addWidget(self.btn_save)

        top_bar.addStretch()

        main_v.addLayout(top_bar)

        # 中间：分割器（左侧类目选择，右侧数据预览）
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(3)

        # 左侧：类目选择面板
        left_panel = QWidget()
        left_lay = QVBoxLayout(left_panel)
        left_lay.setSpacing(6)
        left_lay.setContentsMargins(4, 4, 4, 4)

        col_group = QGroupBox("类目选择（勾选需要保留的列）")
        col_lay = QVBoxLayout(col_group)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setMinimumWidth(200)

        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_layout.setSpacing(4)
        self.scroll_layout.setContentsMargins(4, 4, 4, 4)

        self.scroll.setWidget(self.scroll_content)
        col_lay.addWidget(self.scroll)

        # 共X列 + 全选/取消全选 放在同一行
        bottom_row = QHBoxLayout()
        self.col_count_label = QLabel("共 0 列")
        self.col_count_label.setStyleSheet("color: #888; font-size: 11px;")
        bottom_row.addWidget(self.col_count_label)

        self.btn_select_all = QPushButton("全选")
        self.btn_select_all.clicked.connect(lambda: self._set_all_checks(True))
        self.btn_select_all.setEnabled(False)
        bottom_row.addWidget(self.btn_select_all)

        self.btn_deselect_all = QPushButton("取消全选")
        self.btn_deselect_all.clicked.connect(lambda: self._set_all_checks(False))
        self.btn_deselect_all.setEnabled(False)
        bottom_row.addWidget(self.btn_deselect_all)

        bottom_row.addStretch()
        col_lay.addLayout(bottom_row)

        left_lay.addWidget(col_group)

        # 右侧：数据预览
        right_panel = QWidget()
        right_lay = QVBoxLayout(right_panel)
        right_lay.setSpacing(4)
        right_lay.setContentsMargins(4, 4, 4, 4)

        preview_group = QGroupBox("数据预览")
        preview_lay = QVBoxLayout(preview_group)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(
            "QTableWidget { gridline-color: #ddd; }"
            "QTableWidget::item { padding: 4px; }"
            "QHeaderView::section { background-color: #f0f0f0; padding: 4px; border: 1px solid #ccc; }"
        )
        preview_lay.addWidget(self.table)

        right_lay.addWidget(preview_group, 1)

        # 右下角：免责声明（放在右侧面板底部）
        disc_row = QHBoxLayout()
        disc_row.setContentsMargins(0, 0, 0, 0)
        disc_row.addStretch()
        self.btn_disclaimer = QPushButton("免责声明")
        self.btn_disclaimer.setStyleSheet("color: red; font-size: 13px; border: none;")
        self.btn_disclaimer.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_disclaimer.clicked.connect(self.show_disclaimer)
        disc_row.addWidget(self.btn_disclaimer)
        right_lay.addLayout(disc_row, 0)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([250, 850])

        main_v.addWidget(splitter, 1)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "打开表格文件", "",
            "Excel/CSV Files (*.xlsx *.xls *.csv)"
        )
        if not file_path:
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".csv":
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)

            self.file_path = file_path
            self.file_label.setText(f"已加载: {os.path.basename(file_path)} ({len(self.df)} 行, {len(self.df.columns)} 列)")
            self.file_label.setStyleSheet("color: #4CAF50; font-size: 12px;")

            self._build_column_checkboxes()
            self._preview_data()

            self.btn_select_all.setEnabled(True)
            self.btn_deselect_all.setEnabled(True)
            self.btn_save.setEnabled(True)

            self.statusBar().showMessage(f"已加载 {len(self.df)} 行 x {len(self.df.columns)} 列", 5000)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开文件失败:\n{e}")

    def _build_column_checkboxes(self):
        # 清除旧的复选框
        for cb in self.column_checkboxes:
            cb.deleteLater()
        self.column_checkboxes = []

        # 添加全选/取消全选提示
        hint = QLabel("提示：勾选需要保留的列")
        hint.setStyleSheet("color: #666; font-size: 11px; font-style: italic;")
        self.scroll_layout.addWidget(hint)

        for col in self.df.columns:
            cb = QCheckBox(str(col))
            cb.setChecked(True)
            cb.setStyleSheet("QCheckBox { padding: 4px; font-size: 12px; }")
            self.scroll_layout.addWidget(cb)
            self.column_checkboxes.append(cb)

        self.col_count_label.setText(f"共 {len(self.df.columns)} 列")

    def _preview_data(self):
        # 显示前 100 行数据
        preview_df = self.df.head(100)
        self.table.setRowCount(len(preview_df))
        self.table.setColumnCount(len(preview_df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in preview_df.columns])

        for i in range(len(preview_df)):
            for j in range(len(preview_df.columns)):
                val = preview_df.iloc[i, j]
                item = QTableWidgetItem(str(val) if pd.notna(val) else "")
                self.table.setItem(i, j, item)

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setVisible(True)

    def _set_all_checks(self, checked: bool):
        for cb in self.column_checkboxes:
            cb.setChecked(checked)

    def save_selected(self):
        if self.df is None:
            QMessageBox.warning(self, "警告", "请先打开一个表格文件！")
            return

        # 获取选中的列
        selected_cols = [cb.text() for cb in self.column_checkboxes if cb.isChecked()]

        if not selected_cols:
            QMessageBox.warning(self, "警告", "请至少选择一列！")
            return

        # 选择保存路径
        fmt = self.format_combo.currentText()
        if "xlsx" in fmt:
            ext = ".xlsx"
            filter_str = "Excel Files (*.xlsx)"
        else:
            ext = ".csv"
            filter_str = "CSV Files (*.csv)"

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存为新文件", "",
            filter_str
        )
        if not save_path:
            return

        if not save_path.endswith(ext):
            save_path += ext

        try:
            # 提取选中的列
            result_df = self.df[selected_cols]

            if ext == ".xlsx":
                result_df.to_excel(save_path, index=False)
            else:
                result_df.to_csv(save_path, index=False, encoding="utf-8-sig")

            QMessageBox.information(
                self, "成功",
                f"已保存 {len(result_df)} 行 x {len(selected_cols)} 列\n到:\n{save_path}"
            )
            self.statusBar().showMessage(f"已保存到: {save_path}", 10000)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败:\n{e}")

    def show_disclaimer(self):
        """显示免责声明窗口"""
        dialog = QDialog(self)
        dialog.setWindowTitle("免责声明")
        dialog.resize(500, 400)

        layout = QVBoxLayout(dialog)

        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(True)
        text_browser.setText("""
        <h2>免责声明</h2>
        <p>本软件（以下简称"本软件"）仅供学习和研究使用，任何商业用途均需获得授权。用户使用本软件所产生的任何直接或间接损失，开发者不承担任何法律责任。</p>
        <p>本软件及其相关文档的所有权利均归开发者所有，受著作权法保护。用户在使用本软件时，需遵守国家法律法规，因使用本软件而引起的法律责任由用户自行承担。</p>
        <p>本免责声明的最终解释权归开发者所有。</p>
        """)

        layout.addWidget(text_browser)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.exec()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
    os.system("pause")
