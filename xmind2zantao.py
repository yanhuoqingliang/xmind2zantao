import sys
import os
import shutil
import sqlite3
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QTableWidget, \
    QTableWidgetItem, QFileDialog, QLabel, QHBoxLayout, QHeaderView, QSizePolicy, QMessageBox
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QFont, QColor, QCursor, QIcon
from datetime import datetime

from xmind2testcase.utils import get_xmind_testcase_list, get_xmind_testsuites
from xmind2testcase.zentao import xmind_to_zentao_csv_file


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller 打包后的临时路径
    except Exception:
        base_path = os.path.abspath(".")  # 开发环境中的当前路径
    return os.path.join(base_path, relative_path)


class PreviewWindow(QMainWindow):
    def __init__(self, xmind_file, testcases, x, y):
        super().__init__()
        self.xmind_file = xmind_file
        self.test_cases = testcases  # 将 JSON 数据保存为类的成员变量
        testsuites = get_xmind_testsuites(xmind_file)
        self.suite_count = 0
        for suite in testsuites:
            self.suite_count += len(suite.sub_suites)
        self.setWindowTitle(f"{os.path.basename(xmind_file)} - Preview")
        self.setGeometry(x, y, 1000, 800)
        self.setWindowIcon(QIcon(resource_path("logo.png")))
        self.initUI()

    def initUI(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.title_label = QLabel(f"Preview: {os.path.basename(self.xmind_file)}", self)
        self.title_label.setFont(QFont("Arial", 20, QFont.Bold))
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: #2196F3;")
        main_layout.addWidget(self.title_label)

        button_layout = QHBoxLayout()
        self.testsuites_label = QLabel(f'TestSuites: {self.suite_count}', self)
        self.testsuites_label.setFont(QFont("Arial", 15, QFont.Bold))
        self.testsuites_label.setAlignment(Qt.AlignCenter)
        self.testsuites_label.setStyleSheet("color: #2196F3;")

        self.testcases_label = QLabel(f'TestCases: {len(self.test_cases)}', self)
        self.testcases_label.setFont(QFont("Arial", 15, QFont.Bold))
        self.testcases_label.setAlignment(Qt.AlignCenter)
        self.testcases_label.setStyleSheet("color: #2196F3;")

        self.CSV_button = self.create_button("下载CSV文件", "#FF5722", "#F44336", self.export_csv)
        self.back_button = self.create_button("返回主页", "#FF5722", "#F44336", self.goBack)
        button_layout.addWidget(self.testsuites_label)
        button_layout.addWidget(self.testcases_label)
        button_layout.addWidget(self.CSV_button)
        button_layout.addWidget(self.back_button)
        main_layout.addLayout(button_layout)

        # 创建表格
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(4)  # 根据需要列数调整
        self.table_widget.setHorizontalHeaderLabels(["Suite", "Title", "Summary", "Steps"])

        # 设置表格自适应宽度
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # 让列自适应宽度
        self.table_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 表格扩展填满空间

        # 填充表格内容
        self.populate_table()

        # 添加表格到布局
        main_layout.addWidget(self.table_widget)

        # 初始化窗口大小
        self.show()

    def populate_table(self):
        # 清空表格内容
        self.table_widget.setRowCount(0)

        # 遍历每个测试用例
        for test_case in self.test_cases:
            row_position = self.table_widget.rowCount()
            self.table_widget.insertRow(row_position)

            # 填充套件名、测试名、测试概述
            self.table_widget.setItem(row_position, 0, QTableWidgetItem(test_case['suite']))  # 套件名称
            self.table_widget.setItem(row_position, 1, QTableWidgetItem(test_case['name']))  # 测试名称

            # 初始化步骤文本
            steps_text = ""
            for step in test_case['steps']:
                steps_text += f"Step {step['step_number']}: {step['actions']}\n"
                if step['expectedresults']:
                    steps_text += f"Expected Results: {step['expectedresults']}\n"

            step_item = QTableWidgetItem(steps_text.strip())
            step_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            step_item.setToolTip(steps_text)  # 设置提示框显示完整文本

            self.table_widget.setItem(row_position, 3, step_item)

            # 动态调整行高
            text_length = len(steps_text)
            row_height = 50 + (text_length // 100) * 60  # 依据文本长度动态增加行高
            self.table_widget.setRowHeight(row_position, row_height)

            # 在 "Summary" 列添加按钮
            self.add_summary_buttons(row_position, test_case)

        # 设置列宽，确保文本能够显示完整
        self.table_widget.setColumnWidth(0, 150)
        self.table_widget.setColumnWidth(1, 200)
        self.table_widget.setColumnWidth(2, 250)
        self.table_widget.setColumnWidth(3, 400)

        # 让表格宽度自适应
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def set_table_styles(self):
        """美化表格样式"""
        self.table_widget.setStyleSheet("""
            QTableWidget {
                background-color: #f9f9f9;
                gridline-color: #d3d3d3;
                border-radius: 5px;
                border: 1px solid #e0e0e0;
                font-size: 12px;
                color: #333333;
                padding: 10px;
            }
            QTableWidget::item {
                border: 1px solid #f0f0f0;
                padding: 10px;
                font-size: 12px;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 6px;
                font-size: 14px;
            }
            QTableWidget::item:hover {
                background-color: #f1f1f1;
                border: 1px solid #b0b0b0;
            }
            QTableWidget::item:selected {
                background-color: #cce7ff;
                color: #000000;
                border: 1px solid #4CAF50;
            }
        """)

    def add_summary_buttons(self, row_position, test_case):
        """ 在 'Summary' 列中添加三个按钮：Priority、Preconditions、Summary """
        summary_layout = QHBoxLayout()

        # Priority 按钮
        priority_button = self.create_tip_button("Priority", str(test_case['importance']), "#8BC34A")  # 设置绿色背景
        summary_layout.addWidget(priority_button)

        # Preconditions 按钮
        preconditions_button = self.create_tip_button("Preconditions", test_case['preconditions'], "#2196F3")  # 设置蓝色背景
        summary_layout.addWidget(preconditions_button)

        # Summary 按钮
        summary_button = self.create_tip_button("Summary", test_case['summary'], "#FF9800")  # 设置橙色背景
        summary_layout.addWidget(summary_button)

        # 将按钮布局设置为表格单元格内容
        summary_widget = QWidget(self)
        summary_widget.setLayout(summary_layout)
        self.table_widget.setCellWidget(row_position, 2, summary_widget)  # 设置到 Summary 列

    def create_tip_button(self, text, tooltip, background_color):
        """创建带有tooltip的按钮，并设置不同颜色"""
        button = QPushButton(text, self)
        button.setToolTip(tooltip)
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {background_color};
                color: white;
                padding: 5px;
                border-radius: 5px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.darken_color(background_color)};
            }}
        """)
        return button

    def darken_color(self, color):
        """暗化颜色（使按钮在 hover 时变深）"""
        # 暗化颜色：简单地将颜色的RGB值减少
        color = QColor(color)
        color.setRed(max(color.red() - 20, 0))
        color.setGreen(max(color.green() - 20, 0))
        color.setBlue(max(color.blue() - 20, 0))
        return color.name()

    def goBack(self):
        self.close()
        self.main_window = MainWindow()
        self.main_window.show()
        self.main_window.move(self.main_window.saved_geometry.topLeft())

    def export_csv(self):
        zentao_csv_file = xmind_to_zentao_csv_file(self.xmind_file)
        print('Convert XMind file to zentao csv file successfully: %s' % zentao_csv_file)

        file_name_cvs = os.path.basename(zentao_csv_file)
        print(f"导出 {file_name_cvs} 为 CSV")

        # Check if the file exists
        if not os.path.exists(zentao_csv_file):
            self.show_message("错误", f"文件 {file_name_cvs} 不存在")
            return

        # Use QFileDialog to ask the user where to save the downloaded file
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 CSV 文件", file_name_cvs, "XMind Files (*.csv)")

        # If a path is selected by the user
        if save_path:
            try:
                # Copy the file to the selected location
                shutil.copy(zentao_csv_file, save_path)
                print(f"文件 {file_name_cvs} 已下载到 {save_path}")
                self.show_message("成功", f"文件已成功下载到 {save_path}")
            except Exception as e:
                print(f"下载文件失败: {str(e)}")
                self.show_message("错误", f"下载文件失败: {str(e)}")

    def create_button(self, text, color, hover_color, func):
        button = QPushButton(text, self)
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                font-size: 16px;
                padding: 10px;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
        """)
        button.clicked.connect(func)
        return button

    def show_message(self, title, message):
        QMessageBox.information(self, title, message)


class Database:
    def __init__(self, db_name="records.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        sql = '''
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            create_on TEXT NOT NULL,
            note TEXT,
            is_deleted INTEGER DEFAULT 0
        );
        '''
        self.cursor.execute(sql)
        self.conn.commit()

    def insert_record(self, name, create_on, note=None):
        self.cursor.execute('INSERT INTO records (name, create_on, note) VALUES (?, ?, ?)', (name, create_on, note))
        self.conn.commit()

    def get_all_records(self):
        self.cursor.execute("SELECT * FROM records WHERE is_deleted = 0 ORDER BY create_on DESC")
        return self.cursor.fetchall()

    def delete_record_by_name(self, name):
        self.cursor.execute("DELETE FROM records WHERE name = ?", (name,))
        self.conn.commit()

    def close(self):
        self.conn.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XMind 转换器")
        self.setGeometry(100, 100, 1000, 800)
        self.setWindowIcon(QIcon(resource_path("logo.png")))
        self.db = Database()
        self.saved_geometry = self.geometry()
        self.upload_folder = "upload"
        os.makedirs(self.upload_folder, exist_ok=True)
        self.initUI()

    def initUI(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.title_label = QLabel("XMIND TO TESTCASE", self)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setFont(QFont("Arial", 30, QFont.Bold))
        self.title_label.setStyleSheet("color: #4CAF50;")
        main_layout.addWidget(self.title_label)

        # 创建标签 "请选择文件"
        self.select_file_label = QLabel("--> 点击这里选择您的XMind文件 <--")
        self.select_file_label.setFont(QFont("Arial", 20))
        self.select_file_label.setStyleSheet("color: blue;")
        self.select_file_label.setCursor(QCursor(Qt.PointingHandCursor))
        self.select_file_label.mousePressEvent = self.on_select_file_click
        self.select_file_label.setAlignment(Qt.AlignCenter)
        # 为标签安装事件过滤器
        self.select_file_label.installEventFilter(self)
        main_layout.addWidget(self.select_file_label)

        self.convert_button = self.create_button("开始转换", "#2196F3", "#1976D2", self.startConversion)
        main_layout.addWidget(self.convert_button)

        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(["NAME", "TIME", "ACTIONS"])
        self.style_table(self.table_widget)
        main_layout.addWidget(self.table_widget)

        self.selected_file_path = None
        self.load_data()

    def create_button(self, text, color, hover_color, func):
        button = QPushButton(text, self)
        button.setMinimumHeight(30)
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
        """)
        button.clicked.connect(func)
        return button

    def style_table(self, table_widget):
        table_widget.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                background-color: #fafafa;
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 10px;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        table_widget.setColumnWidth(0, 200)
        table_widget.setColumnWidth(1, 150)
        table_widget.setColumnWidth(2, 350)
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def load_data(self):
        records = self.db.get_all_records()
        self.table_widget.setRowCount(0)  # 清空表格
        for record in records:
            row_position = self.table_widget.rowCount()
            self.table_widget.insertRow(row_position)
            self.table_widget.setItem(row_position, 0, QTableWidgetItem(record[1]))
            self.table_widget.setItem(row_position, 1, QTableWidgetItem(record[2]))
            self.table_widget.setCellWidget(row_position, 2, self.create_action_buttons(record[1]))

            # 设置行高，以确保所有按钮都能显示
            self.table_widget.setRowHeight(row_position, 50)  # 根据需要调整行高

    def create_action_buttons(self, file_name):
        actions_layout = QHBoxLayout()
        actions_layout.setAlignment(Qt.AlignCenter)  # 居中对齐

        actions_layout.addWidget(self.create_button("XMIND", "#FF9800", "#FB8C00", lambda: self.open_xmind(file_name)))
        actions_layout.addWidget(self.create_button("CSV", "#03A9F4", "#0288D1", lambda: self.export_csv(file_name)))
        actions_layout.addWidget(self.create_button("PREVIEW", "#8BC34A", "#7CB342", lambda: self.preview_xmind(file_name)))
        actions_layout.addWidget(self.create_button("DELETE", "#f44336", "#d32f2f", lambda: self.delete_record(file_name)))

        actions_widget = QWidget()
        actions_widget.setLayout(actions_layout)
        actions_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # 固定高度
        actions_widget.setMinimumHeight(40)  # 确保高度足够
        return actions_widget

    def open_xmind(self, file_name):
        # Print the action to the console
        print(f"打开 XMind 文件: {file_name}")

        # Get the full path of the XMind file in the upload folder
        file_path = os.path.join(self.upload_folder, file_name)

        # Check if the file exists
        if not os.path.exists(file_path):
            self.show_message("错误", f"文件 {file_name} 不存在")
            return

        # Use QFileDialog to ask the user where to save the downloaded file
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 XMind 文件", file_name, "XMind Files (*.xmind)")

        # If a path is selected by the user
        if save_path:
            try:
                # Copy the file to the selected location
                shutil.copy(file_path, save_path)
                print(f"文件 {file_name} 已下载到 {save_path}")
                self.show_message("成功", f"文件已成功下载到 {save_path}")
            except Exception as e:
                print(f"下载文件失败: {str(e)}")
                self.show_message("错误", f"下载文件失败: {str(e)}")

    def export_csv(self, file_name):
        # Get the full path of the XMind file in the upload folder
        file_path = os.path.join(self.upload_folder, file_name)

        zentao_csv_file = xmind_to_zentao_csv_file(file_path)
        print('Convert XMind file to zentao csv file successfully: %s' % zentao_csv_file)

        file_name_cvs = os.path.basename(zentao_csv_file)
        print(f"导出 {file_name_cvs} 为 CSV")

        # Check if the file exists
        if not os.path.exists(zentao_csv_file):
            self.show_message("错误", f"文件 {file_name_cvs} 不存在")
            return

        # Use QFileDialog to ask the user where to save the downloaded file
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 CSV 文件", file_name_cvs, "XMind Files (*.csv)")

        # If a path is selected by the user
        if save_path:
            try:
                # Copy the file to the selected location
                shutil.copy(zentao_csv_file, save_path)
                print(f"文件 {file_name_cvs} 已下载到 {save_path}")
                self.show_message("成功", f"文件已成功下载到 {save_path}")
            except Exception as e:
                print(f"下载文件失败: {str(e)}")
                self.show_message("错误", f"下载文件失败: {str(e)}")

    def preview_xmind(self, file_name):
        file_path = os.path.join(self.upload_folder, file_name)
        if not os.path.exists(file_path):
            self.show_message("错误", f"文件 {file_name} 不存在")
            return

        testcases = get_xmind_testcase_list(file_path)

        current_geometry = self.geometry()
        x, y = current_geometry.x() + 20, current_geometry.y() + 20  # 小幅度偏移

        self.preview_window = PreviewWindow(file_path, testcases, x, y)
        self.preview_window.show()
        self.close()

    def delete_record(self, file_name):
        self.db.delete_record_by_name(file_name)
        file_path = os.path.join(self.upload_folder, file_name)
        file_path_csv = os.path.join(self.upload_folder, os.path.splitext(os.path.basename(file_path))[0] + '.csv')
        print(file_path)
        print(file_path_csv)
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"文件 {file_name} 已被删除")
        if os.path.exists(file_path_csv):
            os.remove(file_path_csv)
            print(f"文件 {file_path_csv} 已被删除")
        self.load_data()  # Reload data to update the table

    def on_select_file_click(self, event):
        """处理文件选择"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "XMind Files (*.xmind)", options=options)

        if file_path:
            # 显示选择的文件路径
            file_name = os.path.basename(file_path)
            self.select_file_label.setText(file_name)  # 覆盖 "请选择文件" 文本
            self.selected_file_path = file_path

    def eventFilter(self, source, event):
        """过滤鼠标悬停和离开事件"""
        if source == self.select_file_label:
            if event.type() == QEvent.Enter:
                # 鼠标进入事件，改变文字颜色和加下划线
                self.select_file_label.setStyleSheet("color: green; text-decoration: underline;")
            elif event.type() == QEvent.Leave:
                # 鼠标离开事件，恢复默认样式
                self.select_file_label.setStyleSheet("color: blue;")
        return super().eventFilter(source, event)

    def startConversion(self):
        if self.selected_file_path:
            filename = os.path.basename(self.selected_file_path)
            destination = self.get_unique_file_path(filename)  # 生成唯一文件名

            try:
                shutil.copy(self.selected_file_path, destination)  # 复制文件到目标位置
                create_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # 使用带时间戳的文件名插入记录
                self.db.insert_record(name=os.path.basename(destination), create_on=create_on, note="上传的XMind文件")

                testcases = get_xmind_testcase_list(destination)

                current_geometry = self.geometry()
                x, y = current_geometry.x() + 20, current_geometry.y() + 20  # 小幅度偏移

                self.preview_window = PreviewWindow(destination, testcases, x, y)
                self.preview_window.show()
                self.close()
            except Exception as e:
                self.show_message("上传错误", str(e))
        else:
            self.show_message('提示', '请选择一个文件！')

    def get_unique_file_path(self, filename):
        """确保文件名唯一，如果已存在则添加时间戳"""
        base_filename, file_extension = os.path.splitext(filename)
        new_filename = filename
        while os.path.exists(os.path.join(self.upload_folder, new_filename)):
            timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
            new_filename = f"{base_filename}{timestamp}{file_extension}"
        return os.path.join(self.upload_folder, new_filename)

    def show_message(self, title, message):
        QMessageBox.information(self, title, message)

    def showEvent(self, event):
        self.saved_geometry = self.geometry()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())