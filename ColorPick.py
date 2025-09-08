import sys
import os
import json
import pyautogui
import keyboard
import pyperclip  # 導入剪貼簿操作庫
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QListWidget, QListWidgetItem,
    QHBoxLayout, QPushButton, QInputDialog, QMessageBox
)
from PyQt6.QtGui import QGuiApplication, QColor
from PyQt6.QtCore import QThread, pyqtSignal, QTimer, Qt

# 常數
MAX_COLOR_RECORDS = 20
COLOR_LOG_FILENAME = "color_log.json"
DEFAULT_HOTKEY = 'ctrl+shift+c'

def get_save_path():
    """取得儲存顏色記錄檔案的完整路徑。

    根據程式是作為獨立檔案 (.py) 還是凍結的可執行檔 (.exe) 決定路徑。
    """
    if getattr(sys, 'frozen', False):  # .exe
        return os.path.join(os.path.dirname(sys.executable), COLOR_LOG_FILENAME)
    else:  # .py
        return os.path.join(os.path.abspath(os.path.dirname(__file__)), COLOR_LOG_FILENAME)

class KeyListenerThread(QThread):
    """監聽快捷鍵事件的獨立執行緒。

    當指定的快捷鍵被按下時，會發射 colorCaptured 訊號。
    """
    colorCaptured = pyqtSignal(dict)

    def __init__(self, hotkey=DEFAULT_HOTKEY):
        super().__init__()
        self.hotkey = hotkey

    def run(self):
        """執行緒的主要方法，用於監聽快捷鍵。"""
        keyboard.add_hotkey(self.hotkey, self.capture_color)
        keyboard.wait()

    def capture_color(self):
        """擷取當前滑鼠位置的顏色資訊。

        取得滑鼠座標，截取該點的螢幕像素，並提取其 RGB 和 HEX 值。
        """
        position = pyautogui.position()
        x, y = position.x, position.y
        screen = QGuiApplication.primaryScreen()
        if screen:
            screenshot = screen.grabWindow(0, x, y, 1, 1)
            image = screenshot.toImage()
            color = QColor(image.pixel(0, 0))
            hex_color = color.name().upper()
            color_data = {
                "x": x,
                "y": y,
                "r": color.red(),
                "g": color.green(),
                "b": color.blue(),
                "hex": hex_color
            }
            self.colorCaptured.emit(color_data)

class ColorItemWidget(QWidget):
    """自訂的列表項目 Widget，用於顯示單個顏色記錄的資訊和預覽。"""
    def __init__(self, index, color_data):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)

        color_preview_label = QLabel()
        color_preview_label.setFixedSize(30, 20)
        color_preview_label.setStyleSheet(f"background-color: {color_data['hex']}; border: 1px solid black;")

        info_label = QLabel(
            f"{index:>2}.: ({color_data['x']}, {color_data['y']})  "
            f"RGB({color_data['r']}, {color_data['g']}, {color_data['b']})  HEX: {color_data['hex']}"
        )
        layout.addWidget(color_preview_label)
        layout.addWidget(info_label)

class ColorLoggerApp(QWidget):
    """主要的顏色擷取應用程式視窗。"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python 抓抓工具")
        self.setFixedSize(370, 568)
        self.hotkey = DEFAULT_HOTKEY  # 預設快捷鍵
        self.color_history = self.load_data()
        self.init_ui()
        self.setStyleSheet("""
            QWidget {
                font-family: "Microsoft YaHei UI", "微軟雅黑", sans-serif;
                background-color: #2e3440;
                color: #eceff4;
            }
            QLabel {
                background-color: transparent;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #003d80;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #555;
            }
            QListWidget::item:hover {
                background-color: #434c5e;
                color: #eceff4;
            }
            QListWidget::item:selected {
                background-color: #0056b3;
                color: white;
            }
        """)
        self.start_listener()
        self.start_live_update_timer()

    def init_ui(self):
        """初始化使用者介面元素。"""
        self.main_layout = QVBoxLayout(self)

        # 即時顏色資訊顯示區域
        self.live_info_layout = QHBoxLayout()
        self.live_position_label = QLabel("X: 0 Y: 0")
        self.live_rgb_label = QLabel("RGB: 0 , 0 , 0")
        self.live_hex_label = QLabel("HEX: #000000")
        self.live_preview = QLabel()
        self.live_preview.setFixedSize(20, 20)
        self.live_preview.setStyleSheet("background-color: #000000; border: 1px solid black;")

        self.live_info_layout.addWidget(self.live_position_label)
        self.live_info_layout.addWidget(self.live_rgb_label)
        self.live_info_layout.addWidget(self.live_hex_label)
        self.live_info_layout.addWidget(self.live_preview)
        self.main_layout.addLayout(self.live_info_layout)

        # 顏色記錄列表
        self.color_list = QListWidget()
        self.color_list.itemDoubleClicked.connect(self.copy_hex_color)  # 雙擊複製 HEX
        self.main_layout.addWidget(self.color_list)

        # 按鈕區域
        self.button_layout = QHBoxLayout()
        self.clear_button = QPushButton("清空記錄")
        self.clear_button.clicked.connect(self.clear_history)
        self.change_hotkey_button = QPushButton("更改快捷鍵")
        self.change_hotkey_button.clicked.connect(self.change_hotkey)
        self.button_layout.addWidget(self.clear_button)
        self.button_layout.addWidget(self.change_hotkey_button)
        self.main_layout.addLayout(self.button_layout)

        self.update_color_list_display()

    def start_listener(self):
        """啟動快捷鍵監聽執行緒。"""
        self.listener_thread = KeyListenerThread(self.hotkey)
        self.listener_thread.colorCaptured.connect(self.add_color_record)
        self.listener_thread.start()

    def start_live_update_timer(self):
        """啟動定時器，定期更新即時顏色預覽。"""
        self.live_update_timer = QTimer()
        self.live_update_timer.timeout.connect(self.update_live_color_display)
        self.live_update_timer.start(100)

    def load_data(self):
        """從 JSON 檔案載入顏色記錄。

        如果檔案不存在或內容無效，則返回一個空列表。
        """
        file_path = get_save_path()
        if os.path.exists(file_path):
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    return json.load(file)
            except json.JSONDecodeError:
                print(f"警告: 無法解析 {file_path}，將使用空記錄。")
                return []
        return []

    def save_data(self):
        """將顏色記錄儲存到 JSON 檔案。"""
        file_path = get_save_path()
        with open(file_path, "w", encoding="utf-8") as file:
            json.dump(self.color_history, file, indent=4)

    def add_color_record(self, color_record):
        """新增一個顏色記錄到歷史記錄中。

        如果記錄數量超過上限，則移除最舊的記錄。更新顯示並儲存資料。
        """
        if len(self.color_history) >= MAX_COLOR_RECORDS:
            self.color_history.pop(0)
        self.color_history.append(color_record)
        self.save_data()
        self.update_color_list_display()

    def update_color_list_display(self):
        """更新顏色記錄列表的顯示。"""
        self.color_list.clear()
        for index, record in enumerate(self.color_history, start=1):
            item_widget = ColorItemWidget(index, record)
            list_item = QListWidgetItem()
            list_item.setSizeHint(item_widget.sizeHint())
            self.color_list.addItem(list_item)
            self.color_list.setItemWidget(list_item, item_widget)

    def update_live_color_display(self):
        """更新即時顏色預覽標籤的顯示。"""
        position = pyautogui.position()
        x, y = position.x, position.y
        screen = QGuiApplication.primaryScreen()
        if screen:
            screenshot = screen.grabWindow(0, x, y, 1, 1)
            image = screenshot.toImage()
            color = QColor(image.pixel(0, 0))
            red, green, blue = color.red(), color.green(), color.blue()
            hex_color = color.name().upper()
            self.live_position_label.setText(f"X: {x:<8} Y: {y}")
            self.live_rgb_label.setText(f"RGB: {red:<3}, {green:<3}, {blue:<3}")
            self.live_hex_label.setText(f"HEX: {hex_color}")
            self.live_preview.setStyleSheet(f"background-color: {hex_color}; border: 1px solid black;")

    def copy_hex_color(self, list_item):
        """複製選中列表項目的 HEX 顏色碼到剪貼簿，並顯示一個提示訊息。"""
        index = self.color_list.row(list_item)
        if 0 <= index < len(self.color_history):
            hex_color = self.color_history[index]['hex']
            pyperclip.copy(hex_color)
            QMessageBox.information(self, "複製成功", f"已複製 HEX 顏色碼: {hex_color}", QMessageBox.StandardButton.Ok)

    def clear_history(self):
        """清空所有的顏色記錄，更新顯示並儲存。"""
        self.color_history = []
        self.save_data()
        self.update_color_list_display()

    def change_hotkey(self):
        """更改顏色擷取的快捷鍵。

        彈出一個輸入對話框，讓使用者輸入新的快捷鍵。
        更新快捷鍵監聽器，並處理可能的錯誤。
        """
        new_hotkey, ok = QInputDialog.getText(self, "更改快捷鍵",
                                               "請輸入新的快捷鍵 (例如: ctrl+alt+c):",
                                               text=self.hotkey)
        if ok and new_hotkey:
            self.hotkey = new_hotkey
            self.listener_thread.hotkey = new_hotkey
            try:
                keyboard.unhook_all_hotkeys()
                self.start_listener()
                QMessageBox.information(self, "快捷鍵已更改", f"快捷鍵已更改為: {self.hotkey}", QMessageBox.StandardButton.Ok)
            except Exception as e:
                error_message = f"更改快捷鍵時發生錯誤: {e}\n請檢查輸入的快捷鍵格式是否正確。"
                QMessageBox.critical(self, "錯誤", error_message, QMessageBox.StandardButton.Ok)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ColorLoggerApp()
    window.show()
    sys.exit(app.exec())