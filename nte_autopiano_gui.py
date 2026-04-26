import sys
import os
import time
import ctypes
import mido
import psutil
import win32gui
import win32process
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QListWidget,
                             QPushButton, QCheckBox, QDoubleSpinBox, QSpinBox,
                             QSlider, QStyle, QListWidgetItem, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QIcon

# ==========================================
# DirectInput 底层按键映射 (防反作弊)
# ==========================================
SendInput = ctypes.windll.user32.SendInput

PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort), ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong), ("wParamL", ctypes.c_short), ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long), ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput), ("mi", MouseInput), ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", Input_I)]


# 硬件扫描码映射
SCAN_CODES = {
    'z': 0x2C, 'x': 0x2D, 'c': 0x2E, 'v': 0x2F, 'b': 0x30, 'n': 0x31, 'm': 0x32,
    'a': 0x1E, 's': 0x1F, 'd': 0x20, 'f': 0x21, 'g': 0x22, 'h': 0x23, 'j': 0x24,
    'q': 0x10, 'w': 0x11, 'e': 0x12, 'r': 0x13, 't': 0x14, 'y': 0x15, 'u': 0x16,
    'shift': 0x2A, 'ctrl': 0x1D
}


def press_key(key_char):
    hexKeyCode = SCAN_CODES.get(key_char)
    if not hexKeyCode: return
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def release_key(key_char):
    hexKeyCode = SCAN_CODES.get(key_char)
    if not hexKeyCode: return
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


# ==========================================
# 异环键位映射逻辑
# ==========================================
def generate_key_map():
    mapping = {}
    octaves = {
        3: ['z', 'x', 'c', 'v', 'b', 'n', 'm'],
        4: ['a', 's', 'd', 'f', 'g', 'h', 'j'],
        5: ['q', 'w', 'e', 'r', 't', 'y', 'u']
    }
    for oct_idx, (oct_num, keys) in enumerate(octaves.items()):
        base_midi = 48 + (oct_idx * 12)
        mapping[base_midi + 0] = (keys[0], None)
        mapping[base_midi + 1] = (keys[0], 'shift')
        mapping[base_midi + 2] = (keys[1], None)
        mapping[base_midi + 3] = (keys[2], 'ctrl')
        mapping[base_midi + 4] = (keys[2], None)
        mapping[base_midi + 5] = (keys[3], None)
        mapping[base_midi + 6] = (keys[3], 'shift')
        mapping[base_midi + 7] = (keys[4], None)
        mapping[base_midi + 8] = (keys[4], 'shift')
        mapping[base_midi + 9] = (keys[5], None)
        mapping[base_midi + 10] = (keys[6], 'ctrl')
        mapping[base_midi + 11] = (keys[6], None)
    return mapping


KEY_MAP = generate_key_map()


# ==========================================
# 后台播放线程
# ==========================================
class MidiPlayerThread(QThread):
    progress_sig = pyqtSignal(int, str)  # 当前进度(秒), 时间文本
    status_sig = pyqtSignal(str)
    finished_sig = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.is_playing = False
        self.is_paused = False
        self.seek_time = -1

        # 配置参数
        self.speed = 1.0
        self.pitch_shift = 0
        self.target_process = ""
        self.pause_on_leave = True
        self.force_range = False  # 新增强制压缩音域

        self.total_duration = 0
        self.events = []  # [(time, note, type), ...]

    def load_midi(self, file_path):
        self.file_path = file_path
        try:
            mid = mido.MidiFile(file_path)
            self.events.clear()
            current_time = 0
            for msg in mid:
                current_time += msg.time
                if msg.type in ['note_on', 'note_off']:
                    self.events.append((current_time, msg.note, msg.velocity if msg.type == 'note_on' else 0))
            self.total_duration = current_time
            return True
        except:
            return False

    def get_active_process(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return psutil.Process(pid).name().lower()
        except:
            return ""

    def hit_key(self, note):
        shifted_note = note + self.pitch_shift

        # 强制将音符压缩到异环支持的音域内(48 ~ 83)
        if self.force_range:
            while shifted_note < 48:
                shifted_note += 12
            while shifted_note > 83:
                shifted_note -= 12

        if shifted_note in KEY_MAP:
            key, mod = KEY_MAP[shifted_note]

            if mod: press_key(mod)
            press_key(key)

            release_key(key)
            if mod: release_key(mod)

    def run(self):
        if not self.events: return
        self.is_playing = True

        start_time = time.time()
        current_event_idx = 0
        virtual_time = 0.0
        last_status = ""

        while self.is_playing and current_event_idx < len(self.events):
            if self.seek_time >= 0:
                virtual_time = self.seek_time
                start_time = time.time() - (virtual_time / self.speed)
                # 寻找进度对应的事件索引
                current_event_idx = next((i for i, e in enumerate(self.events) if e[0] >= virtual_time),
                                         len(self.events))
                self.seek_time = -1

            if self.is_paused:
                if last_status != "paused":
                    self.status_sig.emit("已暂停")
                    last_status = "paused"
                start_time = time.time() - (virtual_time / self.speed)
                time.sleep(0.1)
                continue

            # 离开窗口暂停逻辑
            if self.pause_on_leave and self.target_process:
                active_proc = self.get_active_process()
                if active_proc != self.target_process.lower():
                    if last_status != "waiting":
                        self.status_sig.emit("等待进入游戏窗口...")
                        last_status = "waiting"
                    start_time = time.time() - (virtual_time / self.speed)
                    time.sleep(0.5)
                    continue

            if last_status in ["waiting", "paused", ""]:
                self.status_sig.emit("正在播放...")
                last_status = "playing"

            virtual_time = (time.time() - start_time) * self.speed

            # 触发到时间的音符
            while current_event_idx < len(self.events) and self.events[current_event_idx][0] <= virtual_time:
                ev_time, note, vel = self.events[current_event_idx]
                if vel > 0:  # Note On
                    self.hit_key(note)
                current_event_idx += 1

            # 更新UI进度 (限制频率)
            if int(virtual_time * 10) % 5 == 0:
                cur_m, cur_s = divmod(int(virtual_time), 60)
                tot_m, tot_s = divmod(int(self.total_duration), 60)
                time_str = f"{cur_m}:{cur_s:02d} / {tot_m}:{tot_s:02d}"
                self.progress_sig.emit(int(virtual_time), time_str)

            time.sleep(0.005)

        self.is_playing = False
        self.finished_sig.emit()


# ==========================================
# 主界面 UI
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NTE AutoPiano UI - 异环自动弹琴")
        self.resize(850, 550)

        self.light_style = """
            QMainWindow { background-color: #f2f2f5; }
            QLabel, QCheckBox { font-family: 'Microsoft YaHei'; color: #333; }
            QListWidget { background: white; color: #333; border-radius: 8px; border: 1px solid #ddd; padding: 5px; outline: 0; }
            QListWidget::item { padding: 8px; border-bottom: 1px solid #eee; }
            QListWidget::item:selected { background: #e6ccff; color: #6600cc; border-radius: 4px; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { border: 1px solid #ccc; border-radius: 4px; padding: 5px; color: #333; background: white;}
            QComboBox::drop-down { border: 0px; }
            QPushButton { background-color: #b066ff; color: white; border-radius: 15px; padding: 8px; font-weight: bold; }
            QPushButton:hover { background-color: #9933ff; }
            QPushButton:disabled { background-color: #cccccc; }
            QSlider::groove:horizontal { border-radius: 3px; height: 6px; background: #ddd; }
            QSlider::handle:horizontal { background: #b066ff; width: 14px; margin: -4px 0; border-radius: 7px; }
        """

        self.dark_style = """
            QMainWindow { background-color: #1e1e24; }
            QLabel, QCheckBox { font-family: 'Microsoft YaHei'; color: #eee; }
            QListWidget { background: #2b2b36; color: #eee; border-radius: 8px; border: 1px solid #444; padding: 5px; outline: 0; }
            QListWidget::item { padding: 8px; border-bottom: 1px solid #3c3c4a; }
            QListWidget::item:selected { background: #6600cc; color: #fff; border-radius: 4px; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { border: 1px solid #555; border-radius: 4px; padding: 5px; color: #eee; background: #2b2b36;}
            QComboBox::drop-down { border: 0px; }
            QPushButton { background-color: #8a2be2; color: white; border-radius: 15px; padding: 8px; font-weight: bold; }
            QPushButton:hover { background-color: #9932cc; }
            QPushButton:disabled { background-color: #555555; color: #888; }
            QSlider::groove:horizontal { border-radius: 3px; height: 6px; background: #444; }
            QSlider::handle:horizontal { background: #8a2be2; width: 14px; margin: -4px 0; border-radius: 7px; }
        """

        self.setStyleSheet(self.dark_style)

        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        self.midi_dir = os.path.join(base_dir, "midi_file")
        if not os.path.exists(self.midi_dir):
            os.makedirs(self.midi_dir)

        self.player = MidiPlayerThread()
        self.player.progress_sig.connect(self.update_progress)
        self.player.status_sig.connect(self.update_status)
        self.player.finished_sig.connect(self.on_playback_finished)

        self.init_ui()
        self.refresh_list()
        self.refresh_windows()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)

        # 上半部分：列表 + 设置
        top_layout = QHBoxLayout()

        # --- 左侧：播放列表 ---
        left_layout = QVBoxLayout()
        header_layout = QHBoxLayout()
        title_lbl = QLabel("播放列表 (双击播放)")
        title_lbl.setFont(QFont("Microsoft YaHei", 14, QFont.Weight.Bold))
        refresh_btn = QPushButton("点我刷新!!")
        refresh_btn.setFixedSize(90, 30)
        refresh_btn.clicked.connect(self.refresh_list)

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(refresh_btn)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索 MIDI 文件...")
        self.search_input.textChanged.connect(self.filter_list)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.play_selected)

        left_layout.addLayout(header_layout)
        left_layout.addWidget(self.search_input)
        left_layout.addWidget(self.list_widget)

        # --- 右侧：设置 ---
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(20, 0, 0, 0)

        settings_title = QLabel("设置")
        settings_title.setFont(QFont("Microsoft YaHei", 14, QFont.Weight.Bold))
        right_layout.addWidget(settings_title)

        # 自动跳转窗口 (仅作展示或可后续接入win32gui.SetForegroundWindow)
        self.chk_auto_jump = QCheckBox("自动跳转游戏窗口")
        self.chk_auto_jump.setChecked(True)
        # 离开暂停
        self.chk_pause_leave = QCheckBox("离开窗口暂停播放")
        self.chk_pause_leave.setChecked(True)
        self.chk_pause_leave.stateChanged.connect(self.update_settings)

        self.chk_force_range = QCheckBox("强制压缩至可弹奏音域")
        self.chk_force_range.setChecked(False)
        self.chk_force_range.stateChanged.connect(self.update_settings)

        self.chk_dark_mode = QCheckBox("暗色模式")
        self.chk_dark_mode.setChecked(True)
        self.chk_dark_mode.stateChanged.connect(self.toggle_theme)

        proc_layout = QHBoxLayout()
        proc_layout.addWidget(QLabel("指定操作进程"))
        self.proc_combo = QComboBox()
        self.proc_combo.setFixedWidth(160)
        self.proc_combo.currentIndexChanged.connect(self.update_settings)

        self.btn_refresh_proc = QPushButton("↻")
        self.btn_refresh_proc.setFixedSize(30, 30)
        self.btn_refresh_proc.clicked.connect(self.refresh_windows)

        proc_layout.addWidget(self.proc_combo)
        proc_layout.addWidget(self.btn_refresh_proc)
        proc_layout.addStretch()

        speed_layout = QHBoxLayout()
        speed_layout.addWidget(QLabel("速度"))
        self.speed_spin = QDoubleSpinBox()
        self.speed_spin.setRange(0.1, 5.0)
        self.speed_spin.setValue(1.0)
        self.speed_spin.setSingleStep(0.1)
        self.speed_spin.valueChanged.connect(self.update_settings)
        speed_layout.addWidget(self.speed_spin)
        speed_layout.addStretch()

        pitch_layout = QHBoxLayout()
        pitch_layout.addWidget(QLabel("升降调范围"))
        self.pitch_spin = QSpinBox()
        self.pitch_spin.setRange(-48, 48)
        self.pitch_spin.setValue(0)
        self.pitch_spin.valueChanged.connect(self.update_settings)
        pitch_layout.addWidget(self.pitch_spin)
        pitch_layout.addStretch()

        right_layout.addWidget(self.chk_auto_jump)
        right_layout.addWidget(self.chk_pause_leave)
        right_layout.addWidget(self.chk_force_range)
        right_layout.addWidget(self.chk_dark_mode)
        right_layout.addLayout(proc_layout)
        right_layout.addLayout(speed_layout)
        right_layout.addLayout(pitch_layout)

        right_layout.addSpacing(30)
        info_title = QLabel("文件信息")
        info_title.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.info_lbl = QLabel("关注b站视界Seekai喵\n来LingFlame.cn看看喵\n安全模式已启用(DirectInput)")
        self.info_lbl.setStyleSheet("color: #666;")
        right_layout.addWidget(info_title)
        right_layout.addWidget(self.info_lbl)
        right_layout.addStretch()

        top_layout.addLayout(left_layout, 2)
        top_layout.addLayout(right_layout, 1)

        # --- 底部：控制栏 ---
        bottom_layout = QHBoxLayout()

        self.btn_prev = QPushButton("上一首")
        self.btn_prev.setFixedSize(60, 40)
        self.btn_play = QPushButton("播放")
        self.btn_play.setFixedSize(60, 50)
        self.btn_play.clicked.connect(self.toggle_play)
        self.btn_next = QPushButton("下一首")
        self.btn_next.setFixedSize(60, 40)

        self.status_lbl = QLabel("暂未播放")
        self.status_lbl.setFixedWidth(150)

        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setRange(0, 100)
        self.slider.sliderReleased.connect(self.seek_position)

        self.time_lbl = QLabel("0:00 / 0:00")

        bottom_layout.addWidget(self.btn_prev)
        bottom_layout.addWidget(self.btn_play)
        bottom_layout.addWidget(self.btn_next)
        bottom_layout.addSpacing(20)
        bottom_layout.addWidget(self.status_lbl)
        bottom_layout.addWidget(self.slider)
        bottom_layout.addWidget(self.time_lbl)

        main_layout.addLayout(top_layout)
        main_layout.addSpacing(10)
        main_layout.addLayout(bottom_layout)

        self.setCentralWidget(central_widget)
        self.update_settings()

    def refresh_list(self):
        self.list_widget.clear()
        if not os.path.exists(self.midi_dir): return
        for f in os.listdir(self.midi_dir):
            if f.lower().endswith(('.mid', '.midi')):
                self.list_widget.addItem(f)

    def filter_list(self, text):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def refresh_windows(self):
        current_data = self.proc_combo.currentData()
        self.proc_combo.blockSignals(True)
        self.proc_combo.clear()

        windows = []

        def enum_win_cb(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    proc_name = psutil.Process(pid).name()
                    title = win32gui.GetWindowText(hwnd)
                    if title and proc_name.lower() not in ['explorer.exe', 'applicationframehost.exe']:
                        results.append((proc_name, title))
                except:
                    pass

        win32gui.EnumWindows(enum_win_cb, windows)

        seen = set()
        self.proc_combo.addItem("异环 (NTE.exe)", "nte.exe")
        seen.add("nte.exe")

        for proc_name, title in windows:
            if proc_name.lower() not in seen:
                seen.add(proc_name.lower())
                display_title = title if len(title) <= 12 else title[:11] + "..."
                self.proc_combo.addItem(f"{display_title} ({proc_name})", proc_name)

        if current_data:
            idx = self.proc_combo.findData(current_data)
            if idx >= 0:
                self.proc_combo.setCurrentIndex(idx)

        self.proc_combo.blockSignals(False)
        self.update_settings()

    def toggle_theme(self):
        if self.chk_dark_mode.isChecked():
            self.setStyleSheet(self.dark_style)
        else:
            self.setStyleSheet(self.light_style)

    def update_settings(self):
        self.player.speed = self.speed_spin.value()
        self.player.pitch_shift = self.pitch_spin.value()
        target = self.proc_combo.currentData()
        self.player.target_process = target.strip() if target else ""
        self.player.pause_on_leave = self.chk_pause_leave.isChecked()
        self.player.force_range = self.chk_force_range.isChecked()

    def jump_to_game(self):
        if not self.chk_auto_jump.isChecked(): return
        target = self.proc_combo.currentData()
        if not target: return
        target = target.strip().lower()

        def enum_win_cb(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    if psutil.Process(pid).name().lower() == target:
                        results.append(hwnd)
                except:
                    pass

        hwnds = []
        win32gui.EnumWindows(enum_win_cb, hwnds)
        if hwnds:
            try:
                win32gui.SetForegroundWindow(hwnds[0])
            except:
                pass

    def play_selected(self, item):
        file_path = os.path.join(self.midi_dir, item.text())
        if self.player.is_playing:
            self.player.is_playing = False
            self.player.wait()

        if self.player.load_midi(file_path):
            self.slider.setRange(0, int(self.player.total_duration))

            tot_m, tot_s = divmod(int(self.player.total_duration), 60)
            self.time_lbl.setText(f"0:00 / {tot_m}:{tot_s:02d}")

            self.btn_play.setText("暂停")
            self.jump_to_game()
            self.player.start()
        else:
            self.status_lbl.setText("文件加载失败！")

    def update_status(self, text):
        self.status_lbl.setText(text)

    def toggle_play(self):
        if self.player.is_playing:
            self.player.is_paused = not self.player.is_paused
            self.btn_play.setText("播放" if self.player.is_paused else "暂停")
            self.status_lbl.setText("已暂停" if self.player.is_paused else "正在播放...")
        else:
            item = self.list_widget.currentItem()
            if item: self.play_selected(item)

    def seek_position(self):
        if self.player.is_playing:
            self.player.seek_time = self.slider.value()

    def update_progress(self, pos, time_str):
        if not self.slider.isSliderDown():
            self.slider.setValue(pos)
        self.time_lbl.setText(time_str)

    def on_playback_finished(self):
        self.btn_play.setText("播放")
        self.status_lbl.setText("播放结束")
        self.slider.setValue(0)
        self.time_lbl.setText("0:00 / 0:00")

    def closeEvent(self, event):
        self.player.is_playing = False
        self.player.wait()
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())