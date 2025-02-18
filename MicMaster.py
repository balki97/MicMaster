import json
import logging
import os
import sys
import time
from ctypes import POINTER
from threading import Thread

import keyboard
import psutil
import pythoncom
import requests
import winsound
from comtypes import CLSCTX_ALL
from ctypes import cast
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from plyer import notification
from win10toast_click import ToastNotifier
from PyQt5.QtCore import Qt, QEvent, QTimer, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QSlider, QHBoxLayout,
    QCheckBox, QSystemTrayIcon, QMenu, QAction, QDialog, QComboBox, QMessageBox,
    QListWidget, QListWidgetItem, QDialogButtonBox, QAbstractItemView,
    QInputDialog, QProgressBar, QTextEdit
)

import pyaudio
import numpy as np

VERSION = "1.0.0"
SETTINGS_FILE = 'settings.json'
LOG_FILE = 'app.log'

def setup_logging(enable_logging: bool) -> None:
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()
    if enable_logging:
        try:
            handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logging.info("Logging enabled.")
        except Exception as e:
            QMessageBox.critical(None, "Logging Error", f"Failed to enable logging: {e}")

def is_process_running(exe_name: str) -> bool:
    return any(proc.info['name'] == exe_name for proc in psutil.process_iter(['name']))

def check_for_updates_notify(parent):
    try:
        parent.update_status_label.setText("Checking for updates...")
        parent.update_status_label.repaint()
        owner = "balki97"
        repo = "MicMaster"
        api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        response = requests.get(api_url, timeout=10)
        if response.status_code == 404:
            parent.update_status_label.setText("No releases found on GitHub.")
            return
        response.raise_for_status()
        latest_release = response.json()
        download_url = ""
        assets = latest_release.get('assets', [])
        if assets:
            download_url = assets[0].get('browser_download_url', '')
        latest_version = latest_release.get('tag_name', '').lstrip('v')
        if not latest_version:
            logging.warning("Latest version not found.")
            parent.update_status_label.setText("Failed to retrieve version info.")
            return
        if version_tuple(latest_version) > version_tuple(VERSION):
            msg = (f"New version {latest_version} available.\n"
                   f"You have {VERSION}.\n"
                   "Please download the latest version from:\n" + download_url)
            QMessageBox.information(parent, "Update Available", msg)
            parent.update_status_label.setText("Update available.")
        else:
            parent.update_status_label.setText("You are using the latest version.")
            logging.info("Latest version in use.")
    except Exception as e:
        logging.error(f"Error checking updates: {e}")
        parent.update_status_label.setText("Error checking updates.")
    finally:
        if hasattr(parent, 'check_updates_btn') and parent.check_updates_btn is not None:
            parent.check_updates_btn.setEnabled(True)

def version_tuple(v):
    try:
        return tuple(map(int, v.split(".")))
    except Exception:
        return (0,)

class ApplicationSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Applications to Auto-Mute")
        self.setMinimumSize(300, 400)
        layout = QVBoxLayout()
        self.process_list = QListWidget()
        self.process_list.setSelectionMode(QAbstractItemView.MultiSelection)
        layout.addWidget(self.process_list)
        running_apps = sorted({proc.name() for proc in psutil.process_iter()})
        for app in running_apps:
            self.process_list.addItem(QListWidgetItem(app))
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)
    def get_selected_apps(self) -> list:
        return [self.process_list.item(i).text() for i in range(self.process_list.count())
                if self.process_list.item(i).isSelected()]

class ProfileManagementDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Profiles")
        self.setMinimumSize(300, 400)
        self.parent_widget = parent
        layout = QVBoxLayout()
        self.profile_list = QListWidget()
        self.profile_list.addItems(self.parent_widget.profiles)
        self.profile_list.setCurrentRow(self.parent_widget.current_profile_index)
        layout.addWidget(self.profile_list)
        btn_layout = QHBoxLayout()
        self.new_profile_btn = QPushButton("New Profile")
        self.new_profile_btn.clicked.connect(self.create_profile)
        btn_layout.addWidget(self.new_profile_btn)
        self.rename_profile_btn = QPushButton("Rename Profile")
        self.rename_profile_btn.clicked.connect(self.rename_profile)
        btn_layout.addWidget(self.rename_profile_btn)
        self.delete_profile_btn = QPushButton("Delete Profile")
        self.delete_profile_btn.clicked.connect(self.delete_profile)
        btn_layout.addWidget(self.delete_profile_btn)
        layout.addLayout(btn_layout)
        ok_cancel_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)
        ok_cancel_layout.addWidget(self.ok_btn)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        ok_cancel_layout.addWidget(self.cancel_btn)
        layout.addLayout(ok_cancel_layout)
        self.setLayout(layout)
    def create_profile(self):
        profile_name, ok = QInputDialog.getText(self, "New Profile", "Enter profile name:")
        if ok and profile_name:
            if profile_name in self.parent_widget.profiles:
                QMessageBox.warning(self, "Duplicate Profile", "Profile already exists.")
                return
            self.parent_widget.profiles.append(profile_name)
            self.parent_widget.settings['profiles'][profile_name] = self.parent_widget.default_profile_settings()
            self.parent_widget.save_settings()
            self.profile_list.addItem(QListWidgetItem(profile_name))
            logging.info(f"Profile '{profile_name}' created.")
    def rename_profile(self):
        selected = self.profile_list.selectedItems()
        if not selected:
            return
        old_name = selected[0].text()
        new_name, ok = QInputDialog.getText(self, "Rename Profile", "Enter new profile name:", text=old_name)
        if ok and new_name:
            if new_name in self.parent_widget.profiles:
                QMessageBox.warning(self, "Duplicate Profile", "Profile already exists.")
                return
            idx = self.parent_widget.profiles.index(old_name)
            self.parent_widget.profiles[idx] = new_name
            self.parent_widget.settings['profiles'][new_name] = self.parent_widget.settings['profiles'].pop(old_name)
            self.parent_widget.save_settings()
            selected[0].setText(new_name)
            logging.info(f"Profile '{old_name}' renamed to '{new_name}'.")
    def delete_profile(self):
        selected = self.profile_list.selectedItems()
        if not selected:
            return
        profile_name = selected[0].text()
        if profile_name == "Default":
            QMessageBox.warning(self, "Delete Profile", "Default profile cannot be deleted.")
            return
        reply = QMessageBox.question(self, "Delete Profile",
                                     f"Delete profile '{profile_name}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            idx = self.parent_widget.profiles.index(profile_name)
            self.parent_widget.profiles.pop(idx)
            self.parent_widget.settings['profiles'].pop(profile_name)
            self.parent_widget.save_settings()
            self.profile_list.takeItem(self.profile_list.row(selected[0]))
            logging.info(f"Profile '{profile_name}' deleted.")

class LogViewerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("View Logs")
        self.resize(600, 400)
        layout = QVBoxLayout()
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        btn_box = QDialogButtonBox(QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.setLayout(layout)
        self.load_logs()
    def load_logs(self):
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                self.log_text.setPlainText(f.read())
        else:
            self.log_text.setPlainText("Log file not found.")

class SettingsWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        self.setWindowTitle("Settings")
        self.setMinimumSize(400, 600)
        layout = QVBoxLayout()
        profile_layout = QHBoxLayout()
        self.profile_label = QLabel("Current Profile:")
        profile_layout.addWidget(self.profile_label)
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(self.parent_widget.profiles)
        self.profile_combo.setCurrentIndex(self.parent_widget.current_profile_index)
        self.profile_combo.currentIndexChanged.connect(self.switch_profile)
        profile_layout.addWidget(self.profile_combo)
        self.manage_profiles_btn = QPushButton("Manage Profiles")
        self.manage_profiles_btn.clicked.connect(self.manage_profiles)
        profile_layout.addWidget(self.manage_profiles_btn)
        layout.addLayout(profile_layout)
        self.volume_label = QLabel("Default Volume:")
        layout.addWidget(self.volume_label)
        vol_layout = QHBoxLayout()
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setMinimum(0)
        self.volume_slider.setMaximum(100)
        self.volume_slider.setValue(100)
        self.volume_slider.setTickInterval(10)
        self.volume_slider.setTickPosition(QSlider.TicksBelow)
        self.volume_slider.valueChanged.connect(self.update_volume_label)
        vol_layout.addWidget(self.volume_slider)
        self.volume_value_label = QLabel("100%")
        vol_layout.addWidget(self.volume_value_label)
        layout.addLayout(vol_layout)
        self.startup_checkbox = QCheckBox("Start on system boot")
        layout.addWidget(self.startup_checkbox)
        self.notifications_checkbox = QCheckBox("Enable desktop notifications")
        layout.addWidget(self.notifications_checkbox)
        self.sound_notification_checkbox = QCheckBox("Enable sound notifications instead of desktop notifications")
        layout.addWidget(self.sound_notification_checkbox)
        self.theme_label = QLabel("Theme:")
        layout.addWidget(self.theme_label)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        layout.addWidget(self.theme_combo)
        self.enable_auto_mute_checkbox = QCheckBox("Enable Auto-Mute on Specific Applications")
        layout.addWidget(self.enable_auto_mute_checkbox)
        self.auto_mute_label = QLabel("Applications to Auto-Mute:")
        layout.addWidget(self.auto_mute_label)
        self.app_list = QListWidget()
        layout.addWidget(self.app_list)
        apps_btn_layout = QHBoxLayout()
        self.select_apps_btn = QPushButton("Select Applications")
        self.select_apps_btn.clicked.connect(self.select_applications)
        apps_btn_layout.addWidget(self.select_apps_btn)
        self.remove_app_btn = QPushButton("Remove Selected")
        self.remove_app_btn.clicked.connect(self.remove_app)
        apps_btn_layout.addWidget(self.remove_app_btn)
        layout.addLayout(apps_btn_layout)
        self.tray_checkbox = QCheckBox("Enable Minimize to Tray")
        layout.addWidget(self.tray_checkbox)
        self.desktop_shortcut_checkbox = QCheckBox("Create Desktop Shortcut")
        layout.addWidget(self.desktop_shortcut_checkbox)
        btns_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_settings)
        btns_layout.addWidget(self.save_btn)
        self.reset_btn = QPushButton("Reset to Default")
        self.reset_btn.clicked.connect(self.reset_settings)
        btns_layout.addWidget(self.reset_btn)
        self.view_logs_btn = QPushButton("View Logs")
        self.view_logs_btn.clicked.connect(self.open_log_viewer)
        btns_layout.addWidget(self.view_logs_btn)
        layout.addLayout(btns_layout)
        self.setLayout(layout)
        self.load_settings()
        self.enable_auto_mute_checkbox.stateChanged.connect(self.toggle_auto_mute)
    def manage_profiles(self):
        dialog = ProfileManagementDialog(self.parent_widget)
        if dialog.exec_():
            self.profile_combo.clear()
            self.profile_combo.addItems(self.parent_widget.profiles)
            self.profile_combo.setCurrentIndex(self.parent_widget.current_profile_index)
    def switch_profile(self, index: int):
        if index != self.parent_widget.current_profile_index:
            logging.info(f"Switching profile from {self.parent_widget.current_profile_index} to {index}.")
            if index < len(self.parent_widget.profiles):
                self.parent_widget.current_profile_index = index
                self.parent_widget.settings['current_profile'] = index
                self.parent_widget.load_current_profile()
                QMessageBox.information(self, "Profile Switched", f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                logging.info(f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                self.load_settings()
            else:
                logging.error(f"Invalid profile index: {index}. Reverting.")
                QMessageBox.warning(self, "Profile Switch Failed", "Invalid profile. Reverting.")
                self.parent_widget.current_profile_index = 0
                self.parent_widget.settings['current_profile'] = 0
                self.parent_widget.load_current_profile()
                self.profile_combo.setCurrentIndex(0)
                self.load_settings()
    def select_applications(self):
        dialog = ApplicationSelectionDialog(self)
        if dialog.exec_():
            selected_apps = dialog.get_selected_apps()
            for app in selected_apps:
                if not self.is_app_in_list(app):
                    self.app_list.addItem(app)
    def is_app_in_list(self, app_name: str) -> bool:
        return any(self.app_list.item(i).text().lower() == app_name.lower() for i in range(self.app_list.count()))
    def remove_app(self):
        for item in self.app_list.selectedItems():
            self.app_list.takeItem(self.app_list.row(item))
    def toggle_auto_mute(self, state):
        enabled = (state == Qt.Checked)
        self.select_apps_btn.setEnabled(enabled)
        self.auto_mute_label.setEnabled(enabled)
    def load_settings(self):
        profile = self.parent_widget.get_current_profile()
        self.volume_slider.setValue(profile.get('volume', 100))
        self.volume_value_label.setText(f"{profile.get('volume', 100)}%")
        self.startup_checkbox.setChecked(profile.get('startup', False))
        self.notifications_checkbox.setChecked(profile.get('notifications', False))
        self.sound_notification_checkbox.setChecked(profile.get('sound_notifications', False))
        self.theme_combo.setCurrentText(profile.get('theme', 'Dark'))
        self.enable_auto_mute_checkbox.setChecked(profile.get('enable_auto_mute', False))
        self.app_list.clear()
        self.app_list.addItems(profile.get('auto_mute_apps', []))
        self.tray_checkbox.setChecked(profile.get('tray_enabled', False))
        self.desktop_shortcut_checkbox.setChecked(profile.get('create_desktop_shortcut', False))
        self.toggle_auto_mute(self.enable_auto_mute_checkbox.isChecked())
    def update_volume_label(self, value):
        self.volume_value_label.setText(f"{value}%")
    def save_settings(self):
        profile = self.parent_widget.get_current_profile()
        profile['volume'] = self.volume_slider.value()
        profile['startup'] = self.startup_checkbox.isChecked()
        profile['notifications'] = self.notifications_checkbox.isChecked()
        profile['sound_notifications'] = self.sound_notification_checkbox.isChecked()
        profile['theme'] = self.theme_combo.currentText()
        profile['enable_auto_mute'] = self.enable_auto_mute_checkbox.isChecked()
        profile['auto_mute_apps'] = [self.app_list.item(i).text() for i in range(self.app_list.count())]
        profile['tray_enabled'] = self.tray_checkbox.isChecked()
        profile['create_desktop_shortcut'] = self.desktop_shortcut_checkbox.isChecked()
        self.parent_widget.settings['current_profile'] = self.parent_widget.current_profile_index
        self.parent_widget.save_settings()
        if profile['create_desktop_shortcut']:
            self.parent_widget.create_desktop_shortcut_method()
        else:
            self.parent_widget.remove_desktop_shortcut_method()
        if profile['startup']:
            self.parent_widget.add_to_startup()
        else:
            self.parent_widget.remove_from_startup()
        setup_logging(profile.get('enable_logging', True))
        QMessageBox.information(self, "Settings Saved", "Settings saved successfully.")
        self.accept()
    def reset_settings(self):
        profile = self.parent_widget.get_current_profile()
        default = self.parent_widget.default_profile_settings()
        for key in default:
            profile[key] = default[key]
        self.load_settings()
    def open_log_viewer(self):
        dialog = LogViewerDialog(self)
        dialog.exec_()
    def resource_path(self, relative_path: str) -> str:
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

class HotkeyListener(Thread):
    def __init__(self, callback, hotkey: str):
        super().__init__()
        self.callback = callback
        self.hotkey = hotkey
        self.daemon = True
    def run(self):
        try:
            keyboard.add_hotkey(self.hotkey, self.callback)
            logging.info(f"Hotkey '{self.hotkey}' listener started.")
            keyboard.wait()
        except Exception as e:
            logging.error(f"Error in HotkeyListener: {e}")

class AudioStreamThread(Thread):
    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent
        self.running = True
        self.chunk = 1024
        self.format = pyaudio.paInt16
        self.channels = 1
        self.rate = 44100
    def run(self):
        p = pyaudio.PyAudio()
        try:
            stream = p.open(format=self.format,
                            channels=self.channels,
                            rate=self.rate,
                            input=True,
                            frames_per_buffer=self.chunk)
        except Exception as e:
            logging.error(f"Error opening audio stream: {e}")
            QMessageBox.critical(None, "Error", "Failed to open audio stream.")
            return
        while self.running:
            try:
                data = stream.read(self.chunk, exception_on_overflow=False)
                audio_data = np.frombuffer(data, dtype=np.int16)
                peak = np.abs(audio_data).max()
                level = int((peak / 32768) * 100)
                self.parent.audio_level_signal.emit(level)
            except Exception as e:
                logging.error(f"Error reading audio stream: {e}")
                break
        stream.stop_stream()
        stream.close()
        p.terminate()
    def stop(self):
        self.running = False

class MicMaster(QWidget):
    toggle_mute_signal = pyqtSignal()
    audio_level_signal = pyqtSignal(int)
    def __init__(self):
        super().__init__()
        self.is_muted = False
        self.device = None
        self.interface = None
        self.volume = None
        self.current_hotkey = None
        self.tray_icon = None
        self.tray_enabled = False
        self.use_sound_notification = False
        self.notifications_enabled = False
        self.auto_mute_apps = []
        self.enable_auto_mute = False
        self.notifier = ToastNotifier()
        self.profiles = []
        self.current_profile_index = 0
        self.settings = {}
        self.original_mic_off_icon = QIcon(self.resource_path(os.path.join("images", "mic_off.png")))
        self.tinted_mic_off_icon = QIcon(self.tint_pixmap(os.path.join("images", "mic_off.png"), "red"))
        self.mic_on_icon = QIcon(self.resource_path(os.path.join("images", "mic_on.png")))
        self.setWindowIcon(QIcon(self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))))
        self.audio_level_signal.connect(self.update_audio_level_visualization)
        self.initUI()
        self.load_settings()
        setup_logging(self.settings.get('enable_logging', True))
        self.init_device()
        self.init_tray_icon()
        self.load_hotkey()
        self.toggle_mute_signal.connect(self.toggle_mute)
        self.check_for_updates()  # Check updates on startup
        self.audio_thread = AudioStreamThread(self)
        self.audio_thread.start()
    def resource_path(self, relative_path: str) -> str:
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    def tint_pixmap(self, pixmap_path: str, color: str) -> QPixmap:
        pixmap = QPixmap(self.resource_path(pixmap_path))
        tinted = QPixmap(pixmap.size())
        tinted.fill(Qt.transparent)
        painter = QPainter(tinted)
        painter.drawPixmap(0, 0, pixmap)
        painter.setCompositionMode(QPainter.CompositionMode_SourceAtop)
        painter.fillRect(tinted.rect(), QColor(color))
        painter.end()
        return tinted
    def initUI(self):
        layout = QVBoxLayout()
        self.mute_btn = QPushButton("Mute Mic", self)
        self.mute_btn.setIcon(self.original_mic_off_icon)
        self.mute_btn.clicked.connect(self.toggle_mute)
        self.mute_btn.setToolTip("Mute/unmute microphone.")
        layout.addWidget(self.mute_btn)
        vol_layout = QHBoxLayout()
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setMinimum(0)
        self.volume_slider.setMaximum(100)
        self.volume_slider.setValue(100)
        self.volume_slider.setTickInterval(10)
        self.volume_slider.setTickPosition(QSlider.TicksBelow)
        self.volume_slider.valueChanged.connect(self.set_volume)
        vol_layout.addWidget(self.volume_slider)
        self.volume_label = QLabel("100%")
        vol_layout.addWidget(self.volume_label)
        layout.addLayout(vol_layout)
        self.hotkey_label = QLabel("Recorded Hotkey: None", self)
        layout.addWidget(self.hotkey_label)
        self.record_hotkey_btn = QPushButton("Record Hotkey", self)
        self.record_hotkey_btn.clicked.connect(self.start_recording)
        self.record_hotkey_btn.setToolTip("Record a hotkey.")
        layout.addWidget(self.record_hotkey_btn)
        self.stop_record_btn = QPushButton("Stop Recording", self)
        self.stop_record_btn.clicked.connect(self.stop_recording)
        self.stop_record_btn.setToolTip("Stop recording hotkey.")
        layout.addWidget(self.stop_record_btn)
        self.settings_btn = QPushButton("Settings", self)
        self.settings_btn.clicked.connect(self.open_settings)
        self.settings_btn.setToolTip("Open settings.")
        layout.addWidget(self.settings_btn)
        self.check_updates_btn = QPushButton("Check for Updates", self)
        self.check_updates_btn.clicked.connect(lambda: check_for_updates_notify(self))
        self.check_updates_btn.setToolTip("Check for updates.")
        layout.addWidget(self.check_updates_btn)
        self.update_status_label = QLabel("")
        layout.addWidget(self.update_status_label)
        self.help_btn = QPushButton("Help", self)
        self.help_btn.clicked.connect(self.show_help)
        self.help_btn.setToolTip("Show help.")
        layout.addWidget(self.help_btn)
        self.audio_level_label = QLabel("Audio Level: 0%", self)
        layout.addWidget(self.audio_level_label)
        self.audio_level_visual = QProgressBar(self)
        self.audio_level_visual.setMinimum(0)
        self.audio_level_visual.setMaximum(100)
        self.audio_level_visual.setValue(0)
        layout.addWidget(self.audio_level_visual)
        self.version_label = QLabel(f"v{VERSION}", self)
        font = self.version_label.font()
        font.setPointSize(8)
        self.version_label.setFont(font)
        self.version_label.setStyleSheet("color: gray;")
        layout.addWidget(self.version_label, alignment=Qt.AlignRight)
        self.setLayout(layout)
        self.setWindowTitle('MicMaster')
        self.resize(300, 550)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)
    def show_help(self):
        help_text = ("MicMaster Help:\n"
                     "- Toggle mute with the button or hotkey.\n"
                     "- Configure settings and profiles in the Settings menu.\n"
                     "- Enable auto-mute for specified applications.\n"
                     "- Real-time audio level display.\n"
                     "- Minimize to tray to run in background.\n"
                     "- When an update is available, you will be notified with a download link.\n"
                     f"- Current Version: {VERSION}")
        QMessageBox.information(self, "MicMaster Help", help_text)
    def apply_theme(self, theme_name: str):
        if theme_name == 'Dark':
            self.setStyleSheet("""
                QWidget { background-color: #1e1e1e; color: #c5c6c7; font-family: "Segoe UI"; font-size: 10pt; }
                QPushButton { background-color: #282a36; border: 2px solid #44475a; border-radius: 8px; padding: 8px; color: #f8f8f2; }
                QPushButton:hover { background-color: #44475a; }
                QPushButton:pressed { background-color: #6272a4; }
                QSlider::groove:horizontal { height: 8px; background: #44475a; border-radius: 4px; }
                QSlider::handle:horizontal { background: #50fa7b; border: 1px solid #bd93f9; width: 14px; margin: -3px 0; border-radius: 7px; }
                QLabel { color: #f8f8f2; }
                QComboBox { background-color: #282a36; border: 1px solid #44475a; border-radius: 8px; padding: 4px; color: #f8f8f2; }
                QProgressBar { border: 2px solid #44475a; border-radius: 8px; background-color: #282a36; text-align: center; }
                QProgressBar::chunk { background-color: #50fa7b; border-radius: 4px; }
            """)
            logging.info("Dark theme applied.")
        else:
            self.setStyleSheet("""
                QWidget { background-color: #f5f5f5; color: #2e2e2e; font-family: "Segoe UI"; font-size: 10pt; }
                QPushButton { background-color: #ffffff; border: 2px solid #c5c5c5; border-radius: 8px; padding: 8px; color: #2e2e2e; }
                QPushButton:hover { background-color: #e0e0e0; }
                QPushButton:pressed { background-color: #d5d5d5; }
                QSlider::groove:horizontal { height: 8px; background: #c5c5c5; border-radius: 4px; }
                QSlider::handle:horizontal { background: #0078d7; border: 1px solid #005a9e; width: 14px; margin: -3px 0; border-radius: 7px; }
                QLabel { color: #2e2e2e; }
                QComboBox { background-color: #ffffff; border: 1px solid #c5c5c5; border-radius: 8px; padding: 4px; color: #2e2e2e; }
                QProgressBar { border: 2px solid #c5c5c5; border-radius: 8px; background-color: #ffffff; text-align: center; }
                QProgressBar::chunk { background-color: #0078d7; border-radius: 4px; }
            """)
            logging.info("Light theme applied.")
    def open_settings(self):
        settings_window = SettingsWindow(self)
        if settings_window.exec_():
            self.load_settings()
            self.setup_auto_mute()
    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    self.settings = json.load(f)
                self.profiles = list(self.settings.get('profiles', {}).keys())
                if not self.profiles:
                    self.profiles = ['Default']
                    self.settings['profiles'] = {'Default': self.default_profile_settings()}
                self.current_profile_index = self.settings.get('current_profile', 0)
                if self.current_profile_index >= len(self.profiles):
                    self.current_profile_index = 0
                self.apply_profile_settings()
            except Exception as e:
                logging.error(f"Error loading settings: {e}")
                QMessageBox.critical(self, "Error", "Failed to load settings.")
        else:
            self.settings = {'profiles': {'Default': self.default_profile_settings()}, 'current_profile': 0}
            self.profiles = ['Default']
            self.current_profile_index = 0
            self.save_settings()
    def save_settings(self):
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(self.settings, f, indent=4)
            logging.info("Settings saved.")
        except Exception as e:
            logging.error(f"Error saving settings: {e}")
            QMessageBox.critical(self, "Error", "Failed to save settings.")
    def default_profile_settings(self) -> dict:
        return {
            'volume': 100,
            'startup': False,
            'notifications': False,
            'sound_notifications': False,
            'theme': 'Dark',
            'enable_auto_mute': False,
            'auto_mute_apps': [],
            'tray_enabled': False,
            'create_desktop_shortcut': False,
            'hotkey': None
        }
    def get_current_profile(self) -> dict:
        profile_name = self.profiles[self.current_profile_index]
        return self.settings['profiles'][profile_name]
    def apply_profile_settings(self):
        profile = self.get_current_profile()
        if profile:
            self.volume_slider.setValue(profile.get('volume', 100))
            self.use_sound_notification = profile.get('sound_notifications', False)
            self.notifications_enabled = profile.get('notifications', False)
            self.tray_enabled = profile.get('tray_enabled', False)
            self.enable_auto_mute = profile.get('enable_auto_mute', False)
            self.auto_mute_apps = profile.get('auto_mute_apps', [])
            self.hotkey = profile.get('hotkey', None)
            self.apply_theme(profile.get('theme', 'Dark'))
            logging.info(f"Applied profile: {self.profiles[self.current_profile_index]}")
        else:
            logging.error(f"Profile index {self.current_profile_index} does not exist.")
    def add_to_startup(self):
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
            startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
            shortcut_path = os.path.join(startup_folder, 'MicMaster.lnk')
            from win32com.client import Dispatch
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = exe_path
            shortcut.Arguments = ""
            shortcut.WorkingDirectory = os.path.dirname(exe_path)
            shortcut.IconLocation = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
            shortcut.save()
            logging.info("Added to startup.")
        except Exception as e:
            logging.error(f"Error adding to startup: {e}")
            QMessageBox.critical(self, "Error", "Failed to add to startup.")
    def remove_from_startup(self):
        try:
            startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
            shortcut_path = os.path.join(startup_folder, 'MicMaster.lnk')
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                logging.info("Removed from startup.")
        except Exception as e:
            logging.error(f"Error removing from startup: {e}")
            QMessageBox.critical(self, "Error", "Failed to remove startup entry.")
    def init_device(self):
        try:
            pythoncom.CoInitialize()
            devices = AudioUtilities.GetMicrophone()
            if not devices:
                logging.error("No microphone found.")
                QMessageBox.critical(self, "Error", "No microphone found.")
                return
            self.device = devices
            self.interface = self.device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))
            current_volume = self.volume.GetMasterVolumeLevelScalar() * 100
            self.volume_slider.setValue(int(current_volume))
            self.volume_label.setText(f"{int(current_volume)}%")
        except Exception as e:
            logging.error(f"Error initializing microphone control: {e}")
            QMessageBox.critical(self, "Error", "Failed to initialize microphone control.")
    def init_tray_icon(self):
        tray_icon_path = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
        self.tray_icon = QSystemTrayIcon(QIcon(tray_icon_path), self)
        tray_menu = QMenu(self)
        restore_action = QAction("Restore", self)
        restore_action.triggered.connect(self.show_normal)
        tray_menu.addAction(restore_action)
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)
        switch_profile_action = QAction("Switch Profile", self)
        switch_profile_action.triggered.connect(lambda: self.open_settings())
        tray_menu.addAction(switch_profile_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        if self.tray_enabled:
            self.hide()
            self.tray_icon.show()
            self.tray_icon.showMessage("MicMaster", "Minimized to tray", QSystemTrayIcon.Information, 2000)
        elif self.tray_icon:
            self.tray_icon.hide()
    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_normal()
    def show_normal(self):
        self.show()
        self.activateWindow()
    def toggle_tray_option(self, state):
        self.tray_enabled = (state == Qt.Checked)
        profile = self.get_current_profile()
        profile['tray_enabled'] = self.tray_enabled
        self.save_settings()
        if self.tray_enabled:
            self.init_tray_icon()
        else:
            if self.tray_icon:
                self.tray_icon.hide()
    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange and self.isMinimized():
            profile = self.get_current_profile()
            if profile.get('tray_enabled', False):
                QTimer.singleShot(0, self.hide)
                self.tray_icon.showMessage("MicMaster", "Minimized to tray", QSystemTrayIcon.Information, 2000)
        super().changeEvent(event)
    def closeEvent(self, event):
        self.audio_thread.stop()
        self.audio_thread.join()
        event.accept()
    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.hide()
        self.audio_thread.stop()
        self.audio_thread.join()
        QApplication.quit()
    def start_recording(self):
        if getattr(self, 'recording', False):
            return
        self.recording = True
        self.pressed_keys = set()
        self.hotkey_label.setStyleSheet("color: red;")
        self.hotkey_label.setText("Recording hotkey... Press 'Stop Recording'.")
        keyboard.hook(self.record_key)
    def stop_recording(self):
        if not getattr(self, 'recording', False):
            self.hotkey_label.setText("No recording in progress.")
            return
        self.recording = False
        self.hotkey_label.setStyleSheet("color: white;")
        keyboard.unhook_all()
        if not self.pressed_keys:
            self.hotkey_label.setText("Error: No keys recorded.")
            return
        self.hotkey = "+".join(sorted(self.pressed_keys))
        if self.current_hotkey:
            try:
                keyboard.remove_hotkey(self.current_hotkey)
                logging.info(f"Removed previous hotkey: {self.current_hotkey}")
            except ValueError:
                logging.warning(f"Hotkey {self.current_hotkey} not found.")
        try:
            self.current_hotkey = keyboard.add_hotkey(self.hotkey, self.emit_toggle_mute_signal)
            self.hotkey_label.setText(f"Recorded Hotkey: {self.hotkey}")
            logging.info(f"Hotkey recorded: {self.hotkey}")
            profile = self.get_current_profile()
            profile['hotkey'] = self.hotkey
            self.save_settings()
        except Exception as e:
            logging.error(f"Error setting hotkey: {e}")
            self.hotkey_label.setText("Error: Invalid hotkey.")
    @pyqtSlot()
    def emit_toggle_mute_signal(self):
        try:
            self.toggle_mute_signal.emit()
            logging.info("Toggle mute signal emitted.")
        except Exception as e:
            logging.error(f"Error emitting toggle signal: {e}")
    def record_key(self, e):
        if e.event_type == "down" and e.name not in self.pressed_keys:
            self.pressed_keys.add(e.name)
            self.hotkey_label.setText(f"Recording hotkey: {' + '.join(sorted(self.pressed_keys))}")
    def load_hotkey(self):
        profile = self.get_current_profile()
        hotkey = profile.get('hotkey', None)
        if hotkey:
            self.hotkey = hotkey
            try:
                self.hotkey_listener = HotkeyListener(self.emit_toggle_mute_signal, self.hotkey)
                self.hotkey_listener.start()
                self.hotkey_label.setText(f"Recorded Hotkey: {self.hotkey}")
                logging.info(f"Hotkey loaded: {self.hotkey}")
            except Exception as e:
                logging.error(f"Error loading hotkey: {e}")
                self.hotkey_label.setText("Error: Invalid hotkey.")
    def setup_auto_mute(self):
        profile = self.get_current_profile()
        self.enable_auto_mute = profile.get('enable_auto_mute', False)
        self.auto_mute_apps = profile.get('auto_mute_apps', [])
    def check_auto_mute_apps(self):
        if not self.auto_mute_apps or not self.enable_auto_mute:
            return
        try:
            running_apps = [proc.name().lower() for proc in psutil.process_iter()]
            should_mute = any(app.lower() in running_apps for app in self.auto_mute_apps)
            if should_mute and not self.is_muted:
                self.toggle_mute()
            elif not should_mute and self.is_muted:
                self.toggle_mute()
        except Exception as e:
            logging.error(f"Error checking auto-mute: {e}")
    def toggle_mute(self):
        try:
            pythoncom.CoInitialize()
            self.is_muted = not self.is_muted
            self.mute_btn.setText("Unmute Mic" if self.is_muted else "Mute Mic")
            self.mute_btn.setIcon(self.tinted_mic_off_icon if self.is_muted else self.mic_on_icon)
            self.mute_microphone(self.is_muted)
            self.send_notification()
            logging.info(f"Microphone {'muted' if self.is_muted else 'unmuted'}.")
            if self.tray_icon:
                path = self.resource_path(os.path.join("images", "mic_off.png")) if self.is_muted else self.resource_path(os.path.join("images", "mic_on.png"))
                self.tray_icon.setIcon(QIcon(path))
        except Exception as e:
            logging.error(f"Error toggling mute: {e}")
            QMessageBox.critical(self, "Error", "Failed to toggle microphone.")
    def send_notification(self):
        status = "Muted" if self.is_muted else "Unmuted"
        if self.use_sound_notification:
            sound_file = os.path.join("sounds", "mute_sound.wav") if self.is_muted else os.path.join("sounds", "unmute_sound.wav")
            try:
                winsound.PlaySound(self.resource_path(sound_file), winsound.SND_FILENAME | winsound.SND_ASYNC)
            except Exception as e:
                logging.error(f"Error playing sound: {e}")
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION if self.is_muted else winsound.MB_OK)
        elif self.notifications_enabled:
            try:
                self.notifier.show_toast("MicMaster", f"Microphone {status}",
                                         icon_path=self.resource_path(os.path.join("icons", "mic_switch_icon.ico")),
                                         duration=5, threaded=True,
                                         callback_on_click=self.handle_toggle_mute_callback)
            except Exception as e:
                logging.error(f"Error showing notification: {e}")
                notification.notify(title="MicMaster", message=f"Microphone {status}", timeout=2)
    def handle_toggle_mute_callback(self):
        try:
            self.toggle_mute()
            logging.info("Toggled via notification.")
        except Exception as e:
            logging.error(f"Error in notification callback: {e}")
        return 0
    def mute_microphone(self, mute: bool):
        try:
            if self.volume:
                self.volume.SetMute(1 if mute else 0, None)
                logging.info(f"Microphone {'muted' if mute else 'unmuted'}.")
        except Exception as e:
            logging.error(f"Error controlling microphone: {e}")
            QMessageBox.critical(self, "Error", "Failed to control microphone.")
    def set_volume(self, value: int):
        if self.volume:
            vol_level = value / 100.0
            try:
                self.volume.SetMasterVolumeLevelScalar(vol_level, None)
                self.volume_label.setText(f"{value}%")
                logging.info(f"Volume set to {value}%.")
            except Exception as e:
                logging.error(f"Error setting volume: {e}")
                QMessageBox.critical(self, "Error", "Failed to set volume.")
    def update_audio_level_visualization(self, level: int):
        self.audio_level_visual.setValue(level)
        self.audio_level_label.setText(f"Audio Level: {level}%")
    def check_for_updates(self):
        if hasattr(self, 'check_updates_btn') and self.check_updates_btn is not None:
            self.check_updates_btn.setEnabled(False)
        check_for_updates_notify(self)
    def create_desktop_shortcut_method(self):
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
            desktop_folder = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop_folder, 'MicMaster.lnk')
            if not os.path.exists(shortcut_path):
                from win32com.client import Dispatch
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.Arguments = ""
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                icon_location = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
                if os.path.exists(icon_location):
                    shortcut.IconLocation = icon_location
                shortcut.save()
                logging.info("Desktop shortcut created.")
                try:
                    self.notifier.show_toast("MicMaster", "Desktop shortcut created.",
                                               icon_path=self.resource_path(os.path.join("icons", "mic_switch_icon.ico")),
                                               duration=2, threaded=True)
                except Exception as e:
                    logging.error(f"Error showing notification: {e}")
            else:
                logging.info("Desktop shortcut already exists.")
        except Exception as e:
            logging.error(f"Error creating desktop shortcut: {e}")
            QMessageBox.critical(self, "Error", "Failed to create desktop shortcut.")
    def remove_desktop_shortcut_method(self):
        try:
            desktop_folder = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop_folder, 'MicMaster.lnk')
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                logging.info("Desktop shortcut removed.")
            else:
                logging.info("Desktop shortcut does not exist.")
        except Exception as e:
            logging.error(f"Error removing desktop shortcut: {e}")
            QMessageBox.critical(self, "Error", "Failed to remove desktop shortcut.")
    def load_current_profile(self):
        try:
            self.apply_profile_settings()
            setup_logging(self.settings.get('enable_logging', True))
            self.setup_auto_mute()
            self.load_hotkey()
        except Exception as e:
            logging.error(f"Error loading profile: {e}")
            QMessageBox.critical(self, "Error", "Failed to load profile.")
            self.current_profile_index = 0
            self.settings['current_profile'] = 0
            self.apply_profile_settings()
    def get_profiles(self) -> list:
        return self.profiles
    def closeEvent(self, event):
        self.audio_thread.stop()
        self.audio_thread.join()
        event.accept()
    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.hide()
        self.audio_thread.stop()
        self.audio_thread.join()
        QApplication.quit()

class MicMasterApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        icons_dir = os.path.join(os.path.dirname(__file__), "icons")
        if not os.path.exists(icons_dir):
            setup_logging(True)
            logging.error("Icons directory not found.")
            QMessageBox.critical(None, "Error", "Icons directory not found.")
            sys.exit(1)
        icon_path = os.path.join(icons_dir, "mic_switch_icon.ico")
        if not os.path.exists(icon_path):
            setup_logging(True)
            logging.error("Application icon not found.")
            QMessageBox.critical(None, "Error", "Application icon not found.")
            sys.exit(1)
        self.app.setWindowIcon(QIcon(icon_path))
        self.window = MicMaster()
        self.window.show()
    def run(self):
        sys.exit(self.app.exec_())

def main():
    pythoncom.CoInitialize()
    app_instance = MicMasterApp()
    app_instance.run()

if __name__ == '__main__':
    main()
