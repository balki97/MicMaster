import json
import logging
import os
import shutil
import subprocess
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
    QCheckBox, QSystemTrayIcon, QMenu, QAction, QDialog, QSpinBox, QComboBox,
    QMessageBox, QListWidget, QListWidgetItem, QDialogButtonBox, QAbstractItemView,
    QInputDialog, QProgressBar
)

import pyaudio
import numpy as np

VERSION = "1.0.5"
SETTINGS_FILE = 'settings.json'
LOG_FILE = 'app.log'


def setup_logging(enable_logging: bool):
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
            logging.info("Logging has been enabled.")
        except Exception as e:
            QMessageBox.critical(None, "Logging Error", f"Failed to enable logging: {e}")


def is_process_running(exe_name: str) -> bool:
    return any(proc.info['name'] == exe_name for proc in psutil.process_iter(['name']))


def perform_update(new_exe_path: str, current_exe_path: str):
    setup_logging(True)
    logging.info("Updater started.")
    logging.info(f"New executable path: {new_exe_path}")
    logging.info(f"Current executable path: {current_exe_path}")

    try:
        logging.info("Waiting for the main application to exit...")
        for _ in range(30):
            if not is_process_running(os.path.basename(current_exe_path)):
                break
            time.sleep(1)
        else:
            logging.error("Main application did not exit within the expected time.")
            QMessageBox.critical(None, "Update Error", "Failed to update the application because it is still running.")
            sys.exit(1)

        shutil.move(new_exe_path, current_exe_path)
        logging.info("Executable replaced successfully.")
        subprocess.Popen([current_exe_path])
        logging.info("Application restarted successfully.")
    except Exception as e:
        logging.error(f"Error during update: {e}")
        QMessageBox.critical(None, "Update Error", f"An error occurred during the update: {e}")
        sys.exit(1)


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
        return [item.text() for item in self.process_list.selectedItems()]


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

        buttons_layout = QHBoxLayout()
        self.new_profile_btn = QPushButton("New Profile")
        self.new_profile_btn.clicked.connect(self.create_profile)
        buttons_layout.addWidget(self.new_profile_btn)

        self.rename_profile_btn = QPushButton("Rename Profile")
        self.rename_profile_btn.clicked.connect(self.rename_profile)
        buttons_layout.addWidget(self.rename_profile_btn)

        self.delete_profile_btn = QPushButton("Delete Profile")
        self.delete_profile_btn.clicked.connect(self.delete_profile)
        buttons_layout.addWidget(self.delete_profile_btn)

        layout.addLayout(buttons_layout)

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
                QMessageBox.warning(self, "Duplicate Profile", "A profile with this name already exists.")
                return
            self.parent_widget.profiles.append(profile_name)
            self.parent_widget.settings['profiles'][profile_name] = self.parent_widget.default_profile_settings()
            self.parent_widget.save_settings()
            self.profile_list.addItem(QListWidgetItem(profile_name))
            logging.info(f"Profile '{profile_name}' created.")

    def rename_profile(self):
        selected_items = self.profile_list.selectedItems()
        if not selected_items:
            return
        old_name = selected_items[0].text()
        new_name, ok = QInputDialog.getText(self, "Rename Profile", "Enter new profile name:", text=old_name)
        if ok and new_name:
            if new_name in self.parent_widget.profiles:
                QMessageBox.warning(self, "Duplicate Profile", "A profile with this name already exists.")
                return
            index = self.parent_widget.profiles.index(old_name)
            self.parent_widget.profiles[index] = new_name
            self.parent_widget.settings['profiles'][new_name] = self.parent_widget.settings['profiles'].pop(old_name)
            self.parent_widget.save_settings()
            self.profile_list.currentItem().setText(new_name)
            logging.info(f"Profile '{old_name}' renamed to '{new_name}'.")

    def delete_profile(self):
        selected_items = self.profile_list.selectedItems()
        if not selected_items:
            return
        profile_name = selected_items[0].text()
        if profile_name == "Default":
            QMessageBox.warning(self, "Delete Profile", "The 'Default' profile cannot be deleted.")
            return
        reply = QMessageBox.question(
            self,
            "Delete Profile",
            f"Are you sure you want to delete the profile '{profile_name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            index = self.parent_widget.profiles.index(profile_name)
            self.parent_widget.profiles.pop(index)
            self.parent_widget.settings['profiles'].pop(profile_name)
            self.parent_widget.save_settings()
            self.profile_list.takeItem(self.profile_list.row(selected_items[0]))
            logging.info(f"Profile '{profile_name}' deleted.")


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

        # Create a horizontal layout for the slider and its value label
        volume_layout = QHBoxLayout()

        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setMinimum(0)
        self.volume_slider.setMaximum(100)
        self.volume_slider.setValue(100)
        self.volume_slider.setTickInterval(10)
        self.volume_slider.setTickPosition(QSlider.TicksBelow)
        self.volume_slider.valueChanged.connect(self.update_volume_label)
        volume_layout.addWidget(self.volume_slider)

        self.volume_value_label = QLabel("100%")
        volume_layout.addWidget(self.volume_value_label)
        layout.addLayout(volume_layout)

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

        app_buttons_layout = QHBoxLayout()
        self.select_apps_btn = QPushButton("Select Applications")
        self.select_apps_btn.clicked.connect(self.select_applications)
        app_buttons_layout.addWidget(self.select_apps_btn)

        self.remove_app_btn = QPushButton("Remove Selected")
        self.remove_app_btn.clicked.connect(self.remove_app)
        app_buttons_layout.addWidget(self.remove_app_btn)

        layout.addLayout(app_buttons_layout)

        self.tray_checkbox = QCheckBox("Enable Minimize to Tray")
        layout.addWidget(self.tray_checkbox)

        self.desktop_shortcut_checkbox = QCheckBox("Create Desktop Shortcut")
        layout.addWidget(self.desktop_shortcut_checkbox)

        self.enable_logging_checkbox = QCheckBox("Enable Logging")
        layout.addWidget(self.enable_logging_checkbox)

        buttons_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_settings)
        buttons_layout.addWidget(self.save_btn)

        self.reset_btn = QPushButton("Reset to Default")
        self.reset_btn.clicked.connect(self.reset_settings)
        buttons_layout.addWidget(self.reset_btn)

        layout.addLayout(buttons_layout)

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
            logging.info(f"Switching profile from index {self.parent_widget.current_profile_index} to {index}.")
            if index < len(self.parent_widget.profiles):
                self.parent_widget.current_profile_index = index
                self.parent_widget.settings['current_profile'] = index
                self.parent_widget.load_current_profile()
                QMessageBox.information(self, "Profile Switched", f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                logging.info(f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                self.load_settings()
            else:
                logging.error(f"Invalid profile index: {index}. Reverting to the first profile.")
                QMessageBox.warning(self, "Profile Switch Failed", "Selected profile does not exist. Reverting to the default profile.")
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
        self.enable_logging_checkbox.setChecked(profile.get('enable_logging', True))

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

        profile['enable_logging'] = self.enable_logging_checkbox.isChecked()

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

        self.parent_widget.setup_logging(profile['enable_logging'])

        QMessageBox.information(self, "Settings Saved", "Settings have been saved successfully.")
        self.accept()

    def reset_settings(self):
        profile = self.parent_widget.get_current_profile()
        default = self.parent_widget.default_profile_settings()
        for key in default:
            profile[key] = default[key]

        self.load_settings()

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
            logging.error(f"Error in HotkeyListener thread: {e}")


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
            QMessageBox.critical(None, "Error", "Failed to open audio stream for visualization.")
            return

        while self.running:
            try:
                data = stream.read(self.chunk, exception_on_overflow=False)
                audio_data = np.frombuffer(data, dtype=np.int16)
                peak = np.abs(audio_data).max()
                level = int((peak / 32768) * 100)
                self.parent.update_audio_level_visualization(level)
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
        self.enable_logging = False
        self.app_check_timer = QTimer(self)
        self.app_check_timer.timeout.connect(self.check_auto_mute_apps)
        self.app_check_timer.start(5000)
        self.notifier = ToastNotifier()

        self.profiles = []
        self.current_profile_index = 0
        self.settings = {}

        self.original_mic_off_icon = QIcon(self.resource_path(os.path.join("images", "mic_off.png")))
        self.tinted_mic_off_icon = QIcon(self.tint_pixmap(os.path.join("images", "mic_off.png"), "red"))
        self.mic_on_icon = QIcon(self.resource_path(os.path.join("images", "mic_on.png")))

        icon_path = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
        self.setWindowIcon(QIcon(icon_path))

        self.initUI()
        self.load_settings()
        self.setup_logging(self.enable_logging)
        self.init_device()
        self.init_tray_icon()
        self.load_hotkey()

        self.toggle_mute_signal.connect(self.toggle_mute)

        self.check_for_updates()

        self.audio_thread = AudioStreamThread(self)
        self.audio_thread.start()

    def setup_logging(self, enable_logging: bool):
        setup_logging(enable_logging)

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
        self.mute_btn.setToolTip("Mute or unmute your microphone.")
        layout.addWidget(self.mute_btn)

        volume_layout = QHBoxLayout()
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setMinimum(0)
        self.volume_slider.setMaximum(100)
        self.volume_slider.setValue(100)
        self.volume_slider.setTickInterval(10)
        self.volume_slider.setTickPosition(QSlider.TicksBelow)
        self.volume_slider.valueChanged.connect(self.set_volume)
        volume_layout.addWidget(self.volume_slider)

        self.volume_label = QLabel("100%")
        volume_layout.addWidget(self.volume_label)
        layout.addLayout(volume_layout)

        self.hotkey_label = QLabel("Recorded Hotkey: None", self)
        layout.addWidget(self.hotkey_label)

        self.record_hotkey_btn = QPushButton("Record Hotkey", self)
        self.record_hotkey_btn.clicked.connect(self.start_recording)
        self.record_hotkey_btn.setToolTip("Start recording a hotkey combination.")
        layout.addWidget(self.record_hotkey_btn)

        self.stop_record_btn = QPushButton("Stop Recording", self)
        self.stop_record_btn.clicked.connect(self.stop_recording)
        self.stop_record_btn.setToolTip("Stop recording the hotkey combination.")
        layout.addWidget(self.stop_record_btn)

        self.settings_btn = QPushButton("Settings", self)
        self.settings_btn.clicked.connect(self.open_settings)
        self.settings_btn.setToolTip("Open settings to customize the app.")
        layout.addWidget(self.settings_btn)

        self.check_updates_btn = QPushButton("Check for Updates", self)
        self.check_updates_btn.clicked.connect(self.check_for_updates)
        self.check_updates_btn.setToolTip("Check for application updates.")
        layout.addWidget(self.check_updates_btn)

        self.update_status_label = QLabel("")
        layout.addWidget(self.update_status_label)

        self.help_btn = QPushButton("Help", self)
        self.help_btn.clicked.connect(self.show_help)
        self.help_btn.setToolTip("Show help information.")
        layout.addWidget(self.help_btn)

        self.audio_level_label = QLabel("Audio Level: 0%", self)
        layout.addWidget(self.audio_level_label)

        self.audio_level_visual = QProgressBar(self)
        self.audio_level_visual.setMinimum(0)
        self.audio_level_visual.setMaximum(100)
        self.audio_level_visual.setValue(0)
        layout.addWidget(self.audio_level_visual)

        self.setLayout(layout)
        self.setWindowTitle('MicMaster')
        self.resize(300, 500)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)

    def apply_theme(self, theme_name: str):
        if theme_name == 'Dark':
            self.setStyleSheet("""
                /* General Widget Styles */
                QWidget {
                    background-color: #1e1e1e;
                    color: #c5c6c7;
                    font-family: "Segoe UI", sans-serif;
                    font-size: 10pt;
                }

                /* QPushButton Styles */
                QPushButton {
                    background-color: #282a36;
                    border: 2px solid #44475a;
                    border-radius: 8px;
                    padding: 8px;
                    color: #f8f8f2;
                }
                QPushButton:hover {
                    background-color: #44475a;
                }
                QPushButton:pressed {
                    background-color: #6272a4;
                }

                /* QSlider Styles */
                QSlider::groove:horizontal {
                    height: 8px;
                    background: #44475a;
                    border-radius: 4px;
                }
                QSlider::handle:horizontal {
                    background: #50fa7b;
                    border: 1px solid #bd93f9;
                    width: 14px;
                    margin: -3px 0;
                    border-radius: 7px;
                }
                QSlider::handle:horizontal:hover {
                    background: #8be9fd;
                }

                /* QLabel Styles */
                QLabel {
                    color: #f8f8f2;
                }

                /* QComboBox Styles */
                QComboBox {
                    background-color: #282a36;
                    border: 1px solid #44475a;
                    border-radius: 8px;
                    padding: 4px;
                    color: #f8f8f2;
                }
                QComboBox QAbstractItemView {
                    background-color: #282a36;
                    selection-background-color: #44475a;
                    selection-color: #f8f8f2;
                }

                /* QCheckBox Styles */
                QCheckBox {
                    padding: 4px;
                }

                /* QProgressBar Styles */
                QProgressBar {
                    border: 2px solid #44475a;
                    border-radius: 8px;
                    text-align: center;
                    background-color: #282a36;
                }
                QProgressBar::chunk {
                    background-color: #50fa7b;
                    border-radius: 4px;
                }
            """)
            logging.info("Applied Modern Dark theme.")
        else:
            self.setStyleSheet("""
                /* General Widget Styles */
                QWidget {
                    background-color: #f5f5f5;
                    color: #2e2e2e;
                    font-family: "Segoe UI", sans-serif;
                    font-size: 10pt;
                }

                /* QPushButton Styles */
                QPushButton {
                    background-color: #ffffff;
                    border: 2px solid #c5c5c5;
                    border-radius: 8px;
                    padding: 8px;
                    color: #2e2e2e;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
                QPushButton:pressed {
                    background-color: #d5d5d5;
                }

                /* QSlider Styles */
                QSlider::groove:horizontal {
                    height: 8px;
                    background: #c5c5c5;
                    border-radius: 4px;
                }
                QSlider::handle:horizontal {
                    background: #0078d7;
                    border: 1px solid #005a9e;
                    width: 14px;
                    margin: -3px 0;
                    border-radius: 7px;
                }
                QSlider::handle:horizontal:hover {
                    background: #3399ff;
                }

                /* QLabel Styles */
                QLabel {
                    color: #2e2e2e;
                }

                /* QComboBox Styles */
                QComboBox {
                    background-color: #ffffff;
                    border: 1px solid #c5c5c5;
                    border-radius: 8px;
                    padding: 4px;
                    color: #2e2e2e;
                }
                QComboBox QAbstractItemView {
                    background-color: #ffffff;
                    selection-background-color: #0078d7;
                    selection-color: #ffffff;
                }

                /* QCheckBox Styles */
                QCheckBox {
                    padding: 4px;
                }

                /* QProgressBar Styles */
                QProgressBar {
                    border: 2px solid #c5c5c5;
                    border-radius: 8px;
                    text-align: center;
                    background-color: #ffffff;
                }
                QProgressBar::chunk {
                    background-color: #0078d7;
                    border-radius: 4px;
                }
            """)
            logging.info("Applied Modern Light theme.")

    def show_help(self):
        help_message = """
        MicMaster Help:
        - Mute/Unmute your microphone with the 'Mute Mic' button or assigned hotkey.
        - Set your preferences in the Settings menu.
        - Auto-Mute when specific applications are running.
        - View real-time audio level.
        - Minimize to the tray to keep MicMaster running in the background.
        - Check for updates to stay up-to-date with the latest features.
        - Manage multiple profiles for different settings configurations.
        """
        QMessageBox.information(self, "MicMaster Help", help_message)

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
            logging.info("Settings saved successfully.")
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
            'enable_logging': True,
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
            self.enable_logging = profile.get('enable_logging', True)
            self.enable_auto_mute = profile.get('enable_auto_mute', False)
            self.auto_mute_apps = profile.get('auto_mute_apps', [])
            self.hotkey = profile.get('hotkey', None)

            self.apply_theme(profile.get('theme', 'Dark'))
            logging.info(f"Applied profile: {self.profiles[self.current_profile_index]}")
        else:
            logging.error(f"Profile at index {self.current_profile_index} does not exist.")

    def add_to_startup(self):
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
            startup_folder = os.path.join(
                os.getenv('APPDATA'),
                'Microsoft',
                'Windows',
                'Start Menu',
                'Programs',
                'Startup'
            )
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
            startup_folder = os.path.join(
                os.getenv('APPDATA'),
                'Microsoft',
                'Windows',
                'Start Menu',
                'Programs',
                'Startup'
            )
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
                logging.error("No microphone device found.")
                QMessageBox.critical(self, "Error", "No microphone device found.")
                return

            self.device = devices
            self.interface = self.device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))

            current_volume = self.volume.GetMasterVolumeLevelScalar() * 100
            self.volume_slider.setValue(int(current_volume))
            self.volume_label.setText(f"{int(current_volume)}%")
        except Exception as e:
            logging.error(f"Error initializing microphone volume control: {e}")
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

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)

        if self.tray_enabled:
            self.hide()
            self.tray_icon.show()
            self.tray_icon.showMessage(
                "MicMaster",
                "App minimized to tray on startup",
                QSystemTrayIcon.Information,
                2000
            )
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
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized():
                profile = self.get_current_profile()
                if profile.get('tray_enabled', False):
                    QTimer.singleShot(0, self.hide)
                    self.tray_icon.showMessage(
                        "MicMaster",
                        "Application minimized to tray",
                        QSystemTrayIcon.Information,
                        2000
                    )
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
        self.hotkey_label.setText("Recording hotkey... Press 'Stop Recording' when finished.")
        keyboard.hook(self.record_key)

    def stop_recording(self):
        if not getattr(self, 'recording', False):
            self.hotkey_label.setText("No recording in progress.")
            return

        self.recording = False
        self.hotkey_label.setStyleSheet("color: white;")
        keyboard.unhook_all()

        if not self.pressed_keys:
            self.hotkey_label.setText("Error: No keys were recorded.")
            return

        self.hotkey = "+".join(sorted(self.pressed_keys))

        if self.current_hotkey:
            try:
                keyboard.remove_hotkey(self.current_hotkey)
                logging.info(f"Removed previous hotkey: {self.current_hotkey}")
            except ValueError:
                logging.warning(f"Hotkey {self.current_hotkey} not found when trying to remove.")

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
            logging.info("toggle_mute_signal emitted.")
        except Exception as e:
            logging.error(f"Error emitting toggle_mute_signal: {e}")

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
            logging.error(f"Error checking auto-mute applications: {e}")

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
                tray_icon_path = self.resource_path(os.path.join("images", "mic_off.png")) if self.is_muted else self.resource_path(os.path.join("images", "mic_on.png"))
                self.tray_icon.setIcon(QIcon(tray_icon_path))
        except Exception as e:
            logging.error(f"Error toggling mute state: {e}")
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
                self.notifier.show_toast(
                    "MicMaster",
                    f"Microphone {status}",
                    icon_path=self.resource_path(os.path.join("icons", "mic_switch_icon.ico")),
                    duration=5,
                    threaded=True,
                    callback_on_click=self.handle_toggle_mute_callback
                )
            except Exception as e:
                logging.error(f"Error showing interactive notification: {e}")
                notification.notify(
                    title="MicMaster",
                    message=f"Microphone {status}",
                    timeout=2
                )

    def handle_toggle_mute_callback(self):
        try:
            self.toggle_mute()
            logging.info("Microphone toggle triggered via notification click.")
        except Exception as e:
            logging.error(f"Error in notification callback: {e}")
        return 0

    def mute_microphone(self, mute: bool):
        try:
            if self.volume:
                self.volume.SetMute(1 if mute else 0, None)
                logging.info(f"Microphone {'muted' if mute else 'unmuted'}.")
        except Exception as e:
            logging.error(f"Error controlling the microphone: {e}")
            QMessageBox.critical(self, "Error", "Failed to control microphone.")

    def set_volume(self, value: int):
        if self.volume:
            volume_level = value / 100.0
            try:
                self.volume.SetMasterVolumeLevelScalar(volume_level, None)
                self.volume_label.setText(f"{value}%")
                logging.info(f"Volume set to {value}%.")
            except Exception as e:
                logging.error(f"Error setting microphone volume: {e}")
                QMessageBox.critical(self, "Error", "Failed to set volume.")

    def update_audio_level(self):
        try:
            if self.volume:
                current_level = self.volume.GetMasterVolumeLevelScalar() * 100
                self.audio_level_label.setText(f"Audio Level: {int(current_level)}%")
        except Exception as e:
            logging.error(f"Error updating audio level: {e}")

    def update_audio_level_visualization(self, level: int):
        self.audio_level_visual.setValue(level)
        self.audio_level_label.setText(f"Audio Level: {level}%")

    def check_for_updates(self):
        try:
            self.update_status_label.setText("Checking for updates...")
            self.update_status_label.repaint()
            self.check_updates_btn.setEnabled(False)

            owner = "balki97"
            repo = "MicMaster"
            api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
            response = requests.get(api_url, timeout=10)

            if response.status_code == 404:
                self.update_status_label.setText("No releases found. Please create a release on GitHub.")
                QMessageBox.information(
                    self,
                    "No Releases",
                    "No releases found for MicMaster. Please create a release on GitHub to enable update checking."
                )
                return

            response.raise_for_status()
            latest_release = response.json()
            assets = latest_release.get('assets', [])
            download_url = assets[0].get('browser_download_url', '') if assets else ''

            latest_version = latest_release.get('tag_name', '').lstrip('v')

            if not latest_version:
                logging.warning("Latest version not found in the GitHub response.")
                self.update_status_label.setText("Failed to retrieve latest version.")
                QMessageBox.warning(self, "Update Check", "Failed to retrieve the latest version information.")
                return

            if self.is_newer_version(latest_version, VERSION):
                reply = QMessageBox.question(
                    self,
                    "Update Available",
                    f"A new version ({latest_version}) is available. You are using version {VERSION}.\n\nDo you want to download the latest version?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes and download_url:
                    self.download_update(download_url)
                self.update_status_label.setText("Update available.")
            else:
                self.update_status_label.setText("You are using the latest version.")
                logging.info("You are using the latest version.")
        except requests.RequestException as e:
            logging.error(f"Error checking for updates: {e}")
            self.update_status_label.setText("Error checking for updates.")
            QMessageBox.warning(self, "Update Check Failed", "Failed to check for updates.")
        except Exception as e:
            logging.error(f"Unexpected error during update check: {e}")
            self.update_status_label.setText("Error checking for updates.")
            QMessageBox.warning(self, "Update Check Failed", "Failed to check for updates.")
        finally:
            self.check_updates_btn.setEnabled(True)

    def download_update(self, url: str):
        try:
            self.update_status_label.setText("Downloading update...")
            self.update_status_label.repaint()
            response = requests.get(url, stream=True)
            total_length = response.headers.get('content-length')

            update_filename = "MicMaster_update.exe"

            if os.path.exists(update_filename):
                os.remove(update_filename)

            if total_length is None:
                with open(update_filename, 'wb') as f:
                    f.write(response.content)
            else:
                dl = 0
                total_length = int(total_length)
                with open(update_filename, 'wb') as f:
                    for data in response.iter_content(chunk_size=4096):
                        dl += len(data)
                        f.write(data)
                        percent = int(dl * 100 / total_length)
                        self.update_status_label.setText(f"Downloading update... {percent}%")
                        self.update_status_label.repaint()

            QMessageBox.information(
                self,
                "Download Complete",
                "Update downloaded successfully. The application will now update and restart."
            )

            new_exe = os.path.join(os.getcwd(), "MicMaster_update.exe")
            current_exe = sys.executable

            subprocess.Popen([current_exe, "--update", new_exe])

            QApplication.quit()
        except Exception as e:
            logging.error(f"Error downloading update: {e}")
            QMessageBox.critical(self, "Download Failed", "Failed to download the update.")

    def is_newer_version(self, latest: str, current: str) -> bool:
        def version_tuple(v):
            return tuple(map(int, (v.split("."))))
        try:
            return version_tuple(latest) > version_tuple(current)
        except ValueError:
            logging.error(f"Invalid version format. Latest: {latest}, Current: {current}")
            return False

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
                    self.notifier.show_toast(
                        "MicMaster",
                        "Desktop shortcut created successfully.",
                        icon_path=self.resource_path(os.path.join("icons", "mic_switch_icon.ico")),
                        duration=2,
                        threaded=True
                    )
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
            self.setup_logging(self.enable_logging)
            self.setup_auto_mute()
            self.load_hotkey()
        except Exception as e:
            logging.error(f"Error loading current profile: {e}")
            QMessageBox.critical(self, "Error", "Failed to load the selected profile.")
            self.current_profile_index = 0
            self.settings['current_profile'] = 0
            self.apply_profile_settings()

    def get_profiles(self) -> list:
        return self.profiles

    def setup_auto_mute(self):
        profile = self.get_current_profile()
        self.enable_auto_mute = profile.get('enable_auto_mute', False)
        self.auto_mute_apps = profile.get('auto_mute_apps', [])

    def toggle_voice_activation(self, state):
        pass  # Placeholder for future implementation


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
    if len(sys.argv) == 3 and sys.argv[1] == "--update":
        new_exe_path = sys.argv[2]
        current_exe_path = sys.executable
        perform_update(new_exe_path, current_exe_path)
        sys.exit(0)
    app_instance = MicMasterApp()
    app_instance.run()


if __name__ == '__main__':
    main()
