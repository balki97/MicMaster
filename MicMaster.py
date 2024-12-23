import pythoncom
import logging
import psutil
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QSlider, QHBoxLayout,
    QCheckBox, QSystemTrayIcon, QMenu, QAction, QDialog, QSpinBox, QComboBox,
    QMessageBox, QListWidget, QListWidgetItem, QDialogButtonBox, QAbstractItemView,
    QInputDialog
)
from PyQt5.QtCore import Qt, QEvent, QTimer, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from plyer import notification
import keyboard
import sys
import os
import json
import winsound
import requests
import shutil
import subprocess
import time
from threading import Thread
from win10toast_click import ToastNotifier

# Define the current version
VERSION = "1.0.4"

SETTINGS_FILE = 'settings.json'
LOG_FILE = 'app.log'

def setup_logging(enable_logging):
    """Configure logging based on user preference."""
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Remove all existing handlers
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
    else:
        # Logging is disabled; no handlers will be added
        pass

def is_process_running(exe_name):
    """Check if a process with the given executable name is running."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == exe_name:
            return True
    return False

def perform_update(new_exe_path, current_exe_path):
    """Replace the current executable with the new one and restart the application."""
    setup_logging(True)  # Ensure logging is enabled for the updater
    logging.info("Updater started.")
    logging.info(f"New executable path: {new_exe_path}")
    logging.info(f"Current executable path: {current_exe_path}")
    
    try:
        # Wait for the main application to close
        logging.info("Waiting for the main application to exit...")
        for _ in range(30):
            if not is_process_running(os.path.basename(current_exe_path)):
                break
            time.sleep(1)
        else:
            logging.error("Main application did not exit within the expected time.")
            QMessageBox.critical(None, "Update Error", "Failed to update the application because it is still running.")
            sys.exit(1)
        
        # Replace the executable
        shutil.move(new_exe_path, current_exe_path)
        logging.info("Executable replaced successfully.")
        
        # Restart the application
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

        # Populate the list with current running applications
        running_apps = sorted(set(proc.name() for proc in psutil.process_iter()))
        for app in running_apps:
            item = QListWidgetItem(app)
            self.process_list.addItem(item)

        # OK and Cancel buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def get_selected_apps(self):
        selected_items = self.process_list.selectedItems()
        return [item.text() for item in selected_items]

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

        # OK and Cancel buttons
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
            self.profile_list.addItem(profile_name)
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

        # Profile Management
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

        # Default Volume
        self.volume_label = QLabel("Default Volume:")
        layout.addWidget(self.volume_label)

        self.volume_spinbox = QSpinBox()
        self.volume_spinbox.setMinimum(0)
        self.volume_spinbox.setMaximum(100)
        layout.addWidget(self.volume_spinbox)

        # Startup Option
        self.startup_checkbox = QCheckBox("Start on system boot")
        layout.addWidget(self.startup_checkbox)

        # Notification toggle
        self.notifications_checkbox = QCheckBox("Enable desktop notifications")
        layout.addWidget(self.notifications_checkbox)

        # Sound Notification Option
        self.sound_notification_checkbox = QCheckBox("Enable sound notifications instead of desktop notifications")
        layout.addWidget(self.sound_notification_checkbox)

        # Theme Selection
        self.theme_label = QLabel("Theme:")
        layout.addWidget(self.theme_label)

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        layout.addWidget(self.theme_combo)

        # Enable Auto-Mute on Applications
        self.enable_auto_mute_checkbox = QCheckBox("Enable Auto-Mute on Specific Applications")
        layout.addWidget(self.enable_auto_mute_checkbox)

        # Applications to Auto-Mute
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

        # Tray Option
        self.tray_checkbox = QCheckBox("Enable Minimize to Tray")
        layout.addWidget(self.tray_checkbox)

        # Create Desktop Shortcut Option
        self.desktop_shortcut_checkbox = QCheckBox("Create Desktop Shortcut")
        layout.addWidget(self.desktop_shortcut_checkbox)

        # Enable Logging Option
        self.enable_logging_checkbox = QCheckBox("Enable Logging")
        layout.addWidget(self.enable_logging_checkbox)

        # Save and Reset buttons
        buttons_layout = QHBoxLayout()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_settings)
        buttons_layout.addWidget(self.save_btn)

        self.reset_btn = QPushButton("Reset to Default")
        self.reset_btn.clicked.connect(self.reset_settings)
        buttons_layout.addWidget(self.reset_btn)

        layout.addLayout(buttons_layout)

        self.setLayout(layout)

        # Load initial settings
        self.load_settings()

        # Connect checkboxes to enable/disable related widgets
        self.enable_auto_mute_checkbox.stateChanged.connect(self.toggle_auto_mute)

    def manage_profiles(self):
        dialog = ProfileManagementDialog(self.parent_widget)
        if dialog.exec_():
            self.profile_combo.clear()
            self.profile_combo.addItems(self.parent_widget.profiles)
            self.profile_combo.setCurrentIndex(self.parent_widget.current_profile_index)

    def switch_profile(self, index):
        if index != self.parent_widget.current_profile_index:
            logging.info(f"Switching profile from index {self.parent_widget.current_profile_index} to {index}.")
            if index < len(self.parent_widget.profiles):
                self.parent_widget.current_profile_index = index
                self.parent_widget.settings['current_profile'] = index
                self.parent_widget.load_current_profile()
                QMessageBox.information(self, "Profile Switched", f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                logging.info(f"Switched to profile '{self.parent_widget.profiles[index]}'.")
                self.load_settings()  # Refresh the settings fields with the new profile
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

    def is_app_in_list(self, app_name):
        for i in range(self.app_list.count()):
            if self.app_list.item(i).text().lower() == app_name.lower():
                return True
        return False

    def remove_app(self):
        selected_items = self.app_list.selectedItems()
        if not selected_items:
            return
        for item in selected_items:
            self.app_list.takeItem(self.app_list.row(item))

    def toggle_auto_mute(self, state):
        enabled = (state == Qt.Checked)
        # Allow the app list and remove button to always be enabled
        self.app_list.setEnabled(True)
        self.remove_app_btn.setEnabled(True)
        # Enable or disable only the select button based on the checkbox
        self.select_apps_btn.setEnabled(enabled)
        self.auto_mute_label.setEnabled(enabled)

    def load_settings(self):
        profile = self.parent_widget.get_current_profile()
        self.volume_spinbox.setValue(profile.get('volume', 100))
        self.startup_checkbox.setChecked(profile.get('startup', False))
        self.notifications_checkbox.setChecked(profile.get('notifications', False))
        self.sound_notification_checkbox.setChecked(profile.get('sound_notifications', False))
        self.theme_combo.setCurrentText(profile.get('theme', 'Dark'))

        self.enable_auto_mute_checkbox.setChecked(profile.get('enable_auto_mute', False))
        auto_mute_apps = profile.get('auto_mute_apps', [])
        self.app_list.clear()
        self.app_list.addItems(auto_mute_apps)

        self.tray_checkbox.setChecked(profile.get('tray_enabled', False))

        self.desktop_shortcut_checkbox.setChecked(profile.get('create_desktop_shortcut', False))

        self.enable_logging_checkbox.setChecked(profile.get('enable_logging', True))  # Default True

        # Ensure widgets are enabled/disabled based on loaded settings
        self.toggle_auto_mute(self.enable_auto_mute_checkbox.isChecked())

    def save_settings(self):
        profile = self.parent_widget.get_current_profile()
        profile['volume'] = self.volume_spinbox.value()
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

        # Handle Desktop Shortcut
        if profile['create_desktop_shortcut']:
            self.parent_widget.create_desktop_shortcut_method()
        else:
            self.parent_widget.remove_desktop_shortcut_method()

        if profile['startup']:
            self.parent_widget.add_to_startup()
        else:
            self.parent_widget.remove_from_startup()

        # Apply Logging Settings
        self.parent_widget.setup_logging(profile['enable_logging'])

        QMessageBox.information(self, "Settings Saved", "Settings have been saved successfully.")
        self.accept()

    def reset_settings(self):
        profile = self.parent_widget.get_current_profile()
        profile['volume'] = 100
        profile['startup'] = False
        profile['notifications'] = False
        profile['sound_notifications'] = False
        profile['theme'] = 'Dark'

        profile['enable_auto_mute'] = False
        profile['auto_mute_apps'] = []
        self.app_list.clear()

        profile['tray_enabled'] = False
        profile['create_desktop_shortcut'] = False

        profile['enable_logging'] = True

        self.volume_spinbox.setValue(100)
        self.startup_checkbox.setChecked(False)
        self.notifications_checkbox.setChecked(False)
        self.sound_notification_checkbox.setChecked(False)
        self.theme_combo.setCurrentText("Dark")

        self.enable_auto_mute_checkbox.setChecked(False)
        self.app_list.clear()
        self.app_list.setEnabled(True)
        self.select_apps_btn.setEnabled(False)
        self.remove_app_btn.setEnabled(True)
        self.auto_mute_label.setEnabled(False)

        self.tray_checkbox.setChecked(False)
        self.desktop_shortcut_checkbox.setChecked(False)

        self.enable_logging_checkbox.setChecked(True)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

class HotkeyListener(Thread):
    def __init__(self, callback, hotkey):
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

class MicMaster(QWidget):
    toggle_mute_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.is_muted = False
        self.device = None
        self.interface = None
        self.volume = None
        self.hotkey = None
        self.current_hotkey = None
        self.tray_icon = None
        self.tray_enabled = False
        self.use_sound_notification = False
        self.notifications_enabled = False
        self.pressed_keys = set()
        self.auto_mute_apps = []
        self.enable_auto_mute = False
        self.enable_logging = False
        self.app_check_timer = QTimer(self)
        self.app_check_timer.timeout.connect(self.check_auto_mute_apps)
        self.app_check_timer.start(5000)  # Check every 5 seconds
        self.notifier = ToastNotifier()

        # Profiles
        self.profiles = []
        self.current_profile_index = 0
        self.settings = {}
        
        # Preload and tint the mic_off.png icon
        self.original_mic_off_icon = QIcon(self.resource_path(os.path.join("images", "mic_off.png")))
        self.tinted_mic_off_icon = QIcon(self.tint_pixmap(os.path.join("images", "mic_off.png"), "red"))
        self.mic_on_icon = QIcon(self.resource_path(os.path.join("images", "mic_on.png")))

        # Set the application window icon
        icon_path = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
        self.setWindowIcon(QIcon(icon_path))

        # Initialize UI
        self.initUI()

        # Load settings
        self.load_settings()

        # Setup Logging Based on Settings
        self.setup_logging(self.enable_logging)

        self.init_device()
        self.init_tray_icon()
        self.load_hotkey()

        # Connect the signal to ensure toggle_mute runs in the main thread
        self.toggle_mute_signal.connect(self.toggle_mute)

        # Check for updates on launch
        self.check_for_updates()

    def setup_logging(self, enable_logging):
        """Wrapper to configure logging."""
        setup_logging(enable_logging)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def tint_pixmap(self, pixmap_path, color):
        """Tint the given pixmap with the specified color."""
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

        self.tray_checkbox = QCheckBox("Minimize to system tray", self)
        self.tray_checkbox.setChecked(self.tray_enabled)
        self.tray_checkbox.stateChanged.connect(self.toggle_tray_option)
        self.tray_checkbox.setToolTip("Minimize the app to the system tray.")
        layout.addWidget(self.tray_checkbox)

        self.settings_btn = QPushButton("Settings", self)
        self.settings_btn.clicked.connect(self.open_settings)
        self.settings_btn.setToolTip("Open settings to customize the app.")
        layout.addWidget(self.settings_btn)

        self.check_updates_btn = QPushButton("Check for Updates", self)
        self.check_updates_btn.clicked.connect(self.check_for_updates)
        self.check_updates_btn.setToolTip("Check for application updates.")
        layout.addWidget(self.check_updates_btn)

        # Update Status Label
        self.update_status_label = QLabel("")
        layout.addWidget(self.update_status_label)

        self.help_btn = QPushButton("Help", self)
        self.help_btn.clicked.connect(self.show_help)
        self.help_btn.setToolTip("Show help information.")
        layout.addWidget(self.help_btn)

        # Real-Time Audio Level Indicator
        self.audio_level_label = QLabel("Audio Level: 0%", self)
        layout.addWidget(self.audio_level_label)
        self.audio_level_timer = QTimer(self)
        self.audio_level_timer.timeout.connect(self.update_audio_level)
        self.audio_level_timer.start(1000)  # Update every second

        self.setLayout(layout)
        self.setWindowTitle('MicMaster')
        self.resize(300, 400)  # Allow resizing

        # **Remove the Maximize Button**
        # This line removes the maximize button while keeping the window resizable
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)

    def apply_theme(self, theme_name):
        """Apply the selected theme to the application."""
        if theme_name == 'Dark':
            self.setStyleSheet("""
                QWidget {
                    background-color: #2e2e2e;
                    color: #ffffff;
                }
                QPushButton {
                    background-color: #444444;
                    border: none;
                    padding: 10px;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
                QSlider::groove:horizontal {
                    height: 8px;
                    background: #444444;
                    border-radius: 4px;
                }
                QSlider::handle:horizontal {
                    background: #ffffff;
                    border: 1px solid #5c5c5c;
                    width: 14px;
                    margin: -4px 0;
                    border-radius: 7px;
                }
            """)
            logging.info("Applied Dark theme.")
        else:
            self.setStyleSheet("""
                QWidget {
                    background-color: #f0f0f0;
                    color: #000000;
                }
                QPushButton {
                    background-color: #dddddd;
                    border: none;
                    padding: 10px;
                }
                QPushButton:hover {
                    background-color: #cccccc;
                }
                QSlider::groove:horizontal {
                    height: 8px;
                    background: #cccccc;
                    border-radius: 4px;
                }
                QSlider::handle:horizontal {
                    background: #000000;
                    border: 1px solid #5c5c5c;
                    width: 14px;
                    margin: -4px 0;
                    border-radius: 7px;
                }
            """)
            logging.info("Applied Light theme.")

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
            self.load_settings()  # Reload settings after changes
            self.setup_auto_mute()

    def load_settings(self):
        """Load settings from the settings file."""
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

    def default_profile_settings(self):
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

    def get_current_profile(self):
        profile_name = self.profiles[self.current_profile_index]
        return self.settings['profiles'][profile_name]

    def get_profiles(self):
        return self.profiles

    def apply_profile_settings(self):
        profile = self.get_current_profile()
        if profile:
            # Update MicMaster's own UI elements and attributes
            self.volume_slider.setValue(profile.get('volume', 100))
            self.use_sound_notification = profile.get('sound_notifications', False)
            self.notifications_enabled = profile.get('notifications', False)
            self.tray_enabled = profile.get('tray_enabled', False)
            self.enable_logging = profile.get('enable_logging', True)
            self.enable_auto_mute = profile.get('enable_auto_mute', False)
            self.auto_mute_apps = profile.get('auto_mute_apps', [])
            self.hotkey = profile.get('hotkey', None)

            # Apply theme
            self.apply_theme(profile.get('theme', 'Dark'))
            logging.info(f"Applied profile: {self.profiles[self.current_profile_index]}")
        else:
            logging.error(f"Profile at index {self.current_profile_index} does not exist.")

    def add_to_startup(self):
        try:
            # Path to the executable
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(sys.argv[0])

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
            return

    def init_device(self):
        """Initialize the audio device and prepare for volume control."""
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

            # Set initial slider to match system volume
            current_volume = self.volume.GetMasterVolumeLevelScalar() * 100
            self.volume_slider.setValue(int(current_volume))
            self.volume_label.setText(f"{int(current_volume)}%")
        except Exception as e:
            logging.error(f"Error initializing microphone volume control: {e}")
            QMessageBox.critical(self, "Error", "Failed to initialize microphone control.")

    def init_tray_icon(self):
        """Initialize the system tray icon."""
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

    def on_tray_icon_activated(self, reason):
        """Handle system tray icon activation."""
        if reason == QSystemTrayIcon.Trigger:
            self.show_normal()

    def show_normal(self):
        self.show()
        self.activateWindow()

    def toggle_tray_option(self, state):
        """Enable or disable the system tray option based on checkbox state."""
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
        """Handle window state changes like minimizing."""
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized() and self.tray_enabled:
                QTimer.singleShot(0, self.hide)
                self.tray_icon.showMessage(
                    "MicMaster",
                    "Application minimized to tray",
                    QSystemTrayIcon.Information,
                    2000
                )
        super().changeEvent(event)

    def closeEvent(self, event):
        """Handle the close event to exit the application."""
        event.accept()

    def quit_app(self):
        """Exit the application."""
        self.tray_icon.hide()
        QApplication.quit()

    def start_recording(self):
        """Start recording the hotkey from user input."""
        if hasattr(self, 'recording') and self.recording:
            return  # Already recording
        self.recording = True
        self.pressed_keys.clear()
        self.hotkey_label.setStyleSheet("color: red;")
        self.hotkey_label.setText("Recording hotkey... Press 'Stop Recording' when finished.")
        keyboard.hook(self.record_key)

    def stop_recording(self):
        """Stop recording the hotkey and set up the hotkey listener."""
        if not hasattr(self, 'recording') or not self.recording:
            self.hotkey_label.setText("No recording in progress.")
            return

        self.recording = False
        self.hotkey_label.setStyleSheet("color: white;")
        keyboard.unhook_all()

        if not self.pressed_keys:
            self.hotkey_label.setText("Error: No keys were recorded.")
            return

        self.hotkey = "+".join(sorted(self.pressed_keys))

        # Remove existing hotkey if any
        if self.current_hotkey:
            try:
                keyboard.remove_hotkey(self.current_hotkey)
                logging.info(f"Removed previous hotkey: {self.current_hotkey}")
            except ValueError:
                logging.warning(f"Hotkey {self.current_hotkey} not found when trying to remove.")

        try:
            # Instead of directly calling toggle_mute, emit the signal
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
        """Emit the toggle_mute_signal to safely toggle mute in the main thread."""
        try:
            self.toggle_mute_signal.emit()
            logging.info("toggle_mute_signal emitted.")
        except Exception as e:
            logging.error(f"Error emitting toggle_mute_signal: {e}")
        # No return statement

    def record_key(self, e):
        """Record each key press one at a time."""
        if e.event_type == "down" and e.name not in self.pressed_keys:
            self.pressed_keys.add(e.name)
            self.hotkey_label.setText(f"Recording hotkey: {' + '.join(sorted(self.pressed_keys))}")

    def load_hotkey(self):
        """Load and set the hotkey from settings."""
        profile = self.get_current_profile()
        hotkey = profile.get('hotkey', None)
        if hotkey:
            self.hotkey = hotkey
            try:
                # Start a hotkey listener thread
                self.hotkey_listener = HotkeyListener(self.emit_toggle_mute_signal, self.hotkey)
                self.hotkey_listener.start()
                self.hotkey_label.setText(f"Recorded Hotkey: {self.hotkey}")
                logging.info(f"Hotkey loaded: {self.hotkey}")
            except Exception as e:
                logging.error(f"Error loading hotkey: {e}")
                self.hotkey_label.setText("Error: Invalid hotkey.")

    def setup_auto_mute(self):
        """Set up auto-mute based on settings."""
        profile = self.get_current_profile()
        self.enable_auto_mute = profile.get('enable_auto_mute', False)
        self.auto_mute_apps = profile.get('auto_mute_apps', [])

    def check_auto_mute_apps(self):
        """Check if any of the auto-mute applications are running."""
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
        """Toggle the mute/unmute state of the microphone."""
        try:
            pythoncom.CoInitialize()
            self.is_muted = not self.is_muted
            self.mute_btn.setText("Unmute Mic" if self.is_muted else "Mute Mic")
            # Update mute button icon
            if self.is_muted:
                self.mute_btn.setIcon(self.tinted_mic_off_icon)
            else:
                self.mute_btn.setIcon(self.mic_on_icon)
            self.mute_microphone(self.is_muted)
            self.send_notification()
            logging.info(f"Microphone {'muted' if self.is_muted else 'unmuted'}.")
    
            # Update tray icon if initialized
            if self.tray_icon:
                if self.is_muted:
                    tray_icon_path = self.resource_path(os.path.join("images", "mic_off.png"))
                else:
                    tray_icon_path = self.resource_path(os.path.join("images", "mic_on.png"))
                self.tray_icon.setIcon(QIcon(tray_icon_path))

        except Exception as e:
            logging.error(f"Error toggling mute state: {e}")
            QMessageBox.critical(self, "Error", "Failed to toggle microphone.")

    def send_notification(self):
        """Send a desktop notification with action buttons for mic mute/unmute."""
        status = "Muted" if self.is_muted else "Unmuted"

        if self.notifications_enabled:
            if self.use_sound_notification:
                # Using winsound with .wav files
                sound_file = os.path.join("sounds", "mute_sound.wav") if self.is_muted else os.path.join("sounds", "unmute_sound.wav")
                try:
                    winsound.PlaySound(self.resource_path(sound_file), winsound.SND_FILENAME | winsound.SND_ASYNC)
                except Exception as e:
                    logging.error(f"Error playing sound: {e}")
                    # Fallback to system beep
                    if self.is_muted:
                        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                    else:
                        winsound.MessageBeep(winsound.MB_OK)
            else:
                # Send interactive notification
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
                    # Fallback to non-interactive notification
                    notification.notify(
                        title="MicMaster",
                        message=f"Microphone {status}",
                        timeout=2
                    )

    def handle_toggle_mute_callback(self):
        """Handle the callback from the notification click."""
        try:
            self.toggle_mute()
            logging.info("Microphone toggle triggered via notification click.")
        except Exception as e:
            logging.error(f"Error in notification callback: {e}")
        # Do not return anything

    def mute_microphone(self, mute):
        """Mute or unmute the microphone using pycaw."""
        try:
            if self.volume:
                self.volume.SetMute(1 if mute else 0, None)
                logging.info(f"Microphone {'muted' if mute else 'unmuted'}.")
        except Exception as e:
            logging.error(f"Error controlling the microphone: {e}")
            QMessageBox.critical(self, "Error", "Failed to control microphone.")

    def set_volume(self, value):
        """Set the microphone volume."""
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
        """Update the real-time audio level indicator."""
        try:
            if self.volume:
                current_level = self.volume.GetMasterVolumeLevelScalar() * 100
                self.audio_level_label.setText(f"Audio Level: {int(current_level)}%")
        except Exception as e:
            logging.error(f"Error updating audio level: {e}")

    def check_for_updates(self):
        """Check GitHub for the latest release and notify the user if an update is available."""
        try:
            self.update_status_label.setText("Checking for updates...")
            self.update_status_label.repaint()
            self.check_updates_btn.setEnabled(False)

            owner = "balki97" 
            repo = "MicMaster"
            api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
            response = requests.get(api_url, timeout=10)

            if response.status_code == 404:
                # Repository or releases not found
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
            if assets:
                download_url = assets[0].get('browser_download_url', '')
            else:
                download_url = ''

            latest_version = latest_release.get('tag_name', '').lstrip('v')

            if not latest_version:
                logging.warning("Latest version not found in the GitHub response.")
                self.update_status_label.setText("Failed to retrieve latest version.")
                QMessageBox.warning(self, "Update Check", "Failed to retrieve the latest version information.")
                return

            # Compare versions
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

    def download_update(self, url):
        """Download the latest update executable and initiate the update process."""
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
                        done = int(50 * dl / total_length)
                        # Optionally, update a progress bar here

            QMessageBox.information(
                self,
                "Download Complete",
                "Update downloaded successfully. The application will now update and restart."
            )

            # Path to the new executable
            new_exe = os.path.join(os.getcwd(), "MicMaster_update.exe")
            current_exe = sys.executable

            # Launch the updater process
            subprocess.Popen([current_exe, "--update", new_exe])

            # Exit the main application to allow the updater to replace the executable
            QApplication.quit()
        except Exception as e:
            logging.error(f"Error downloading update: {e}")
            QMessageBox.critical(self, "Download Failed", "Failed to download the update.")

    def is_newer_version(self, latest, current):
        """Compare two version strings."""
        def version_tuple(v):
            return tuple(map(int, (v.split("."))))
        try:
            return version_tuple(latest) > version_tuple(current)
        except ValueError:
            logging.error(f"Invalid version format. Latest: {latest}, Current: {current}")
            return False

    def create_desktop_shortcut_method(self):
        """Create a desktop shortcut for the application."""
        try:
            # Path to the executable
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(sys.argv[0])

            desktop_folder = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop_folder, 'MicMaster.lnk')

            if not os.path.exists(shortcut_path):
                from win32com.client import Dispatch
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.Arguments = ""
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                # Access resource_path from parent
                icon_location = self.resource_path(os.path.join("icons", "mic_switch_icon.ico"))
                if os.path.exists(icon_location):
                    shortcut.IconLocation = icon_location
                shortcut.save()
                logging.info("Desktop shortcut created.")

                # Notify success
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
        """Remove the desktop shortcut for the application."""
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
            return

    def load_current_profile(self):
        """Load the settings for the current profile."""
        try:
            self.apply_profile_settings()
            self.setup_logging(self.enable_logging)
            self.setup_auto_mute()
            self.load_hotkey()
        except Exception as e:
            logging.error(f"Error loading current profile: {e}")
            QMessageBox.critical(self, "Error", "Failed to load the selected profile.")
            # Optionally, revert to the default profile
            self.current_profile_index = 0
            self.settings['current_profile'] = 0
            self.apply_profile_settings()

    def setup_auto_mute(self):
        """Set up auto-mute based on settings."""
        profile = self.get_current_profile()
        self.enable_auto_mute = profile.get('enable_auto_mute', False)
        self.auto_mute_apps = profile.get('auto_mute_apps', [])

    def toggle_voice_activation(self, state):
        """Placeholder for voice activation toggle."""
        pass  # Not implemented in this version

class MicMasterApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        # Ensure the icons directory exists
        icons_dir = os.path.join(os.path.dirname(__file__), "icons")
        if not os.path.exists(icons_dir):
            setup_logging(True)
            logging.error("Icons directory not found.")
            QMessageBox.critical(None, "Error", "Icons directory not found.")
            sys.exit(1)

        icon_path = os.path.join(os.path.dirname(__file__), "icons", "mic_switch_icon.ico")
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
    # Check if the script is launched with the update argument
    if len(sys.argv) == 3 and sys.argv[1] == "--update":
        new_exe_path = sys.argv[2]
        current_exe_path = sys.executable
        perform_update(new_exe_path, current_exe_path)
        sys.exit(0)
    app_instance = MicMasterApp()
    app_instance.run()

if __name__ == '__main__':
    main()