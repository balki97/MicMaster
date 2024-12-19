# MicMaster

![MicMaster Logo](icons/mic_switch_icon.ico)

**MicMaster** is a powerful and user-friendly microphone management tool designed exclusively for Windows users. Whether you're a gamer, content creator, or professional, MicMaster provides seamless control over your microphone settings, ensuring optimal audio performance tailored to your needs.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Settings](#settings)
- [Hotkey Configuration](#hotkey-configuration)
- [Auto-Mute Applications](#auto-mute-applications)
- [System Tray Integration](#system-tray-integration)
- [Checking for Updates](#checking-for-updates)
- [Contributing](#contributing)
- [License](#license)

## Features

- **Mute/Unmute Microphone:** Instantly control your microphone's mute status with a single click or a customizable hotkey.
- **Volume Control:** Adjust your microphone's volume levels effortlessly using the integrated slider.
- **Custom Hotkeys:** Assign personalized hotkey combinations to toggle mute/unmute functionality for quick access.
- **Auto-Mute Specific Applications:** Automatically mute your microphone when certain applications are running, enhancing privacy and reducing interruptions.
- **Real-Time Audio Level Indicator:** Monitor your microphone's audio levels in real-time to ensure optimal sound quality.
- **System Tray Integration:** Minimize MicMaster to the system tray, allowing it to run unobtrusively in the background.
- **Desktop Shortcut Management:** Easily create or remove a desktop shortcut for quick access to MicMaster.
- **Theme Support:** Choose between Dark and Light themes to match your aesthetic preferences.
- **Startup Configuration:** Configure MicMaster to launch automatically when your system boots up.
- **Notifications:** Receive desktop notifications or sound alerts when muting/unmuting your microphone.
- **Update Checker:** Stay informed about the latest MicMaster releases directly from GitHub.

## Requirements

- **Operating System:** Windows 7 or later (64-bit recommended)
- **Python:** Python 3.7 or higher *(Only required if building from source)*
- **Dependencies:** All dependencies are bundled with the packaged executable. If building from source, ensure the following Python packages are installed:
  - PyQt5
  - pycaw
  - psutil
  - keyboard
  - requests
  - plyer
  - comtypes
  - win32com

## Installation

### Using the Packaged Executable

1. **Download the Latest Release:**
   - Navigate to the [Releases](https://github.com/yourusername/MicMaster/releases) section of this repository.
   - Download the latest `MicMaster.exe` file.

2. **Run the Installer:**
   - Double-click the downloaded `MicMaster.exe` file.
   - Follow the on-screen instructions to install MicMaster on your system.

3. **Launch MicMaster:**
   - After installation, MicMaster will be available in your Start Menu.
   - Open the application to begin managing your microphone settings.

### Building from Source *(Optional)*

If you prefer to build MicMaster from source, follow these steps:

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/yourusername/MicMaster.git
   cd MicMaster
   ```

2. **Create a Virtual Environment:**
   ```bash
    python -m venv venv
    venv\Scripts\activate
   ```

3. **Install Dependencies:**
   ```bash
    pip install -r requirements.txt
   ```

4. **Run the Application:**
   ```bash
    python MicMaster.py
   ```