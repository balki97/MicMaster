# MicMaster

![MicMaster Logo](icons/mic_switch_icon.ico)

**MicMaster** is a powerful and user-friendly microphone management tool designed exclusively for Windows users. Whether you're a gamer, content creator, or professional, MicMaster provides seamless control over your microphone settings, ensuring optimal audio performance tailored to your needs.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Download and Run](#download-and-run)
- [Usage](#usage)
- [Settings](#settings)
- [Hotkey Configuration](#hotkey-configuration)
- [Auto-Mute Applications](#auto-mute-applications)
- [System Tray Integration](#system-tray-integration)
- [Checking for Updates](#checking-for-updates)
- [Contributing](#contributing)

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

## Download and Run

**MicMaster** is distributed as a standalone executable for Windows. To get started:

1. **Download the Latest Release:**
   - Navigate to the [Releases](https://github.com/balki97/MicMaster/releases) section of this repository.
   - Download the latest `MicMaster.exe` file.

2. **Run the Application:**
   - Locate the downloaded `MicMaster.exe` file in your Downloads folder or the location you chose.
   - Double-click `MicMaster.exe` to launch the application.
   - No installation is required; MicMaster runs directly from the executable.

## Usage

Upon launching MicMaster, you'll be greeted with a straightforward interface:

- **Mute/Unmute Button:** Click to toggle your microphone's mute status.
- **Volume Slider:** Adjust the microphone volume to your preferred level.
- **Record Hotkey:** Assign a custom hotkey combination for quick mute/unmute actions.
- **Stop Recording:** Finish recording your hotkey combination.
- **Settings:** Access detailed configuration options to tailor MicMaster to your preferences.
- **Check for Updates:** Ensure you're running the latest version of MicMaster.
- **Help:** Access help documentation and usage tips.

## Settings

Customize MicMaster to fit your workflow by accessing the **Settings** window:

- **Default Volume:** Set your preferred microphone volume level.
- **Start on System Boot:** Enable MicMaster to launch automatically when Windows starts.
- **Desktop Notifications:** Toggle desktop notifications for mute/unmute actions.
- **Sound Notifications:** Choose sound alerts instead of desktop notifications.
- **Theme Selection:** Switch between Dark and Light themes.
- **Auto-Mute Applications:** Specify applications that will automatically mute your microphone when running.
- **System Tray Integration:** Enable or disable minimizing MicMaster to the system tray.
- **Desktop Shortcut:** Create or remove a desktop shortcut for easy access.

## Hotkey Configuration

Assigning a custom hotkey allows you to mute or unmute your microphone without navigating the application:

1. **Record Hotkey:**
   - Click the **Record Hotkey** button.
   - Press your desired key combination (e.g., `Ctrl + Alt + M`).
   - Click **Stop Recording** to finalize the hotkey.

2. **Using the Hotkey:**
   - Press your assigned key combination anytime to toggle your microphone's mute status.

## Auto-Mute Applications

Enhance your privacy by automatically muting your microphone when specific applications are active:

1. **Access Settings:**
   - Open the **Settings** window.

2. **Enable Auto-Mute:**
   - Check the **Enable Auto-Mute on Specific Applications** option.

3. **Select Applications:**
   - Click **Select Applications** to choose from a list of currently running applications.
   - Add or remove applications as needed.

4. **Save Settings:**
   - Click **Save** to apply the changes.

## System Tray Integration

Keep MicMaster running in the background without occupying space on your taskbar:

1. **Enable Tray Option:**
   - In the main interface, check the **Minimize to system tray** option.

2. **Minimize Application:**
   - Minimize MicMaster, and it will reside in the system tray.

3. **Restore Application:**
   - Click the MicMaster icon in the system tray to restore the main window.

## Checking for Updates

Stay up-to-date with the latest features and improvements:

1. **Manual Update Check:**
   - Click the **Check for Updates** button in the main interface.
   - If a new version is available, follow the prompts to download and install it.

2. **Automatic Notifications:**
   - MicMaster can notify you of updates upon launch if the feature is enabled in settings.

## Contributing

Contributions are welcome! To contribute to MicMaster:

1. **Fork the Repository:**
   - Click the **Fork** button on the repository page.

2. **Clone Your Fork:**
   ```bash
    git clone https://github.com/balki97/MicMaster.git
    cd MicMaster
    ```

3. **Create a New Branch:**
    ```bash
    git checkout -b feature/YourFeatureName
    ```

4. **Make Your Changes:**
    - Implement your feature or bug fix.

5. **Commit Your Changes:**
    ```bash
    git commit -m "Add feature: YourFeatureName"
    ```

6. **Push to Your Fork:**
    ```bash
    git push origin feature/YourFeatureName
    ```

7. **Create a Pull Request:**
    - Navigate to your fork on GitHub and click **Compare & pull request**.

Please ensure that your contributions adhere to the project's coding standards and include relevant tests where applicable.

---

© 2024 [Balki](https://github.com/balki97). All rights reserved.

---
