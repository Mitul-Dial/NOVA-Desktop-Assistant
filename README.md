<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" />
</p>

<h1 align="center">N · O · V · A</h1>
<p align="center"><strong>Neural Omni-capable Voice Assistant</strong></p>

<p align="center">
A sleek, futuristic desktop voice assistant for Windows. Activate NOVA with your voice, open apps, browse folders, close programs — all hands-free.
</p>

---

## ✨ Features

- 🎤 **Wake Word Detection** — Say *"Nova"* to activate (understands natural variations like "Noah", "Nora", etc.)
- 🚀 **Open Applications** — *"Open Chrome"*, *"Open Spotify"*, *"Open VS Code"* — any installed app
- 📂 **Drive & Folder Navigation** — *"Open Internship from M drive"* — works with any drive letter
- ❌ **Close Applications** — *"Close Chrome"*, *"Close Discord"*
- 💻 **Terminal Access** — *"Open Terminal"* or *"Open Command Prompt"*
- ⚡ **Smart Window Switching** — If an app is already open, NOVA switches to it instead of opening a duplicate
- 🎨 **Premium Dark UI** — Glassmorphism-inspired design with smooth animations
- ⚙️ **Custom Commands** — Add your own voice triggers for websites and apps
- 🔄 **Auto-Start** — Optional setting to activate listening on launch

---

## 🚀 Quick Start

### Prerequisites

- **Windows 10/11**
- **Python 3.10+** — [Download here](https://www.python.org/downloads/)
  - ✅ Make sure to check **"Add Python to PATH"** during installation
- **Microphone** — Any working microphone for voice commands

### Option 1: One-Click Build (Recommended)

1. **Clone or download** this repository:
   ```bash
   git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
   cd NOVA-Desktop-Assistant
   ```

2. **Double-click `build.bat`** — it will:
   - Install all dependencies
   - Generate the app icon
   - Build `NOVA.exe`
   - Create a Desktop shortcut

3. **Double-click "NOVA" on your Desktop** — done! 🎉

### Option 2: Run from Python

```bash
# Clone the repo
git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
cd NOVA-Desktop-Assistant

# Install dependencies
pip install -r requirements.txt

# Generate icon (first time only)
python generate_icon.py

# Run NOVA
python "NOVA Desktop Assistant.py"
```

---

## 🗣️ Voice Commands

| Command | What it does |
|---------|-------------|
| *"Nova"* | Wake up the assistant |
| *"Open Chrome"* | Opens Google Chrome (or switches to it) |
| *"Open [app name]"* | Opens any installed application |
| *"Open M drive"* | Opens M:\ drive in File Explorer |
| *"Open [folder] from [X] drive"* | Opens a specific folder from any drive |
| *"Close Chrome"* | Closes Google Chrome |
| *"Close [app name]"* | Closes the specified application |
| *"Open Terminal"* | Opens Command Prompt |

> 💡 **Tip:** You can add custom voice commands through **Settings > + Add New** for any website URL or file path.

---

## 📁 Project Structure

```
NOVA-Desktop-Assistant/
├── NOVA Desktop Assistant.py   # Main application
├── generate_icon.py            # Icon generator script
├── nova.ico                    # App icon (auto-generated)
├── build.bat                   # One-click build script
├── requirements.txt            # Python dependencies
├── LICENSE                     # MIT License
└── README.md                   # This file
```

---

## 🛠️ Building the EXE Manually

If you prefer to build manually:

```bash
pip install -r requirements.txt
python generate_icon.py

pyinstaller --onefile --windowed --name "NOVA" --icon "nova.ico" ^
    --add-data "nova.ico;." ^
    --hidden-import "pyttsx3.drivers" ^
    --hidden-import "pyttsx3.drivers.sapi5" ^
    --hidden-import "win32com.client" ^
    --hidden-import "pythoncom" ^
    --hidden-import "comtypes" ^
    --hidden-import "customtkinter" ^
    --hidden-import "speech_recognition" ^
    "NOVA Desktop Assistant.py"
```

The EXE will be in the `dist/` folder.

---

## 📝 Notes

- **User data** (custom commands, settings) is stored in `%APPDATA%\NOVA\` and persists across updates
- NOVA requires an **internet connection** for speech recognition (uses Google's speech-to-text API)
- The EXE is a single portable file — you can move it anywhere

---

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/Mitul-Dial">Mitul Dial</a>
</p>
