<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Platform-Windows_10%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" />
</p>

<h1 align="center">N · O · V · A</h1>
<p align="center"><strong>Neural Omni-capable Voice Assistant</strong></p>

<p align="center">
A hands-free voice-controlled desktop assistant for Windows that lets you open apps, navigate folders, close programs, and execute custom commands — all using your voice.
</p>

---

## 🤔 What is NOVA?

NOVA is a **desktop voice assistant** built entirely in Python. It runs as a standalone Windows application with a premium dark-themed interface. Once activated, NOVA continuously listens for your voice and performs actions on your computer — no typing, no clicking.

Think of it like having your own personal Jarvis — say *"Nova"* to wake it up, then tell it what to do.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🎤 **Wake Word Detection** | Say *"Nova"* to activate. Understands natural variations like *"Noah"*, *"Nora"*, etc. |
| 🚀 **Open Any Application** | *"Open Chrome"*, *"Open Spotify"*, *"Open VS Code"* — works with any installed app on your PC |
| 📂 **Folder & Drive Navigation** | *"Open M drive"* or *"Open Internship from M drive"* — browse any drive and folder by voice |
| ❌ **Close Applications** | *"Close Chrome"*, *"Close Discord"* — shut down apps hands-free |
| 💻 **Terminal Access** | *"Open Terminal"* or *"Open Command Prompt"* |
| 🔄 **Smart Window Switching** | If the app is already running, NOVA switches to it instead of opening a duplicate |
| ⚙️ **Custom Commands** | Add your own voice triggers for URLs, file paths, or programs through the settings panel |
| 🔁 **Auto-Start Mode** | Toggle auto-activation so NOVA starts listening the moment you launch it |
| 🎨 **Premium Dark UI** | Futuristic glassmorphism-inspired design with smooth pulse animations |
| 🖥️ **Single EXE** | Builds into one portable `.exe` file — no installation needed |

---

## 💡 Why Use NOVA?

- **Hands-free productivity** — Control your PC without touching the keyboard or mouse
- **Instant app launching** — Opens any app in seconds, faster than searching through the Start Menu
- **Works offline for UI** — The interface and app management run locally; only speech recognition needs internet
- **Lightweight** — Minimal CPU and memory usage when idle, runs quietly in the background
- **Fully customizable** — Add unlimited custom voice commands through the built-in settings panel
- **No account needed** — No sign-ups, no cloud services, no data collection — it runs 100% on your machine
- **Open source** — Read, modify, and improve the code however you like

---

## 🚀 Getting Started

### Prerequisites

Before you start, make sure you have:

1. **Windows 10 or 11**
2. **Python 3.10 or newer** — [Download from python.org](https://www.python.org/downloads/)
   > ⚠️ During Python installation, **check the box** that says **"Add Python to PATH"** — this is important!
3. **A working microphone** — Built-in laptop mic, headset, or external mic
4. **Internet connection** — Required for voice recognition (uses Google Speech-to-Text)

---

## 📦 Installation & Build

### Option 1: One-Click Build (Recommended)

The easiest way — just double-click one file and you're done:

1. **Download this repository:**
   - Click the green **"Code"** button on GitHub → **"Download ZIP"**
   - Extract the ZIP to any folder
   - Or clone with Git:
     ```bash
     git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
     ```

2. **Open the folder** and **double-click `build.bat`**
   - It automatically creates a virtual environment
   - Installs all dependencies (nothing pollutes your global Python)
   - Builds the `NOVA.exe`
   - Places a shortcut on your Desktop

3. **Double-click "NOVA" on your Desktop** — the app opens! 🎉

### Option 2: Run Directly with Python

If you just want to run it without building an EXE:

```bash
# 1. Clone the repo
git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
cd NOVA-Desktop-Assistant

# 2. Create a virtual environment
python -m venv .venv

# 3. Activate it
.venv\Scripts\activate

# 4. Install dependencies
pip install -r requirements.txt

# 5. Generate the icon (first time only)
python generate_icon.py

# 6. Run NOVA
python "NOVA Desktop Assistant.py"
```

---

## 🗣️ How to Use

### Step 1: Launch NOVA
Double-click the **NOVA** shortcut on your Desktop (or run the Python script). The app opens with a sleek dark interface.

### Step 2: Activate the Assistant
You have two ways:
- **Click the "ACTIVATE" button** in the center of the screen
- **Enable Auto-Start** in Settings so it activates automatically every time

### Step 3: Say the Wake Word
Once activated, NOVA continuously listens in the background. Say:
> **"Nova"**

NOVA will respond with *"Yes boss!"* and start listening for your command.

### Step 4: Give a Command
After the wake word, say your command clearly:

| What You Say | What Happens |
|-------------|-------------|
| *"Open Chrome"* | Opens Google Chrome (or switches to it if already open) |
| *"Open Spotify"* | Opens Spotify |
| *"Open VS Code"* | Opens Visual Studio Code |
| *"Open [any app]"* | Opens any installed application |
| *"Open M drive"* | Opens M:\ in File Explorer |
| *"Open photos from D drive"* | Opens the "photos" folder on D:\ |
| *"Open [folder] from [X] drive"* | Opens any folder from any drive |
| *"Close Chrome"* | Closes Google Chrome |
| *"Close [any app]"* | Closes the specified application |
| *"Open Terminal"* | Opens Command Prompt |

### Step 5: Add Custom Commands (Optional)
1. Click **⚙ Settings** in the top-right corner
2. Click **+ Add New**
3. Enter a **Voice Trigger** (e.g., *"google"*) and a **Target URL or path**
4. Now say *"Nova... open google"* and it opens!

### Step 6: Deactivate
Click the red **DEACTIVATE** button to stop listening. Close the window to exit completely.

---

## 📁 Project Structure

```
NOVA-Desktop-Assistant/
├── NOVA Desktop Assistant.py   # Main application source code
├── generate_icon.py            # Script to generate the app icon
├── nova.ico                    # App icon (auto-generated)
├── build.bat                   # One-click build script
├── requirements.txt            # Python dependencies
├── LICENSE                     # MIT License
└── README.md                   # This file
```

---

## 🛠️ Manual EXE Build

If you want to build the EXE manually without `build.bat`:

```bash
# Activate your virtual environment first
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Generate icon
python generate_icon.py

# Build EXE
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

The EXE will be at `dist/NOVA.exe`.

---

## 📝 Notes

- Your **custom commands and settings** are saved in `%APPDATA%\NOVA\` — they persist even if you update or rebuild the app
- NOVA requires an **internet connection** for speech recognition (it uses Google's free Speech-to-Text API)
- The built EXE is a **single portable file** — you can copy it to any Windows PC and run it directly
- First launch may take a few seconds as Windows verifies the executable

---

## ⚠️ Disclaimer

This software is provided **"as-is"** without any warranty of any kind, express or implied. By downloading, installing, or using NOVA, you acknowledge and agree that:

- **You use this software entirely at your own risk.** The developer is **not responsible** for any damage, data loss, system issues, or any other consequences that may arise from using this program.
- This is an **open-source hobby project** and is not guaranteed to be free of bugs or errors.
- NOVA interacts with your operating system (opening/closing apps, accessing folders). While it only performs actions you explicitly command, the developer takes **no liability** for any unintended behavior.
- Voice recognition accuracy depends on your microphone quality, internet connection, and environment — it may occasionally misinterpret commands.

**Use responsibly.** If you are unsure, review the source code before running.

---

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/Mitul-Dial">Mitul Dial</a>
</p>
