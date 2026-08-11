# NOVA Desktop Assistant

## Overview

NOVA is a hands-free, voice-controlled desktop assistant for Windows built entirely in Python. It runs as a standalone application featuring a premium dark-themed interface. Once activated, NOVA continuously listens for voice input and executes system operations, allowing users to open applications, navigate directories, close programs, and trigger custom workflows without keyboard or mouse interaction.

## Features

- **Wake Word Detection**: Activate the assistant naturally using the primary wake word "Nova" or recognized phonetic variations.
- **Application Management**: Launch or terminate installed applications system-wide (e.g., Chrome, Spotify, VS Code).
- **Directory Navigation**: Browse drives and folders directly via voice commands.
- **Smart Window Switching**: Detects if an application is already running and brings it to the foreground to prevent duplicate instances.
- **Terminal Access**: Quickly launch the Command Prompt or Terminal via voice.
- **Custom Commands**: Define custom voice triggers mapped to specific URLs, file paths, or executables through the built-in settings interface.
- **Auto-Start Mode**: Configure the application to begin listening immediately upon launch.
- **Premium User Interface**: Features a futuristic, dark-themed GUI built with CustomTkinter, incorporating smooth pulse animations and responsive feedback.
- **Standalone Executable**: Packaged as a single portable executable file requiring no complex installation.

## How It Works

The application operates on a continuous, multi-threaded listening loop once activated:

1. **User Input**: The user speaks the wake word followed by a command.
2. **Speech Processing**: Audio is captured via the microphone and processed using Google Speech-to-Text.
3. **Intent Handling**: The recognized text is parsed to extract the core intent (e.g., "open", "close") and the target (e.g., "Chrome", "M drive").
4. **Action Execution**: The system executes the requested operation using Windows API integrations (pywin32) or standard OS commands.
5. **Response**: NOVA provides auditory feedback confirming the action using offline text-to-speech (pyttsx3) and updates the GUI status.

## Tech Stack

| Technology | Purpose |
|------------|---------|
| **Python 3.10+** | Core application logic and multi-threading |
| **CustomTkinter** | Modern, dark-themed graphical user interface |
| **SpeechRecognition** | Audio capture and cloud-based speech-to-text processing |
| **PyTTSX3** | Offline text-to-speech engine for vocal feedback |
| **PyWin32** | Deep Windows OS integration for window management |
| **PyInstaller** | Bundling the application into a standalone executable |

## Project Structure

```text
NOVA-Desktop-Assistant/
├── assets/                     # Project screenshots
├── build.bat                   # Automated build script
├── generate_icon.py            # Icon generation utility
├── LICENSE                     # MIT License agreement
├── nova.ico                    # Compiled application icon
├── NOVA Desktop Assistant.py   # Main application source code
├── nova_commands.json          # Base configuration for commands
├── README.md                   # Project documentation
└── requirements.txt            # Python dependencies
```

## Installation

### Prerequisites

- Python 3.10 or newer
- An active microphone (built-in or external)
- An active internet connection (required for speech recognition)

### Option 1: Automated Build (Recommended)

1. Download or clone this repository:
   ```bash
   git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
   ```
2. Navigate to the project directory and execute `build.bat`.
   - This script automatically creates a virtual environment, installs all required dependencies, compiles the application into a single `.exe` file, and places a shortcut on your Desktop.
3. Launch the application using the created shortcut.

### Option 2: Run from Source

1. Clone the repository and navigate to the folder:
   ```bash
   git clone https://github.com/Mitul-Dial/NOVA-Desktop-Assistant.git
   cd NOVA-Desktop-Assistant
   ```
2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```
3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Generate the required application icon (first run only):
   ```bash
   python generate_icon.py
   ```
5. Execute the main script:
   ```bash
   python "NOVA Desktop Assistant.py"
   ```

## Configuration

NOVA automatically manages its configuration and user settings locally. No complex environment variables or secrets are required to run the application.

User data, including custom commands and UI preferences, is stored securely in:
`%APPDATA%\NOVA\`

This ensures that configurations persist across updates or rebuilds of the executable.

## Usage

1. Launch the NOVA application.
2. Click the central **ACTIVATE** button (or enable Auto-Start in the settings).
3. Say the wake word: **"Nova"**.
4. Issue a command. Examples include:
   - *"Open Chrome"*
   - *"Close Spotify"*
   - *"Open M drive"*
   - *"Open photos from D drive"*
   - *"Open Terminal"*
5. To configure custom commands, click the **Settings** icon, select **Add New**, and map a custom voice trigger to a URL or file path.

## Screenshots

### Main Interface
![NOVA Main Interface](assets/Main%20desktop.jfif)

### Active Listening Mode
![NOVA Active Listening Mode](assets/AI%20Turned%20On.jfif)

### Command Library
![NOVA Command Library](assets/Command%20Library.jfif)

### Custom Function Setup
![NOVA Custom Function Setup](assets/Save%20New%20Function.jfif)

## Technical Highlights

- **Multi-threaded Architecture**: The UI runs on the main thread while speech recognition and system operations are handled asynchronously, ensuring a responsive interface at all times.
- **Fuzzy Matching Logic**: Employs string similarity algorithms to match voice inputs with installed applications, increasing reliability even with minor misinterpretations.
- **Native OS Integration**: Utilizes COM interfaces and the Windows API to intelligently detect running applications and switch window focus rather than opening redundant instances.
- **Dynamic GUI Elements**: Features mathematically calculated rendering for continuous pulse and glow animations in the CustomTkinter canvas without relying on external image assets for state changes.

## Future Improvements

- Implementation of an offline wake-word engine (such as Porcupine) to reduce latency and reliance on cloud services for initial activation.
- Cross-platform support for macOS and Linux environments.
- Integration with large language models (LLMs) for complex query resolution and conversational capabilities.
- Advanced window management (e.g., resizing, snapping) via voice commands.

## Author

Mitul Dial

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
