import customtkinter as ctk
import speech_recognition as sr
import pyttsx3
import webbrowser
import threading
import os
import sys
import subprocess
import json
import math

# Try to import pythoncom ensuring it works in threads on Windows
try:
    import pythoncom
except ImportError:
    pythoncom = None

CONFIG_FILE = "nova_commands.json"
SETTINGS_FILE = "nova_settings.json"

# ──────────────────────────────────────────────
#  REFINED DESIGN SYSTEM
# ──────────────────────────────────────────────
COLORS = {
    "bg":           "#0A0A0F",   # Deep space background
    "surface":      "#13131A",   # Primary surface
    "surface2":     "#1A1A24",   # Secondary surface
    "surface3":     "#20202D",   # Hover states
    "border":       "#2A2A38",   # Subtle borders
    "border_focus": "#3A3A4C",   # Focused borders
    "text":         "#F0F0F5",   # Primary text
    "text_dim":     "#8A8A9E",   # Secondary text
    "text_muted":   "#5A5A6E",   # Tertiary text
    "accent":       "#6B5EFF",   # Primary accent (vibrant purple)
    "accent_hover": "#5A4BE5",   # Hover state
    "accent_glow":  "#8F82FF",   # Glow/highlight
    "accent_dim":   "#4A3DC8",   # Dimmed accent
    "danger":       "#FF4757",   # Delete/stop
    "danger_hover": "#E63946",
    "success":      "#2ECC71",   # Active/success
    "success_glow": "#4ADE80",
    "warn":         "#FFA726",   # Edit/warning
    "warn_hover":   "#FB8C00",
}

FONT = "Segoe UI"
SPACING = {
    "xs": 4,
    "sm": 8,
    "md": 12,
    "lg": 16,
    "xl": 24,
    "xxl": 32,
}


class NovaAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ── Window Configuration ──
        self.title("N O V A")
        self.geometry("1000x680")
        self.minsize(880, 600)
        self.configure(fg_color=COLORS["bg"])
        ctk.set_appearance_mode("Dark")

        # ── State Variables ──
        self.is_running = False
        self.recognizer = sr.Recognizer()
        self.thread = None
        self.apps = {}
        self.custom_commands = {}
        self._pulse_after_id = None
        self._glow_after_id = None
        self._pulse_angle = 0

        # ── Initialize ──
        self.load_custom_commands()
        self.load_settings()

        # ── Layout ──
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()
        threading.Thread(target=self.load_installed_apps, daemon=True).start()

        # Auto-activate if enabled
        if self.stay_active:
            self.after(600, self.start_assistant)

    # ═══════════════════════════════════════════
    #  UI CONSTRUCTION
    # ═══════════════════════════════════════════
    def _build_ui(self):
        # ── Header Bar with Gradient ──
        header = ctk.CTkFrame(self, height=56, fg_color=COLORS["surface"], corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        header.grid_propagate(False)

        # Logo section
        logo_frame = ctk.CTkFrame(header, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=SPACING["lg"], pady=0)

        ctk.CTkLabel(
            logo_frame, text="●",
            font=ctk.CTkFont(family=FONT, size=20),
            text_color=COLORS["accent_glow"]
        ).pack(side="left", padx=(0, SPACING["sm"]))

        ctk.CTkLabel(
            logo_frame, text="N O V A",
            font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
            text_color=COLORS["text"]
        ).pack(side="left")

        # Status indicator (top right)
        self.header_status = ctk.CTkLabel(
            header, text="● Standby",
            font=ctk.CTkFont(family=FONT, size=12),
            text_color=COLORS["text_dim"]
        )
        self.header_status.grid(row=0, column=1, padx=SPACING["md"])

        # Settings button
        self.menu_btn = ctk.CTkButton(
            header, text="⚙ Settings",
            font=ctk.CTkFont(family=FONT, size=12),
            width=100, height=36, corner_radius=10,
            fg_color="transparent",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_dim"],
            hover_color=COLORS["surface2"],
            command=self.open_menu
        )
        self.menu_btn.grid(row=0, column=2, padx=SPACING["lg"], pady=10)

        # ── Main Content Area ──
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=SPACING["xxl"], pady=SPACING["xl"])
        content.grid_rowconfigure((0, 6), weight=1)
        content.grid_columnconfigure(0, weight=1)

        # Title with spacing
        ctk.CTkLabel(
            content, text="N · O · V · A",
            font=ctk.CTkFont(family=FONT, size=56, weight="bold"),
            text_color=COLORS["text"]
        ).grid(row=1, column=0, pady=(0, SPACING["xs"]))

        # Subtitle
        self.subtitle_label = ctk.CTkLabel(
            content, text="Neural Omni-capable Voice Assistant",
            font=ctk.CTkFont(family=FONT, size=14),
            text_color=COLORS["text_dim"]
        )
        self.subtitle_label.grid(row=2, column=0, pady=(0, SPACING["sm"]))

        # Version/Status badge
        self.status_badge = ctk.CTkLabel(
            content, text="Ready",
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=COLORS["text_muted"],
            fg_color=COLORS["surface2"],
            corner_radius=12,
            padx=SPACING["md"],
            pady=SPACING["xs"]
        )
        self.status_badge.grid(row=3, column=0, pady=(0, SPACING["xxl"]))

        # ── Animated Ring Container ──
        ring_size = 240
        self.ring_canvas = ctk.CTkCanvas(
            content, width=ring_size, height=ring_size,
            bg=COLORS["bg"], highlightthickness=0
        )
        self.ring_canvas.grid(row=4, column=0, pady=SPACING["md"])

        # Draw static ring elements
        center = ring_size // 2
        # Outer subtle ring
        self.ring_canvas.create_oval(
            center - 105, center - 105, center + 105, center + 105,
            outline=COLORS["border"], width=1, tags="static"
        )

        # ── Main Control Button ──
        self.toggle_btn = ctk.CTkButton(
            content, text="ACTIVATE",
            font=ctk.CTkFont(family=FONT, size=16, weight="bold"),
            width=190, height=190, corner_radius=95,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_width=0,
            text_color="#FFFFFF",
            command=self.toggle_assistant
        )
        self.toggle_btn.grid(row=4, column=0)

        # ── Status Footer ──
        status_container = ctk.CTkFrame(
            content, fg_color=COLORS["surface"], corner_radius=14,
            border_width=1, border_color=COLORS["border"]
        )
        status_container.grid(row=5, column=0, pady=(SPACING["xxl"], 0), sticky="ew", padx=SPACING["xxl"] * 3)

        self.status_log = ctk.CTkLabel(
            status_container,
            text='Say  "Nova"  to begin',
            font=ctk.CTkFont(family=FONT, size=13),
            text_color=COLORS["text_dim"],
            height=52
        )
        self.status_log.pack(pady=SPACING["md"], padx=SPACING["lg"])

    # ═══════════════════════════════════════════
    #  ENHANCED ANIMATIONS
    # ═══════════════════════════════════════════
    def _start_pulse(self):
        """Start smooth rotating pulse animation"""
        self._pulse_angle = 0
        self._pulse_tick()

    def _pulse_tick(self):
        if not self.is_running:
            self.ring_canvas.delete("pulse")
            return

        center = 120
        radius = 100
        self._pulse_angle = (self._pulse_angle + 2) % 360

        # Create arc segments
        self.ring_canvas.delete("pulse")

        # Main rotating arc
        start_angle = self._pulse_angle
        extent = 120
        self.ring_canvas.create_arc(
            center - radius, center - radius,
            center + radius, center + radius,
            start=start_angle, extent=extent,
            outline=COLORS["accent_glow"], width=3,
            style="arc", tags="pulse"
        )

        # Secondary arc (opposite side, dimmer)
        start_angle2 = (self._pulse_angle + 180) % 360
        extent2 = 80
        self.ring_canvas.create_arc(
            center - radius, center - radius,
            center + radius, center + radius,
            start=start_angle2, extent=extent2,
            outline=COLORS["accent_dim"], width=2,
            style="arc", tags="pulse"
        )

        # Glow effect (inner pulse)
        pulse_phase = (self._pulse_angle % 360) / 360.0
        inner_radius = 95 + 5 * math.sin(pulse_phase * 2 * math.pi)
        alpha = int(50 + 30 * math.sin(pulse_phase * 2 * math.pi))
        
        self._pulse_after_id = self.after(20, self._pulse_tick)  # 50 fps

    def _stop_pulse(self):
        if self._pulse_after_id:
            self.after_cancel(self._pulse_after_id)
            self._pulse_after_id = None
        self.ring_canvas.delete("pulse")

    def _start_button_glow(self):
        """Subtle breathing glow effect on button"""
        self._glow_phase = 0
        self._button_glow_tick()

    def _button_glow_tick(self):
        if not self.is_running:
            return

        self._glow_phase = (self._glow_phase + 1) % 120
        intensity = 0.5 + 0.5 * math.sin((self._glow_phase / 120) * 2 * math.pi)
        
        # Subtle color shift
        if intensity > 0.7:
            self.toggle_btn.configure(fg_color=COLORS["danger"])
        else:
            self.toggle_btn.configure(fg_color=COLORS["danger_hover"])

        self._glow_after_id = self.after(50, self._button_glow_tick)

    def _stop_button_glow(self):
        if self._glow_after_id:
            self.after_cancel(self._glow_after_id)
            self._glow_after_id = None

    # ═══════════════════════════════════════════
    #  SETTINGS PANEL
    # ═══════════════════════════════════════════
    def open_menu(self):
        win = ctk.CTkToplevel(self)
        win.title("Settings")
        win.geometry("560x520")
        win.configure(fg_color=COLORS["bg"])
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        # Center on parent
        x = self.winfo_x() + (self.winfo_width() // 2) - 280
        y = self.winfo_y() + (self.winfo_height() // 2) - 260
        win.geometry(f"+{x}+{y}")

        # ── Header ──
        header = ctk.CTkFrame(win, fg_color=COLORS["surface"], corner_radius=0, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(fill="both", expand=True, padx=SPACING["lg"])

        ctk.CTkLabel(
            header_content, text="Command Library",
            font=ctk.CTkFont(family=FONT, size=18, weight="bold"),
            text_color=COLORS["text"]
        ).pack(side="left", pady=SPACING["lg"])

        # Add button with icon
        ctk.CTkButton(
            header_content, text="+ Add New",
            font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
            width=110, height=38, corner_radius=10,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            command=lambda: self.open_add_edit_dialog(None, None, win)
        ).pack(side="right", pady=SPACING["md"])

        # ── Commands List ──
        list_frame = ctk.CTkFrame(win, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=SPACING["lg"], pady=SPACING["md"])

        self.funcs_scroll = ctk.CTkScrollableFrame(
            list_frame,
            fg_color=COLORS["surface"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["border"],
            scrollbar_button_color=COLORS["border"],
            scrollbar_button_hover_color=COLORS["accent"]
        )
        self.funcs_scroll.pack(fill="both", expand=True)
        self.refresh_functions_list(self.funcs_scroll, win)

        # ── Footer Settings ──
        footer = ctk.CTkFrame(win, fg_color=COLORS["surface2"], corner_radius=12, height=70)
        footer.pack(fill="x", padx=SPACING["lg"], pady=(SPACING["md"], SPACING["lg"]))
        footer.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(fill="both", expand=True, padx=SPACING["lg"], pady=SPACING["md"])

        # Left side - toggle
        left_frame = ctk.CTkFrame(footer_content, fg_color="transparent")
        left_frame.pack(side="left", fill="y")

        self.stay_active_var = ctk.BooleanVar(value=self.stay_active)
        stay_switch = ctk.CTkSwitch(
            left_frame, text="",
            width=48, height=24,
            command=self.toggle_stay_active,
            variable=self.stay_active_var,
            button_color=COLORS["text"],
            progress_color=COLORS["accent"],
            fg_color=COLORS["border"]
        )
        stay_switch.pack(side="left")

        info_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        info_frame.pack(side="left", padx=(SPACING["md"], 0))

        ctk.CTkLabel(
            info_frame, text="Auto-Start",
            font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
            text_color=COLORS["text"], anchor="w"
        ).pack(anchor="w")

        ctk.CTkLabel(
            info_frame, text="Activate automatically on launch",
            font=ctk.CTkFont(family=FONT, size=11),
            text_color=COLORS["text_muted"], anchor="w"
        ).pack(anchor="w")

    def refresh_functions_list(self, scroll_frame, parent_window):
        """Enhanced command list with better cards"""
        for w in scroll_frame.winfo_children():
            w.destroy()

        if not self.custom_commands:
            empty_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            empty_frame.pack(pady=60)

            ctk.CTkLabel(
                empty_frame, text="📋",
                font=ctk.CTkFont(size=48),
                text_color=COLORS["text_muted"]
            ).pack()

            ctk.CTkLabel(
                empty_frame, text="No commands yet",
                font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
                text_color=COLORS["text_dim"]
            ).pack(pady=(SPACING["sm"], 0))

            ctk.CTkLabel(
                empty_frame, text="Click 'Add New' to create your first command",
                font=ctk.CTkFont(family=FONT, size=12),
                text_color=COLORS["text_muted"]
            ).pack()
            return

        for idx, (name, link) in enumerate(self.custom_commands.items()):
            # Command card
            card = ctk.CTkFrame(
                scroll_frame,
                fg_color=COLORS["surface2"],
                corner_radius=12,
                border_width=1,
                border_color=COLORS["border"]
            )
            card.pack(fill="x", pady=SPACING["xs"], padx=SPACING["sm"])

            # Content wrapper
            content = ctk.CTkFrame(card, fg_color="transparent")
            content.pack(fill="x", padx=SPACING["md"], pady=SPACING["md"])

            # Left - Info
            info_frame = ctk.CTkFrame(content, fg_color="transparent")
            info_frame.pack(side="left", fill="x", expand=True)

            ctk.CTkLabel(
                info_frame, text=name.title(),
                font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                text_color=COLORS["text"], anchor="w"
            ).pack(anchor="w")

            link_display = link if len(link) < 40 else link[:37] + "..."
            ctk.CTkLabel(
                info_frame, text=link_display,
                font=ctk.CTkFont(family=FONT, size=11),
                text_color=COLORS["text_dim"], anchor="w"
            ).pack(anchor="w", pady=(2, 0))

            # Right - Actions
            actions = ctk.CTkFrame(content, fg_color="transparent")
            actions.pack(side="right")

            # Edit button
            ctk.CTkButton(
                actions, text="✎",
                width=36, height=36, corner_radius=8,
                font=ctk.CTkFont(size=14),
                fg_color=COLORS["surface3"],
                hover_color=COLORS["warn"],
                text_color=COLORS["text_dim"],
                border_width=1,
                border_color=COLORS["border_focus"],
                command=lambda n=name, l=link: self.open_add_edit_dialog(n, l, parent_window)
            ).pack(side="left", padx=2)

            # Delete button
            ctk.CTkButton(
                actions, text="×",
                width=36, height=36, corner_radius=8,
                font=ctk.CTkFont(size=18),
                fg_color=COLORS["surface3"],
                hover_color=COLORS["danger"],
                text_color=COLORS["text_dim"],
                border_width=1,
                border_color=COLORS["border_focus"],
                command=lambda n=name: self.delete_function(n, scroll_frame, parent_window)
            ).pack(side="left", padx=2)

    def open_add_edit_dialog(self, edit_name, edit_link, parent_window):
        """Modern add/edit dialog"""
        dlg = ctk.CTkToplevel(self)
        dlg.title("Edit Function" if edit_name else "New Function")
        dlg.geometry("480x340")
        dlg.configure(fg_color=COLORS["bg"])
        dlg.transient(parent_window)
        dlg.grab_set()
        dlg.resizable(False, False)

        # Center on parent
        x = parent_window.winfo_x() + (parent_window.winfo_width() // 2) - 240
        y = parent_window.winfo_y() + (parent_window.winfo_height() // 2) - 170
        dlg.geometry(f"+{x}+{y}")

        # Main card
        card = ctk.CTkFrame(
            dlg, fg_color=COLORS["surface"],
            corner_radius=16, border_width=1, border_color=COLORS["border"]
        )
        card.pack(fill="both", expand=True, padx=SPACING["lg"], pady=SPACING["lg"])

        # Header
        header = ctk.CTkLabel(
            card,
            text="Edit Function" if edit_name else "New Function",
            font=ctk.CTkFont(family=FONT, size=16, weight="bold"),
            text_color=COLORS["text"]
        )
        header.pack(pady=(SPACING["lg"], SPACING["xl"]), padx=SPACING["lg"])

        # Form
        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(fill="both", expand=True, padx=SPACING["xl"], pady=(0, SPACING["lg"]))

        # Name field
        ctk.CTkLabel(
            form, text="Voice Trigger",
            font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
            text_color=COLORS["text_dim"], anchor="w"
        ).pack(fill="x", pady=(0, SPACING["xs"]))

        name_entry = ctk.CTkEntry(
            form, height=42, corner_radius=10,
            font=ctk.CTkFont(family=FONT, size=13),
            fg_color=COLORS["surface2"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            placeholder_text='e.g. "google"'
        )
        name_entry.pack(fill="x", pady=(0, SPACING["lg"]))
        if edit_name:
            name_entry.insert(0, edit_name)

        # Link field
        ctk.CTkLabel(
            form, text="Target (URL or Path)",
            font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
            text_color=COLORS["text_dim"], anchor="w"
        ).pack(fill="x", pady=(0, SPACING["xs"]))

        link_entry = ctk.CTkEntry(
            form, height=42, corner_radius=10,
            font=ctk.CTkFont(family=FONT, size=13),
            fg_color=COLORS["surface2"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            placeholder_text="https://example.com"
        )
        link_entry.pack(fill="x", pady=(0, SPACING["xl"]))
        if edit_link:
            link_entry.insert(0, edit_link)

        def save():
            new_name = name_entry.get().strip().lower()
            new_link = link_entry.get().strip()
            if not new_name or not new_link:
                return
            if edit_name and edit_name != new_name:
                self.custom_commands.pop(edit_name, None)
            self.custom_commands[new_name] = new_link
            self.save_custom_commands()
            dlg.destroy()
            if hasattr(self, "funcs_scroll") and self.funcs_scroll.winfo_exists():
                self.refresh_functions_list(self.funcs_scroll, parent_window)

        # Save button
        ctk.CTkButton(
            form, text="Save Function",
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
            height=44, corner_radius=10,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            command=save
        ).pack(fill="x")

    def delete_function(self, name, scroll_frame, parent_window):
        self.custom_commands.pop(name, None)
        self.save_custom_commands()
        self.refresh_functions_list(scroll_frame, parent_window)

    # ═══════════════════════════════════════════
    #  DATA PERSISTENCE
    # ═══════════════════════════════════════════
    def load_custom_commands(self):
        if not os.path.exists(CONFIG_FILE):
            for legacy in ("stark_commands.json", "anna_commands.json"):
                if os.path.exists(legacy):
                    try:
                        with open(legacy, "r") as f:
                            self.custom_commands = json.load(f)
                        break
                    except:
                        pass
        else:
            try:
                with open(CONFIG_FILE, "r") as f:
                    self.custom_commands = json.load(f)
            except:
                self.custom_commands = {}

    def save_custom_commands(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.custom_commands, f, indent=4)
        except:
            pass

    def load_installed_apps(self):
        try:
            paths = [
                os.path.join(os.environ.get("ProgramData", ""), "Microsoft", "Windows", "Start Menu", "Programs"),
                os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs"),
            ]
            for p in paths:
                if not os.path.exists(p):
                    continue
                for root, _, files in os.walk(p):
                    for f in files:
                        if f.lower().endswith((".lnk", ".exe", ".url")):
                            self.apps[os.path.splitext(f)[0].lower()] = os.path.join(root, f)

            ps = "Get-StartApps | Select-Object Name, AppID | ConvertTo-Json -Compress"
            res = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                capture_output=True, text=True, encoding="utf-8", errors="ignore"
            )
            if res.returncode == 0 and res.stdout.strip():
                data = json.loads(res.stdout)
                if isinstance(data, dict):
                    data = [data]
                for item in data:
                    n, i = item.get("Name", "").lower(), item.get("AppID")
                    if n and i:
                        self.apps[n] = i
        except:
            pass

    def load_settings(self):
        self.stay_active = False
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
                self.stay_active = data.get("stay_active", False)
            except:
                pass

    def save_settings(self):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump({"stay_active": self.stay_active}, f, indent=4)
        except:
            pass

    def toggle_stay_active(self):
        self.stay_active = self.stay_active_var.get()
        self.save_settings()

    # ═══════════════════════════════════════════
    #  VOICE ASSISTANT CONTROL
    # ═══════════════════════════════════════════
    def toggle_assistant(self):
        if not self.is_running:
            self.start_assistant()
        else:
            self.stop_assistant()

    def start_assistant(self):
        self.is_running = True

        # Update UI to active state
        self.toggle_btn.configure(
            text="LISTENING",
            fg_color=COLORS["danger"],
            hover_color=COLORS["danger_hover"]
        )
        self.subtitle_label.configure(
            text='Listening for "Nova" command',
            text_color=COLORS["success_glow"]
        )
        self.status_badge.configure(
            text="Online",
            text_color=COLORS["success"],
            fg_color=COLORS["surface2"]
        )
        self.header_status.configure(
            text="● Online",
            text_color=COLORS["success_glow"]
        )

        # Start animations
        self._start_pulse()
        self._start_button_glow()

        self.speak("Nova online.")

        # Start listening thread
        self.thread = threading.Thread(target=self.run_loop, daemon=True)
        self.thread.start()

    def stop_assistant(self):
        self.is_running = False

        # Stop animations
        self._stop_pulse()
        self._stop_button_glow()

        # Update UI to standby state
        self.toggle_btn.configure(
            text="ACTIVATE",
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"]
        )
        self.subtitle_label.configure(
            text="Neural Omni-capable Voice Assistant",
            text_color=COLORS["text_dim"]
        )
        self.status_badge.configure(
            text="Ready",
            text_color=COLORS["text_muted"],
            fg_color=COLORS["surface2"]
        )
        self.header_status.configure(
            text="● Standby",
            text_color=COLORS["text_dim"]
        )
        self.status_log.configure(text='Say  "Nova"  to begin')

    def init_engine(self):
        engine = pyttsx3.init()
        engine.setProperty("rate", 165)
        voices = engine.getProperty("voices")
        for v in voices:
            if "zira" in v.name.lower() or "female" in v.name.lower():
                engine.setProperty("voice", v.id)
                break
        return engine

    def speak(self, text):
        self.status_log.configure(text=f"Nova: {text}")
        try:
            if pythoncom:
                pythoncom.CoInitialize()
            engine = self.init_engine()
            engine.say(text)
            engine.runAndWait()
            if pythoncom:
                pythoncom.CoUninitialize()
        except:
            pass

    # ═══════════════════════════════════════════
    #  COMMAND PROCESSING
    # ═══════════════════════════════════════════
    def close_application(self, app_name):
        self.speak(f"Closing {app_name}")
        try:
            proc = app_name
            if "chrome" in app_name:
                proc = "chrome"
            elif "edge" in app_name:
                proc = "msedge"
            elif "calc" in app_name:
                proc = "calculator"
            elif "notepad" in app_name:
                proc = "notepad"
            elif "code" in app_name:
                proc = "code"
            if not proc.endswith(".exe"):
                proc += ".exe"
            subprocess.Popen(
                f"taskkill /F /IM {proc}",
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except:
            pass

    def process_command(self, command):
        command = command.lower()

        # Close application
        if command.startswith("close "):
            self.close_application(command.replace("close ", "").strip())
            return

        # Custom commands
        for key, link in self.custom_commands.items():
            if f"open {key}" in command or key in command:
                self.speak(f"Opening {key}")
                if link.startswith("http") or link.startswith("www"):
                    webbrowser.open(link)
                else:
                    try:
                        os.startfile(link)
                    except:
                        subprocess.Popen(["explorer", link], shell=True)
                return

        # System applications
        if command.startswith("open "):
            target = command.replace("open ", "").strip()
            launch_path = self.apps.get(target)
            if not launch_path:
                matches = sorted([k for k in self.apps if target in k], key=len)
                if matches:
                    launch_path = self.apps[matches[0]]
                    target = matches[0]
            if launch_path:
                self.speak(f"Opening {target}")
                try:
                    if os.path.exists(launch_path):
                        os.startfile(launch_path)
                    else:
                        subprocess.Popen(["explorer", f"shell:AppsFolder\\{launch_path}"])
                except:
                    pass
                return

        # Command prompt
        if "command prompt" in command or "terminal" in command:
            self.speak("Opening Terminal")
            os.system("start cmd")

    # ═══════════════════════════════════════════
    #  LISTENING LOOP
    # ═══════════════════════════════════════════
    def run_loop(self):
        if pythoncom:
            pythoncom.CoInitialize()

        while self.is_running:
            try:
                with sr.Microphone() as source:
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                    try:
                        audio = self.recognizer.listen(source, timeout=1, phrase_time_limit=5)
                    except sr.WaitTimeoutError:
                        continue

                    if not self.is_running:
                        break

                    try:
                        word = self.recognizer.recognize_google(audio).lower()
                        print(f"Heard: {word}")
                    except:
                        continue

                    # Wake word triggers
                    triggers = [
                        "nova", "noah", "nora", "novah", "nover",
                        "know va", "now a", "no va"
                    ]
                    if any(t in word for t in triggers):
                        self.speak("Yes?")
                        self.subtitle_label.configure(
                            text="Processing command...",
                            text_color=COLORS["warn"]
                        )
                        self.status_badge.configure(
                            text="Processing",
                            text_color=COLORS["warn"]
                        )

                        with sr.Microphone() as src2:
                            self.recognizer.adjust_for_ambient_noise(src2, duration=0.5)
                            try:
                                a2 = self.recognizer.listen(src2, timeout=5, phrase_time_limit=5)
                                cmd = self.recognizer.recognize_google(a2)
                                self.status_log.configure(text=f"Command: {cmd}")
                                self.process_command(cmd)
                            except sr.WaitTimeoutError:
                                self.speak("Timed out.")
                            except sr.UnknownValueError:
                                self.speak("Could not understand.")

                        # Reset to listening state
                        self.subtitle_label.configure(
                            text='Listening for "Nova" command',
                            text_color=COLORS["success_glow"]
                        )
                        self.status_badge.configure(
                            text="Online",
                            text_color=COLORS["success"]
                        )

            except Exception as e:
                print(f"Error: {e}")

        if pythoncom:
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    app = NovaAssistant()
    app.mainloop()