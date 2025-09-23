# gui_manager.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import SUCCESS, DANGER, INFO, PRIMARY
import yaml
import aiohttp
import asyncio
import threading
import xml.etree.ElementTree as ET
import json
from pathlib import Path
from typing import Dict, List, Optional, Callable, Any, Tuple
import logging

# -----------------------
# Logging
# -----------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("LyriSyncGUI")

CONFIG_FILE = "lyrisync_config.yaml"


# =======================
# Config helpers
# =======================
def _default_config() -> Dict[str, Any]:
    """
    Defaults match your consolidated YAML:
      - top-level 'connections' (preferred)
      - plus legacy single-endpoint fields in 'settings' (for backward compatibility)
    """
    return {
        "roles": [],
        "ui": {"theme": "darkly"},
        "settings": {
            # legacy single-endpoint fields (main.py can still read these)
            "vmix_api_url": "http://localhost:8088/api",
            "openlp_ws_url": "ws://localhost:4317",   # legacy single WS entry
            "vmix_title_input": "SongTitle",
            "vmix_title_field": "Message.Text",

            # global behavior
            "api_port": 5000,
            "splash_enabled": True,
            "poll_interval_sec": 2,
            "overlay_channel": 1,
            "auto_overlay_on_send": True,
            "auto_overlay_out_on_clear": True,
            "overlay_always_on": False,
            "auto_clear_idle_sec": 0,
            "max_chars_per_line": 36,   # conservative default wrap
            "clear_on_blank": True,
            "text_layer_above": False,  # optional vMix title layer behavior

            # NOTE: connections now live at TOP-LEVEL; this remains only for backward compat.
        },
        # --- PREFERRED location for multi-connection bridge config ---
        "connections": []  # list of { name, openlp_ip, http_port, ws_port, vmix_api_url, mappings: [{input, field}, ...] }
    }


def load_config() -> Dict[str, Any]:
    path = Path(CONFIG_FILE)
    if not path.exists():
        return _default_config()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        base = _default_config()
        # shallow merge first
        base.update({k: v for k, v in data.items() if k != "ui" and k != "settings" and k != "connections"})
        # nested merges for ui, settings
        base["ui"] = {**_default_config()["ui"], **(data.get("ui") or {})}
        base["settings"] = {**_default_config()["settings"], **(data.get("settings") or {})}
        # top-level connections (preferred)
        base["connections"] = list(data.get("connections") or [])
        # roles
        base["roles"] = list(data.get("roles") or [])
        return base
    except Exception as e:
        logger.error("Failed to load config: %s", e)
        messagebox.showerror("Config Error", f"Failed to load configuration:\n{e}\nUsing defaults.")
        return _default_config()


def save_config(config: Dict[str, Any]) -> bool:
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            yaml.safe_dump(config, f, default_flow_style=False, allow_unicode=True)
        return True
    except Exception as e:
        logger.error("Failed to save config: %s", e)
        messagebox.showerror("Config Error", f"Failed to save configuration:\n{e}")
        return False


# =======================
# Async vMix discovery
# =======================
class AsyncVmixDiscoverer:
    def __init__(self):
        self._session: Optional[aiohttp.ClientSession] = None
        self._lock = threading.Lock()

    async def _get_session(self) -> aiohttp.ClientSession:
        with self._lock:
            if self._session is None or self._session.closed:
                self._session = aiohttp.ClientSession()
            return self._session

    async def discover_vmix_inputs(self, api_url: str) -> Tuple[List[str], Dict[str, List[str]]]:
        try:
            session = await self._get_session()
            async with session.get(api_url, timeout=5) as resp:
                if resp.status != 200:
                    raise RuntimeError(f"HTTP {resp.status}: {await resp.text()}")
                xml_text = await resp.text()
        except asyncio.TimeoutError:
            raise RuntimeError("vMix discovery timed out after 5 seconds")
        except Exception as e:
            raise RuntimeError(f"vMix discovery failed: {e}")

        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError as e:
            raise RuntimeError(f"Failed to parse vMix XML: {e}")

        input_names: List[str] = []
        fields_by_input: Dict[str, List[str]] = {}
        seen = set()

        for node in root.findall(".//inputs/input"):
            name = node.get("title") or node.get("shortTitle") or node.get("number") or "Unknown"
            if name not in seen:
                input_names.append(name)
                seen.add(name)

            fields: List[str] = []
            data_node = node.find("data")
            if data_node is not None:
                for t in data_node.findall("text"):
                    nm = t.get("name")
                    if nm and nm not in fields:
                        fields.append(nm)
            fields_by_input[name] = fields

        return input_names, fields_by_input

    async def close(self):
        with self._lock:
            if self._session and not self._session.closed:
                try:
                    asyncio.create_task(self._session.close())
                except Exception:
                    pass
                self._session = None


# =======================
# GUI
# =======================
class LyriSyncGUI:
    """
    LyriSync+ GUI
    - Roles tab: manage StreamDeck role mappings.
    - Live Status tab: test lyrics (multiline + auto-grow), overlay, recording.
    - Settings dialog: configure vMix, OpenLP, API, overlays, import JSON connections.
    - Works with top-level 'connections' in the YAML (preferred).
    """

    def __init__(
        self,
        master: tk.Tk,
        config: Dict[str, Any],
        on_config_save: Callable[[Dict[str, Any]], bool],
        action_callback: Optional[Callable[[Any], Any]] = None,
    ):
        self.master = master
        self.config = config
        self.on_config_save = on_config_save
        self.action_callback = action_callback

        self.discoverer = AsyncVmixDiscoverer()
        self._vmix_inputs: List[str] = []
        self._fields_by_input: Dict[str, List[str]] = {}

        # Window
        self.master.title("LyriSync+")
        self.master.geometry("980x640")
        self.master.minsize(820, 540)

        # Theme
        initial_theme = (self.config.get("ui") or {}).get("theme", "darkly")
        self.style = tb.Style(initial_theme)

        # Async loop for discovery/tasks
        self.loop = asyncio.new_event_loop()
        self.async_thread = threading.Thread(target=self._run_async_loop, daemon=True)
        self.async_thread.start()

        # Build UI
        self._build_ui()

        # Roles initial fill
        self.refresh_roles_list()

        # Close handling
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

    def _run_async_loop(self):
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_forever()
        finally:
            try:
                self.loop.close()
            except Exception:
                pass

    def thread_safe(self, fn: Callable, *args, **kwargs):
        self.master.after(0, lambda: fn(*args, **kwargs))

    def _on_close(self):
        async def _cleanup():
            await self.discoverer.close()
        try:
            asyncio.run_coroutine_threadsafe(_cleanup(), self.loop)
            self.loop.call_soon_threadsafe(self.loop.stop)
        except Exception:
            pass
        self.master.destroy()

    # -------------------
    # UI
    # -------------------
    def _build_ui(self):
        header = ttk.Frame(self.master)
        header.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(header, text="LyriSync+", font=("Segoe UI", 16, "bold")).pack(side="left")

        status_frame = ttk.Frame(header)
        status_frame.pack(side="right")

        self._openlp_led = self._led_group(status_frame, "OpenLP")
        self._vmix_led = self._led_group(status_frame, "vMix")
        self._ovr_led = self._led_group(status_frame, "Overlay")
        self._rec_led = self._led_group(status_frame, "Recording")

        right_tools = ttk.Frame(header)
        right_tools.pack(side="right", padx=(0, 12))
        ttk.Label(right_tools, text="Theme:").pack(side="left", padx=(0, 6))
        themes = ["darkly", "flatly", "cosmo", "pulse", "cyborg", "sandstone", "superhero", "morph", "journal", "simplex"]
        self.theme_var = tk.StringVar(value=self.style.theme.name)
        theme_dd = ttk.Combobox(right_tools, values=themes, textvariable=self.theme_var, state="readonly", width=12)
        theme_dd.pack(side="left")
        theme_dd.bind("<<ComboboxSelected>>", self._apply_theme)
        ttk.Button(right_tools, text="Settings", command=self.open_settings_dialog, bootstyle=PRIMARY).pack(side="left", padx=(10, 0))

        notebook = ttk.Notebook(self.master)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.roles_frame = ttk.Frame(notebook)
        self.status_frame = ttk.Frame(notebook)
        notebook.add(self.roles_frame, text="🎭 Roles & Decks")
        notebook.add(self.status_frame, text="📡 Live Status")

        self._build_roles_tab()
        self._build_status_tab()

    def _led_group(self, parent: ttk.Frame, caption: str) -> ttk.Label:
        frame = ttk.Frame(parent)
        frame.pack(side="left", padx=8)
        ttk.Label(frame, text=caption, font=("Segoe UI", 9)).pack()
        lbl = ttk.Label(frame, text="●", font=("Segoe UI", 12), foreground="#c43c3c")
        lbl.pack()
        return lbl

    def _apply_theme(self, _event=None):
        chosen = self.theme_var.get()
        try:
            self.style.theme_use(chosen)
            self.config.setdefault("ui", {})["theme"] = chosen
            save_config(self.config)
        except Exception as e:
            logger.error("Failed to apply theme: %s", e)
            messagebox.showerror("Theme Error", f"Failed to apply theme:\n{e}")

    # -------------------
    # Roles
    # -------------------
    def _build_roles_tab(self):
        tree_frame = ttk.Frame(self.roles_frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        yscroll = ttk.Scrollbar(tree_frame)
        yscroll.pack(side="right", fill="y")
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        xscroll.pack(side="bottom", fill="x")

        self.roles_tree = ttk.Treeview(
            tree_frame,
            columns=("Name", "Decks", "Buttons"),
            show="headings",
            height=14,
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set,
        )
        yscroll.config(command=self.roles_tree.yview)
        xscroll.config(command=self.roles_tree.xview)

        self.roles_tree.heading("Name", text="Role Name")
        self.roles_tree.heading("Decks", text="Deck IDs")
        self.roles_tree.heading("Buttons", text="Button Mappings")

        self.roles_tree.column("Name", width=160, minwidth=120)
        self.roles_tree.column("Decks", width=120, minwidth=80)
        self.roles_tree.column("Buttons", width=360, minwidth=220)

        self.roles_tree.pack(fill="both", expand=True)

        btns = ttk.Frame(self.roles_frame)
        btns.pack(pady=(6, 4))
        ttk.Button(btns, text="➕ Add Role", command=self.add_role).pack(side="left", padx=5)
        ttk.Button(btns, text="✏️ Edit Role", command=self.edit_role).pack(side="left", padx=5)
        ttk.Button(btns, text="❌ Delete Role", command=self.delete_role, bootstyle=DANGER).pack(side="left", padx=5)
        ttk.Button(btns, text="🔄 Refresh", command=self.refresh_roles_list).pack(side="left", padx=5)

    def refresh_roles_list(self):
        try:
            for iid in self.roles_tree.get_children():
                self.roles_tree.delete(iid)
            for role in self.config.get("roles", []):
                decks = ", ".join(str(d) for d in role.get("decks", []))
                buttons = ", ".join([f"{k} → {v}" for k, v in role.get("buttons", {}).items()])
                self.roles_tree.insert("", "end", values=(role.get("name", "Unnamed"), decks, buttons))
        except Exception as e:
            logger.error("Refresh roles failed: %s", e)
            messagebox.showerror("Roles Error", f"Failed to refresh roles:\n{e}")

    def add_role(self):
        RoleEditorDialog(self.master, None, None, self.config, self._on_role_saved).show()

    def edit_role(self):
        sel = self.roles_tree.selection()
        if not sel:
            messagebox.showwarning("Select Role", "Please select a role to edit.")
            return
        try:
            idx = self.roles_tree.index(sel[0])
            role = self.config["roles"][idx]
            RoleEditorDialog(self.master, role, idx, self.config, self._on_role_saved).show()
        except Exception as e:
            messagebox.showerror("Edit Error", f"Failed to edit role:\n{e}")

    def delete_role(self):
        sel = self.roles_tree.selection()
        if not sel:
            return
        idx = self.roles_tree.index(sel[0])
        role_name = self.config["roles"][idx].get("name", "Unnamed")
        if messagebox.askyesno("Confirm Delete", f"Delete role '{role_name}'?"):
            del self.config["roles"][idx]
            self.refresh_roles_list()
            save_config(self.config)

    def _on_role_saved(self, new_role, role_index):
        if role_index is not None:
            self.config["roles"][role_index] = new_role
        else:
            self.config["roles"].append(new_role)
        self.refresh_roles_list()
        save_config(self.config)

    # -------------------
    # Status tab (multiline Test Lyrics + auto-grow)
    # -------------------
    def _build_status_tab(self):
        # Controls frame
        test = ttk.LabelFrame(self.status_frame, text="Controls", padding=10)
        test.pack(fill="x", padx=10, pady=(12, 8))

        ttk.Label(test, text="Test Lyrics:").grid(row=0, column=0, sticky="nw", pady=(5, 0))

        # Multiline Text with scrollbar (double width, short height; auto-grows up to 6)
        text_wrap_frame = ttk.Frame(test)
        text_wrap_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=5)

        self._lyrics_text = tk.Text(
            text_wrap_frame,
            height=2,               # start at 2 lines
            width=84,               # wider
            wrap="word",
            font=("Segoe UI", 10)
        )
        yscroll = ttk.Scrollbar(text_wrap_frame, orient="vertical", command=self._lyrics_text.yview)
        self._lyrics_text.configure(yscrollcommand=yscroll.set)

        self._lyrics_text.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        # default text
        self._lyrics_text.insert("1.0", "SAMPLE LYRICS")

        # auto-grow binding (on change and key press)
        self._lyrics_text.bind("<<Modified>>", self._autogrow_text)
        self._lyrics_text.bind("<KeyRelease>", self._autogrow_text)

        ttk.Button(test, text="Show Lyrics", command=self._send_test_lyrics, bootstyle=SUCCESS).grid(row=0, column=2, padx=6, pady=5, sticky="n")
        ttk.Button(test, text="Clear", command=self._clear_lyrics, bootstyle=DANGER).grid(row=0, column=3, padx=0, pady=5, sticky="n")

        actions = ttk.Frame(test)
        actions.grid(row=1, column=0, columnspan=4, sticky="w", pady=(10, 0))
        ttk.Button(actions, text="Toggle Overlay", command=lambda: self._trigger_action("toggle_overlay")).pack(side="left", padx=5)
        ttk.Button(actions, text="Start Recording", command=lambda: self._trigger_action("start_recording"), bootstyle=SUCCESS).pack(side="left", padx=5)
        ttk.Button(actions, text="Stop Recording", command=lambda: self._trigger_action("stop_recording"), bootstyle=DANGER).pack(side="left", padx=5)

        # Connection status
        conn = ttk.LabelFrame(self.status_frame, text="Connection Status", padding=10)
        conn.pack(fill="x", padx=10, pady=(6, 10))
        ttk.Label(conn, text="vMix:").grid(row=0, column=0, sticky="w")
        self.vmix_status_var = tk.StringVar(value="Disconnected")
        ttk.Label(conn, textvariable=self.vmix_status_var, foreground="red").grid(row=0, column=1, padx=4)

        ttk.Label(conn, text="OpenLP:").grid(row=0, column=2, sticky="w")
        self.openlp_status_var = tk.StringVar(value="Disconnected")
        ttk.Label(conn, textvariable=self.openlp_status_var, foreground="red").grid(row=0, column=3, padx=4)

        # column/row growth
        test.columnconfigure(1, weight=1)
        test.rowconfigure(0, weight=1)
        text_wrap_frame.columnconfigure(0, weight=1)
        text_wrap_frame.rowconfigure(0, weight=1)

    def _autogrow_text(self, event=None):
        """Auto-adjust Text height between 2 and 6 lines based on content."""
        try:
            if event and str(event.type) == "<<Modified>>":
                self._lyrics_text.edit_modified(False)
            total_lines = int(self._lyrics_text.index("end-1c").split(".")[0])
            min_lines, max_lines = 2, 6
            new_h = max(min_lines, min(max_lines, total_lines))
            current_h = int(self._lyrics_text.cget("height"))
            if new_h != current_h:
                self._lyrics_text.configure(height=new_h)
        except Exception:
            pass

    # -------------------
    # Actions
    # -------------------
    def _trigger_action(self, action: str):
        if callable(self.action_callback):
            try:
                self.action_callback(action)
            except Exception as e:
                messagebox.showerror("Action Error", f"{e}")

    def _send_test_lyrics(self):
        if callable(self.action_callback):
            txt = self._lyrics_text.get("1.0", "end-1c").strip()
            if txt:
                txt = txt.upper()  # enforce all caps on send
                self.action_callback(("set_lyrics_text", txt))
                self.action_callback("show_lyrics")

    def _clear_lyrics(self):
        if callable(self.action_callback):
            self.action_callback("clear_lyrics")

    def set_recording(self, is_on: bool):
        self.thread_safe(self._rec_led.configure, foreground="#2ca34a" if is_on else "#c43c3c")

    def set_overlay(self, is_on: bool):
        self.thread_safe(self._ovr_led.configure, foreground="#2ca34a" if is_on else "#c43c3c")

    def set_conn_status(self, vmix_ok=None, openlp_ok=None):
        if vmix_ok is not None:
            self.thread_safe(self._vmix_led.configure, foreground="#2ca34a" if vmix_ok else "#c43c3c")
            self.thread_safe(self.vmix_status_var.set, "Connected" if vmix_ok else "Disconnected")
        if openlp_ok is not None:
            self.thread_safe(self._openlp_led.configure, foreground="#2ca34a" if openlp_ok else "#c43c3c")
            self.thread_safe(self.openlp_status_var.set, "Connected" if openlp_ok else "Disconnected")

    def open_settings_dialog(self):
        SettingsDialog(self.master, self.config, self.discoverer, self._apply_settings).show()

    def _apply_settings(self, new_settings: Dict[str, Any]):
        """
        Called by SettingsDialog on Save. We save 'settings' here.
        The dialog itself updates top-level 'connections' directly on the config
        BEFORE calling this, so a single save writes both.
        """
        self.config["settings"] = new_settings
        save_config(self.config)


# =======================
# Role Editor
# =======================
class RoleEditorDialog:
    def __init__(self, parent, role, role_index, config, on_save):
        self.parent = parent
        self.role = role or {}
        self.role_index = role_index
        self.config = config
        self.on_save = on_save
        self.window: Optional[tk.Toplevel] = None

    def show(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Role Editor")
        self.window.geometry("460x360")
        self.window.transient(self.parent)
        self.window.grab_set()

        frm = ttk.Frame(self.window, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Role Name:").grid(row=0, column=0, sticky="w", pady=6)
        self.name_var = tk.StringVar(value=self.role.get("name", ""))
        name_e = ttk.Entry(frm, textvariable=self.name_var)
        name_e.grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Deck IDs (comma-separated):").grid(row=1, column=0, sticky="w", pady=6)
        self.decks_var = tk.StringVar(value=", ".join(str(d) for d in self.role.get("decks", [])))
        ttk.Entry(frm, textvariable=self.decks_var).grid(row=1, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Button Mappings (key:action, comma-separated):").grid(row=2, column=0, sticky="w", pady=6)
        self.buttons_var = tk.StringVar(
            value=", ".join([f"{k}:{v}" for k, v in (self.role.get("buttons", {}) or {}).items()])
        )
        ttk.Entry(frm, textvariable=self.buttons_var).grid(row=2, column=1, sticky="ew", padx=8)

        help_txt = "Example: 0:show_lyrics, 1:clear_lyrics, 2:toggle_overlay"
        ttk.Label(frm, text=help_txt, foreground="gray").grid(row=3, column=1, sticky="w", padx=8, pady=(2, 10))

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="💾 Save", command=self._save, bootstyle=SUCCESS).pack(side="left", padx=8)
        ttk.Button(btns, text="Cancel", command=self.window.destroy).pack(side="left", padx=8)

        frm.columnconfigure(1, weight=1)
        name_e.focus()

    def _save(self):
        try:
            name = (self.name_var.get() or "").strip()
            if not name:
                messagebox.showerror("Validation", "Role name is required.")
                return

            decks: List[int] = []
            for part in (self.decks_var.get() or "").split(","):
                part = part.strip()
                if part.isdigit():
                    decks.append(int(part))

            buttons: Dict[str, str] = {}
            for item in (self.buttons_var.get() or "").split(","):
                item = item.strip()
                if ":" in item:
                    k, v = item.split(":", 1)
                    buttons[k.strip()] = v.strip()

            self.on_save({"name": name, "decks": decks, "buttons": buttons}, self.role_index)
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save role:\n{e}")


# =======================
# Settings dialog
# =======================
class SettingsDialog:
    def __init__(self, parent, config: Dict[str, Any], discoverer: AsyncVmixDiscoverer, on_apply: Callable[[Dict[str, Any]], None]):
        self.parent = parent
        self.config = config
        self.discoverer = discoverer
        self.on_apply = on_apply
        self.window: Optional[tk.Toplevel] = None

        s = self.config.get("settings", {})

        # Back-compat single-endpoint vMix/OpenLP fields
        self.vmix_api_var = tk.StringVar(value=s.get("vmix_api_url", "http://localhost:8088/api"))
        self.openlp_ws_var = tk.StringVar(value=s.get("openlp_ws_url", "ws://localhost:4317"))

        # General/global fields
        self.api_port_var = tk.StringVar(value=str(s.get("api_port", 5000)))
        self.input_var = tk.StringVar(value=s.get("vmix_title_input", "SongTitle"))
        self.field_var = tk.StringVar(value=s.get("vmix_title_field", "Message.Text"))
        self.splash_var = tk.BooleanVar(value=bool(s.get("splash_enabled", True)))
        self.poll_var = tk.StringVar(value=str(s.get("poll_interval_sec", 2)))
        self.overlay_var = tk.StringVar(value=str(s.get("overlay_channel", 1)))
        self.aoin_var = tk.BooleanVar(value=bool(s.get("auto_overlay_on_send", True)))
        self.aoout_var = tk.BooleanVar(value=bool(s.get("auto_overlay_out_on_clear", True)))
        self.always_on_var = tk.BooleanVar(value=bool(s.get("overlay_always_on", False)))
        self.idle_var = tk.StringVar(value=str(s.get("auto_clear_idle_sec", 0)))
        self.wrap_var = tk.StringVar(value=str(s.get("max_chars_per_line", 36)))
        self.cob_var = tk.BooleanVar(value=bool(s.get("clear_on_blank", True)))
        self.text_layer_above_var = tk.BooleanVar(value=bool(s.get("text_layer_above", False)))

        # PREFERRED multi-connection block (TOP-LEVEL)
        self.connections: List[Dict[str, Any]] = list(self.config.get("connections", []))

    def show(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Settings")
        self.window.geometry("820x720")
        self.window.transient(self.parent)
        self.window.grab_set()

        main = ttk.Frame(self.window, padding=12)
        main.pack(fill="both", expand=True)

        # vMix (single-endpoint legacy/testing)
        ttk.Label(main, text="vMix (Single-Endpoint / Legacy)", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        ttk.Label(main, text="vMix API URL:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.vmix_api_var, width=44).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(main, text="vMix Input:").grid(row=2, column=0, sticky="w", pady=5)
        self.input_combo = ttk.Combobox(main, textvariable=self.input_var, state="readonly", width=30)
        self.input_combo.grid(row=2, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="Discover", command=self._discover_inputs).grid(row=2, column=2, padx=6)

        ttk.Label(main, text="vMix Field:").grid(row=3, column=0, sticky="w", pady=5)
        self.field_combo = ttk.Combobox(main, textvariable=self.field_var, state="readonly", width=30)
        self.field_combo.grid(row=3, column=1, sticky="ew", padx=6)

        # OpenLP (single-endpoint legacy/testing)
        ttk.Label(main, text="OpenLP (Single-Endpoint / Legacy)", font=("Segoe UI", 12, "bold")).grid(row=4, column=0, columnspan=4, sticky="w", pady=(18, 10))
        ttk.Label(main, text="OpenLP WS URL:").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.openlp_ws_var, width=44).grid(row=5, column=1, sticky="ew", padx=6)

        # API
        ttk.Label(main, text="API / App Settings", font=("Segoe UI", 12, "bold")).grid(row=6, column=0, columnspan=4, sticky="w", pady=(18, 10))
        ttk.Label(main, text="LyriSync+ API Port:").grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.api_port_var, width=10).grid(row=7, column=1, sticky="w", padx=6)

        # Overlay behavior
        ttk.Label(main, text="Overlay Settings", font=("Segoe UI", 12, "bold")).grid(row=8, column=0, columnspan=4, sticky="w", pady=(18, 10))
        ttk.Label(main, text="Overlay Channel (1-4):").grid(row=9, column=0, sticky="w", pady=5)
        ttk.Combobox(main, textvariable=self.overlay_var, values=["1", "2", "3", "4"], state="readonly", width=6).grid(row=9, column=1, sticky="w", padx=6)

        ttk.Checkbutton(main, text="Auto Overlay on Send", variable=self.aoin_var).grid(row=10, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(main, text="Overlay Out on Clear", variable=self.aoout_var).grid(row=11, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(main, text="Overlay Always On", variable=self.always_on_var).grid(row=12, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(main, text="Clear on Blank Slide", variable=self.cob_var).grid(row=13, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(main, text="Title layer above text (vMix title)", variable=self.text_layer_above_var).grid(row=14, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(main, text="Show Splash Screen", variable=self.splash_var).grid(row=15, column=0, columnspan=2, sticky="w", pady=4)

        # Text / timing
        ttk.Label(main, text="Max Chars per Line:").grid(row=16, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.wrap_var, width=10).grid(row=16, column=1, sticky="w", padx=6)

        ttk.Label(main, text="Auto-Clear Idle (sec, 0=off):").grid(row=17, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.idle_var, width=10).grid(row=17, column=1, sticky="w", padx=6)

        ttk.Label(main, text="Poll Interval (sec):").grid(row=18, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.poll_var, width=10).grid(row=18, column=1, sticky="w", padx=6)

        # JSON import/export for multi-connection
        sep = ttk.Separator(main); sep.grid(row=19, column=0, columnspan=4, sticky="ew", pady=(14, 8))
        ttk.Label(main, text="Multi-Connection (Top-Level 'connections')", font=("Segoe UI", 12, "bold")).grid(row=20, column=0, columnspan=4, sticky="w", pady=(0, 8))
        self.conn_info_var = tk.StringVar(value=self._connections_summary())
        ttk.Label(main, textvariable=self.conn_info_var).grid(row=21, column=0, columnspan=3, sticky="w")
        btns_json = ttk.Frame(main); btns_json.grid(row=21, column=3, sticky="e")
        ttk.Button(btns_json, text="Import JSON…", command=self._import_json, bootstyle=INFO).pack(side="left", padx=4)
        ttk.Button(btns_json, text="Export JSON…", command=self._export_json).pack(side="left")

        # Bottom buttons
        btns = ttk.Frame(main)
        btns.grid(row=22, column=0, columnspan=4, pady=18, sticky="e")
        ttk.Button(btns, text="Test vMix Connection", command=self._test_vmix).pack(side="left", padx=6)
        ttk.Button(btns, text="Save", command=self._save_settings, bootstyle=SUCCESS).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.window.destroy).pack(side="left", padx=6)

        main.columnconfigure(1, weight=1)

    # ---- Discovery / Test (async) ----
    def _discover_inputs(self):
        async def _task():
            try:
                api = self.vmix_api_var.get().strip() or "http://localhost:8088/api"
                inputs, fmap = await self.discoverer.discover_vmix_inputs(api)
                def _apply():
                    self.input_combo["values"] = inputs
                    if inputs and self.input_var.get() not in inputs:
                        self.input_var.set(inputs[0])
                    current = self.input_var.get()
                    self.field_combo["values"] = fmap.get(current, [])
                    if fmap.get(current) and self.field_var.get() not in fmap[current]:
                        self.field_var.set(fmap[current][0])
                self.window.after(0, _apply)
            except Exception as e:
                self.window.after(0, lambda: messagebox.showerror("Discovery Error", f"Failed to discover vMix inputs:\n{e}"))
        asyncio.run_coroutine_threadsafe(_task(), asyncio.get_event_loop())

    def _test_vmix(self):
        async def _task():
            try:
                api = self.vmix_api_var.get().strip() or "http://localhost:8088/api"
                inputs, _ = await self.discoverer.discover_vmix_inputs(api)
                self.window.after(0, lambda: messagebox.showinfo("vMix", f"Connected. Found {len(inputs)} input(s)."))
            except Exception as e:
                self.window.after(0, lambda: messagebox.showerror("vMix", f"Connection failed:\n{e}"))
        asyncio.run_coroutine_threadsafe(_task(), asyncio.get_event_loop())

    # ---- Connections import/export
    def _connections_summary(self) -> str:
        n = len(self.connections)
        if n == 0:
            return "No connections configured."
        names = ", ".join((c.get("name") or f"#{i+1}") for i, c in enumerate(self.connections))
        return f"{n} connection(s): {names}"

    def _import_json(self):
        path = filedialog.askopenfilename(
            title="Select JSON config",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "connections" in data:
                conns = data["connections"]
            else:
                conns = data
            if not isinstance(conns, list):
                raise ValueError('JSON must be a list of connection objects or {"connections": [...]}')
            self.connections = conns
            self.conn_info_var.set(self._connections_summary())
            messagebox.showinfo("Import JSON", f"Loaded {len(conns)} connection(s). Save settings to persist.")
        except Exception as e:
            messagebox.showerror("Import JSON", f"Failed to import:\n{e}")

    def _export_json(self):
        path = filedialog.asksaveasfilename(
            title="Export JSON config",
            defaultextension=".json",
            initialfile="openlp_vmix_connections.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            payload = {"connections": self.connections}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Export JSON", f"Saved {len(self.connections)} connection(s) to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export JSON", f"Failed to export:\n{e}")

    # ---- Save settings
    def _save_settings(self):
        try:
            new_settings = {
                # legacy/single-endpoint (kept for backward compat and quick testing)
                "vmix_api_url": (self.vmix_api_var.get().strip() or "http://localhost:8088/api"),
                "openlp_ws_url": (self.openlp_ws_var.get().strip() or "ws://localhost:4317"),
                "vmix_title_input": (self.input_var.get().strip() or "SongTitle"),
                "vmix_title_field": (self.field_var.get().strip() or "Message.Text"),

                # global behavior
                "api_port": max(1024, min(65535, int(self.api_port_var.get().strip() or "5000"))),
                "splash_enabled": bool(self.splash_var.get()),
                "poll_interval_sec": max(1, int(self.poll_var.get().strip() or "2")),
                "overlay_channel": max(1, min(4, int(self.overlay_var.get().strip() or "1"))),
                "auto_overlay_on_send": bool(self.aoin_var.get()),
                "auto_overlay_out_on_clear": bool(self.aoout_var.get()),
                "overlay_always_on": bool(self.always_on_var.get()),
                "auto_clear_idle_sec": max(0, int(self.idle_var.get().strip() or "0")),
                "max_chars_per_line": max(10, int(self.wrap_var.get().strip() or "36")),
                "clear_on_blank": bool(self.cob_var.get()),
                "text_layer_above": bool(self.text_layer_above_var.get()),
            }
        except ValueError as e:
            messagebox.showerror("Settings", f"Invalid numeric value:\n{e}")
            return

        try:
            # IMPORTANT: persist connections at TOP-LEVEL
            self.config["connections"] = list(self.connections)
            # then persist other settings via callback (which saves the whole config)
            self.on_apply(new_settings)
            messagebox.showinfo("Settings", "Settings saved.")
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Settings", f"Failed to apply settings:\n{e}")
