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
from preach_info_db import PreachInfoDB

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
    return {
        "roles": [],
        "ui": {"theme": "darkly"},
        "settings": {
            # legacy single-connection fields are still supported/used by main.py
            "vmix_api_url": "http://localhost:8088/api",
            "openlp_ws_url": "ws://localhost:4317",
            "api_port": 5000,
            "vmix_title_input": "SongTitle",
            "vmix_title_field": "Message.Text",

            "splash_enabled": True,
            "poll_interval_sec": 2,
            "overlay_channel": 1,
            "auto_overlay_on_send": True,
            "auto_overlay_out_on_clear": True,
            "overlay_always_on": False,
            "auto_clear_idle_sec": 0,
            "max_chars_per_line": 36,        # conservative default wrap
            "clear_on_blank": True,
            "text_layer_above": False,       # optional vMix title layer behavior

            # NEW: multi-connection list (each connection has openlp/vmix + mappings)
            "connections": [],
            "preach_db_path": "lyrisync_preach.db",
        }
    }


def load_config() -> Dict[str, Any]:
    path = Path(CONFIG_FILE)
    if not path.exists():
        return _default_config()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        base = _default_config()
        base.update(data)
        base["ui"] = {**_default_config()["ui"], **(data.get("ui") or {})}
        base["settings"] = {**_default_config()["settings"], **(data.get("settings") or {})}
        base["roles"] = data.get("roles") or []
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
# Connection Editor (used in tab + settings)
# =======================
class ConnectionEditorDialog:
    """
    Quick editor for one OpenLP→vMix connection with mappings.
    Produces/edits a dict compatible with settings['connections'].
    """
    def __init__(self, parent, on_save: Callable[[Dict[str, Any]], None], seed: Optional[Dict[str, Any]] = None):
        self.parent = parent
        self.on_save = on_save
        seed = seed or {}
        self.win: Optional[tk.Toplevel] = None

        # basic fields
        self.name_var = tk.StringVar(value=seed.get("name", "Connection"))
        self.openlp_ip_var = tk.StringVar(value=seed.get("openlp_ip", "127.0.0.1"))
        self.http_port_var = tk.StringVar(value=str(seed.get("http_port", 4316)))
        self.ws_port_var = tk.StringVar(value=str(seed.get("ws_port", 4317)))
        self.vmix_api_var = tk.StringVar(value=seed.get("vmix_api_url", "http://localhost:8088/api"))

        # mappings list: list[{"input": "...", "field": "..."}]
        self._mappings: List[Dict[str, str]] = list(seed.get("mappings", []))

    def show(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("Connection")
        self.win.geometry("560x500")
        self.win.transient(self.parent)
        self.win.grab_set()

        frm = ttk.Frame(self.win, padding=12)
        frm.pack(fill="both", expand=True)

        r = 0
        ttk.Label(frm, text="Name:").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.name_var, width=30).grid(row=r, column=1, sticky="w", padx=6); r+=1

        ttk.Label(frm, text="OpenLP IP:").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.openlp_ip_var, width=18).grid(row=r, column=1, sticky="w", padx=6); r+=1

        ttk.Label(frm, text="HTTP Port:").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.http_port_var, width=10).grid(row=r, column=1, sticky="w", padx=6); r+=1

        ttk.Label(frm, text="WS Port:").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.ws_port_var, width=10).grid(row=r, column=1, sticky="w", padx=6); r+=1

        ttk.Label(frm, text="vMix API URL:").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.vmix_api_var, width=36).grid(row=r, column=1, sticky="w", padx=6); r+=1

        # mappings section
        sep = ttk.Separator(frm); sep.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8,6)); r+=1
        ttk.Label(frm, text="Mappings (vMix Input → Field)").grid(row=r, column=0, columnspan=2, sticky="w", pady=(0,6)); r+=1

        map_frame = ttk.Frame(frm); map_frame.grid(row=r, column=0, columnspan=2, sticky="nsew")
        frm.rowconfigure(r, weight=1); frm.columnconfigure(1, weight=1)

        cols = ("input", "field")
        self.map_tree = ttk.Treeview(map_frame, columns=cols, show="headings", height=6)
        for c in cols:
            self.map_tree.heading(c, text=c.capitalize())
            self.map_tree.column(c, width=200 if c=="input" else 250, stretch=True)
        self.map_tree.pack(side="left", fill="both", expand=True)

        for m in self._mappings:
            self.map_tree.insert("", "end", values=(m.get("input",""), m.get("field","")))

        yscroll = ttk.Scrollbar(map_frame, orient="vertical", command=self.map_tree.yview)
        self.map_tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")

        ctrl = ttk.Frame(frm); ctrl.grid(row=r+1, column=0, columnspan=2, sticky="w", pady=(6,0))
        in_var = tk.StringVar(); field_var = tk.StringVar()
        ttk.Entry(ctrl, textvariable=in_var, width=20).pack(side="left", padx=(0,6))
        ttk.Entry(ctrl, textvariable=field_var, width=24).pack(side="left", padx=(0,6))
        ttk.Button(ctrl, text="Add", command=lambda:self._add_map(in_var, field_var)).pack(side="left")
        ttk.Button(ctrl, text="Delete Selected", command=self._del_selected_map).pack(side="left", padx=6)

        btns = ttk.Frame(frm); btns.grid(row=r+2, column=0, columnspan=2, sticky="e", pady=12)
        ttk.Button(btns, text="Save", command=self._save, bootstyle=SUCCESS).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.win.destroy).pack(side="left")

    def _add_map(self, in_var: tk.StringVar, field_var: tk.StringVar):
        i = (in_var.get() or "").strip()
        f = (field_var.get() or "").strip()
        if not i or not f:
            messagebox.showerror("Mapping", "Both Input and Field are required.")
            return
        self.map_tree.insert("", "end", values=(i, f))
        in_var.set(""); field_var.set("")

    def _del_selected_map(self):
        for iid in self.map_tree.selection():
            self.map_tree.delete(iid)

    def _save(self):
        try:
            name = (self.name_var.get() or "").strip()
            if not name:
                raise ValueError("Name is required.")
            ip = (self.openlp_ip_var.get() or "").strip()
            http_port = int(self.http_port_var.get() or "4316")
            ws_port = int(self.ws_port_var.get() or "4317")
            vmix_api = (self.vmix_api_var.get() or "").strip()
            if not vmix_api.startswith("http://") and not vmix_api.startswith("https://"):
                raise ValueError("vMix API URL must start with http:// or https://")

            mappings: List[Dict[str,str]] = []
            for iid in self.map_tree.get_children():
                v = self.map_tree.item(iid, "values")
                mappings.append({"input": v[0], "field": v[1]})
            if not mappings:
                raise ValueError("Add at least one mapping (Input → Field).")

            payload = {
                "name": name,
                "openlp_ip": ip,
                "http_port": http_port,
                "ws_port": ws_port,
                "vmix_api_url": vmix_api,
                "mappings": mappings,
            }
            self.on_save(payload)
            self.win.destroy()
        except Exception as e:
            messagebox.showerror("Save Connection", str(e))


# =======================
# GUI
# =======================
class LyriSyncGUI:
    """
    LyriSync+ GUI
    - Roles tab: manage StreamDeck role mappings.
    - Connections tab: add/edit/delete OpenLP→vMix connections (with mappings).
    - Live Status tab: test lyrics (multiline + auto-grow), overlay, recording.
    - Settings dialog: configure legacy single-connection settings + Quick Add + JSON import.
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
        self.preach_db = PreachInfoDB(
            (self.config.get("settings") or {}).get("preach_db_path", "lyrisync_preach.db")
        )
        self._preach_rows: List[Dict[str, Any]] = []

        # Window
        self.master.title("LyriSync+")
        self.master.geometry("1024x680")
        self.master.minsize(880, 560)

        # Theme
        initial_theme = (self.config.get("ui") or {}).get("theme", "darkly")
        self.style = tb.Style(initial_theme)

        # Async loop for discovery/tasks (used by discover/test vMix)
        self.loop = asyncio.new_event_loop()
        self.async_thread = threading.Thread(target=self._run_async_loop, daemon=True)
        self.async_thread.start()

        # Build UI
        self._build_ui()

        # Roles initial fill
        self.refresh_roles_list()
        # Connections initial fill
        self.refresh_connections_list()
        # Preach info initial fill
        self.refresh_preach_list()

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
        self.connections_frame = ttk.Frame(notebook)
        self.status_frame = ttk.Frame(notebook)
        self.preach_frame = ttk.Frame(notebook)

        notebook.add(self.roles_frame, text="🎭 Roles & Decks")
        notebook.add(self.connections_frame, text="🔗 Connections")
        notebook.add(self.status_frame, text="📡 Live Status")
        notebook.add(self.preach_frame, text="📖 Preach Info")

        self._build_roles_tab()
        self._build_connections_tab()
        self._build_status_tab()
        self._build_preach_tab()

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
    # Connections tab
    # -------------------
    def _build_connections_tab(self):
        wrap = ttk.Frame(self.connections_frame)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        yscroll = ttk.Scrollbar(wrap)
        yscroll.pack(side="right", fill="y")

        self.conn_tree = ttk.Treeview(
            wrap,
            columns=("Name", "OpenLP", "Ports", "vMix API", "Mappings"),
            show="headings",
            height=14,
            yscrollcommand=yscroll.set,
        )
        yscroll.config(command=self.conn_tree.yview)

        self.conn_tree.heading("Name", text="Name")
        self.conn_tree.heading("OpenLP", text="OpenLP IP")
        self.conn_tree.heading("Ports", text="HTTP/WS")
        self.conn_tree.heading("vMix API", text="vMix API URL")
        self.conn_tree.heading("Mappings", text="Mappings")

        self.conn_tree.column("Name", width=160)
        self.conn_tree.column("OpenLP", width=120)
        self.conn_tree.column("Ports", width=100)
        self.conn_tree.column("vMix API", width=260)
        self.conn_tree.column("Mappings", width=320)

        self.conn_tree.pack(fill="both", expand=True)

        btns = ttk.Frame(self.connections_frame)
        btns.pack(pady=(6, 4))

        ttk.Button(btns, text="➕ Add", command=self._add_connection).pack(side="left", padx=5)
        ttk.Button(btns, text="✏️ Edit", command=self._edit_connection).pack(side="left", padx=5)
        ttk.Button(btns, text="❌ Delete", command=self._delete_connection, bootstyle=DANGER).pack(side="left", padx=5)
        ttk.Button(btns, text="📥 Import JSON…", command=self._import_connections_json, bootstyle=INFO).pack(side="left", padx=5)
        ttk.Button(btns, text="💾 Save", command=self._save_connections, bootstyle=SUCCESS).pack(side="left", padx=5)

    def refresh_connections_list(self):
        try:
            for iid in self.conn_tree.get_children():
                self.conn_tree.delete(iid)
            conns = self.config.get("settings", {}).get("connections", []) or []
            for c in conns:
                maps = ", ".join([f"{m.get('input','')}→{m.get('field','')}" for m in c.get("mappings", [])])
                ports = f"{c.get('http_port',4316)}/{c.get('ws_port',4317)}"
                self.conn_tree.insert(
                    "", "end",
                    values=(
                        c.get("name", "Connection"),
                        c.get("openlp_ip", "127.0.0.1"),
                        ports,
                        c.get("vmix_api_url", ""),
                        maps
                    )
                )
        except Exception as e:
            logger.error("Refresh connections failed: %s", e)
            messagebox.showerror("Connections Error", f"Failed to refresh connections:\n{e}")

    def _add_connection(self):
        def _on_save(new_conn: Dict[str, Any]):
            self.config["settings"].setdefault("connections", []).append(new_conn)
            self.refresh_connections_list()
        ConnectionEditorDialog(self.master, on_save=_on_save).show()

    def _edit_connection(self):
        sel = self.conn_tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a connection to edit.")
            return
        idx = self.conn_tree.index(sel[0])
        conns = self.config["settings"].get("connections", [])
        if idx >= len(conns):
            return
        def _on_save(updated: Dict[str, Any]):
            conns[idx] = updated
            self.refresh_connections_list()
        ConnectionEditorDialog(self.master, on_save=_on_save, seed=conns[idx]).show()

    def _delete_connection(self):
        sel = self.conn_tree.selection()
        if not sel:
            return
        idx = self.conn_tree.index(sel[0])
        conns = self.config["settings"].get("connections", [])
        if idx >= len(conns):
            return
        name = conns[idx].get("name", "Connection")
        if messagebox.askyesno("Confirm Delete", f"Delete connection '{name}'?"):
            del conns[idx]
            self.refresh_connections_list()

    def _import_connections_json(self):
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
            self.config["settings"]["connections"] = conns
            self.refresh_connections_list()
            messagebox.showinfo("Import JSON", f"Loaded {len(conns)} connection(s). Click Save to persist.")
        except Exception as e:
            messagebox.showerror("Import JSON", f"Failed to import:\n{e}")

    def _save_connections(self):
        if save_config(self.config):
            messagebox.showinfo("Connections", "Connections saved to config.")

    # -------------------
    # Status tab (multiline Test Lyrics + auto-grow)
    # -------------------
    def _build_status_tab(self):
        # Controls frame
        test = ttk.LabelFrame(self.status_frame, text="Controls", padding=10)
        test.pack(fill="x", padx=10, pady=(12, 8))

        ttk.Label(test, text="Test Lyrics:").grid(row=0, column=0, sticky="nw", pady=(5, 0))

        # Multiline Text with scrollbar
        text_wrap_frame = ttk.Frame(test)
        text_wrap_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=5)

        self._lyrics_text = tk.Text(
            text_wrap_frame,
            height=2,               # start at 2 lines (shorter)
            width=58,               # wider typing area
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

    # -------------------
    # Preach info tab (SQLite-backed)
    # -------------------
    def _build_preach_tab(self):
        outer = ttk.Frame(self.preach_frame)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        tree_wrap = ttk.Frame(outer)
        tree_wrap.pack(fill="both", expand=True)

        yscroll = ttk.Scrollbar(tree_wrap)
        yscroll.pack(side="right", fill="y")

        self.preach_tree = ttk.Treeview(
            tree_wrap,
            columns=("ID", "Name", "Title", "Scriptures", "Inspirations", "Subjects"),
            show="headings",
            height=12,
            yscrollcommand=yscroll.set,
        )
        yscroll.config(command=self.preach_tree.yview)

        self.preach_tree.heading("ID", text="ID")
        self.preach_tree.heading("Name", text="Preacher Name")
        self.preach_tree.heading("Title", text="Message Title")
        self.preach_tree.heading("Scriptures", text="Scriptures")
        self.preach_tree.heading("Inspirations", text="Inspirations")
        self.preach_tree.heading("Subjects", text="Subjects")

        self.preach_tree.column("ID", width=55, anchor="center")
        self.preach_tree.column("Name", width=150)
        self.preach_tree.column("Title", width=170)
        self.preach_tree.column("Scriptures", width=180)
        self.preach_tree.column("Inspirations", width=200)
        self.preach_tree.column("Subjects", width=170)
        self.preach_tree.pack(fill="both", expand=True)

        manage = ttk.Frame(outer)
        manage.pack(fill="x", pady=(8, 4))
        ttk.Button(manage, text="➕ Add", command=self._add_preach).pack(side="left", padx=5)
        ttk.Button(manage, text="✏️ Edit", command=self._edit_preach).pack(side="left", padx=5)
        ttk.Button(manage, text="❌ Delete", command=self._delete_preach, bootstyle=DANGER).pack(side="left", padx=5)
        ttk.Button(manage, text="🔄 Refresh", command=self.refresh_preach_list).pack(side="left", padx=5)

        overlays = ttk.LabelFrame(outer, text="Display as Overlay", padding=10)
        overlays.pack(fill="x", pady=(6, 4))
        ttk.Label(
            overlays,
            text="Select one row, then press a field to send it live to the configured vMix title overlay.",
        ).pack(anchor="w", pady=(0, 8))

        btns = ttk.Frame(overlays)
        btns.pack(fill="x")
        ttk.Button(btns, text="Name", command=lambda: self._show_preach_field("name"), bootstyle=INFO).pack(side="left", padx=5)
        ttk.Button(btns, text="Title", command=lambda: self._show_preach_field("title"), bootstyle=INFO).pack(side="left", padx=5)
        ttk.Button(btns, text="Scriptures", command=lambda: self._show_preach_field("scriptures"), bootstyle=INFO).pack(side="left", padx=5)
        ttk.Button(btns, text="Inspirations", command=lambda: self._show_preach_field("inspirations"), bootstyle=INFO).pack(side="left", padx=5)
        ttk.Button(btns, text="Subjects", command=lambda: self._show_preach_field("subjects"), bootstyle=INFO).pack(side="left", padx=5)

    def refresh_preach_list(self):
        try:
            for iid in self.preach_tree.get_children():
                self.preach_tree.delete(iid)
            self._preach_rows = self.preach_db.list_entries()
            for row in self._preach_rows:
                self.preach_tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("id"),
                        row.get("name", ""),
                        row.get("title", ""),
                        row.get("scriptures", ""),
                        row.get("inspirations", ""),
                        row.get("subjects", ""),
                    ),
                )
        except Exception as e:
            logger.error("Refresh preach info failed: %s", e)
            messagebox.showerror("Preach Info", f"Failed to load preach info:\n{e}")

    def _selected_preach_row(self) -> Optional[Dict[str, Any]]:
        sel = self.preach_tree.selection()
        if not sel:
            return None
        idx = self.preach_tree.index(sel[0])
        if idx < 0 or idx >= len(self._preach_rows):
            return None
        return self._preach_rows[idx]

    def _add_preach(self):
        PreachInfoEditorDialog(self.master, on_save=self._on_preach_created).show()

    def _edit_preach(self):
        row = self._selected_preach_row()
        if not row:
            messagebox.showwarning("Preach Info", "Select a record to edit.")
            return
        PreachInfoEditorDialog(self.master, on_save=self._on_preach_updated, seed=row).show()

    def _delete_preach(self):
        row = self._selected_preach_row()
        if not row:
            return
        label = row.get("title") or row.get("name") or f"ID {row.get('id')}"
        if messagebox.askyesno("Delete", f"Delete preach info '{label}'?"):
            try:
                self.preach_db.delete_entry(int(row.get("id")))
                self.refresh_preach_list()
            except Exception as e:
                messagebox.showerror("Delete", f"Failed to delete record:\n{e}")

    def _on_preach_created(self, payload: Dict[str, Any]):
        self.preach_db.create_entry(payload)
        self.refresh_preach_list()

    def _on_preach_updated(self, payload: Dict[str, Any]):
        row_id = int(payload.get("id"))
        self.preach_db.update_entry(row_id, payload)
        self.refresh_preach_list()

    def _show_preach_field(self, field_name: str):
        row = self._selected_preach_row()
        if not row:
            messagebox.showwarning("Preach Info", "Select one record first.")
            return
        text = (row.get(field_name) or "").strip()
        if not text:
            messagebox.showwarning("Preach Info", f"Selected record has no {field_name}.")
            return
        if callable(self.action_callback):
            try:
                self.action_callback(("set_lyrics_text", text))
                self.action_callback("show_lyrics")
            except Exception as e:
                messagebox.showerror("Overlay", f"Failed to show overlay text:\n{e}")

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

    # -------------------
    # LED + connection updates (called from main)
    # -------------------
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

    # -------------------
    # Settings
    # -------------------
    def open_settings_dialog(self):
        SettingsDialog(self.master, self.config, self.discoverer, self._apply_settings).show()

    def _apply_settings(self, new_settings: Dict[str, Any]):
        self.config["settings"] = new_settings
        save_config(self.config)


# =======================
# Preach Info Editor
# =======================
class PreachInfoEditorDialog:
    def __init__(
        self,
        parent,
        on_save: Callable[[Dict[str, Any]], None],
        seed: Optional[Dict[str, Any]] = None,
    ):
        self.parent = parent
        self.on_save = on_save
        self.seed = seed or {}
        self.window: Optional[tk.Toplevel] = None

    def show(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Preach Info")
        self.window.geometry("620x470")
        self.window.transient(self.parent)
        self.window.grab_set()

        frm = ttk.Frame(self.window, padding=12)
        frm.pack(fill="both", expand=True)

        self.name_var = tk.StringVar(value=self.seed.get("name", ""))
        self.title_var = tk.StringVar(value=self.seed.get("title", ""))
        self.scriptures_var = tk.StringVar(value=self.seed.get("scriptures", ""))
        self.subjects_var = tk.StringVar(value=self.seed.get("subjects", ""))

        ttk.Label(frm, text="Preacher Name:").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.name_var, width=50).grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Title:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.title_var, width=50).grid(row=1, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Scriptures:").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.scriptures_var, width=50).grid(row=2, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Subjects:").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.subjects_var, width=50).grid(row=3, column=1, sticky="ew", padx=8)

        ttk.Label(frm, text="Inspirations:").grid(row=4, column=0, sticky="nw", pady=6)
        self.inspirations_text = tk.Text(frm, height=10, width=52, wrap="word", font=("Segoe UI", 10))
        self.inspirations_text.grid(row=4, column=1, sticky="nsew", padx=8, pady=6)
        self.inspirations_text.insert("1.0", self.seed.get("inspirations", ""))

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Save", command=self._save, bootstyle=SUCCESS).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.window.destroy).pack(side="left")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(4, weight=1)

    def _save(self):
        try:
            payload = {
                "name": (self.name_var.get() or "").strip(),
                "title": (self.title_var.get() or "").strip(),
                "scriptures": (self.scriptures_var.get() or "").strip(),
                "inspirations": (self.inspirations_text.get("1.0", "end-1c") or "").strip(),
                "subjects": (self.subjects_var.get() or "").strip(),
            }
            if self.seed.get("id") is not None:
                payload["id"] = int(self.seed["id"])
            if not any(payload.get(k) for k in ("name", "title", "scriptures", "inspirations", "subjects")):
                raise ValueError("Please fill at least one field.")
            self.on_save(payload)
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Preach Info", str(e))


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
        self.vmix_api_var = tk.StringVar(value=s.get("vmix_api_url", "http://localhost:8088/api"))
        self.openlp_ws_var = tk.StringVar(value=s.get("openlp_ws_url", "ws://localhost:4317"))
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

        # connections viewer (imported JSON lives here until Save)
        self.connections: List[Dict[str, Any]] = list(s.get("connections", []))

    def show(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Settings")
        self.window.geometry("800x740")
        self.window.transient(self.parent)
        self.window.grab_set()

        main = ttk.Frame(self.window, padding=12)
        main.pack(fill="both", expand=True)

        # vMix
        ttk.Label(main, text="vMix Settings", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        ttk.Label(main, text="vMix API URL:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.vmix_api_var, width=44).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(main, text="vMix Input:").grid(row=2, column=0, sticky="w", pady=5)
        self.input_combo = ttk.Combobox(main, textvariable=self.input_var, state="readonly", width=30)
        self.input_combo.grid(row=2, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="Discover", command=self._discover_inputs).grid(row=2, column=2, padx=6)

        ttk.Label(main, text="vMix Field:").grid(row=3, column=0, sticky="w", pady=5)
        self.field_combo = ttk.Combobox(main, textvariable=self.field_var, state="readonly", width=30)
        self.field_combo.grid(row=3, column=1, sticky="ew", padx=6)

        # OpenLP
        ttk.Label(main, text="OpenLP Settings", font=("Segoe UI", 12, "bold")).grid(row=4, column=0, columnspan=4, sticky="w", pady=(18, 10))
        ttk.Label(main, text="OpenLP WS URL:").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.openlp_ws_var, width=44).grid(row=5, column=1, sticky="ew", padx=6)

        # API
        ttk.Label(main, text="API Settings", font=("Segoe UI", 12, "bold")).grid(row=6, column=0, columnspan=4, sticky="w", pady=(18, 10))
        ttk.Label(main, text="LyriSync+ API Port:").grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.api_port_var, width=10).grid(row=7, column=1, sticky="w", padx=6)

        # Overlay
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

        # JSON / Quick Add
        sep = ttk.Separator(main); sep.grid(row=19, column=0, columnspan=4, sticky="ew", pady=(14, 8))
        ttk.Label(main, text="Multi-Connection", font=("Segoe UI", 12, "bold")).grid(row=20, column=0, columnspan=4, sticky="w", pady=(0, 8))
        self.conn_info_var = tk.StringVar(value=self._connections_summary())
        ttk.Label(main, textvariable=self.conn_info_var).grid(row=21, column=0, columnspan=2, sticky="w")
        btn_group = ttk.Frame(main); btn_group.grid(row=21, column=2, columnspan=2, sticky="e")
        ttk.Button(btn_group, text="Quick Add Connection…", command=self._quick_add_connection).pack(side="left", padx=(0,6))
        ttk.Button(btn_group, text="Import JSON…", command=self._import_json, bootstyle=INFO).pack(side="left")

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

    # ---- Quick Add + JSON import
    def _connections_summary(self) -> str:
        n = len(self.connections)
        return f"Imported connections: {n}" if n else "No imported connections."

    def _quick_add_connection(self):
        def _on_save(new_conn: Dict[str, Any]):
            self.connections.append(new_conn)
            self.conn_info_var.set(self._connections_summary())
            messagebox.showinfo("Connection", f"Added “{new_conn.get('name','Connection')}”. Save settings to persist.")
        ConnectionEditorDialog(self.window, on_save=_on_save).show()

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

    # ---- Save settings
    def _save_settings(self):
        try:
            new_settings = {
                "vmix_api_url": (self.vmix_api_var.get().strip() or "http://localhost:8088/api"),
                "openlp_ws_url": (self.openlp_ws_var.get().strip() or "ws://localhost:4317"),
                "api_port": max(1024, min(65535, int(self.api_port_var.get().strip() or "5000"))),
                "vmix_title_input": (self.input_var.get().strip() or "SongTitle"),
                "vmix_title_field": (self.field_var.get().strip() or "Message.Text"),
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
                # bring in any quick-added / imported connections
                "connections": list(self.connections),
                "preach_db_path": self.config.get("settings", {}).get("preach_db_path", "lyrisync_preach.db"),
            }
        except ValueError as e:
            messagebox.showerror("Settings", f"Invalid numeric value:\n{e}")
            return

        try:
            self.on_apply(new_settings)
            messagebox.showinfo("Settings", "Settings saved.")
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Settings", f"Failed to apply settings:\n{e}")
