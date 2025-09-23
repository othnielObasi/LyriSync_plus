# main.py
import asyncio
import threading
import time
from typing import Dict, Any, List, Tuple, Optional

from flask import Flask, request, jsonify
import ttkbootstrap as tb

from gui_manager import LyriSyncGUI, load_config, save_config
from vmix_openlp_handler import VmixController, OpenLPController
from splash_screen import show_splash

# -------------------------
# App state
# -------------------------
app_state = {
    "lyrics": "",
    "overlay_on": False,
    "recording": False,
    "connections_ok": 0,   # how many OpenLP sockets are connected
}
last_lyrics_ts = 0.0

settings: Dict[str, Any] = {}
gui_ref: Optional[LyriSyncGUI] = None
shared_lock = threading.Lock()
shutdown_event = threading.Event()

# Runtime connection bundles
class ConnBundle:
    def __init__(self, name: str, openlp_ip: str, ws_port: int, vmix_api_url: str,
                 mappings: List[Dict[str, str]]):
        self.name = name
        self.openlp_ip = openlp_ip
        self.ws_port = ws_port
        self.vmix_api_url = vmix_api_url
        self.mappings = mappings  # list of {"input": "...", "field": "..."}

        self.vmix = VmixController(api_url=vmix_api_url)
        self.openlp = OpenLPController(ws_url=f"ws://{openlp_ip}:{ws_port}")

# List of all configured connections
connections: List[ConnBundle] = []

# -------------------------
# Helpers
# -------------------------
def soft_wrap(text: str, max_chars: int = 36) -> str:
    """Soft-wrap text into up to 2 lines around max_chars per line."""
    if not text or max_chars <= 0:
        return text or ""
    words = text.strip().split()
    if not words:
        return ""
    line1, line2 = "", ""
    for w in words:
        cand = (line1 + " " + w).strip() if line1 else w
        if len(cand) <= max_chars or not line1:
            line1 = cand
        else:
            cand2 = (line2 + " " + w).strip() if line2 else w
            line2 = cand2
    return line1 if not line2 else (line1 + "\n" + line2)

async def send_to_all_vmix(text: str):
    """Send wrapped/uppercased text to every mapped vMix field across all connections."""
    max_chars = int(settings.get("max_chars_per_line", 36))
    text_u = (text or "").upper()
    payload = soft_wrap(text_u, max_chars=max_chars)

    tasks = []
    for bundle in connections:
        for m in bundle.mappings:
            inp = m.get("input") or settings.get("vmix_title_input", "SongTitle")
            field = m.get("field") or settings.get("vmix_title_field", "Message.Text")
            tasks.append(bundle.vmix.send_title_text(inp, field, payload))
    if tasks:
        await asyncio.gather(*tasks, return_exceptions=True)

async def clear_all_vmix():
    tasks = []
    for bundle in connections:
        for m in bundle.mappings:
            inp = m.get("input") or settings.get("vmix_title_input", "SongTitle")
            field = m.get("field") or settings.get("vmix_title_field", "Message.Text")
            tasks.append(bundle.vmix.send_title_text(inp, field, ""))
    if tasks:
        await asyncio.gather(*tasks, return_exceptions=True)

async def overlay_action(action: str):
    """Perform overlay action using the first connection's vMix (global overlay)."""
    if not connections:
        return
    ch = int(settings.get("overlay_channel", 1))
    try:
        await connections[0].vmix.trigger_overlay(ch, action=action)
    except Exception:
        pass

async def ensure_always_on_overlay():
    try:
        if settings.get("overlay_always_on", False):
            await overlay_action("On")
    except Exception:
        pass

def update_lyrics(text: str):
    global last_lyrics_ts
    with shared_lock:
        app_state["lyrics"] = text or ""
        last_lyrics_ts = time.time()

# -------------------------
# Actions
# -------------------------
async def handle_action(action):
    """Action router (supports tuple ('set_lyrics_text', text))."""
    if isinstance(action, tuple) and action[0] == "set_lyrics_text":
        update_lyrics(action[1])
        return

    if action == "show_lyrics":
        with shared_lock:
            text = app_state["lyrics"]
        await send_to_all_vmix(text)
        if settings.get("overlay_always_on", False):
            await overlay_action("On")
        elif settings.get("auto_overlay_on_send", True):
            await overlay_action("In")

    elif action == "clear_lyrics":
        await clear_all_vmix()
        if settings.get("overlay_always_on", False):
            pass  # overlay remains on; text is blanked
        elif settings.get("auto_overlay_out_on_clear", True):
            await overlay_action("Out")

    elif action == "toggle_overlay":
        await overlay_action("In")  # vMix treats OverlayInputX as toggle-ish on "In"

    elif action == "start_recording":
        if connections:
            await connections[0].vmix.start_recording()
        with shared_lock:
            app_state["recording"] = True
        if gui_ref:
            gui_ref.thread_safe(gui_ref.set_recording, True)

    elif action == "stop_recording":
        if connections:
            await connections[0].vmix.stop_recording()
        with shared_lock:
            app_state["recording"] = False
        if gui_ref:
            gui_ref.thread_safe(gui_ref.set_recording, False)

# -------------------------
# Flask API (broadcast to all connections)
# -------------------------
api = Flask(__name__)

@api.route("/api/show_lyrics", methods=["POST"])
def api_show():
    data = request.json or {}
    if "text" in data:
        update_lyrics(str(data.get("text") or ""))
    asyncio.run_coroutine_threadsafe(handle_action("show_lyrics"), loop)
    return jsonify(success=True)

@api.route("/api/clear_lyrics", methods=["POST"])
def api_clear():
    asyncio.run_coroutine_threadsafe(handle_action("clear_lyrics"), loop)
    return jsonify(success=True)

@api.route("/api/toggle_overlay", methods=["POST"])
def api_overlay():
    asyncio.run_coroutine_threadsafe(handle_action("toggle_overlay"), loop)
    return jsonify(success=True)

@api.route("/api/start_recording", methods=["POST"])
def api_start():
    asyncio.run_coroutine_threadsafe(handle_action("start_recording"), loop)
    return jsonify(success=True)

@api.route("/api/stop_recording", methods=["POST"])
def api_stop():
    asyncio.run_coroutine_threadsafe(handle_action("stop_recording"), loop)
    return jsonify(success=True)

@api.route("/api/status")
def api_status():
    with shared_lock:
        return jsonify(app_state)

def run_api():
    port = int(settings.get("api_port", 5000))
    print(f"[API] Running on http://0.0.0.0:{port}")
    api.run(host="0.0.0.0", port=port, threaded=True)

# -------------------------
# Watchers
# -------------------------
async def idle_watcher():
    """Auto-clear after idle seconds."""
    global last_lyrics_ts
    while not shutdown_event.is_set():
        try:
            idle = int(settings.get("auto_clear_idle_sec", 0))
            cur_ts = last_lyrics_ts
            if idle > 0 and cur_ts and (time.time() - cur_ts) >= idle:
                await handle_action("clear_lyrics")
                last_lyrics_ts = 0.0
        except Exception:
            pass
        await asyncio.sleep(1)

async def health_watcher():
    """Aggregate connection LEDs: any OpenLP connected? vMix reachable?"""
    while not shutdown_event.is_set():
        vmix_ok = False
        openlp_ok_count = 0

        # vMix check = first connection (or False if none)
        try:
            if connections:
                status = await connections[0].vmix.get_status()
                vmix_ok = bool(status)
        except Exception:
            vmix_ok = False

        # OpenLP: count running threads
        for b in connections:
            try:
                openlp_ok = b.openlp.running and (b.openlp._thread and b.openlp._thread.is_alive())
                if openlp_ok:
                    openlp_ok_count += 1
            except Exception:
                pass

        with shared_lock:
            app_state["connections_ok"] = openlp_ok_count

        if gui_ref:
            gui_ref.thread_safe(gui_ref.set_conn_status, vmix_ok, openlp_ok_count > 0)

        await asyncio.sleep(max(1, int(settings.get("poll_interval_sec", 2))))

async def poll_status():
    """Record/Overlay LEDs from first vMix connection, if any."""
    while not shutdown_event.is_set():
        try:
            if connections:
                st = await connections[0].vmix.get_status() or {}
                rec = str(st.get("recording", "")).lower() == "true"
                ov = str(st.get("overlay1", "")).lower() == "true"
                with shared_lock:
                    app_state["recording"] = rec
                    app_state["overlay_on"] = ov
                if gui_ref:
                    gui_ref.thread_safe(gui_ref.set_recording, rec)
                    gui_ref.thread_safe(gui_ref.set_overlay, ov)
        except Exception:
            pass
        await asyncio.sleep(int(settings.get("poll_interval_sec", 2)))

# -------------------------
# OpenLP Event wiring
# -------------------------
def make_openlp_callbacks(bundle: ConnBundle):
    def on_connect():
        if gui_ref:
            gui_ref.thread_safe(gui_ref.set_conn_status, None, True)

    def on_disconnect():
        # health_watcher recalculates aggregate; nothing extra required
        if gui_ref:
            gui_ref.thread_safe(gui_ref.set_conn_status, None, False)

    def on_new_lyrics(payload: Tuple[str, bool]):
        txt = payload[0] if isinstance(payload, tuple) else str(payload)
        is_blank = bool(payload[1]) if isinstance(payload, tuple) else (txt.strip() == "")
        update_lyrics(txt)
        if is_blank and settings.get("clear_on_blank", True):
            asyncio.run_coroutine_threadsafe(handle_action("clear_lyrics"), loop)
        else:
            asyncio.run_coroutine_threadsafe(handle_action("show_lyrics"), loop)

    bundle.openlp.on_connect = on_connect
    bundle.openlp.on_disconnect = on_disconnect
    bundle.openlp.on_new_lyrics = on_new_lyrics

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    # Load config and build connections
    config = load_config()
    settings = config.get("settings", {}) or {}

    # Build connection bundles from GUI-managed list
    connections = []
    for c in (config.get("connections") or []):
        mappings = c.get("mappings") or []
        bundle = ConnBundle(
            name=c.get("name") or "Connection",
            openlp_ip=c.get("openlp_ip", "127.0.0.1"),
            ws_port=int(c.get("ws_port", 4317)),
            vmix_api_url=c.get("vmix_api_url", settings.get("vmix_api_url", "http://localhost:8088/api")),
            mappings=mappings,
        )
        make_openlp_callbacks(bundle)
        connections.append(bundle)

    # Fallback: if no connections defined, create one from legacy settings
    if not connections:
        legacy_ws = settings.get("openlp_ws_url", "ws://localhost:4317")
        # extract host/port
        import re
        m = re.match(r"^ws://([^:/]+):(\d+)", legacy_ws.strip())
        host = m.group(1) if m else "localhost"
        port = int(m.group(2)) if m else 4317
        mappings = [{
            "input": settings.get("vmix_title_input", "SongTitle"),
            "field": settings.get("vmix_title_field", "Message.Text")
        }]
        bundle = ConnBundle(
            name="Default",
            openlp_ip=host,
            ws_port=port,
            vmix_api_url=settings.get("vmix_api_url", "http://localhost:8088/api"),
            mappings=mappings,
        )
        make_openlp_callbacks(bundle)
        connections.append(bundle)

    # Start all OpenLP listeners
    for b in connections:
        b.openlp.start()

    # GUI
    chosen_theme = (config.get("ui", {}) or {}).get("theme", "darkly")
    root = tb.Window(themename=chosen_theme)
    gui_ref = LyriSyncGUI(root, config, save_config,
                          action_callback=lambda a: asyncio.run_coroutine_threadsafe(handle_action(a), loop))

    # Splash (optional)
    if bool(settings.get("splash_enabled", True)):
        try:
            show_splash("splash.png", duration_ms=1600)
        except Exception:
            pass

    # Async loop + background tasks
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    loop.run_until_complete(ensure_always_on_overlay())
    threading.Thread(target=run_api, daemon=True).start()
    asyncio.run_coroutine_threadsafe(idle_watcher(), loop)
    asyncio.run_coroutine_threadsafe(health_watcher(), loop)
    asyncio.run_coroutine_threadsafe(poll_status(), loop)

    # Graceful shutdown
    def on_closing():
        shutdown_event.set()
        for b in connections:
            try:
                b.openlp.stop()
            except Exception:
                pass
        # close aiohttp sessions
        async def _close_all():
            for b in connections:
                try:
                    await b.vmix.close()
                except Exception:
                    pass
        loop.run_until_complete(_close_all())
        try:
            loop.stop()
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    try:
        root.mainloop()
    finally:
        try:
            loop.close()
        except Exception:
            pass
