# main.py
import asyncio
import threading
import time
from typing import Tuple

import ttkbootstrap as tb
from flask import Flask, request, jsonify

from gui_manager import LyriSyncGUI, load_config, save_config
from vmix_openlp_handler import VmixController, OpenLPController

# -------------------------
# Global state & locks
# -------------------------
state = {
    "lyrics": "",
    "overlay_on": False,
    "recording": False,
}
last_lyrics_ts = 0.0
lock = threading.Lock()
shutdown_evt = threading.Event()

gui: LyriSyncGUI | None = None
vmix: VmixController | None = None
openlp: OpenLPController | None = None
settings = {}
loop: asyncio.AbstractEventLoop | None = None

# -------------------------
# Helpers
# -------------------------
def soft_wrap(text: str, max_chars: int) -> str:
    """
    Wrap to at most two lines. Respects word boundaries where possible.
    """
    text = (text or "").strip()
    if not text:
        return ""
    words = text.split()
    line1, line2 = "", ""
    for w in words:
        cand = (line1 + " " + w).strip() if line1 else w
        if len(cand) <= max_chars or not line1:
            line1 = cand
        else:
            cand2 = (line2 + " " + w).strip() if line2 else w
            line2 = cand2
    return line1 if not line2 else f"{line1}\n{line2}"

async def update_leds_from_status():
    try:
        s = await vmix.get_status()
        rec = str(s.get("recording", "")).lower() == "true"
        ov1 = str(s.get("overlay1", "")).lower() == "true"
        with lock:
            state["recording"] = rec
            state["overlay_on"] = ov1
        if gui:
            gui.thread_safe(gui.set_recording, rec)
            gui.thread_safe(gui.set_overlay, ov1)
    except Exception:
        # keep LEDs red if status fails
        pass

# -------------------------
# Action dispatcher
# -------------------------
async def handle_action(action):
    """
    Supports:
      - ("set_lyrics_text", str)
      - "show_lyrics"
      - "clear_lyrics"
      - "toggle_overlay"
      - "start_recording"
      - "stop_recording"
    """
    global last_lyrics_ts
    if action is None:
        return

    # Read settings
    title_input = settings.get("vmix_title_input", "SongTitle")
    title_field = settings.get("vmix_title_field", "Message.Text")
    ch = int(settings.get("overlay_channel", 1))
    always_on = bool(settings.get("overlay_always_on", False))
    auto_in = bool(settings.get("auto_overlay_on_send", True))
    auto_out = bool(settings.get("auto_overlay_out_on_clear", True))
    max_chars = int(settings.get("max_chars_per_line", 48))

    # Set/Send/Clear
    if isinstance(action, tuple) and action[0] == "set_lyrics_text":
        text = (action[1] or "").strip().upper()
        with lock:
            state["lyrics"] = text
            last_lyrics_ts = time.time()
        return

    if action == "show_lyrics":
        with lock:
            text = state["lyrics"]
        wrapped = soft_wrap(text, max_chars)
        await vmix.send_title_text(title_input, title_field, wrapped)
        if always_on:
            # force on
            try:
                await vmix.trigger_overlay(ch, "On")
            except Exception:
                pass
        elif auto_in:
            try:
                await vmix.trigger_overlay(ch, "In")
            except Exception:
                pass
        await update_leds_from_status()
        return

    if action == "clear_lyrics":
        await vmix.send_title_text(title_input, title_field, "")
        if not always_on and auto_out:
            try:
                await vmix.trigger_overlay(ch, "Out")
            except Exception:
                pass
        await update_leds_from_status()
        return

    if action == "toggle_overlay":
        await vmix.trigger_overlay(ch, "In")  # vMix treats repeated In/Out as toggle for title overlays
        await update_leds_from_status()
        return

    if action == "start_recording":
        await vmix.start_recording()
        await update_leds_from_status()
        return

    if action == "stop_recording":
        await vmix.stop_recording()
        await update_leds_from_status()
        return

# -------------------------
# Idle auto-clear & health
# -------------------------
async def idle_watcher():
    global last_lyrics_ts
    while not shutdown_evt.is_set():
        try:
            idle = int(settings.get("auto_clear_idle_sec", 0))
            if idle > 0:
                with lock:
                    ts = last_lyrics_ts
                if ts and (time.time() - ts) >= idle:
                    await handle_action("clear_lyrics")
                    with lock:
                        last_lyrics_ts = 0.0
        except Exception:
            pass
        await asyncio.sleep(1)

async def health_watcher():
    while not shutdown_evt.is_set():
        await update_leds_from_status()
        await asyncio.sleep(max(1, int(settings.get("poll_interval_sec", 2))))

# -------------------------
# OpenLP wiring
# -------------------------
def on_openlp_new(payload: Tuple[str, bool]):
    # payload = (text, is_blank)
    text, is_blank = payload
    text = (text or "").strip().upper()
    with lock:
        state["lyrics"] = text
        # timestamp only when not blank
        if text:
            global last_lyrics_ts
            last_lyrics_ts = time.time()

    # Drive vMix based on blank/text
    if is_blank and settings.get("clear_on_blank", True):
        asyncio.run_coroutine_threadsafe(handle_action("clear_lyrics"), loop)
    else:
        asyncio.run_coroutine_threadsafe(handle_action("show_lyrics"), loop)

def on_openlp_connect():
    if gui:
        gui.thread_safe(gui.set_conn_status, None, True)

def on_openlp_disconnect():
    if gui:
        gui.thread_safe(gui.set_conn_status, None, False)

# -------------------------
# Flask mini API (optional)
# -------------------------
api = Flask(__name__)

@api.route("/api/show_lyrics", methods=["POST"])
def api_show_lyrics():
    data = request.json or {}
    txt = str(data.get("text", "")).upper()
    with lock:
        state["lyrics"] = txt
    asyncio.run_coroutine_threadsafe(handle_action("show_lyrics"), loop)
    return jsonify(ok=True)

@api.route("/api/clear_lyrics", methods=["POST"])
def api_clear_lyrics():
    asyncio.run_coroutine_threadsafe(handle_action("clear_lyrics"), loop)
    return jsonify(ok=True)

@api.route("/api/status")
def api_status():
    with lock:
        return jsonify(state)

def run_api():
    port = int(settings.get("api_port", 5000))
    api.run(port=port, threaded=True)

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    # Load configuration
    config = load_config()
    settings = config.get("settings", {}) or {}

    # Controllers
    vmix = VmixController(api_url=settings.get("vmix_api_url", "http://localhost:8088/api"))
    openlp = OpenLPController(ws_url=settings.get("openlp_ws_url", "ws://localhost:4317"))

    # Async loop
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    # OpenLP callbacks
    openlp.on_new_lyrics = on_openlp_new
    openlp.on_connect = on_openlp_connect
    openlp.on_disconnect = on_openlp_disconnect
    openlp.start()

    # GUI
    root = tb.Window(themename=(config.get("ui", {}) or {}).get("theme", "darkly"))
    gui = LyriSyncGUI(
        root,
        config,
        save_config,
        action_callback=lambda a: asyncio.run_coroutine_threadsafe(handle_action(a), loop),
    )

    # Kick off background tasks
    asyncio.run_coroutine_threadsafe(health_watcher(), loop)
    asyncio.run_coroutine_threadsafe(idle_watcher(), loop)

    # Optional: start API server
    threading.Thread(target=run_api, daemon=True).start()

    def on_close():
        shutdown_evt.set()
        try:
            openlp.stop()
        except Exception:
            pass
        try:
            loop.call_soon_threadsafe(loop.stop)
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    try:
        root.mainloop()
    finally:
        try:
            loop.close()
        except Exception:
            pass
