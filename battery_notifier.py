import os, sys, time, threading, signal, configparser, logging
import psutil, winsound, msvcrt
from winotify import Notification
from logging.handlers import RotatingFileHandler

# ── Paths ─────────────────────────────────────────────────────────────────────
_TEMP    = os.getenv("TEMP",    os.getcwd())
_APPDATA = os.getenv("APPDATA", os.getcwd())

LOCKFILE_PATH  = os.path.join(_TEMP,    "battery_notifier.lock")
FLAG_FILE      = os.path.join(_APPDATA, "BatteryNotifierShortcutFlag.txt")
SHORTCUT_PATH  = os.path.join(_APPDATA, r"Microsoft\Windows\Start Menu\Programs\Startup\BatteryNotifier.lnk")
CONFIG_PATH    = os.path.join(_APPDATA, "BatteryNotifierConfig.ini")
LOG_PATH       = os.path.join(_APPDATA, "BatteryNotifier.log")
# ICO_PATH and ALARM_PATH are set after _resource_base() is defined below

# ── Logging (rotating, max 1MB, keep 2 backups) ───────────────────────────────
def _setup_logging() -> logging.Logger:
    logger = logging.getLogger("BatteryNotifier")
    logger.setLevel(logging.DEBUG)
    handler = RotatingFileHandler(LOG_PATH, maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(handler)
    # Also print to console when running from terminal
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(console)
    return logger

log = _setup_logging()

# ── Config — thread-safe, live-reloadable ────────────────────────────────────
class AppConfig:
    """
    Holds thresholds read from config.ini.
    Safe to read from any thread via properties.
    Call reload() to pick up file changes without restarting.
    """
    _DEFAULTS = {"low": 30, "full": 90, "repeat_minutes": 30}

    def __init__(self) -> None:
        self._lock  = threading.Lock()
        self._low   = self._DEFAULTS["low"]
        self._full  = self._DEFAULTS["full"]
        self._rpt   = self._DEFAULTS["repeat_minutes"]
        self._mtime = 0.0
        self._ensure_file()
        self.reload()

    # ── public read-only properties (no lock needed for int reads on CPython) ─
    @property
    def low(self)            -> int: return self._low
    @property
    def full(self)           -> int: return self._full
    @property
    def repeat_seconds(self) -> int: return self._rpt * 60

    def snapshot(self) -> tuple[int, int, int]:
        """
        Return (low, full, repeat_seconds) atomically under the lock.
        Use this in the monitor loop instead of reading properties one-by-one,
        so a config reload between reads can never produce an inconsistent state
        (e.g. low=35 from old config paired with full=85 from new config).
        """
        with self._lock:
            return self._low, self._full, self._rpt * 60

    def _ensure_file(self) -> None:
        """Write default config.ini if it doesn't exist yet."""
        if os.path.exists(CONFIG_PATH):
            return
        cfg = configparser.ConfigParser()
        cfg["thresholds"] = {k: str(v) for k, v in self._DEFAULTS.items()}
        with open(CONFIG_PATH, "w") as f:
            cfg.write(f)
        log.info(f"Default config created at {CONFIG_PATH}")

    def reload(self) -> bool:
        """
        Re-read config.ini. Returns True if values actually changed.
        Called at startup and by the watcher thread when mtime changes.
        """
        try:
            mtime = os.stat(CONFIG_PATH).st_mtime
        except OSError:
            log.warning("Config file missing — keeping current values")
            return False

        if mtime == self._mtime:
            return False   # file unchanged

        cfg = configparser.ConfigParser()
        cfg.read(CONFIG_PATH)
        try:
            low    = int(cfg.get("thresholds", "low",            fallback="30"))
            full   = int(cfg.get("thresholds", "full",           fallback="90"))
            repeat = int(cfg.get("thresholds", "repeat_minutes", fallback="30"))
            if not (0 < low < full < 100):
                raise ValueError(f"low={low} must be < full={full}, both in 1–99")
            if repeat < 1:
                raise ValueError(f"repeat_minutes={repeat} must be >= 1")
        except (ValueError, configparser.Error) as e:
            log.warning(f"Config parse error ({e}) — keeping current values")
            return False

        with self._lock:
            changed      = (low != self._low or full != self._full or repeat != self._rpt)
            self._low    = low
            self._full   = full
            self._rpt    = repeat
            self._mtime  = mtime

        if changed:
            log.info(f"Config reloaded → LOW={low}%  FULL={full}%  REPEAT={repeat}min")
        return changed


def _watch_config(cfg: AppConfig, stop: threading.Event) -> None:
    """Background thread: checks config.ini mtime every 15 s and reloads if changed."""
    while not stop.wait(15):
        cfg.reload()

# ── Resource path resolution (works for frozen EXE and plain .py) ────────────
def _resource_base() -> str:
    """
    Returns the directory that contains bundled resource files.
    - Frozen EXE  : sys._MEIPASS  (PyInstaller temp extraction folder)
    - Plain .py   : directory of THIS FILE — not os.getcwd(), so the app
                    works correctly when launched from a different working
                    directory (e.g. via a Windows startup shortcut).
    """
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS")
    return os.path.dirname(os.path.abspath(__file__))

_BASE      = _resource_base()
ALARM_PATH = os.path.join(_BASE, "Alarm01.wav")
ICO_PATH   = os.path.join(_BASE, "battery.ico")

# ── Module-level handle for lock file ────────────────────────────────────────
_lockfile     = None
_alarm_active = False

# ── Notification ──────────────────────────────────────────────────────────────
def show_toast(title: str, message: str) -> None:
    try:
        Notification(app_id="Battery Notifier", title=title, msg=message, duration="short").show()
        log.info(f"Toast: [{title}] {message}")
    except Exception as e:
        log.error(f"Toast failed: {e}")

# ── Single-instance lock ──────────────────────────────────────────────────────
def single_instance_lock() -> bool:
    global _lockfile
    try:
        _lockfile = open(LOCKFILE_PATH, "w")
        msvcrt.locking(_lockfile.fileno(), msvcrt.LK_NBLCK, 1)
        return True
    except OSError:
        return False

# ── Alarm helpers ─────────────────────────────────────────────────────────────
def play_alarm() -> None:
    global _alarm_active
    if _alarm_active:
        return
    # Warn if WAV is missing instead of silent failure
    if not os.path.exists(ALARM_PATH):
        log.warning(f"Alarm file not found: {ALARM_PATH}")
        show_toast("⚠️ Alarm File Missing", f"Place Alarm01.wav next to the app.")
        return
    try:
        winsound.PlaySound(ALARM_PATH, winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP)
        _alarm_active = True
        log.debug("Alarm started")
    except Exception as e:
        log.error(f"Alarm error: {e}")

def stop_alarm() -> None:
    global _alarm_active
    if not _alarm_active:
        return
    winsound.PlaySound(None, winsound.SND_PURGE)
    _alarm_active = False
    log.debug("Alarm stopped")

# ── Startup shortcut ──────────────────────────────────────────────────────────
def add_startup_once() -> None:
    if os.path.exists(FLAG_FILE):
        return
    try:
        from win32com.client import Dispatch

        frozen = getattr(sys, "frozen", False)

        if frozen:
            # PyInstaller EXE — point shortcut directly at the EXE
            target  = sys.executable
            args    = ""
            workdir = os.path.dirname(target)
            log.info("Startup shortcut mode: frozen EXE")
        else:
            # Plain .py script — launch via pythonw.exe (no console window)
            # pythonw.exe lives alongside python.exe in the same Scripts folder
            pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            if not os.path.exists(pythonw):
                # Fallback: use python.exe (shows a brief console flash)
                pythonw = sys.executable
                log.warning("pythonw.exe not found — falling back to python.exe")
            target  = pythonw
            args    = f'"{os.path.abspath(__file__)}"'
            workdir = os.path.dirname(os.path.abspath(__file__))
            log.info(f"Startup shortcut mode: script via {os.path.basename(target)}")

        shell = Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(SHORTCUT_PATH)
        sc.Targetpath       = target
        sc.Arguments        = args
        sc.WorkingDirectory = workdir
        sc.IconLocation     = target
        sc.save()

        with open(FLAG_FILE, "w") as f:
            f.write("created")
        show_toast("✅ Battery Notifier Installed", "Will run automatically on startup.")
        log.info(f"Startup shortcut created → {SHORTCUT_PATH}")
    except Exception as e:
        log.error(f"Startup shortcut failed: {e}")

# ── Shutdown event — set this to stop all threads cleanly ────────────────────
_stop = threading.Event()

# ── Cleanup ───────────────────────────────────────────────────────────────────
def cleanup() -> None:
    log.info("Cleanup called — shutting down")
    _stop.set()          # signals monitor loop and watcher thread to exit
    stop_alarm()
    try:
        if _lockfile:
            _lockfile.close()
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
        # ✅ FLAG_FILE and SHORTCUT_PATH intentionally NOT deleted
        #    so auto-start survives across restarts
    except Exception as e:
        log.error(f"Cleanup error: {e}")
    log.info("Shutdown complete")
    sys.exit(0)          # normal Python exit — runs atexit, flushes handlers

# ── Graceful Ctrl+C / SIGTERM ─────────────────────────────────────────────────
def _signal_handler(sig, frame):
    log.info(f"Signal {sig} received")
    cleanup()

signal.signal(signal.SIGINT,  _signal_handler)
signal.signal(signal.SIGTERM, _signal_handler)

# ── Native tray icon (Win32) with live battery % tooltip ─────────────────────
_tray_hwnd: int = 0

def update_tray_tooltip(percent: float, plugged: bool) -> None:
    """Update tray tooltip text with live battery status."""
    global _tray_hwnd
    if not _tray_hwnd:
        return
    try:
        import win32gui, win32con
        status = "🔌" if plugged else "🔋"
        tip    = f"Battery Notifier — {percent:.0f}% {status}"
        nid    = (
            _tray_hwnd, 0,
            win32gui.NIF_TIP,
            win32con.WM_USER + 20,
            0,
            tip
        )
        win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
    except Exception as e:
        log.debug(f"Tooltip update failed: {e}")

def native_tray() -> None:
    global _tray_hwnd
    import win32gui, win32api, win32con

    WM_TRAY  = win32con.WM_USER + 20
    CMD_EXIT = 1000

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == win32con.WM_DESTROY:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, 0))
            cleanup()
        elif msg == WM_TRAY:
            if lparam == win32con.WM_RBUTTONUP:
                menu = win32gui.CreatePopupMenu()
                win32gui.AppendMenu(menu, win32con.MF_STRING, CMD_EXIT, "Exit Battery Notifier")
                pos = win32gui.GetCursorPos()
                win32gui.SetForegroundWindow(hwnd)
                win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, hwnd, None)
            elif lparam == win32con.WM_LBUTTONDBLCLK:
                cleanup()
        elif msg == win32con.WM_COMMAND and wparam == CMD_EXIT:
            win32gui.DestroyWindow(hwnd)
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    hinst      = win32api.GetModuleHandle()
    class_name = "BatteryNotifierTray"

    wc               = win32gui.WNDCLASS()
    wc.lpfnWndProc   = wnd_proc
    wc.lpszClassName = class_name
    wc.hInstance     = hinst
    win32gui.RegisterClass(wc)

    hwnd = win32gui.CreateWindow(class_name, "BatteryNotifier", 0, 0, 0, 0, 0, 0, 0, hinst, None)
    _tray_hwnd = hwnd

    # Load custom icon if available, else fallback to system icon
    if os.path.exists(ICO_PATH):
        try:
            hicon = win32gui.LoadImage(
                0, ICO_PATH, win32con.IMAGE_ICON,
                16, 16, win32con.LR_LOADFROMFILE
            )
            log.info("Custom battery.ico loaded")
        except Exception as e:
            log.warning(f"Custom icon failed ({e}), using default")
            hicon = win32gui.LoadIcon(0, win32con.IDI_INFORMATION)
    else:
        hicon = win32gui.LoadIcon(0, win32con.IDI_INFORMATION)
        log.info("No battery.ico found — using default icon. Place battery.ico next to app for custom icon.")

    nid = (
        hwnd, 0,
        win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
        WM_TRAY, hicon, "Battery Notifier — Starting..."
    )
    win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
    log.info("Tray icon created")
    win32gui.PumpMessages()

# ── Main loop ─────────────────────────────────────────────────────────────────
def main() -> None:
    log.info("=" * 50)
    log.info("Battery Notifier starting")

    if not single_instance_lock():
        show_toast("⚠️ Already Running", "Battery Notifier is already active.")
        log.warning("Another instance detected — exiting")
        time.sleep(2)
        sys.exit(0)

    cfg = AppConfig()

    add_startup_once()

    # Start tray + config watcher as daemon threads
    threading.Thread(target=native_tray,                          daemon=True).start()
    threading.Thread(target=_watch_config, args=(cfg, _stop),     daemon=True).start()

    plugged_prev      = None
    low_on            = False
    full_on           = False
    startup_complete  = False
    low_last_alerted :float = 0.0
    full_last_alerted:float = 0.0

    # Short startup delay so tray is ready before first tooltip update
    time.sleep(1)
    startup_complete = True
    log.info("Startup complete — entering monitor loop")

    while not _stop.is_set():
        try:
            battery = psutil.sensors_battery()

            if battery is None:
                stop_alarm()
                low_on = full_on = False
                log.debug("No battery detected — sleeping 30s")
                _stop.wait(30)   # interruptible sleep — exits immediately on shutdown
                continue

            percent = battery.percent
            plugged = battery.power_plugged

            # Read all thresholds atomically in one lock acquisition
            LOW, FULL, REPEAT_SECONDS = cfg.snapshot()

            # Update live tray tooltip every loop
            update_tray_tooltip(percent, plugged)

            # ── Plug state changed ────────────────────────────────────────────
            if plugged != plugged_prev:
                stop_alarm()
                low_on = full_on = False
                low_last_alerted = full_last_alerted = 0.0
                if startup_complete:
                    if plugged:
                        show_toast("🔌 Plugged In",  f"{percent:.0f}%")
                    else:
                        show_toast("🔋 Unplugged",   f"{percent:.0f}%")
                    log.info(f"Plug state changed → {'plugged' if plugged else 'unplugged'} at {percent:.0f}%")
                plugged_prev = plugged

            now = time.monotonic()

            # ── Low battery (with repeat every REPEAT_SECONDS) ────────────────
            if not plugged and percent <= LOW:
                if not low_on or (now - low_last_alerted) >= REPEAT_SECONDS:
                    show_toast("🔴 Battery Low", f"{percent:.0f}% — Plug in charger!")
                    play_alarm()
                    low_on           = True
                    low_last_alerted = now
                    log.warning(f"Low battery alert: {percent:.0f}%")
            elif low_on and (plugged or percent > LOW):
                stop_alarm()
                low_on           = False
                low_last_alerted = 0.0
                log.info(f"Low battery cleared: {percent:.0f}%")

            # ── Battery full (with repeat every REPEAT_SECONDS) ───────────────
            if plugged and percent >= FULL:
                if not full_on or (now - full_last_alerted) >= REPEAT_SECONDS:
                    show_toast("🟢 Battery Full", f"{percent:.0f}% — Unplug charger.")
                    play_alarm()
                    full_on           = True
                    full_last_alerted = now
                    log.warning(f"Full battery alert: {percent:.0f}%")
            elif full_on and (not plugged or percent < FULL):
                stop_alarm()
                full_on           = False
                full_last_alerted = 0.0
                log.info(f"Full battery cleared: {percent:.0f}%")

        except Exception as e:
            stop_alarm()
            log.error(f"Battery loop error: {e}", exc_info=True)

        _stop.wait(10)   # interruptible — exits in <1s on shutdown instead of sleeping full 10s

    log.info("Monitor loop exited")


if __name__ == "__main__":
    main()
