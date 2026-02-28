import time
from dataclasses import dataclass
from typing import Optional

import win32api
import win32con
import win32gui


@dataclass
class TargetWindow:
    hwnd: int
    title: str


class BackgroundInput:
    """Send keystrokes to a specific window (hwnd) using Win32 messages.

    Notes:
    - Works best with classic Win32 apps that process WM_KEYDOWN/WM_KEYUP.
    - Some modern apps/games (or apps using raw input/DirectInput) may ignore these.
    """

    @staticmethod
    def find_window_by_title_substring(title_substring: str) -> Optional[TargetWindow]:
        title_substring = (title_substring or "").strip().lower()
        if not title_substring:
            return None

        matches: list[TargetWindow] = []

        def enum_handler(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
            if title_substring in title.lower():
                matches.append(TargetWindow(hwnd=hwnd, title=title))

        win32gui.EnumWindows(enum_handler, None)
        return matches[0] if matches else None

    @staticmethod
    def set_foreground(hwnd: int) -> None:
        # Best-effort: Windows may block focus stealing depending on foreground lock.
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            # Fallback: try bringing to top.
            try:
                win32gui.BringWindowToTop(hwnd)
            except Exception:
                pass

    @staticmethod
    def _vk_to_scan(vk: int) -> int:
        return win32api.MapVirtualKey(vk, 0)

    @staticmethod
    def send_vk(hwnd: int, vk: int, keyup: bool = False) -> None:
        scan = BackgroundInput._vk_to_scan(vk)
        lparam = 0x00000001 | (scan << 16)
        msg = win32con.WM_KEYUP if keyup else win32con.WM_KEYDOWN
        # 盡量用同步送入，避免某些視窗佇列忙碌時吞訊息
        try:
            win32gui.SendMessage(hwnd, msg, vk, lparam)
        except Exception:
            win32api.PostMessage(hwnd, msg, vk, lparam)

    @staticmethod
    def press_vk(hwnd: int, vk: int, hold_seconds: float = 0.0) -> None:
        BackgroundInput.send_vk(hwnd, vk, keyup=False)
        if hold_seconds and hold_seconds > 0:
            time.sleep(hold_seconds)
        BackgroundInput.send_vk(hwnd, vk, keyup=True)

    @staticmethod
    def send_text(hwnd: int, text: str, interval: float = 0.0) -> None:
        # WM_CHAR expects UTF-16 code units for BMP chars.
        # Best-effort: 喚醒目標視窗（不保證不搶焦點）
        try:
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        except Exception:
            pass

        for ch in text:
            try:
                win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)
            except Exception:
                win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)
            if interval and interval > 0:
                time.sleep(interval)

    @staticmethod
    def _pack_lparam_xy(x: int, y: int) -> int:
        # LPARAM: low word = x, high word = y (signed 16-bit)
        x = int(x) & 0xFFFF
        y = int(y) & 0xFFFF
        return x | (y << 16)

    @staticmethod
    def click(hwnd: int, x: int = 10, y: int = 10, button: str = "left") -> None:
        """Best-effort background click.

        Notes:
        - Coordinates are client-area coords relative to the target window.
        - Many games (including MapleStory) will ignore these messages.
        """
        btn = (button or "left").lower()
        if btn == "right":
            down, up = win32con.WM_RBUTTONDOWN, win32con.WM_RBUTTONUP
            wparam = win32con.MK_RBUTTON
        else:
            down, up = win32con.WM_LBUTTONDOWN, win32con.WM_LBUTTONUP
            wparam = win32con.MK_LBUTTON

        lparam = BackgroundInput._pack_lparam_xy(x, y)
        win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
        win32api.PostMessage(hwnd, down, wparam, lparam)
        win32api.PostMessage(hwnd, up, 0, lparam)

    @staticmethod
    def mouse_down(hwnd: int, x: int = 10, y: int = 10, button: str = "left") -> None:
        btn = (button or "left").lower()
        if btn == "right":
            down = win32con.WM_RBUTTONDOWN
            wparam = win32con.MK_RBUTTON
        else:
            down = win32con.WM_LBUTTONDOWN
            wparam = win32con.MK_LBUTTON
        lparam = BackgroundInput._pack_lparam_xy(x, y)
        win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
        win32api.PostMessage(hwnd, down, wparam, lparam)

    @staticmethod
    def mouse_up(hwnd: int, x: int = 10, y: int = 10, button: str = "left") -> None:
        btn = (button or "left").lower()
        if btn == "right":
            up = win32con.WM_RBUTTONUP
        else:
            up = win32con.WM_LBUTTONUP
        lparam = BackgroundInput._pack_lparam_xy(x, y)
        win32api.PostMessage(hwnd, up, 0, lparam)


VK_MAP = {
    "UP": win32con.VK_UP,
    "DOWN": win32con.VK_DOWN,
    "LEFT": win32con.VK_LEFT,
    "RIGHT": win32con.VK_RIGHT,
    "SPACE": win32con.VK_SPACE,
    "ENTER": win32con.VK_RETURN,
    "ESC": win32con.VK_ESCAPE,
    "TAB": win32con.VK_TAB,
    "SHIFT": win32con.VK_SHIFT,
    "CTRL": win32con.VK_CONTROL,
    "CONTROL": win32con.VK_CONTROL,
    "ALT": win32con.VK_MENU,
    "BACKSPACE": win32con.VK_BACK,
    "BKSP": win32con.VK_BACK,
    "DEL": win32con.VK_DELETE,
    "DELETE": win32con.VK_DELETE,
    "HOME": win32con.VK_HOME,
    "END": win32con.VK_END,
    "PGUP": win32con.VK_PRIOR,
    "PAGEUP": win32con.VK_PRIOR,
    "PGDN": win32con.VK_NEXT,
    "PAGEDOWN": win32con.VK_NEXT,
    "INSERT": win32con.VK_INSERT,
    "INS": win32con.VK_INSERT,
    "F1": win32con.VK_F1,
    "F2": win32con.VK_F2,
    "F3": win32con.VK_F3,
    "F4": win32con.VK_F4,
    "F5": win32con.VK_F5,
    "F6": win32con.VK_F6,
    "F7": win32con.VK_F7,
    "F8": win32con.VK_F8,
    "F9": win32con.VK_F9,
    "F10": win32con.VK_F10,
    "F11": win32con.VK_F11,
    "F12": win32con.VK_F12,
    # OEM 符號鍵：不同 pywin32 版本常數名字可能不存在，這裡用 getattr 安全取值。
    "-": getattr(win32con, "VK_OEM_MINUS", None),
    "=": getattr(win32con, "VK_OEM_PLUS", None),
    "[": getattr(win32con, "VK_OEM_4", None),
    "]": getattr(win32con, "VK_OEM_6", None),
    "\\": getattr(win32con, "VK_OEM_5", None),
    ";": getattr(win32con, "VK_OEM_1", None),
    "'": getattr(win32con, "VK_OEM_7", None),
    ",": getattr(win32con, "VK_OEM_COMMA", None),
    ".": getattr(win32con, "VK_OEM_PERIOD", None),
    "/": getattr(win32con, "VK_OEM_2", None),
    "`": getattr(win32con, "VK_OEM_3", None),
}


def parse_command_to_vk(cmd: str) -> Optional[int]:
    cmd = (cmd or "").strip().upper()
    if cmd in VK_MAP:
        vk = VK_MAP[cmd]
        return vk if isinstance(vk, int) else None

    if len(cmd) == 1:
        # Letters/numbers: use virtual-key code.
        ch = cmd
        if "A" <= ch <= "Z":
            return ord(ch)
        if "0" <= ch <= "9":
            return ord(ch)
    return None
