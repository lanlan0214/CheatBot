"""
獨立按鍵測試腳本
=====================
執行後會：
1. 印出是否以管理員身分執行
2. 印出找到的所有視窗標題（幫你確認目標視窗名稱）
3. 等 4 秒讓你自己點一下遊戲視窗
4. 然後試 4 種方式各按 LEFT 3 秒，看哪種有效

用法（必須要以管理員身分執行！）：
    以系統管理員身分開啟 cmd，然後：
    cd /d "C:\\Users\\frank\\OneDrive\\桌面\\測試左右鍵\\desktop-automation-app"
    C:\\Users\\frank\\OneDrive\\桌面\\測試左右鍵\\venv\\Scripts\\python.exe test_keys.py
"""

import ctypes
import ctypes.wintypes
import time
import sys
import win32api
import win32gui
import win32con

# ── 檢查管理員權限 ─────────────────────────────────────────────────────────────
is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
print("=" * 60)
if is_admin:
    print("✅ 目前以【系統管理員】身分執行 — 可以對高權限視窗送鍵")
else:
    print("❌ 目前以【一般使用者】身分執行！")
    print("   楓之谷以管理員跑，普通權限會被 UIPI 封鎖 SendInput/PostMessage")
    print("   請關掉這個視窗，改用「以系統管理員身分開啟」的 cmd 再試一次！")
    print()
    print("   方法：在開始功能表搜尋 cmd，右鍵 → 以系統管理員身分執行")
    print("=" * 60)
    input("按 Enter 退出...")
    sys.exit(1)
print("=" * 60)

# ── 列出所有可見視窗 ──────────────────────────────────────────────────────────
print("=" * 60)
print("目前所有可見視窗標題：")
titles = []
def _enum(hwnd, _):
    if win32gui.IsWindowVisible(hwnd):
        t = win32gui.GetWindowText(hwnd)
        if t:
            titles.append((hwnd, t))
win32gui.EnumWindows(_enum, None)
for hwnd, t in titles:
    print(f"  HWND={hwnd:#010x}  {t}")
print("=" * 60)

# ── 自動找楓之谷視窗 ──────────────────────────────────────────────────────────
target_hwnd = None
SEARCH_KEYWORDS = ["楓之谷", "maplestory", "maple", "別桀谷", "冒險島"]
for hwnd, title in titles:
    if any(k.lower() in title.lower() for k in SEARCH_KEYWORDS):
        target_hwnd = hwnd
        print(f"\n✅ 自動偵測到目標視窗：HWND={hwnd:#010x}  《{title}》")
        break

if target_hwnd is None:
    print("\n⚠️  找不到楓之谷視窗，請手動輸入 HWND（十六進位，如 0x00123456）:")
    raw = input("HWND > ").strip()
    target_hwnd = int(raw, 16) if raw.startswith("0x") else int(raw)

VK_LEFT  = 0x25   # Virtual Key: LEFT arrow
VK_RIGHT = 0x27   # Virtual Key: RIGHT arrow
HOLD_SECONDS = 3  # 每種方法測試秒數

# ── SendInput helpers ─────────────────────────────────────────────────────────
_KEYEVENTF_SCANCODE = 0x0008
_KEYEVENTF_KEYUP    = 0x0002
_INPUT_KEYBOARD     = 1

class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk',         ctypes.c_ushort),
        ('wScan',       ctypes.c_ushort),
        ('dwFlags',     ctypes.c_uint),
        ('time',        ctypes.c_uint),
        ('dwExtraInfo', ctypes.c_uint64),
    ]

class _INPUT(ctypes.Structure):
    _fields_ = [
        ('type', ctypes.c_uint),
        ('ki',   _KEYBDINPUT),
    ]

def send_input_scancode(vk: int, keyup: bool = False) -> bool:
    try:
        scan = ctypes.windll.user32.MapVirtualKeyW(vk, 0)
        flags = _KEYEVENTF_SCANCODE | (_KEYEVENTF_KEYUP if keyup else 0)
        inp = _INPUT(
            type=_INPUT_KEYBOARD,
            ki=_KEYBDINPUT(wVk=0, wScan=scan, dwFlags=flags, time=0, dwExtraInfo=0),
        )
        return ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT)) == 1
    except Exception as e:
        print(f"  SendInput 例外: {e}")
        return False

def send_input_vk(vk: int, keyup: bool = False) -> bool:
    """用 VK code（非掃描碼）的 SendInput"""
    try:
        flags = _KEYEVENTF_KEYUP if keyup else 0
        inp = _INPUT(
            type=_INPUT_KEYBOARD,
            ki=_KEYBDINPUT(wVk=vk, wScan=0, dwFlags=flags, time=0, dwExtraInfo=0),
        )
        return ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT)) == 1
    except Exception as e:
        print(f"  SendInput VK 例外: {e}")
        return False

def post_message_key(hwnd: int, vk: int, keyup: bool = False):
    """背景模式：PostMessage 直接送進視窗，不需要焦點"""
    scan = win32api.MapVirtualKey(vk, 0)
    try:
        if keyup:
            lparam = (1 | (scan << 16) | (1 << 30) | (1 << 31))
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, lparam)
        else:
            lparam = (1 | (scan << 16))
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, lparam)
    except Exception as e:
        print(f"  PostMessage 失敗: {e}")
        print("  → 需要管理員權限才能 PostMessage 給高權限視窗")

def focus_window(hwnd: int):
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.2)
        active = win32gui.GetForegroundWindow()
        if active == hwnd:
            print(f"  ✅ 視窗已取得焦點")
        else:
            print(f"  ⚠️  焦點視窗是 {active:#010x}，不是目標 {hwnd:#010x}")
    except Exception as e:
        print(f"  SetForegroundWindow 失敗: {e}")

def hold_key(press_fn, release_fn, seconds: float, label: str):
    print(f"\n  [按下 {label}] {seconds}秒...")
    press_fn()
    time.sleep(seconds)
    release_fn()
    print(f"  [放開 {label}]")

# ── 主測試 ────────────────────────────────────────────────────────────────────
print(f"\n⏳ 4 秒後開始測試，請把 cmd/終端機縮小，讓楓之谷視窗可以看到！")
for i in range(4, 0, -1):
    print(f"   {i}...")
    time.sleep(1)

# ===== 方法 A：SendInput + 掃描碼（焦點模式）=====
print("\n" + "=" * 60)
print("測試 A：SendInput + 掃描碼（先切焦點到遊戲）")
focus_window(target_hwnd)
hold_key(
    lambda: send_input_scancode(VK_LEFT, False),
    lambda: send_input_scancode(VK_LEFT, True),
    HOLD_SECONDS,
    "LEFT (SendInput scancode)"
)
time.sleep(1)

# ===== 方法 B：SendInput + VK code（焦點模式）=====
print("\n" + "=" * 60)
print("測試 B：SendInput + VK code（先切焦點到遊戲）")
focus_window(target_hwnd)
hold_key(
    lambda: send_input_vk(VK_LEFT, False),
    lambda: send_input_vk(VK_LEFT, True),
    HOLD_SECONDS,
    "LEFT (SendInput VK)"
)
time.sleep(1)

# ===== 方法 C：PostMessage（背景模式，不需焦點）=====
print("\n" + "=" * 60)
print("測試 C：PostMessage WM_KEYDOWN（不切焦點，背景模式）")
print(f"  目標 HWND: {target_hwnd:#010x}")
# 持續送 WM_KEYDOWN（模擬長按）
print(f"  [開始連送 WM_KEYDOWN {HOLD_SECONDS}秒]...")
end_t = time.time() + HOLD_SECONDS
post_message_key(target_hwnd, VK_LEFT, keyup=False)
while time.time() < end_t:
    time.sleep(0.05)
post_message_key(target_hwnd, VK_LEFT, keyup=True)
print(f"  [放開 WM_KEYUP]")
time.sleep(1)

# ===== 方法 D：keybd_event（舊版，對照組）=====
print("\n" + "=" * 60)
print("測試 D：keybd_event（舊方法，對照用）")
focus_window(target_hwnd)
hold_key(
    lambda: win32api.keybd_event(VK_LEFT, 0, 0, 0),
    lambda: win32api.keybd_event(VK_LEFT, 0, win32con.KEYEVENTF_KEYUP, 0),
    HOLD_SECONDS,
    "LEFT (keybd_event)"
)
time.sleep(1)

# ===== 方法 E：單獨按技能鍵 X（SendInput 掃描碼）=====
VK_X = 0x58
print("\n" + "=" * 60)
print("測試 E：雨技能鍵 X 單独 SendInput scancode（再按 5 次，每次間隔0.5秒）")
focus_window(target_hwnd)
for i in range(5):
    print(f"  X 第 {i+1} 次...")
    ok = send_input_scancode(VK_X, False)
    time.sleep(0.1)
    send_input_scancode(VK_X, True)
    time.sleep(0.5)
print("測試 E 完成")
time.sleep(1)

# ===== 方法 F：單獨按技能鍵 X（keybd_event）=====
print("\n" + "=" * 60)
print("測試 F：雨技能鍵 X 單獨 keybd_event（再按 5 次，每次間隔0.5秒）")
focus_window(target_hwnd)
for i in range(5):
    print(f"  X 第 {i+1} 次...")
    win32api.keybd_event(VK_X, 0, 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(VK_X, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
print("測試 F 完成")
time.sleep(1)

# ===== 方法 G：檡擬實際宏：按LEFT 2秒 → 放開 → 按X → 再按LEFT =====
print("\n" + "=" * 60)
print("測試 G：[LEFT 2秒] → [放，等150ms] → [X按下120ms再放] → [等150ms] → [LEFT 2秒], 共 3 循環")
focus_window(target_hwnd)
for cycle in range(3):
    print(f"  循環 {cycle+1}: 按 LEFT 2 秒...")
    win32api.keybd_event(VK_LEFT, 0, 0, 0)
    time.sleep(2)
    win32api.keybd_event(VK_LEFT, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.15)   # 人類放開方向鍵後的自然延遲
    print(f"  循環 {cycle+1}: 按 X...")
    win32api.keybd_event(VK_X, 0, 0, 0)
    time.sleep(0.12)   # 按鍵持續時間
    win32api.keybd_event(VK_X, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.15)   # 放開技能鍵後稍停
print("測試 G 完成")

print("\n" + "=" * 60)
print("✅ 全部測試完成！")
print("請告訴我哪個注能鍵 X 有動（E / F / G / 全部沒動）")
