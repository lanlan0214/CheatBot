# -*- coding: utf-8 -*-
import ctypes, time, sys
import win32api, win32gui, win32con

if not ctypes.windll.shell32.IsUserAnAdmin():
    print("請用管理員身分執行！"); input(); sys.exit(1)
print("✅ 管理員確認\n")

# 列出視窗
wins = []
def _e(h, _):
    if win32gui.IsWindowVisible(h):
        t = win32gui.GetWindowText(h)
        if t: wins.append((h, t))
win32gui.EnumWindows(_e, None)
for i,(h,t) in enumerate(wins):
    print(f"  [{i:2d}] {h:#010x}  {t}")

raw = input("\n選視窗編號或HWND hex > ").strip()
hwnd = int(raw,16) if raw.lower().startswith("0x") else wins[int(raw)][0]
print(f"目標: {win32gui.GetWindowText(hwnd)}")

# SendInput 結構
class KI(ctypes.Structure):
    _fields_=[('wVk',ctypes.c_ushort),('wScan',ctypes.c_ushort),
              ('dwFlags',ctypes.c_uint),('time',ctypes.c_uint),('dwExtraInfo',ctypes.c_uint64)]
class INP(ctypes.Structure):
    _fields_=[('type',ctypes.c_uint),('ki',KI)]

KEYUP   = 0x0002
SCAN    = 0x0008
EXT     = 0x0001
VK_X    = 0x58
SCAN_X  = ctypes.windll.user32.MapVirtualKeyW(VK_X, 0)  # ==> 0x2D

def si(vk=0, scan=0, flags=0, up=False):
    f = flags | (KEYUP if up else 0)
    i = INP(1, KI(vk, scan, f, 0, 0))
    return ctypes.windll.user32.SendInput(1, ctypes.byref(i), ctypes.sizeof(INP))

def ke(up=False):
    win32api.keybd_event(VK_X, 0, win32con.KEYEVENTF_KEYUP if up else 0, 0)

def ke_scan(up=False):
    scan = win32api.MapVirtualKey(VK_X, 0)
    win32api.keybd_event(VK_X, scan, win32con.KEYEVENTF_KEYUP if up else 0, 0)

# 切焦點
def focus():
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.4)
        ok = win32gui.GetForegroundWindow()==hwnd
        print(f"  焦點: {'OK' if ok else 'NG'}")
    except Exception as e:
        print(f"  焦點失敗: {e}")

# 倒數
print("\n5秒後開始測試，請縮小cmd讓遊戲可見")
for i in range(5,0,-1): print(f"  {i}..."); time.sleep(1)

METHODS = [
    ("M1  keybd_event (無scan)",
        lambda: ke(False), lambda: ke(True)),
    ("M2  keybd_event (有scan碼)",
        lambda: ke_scan(False), lambda: ke_scan(True)),
    ("M3  SendInput VK only",
        lambda: si(vk=VK_X), lambda: si(vk=VK_X, up=True)),
    ("M4  SendInput SCAN only",
        lambda: si(scan=SCAN_X, flags=SCAN), lambda: si(scan=SCAN_X, flags=SCAN, up=True)),
    ("M5  SendInput VK+SCAN",
        lambda: si(vk=VK_X, scan=SCAN_X), lambda: si(vk=VK_X, scan=SCAN_X, up=True)),
    ("M6  SendInput VK+SCAN+EXT",
        lambda: si(vk=VK_X, scan=SCAN_X, flags=EXT), lambda: si(vk=VK_X, scan=SCAN_X, flags=EXT, up=True)),
    ("M7  SendInput SCAN+EXT",
        lambda: si(scan=SCAN_X, flags=SCAN|EXT), lambda: si(scan=SCAN_X, flags=SCAN|EXT, up=True)),
]

for label, dn, up in METHODS:
    print(f"\n{'='*50}")
    print(f"{label}  (按3次X，每次間隔1秒)")
    focus()
    time.sleep(0.3)
    for n in range(3):
        print(f"  X {n+1}...")
        dn()
        time.sleep(0.15)
        up()
        time.sleep(1.0)
    ans = input("  有效果嗎？(y/n/skip) > ").strip().lower()
    if ans == 'y':
        print(f"\n✅ {label} 有效！終止測試。")
        break
    time.sleep(0.5)
else:
    print("\n全部方法測試完畢")

input("\n按Enter結束")
