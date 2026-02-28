from tkinter import (
    Tk,
    Button,
    Label,
    StringVar,
    Entry,
    Frame,
    Checkbutton,
    IntVar,
    Radiobutton,
)
from tkinter import ttk
from tkinter import filedialog, messagebox
import time
import json
import threading
import sys
import webbrowser

import pyautogui
import pygetwindow
import win32gui
import win32con
import keyboard

from core.win_background_input import BackgroundInput, parse_command_to_vk
from core.update_manager import check_for_update, prepare_and_launch_update
from app_config import APP_VERSION, UPDATE_MANIFEST_URL


class MainWindow:
    def __init__(self, master: Tk):
        self.master = master
        master.title("Desktop Automation App")
        master.resizable(False, False)

        # 盡量用 ttk 主題讓整體更一致（不同 Windows 主題也比較不突兀）
        try:
            style = ttk.Style(master)
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        # 全域邊距
        self._padx = 12
        self._pady = 8

        # 宏執行狀態（避免 callback 找不到屬性）
        self._stop_flag = False
        self._worker_thread = None

    # 暫停/繼續（F9 全域快捷鍵）
        self._paused = False
        self._pause_event = threading.Event()
        self._pause_event.set()  # set = not paused

        # 紀錄目前選擇到的視窗物件
        self.selected_window = None
        self.selected_hwnd = None

        # --- 視窗選擇區 ---
        header_frame = Frame(master)
        header_frame.pack(padx=self._padx, pady=(self._pady, 6), fill="x")

        Label(
            header_frame,
            text="① 選擇目標視窗",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        self.label = Label(header_frame, text="視窗標題 (例如: LINE / Discord / MapleStory)")
        self.label.pack(anchor="w")

        self.window_title_var = StringVar()
        self.window_title_entry = Entry(
            header_frame, textvariable=self.window_title_var, width=40
        )
        self.window_title_entry.pack(pady=2)

        # 目前視窗清單（給朋友用：不用猜標題）
        self.window_choices = []
        self.window_choice_var = StringVar(value="")

        picker_frame = Frame(header_frame)
        picker_frame.pack(fill="x", pady=(2, 0))

        self.window_picker = ttk.Combobox(
            picker_frame,
            textvariable=self.window_choice_var,
            values=self.window_choices,
            width=44,
            state="readonly",
        )
        self.window_picker.grid(row=0, column=0, padx=(0, 6), sticky="w")

        self.refresh_windows_button = Button(
            picker_frame,
            text="重新整理視窗",
            command=self.refresh_window_list,
            width=12,
        )
        self.refresh_windows_button.grid(row=0, column=1, padx=(0, 6))

        self.apply_window_button = Button(
            picker_frame,
            text="套用選擇",
            command=self.apply_selected_window_title,
            width=10,
        )
        self.apply_window_button.grid(row=0, column=2)

        self.select_window_button = Button(
            header_frame,
            text="選擇視窗並切到前景",
            command=self.select_window,
        )
        self.select_window_button.pack(pady=4)

    # --- （已移除）單次指令區 ---
    # 單次指令（直接輸入字串/方向鍵/點擊）在遊戲上成功率低、容易誤用，
    # 目前不提供給一般使用者，保留『走路+放技能』作為主要入口。

        # --- ② 快速開始（讓使用者知道怎麼用） ---
        guide_frame = Frame(master)
        guide_frame.pack(padx=self._padx, pady=(0, 6), fill="x")

        Label(
            guide_frame,
            text="② 快速開始",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")
        Label(
            guide_frame,
            text=(
                "1) 從下拉選一個目標視窗 → 按『選擇視窗並切到前景』\n"
                "2) 選『遊戲/前景模式』（成功率最高）\n"
                "3) 到④設定走路秒數/技能鍵序列 → 按『走路+放技能』開始\n"
                "4) 停止：按『停止』；暫停/繼續：按 F9"
            ),
            justify="left",
        ).pack(anchor="w", pady=(2, 0))

        # --- ③ 模式 與控制（不再提供單次指令，所以也不顯示 delay/repeat） ---
        control_frame = Frame(master)
        control_frame.pack(padx=self._padx, pady=6, fill="x")

        Label(
            control_frame,
            text="③ 模式與控制",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # 模式互斥：只能選一個，避免同時勾選造成邏輯互相覆蓋
        # 0=前景(一般/遊戲)、1=背景(PostMessage)、2=強制後台(盡力)
        self.mode_var = IntVar(value=0)
        Label(control_frame, text="執行模式（只能選一個）：").grid(
            row=1, column=0, columnspan=3, pady=(2, 0), sticky="w"
        )

        self.mode_fg_radio = Radiobutton(
            control_frame,
            text="遊戲/前景模式（會切到目標視窗，成功率最高）",
            variable=self.mode_var,
            value=0,
        )
        self.mode_fg_radio.grid(row=2, column=0, columnspan=3, sticky="w")

        self.mode_bg_radio = Radiobutton(
            control_frame,
            text="背景模式（不切焦點，部分程式可能無效）",
            variable=self.mode_var,
            value=1,
        )
        self.mode_bg_radio.grid(row=3, column=0, columnspan=3, sticky="w")

        self.mode_force_bg_radio = Radiobutton(
            control_frame,
            text="強制後台（盡力：即使在遊戲也硬送背景訊息，可能無效）",
            variable=self.mode_var,
            value=2,
        )
        self.mode_force_bg_radio.grid(row=4, column=0, columnspan=3, sticky="w")

        # 執行 & 腳本按鍵
        button_frame = Frame(control_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10, sticky="w")

        self.macro_button = Button(
            button_frame,
            text="走路+放技能",
            command=self.start_walk_cast_macro,
            width=12,
        )
        self.macro_button.grid(row=0, column=0, padx=(0, 6))

        self.stop_button = Button(
            button_frame,
            text="停止",
            command=self.stop_running,
            width=8,
        )
        self.stop_button.grid(row=0, column=1, padx=3)

        self.pause_button = Button(
            button_frame,
            text="暫停(F9)",
            command=self.toggle_pause,
            width=10,
        )
        self.pause_button.grid(row=0, column=2, padx=3)

        self.save_script_button = Button(
            button_frame,
            text="儲存腳本",
            command=self.save_script,
            width=10,
        )
        self.save_script_button.grid(row=1, column=0, padx=(0, 6), pady=(6, 0))

        self.load_script_button = Button(
            button_frame,
            text="載入腳本",
            command=self.load_script,
            width=10,
        )
        self.load_script_button.grid(row=1, column=1, padx=3, pady=(6, 0))

        self.update_button = Button(
            button_frame,
            text="檢查更新",
            command=self.check_for_updates,
            width=10,
        )
        self.update_button.grid(row=1, column=2, padx=3, pady=(6, 0))

        # --- 長按/技能/循環（宏設定）---
        macro_frame = Frame(master)
        macro_frame.pack(padx=self._padx, pady=6, fill="x")

        Label(
            macro_frame,
            text="④ 走路 + 放技能（主要功能）",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        Label(
            macro_frame,
            text=(
                "【走路+放技能】設定（按下『走路+放技能』就會：LEFT 按住 N 秒 → RIGHT 按住 N 秒。\n"
                "長按期間會每隔『技能間隔』輪播按一次『技能鍵序列』。"
            ),
            justify="left",
        ).grid(row=1, column=0, columnspan=3, sticky="w")

        Label(macro_frame, text="長按秒數 (左/右各按住幾秒，例如 10)").grid(
            row=2, column=0, sticky="w"
        )
        self.hold_seconds_var = StringVar(value="10")
        Entry(macro_frame, textvariable=self.hold_seconds_var, width=10).grid(
            row=3, column=0, pady=2, sticky="w"
        )

        self.cast_k_var = IntVar(value=1)
        Checkbutton(
            macro_frame,
            text="長按期間穿插技能鍵(輪播)",
            variable=self.cast_k_var,
        ).grid(row=2, column=1, padx=(15, 0), sticky="w")

        Label(macro_frame, text="技能鍵序列 (例如 K,1,2)").grid(
            row=3, column=1, padx=(15, 0), sticky="w"
        )
        self.skill_keys_var = StringVar(value="K")
        Entry(macro_frame, textvariable=self.skill_keys_var, width=18).grid(
            row=4, column=1, pady=2, sticky="w"
        )

        Label(macro_frame, text="技能間隔秒數 (例如 0.5)").grid(
            row=4, column=0, sticky="w"
        )
        self.k_interval_var = StringVar(value="0.5")
        Entry(macro_frame, textvariable=self.k_interval_var, width=10).grid(
            row=5, column=0, pady=2, sticky="w"
        )

        self.loop_forever_var = IntVar(value=1)
        Checkbutton(
            macro_frame,
            text="無限迴圈（直到按停止）",
            variable=self.loop_forever_var,
        ).grid(row=5, column=1, padx=(15, 0), sticky="w")

        Label(macro_frame, text="左右切換停頓秒數 (例如 0.05)").grid(
            row=6, column=0, sticky="w"
        )
        self.switch_gap_var = StringVar(value="0.05")
        Entry(macro_frame, textvariable=self.switch_gap_var, width=10).grid(
            row=7, column=0, pady=2, sticky="w"
        )

        # 相容舊版腳本欄位（UI 已移除單次指令區，但儲存/載入仍需這些變數）
        self.keys_var = StringVar(value="")
        self.delay_var = StringVar(value="0.1")
        self.repeat_var = StringVar(value="1")

        Button(
            macro_frame,
            text="套用範例：左10右10 + 每0.5秒按K,1,2",
            command=self.apply_macro_example,
        ).grid(row=8, column=0, columnspan=3, pady=(6, 0), sticky="w")

        # 狀態顯示
        status_frame = Frame(master)
        status_frame.pack(padx=self._padx, pady=(0, 10), fill="x")

        self.status_var = StringVar(value="尚未選擇視窗")
        self.status_label = Label(status_frame, textvariable=self.status_var, fg="blue")
        self.status_label.pack(anchor="w")

        self.version_var = StringVar(value=f"版本：v{APP_VERSION}")
        Label(status_frame, textvariable=self.version_var, fg="gray").pack(anchor="w")

        # 初次進來就先抓一次清單（需在 status_var 建立後）
        self.refresh_window_list()

        # F9（視窗有焦點時一定有效）
        try:
            self.master.bind_all("<F9>", lambda _e: self.toggle_pause())
        except Exception:
            pass

        # F9 全域快捷鍵：任何地方按到都能暫停/繼續（可能需要管理員權限）
        try:
            keyboard.add_hotkey("F9", lambda: self.master.after(0, self.toggle_pause))
        except Exception as e:
            self.status_var.set(
                f"提示：此電腦無法註冊『全域』F9 熱鍵({e})；工具視窗有焦點時 F9 仍可用。"
            )

    # ---- 視窗選擇 ----
    def refresh_window_list(self):
        """列出目前可見視窗標題（去掉空標題與本工具視窗），讓使用者直接點選。"""
        try:
            current_app_title = self.master.title()
        except Exception:
            current_app_title = ""

        titles: list[str] = []

        def enum_handler(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not title:
                return
            # 避免抓到自己
            if current_app_title and title == current_app_title:
                return
            titles.append(title)

        try:
            win32gui.EnumWindows(enum_handler, None)
        except Exception as e:
            self.status_var.set(f"重新整理視窗清單失敗: {e}")
            return

        # 去重、排序
        titles = sorted(set(titles))
        self.window_choices = titles
        self.window_picker["values"] = self.window_choices
        if titles:
            # 預設選第一個或保留原選擇
            if self.window_choice_var.get() not in titles:
                self.window_choice_var.set(titles[0])
            self.status_var.set(f"已更新視窗清單，共 {len(titles)} 個")
        else:
            self.window_choice_var.set("")
            self.status_var.set("找不到可見視窗，請先開啟目標程式")

    def apply_selected_window_title(self):
        """把下拉選到的視窗標題套用到輸入框。"""
        title = (self.window_choice_var.get() or "").strip()
        if not title:
            self.status_var.set("請先從下拉選擇一個視窗")
            return
        self.window_title_var.set(title)
        self.status_var.set(f"已套用視窗標題: {title}")

    def select_window(self):
        """根據視窗標題尋找實際視窗並切到前景。"""
        title = self.window_title_var.get().strip()
        if not title:
            self.status_var.set("請先輸入視窗標題，例如 LINE")
            return

        try:
            windows = pygetwindow.getWindowsWithTitle(title)
        except Exception as e:
            self.status_var.set(f"找視窗時發生錯誤: {e}")
            return

        if not windows:
            self.selected_window = None
            self.status_var.set(f"找不到包含『{title}』的視窗，請確認程式是否有開啟")
            return

        win = windows[0]
        try:
            win.activate()
            time.sleep(0.2)
        except Exception as e:
            self.status_var.set(f"切換視窗失敗: {e}")
            return

        self.selected_window = win
        # pygetwindow on Windows exposes _hWnd / hwnd depending on version
        self.selected_hwnd = getattr(win, "_hWnd", None) or getattr(win, "hwnd", None)
        self.status_var.set(f"已選擇並切換到視窗: {win.title}")

    def _bring_target_to_front(self) -> bool:
        """把目標視窗從最小化還原並嘗試置頂（前景模式用）。"""
        if self.selected_hwnd:
            hwnd = self.selected_hwnd
        elif self.selected_window is not None:
            hwnd = getattr(self.selected_window, "_hWnd", None) or getattr(
                self.selected_window, "hwnd", None
            )
        else:
            hwnd = None

        if not hwnd:
            return False

        try:
            # 如果最小化，先還原
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass

        try:
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            return True
        except Exception:
            # 失敗也不要中斷流程：後續仍可能用 pygetwindow.activate() 成功
            return False

    # ---- 指令與執行 ----
    def set_command(self, cmd: str):
        """快速把常用指令填到輸入框。"""
        self.keys_var.set(cmd)

    def input_keys(self):
        """解析指令，支援滑鼠 / 鍵盤，並可重複多次。"""
        cmd = self.keys_var.get().strip().upper()
        delay_str = self.delay_var.get().strip()
        repeat_str = self.repeat_var.get().strip()

        # 檢查延遲
        try:
            delay = float(delay_str) if delay_str else 0.1
        except ValueError:
            self.status_var.set("延遲秒數格式錯誤，請輸入數字，例如 0.1")
            return

        # 檢查次數
        try:
            repeat = int(repeat_str) if repeat_str else 1
            if repeat < 1:
                raise ValueError
        except ValueError:
            self.status_var.set("執行次數格式錯誤，請輸入大於等於 1 的整數")
            return

        if not cmd:
            self.status_var.set("請先輸入指令，例如 123、LEFTCLICK")
            return

        # 模式互斥
        mode = int(self.mode_var.get())
        bg_mode = mode in (1, 2)

        # 先把目前會怎麼送輸入講清楚，避免使用者覺得「沒反應」其實是送到別的地方
        mode_text = (
            "背景" if bg_mode else "前景"
        )
        try:
            target_text = (
                self.selected_window.title
                if (not bg_mode and self.selected_window is not None)
                else (self.window_title_var.get().strip() or "(未填標題)")
            )
        except Exception:
            target_text = "(未知)"
        self.status_var.set(f"準備執行：模式={mode_text}，目標={target_text}，指令={cmd} x{repeat}")

        # 背景模式需要 hwnd；若沒有，嘗試用標題再找一次
        hwnd = self.selected_hwnd
        if bg_mode and not hwnd:
            found = BackgroundInput.find_window_by_title_substring(
                self.window_title_var.get().strip()
            )
            if found:
                hwnd = found.hwnd
                self.selected_hwnd = hwnd
            else:
                self.status_var.set(
                    "背景模式：找不到目標視窗 hwnd。請先輸入更精確的視窗標題，或按『選擇視窗並切到前景』一次。"
                )
                return

        # 非背景模式：切回選擇的視窗（原本行為）
        if (not bg_mode) and self.selected_window is not None:
            try:
                # 使用者希望視窗縮小時按執行會自動拉到最前面
                self._bring_target_to_front()
                self.selected_window.activate()
                time.sleep(0.2)
            except Exception as e:
                self.status_var.set(f"切回目標視窗失敗: {e}")
                return
        elif (not bg_mode) and self.selected_window is None:
            # 前景模式但沒選視窗：會送到「目前正在打字的地方」
            self.status_var.set(
                "前景模式：你尚未按『選擇視窗並切到前景』，輸入會送到目前游標所在的視窗。"
            )

        # 依照次數重複執行
        try:
            for i in range(repeat):
                if cmd == "LEFTCLICK":
                    if bg_mode and hwnd:
                        BackgroundInput.click(hwnd, x=10, y=10, button="left")
                    else:
                        pyautogui.click(button="left")
                elif cmd == "RIGHTCLICK":
                    if bg_mode and hwnd:
                        BackgroundInput.click(hwnd, x=10, y=10, button="right")
                    else:
                        pyautogui.click(button="right")
                else:
                    if bg_mode and hwnd:
                        # 背景模式：
                        # 1) 多字元（例如 KKKK / ABC）一律當「文字」走 WM_CHAR，最容易在記事本看到
                        # 2) 單一字元：優先用 WM_CHAR，Windows 一般輸入框反應最好；
                        #    若是特殊鍵（UP/DOWN/F1...）才走 VK
                        if len(cmd) > 1 and cmd.isprintable():
                            BackgroundInput.send_text(hwnd, cmd, interval=delay)
                        else:
                            vk = parse_command_to_vk(cmd)
                            if vk is not None and cmd not in {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}:
                                BackgroundInput.press_vk(hwnd, vk, hold_seconds=0)
                            else:
                                BackgroundInput.send_text(hwnd, cmd, interval=delay)
                    else:
                        if cmd in {"UP", "DOWN", "LEFT", "RIGHT"}:
                            pyautogui.press(cmd.lower())
                        else:
                            pyautogui.typewrite(cmd, interval=delay)

                if repeat > 1 and i != repeat - 1:
                    time.sleep(delay)

            self.status_var.set(f"指令已執行: {cmd}，共 {repeat} 次")
        except Exception as e:
            self.status_var.set(f"執行失敗: {e}")

    # ---- 宏：左按住/右按住，期間穿插技能鍵 ----
    def stop_running(self):
        self._stop_flag = True
        # 避免卡在暫停等待
        try:
            self._pause_event.set()
        except Exception:
            pass
        self.status_var.set("已送出停止指令，等待目前動作結束…")

    def toggle_pause(self):
        # 暫停不等於停止：暫停時宏會停在安全點，直到再次按 F9/按鈕才繼續
        self._paused = not self._paused
        if self._paused:
            self._pause_event.clear()
            self.pause_button.configure(text="繼續(F9)")
            self.status_var.set("已暫停（F9 或按『繼續』可恢復）")
        else:
            self._pause_event.set()
            self.pause_button.configure(text="暫停(F9)")
            self.status_var.set("已繼續")

    def start_walk_cast_macro(self):
        if self._worker_thread and self._worker_thread.is_alive():
            self.status_var.set("已有任務正在執行，請先按停止")
            return

        # 讀取設定
        try:
            hold_seconds = float(self.hold_seconds_var.get().strip() or "0")
            if hold_seconds <= 0:
                raise ValueError
        except ValueError:
            self.status_var.set("長按秒數格式錯誤，請輸入大於 0 的數字")
            return

        try:
            k_interval = float(self.k_interval_var.get().strip() or "0")
            if k_interval <= 0:
                k_interval = 0.5
        except ValueError:
            self.status_var.set("技能間隔秒數格式錯誤，請輸入數字，例如 0.5")
            return

        try:
            switch_gap = float(self.switch_gap_var.get().strip() or "0")
            if switch_gap < 0:
                raise ValueError
        except ValueError:
            self.status_var.set("左右切換停頓秒數格式錯誤，請輸入大於等於 0 的數字")
            return

        # 解析技能輪播序列
        raw_keys = (self.skill_keys_var.get() or "").strip()
        skills = [s.strip().upper() for s in raw_keys.split(",") if s.strip()]
        if self.cast_k_var.get() == 1 and not skills:
            self.status_var.set("技能鍵序列不可為空，例如 K 或 K,1,2")
            return

        self._stop_flag = False

        loop_forever = self.loop_forever_var.get() == 1

        # 顯示本次宏使用的模式，避免誤會「為何會切前景 / 為何背景沒動」
        mode = int(self.mode_var.get())
        bg_mode = mode in (1, 2)
        mode_text = "背景" if bg_mode else "前景"

        # 前景模式：先把視窗還原/置頂，避免使用者覺得沒反應
        if not bg_mode:
            self._bring_target_to_front()

        self._worker_thread = threading.Thread(
            target=self._walk_cast_worker,
            args=(hold_seconds, k_interval, switch_gap, skills, loop_forever),
            daemon=True,
        )
        self._worker_thread.start()
        # 讓使用者看得到已經開始跑
        if loop_forever:
            self.status_var.set(
                f"宏開始：LEFT 按住 → RIGHT 按住（無限迴圈，模式={mode_text}）。"
            )
        else:
            self.status_var.set(f"宏開始：LEFT 按住 → RIGHT 按住（模式={mode_text}）。")

    def _walk_cast_worker(
        self,
        hold_seconds: float,
        k_interval: float,
        switch_gap: float,
        skills: list[str],
        loop_forever: bool,
    ):
        mode = int(self.mode_var.get())
        bg_mode = mode in (1, 2)
        hwnd = self.selected_hwnd
        if bg_mode and not hwnd:
            found = BackgroundInput.find_window_by_title_substring(
                self.window_title_var.get().strip()
            )
            if found:
                hwnd = found.hwnd
                self.selected_hwnd = hwnd

        cast_k = self.cast_k_var.get() == 1

        # 背景模式找不到 hwnd 就直接停止，避免看起來像「有跑但沒動作」
        if bg_mode and not hwnd:
            try:
                self.master.after(
                    0,
                    lambda: self.status_var.set(
                        "背景模式找不到目標視窗，請先選擇視窗或輸入更精確標題"
                    ),
                )
            except Exception:
                pass
            return

        # 若不是背景模式，就盡量把目標視窗切回前景（提高成功率）
        if (not bg_mode) and self.selected_window is not None:
            try:
                self.selected_window.activate()
                time.sleep(0.2)
            except Exception:
                pass

        def set_status(text: str):
            try:
                self.master.after(0, lambda t=text: self.status_var.set(t))
            except Exception:
                pass

        def release_dir_key(dir_key: str):
            """Best-effort release to prevent stuck keys (e.g., RIGHT still held).

            On some games/windows the key-up for the opposite direction may not be
            delivered, which results in continuous movement to one side.  To make
            the macro more reliable we call this on *both* directions before
            pressing the next one and again when we're done holding.
            """
            try:
                if bg_mode and hwnd:
                    vk = parse_command_to_vk(dir_key)
                    if vk is not None:
                        BackgroundInput.send_vk(hwnd, vk, keyup=True)
                else:
                    pyautogui.keyUp(dir_key.lower())
            except Exception:
                pass

        def release_all_dir_keys():
            release_dir_key("LEFT")
            release_dir_key("RIGHT")

        # 先放開左右鍵，避免上一次意外中斷造成卡鍵
        # (有時候只放開相反方向不夠，故兩邊都做一次)
        release_all_dir_keys()


        skill_index = 0

        def next_skill() -> str | None:
            nonlocal skill_index
            if not skills:
                return None
            key = skills[skill_index % len(skills)]
            skill_index += 1
            return key

        def press_skill_once():
            if not cast_k:
                return
            skill = next_skill()
            if not skill:
                return
            try:
                if bg_mode and hwnd:
                    vk = parse_command_to_vk(skill)
                    if vk is not None:
                        BackgroundInput.press_vk(hwnd, vk, hold_seconds=0)
                else:
                    pyautogui.press(skill.lower())
            except Exception:
                # 不要讓技能鍵失敗中斷整個宏
                return

        def hold_dir(dir_key: str, seconds: float):
            start = time.monotonic()
            last_k = 0.0
            paused_total = 0.0
            key_is_down = False

            try:
                # 為了防止方向鍵同時被壓住或上一輪沒有正確放開，我們先放開
                # 兩邊方向鍵，再開始按下當前方向。
                release_all_dir_keys()

                if bg_mode and hwnd:
                    vk = parse_command_to_vk(dir_key)
                    if vk is None:
                        return
                    BackgroundInput.send_vk(hwnd, vk, keyup=False)
                    key_is_down = True
                else:
                    pyautogui.keyDown(dir_key.lower())
                    key_is_down = True

                while not self._stop_flag and (time.monotonic() - start - paused_total) < seconds:
                    # 暫停：立刻放開方向鍵，恢復後再按回去
                    if not self._pause_event.is_set():
                        if key_is_down:
                            release_dir_key(dir_key)
                            key_is_down = False
                        set_status("已暫停（方向鍵已放開，F9 可恢復）")
                        pause_started = time.monotonic()
                        while not self._stop_flag and not self._pause_event.is_set():
                            time.sleep(0.02)
                        paused_total += (time.monotonic() - pause_started)
                        if self._stop_flag:
                            break
                        if bg_mode and hwnd:
                            vk = parse_command_to_vk(dir_key)
                            if vk is not None:
                                BackgroundInput.send_vk(hwnd, vk, keyup=False)
                                key_is_down = True
                        else:
                            pyautogui.keyDown(dir_key.lower())
                            key_is_down = True

                    now = time.monotonic()
                    if cast_k and (now - last_k) >= k_interval:
                        press_skill_once()
                        last_k = now
                    time.sleep(0.02)
            finally:
                if key_is_down:
                    try:
                        if bg_mode and hwnd:
                            vk = parse_command_to_vk(dir_key)
                            if vk is not None:
                                BackgroundInput.send_vk(hwnd, vk, keyup=True)
                        else:
                            pyautogui.keyUp(dir_key.lower())
                    except Exception:
                        pass

        loop_count = 0
        try:
            while not self._stop_flag:
                loop_count += 1
                set_status("宏進行中：按住 LEFT…")
                hold_dir("LEFT", hold_seconds)
                # 再保險一次釋放所有方向鍵，並稍微讓控制系統有空隙
                release_all_dir_keys()
                time.sleep(switch_gap)
                if self._stop_flag:
                    break
                set_status("宏進行中：按住 RIGHT…")
                hold_dir("RIGHT", hold_seconds)
                release_all_dir_keys()
                time.sleep(switch_gap)

                if not loop_forever:
                    break
        finally:
            # 最後再做一次全面放鍵，避免中斷後殘留方向鍵
            release_all_dir_keys()

        # 更新狀態（回主執行緒）
        try:
            self.master.after(
                0,
                lambda: self.status_var.set(
                    f"宏停止。已完成迴圈: {loop_count}"
                ),
            )
        except Exception:
            pass

    def apply_macro_example(self):
        """把常用範例一鍵填好，讓新手知道怎麼用。"""
        self.hold_seconds_var.set("10")
        self.cast_k_var.set(1)
        self.skill_keys_var.set("K,1,2")
        self.k_interval_var.set("0.5")
        try:
            self.loop_forever_var.set(1)
        except Exception:
            pass
        self.status_var.set(
            "已套用範例：請按『走路+放技能』(不是『開始執行』)；停止請按『停止』"
        )

    # ---- 腳本儲存 / 載入 ----
    def save_script(self):
        """將目前設定存成 JSON 腳本。"""
        data = {
            "window_title": self.window_title_var.get().strip(),
            "command": self.keys_var.get().strip(),
            "delay": self.delay_var.get().strip(),
            "repeat": self.repeat_var.get().strip(),
            # 宏設定（舊腳本沒有也沒關係）
            "macro_hold_seconds": self.hold_seconds_var.get().strip(),
            "macro_cast_skills": int(self.cast_k_var.get()),
            "macro_skill_keys": self.skill_keys_var.get().strip(),
            "macro_interval": self.k_interval_var.get().strip(),
            "macro_switch_gap": self.switch_gap_var.get().strip(),
            "macro_loop_forever": int(getattr(self, "loop_forever_var", IntVar(value=1)).get()),
            "mode": int(self.mode_var.get()),
        }

        path = filedialog.asksaveasfilename(
            title="儲存腳本",
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.status_var.set(f"腳本已儲存: {path}")
        except Exception as e:
            messagebox.showerror("儲存失敗", f"無法儲存腳本: {e}")

    def load_script(self):
        """從 JSON 腳本載入設定。"""
        path = filedialog.askopenfilename(
            title="載入腳本",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("載入失敗", f"無法讀取腳本: {e}")
            return

        self.window_title_var.set(str(data.get("window_title", "")))
        self.keys_var.set(str(data.get("command", "")))
        self.delay_var.set(str(data.get("delay", "0.1")))
        self.repeat_var.set(str(data.get("repeat", "1")))

        # 宏設定（沒有就保留目前 UI 的值）
        if "macro_hold_seconds" in data:
            self.hold_seconds_var.set(str(data.get("macro_hold_seconds", "10")))
        if "macro_cast_skills" in data:
            try:
                self.cast_k_var.set(int(data.get("macro_cast_skills", 1)))
            except Exception:
                pass
        if "macro_skill_keys" in data:
            self.skill_keys_var.set(str(data.get("macro_skill_keys", "K")))
        if "macro_interval" in data:
            self.k_interval_var.set(str(data.get("macro_interval", "0.5")))
        if "macro_switch_gap" in data and hasattr(self, "switch_gap_var"):
            self.switch_gap_var.set(str(data.get("macro_switch_gap", "0.05")))
        if "macro_loop_forever" in data and hasattr(self, "loop_forever_var"):
            try:
                self.loop_forever_var.set(int(data.get("macro_loop_forever", 1)))
            except Exception:
                pass
        if "mode" in data:
            try:
                self.mode_var.set(int(data.get("mode", 0)))
            except Exception:
                pass

        self.status_var.set(f"腳本已載入: {path}")

    # ---- 線上更新 ----
    def check_for_updates(self):
        if not UPDATE_MANIFEST_URL.strip():
            self.status_var.set("尚未設定更新網址：請先在 app_config.py 填入 UPDATE_MANIFEST_URL")
            return

        self.update_button.configure(state="disabled")
        self.status_var.set("正在檢查更新…")
        threading.Thread(target=self._check_update_worker, daemon=True).start()

    def _check_update_worker(self):
        try:
            result = check_for_update(APP_VERSION, UPDATE_MANIFEST_URL, timeout=8)
            self.master.after(0, lambda r=result: self._on_update_checked(r))
        except Exception as e:
            self.master.after(0, lambda: self._on_update_failed(str(e)))

    def _on_update_failed(self, message: str):
        self.update_button.configure(state="normal")
        self.status_var.set(f"檢查更新失敗: {message}")

    def _on_update_checked(self, result: dict):
        self.update_button.configure(state="normal")

        if not result.get("has_update"):
            self.status_var.set(f"已是最新版本 (v{APP_VERSION})")
            return

        latest = result.get("latest_version", "?")
        notes = (result.get("notes") or "").strip()
        notes_text = f"\n\n更新內容：\n{notes}" if notes else ""
        ok = messagebox.askyesno(
            "發現新版本",
            f"目前版本：v{APP_VERSION}\n最新版本：v{latest}\n\n是否立即下載並更新？{notes_text}",
        )
        if not ok:
            self.status_var.set("已取消更新")
            return

        download_url = (result.get("download_url") or "").strip()
        if not download_url:
            self.status_var.set("更新資料缺少 download_url")
            return

        self.update_button.configure(state="disabled")
        self.status_var.set("正在下載更新並準備安裝…")
        threading.Thread(
            target=self._apply_update_worker,
            args=(download_url, latest),
            daemon=True,
        ).start()

    def _apply_update_worker(self, download_url: str, latest_version: str):
        try:
            if not getattr(sys, "frozen", False):
                self.master.after(
                    0,
                    lambda: self._on_update_dev_mode(download_url, latest_version),
                )
                return

            prepare_and_launch_update(download_url)
            self.master.after(0, lambda: self._on_update_ready_to_restart(latest_version))
        except Exception as e:
            self.master.after(0, lambda: self._on_update_failed(str(e)))

    def _on_update_dev_mode(self, download_url: str, latest_version: str):
        self.update_button.configure(state="normal")
        self.status_var.set(f"偵測到新版本 v{latest_version}（目前為原始碼模式）")
        open_link = messagebox.askyesno(
            "原始碼模式",
            "目前不是 exe 執行（無法自動覆蓋自己）。\n要開啟下載頁面手動更新嗎？",
        )
        if open_link:
            try:
                webbrowser.open(download_url)
            except Exception:
                pass

    def _on_update_ready_to_restart(self, latest_version: str):
        self.status_var.set(f"已準備更新到 v{latest_version}，程式即將關閉並重啟")
        messagebox.showinfo("更新準備完成", "更新已下載，按確定後將自動關閉並重新開啟新版本。")
        try:
            self.master.destroy()
        except Exception:
            pass


def run_app():
    root = Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
