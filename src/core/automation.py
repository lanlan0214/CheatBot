class Automation:
    def __init__(self):
        import pyautogui
        self.pyautogui = pyautogui

    def input_keys(self, keys: str, delay: float = 0.0):
        self.pyautogui.write(keys, interval=delay)