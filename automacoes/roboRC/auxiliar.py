import pyautogui
import time

time.sleep(5)

x, y = pyautogui.position()
print(f"x={x}, y={y}")