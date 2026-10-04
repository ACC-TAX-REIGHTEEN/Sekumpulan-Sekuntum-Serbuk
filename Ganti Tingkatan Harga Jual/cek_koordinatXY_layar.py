import pyautogui
import time

print("Tekan Ctrl+C di terminal untuk berhenti.")
try:
    while True:
        x, y = pyautogui.position()
        print(f"Posisi Mouse: X={x}, Y={y}", end='\r')
        time.sleep(0.1)
except KeyboardInterrupt:
    print("Selesai.")