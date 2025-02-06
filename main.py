import os
import time

import pyscreeze

import config as CONFIG
import pygetwindow as gw
import win32gui
import win32con
import pyautogui

screen_width, screen_height = pyautogui.size()
print(f"Resolución de pantalla detectada: {screen_width}x{screen_height}")

def main():
    if focus_existing_window("Nuevo Telecrédito"):
        print("La ventana ya estaba abierta y fue enfocada.")
    elif focus_existing_window("Banco de Crédito >>BCP>>"):
        print("La ventana ya estaba abierta y fue enfocada.")
    else:
        print("Abriendo Google Chrome...")
        os.system("start https://www.google.com")  # Windows
        time.sleep(3)
        print("Escribiendo URL de BCP TLC")
        pyautogui.click(int(screen_width * 0.20), int(screen_height * 0.08))  # Hacer focus en la barra de direcciones
        pyautogui.typewrite("https://www.tlcbcp.com/#/h/saldos-movimientos/v2/detalle-cuenta/movimientos-historicos")
        pyautogui.press("enter")

    time.sleep(1)
    pyautogui.hotkey("tab")
    number_positions = locate_numbers()
    press_pin(CONFIG.TELECREDITO_PASSWORD, number_positions)

def focus_existing_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        win = windows[0]
        win32gui.ShowWindow(win._hWnd, win32con.SW_MAXIMIZE)  # Maximizar la ventana
        win32gui.SetForegroundWindow(win._hWnd)  # Traer al frente
        print(f"Ventana '{title}' enfocada.")
        return True
    return False

def locate_numbers():
    number_positions = {}
    for number in range(10):
        pyscreeze.screenshot(f"images/{number}.png", region=(0, 0, screen_width, screen_height))
        image_path = f"images/{number}.png"  # Ruta de la imagen del número
        position = pyautogui.locateOnScreen(image_path, confidence=0.8)
        if position:
            number_positions[number] = pyautogui.center(position)
        else:
            raise ValueError(f"No se encontró la imagen para el número {number}")
    return number_positions

def press_pin(pin, number_positions):
    for digit in pin:
        digit = int(digit)
        if digit in number_positions:
            pyautogui.moveTo(number_positions[digit], duration=0.2)
            pyautogui.click()
        else:
            raise ValueError(f"Posición no encontrada para el dígito {digit}")

main()