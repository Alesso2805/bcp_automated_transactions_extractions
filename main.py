import os
import pandas as pd
import re
import time
from datetime import datetime, timedelta
import config as CONFIG
import pygetwindow as gw
import win32gui
import win32con
import pyautogui
import pyperclip

df = pd.read_excel(r'C:\Users\Alessandro\Desktop\Caja - 04022025 - Agresivo.xlsx', sheet_name='BCP_AGR')

columna = 'Unnamed: 5'  # Cambia esto al nombre de la columna correcta
fila_inicio = 10  # Cambia esto al número de fila desde donde quieres empezar (1-indexed)

screen_width, screen_height = pyautogui.size()
print(f"Resolución de pantalla detectada: {screen_width}x{screen_height}")

def main():
    if focus_existing_window("Nuevo Telecrédito"):
        print("La ventana ya estaba abierta y fue enfocada.")
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
    time.sleep(0.5)
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.25))
    if not wait_for_color(int(screen_width * 0.20), int(screen_height * 0.76), (72,147,216)):
        pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
        pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.30))
    pyautogui.scroll(-500)
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.50))
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.30))
    wait_for_color(int(screen_width * 0.316), int(screen_height * 0.804), (212,102,40))
    pyautogui.click(int(screen_width * 0.35), int(screen_height * 0.36))

    pyautogui.click(int(screen_width * 0.22), int(screen_height * 0.72))
    yesterday = datetime.now() - timedelta(days=2)
    yesterday = yesterday.strftime("%d%m%Y")
    pyautogui.typewrite(yesterday)
    pyautogui.click(int(screen_width * 0.41), int(screen_height * 0.72))
    pyautogui.typewrite(yesterday)
    pyautogui.click(int(screen_width * 0.85), int(screen_height * 0.72))

    num_pages = detect_number_of_pages()
    print(f"Number of pages detected: {num_pages}")

    for valor in df[columna].iloc[fila_inicio - 1:]:
        for page in range(num_pages):
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'f')
            pyautogui.typewrite(str(valor))
            pyautogui.press('enter')
            time.sleep(2)
            if search_successful():
                print(f"Match found for {valor} on page {page + 1}")
                break
            else:
                pyautogui.scroll(-1800)
                go_to_next_page(page)
                print(f"No match found for {valor} on page {page + 1}, moving to next page")

def detect_number_of_pages():
    page_buttons_coords = [
        (1368, 1164),  # Button 1
        (1433, 1156),  # Button 2
        (1511, 1154),  # Button 3
        (1580, 1168)   # Button 4 (if exists)
    ]
    num_pages = 0
    for coord in page_buttons_coords:
        try:
            pyautogui.click(coord)
            num_pages += 1
        except:
            break
    return num_pages

def search_successful():
    # Wait for a short period to allow the search to complete
    time.sleep(1)

    # Check if the search box is highlighted, indicating a match was found
    search_box_coords = (screen_width // 2, screen_height // 2)  # Adjust as needed
    search_box_color = pyautogui.pixel(search_box_coords[0], search_box_coords[1])

    # Define the color that indicates a match was found (adjust as needed)
    match_color = (241,155,74)  # Example color

    return search_box_color == match_color

def go_to_next_page(current_page):
    page_buttons_coords = [
        (1368, 1164),  # Button 1
        (1433, 1156),  # Button 2
        (1511, 1154),  # Button 3
        (1580, 1168)   # Button 4 (if exists)
    ]
    if current_page < len(page_buttons_coords):
        pyautogui.click(page_buttons_coords[current_page])
    else:
        print("No more pages to navigate.")

def focus_existing_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        win = windows[0]
        win32gui.ShowWindow(win._hWnd, win32con.SW_MAXIMIZE)  # Maximizar la ventana
        win32gui.SetForegroundWindow(win._hWnd)  # Traer al frente
        print(f"Ventana '{title}' enfocada.")
        return True
    return False

def press_pin(numbers_list, pin):
    pixel_positions = [
        (1120, 620),  # Position 0
        (1230, 620),  # Position 1
        (1335, 620),  # Position 2
        (1435, 620),  # Position 3
        (1120, 700),  # Position 4
        (1230, 700),  # Position 5
        (1335, 700),  # Position 6
        (1435, 700),  # Position 7
        (1230, 770),  # Position 8
        (1335, 770),  # Position 9
    ]
    number_to_pixel = {number: pixel for number, pixel in zip(numbers_list, pixel_positions)}
    pixels_to_press = []
    for digit_char in pin:
        digit = int(digit_char)
        pixel = number_to_pixel.get(digit)
        if pixel is None:
            raise ValueError(f"Digit '{digit}' not found in the numbers list.")
        pixels_to_press.append(pixel)
    return pixels_to_press

def wait_for_color(x, y, target_color, timeout=5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_color = pyautogui.pixel(x, y)
        if current_color == target_color:
            print(f"Color detectado en posición ({x}, {y}): {current_color}")
            return True
        time.sleep(0.5)
    print("Tiempo de espera agotado. No se detectó el color.")
    return False

main()