import os
import pandas as pd
import re
import time
from datetime import datetime, timedelta
from openpyxl.reader.excel import load_workbook
import config as CONFIG
import pygetwindow as gw
import win32gui
import win32con
import pyautogui
import pyperclip

two_yesterday = datetime.now() - timedelta(days=3)
two_yesterday = two_yesterday.strftime("%d%m%Y")

excel_path = rf'C:\Users\USER\Desktop\Caja - {two_yesterday} - Agresivo.xlsx'
second_excel_path = rf'C:\Users\USER\Desktop\Solicitudes - {two_yesterday} - AGR.xlsx'


df = pd.read_excel(excel_path, sheet_name='AccountDetail')
df_second = pd.read_excel(second_excel_path, sheet_name='Sheet1')


columna = 'Unnamed: 5'  # Cambia esto al nombre de la columna correcta
fila_inicio = 8  # Cambia esto al número de fila desde donde quieres empezar (1-indexed)

target_column = 'Unnamed: 11'
start_row = 8

screen_width, screen_height = pyautogui.size()
print(f"Resolución de pantalla detectada: {screen_width}x{screen_height}")

def buscar_y_extraer_valores(valor, df):
    fila = df[df['Dni'] == valor]
    if not fila.empty:
        nombre = fila['Nombre'].values[0]
        codigo_flip = fila['Codigo Flip'].values[0]
        return nombre, codigo_flip
    else:
        return None, None


def main():
    if focus_existing_window("Nuevo Telecrédito"):
        print("La ventana ya estaba abierta y fue enfocada.")
    time.sleep(0.5)
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.15))
    time.sleep(0.5)
    pyautogui.hotkey('tab')
    pyautogui.typewrite('agr')
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
    wait_for_color(int(screen_width * 0.39), int(screen_height * 0.74), (232, 246, 252))
    pyautogui.scroll(-500)
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.60))
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.30))
    wait_for_color(int(screen_width * 0.316), int(screen_height * 0.804), (212,102,40))
    pyautogui.click(int(screen_width * 0.35), int(screen_height * 0.36))
    time.sleep(0.5)
    # escribir el dia de antes de ayer
    pyautogui.hotkey('tab')
    pyautogui.typewrite(two_yesterday)
    pyautogui.hotkey('tab')
    yesterday = datetime.now() - timedelta(days=2)
    yesterday = yesterday.strftime("%d%m%Y")
    pyautogui.typewrite(yesterday)
    # press Search button
    pyautogui.click(int(screen_width * 0.87), int(screen_height * 0.715))

    num_pages = detect_number_of_pages()
    print(f"Number of pages detected: {num_pages}")
    for valor in df[columna].iloc[fila_inicio - 1:]:
        for page in range(num_pages):
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'f')
            pyautogui.typewrite(str(valor))
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(0.5)
            if search_color_on_screen((255,150,50)):
                print(f"Match found for {valor} on page {page + 1}")
                break
            else:
                pyautogui.scroll(-4000)
                time.sleep(0.5)
                go_to_next_page(page)
                time.sleep(0.5)
                print(f"No match found for {valor} on page {page + 1}, moving to next page")

# def detect_agressive_independent():
#     attempt = 0
#     while not wait_for_color(int(screen_width * 0.39), int(screen_height * 0.74), (232,246,252)):
#         if attempt == 0:
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.25))
#         elif attempt == 1:
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.30))
#         elif attempt == 2:
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.20))
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.35))
#         elif attempt == 3:
#             pyautogui.scroll(-500)
#             pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.35))
#         attempt += 1
#         wait_for_color(int(screen_width * 0.20), int(screen_height * 0.76), (72,147,216))

def detect_number_of_pages():
    page_buttons_coords = [
        (screen_width * 0.51, screen_height * 0.82),  # Button 1
        (screen_width * 0.54, screen_height * 0.82),  # Button 2
        (screen_width * 0.57, screen_height * 0.82),  # Button 3
        (screen_width * 0.60, screen_height * 0.82),   # Button 4 (if exists)
        (screen_width * 0.63, screen_height * 0.82),  # Button 5 (if exists)
        (screen_width * 0.66, screen_height * 0.82)   # Button 6 (if exists)
    ]
    num_pages = 0
    for coord in page_buttons_coords:
        try:
            pyautogui.click(coord)
            num_pages += 1
        except:
            break
    return num_pages

def search_color_on_screen(target_color, timeout=2):
    """
    Search for a specific color on the screen and return its coordinates.

    Args:
        target_color (tuple): The RGB color to search for, e.g., (241, 155, 74).
        timeout (int): The maximum time to search for the color.

    Returns:
        tuple: The coordinates (x, y) of the color if found, otherwise None.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        screenshot = pyautogui.screenshot()
        width, height = screenshot.size
        for x in range(width):
            for y in range(height):
                if screenshot.getpixel((x, y)) == target_color:
                    perform_action_and_insert_value(x, y, target_color)
                    return (x, y)
        time.sleep(0.5)
    print(f"Timeout reached. Color {target_color} not detected on the screen.")
    return None

def perform_action_and_insert_value(x, y, target_color):
    print(f"Color {target_color} detected at position ({x}, {y})")
    pyautogui.click(int(screen_width * 0.94), y)
    wait_for_color(int(screen_width * 0.55), int(screen_height * 0.24), (255,255,255))
    time.sleep(0.5)
    pyautogui.doubleClick(int(screen_width * 0.484), int(screen_height * 0.813))
    pyautogui.hotkey('ctrl', 'c')
    copied_value = pyperclip.paste()
    print(f"Copied value: {copied_value}")

    current_row = start_row
    while current_row < len(df) and pd.notna(df.at[current_row, target_column]):
        current_row += 1

    df.at[current_row, target_column] = copied_value
    print(f"Value '{copied_value}' inserted into row {current_row}, column {target_column}")

    workbook = load_workbook(excel_path)
    sheet = workbook['AccountDetail']
    cell = sheet.cell(row=current_row + 1, column=df.columns.get_loc(target_column) + 1)
    cell.value = copied_value

    # Buscar y extraer valores del segundo archivo
    nombre, codigo_flip = buscar_y_extraer_valores(copied_value, df_second)
    if nombre and codigo_flip:
        df.at[current_row, 'Nombre'] = nombre
        df.at[current_row, 'Codigo Flip'] = codigo_flip
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Nombre') + 1).value = nombre
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Codigo Flip') + 1).value = codigo_flip
        print(f"Nombre '{nombre}' y Codigo Flip '{codigo_flip}' insertados en la fila {current_row}")

    workbook.save(excel_path)
    print("DataFrame saved to Excel file")

    pyautogui.click(int(screen_width * 0.21), int(screen_height * 0.289))
    time.sleep(0.5)

def go_to_next_page(current_page):
    page_buttons_coords = [
        (screen_width * 0.51, screen_height * 0.82),  # Button 1
        (screen_width * 0.54, screen_height * 0.82),  # Button 2
        (screen_width * 0.57, screen_height * 0.82),  # Button 3
        (screen_width * 0.60, screen_height * 0.82),   # Button 4 (if exists)
        (screen_width * 0.63, screen_height * 0.82),  # Button 5 (if exists)
        (screen_width * 0.66, screen_height * 0.82)   # Button 6 (if exists)
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

def wait_for_color(x, y, target_color, timeout=2):
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


