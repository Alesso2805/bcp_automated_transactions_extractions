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

two_yesterday = (datetime.now() - timedelta(days=2)).strftime("%d%m%Y")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")

excel_path = rf'C:\Users\Flip\Desktop\Caja - {two_yesterday} - Agresivo.xlsx'
second_excel_path = rf'C:\Users\Flip\Desktop\Solicitudes - {two_yesterday} - AGR.xlsx'

df = pd.read_excel(excel_path, sheet_name='AccountDetail')
df_second = pd.read_excel(second_excel_path, sheet_name='Sheet1')

columna = 'Unnamed: 5'
fila_inicio = 8
target_column = 'Unnamed: 11'
start_row = 8

screen_width, screen_height = pyautogui.size()
print(f"Resolución de pantalla detectada: {screen_width}x{screen_height}")

def buscar_y_extraer_valores(valor, df):
    fila = df[df['Dni'] == valor]
    if not fila.empty:
        return fila['Nombre'].values[0], fila['Codigo Flip'].values[0], fila['Cantidad'].values[0]
    return None, None, None

def focus_existing_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        win = windows[0]
        win32gui.ShowWindow(win._hWnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(win._hWnd)
        print(f"Ventana '{title}' enfocada.")
        return True
    return False

def wait_for_color(x, y, target_color, timeout=2):
    start_time = time.time()
    while time.time() - start_time < timeout:
        if pyautogui.pixel(x, y) == target_color:
            print(f"Color detectado en posición ({x}, {y}): {target_color}")
            return True
        time.sleep(0.5)
    print("Tiempo de espera agotado. No se detectó el color.")
    return False

def perform_action_and_insert_value(x, y, target_color, current_row):
    print(f"Color {target_color} detected at position ({x}, {y})")
    pyautogui.click(int(screen_width * 0.94), y)
    wait_for_color(int(screen_width * 0.55), int(screen_height * 0.24), (255, 255, 255))
    time.sleep(1.5)
    pyautogui.doubleClick(int(screen_width * 0.484), int(screen_height * 0.813))
    pyautogui.hotkey('ctrl', 'c')
    copied_value = pyperclip.paste()
    print(f"Copied value: {copied_value}")

    while current_row < len(df) and pd.notna(df.at[current_row, target_column]):
        current_row += 1

    df.at[current_row, target_column] = copied_value
    print(f"Value '{copied_value}' inserted into row {current_row}, column {target_column}")

    workbook = load_workbook(excel_path)
    sheet = workbook['AccountDetail']
    sheet.cell(row=current_row + 1, column=df.columns.get_loc(target_column) + 1).value = copied_value

    nombre, codigo_flip, monto = buscar_y_extraer_valores(copied_value, df_second)
    if nombre and codigo_flip and monto:
        df.at[current_row, 'Unnamed: 9'] = nombre
        df.at[current_row, 'Unnamed: 15'] = codigo_flip
        df.at[current_row, 'Unnamed: 13'] = monto
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 9') + 1).value = nombre
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 15') + 1).value = codigo_flip
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 13') + 1).value = monto
        print(f"Nombre '{nombre}', Codigo Flip '{codigo_flip}' y Cantidad '{monto}' insertados en la fila {current_row}")

        workbook.save(excel_path)
        print("DataFrame saved to Excel file")
        pyautogui.click(int(screen_width * 0.21), int(screen_height * 0.289))
        time.sleep(0.5)

def detect_number_of_pages():
    page_buttons_coords = [
        (screen_width * 0.51, screen_height * 0.82),
        (screen_width * 0.54, screen_height * 0.82),
        (screen_width * 0.57, screen_height * 0.82),
        (screen_width * 0.60, screen_height * 0.82),
        (screen_width * 0.63, screen_height * 0.82),
        (screen_width * 0.66, screen_height * 0.82)
    ]
    num_pages = 0
    for coord in page_buttons_coords:
        try:
            pyautogui.click(coord)
            num_pages += 1
        except:
            break
    return num_pages

def search_color_on_screen(target_color, current_row, timeout=2):
    start_time = time.time()
    while time.time() - start_time < timeout:
        screenshot = pyautogui.screenshot()
        width, height = screenshot.size
        for x in range(width):
            for y in range(height):
                if screenshot.getpixel((x, y)) == target_color:
                    perform_action_and_insert_value(x, y, target_color, current_row)
                    return (x, y)
        time.sleep(0.5)
    print(f"Timeout reached. Color {target_color} not detected on the screen.")
    return None

def go_to_next_page(current_page):
    page_buttons_coords = [
        (screen_width * 0.51, screen_height * 0.82),
        (screen_width * 0.54, screen_height * 0.82),
        (screen_width * 0.57, screen_height * 0.82),
        (screen_width * 0.60, screen_height * 0.82),
        (screen_width * 0.63, screen_height * 0.82),
        (screen_width * 0.66, screen_height * 0.82)
    ]
    if current_page < len(page_buttons_coords):
        pyautogui.click(page_buttons_coords[current_page])
    else:
        print("No more pages to navigate.")

def main():
    if focus_existing_window("Nuevo Telecrédito"):
        print("La ventana ya estaba abierta y fue enfocada.")
    time.sleep(0.5)
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.15))
    time.sleep(0.5)
    pyautogui.hotkey('tab')
    pyautogui.typewrite('agr')
    time.sleep(0.4)
    pyautogui.click(int(screen_width * 0.50), int(screen_height * 0.24))
    wait_for_color(int(screen_width * 0.39), int(screen_height * 0.74), (232, 246, 252))
    pyautogui.scroll(-500)
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.60))
    pyautogui.click(int(screen_width * 0.25), int(screen_height * 0.30))
    wait_for_color(int(screen_width * 0.316), int(screen_height * 0.804), (212, 102, 40))
    pyautogui.click(int(screen_width * 0.35), int(screen_height * 0.36))
    time.sleep(0.5)
    for i in range(4):
        pyautogui.hotkey('tab')
        pyautogui.hotkey('backspace')
    pyautogui.click(int(screen_width * 0.22), int(screen_height * 0.72))
    pyautogui.typewrite(two_yesterday)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(yesterday)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    num_pages = detect_number_of_pages()
    print(f"Number of pages detected: {num_pages}")

    current_row = start_row
    for valor in df[columna].iloc[fila_inicio - 1:]:
        match_found = False
        while not match_found:
            for page in range(num_pages):
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'f')
                pyautogui.typewrite(str(valor))
                pyautogui.press('enter')
                pyautogui.press('enter')
                time.sleep(0.5)
                if search_color_on_screen((255, 150, 50), current_row):
                    print(f"Match found for {valor} on page {page + 1}")
                    match_found = True
                    break
                else:
                    pyautogui.scroll(-4000)
                    time.sleep(0.5)
                    go_to_next_page(page)
                    time.sleep(0.5)
                    print(f"No match found for {valor} on page {page + 1}, moving to next page")
            if not match_found:
                print(f"Value {valor} not found, moving to the next cell")
                current_row += 1
                break

main()