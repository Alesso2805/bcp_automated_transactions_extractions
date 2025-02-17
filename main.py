import os

import openpyxl
import pandas as pd
import re
import time
from datetime import datetime, timedelta
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Alignment
from pdfminer.image import align32

import config as CONFIG
import pygetwindow as gw
import win32gui
import win32con
import pyautogui
import pyperclip

two_yesterday = (datetime.now() - timedelta(days=3)).strftime("%d%m%Y")
yesterday = (datetime.now() - timedelta(days=2)).strftime("%d%m%Y")

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

def buscar_y_extraer_valores(valor, monto, df):
    print(f"Buscando valores para Dni={valor}, Cantidad={monto}, y Procesado='NO'")
    fila = df[(df['Dni'] == valor) & (df['Cantidad'] == monto) & (df['Procesado'] == 'NO')]
    if not fila.empty:
        index = fila.index[0]
        nombre = fila['Nombre'].values[0]
        codigo_flip = fila['Codigo Flip'].values[0]
        cantidad = fila['Cantidad'].values[0]

        # Mark the row as processed
        df.at[index, 'Procesado'] = 'SI'

        # Save the changes to the Excel file
        workbook = load_workbook(second_excel_path)
        sheet = workbook['Sheet1']
        sheet.cell(row=index + 2, column=df.columns.get_loc('Procesado') + 1).value = 'SI'
        workbook.save(second_excel_path)

        print(f"Valores extraídos: Nombre={nombre}, Codigo Flip={codigo_flip}, Cantidad={cantidad}")
        return nombre, codigo_flip, cantidad
    print(f"No se encontraron valores para Dni={valor}, Cantidad={monto}, y Procesado='NO'")
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
    wait_for_color(int(screen_width * 0.746), int(screen_height * 0.412), (96,108,127))
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

    monto = df.at[current_row - 1, 'Unnamed: 3']
    print(f"Obteniendo valores del segundo Excel para Dni={copied_value} y Monto={monto}")
    nombre, codigo_flip, cantidad = buscar_y_extraer_valores(copied_value, monto, df_second)
    if nombre and codigo_flip and cantidad:
        df.at[current_row, 'Unnamed: 9'] = nombre
        df.at[current_row, 'Unnamed: 15'] = codigo_flip
        df.at[current_row, 'Unnamed: 13'] = cantidad
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 9') + 1).value = nombre
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 15') + 1).value = codigo_flip
        sheet.cell(row=current_row + 1, column=df.columns.get_loc('Unnamed: 13') + 1).value = cantidad
        print(f"Nombre '{nombre}', Codigo Flip '{codigo_flip}' y Cantidad '{cantidad}' insertados en la fila {current_row}")
    else:
        columns_to_check = ['Unnamed: 9', 'Unnamed: 15', 'Unnamed: 13']
        for col in columns_to_check:
            df.at[current_row, col] = 'PENDIENTE'
            sheet.cell(row=current_row + 1, column=df.columns.get_loc(col) + 1).value = 'PENDIENTE'
        print(f"Valores no encontrados, se insertó 'PENDIENTE' en las columnas designadas para la fila {current_row}")

    workbook.save(excel_path)
    print("DataFrame saved to Excel file")
    pyautogui.click(int(screen_width * 0.21), int(screen_height * 0.289))
    time.sleep(0.5)

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


def go_to_next_page(current_page, num_pages):
    base_x = screen_width * 0.51
    y = 881  # Fixed y-coordinate for the buttons
    target_color = (255, 120, 0)  # RGB color to detect the button for page 1

    # Find the button for page 1
    for x in range(int(screen_width * 0.51), screen_width, 1):
        if pyautogui.pixelMatchesColor(x, y, target_color):
            base_x = x
            pyautogui.moveTo(base_x, y)
            break
    else:
        print("Button for page 1 not found.")
        return None, None

    # Click the button for the current page
    x = base_x + (current_page * 60)
    if pyautogui.onScreen(x, y):
        pyautogui.moveTo(x, y)
        pyautogui.click(x, y)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.hotkey('enter')
        pyautogui.hotkey('enter')
        time.sleep(0.5)
    else:
        print(f"Button for page {current_page + 1} is not on screen.")

    return base_x, y

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
    time.sleep(0.1)
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
    wait_for_color(int(screen_width * 0.823), int(screen_height * 0.878), (96,108,127))
    pyautogui.scroll(-4000)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    copied_text = pyperclip.paste()
    page_numbers = re.findall(r'\d+', copied_text)
    if page_numbers:
        num_pages = int(page_numbers[-1])
        print(f"Number of pages detected: {num_pages}")
    else:
        print("No page numbers found in the copied text.")

    current_row = start_row
    base_x, y = go_to_next_page(0, num_pages)  # Get the coordinates for page 1
    if base_x is None or y is None:
        return

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
                    go_to_next_page(page, num_pages)
                    time.sleep(0.5)
                    print(f"No match found for {valor} on page {page + 1}, moving to next page")
            if not match_found:
                print(f"Value {valor} not found, moving to the next cell")
                pyautogui.moveTo(base_x, y)  # Click on the coordinates of page 1
                pyautogui.click(base_x, y)
                current_row += 1
                break

    # Filter rows with 'PENDIENTE' and save to a new sheet
    pendientes_df = df[df.isin(['PENDIENTE']).any(axis=1)]
    custom_headers = ['Fecha', 'Fecha Valuta', 'Descripción Operación', 'Monto', 'Sucursal - Agencia', 'N° Operación',
                      'Usuario', 'x', 'x', 'NOMBRE', 'x', 'DNI', 'x', 'Monto', 'x', 'CODIGO FLIP']
    pendientes_df.columns = custom_headers

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
        pendientes_df.to_excel(writer, sheet_name='PENDIENTES', index=False)

    # Load the workbook and select the sheet
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook['PENDIENTES']

    # Center align all cells and adjust column widths
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    # Save the workbook
    workbook.save(excel_path)
    print("Rows with 'PENDIENTE' moved to a new sheet named 'PENDIENTES' with custom headers and formatted cells")

main()
