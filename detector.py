import pyautogui

def detect_click_coordinates():
    print("Haz clic en cualquier lugar para obtener las coordenadas. Presiona Ctrl+C para salir.")
    try:
        while True:
            x, y = pyautogui.position()
            print(f"Coordenadas del clic: ({x}, {y})")
            pyautogui.sleep(1)  # Espera 1 segundo antes de capturar la siguiente posición
    except KeyboardInterrupt:
        print("Detección de coordenadas finalizada.")

detect_click_coordinates()