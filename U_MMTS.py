import pyautogui
import time

def move_mouse_to_segment(segment, duration=0.5):
    """
    Mueve el ratón al centro del segmento especificado en la pantalla.

    Parámetros:
        segment (str): Segmento de la pantalla ('top_left', 'top_right', 'bottom_left', 'bottom_right').

    """
    # Obtener tamaño de la pantalla
    screen_width, screen_height = pyautogui.size()

    # Calcular posiciones centrales de los segmentos
    positions = {
        'top_left': (screen_width // 4, screen_height // 4),
        'top_right': (3 * screen_width // 4, screen_height // 4),
        'bottom_left': (screen_width // 4, 3 * screen_height // 4),
        'bottom_right': (3 * screen_width // 4, 3 * screen_height // 4),
        'center': (screen_width // 2, screen_height // 2),
        'corner_top_left': (10, 10),
        'corner_top_right': (screen_width - 10, 10),
        'corner_bottom_left': (10, screen_height - 10),
        'corner_bottom_right': (screen_width - 10, screen_height - 10)
    }

    # Validar segmento
    if segment not in positions:
        raise ValueError("Segmento inválido. Usa: 'top_left', 'top_right', 'bottom_left' o 'bottom_right'.")

    # Mover el ratón
    x, y = positions[segment]
    pyautogui.moveTo(x, y, duration=duration)



# Ejemplo de uso
'''
try:
    move_mouse_to_segment('top_left')  # Cambia el segmento aquí según sea necesario
    move_mouse_to_segment('top_right')  # Cambia el segmento aquí según sea necesario
    move_mouse_to_segment('bottom_left')  # Cambia el segmento aquí según sea necesario
    move_mouse_to_segment('bottom_right')  # Cambia el segmento aquí según sea necesario
except Exception as e:
    print(f"Error: {e}")
'''