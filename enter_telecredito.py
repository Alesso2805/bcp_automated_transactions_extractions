import pyautogui
import time
import webbrowser
import re
import pyperclip
import my_utils
from U_MMTC import search_color
from U_MMTS import move_mouse_to_segment
from U_CPC import check_pixel_color
import config as CONFIG
import traceback

def login_to_telecredito():
    bandera_repeat = False
    seconds_to_add = 1

    # Abrir Google Chrome con la URL específica
    print("Abriendo telecredito")
    url = "https://www.telecreditobcp.com/tlcnp/"
    webbrowser.open(url)

    # Esperar a que se cargue el navegador
    time.sleep(seconds_to_add+5)

    # Se realiza la navegación al nuevo telecredito
    move_mouse_to_segment("center")
    pyautogui.click()
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")

    print("Esperando 10 segundos.")
    time.sleep(seconds_to_add+10)

    result = check_pixel_color(1594, 666, (76, 76, 76))
    if result:
        print("Se intenta dar click en el botón naranja utilizando buscador de color (este es preventivo por si estamos en la página de sesión vencida).")
        move_mouse_to_segment("center")
        result = search_color((255, 120, 0), 10, 200)
        pyautogui.moveTo(result)
        pyautogui.click()
    else:
        print("No se encontró la pantalla de sesión vencida")

    time.sleep(seconds_to_add+3)

    print("Esperando que el botón naranja aparezca en pantalla (cualquier objeto naranja podría confundir a esta función)")
    my_utils.verify_color_on_screen((255, 120, 0), 8)

    time.sleep(seconds_to_add+1)

    print("Se da click para permitir navegación por tabs")
    move_mouse_to_segment("bottom_right")

    # Esperar 1 segundo
    time.sleep(seconds_to_add+1)

    # Dar 1 tab
    print("Se da tab")
    pyautogui.press('tab')

    # Esperar 1 segundo
    time.sleep(seconds_to_add+1)

    # Utilizar Ctrl + A y Ctrl + C
    print("Se inicia el proceso de copiado")
    pyautogui.hotkey('ctrl', 'a')
    # Esperar 1 segundo
    time.sleep(seconds_to_add+1)
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'c')

    print("Contenido copiado al portapapeles.")

    # Obtener texto del portapapeles
    texto = pyperclip.paste()

    print("start of copied text...")
    print('"', texto[0:50], '..."')

    try:
        # Utilizar una expresión regular para encontrar solo números individuales
        numeros_en_linea = re.findall(r'^\s*(\d)\s*$', texto, re.MULTILINE)

        # Convertir los números de cadenas a enteros
        lista_numeros = [int(num) for num in numeros_en_linea]

        # Mostrar la lista resultante
        print(lista_numeros)

        pin = CONFIG.TELECREDITO_PASSWORD
        print("Iniciando utilización de press pin.")
        result = my_utils.press_pin(lista_numeros, pin)
        print("Sequence of coordinates to press for pin:", result)

        for pos in result:
            pyautogui.moveTo(pos, duration=0.2)
            time.sleep(seconds_to_add/2+0.3)
            pyautogui.click()

        # Wait a sec
        time.sleep(seconds_to_add+1)

        # Click en el boton de "Continuar"
        print("Se intenta dar click en el botón naranja utilizando buscador de color.")
        move_mouse_to_segment("bottom_right")
        result = search_color((255, 120, 0), 10, 200)
        pyautogui.moveTo(result)
        pyautogui.click()

        time.sleep(seconds_to_add+10)
    except Exception as e:
        print("An error occurred:")
        traceback.print_exc()  # Imprime la traza completa (stacktrace)
        bandera_repeat = True

    for i in range(5):
        time.sleep(2)
        result = check_pixel_color(10, 260, (0, 42, 141))
        bandera_repeat = not result
        if result == True:
            break

    if bandera_repeat == True:
        close_browser() #CERRAMOS EL INTENTO ANTERIOR
        time.sleep(seconds_to_add+1)

        # Abrir Google Chrome con la URL específica
        print("Abriendo telecredito")
        url = "https://www.telecreditobcp.com/tlcnp/"
        webbrowser.open(url)

        # Esperar a que se cargue el navegador
        print("Esperando a que cargue el navegador")
        aux = 0
        while aux < 10:
            if pyautogui.pixelMatchesColor(1700, 800, (255, 255, 255)):
                break
            else:
                aux+=1
                time.sleep(seconds_to_add+1)

        time.sleep(seconds_to_add+3)

        # Verificar el color antes de dar clic en la primera posición
        print("Dando click en el botón que nos lleva al nuevo telecrédito")
        x1, y1 = 473, 525
        color1 = (255, 104, 46) # Naranja
        if pyautogui.pixelMatchesColor(x1, y1, color1):
            pyautogui.click(x1, y1)
            print(f"Clic realizado en ({x1}, {y1}) con el color correcto {color1}.")
        else:
            print(f"El color en ({x1}, {y1}) no coincide con {color1}. Abortando clic.")
            exit()

        print("Esperando 5 segundos.")
        time.sleep(seconds_to_add+5)
        print("Esperando 5 segundos.")
        time.sleep(seconds_to_add+5)
        print("Esperando 5 segundos.")
        time.sleep(seconds_to_add+5)

        result = check_pixel_color(1594, 666, (76, 76, 76))
        if result:
            print("Se intenta dar click en el botón naranja utilizando buscador de color (este es preventivo por si estamos en la página de sesión vencida).")
            move_mouse_to_segment("center")
            result = search_color((255, 120, 0), 10, 200)
            pyautogui.moveTo(result)
            pyautogui.click()
        else:
            print("No se encontró la pantalla de sesión vencida")

        # Se utiliza copiar y pegar para verificar el texto en pantalla.
        print("Se verifica que no hayamos entrado a la pantalla 'Tu Sesión ha expirado'")
        pyautogui.hotkey('ctrl', 'a')
        # Esperar 1 segundo
        time.sleep(seconds_to_add+1)
        pyautogui.hotkey('ctrl', 'c')
        # Obtener texto del portapapeles
        texto = pyperclip.paste()
        if 'expirado' in texto:
            print("ESTOY AQUI")
            #ir al centro
            #presionar el boton naranja que encontremos
            exit()


        print("Esperando que el botón naranja aparezca en pantalla (cualquier objeto naranja podría confundir a esta función)")
        my_utils.verify_color_on_screen((255, 120, 0), 8)

        # La pantalla blanca para prepararnos para el tab
        print("Se da click en la pantalla blanca para poder hacer tab.")
        x2, y2 = 1718, 452
        pyautogui.click(x2, y2)

        # Esperar 1 segundo
        time.sleep(seconds_to_add+1)

        # Dar 1 tab
        print("Se da tab")
        pyautogui.press('tab')

        # Esperar 1 segundo
        time.sleep(seconds_to_add+1)

        # Utilizar Ctrl + A y Ctrl + C
        print("Se inicia el proceso de copiado")
        pyautogui.hotkey('ctrl', 'a')
        # Esperar 1 segundo
        time.sleep(seconds_to_add+1)
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.hotkey('ctrl', 'c')

        print("Contenido copiado al portapapeles.")

        # Obtener texto del portapapeles
        texto = pyperclip.paste()

        print("start of copied text...")
        print('"' + texto[0:50] + '..."')

        # Utilizar una expresión regular para encontrar solo números individuales
        numeros_en_linea = re.findall(r'^\s*(\d)\s*$', texto, re.MULTILINE)

        # Convertir los números de cadenas a enteros
        lista_numeros = [int(num) for num in numeros_en_linea]

        # Mostrar la lista resultante
        print(lista_numeros)

        pin = CONFIG.TELECREDITO_PASSWORD
        result = my_utils.press_pin(lista_numeros, pin)
        print("Sequence of coordinates to press for pin:", result)

        for pos in result:
            pyautogui.moveTo(pos, duration=0.2)
            time.sleep(seconds_to_add+0.5)
            pyautogui.click()

        # Wait a sec
        time.sleep(seconds_to_add+1)

        # Click en el boton de "Continuar"
        print("Se intenta dar click en el botón naranja utilizando buscador de color.")
        move_mouse_to_segment("bottom_right")
        result = search_color((255, 120, 0), 10, 200)
        pyautogui.moveTo(result)
        pyautogui.click()




def close_browser():
    move_mouse_to_segment("corner_top_right")
    pyautogui.click()