from contextlib import contextmanager
import datetime
import config as CONFIG
import pyautogui

def obtener_dias_contiguos(custom_date=None):
    # En este caso recibiría un estilo dd/mm/yyyy si pasamos una fecha
    with open(CONFIG.DIAS_INHABILES_FILE) as f:
        # Leemos y limpiamos las líneas no vacías
        fechas = [line.strip() for line in f if line.strip()]
    # Convertimos a objetos date (formato dd/mm/yyyy)
    fechas = sorted(datetime.datetime.strptime(f, "%d/%m/%Y").date() for f in fechas)

    # Separamos en bloques de días consecutivos
    bloques = []
    bloque = [fechas[0]]
    for fecha in fechas[1:]:
        if (fecha - bloque[-1]).days == 1:
            bloque.append(fecha)
        else:
            bloques.append(bloque)
            bloque = [fecha]
    bloques.append(bloque)

    # Definimos la fecha de referencia
    if custom_date:
        if isinstance(custom_date, str):
            try:
                referencia = datetime.datetime.strptime(custom_date, "%d/%m/%Y").date()
            except ValueError:
                raise ValueError("La fecha personalizada debe estar en formato dd/mm/yyyy")
        elif isinstance(custom_date, datetime.date):
            referencia = custom_date
        else:
            raise TypeError("custom_date debe ser un string o un objeto datetime.date")
    else:
        referencia = datetime.date.today()

    # Buscamos el bloque contiguo "inmediato" anterior a la fecha de referencia.
    # Si la referencia está dentro de un bloque, devolvemos solo los días anteriores a ella.
    # Si la referencia es exactamente 1 día después del último día del bloque, se devuelve el bloque completo.
    for b in bloques:
        if b[0] <= referencia <= b[-1]:
            # Si la fecha de referencia está dentro del bloque, retornamos solo los días anteriores
            return [d for d in b if d < referencia]
        elif referencia == b[-1] + datetime.timedelta(days=1):
            return b
    return []




def obtener_fecha_limite(style="", do_not_jump_to_last=False, simple_time_delta=0, custom_date=None, recursions=0) -> str:
    """
    Calcula la fecha límite basándose en una fecha base y considerando los días inhábiles definidos en un archivo.
    
    Parámetros:
      - style (str): Estilo de formato para la salida. Si es "/", retorna 'dd/mm/yyyy'; sino 'ddmmyyyy'.
      - do_not_jump_to_last (bool): Si True, no retrocede un día adicional al encontrar días inhábiles.
      - simple_time_delta (int): Días a sumar (positivo) o restar (negativo) a la fecha calculada.
      - custom_date (str): Fecha base en 'dd/mm/yyyy' o 'ddmmyyyy'. Si no se da, usa la fecha actual.
      - recursions (int): Veces que se aplica la función recursivamente para navegar hacia fechas anteriores.
    
    Retorna:
      Una cadena con la fecha límite formateada según 'style'.
    """
    # Leer días inhábiles desde el archivo
    try:
        with open(CONFIG.DIAS_INHABILES_FILE, 'r', encoding='utf-8') as archivo:
            lineas = archivo.readlines()
        dias_inhabiles_str = [linea.strip() for linea in lineas if linea.strip()]
    except FileNotFoundError:
        dias_inhabiles_str = []
    
    # Convertir fechas a objetos date
    dias_inhabiles = []
    for fecha in dias_inhabiles_str:
        try:
            dia = datetime.datetime.strptime(fecha, '%d/%m/%Y').date()
            dias_inhabiles.append(dia)
        except ValueError:
            continue
    
    # Determinar fecha base
    if custom_date:
        try:
            fecha_base = datetime.datetime.strptime(custom_date, '%d/%m/%Y').date()
        except ValueError:
            try:
                fecha_base = datetime.datetime.strptime(custom_date, '%d%m%Y').date()
            except ValueError:
                raise ValueError("Formato de fecha inválido. Use 'dd/mm/yyyy' o 'ddmmyyyy'.")
    else:
        fecha_base = datetime.datetime.now().date()
    
    # Llamar a la función interna con parámetros adecuados
    fecha_limite = _obtener_fecha_limite(
        fecha_base=fecha_base,
        dias_inhabiles=dias_inhabiles,
        do_not_jump_to_last=do_not_jump_to_last,
        simple_time_delta=simple_time_delta,
        recursions=recursions
    )
    
    # Formatear resultado
    return fecha_limite.strftime('%d/%m/%Y' if style == "/" else '%d-%m-%Y' if style == "-" else '%d%m%Y')


def _obtener_fecha_limite(fecha_base, dias_inhabiles, do_not_jump_to_last, simple_time_delta, recursions) -> datetime.date:
    """
    Función interna que maneja cálculos de fechas con objetos date.
    """
    ayer = fecha_base - datetime.timedelta(days=1)
    fecha_limite = ayer
    
    # Ajustar por días inhábiles
    if ayer in dias_inhabiles:
        fecha_actual = ayer
        while fecha_actual in dias_inhabiles:
            fecha_anterior = fecha_actual - datetime.timedelta(days=1)
            if fecha_anterior in dias_inhabiles:
                fecha_limite = fecha_anterior
                fecha_actual = fecha_anterior
            else:
                break
        if not do_not_jump_to_last:
            fecha_limite -= datetime.timedelta(days=1)
    
    # Manejar recursión
    if recursions > 0:
        fecha_limite = _obtener_fecha_limite(
            fecha_base=fecha_limite,
            dias_inhabiles=dias_inhabiles,
            do_not_jump_to_last=do_not_jump_to_last,
            simple_time_delta=simple_time_delta,
            recursions=recursions - 1
        )
    else:
        fecha_limite += datetime.timedelta(days=simple_time_delta)
    
    return fecha_limite



def es_dia_habil(custom_date=None):
    # Determinar la fecha de referencia
    if custom_date:
        if isinstance(custom_date, str):
            try:
                fecha = datetime.datetime.strptime(custom_date, "%d/%m/%Y").date()
            except ValueError:
                raise ValueError("La fecha debe estar en formato dd/mm/yyyy")
        elif isinstance(custom_date, datetime.date):
            fecha = custom_date
        else:
            raise TypeError("custom_date debe ser un string o un objeto datetime.date")
    else:
        fecha = datetime.date.today()

    # Cargar y convertir las fechas inhábiles desde el archivo
    with open(CONFIG.DIAS_INHABILES_FILE, 'r', encoding='utf-8') as archivo:
        lineas = [linea.strip() for linea in archivo if linea.strip()]
    dias_inhabiles = [datetime.datetime.strptime(f, "%d/%m/%Y").date() for f in lineas]

    return fecha not in dias_inhabiles

def es_dia_after_rest(custom_date=None):
    # Determinar la fecha de referencia
    if custom_date:
        if isinstance(custom_date, str):
            try:
                fecha = datetime.datetime.strptime(custom_date, "%d/%m/%Y").date()
            except ValueError:
                raise ValueError("La fecha debe estar en formato dd/mm/yyyy")
        elif isinstance(custom_date, datetime.date):
            fecha = custom_date
        else:
            raise TypeError("custom_date debe ser un string o un objeto datetime.date")
    else:
        fecha = datetime.date.today()

    # Cargar y convertir las fechas inhábiles desde el archivo
    with open(CONFIG.DIAS_INHABILES_FILE, 'r', encoding='utf-8') as archivo:
        lineas = [linea.strip() for linea in archivo if linea.strip()]
    dias_inhabiles = [datetime.datetime.strptime(f, "%d/%m/%Y").date() for f in lineas]

    yesterday = fecha - datetime.timedelta(days=1)

    if yesterday in dias_inhabiles and fecha not in dias_inhabiles:
        return True
    else:
        return False

@contextmanager
def sin_failsafe():
    # Guardamos el estado original del failsafe
    estado_original = pyautogui.FAILSAFE
    # Desactivamos el failsafe temporalmente
    pyautogui.FAILSAFE = False
    try:
        yield
    finally:
        # Restauramos el estado original, asegurando que siempre se reactive
        pyautogui.FAILSAFE = estado_original

# Ejemplo de uso
if __name__ == '__main__':
    fecha_resultante = obtener_fecha_limite()
    print("La fecha límite es:", fecha_resultante)
