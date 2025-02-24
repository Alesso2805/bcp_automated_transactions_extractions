import pyautogui
from PIL import ImageGrab

def check_pixel_color(x, y, color):
    """
    Check if the pixel at the specified (x, y) position has the given color.

    Args:
        x (int): The x-coordinate of the pixel.
        y (int): The y-coordinate of the pixel.
        color (tuple): The RGB color to check, e.g., (76, 76, 76).

    Returns:
        bool: True if the pixel color matches, False otherwise.
    """
    # Take a screenshot of the entire screen
    screen = ImageGrab.grab()
    # Get the color of the pixel at the specified position
    pixel_color = screen.getpixel((x, y))
    # Check if the color matches
    return pixel_color == color