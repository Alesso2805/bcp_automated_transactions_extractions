import pyautogui
from PIL import Image
import time
import sys

def get_screen_size():
    screenWidth, screenHeight = pyautogui.size()
    return screenWidth, screenHeight

def is_within_bounds(x, y, screen_width, screen_height):
    return 0 <= x < screen_width and 0 <= y < screen_height

def get_pixel_color(x, y, screenshot):
    return screenshot.getpixel((x, y))

def color_matches(target_color, current_color):
    return target_color == current_color

def search_color(target_color, step_size=3, max_steps=1000, delay=0):
    """
    Searches for the target_color on the screen starting from the current mouse position,
    expanding outward in a "sonar pulse" pattern with the given step_size.

    :param target_color: Tuple of (R, G, B)
    :param step_size: The distance between each step in pixels
    :param max_steps: Maximum number of steps to prevent infinite loops
    :param delay: Optional delay between steps for visualization
    :return: Position tuple (x, y) if found, else None
    """
    # Capture the screen once to optimize performance
    screenshot = pyautogui.screenshot()

    # Get current mouse position
    start_x, start_y = pyautogui.position()
    print(f"Starting search from position: ({start_x}, {start_y})")
    print(f"Target color: {target_color}")

    screen_width, screen_height = get_screen_size()

    for step in range(1, max_steps + 1):
        d = step * step_size
        positions = []

        # Generate positions in a square "ring" around the starting point
        for dx in range(-d, d + 1, step_size):
            dy_top = -d
            dy_bottom = d
            x = start_x + dx
            y_top = start_y + dy_top
            y_bottom = start_y + dy_bottom

            positions.append((x, y_top))
            positions.append((x, y_bottom))

        for dy in range(-d + step_size, d, step_size):
            dx_left = -d
            dx_right = d
            y = start_y + dy
            x_left = start_x + dx_left
            x_right = start_x + dx_right

            positions.append((x_left, y))
            positions.append((x_right, y))

        # Remove duplicates
        positions = list(set(positions))

        #print(f"Step {step}: Checking {len(positions)} positions at distance {d}")

        for pos in positions:
            x, y = pos
            if is_within_bounds(x, y, screen_width, screen_height):
                current_color = get_pixel_color(x, y, screenshot)
                if color_matches(target_color, current_color):
                    print(f"Color {target_color} found at position: ({x}, {y})")
                    return (x, y)
            else:
                # Position is out of screen bounds; skip
                continue

        if delay > 0:
            time.sleep(delay)  # Optional delay for visualization

    print(f"Color {target_color} not found within {max_steps} steps.")
    return None

'''
if __name__ == "__main__":
    # Example usage
    # Define the target color you want to search for (R, G, B)
    target_color = (178, 195, 162)  # White color

    # Optional: Define step size and maximum steps
    step_size = 10  # Pixels
    max_steps = 100  # Adjust as needed
    delay = 0  # Seconds

    print("Please position your mouse at the starting point and press Enter...")
    input()

    result = search_color(target_color, step_size, max_steps, delay)
    if result:
        print(f"Target color found at: {result}")
        pyautogui.moveTo(result)
    else:
        print("Target color not found on the screen.")
'''