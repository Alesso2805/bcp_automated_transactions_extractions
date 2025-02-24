from PIL import ImageGrab
import numpy as np
import time

def verify_color_on_screen(target_color, n_chances, tolerance=0, delay=1):
    """
    Verifies if the screen contains a specific color over multiple attempts.

    Parameters:
        target_color (tuple): The RGB color to search for (e.g., (255, 0, 0)).
        n_chances (int): Number of attempts to check the screen.
        tolerance (int, optional): Acceptable deviation for each RGB channel. Defaults to 0.
        delay (float, optional): Delay in seconds between attempts. Defaults to 1 second.

    Returns:
        tuple: (found (bool), count (int))
            - found: True if the color was found at least once, else False.
            - count: Number of times the color was found across all attempts.
    """
    count = 0
    for attempt in range(n_chances):
        # Capture the screen
        screenshot = ImageGrab.grab()
        screenshot_np = np.array(screenshot)

        # If the image is in RGBA, ignore the alpha channel
        if screenshot_np.shape[2] == 4:
            screenshot_np = screenshot_np[:, :, :3]

        # Define the lower and upper bounds for the target color with tolerance
        lower = np.array([max(c - tolerance, 0) for c in target_color])
        upper = np.array([min(c + tolerance, 255) for c in target_color])

        # Create a mask where the color matches within the tolerance
        mask = np.all((screenshot_np >= lower) & (screenshot_np <= upper), axis=2)

        # Check if any pixel matches
        if np.any(mask):
            count += 1
            print(f"Attempt {attempt + 1}: Color found.")
        else:
            print(f"Attempt {attempt + 1}: Color not found.")

        # Wait before the next attempt
        time.sleep(delay)

    found = count > 0
    return found, count

def press_pin(numbers_list, pin):
    """
    Given a list of numbers and a PIN, returns the corresponding list of pixel positions to press.

    :param numbers_list: List[int] - A list of 10 unique numbers representing the current mapping.
    :param pin: str - The PIN string composed of digits, e.g., "329873".
    :return: List[Tuple[int, int]] - A list of (x, y) pixel positions corresponding to the PIN digits.
    """
    # Define the pixel positions for positions 0 through 9
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

    # Create a mapping from number to its corresponding pixel position
    number_to_pixel = {number: pixel for number, pixel in zip(numbers_list, pixel_positions)}

    # Initialize a list to store the pixel positions to press
    pixels_to_press = []

    # Iterate over each digit in the PIN
    for digit_char in pin:
        # Convert the character to an integer
        digit = int(digit_char)

        # Retrieve the corresponding pixel position
        pixel = number_to_pixel.get(digit)

        if pixel is None:
            raise ValueError(f"Digit '{digit}' not found in the numbers list.")

        # Append the pixel position to the list
        pixels_to_press.append(pixel)

    return pixels_to_press
