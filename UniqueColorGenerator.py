import colorsys
import random

def distance(color1, color2):
    r1, g1, b1 = color1
    r2, g2, b2 = color2
    return ((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2) ** 0.5

def generate_random_color():
    # Generate a random hue
    h = random.random()
    # Generate a random saturation within the recommended range
    s = random.uniform(0.6, 1.0)
    # Generate a random brightness within the recommended range
    v = random.uniform(0.6, 1.0)
    # Convert the HSV color to RGB
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    # Ensure the color has enough contrast with white
    while (r * 0.299 + g * 0.587 + b * 0.114) > 0.5:
        v = random.uniform(0.3, 0.6)
        r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return (int(r * 255), int(g * 255), int(b * 255))

class UniqueColorGenerator:
    def __init__(self, min_distance=100):
        self.generated_colors = []
        self.min_distance = min_distance

    def generate_unique_color(self):
        while True:
            new_color = generate_random_color()
            if all(distance(new_color, color) >= self.min_distance for color in self.generated_colors):
                self.generated_colors.append(new_color)
                # print(f"New color: {new_color}")  # Print new color to console
                return new_color

