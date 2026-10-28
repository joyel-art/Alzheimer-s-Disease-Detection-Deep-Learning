from PIL import Image, ImageDraw
import os

# Create a colorful image
color_img = Image.new('RGB', (224, 224), color='red')
draw = ImageDraw.Draw(color_img)
draw.rectangle([50, 50, 150, 150], fill='blue')
color_img.save('color_test.jpg')

# Create a grayscale image (simulating MRI)
gray_img = Image.new('RGB', (224, 224), color='black')
draw_gray = ImageDraw.Draw(gray_img)
# Draw a brain-like shape (ellipse) with some gray intensity
draw_gray.ellipse([60, 40, 160, 180], fill=(150, 150, 150))
gray_img.save('gray_test.jpg')

print("Created color_test.jpg and gray_test.jpg")
