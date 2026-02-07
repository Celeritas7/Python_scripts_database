import os
from PIL import Image

# Define input and output directories
input_folder = r"D:\####Technical\#Mechanical\Database\#Mind map\#3. SOM\Sample\main"
output_folder = r"D:\####Technical\#Mechanical\Database\#Mind map\#3. SOM\Sample"

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Loop through all image files
for filename in os.listdir(input_folder):
    if filename.lower().endswith((".jpg", ".jpeg", ".png")):
        img_path = os.path.join(input_folder, filename)
        img = Image.open(img_path)
        width, height = img.size

        # Define crop boxes
        left_box = (0, 0, width // 2, height)
        right_box = (width // 2, 0, width, height)

        # Crop and save
        left_half = img.crop(left_box)
        right_half = img.crop(right_box)

        base_name = os.path.splitext(filename)[0]
        left_half.save(os.path.join(output_folder, f"{base_name}_left.jpg"))
        right_half.save(os.path.join(output_folder, f"{base_name}_right.jpg"))

print("✅ All images split and saved.")