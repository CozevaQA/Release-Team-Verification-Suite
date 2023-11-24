from PIL import Image

def compare_images(image1_path, image2_path):
    # Open the images
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path)

    # Ensure the images have the same size
    if image1.size != image2.size:
        return False

    # Compare pixel values
    pixel_pairs = zip(image1.getdata(), image2.getdata())
    for pixel1, pixel2 in pixel_pairs:
        if pixel1 != pixel2:
            return False

    return True

# Example usage
image1_path = "assets/images/GreenDot.png"
image2_path = "assets/images/OrangeDot.png"

result = compare_images(image1_path, image2_path)

if result:
    print("The images are identical.")
else:
    print("The images are different.")