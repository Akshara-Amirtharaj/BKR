import base64

# Path to your image file (e.g., 'signature.png')
image_path = "signature.png"

# Read and encode the image
with open(image_path, "rb") as image_file:
    base64_encoded = base64.b64encode(image_file.read()).decode('utf-8')

# Print the Base64-encoded string
print(base64_encoded)
