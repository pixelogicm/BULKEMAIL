# BULKEMAIL

A bulk email sender application with purchase order image blurring functionality.

## Features

### Image Blurring for Purchase Orders
This application now includes advanced image processing capabilities to blur text areas in purchase order images. This is useful for:
- Protecting sensitive information while maintaining document structure
- Creating samples for demonstration purposes
- Obscuring confidential details in shared documents

#### Supported Text Areas
The application can blur the following areas in purchase order images:
- Company addresses (bill-to and ship-to)
- Invoice/PO numbers and dates
- Item details and descriptions
- Financial totals and pricing
- Contact information and terms
- Any other major text blocks

#### Blur Features
- **Moderate Blur Effect**: Applies a light blur that obscures details while keeping document structure visible
- **Adjustable Blur Strength**: Customizable blur intensity (range: 5-30)
- **Predefined Text Areas**: Automatically identifies common purchase order sections
- **Auto-Detection Mode**: Optional automatic text area detection using computer vision
- **Multiple Output Formats**: Supports PNG, JPEG, and other common image formats

## Installation

1. Clone the repository:
```bash
git clone https://github.com/pixelogicm/BULKEMAIL.git
cd BULKEMAIL
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. For GUI support (optional):
```bash
sudo apt-get install python3-tk  # On Ubuntu/Debian
```

## Usage

### Command Line Mode
Blur a purchase order image directly from the command line:

```bash
python3 bulk_email_sender_manual_login.py <image_path>
```

Example:
```bash
python3 bulk_email_sender_manual_login.py invoice.png
```

### GUI Mode (when display available)
Run the application with GUI:

```bash
python3 bulk_email_sender_manual_login.py
```

### Python Module Usage
Use the image blurring functionality in your own code:

```python
from image_blur import PurchaseOrderBlurrer

# Create blurrer instance
blurrer = PurchaseOrderBlurrer()

# Set custom blur strength (optional)
blurrer.set_blur_strength(20)

# Blur an image
result_path = blurrer.blur_purchase_order(
    "invoice.png",
    "invoice_blurred.png", 
    use_auto_detection=False
)

print(f"Blurred image saved to: {result_path}")
```

## Testing

Create a sample purchase order and test the blurring:

```bash
python3 create_sample_image.py
python3 verify_blur.py
```

## Requirements

- Python 3.7+
- Pillow (PIL) for image processing
- OpenCV for computer vision operations
- NumPy for numerical operations
- Tkinter for GUI (optional)

## License

This project is open source. Please ensure you have permission to process any images containing sensitive or confidential information.