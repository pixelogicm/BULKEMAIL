#!/usr/bin/env python3
"""
Image Blur Module for Purchase Order Processing
Blurs text areas in purchase order images to obscure sensitive information
while maintaining document structure.
"""

import cv2
import numpy as np
from PIL import Image, ImageFilter
import os
from typing import List, Tuple, Optional


class PurchaseOrderBlurrer:
    """
    A class to blur specific text areas in purchase order images.
    Provides moderate blur effect to maintain document structure while obscuring details.
    """
    
    def __init__(self):
        """Initialize the blurrer with default settings."""
        self.blur_strength = 15  # Moderate blur - not too deep
        self.gaussian_kernel_size = (15, 15)  # Kernel size for Gaussian blur
        
    def load_image(self, image_path: str) -> Optional[np.ndarray]:
        """
        Load an image from file path.
        
        Args:
            image_path (str): Path to the image file
            
        Returns:
            np.ndarray: Loaded image as numpy array or None if failed
        """
        try:
            image = cv2.imread(image_path)
            if image is None:
                print(f"Error: Could not load image from {image_path}")
                return None
            return image
        except Exception as e:
            print(f"Error loading image: {e}")
            return None
    
    def detect_text_regions(self, image: np.ndarray) -> List[Tuple[int, int, int, int]]:
        """
        Detect text regions in the image using OpenCV text detection.
        
        Args:
            image (np.ndarray): Input image
            
        Returns:
            List[Tuple[int, int, int, int]]: List of bounding boxes (x, y, w, h) for text regions
        """
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Apply morphological operations to detect text regions
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 2))
        morph = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel)
        
        # Find contours
        contours, _ = cv2.findContours(morph, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        text_regions = []
        height, width = image.shape[:2]
        
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            
            # Filter out very small regions and regions that are too large
            if (w > 30 and h > 10 and 
                w < width * 0.8 and h < height * 0.1 and
                w * h > 300):  # Minimum area
                text_regions.append((x, y, w, h))
        
        return text_regions
    
    def define_common_text_areas(self, image: np.ndarray) -> List[Tuple[int, int, int, int]]:
        """
        Define common text areas found in purchase orders based on typical layout.
        
        Args:
            image (np.ndarray): Input image
            
        Returns:
            List[Tuple[int, int, int, int]]: List of predefined text areas (x, y, w, h)
        """
        height, width = image.shape[:2]
        
        # Define typical areas for purchase order elements
        text_areas = []
        
        # Header area (company name, PO number)
        text_areas.append((0, 0, width, int(height * 0.15)))
        
        # Addresses area (bill to, ship to)
        text_areas.append((0, int(height * 0.15), width, int(height * 0.25)))
        
        # Items table area
        text_areas.append((0, int(height * 0.35), width, int(height * 0.45)))
        
        # Totals area
        text_areas.append((int(width * 0.6), int(height * 0.75), int(width * 0.4), int(height * 0.15)))
        
        # Footer area (terms, contact info)
        text_areas.append((0, int(height * 0.85), width, int(height * 0.15)))
        
        return text_areas
    
    def apply_blur_to_region(self, image: np.ndarray, x: int, y: int, w: int, h: int) -> np.ndarray:
        """
        Apply blur to a specific region of the image.
        
        Args:
            image (np.ndarray): Input image
            x, y, w, h (int): Bounding box coordinates
            
        Returns:
            np.ndarray: Image with blurred region
        """
        # Ensure coordinates are within image bounds
        height, width = image.shape[:2]
        x = max(0, min(x, width - 1))
        y = max(0, min(y, height - 1))
        w = max(1, min(w, width - x))
        h = max(1, min(h, height - y))
        
        # Extract the region
        region = image[y:y+h, x:x+w]
        
        # Apply Gaussian blur
        blurred_region = cv2.GaussianBlur(region, self.gaussian_kernel_size, self.blur_strength)
        
        # Replace the region in the original image
        result = image.copy()
        result[y:y+h, x:x+w] = blurred_region
        
        return result
    
    def blur_purchase_order(self, image_path: str, output_path: str = None, 
                           use_auto_detection: bool = False) -> str:
        """
        Blur text areas in a purchase order image.
        
        Args:
            image_path (str): Path to input image
            output_path (str): Path for output image (optional)
            use_auto_detection (bool): Whether to use automatic text detection
            
        Returns:
            str: Path to the blurred image
        """
        # Load the image
        image = self.load_image(image_path)
        if image is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        # Generate output path if not provided
        if output_path is None:
            name, ext = os.path.splitext(image_path)
            output_path = f"{name}_blurred{ext}"
        
        # Get text areas to blur
        if use_auto_detection:
            text_areas = self.detect_text_regions(image)
            print(f"Auto-detected {len(text_areas)} text regions")
        else:
            text_areas = self.define_common_text_areas(image)
            print(f"Using {len(text_areas)} predefined text areas")
        
        # Apply blur to each text area
        blurred_image = image.copy()
        for x, y, w, h in text_areas:
            blurred_image = self.apply_blur_to_region(blurred_image, x, y, w, h)
            print(f"Blurred region: ({x}, {y}, {w}, {h})")
        
        # Save the result
        success = cv2.imwrite(output_path, blurred_image)
        if not success:
            raise ValueError(f"Could not save blurred image to: {output_path}")
        
        print(f"Blurred image saved to: {output_path}")
        return output_path
    
    def set_blur_strength(self, strength: int):
        """
        Set the blur strength (higher = more blurred).
        
        Args:
            strength (int): Blur strength (recommended range: 5-30)
        """
        self.blur_strength = max(1, min(strength, 50))
        # Adjust kernel size based on strength
        kernel_size = max(3, self.blur_strength)
        if kernel_size % 2 == 0:  # Ensure odd kernel size
            kernel_size += 1
        self.gaussian_kernel_size = (kernel_size, kernel_size)
        print(f"Blur strength set to: {self.blur_strength}")


def main():
    """Main function for testing the blur functionality."""
    blurrer = PurchaseOrderBlurrer()
    
    # Test with the sample image
    input_image = "image1.png"
    if not os.path.exists(input_image):
        print(f"Error: {input_image} not found. Please run create_sample_image.py first.")
        return
    
    try:
        # Apply moderate blur with predefined areas
        output_path = blurrer.blur_purchase_order(
            input_image, 
            "image1_blurred.png", 
            use_auto_detection=False
        )
        print(f"\nSuccess! Blurred purchase order saved as: {output_path}")
        
        # Also test with auto-detection
        output_path_auto = blurrer.blur_purchase_order(
            input_image, 
            "image1_auto_blurred.png", 
            use_auto_detection=True
        )
        print(f"Auto-detected blur version saved as: {output_path_auto}")
        
    except Exception as e:
        print(f"Error processing image: {e}")


if __name__ == "__main__":
    main()