#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 20:47:04 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_crop_images.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_crop_images.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from glob import glob

import cv2
import numpy as np
from PIL import Image, ImageOps
from ..utils._wait import wait
from ..utils._remove import remove

class ImageCropper:
    @staticmethod
    def _get_image_dpi(img):
        """Get the DPI of an image, defaulting to 96 if not available."""
        try:
            dpi_x, dpi_y = img.info.get("dpi", (96, 96))
            # Some images might return 0 for DPI
            return max(dpi_x, 1)  # Ensure DPI is at least 1
        except:
            return 96

    @staticmethod
    def _mm_to_pixels(mm, dpi):
        """Convert millimeters to pixels at the given DPI."""
        inches = mm / 25.4  # 1 inch = 25.4 mm
        return int(inches * dpi)

    @staticmethod
    def _find_content_area(image_path):
        """
        Find the content area in an image by identifying non-white pixels.
        This function detects the bounding rectangle around all non-white
        content in the image.
        Args:
            image_path (str): Path to the image file
        Returns:
            tuple: (x, y, width, height) coordinates of the bounding rectangle
        Raises:
            FileNotFoundError: If the image cannot be read
        """
        # Check if it's a GIF file
        if image_path.lower().endswith(".gif"):
            try:
                # Use PIL for GIF files
                with Image.open(image_path) as img:
                    # Convert to RGB for consistent processing
                    img = img.convert("RGB")
                    # Convert to numpy array for processing
                    img_array = np.array(img)
                    # Convert to grayscale
                    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
                    # Threshold the image to get all non-white areas
                    _, thresh = cv2.threshold(
                        gray, 254, 255, cv2.THRESH_BINARY_INV
                    )
                    # Find all non-zero points
                    points = cv2.findNonZero(thresh)
                    # Get the bounding rectangle
                    if points is not None:
                        x, y, w, h = cv2.boundingRect(points)
                        return x, y, w, h
                    else:
                        # If no points found, return the whole image
                        h, w = img_array.shape[:2]
                        return 0, 0, w, h
            except Exception as e:
                print(f"Error processing GIF with PIL: {e}")
                # Fall back to default dimensions
                return 0, 0, 100, 100

        # For non-GIF files, use OpenCV
        img = cv2.imread(image_path)
        if img is None:
            raise FileNotFoundError(f"Unable to read image file: {image_path}")

        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # Threshold the image to get all non-white areas.
        # Set the threshold close to 255 to exclude only white or nearly white areas.
        _, thresh = cv2.threshold(gray, 254, 255, cv2.THRESH_BINARY_INV)
        # Find all non-zero points (colored or non-white pixels)
        points = cv2.findNonZero(thresh)
        # Get the bounding rectangle around the non-zero points
        if points is not None:
            x, y, w, h = cv2.boundingRect(points)
            return x, y, w, h
        else:
            # If no points found, return the whole image
            h, w = img.shape[:2]
            return 0, 0, w, h

    @classmethod
    def _crop_image(cls, input_path, margin=30):
        """
        Crop an image to its content area plus a specified margin.
        This function detects the content area and crops the image around it,
        adding a margin of the specified number of pixels.
        Args:
            input_path (str): Path to the input image
            margin (int, optional): Margin in pixels to add around content. Defaults to 30.
        Returns:
            PIL.Image: Cropped image
        Raises:
            FileNotFoundError: If the image cannot be read
        """
        # Check if it's a GIF file
        if input_path.lower().endswith(".gif"):
            try:
                # Use PIL for GIF files
                with Image.open(input_path) as img:
                    # Check if GIF has multiple frames
                    is_animated = hasattr(img, "n_frames") and img.n_frames > 1
                    # Find content area
                    x, y, w, h = cls._find_content_area(input_path)
                    # Calculate coordinates with margin
                    x_start = max(x - margin, 0)
                    y_start = max(y - margin, 0)
                    x_end = min(x + w + margin, img.width)
                    y_end = min(y + h + margin, img.height)

                    if is_animated:
                        # Process all frames for animated GIFs
                        frames = []
                        for frame_idx in range(img.n_frames):
                            img.seek(frame_idx)
                            frame_copy = img.copy()
                            cropped_frame = frame_copy.crop(
                                (x_start, y_start, x_end, y_end)
                            )
                            frames.append(cropped_frame)
                        # Return the first frame for now (we'll handle animation later)
                        return frames[0]
                    else:
                        # Crop the image (single frame)
                        cropped_img = img.crop(
                            (x_start, y_start, x_end, y_end)
                        )
                        return cropped_img
            except Exception as e:
                print(f"Error processing GIF with PIL: {e}")
                # Continue with OpenCV method as fallback

        # For non-GIF files, use OpenCV
        img = cv2.imread(input_path)
        if img is None:
            raise FileNotFoundError(f"Unable to read image file: {input_path}")

        # Find the content area
        x, y, w, h = cls._find_content_area(input_path)
        # Calculate the coordinates with margin, clamping to the image boundaries
        x_start = max(x - margin, 0)
        y_start = max(y - margin, 0)
        x_end = min(x + w + margin, img.shape[1])
        y_end = min(y + h + margin, img.shape[0])

        # Crop the image using the bounding rectangle with margin
        cropped_img = img[y_start:y_end, x_start:x_end]
        # Convert OpenCV image to PIL Image
        pil_img = Image.fromarray(cv2.cvtColor(cropped_img, cv2.COLOR_BGR2RGB))
        return pil_img

    @staticmethod
    def _save_image(image, output_path):
        """
        Save an image to the specified output path.
        This function ensures the output path has a valid extension (defaults to PNG)
        and saves the image.
        Args:
            image (PIL.Image): The image to save
            output_path (str): The path where the image will be saved
        """
        # Check if output path has extension
        base, ext = os.path.splitext(output_path)
        if not ext:
            # Default to PNG if no extension
            output_path = f"{base}.png"

        # Save the image
        image.save(output_path)

        wait(
            wait_condition_func=lambda: os.path.exists(
                output_path
            ),
            success_msg=f"Image saved to: {output_path}",
            failure_msg=f"Failed to save image to: {output_path}",
        )

    @classmethod
    def add_margins(
        cls, image, width=None, height=None, width_mm=None, height_mm=None
    ):
        """
        Add white margins to an image to achieve target dimensions.
        Args:
            image (PIL.Image): The input image
            width (int, optional): Target width in pixels
            height (int, optional): Target height in pixels
            width_mm (float, optional): Target width in millimeters
            height_mm (float, optional): Target height in millimeters
        Returns:
            PIL.Image: Image with added margins
        """
        # Get the DPI of the original image
        dpi = cls._get_image_dpi(image)

        # Convert mm to pixels if provided
        if width_mm is not None:
            width = cls._mm_to_pixels(width_mm, dpi)
        if height_mm is not None:
            height = cls._mm_to_pixels(height_mm, dpi)

        def _add_margin_up_to_width(image, width=None):
            """Add horizontal margins to reach target width"""
            if width is not None:
                image_width = image.width
                if image_width < width:
                    left_margin = (width - image_width) // 2
                    right_margin = width - image_width - left_margin
                    return ImageOps.expand(
                        image,
                        (left_margin, 0, right_margin, 0),
                        fill=(255, 255, 255, 1),
                    )
            return image

        def _add_margin_up_to_height(image, height=None):
            """Add vertical margins to reach target height"""
            if height is not None:
                image_height = image.height
                if image_height < height:
                    top_margin = (height - image_height) // 2
                    bottom_margin = height - image_height - top_margin
                    return ImageOps.expand(
                        image,
                        (0, top_margin, 0, bottom_margin),
                        fill=(255, 255, 255, 1),
                    )
            return image

        # Ensure image is in RGBA mode for transparent margins
        if image.mode != "RGBA":
            image = image.convert("RGBA")

        image = _add_margin_up_to_width(image, width)
        image = _add_margin_up_to_height(image, height)

        return image

    @classmethod
    def crop_images(
        cls,
        input_paths,
        margin=30,
        width_mm=None,
        height_mm=None,
        output_dir=None,
        keep_orig=True,
        force=True,
    ):
        """
        Crop multiple images to their content area plus margin.
        This function processes one or more images, detecting content areas,
        applying specified margins, and optionally adjusting to target dimensions.
        Args:
            input_paths (list): List of paths to input images, wildcards allowed
            margin (int, optional): Margin in pixels to add around content. Defaults to 30.
            width_mm (float, optional): Target width in millimeters
            height_mm (float, optional): Target height in millimeters
            output_dir (str, optional): Directory to save output images
            keep_orig (bool, optional): Whether to remove input files after processing. Defaults to True.
            force (bool, optional): Whether to overwrite existing files. Defaults to True.
        """
        # Expand file paths (in case wildcards were used)
        input_files = []
        for path in input_paths:
            if os.path.isfile(path):
                input_files.append(path)
            else:
                # Use glob to expand wildcards
                expanded = glob(path)
                input_files.extend(expanded)

        if not input_files:
            print("No input files found!")
        else:
            for input_path in input_files:
                try:
                    # Determine output path
                    if output_dir:
                        # Ensure output directory exists
                        os.makedirs(output_dir, exist_ok=True)
                        # Create output path in specified directory
                        filename = os.path.basename(input_path)
                        basename, ext = os.path.splitext(filename)
                        # Remove previous '_cropped' suffixes
                        basename = basename.replace("_cropped", "")
                        output_path = os.path.join(
                            output_dir, f"{basename}_cropped{ext}"
                        )
                    else:
                        # Default: save in same directory as input
                        basename, ext = os.path.splitext(input_path)
                        # Remove previous '_cropped' suffixes
                        basename = basename.replace("_cropped", "")
                        output_path = f"{basename}_cropped{ext}"

                    if os.path.exists(output_path) and (not force):
                        print(
                            f"Output path: {output_path} exists. Skipping..."
                        )
                        continue

                    # Crop the image
                    cropped_img = cls._crop_image(input_path, margin=margin)

                    # Add margins if dimensions specified
                    if width_mm or height_mm:
                        cropped_img = cls.add_margins(
                            cropped_img, width_mm=width_mm, height_mm=height_mm
                        )

                    # Get the DPI for reporting
                    dpi = cls._get_image_dpi(cropped_img)

                    # Calculate actual dimensions for reporting
                    width_px = cropped_img.width
                    height_px = cropped_img.height

                    # Save the processed image
                    cls._save_image(cropped_img, output_path)

                    if not keep_orig:
                        remove(input_path)

                except Exception as e:
                    print(f"Error processing {input_path}: {e}")


def _parse_args():
    """
    Parse command line arguments for the crop_images function.
    Returns:
        argparse.Namespace: Parsed argument object
    """
    import argparse

    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Crop images to content area with margin"
    )
    parser.add_argument(
        "input_paths",
        nargs="+",
        help="Path(s) to input image(s), accepts wildcards",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="If true remove output path if exists.",
    )
    parser.add_argument(
        "--margin",
        type=int,
        default=30,
        help="Margin pixel size around the content area.",
    )
    parser.add_argument(
        "--width_mm",
        type=float,
        help="Target width in millimeters (adds margins if needed)",
    )
    parser.add_argument(
        "--height_mm",
        type=float,
        help="Target height in millimeters (adds margins if needed)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI to use for mm to pixel conversion (default: 300)",
    )
    parser.add_argument(
        "--output_dir",
        type=str,
        help="Directory to save cropped images (default: same as input)",
    )
    args = parser._parse_args()
    return args


# For backwards compatibility
def crop_images(*args, **kwargs):
    return ImageCropper.crop_images(*args, **kwargs)


if __name__ == "__main__":
    args = _parse_args()
    ImageCropper.crop_images(
        input_paths=args.input_paths,
        margin=args.margin,
        width_mm=args.width_mm,
        height_mm=args.height_mm,
        output_dir=args.output_dir,
        force=args.force,
    )

# EOF