#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Image Extractor Module
=====================

Extracts images from PDF files and saves them as separate files.
Provides detailed information about extracted images.
"""

import os
import logging
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False
    print("PyMuPDF not available. Install with: pip install PyMuPDF")

try:
    from PIL import Image
    import io
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Pillow not available. Install with: pip install Pillow")


class ImageExtractor:
    """
    Image extraction class for PDF files
    
    Features:
    - Extract all images from PDF
    - Save images in various formats
    - Provide detailed image metadata
    - Handle different image types and formats
    """
    
    def __init__(self, output_dir: str = "output/images"):
        """
        Initialize image extractor
        
        Args:
            output_dir (str): Directory to save extracted images
        """
        self.output_dir = output_dir
        self.logger = logging.getLogger(__name__)
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
    def extract_images(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """
        Extract images from PDF
        
        Args:
            pdf_path (str): Path to PDF file
            pages_range (tuple): Optional page range (start, end)
            
        Returns:
            list: List of extracted image information
        """
        try:
            if PYMUPDF_AVAILABLE:
                return self._extract_with_pymupdf(pdf_path, pages_range)
            else:
                self.logger.warning("No image extraction library available")
                return []
                
        except Exception as e:
            self.logger.error(f"Error extracting images: {str(e)}")
            return []
            
    def _extract_with_pymupdf(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Extract images using PyMuPDF"""
        images_data = []
        
        if not PYMUPDF_AVAILABLE:
            return []
            
        try:
            # Open PDF document
            pdf_document = fitz.open(pdf_path)
            total_pages = len(pdf_document)
            
            self.logger.info(f"Extracting images from PDF with {total_pages} pages")
            
            # Determine page range
            start_page = pages_range[0] - 1 if pages_range else 0
            end_page = min(pages_range[1], total_pages) if pages_range else total_pages
            
            pdf_name = Path(pdf_path).stem
            
            for page_num in range(start_page, end_page):
                try:
                    page = pdf_document[page_num]
                    image_list = page.get_images()
                    
                    self.logger.info(f"Found {len(image_list)} images on page {page_num + 1}")
                    
                    for img_index, img in enumerate(image_list):
                        try:
                            # Get image data
                            xref = img[0]
                            pix = fitz.Pixmap(pdf_document, xref)
                            
                            # Skip if image is too small (likely decorative)
                            if pix.width < 50 or pix.height < 50:
                                pix = None
                                continue
                                
                            # Generate filename
                            image_filename = f"{pdf_name}_page_{page_num + 1}_img_{img_index + 1}"
                            
                            # Determine format and save
                            if pix.n - pix.alpha < 4:  # GRAY or RGB
                                image_format = "png"
                                image_path = os.path.join(self.output_dir, f"{image_filename}.png")
                                pix.save(image_path)
                            else:  # CMYK: convert to RGB first
                                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                                image_format = "png"
                                image_path = os.path.join(self.output_dir, f"{image_filename}.png")
                                pix1.save(image_path)
                                pix1 = None
                                
                            # Get additional image information
                            image_info = {
                                'image_id': len(images_data) + 1,
                                'filename': os.path.basename(image_path),
                                'full_path': image_path,
                                'page': page_num + 1,
                                'width': pix.width,
                                'height': pix.height,
                                'format': image_format,
                                'colorspace': pix.colorspace.name if pix.colorspace else 'Unknown',
                                'has_alpha': bool(pix.alpha),
                                'file_size': os.path.getsize(image_path) if os.path.exists(image_path) else 0,
                                'dpi': self._estimate_dpi(pix, page),
                                'extraction_method': 'PyMuPDF'
                            }
                            
                            # Add PIL-based analysis if available
                            if PIL_AVAILABLE and os.path.exists(image_path):
                                pil_info = self._analyze_with_pil(image_path)
                                image_info.update(pil_info)
                                
                            images_data.append(image_info)
                            self.logger.info(f"Extracted image: {image_filename}")
                            
                            pix = None  # Free memory
                            
                        except Exception as e:
                            self.logger.warning(f"Error extracting image {img_index} from page {page_num + 1}: {str(e)}")
                            continue
                            
                except Exception as e:
                    self.logger.warning(f"Error processing page {page_num + 1}: {str(e)}")
                    continue
                    
            pdf_document.close()
            
            self.logger.info(f"Successfully extracted {len(images_data)} images")
            
        except Exception as e:
            self.logger.error(f"Error with PyMuPDF extraction: {str(e)}")
            
        return images_data
        
    def _analyze_with_pil(self, image_path: str) -> Dict[str, Any]:
        """Analyze image using PIL for additional information"""
        pil_info = {}
        
        if not PIL_AVAILABLE:
            return pil_info
            
        try:
            with Image.open(image_path) as img:
                pil_info.update({
                    'mode': img.mode,
                    'has_transparency': img.mode in ('RGBA', 'LA') or 'transparency' in img.info,
                    'is_animated': getattr(img, 'is_animated', False),
                    'n_frames': getattr(img, 'n_frames', 1),
                })
                
                # Get EXIF data if available
                if hasattr(img, '_getexif') and img._getexif():
                    pil_info['has_exif'] = True
                else:
                    pil_info['has_exif'] = False
                    
                # Estimate image quality/complexity
                pil_info['estimated_quality'] = self._estimate_image_quality(img)
                
        except Exception as e:
            self.logger.warning(f"Error analyzing image with PIL: {str(e)}")
            
        return pil_info
        
    def _estimate_dpi(self, pix, page) -> Tuple[float, float]:
        """Estimate DPI of extracted image"""
        try:
            # Get page dimensions in points (1/72 inch)
            page_rect = page.rect
            page_width_inches = page_rect.width / 72
            page_height_inches = page_rect.height / 72
            
            # Calculate DPI
            dpi_x = pix.width / page_width_inches if page_width_inches > 0 else 0
            dpi_y = pix.height / page_height_inches if page_height_inches > 0 else 0
            
            return (round(dpi_x, 2), round(dpi_y, 2))
            
        except Exception:
            return (0, 0)
            
    def _estimate_image_quality(self, img: Image.Image) -> str:
        """Estimate image quality based on various factors"""
        try:
            width, height = img.size
            total_pixels = width * height
            
            # Basic quality estimation based on size and mode
            if total_pixels < 10000:  # Less than 100x100
                quality = "low"
            elif total_pixels < 100000:  # Less than ~316x316
                quality = "medium"
            else:
                quality = "high"
                
            # Adjust based on color mode
            if img.mode in ('1', 'L'):  # Monochrome or grayscale
                if quality == "high":
                    quality = "medium"
            elif img.mode in ('RGB', 'RGBA'):
                pass  # Keep original assessment
            else:
                quality = "unknown"
                
            return quality
            
        except Exception:
            return "unknown"
            
    def create_image_catalog(self, images_data: List[Dict[str, Any]], catalog_path: str = None) -> str:
        """
        Create a catalog/index of extracted images
        
        Args:
            images_data (list): List of image information
            catalog_path (str): Path to save catalog file
            
        Returns:
            str: Path to created catalog file
        """
        if not catalog_path:
            catalog_path = os.path.join(self.output_dir, "image_catalog.txt")
            
        try:
            with open(catalog_path, 'w', encoding='utf-8') as f:
                f.write("PDF Image Extraction Catalog\n")
                f.write("=" * 50 + "\n\n")
                
                if not images_data:
                    f.write("No images were extracted.\n")
                    return catalog_path
                    
                # Summary
                f.write(f"Total Images Extracted: {len(images_data)}\n")
                pages_with_images = len(set(img['page'] for img in images_data))
                f.write(f"Pages with Images: {pages_with_images}\n")
                total_size = sum(img.get('file_size', 0) for img in images_data)
                f.write(f"Total Size: {self._format_file_size(total_size)}\n\n")
                
                # Detailed listing
                f.write("Detailed Image List:\n")
                f.write("-" * 50 + "\n\n")
                
                for img in images_data:
                    f.write(f"Image ID: {img.get('image_id', 'N/A')}\n")
                    f.write(f"Filename: {img.get('filename', 'N/A')}\n")
                    f.write(f"Page: {img.get('page', 'N/A')}\n")
                    f.write(f"Dimensions: {img.get('width', 0)} x {img.get('height', 0)} pixels\n")
                    f.write(f"Format: {img.get('format', 'Unknown')}\n")
                    f.write(f"File Size: {self._format_file_size(img.get('file_size', 0))}\n")
                    f.write(f"Color Mode: {img.get('mode', 'Unknown')}\n")
                    f.write(f"DPI: {img.get('dpi', (0, 0))}\n")
                    f.write(f"Quality Estimate: {img.get('estimated_quality', 'Unknown')}\n")
                    f.write("\n")
                    
            self.logger.info(f"Image catalog created: {catalog_path}")
            
        except Exception as e:
            self.logger.error(f"Error creating image catalog: {str(e)}")
            
        return catalog_path
        
    def _format_file_size(self, size_bytes: int) -> str:
        """Format file size in human-readable format"""
        if size_bytes == 0:
            return "0 B"
            
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
            
        return f"{size_bytes:.1f} {size_names[i]}"
        
    def get_extraction_summary(self, images_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Get summary statistics for extracted images"""
        if not images_data:
            return {"total_images": 0}
            
        summary = {
            "total_images": len(images_data),
            "pages_with_images": len(set(img['page'] for img in images_data)),
            "formats": list(set(img.get('format', 'Unknown') for img in images_data)),
            "total_file_size": sum(img.get('file_size', 0) for img in images_data),
            "average_width": sum(img.get('width', 0) for img in images_data) / len(images_data),
            "average_height": sum(img.get('height', 0) for img in images_data) / len(images_data),
            "largest_image": max(images_data, key=lambda x: x.get('width', 0) * x.get('height', 0)),
            "smallest_image": min(images_data, key=lambda x: x.get('width', 0) * x.get('height', 0))
        }
        
        return summary
