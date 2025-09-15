#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Excel Converter - Main Module
====================================

A comprehensive tool to convert PDF files to Excel format with detailed extraction
of text, tables, images, and metadata.
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

try:
    from tqdm import tqdm
except ImportError:
    print("tqdm not installed. Install with: pip install tqdm")
    
# Import utility modules
try:
    from utils.pdf_reader import PDFReader
    from utils.table_extractor import TableExtractor  
    from utils.image_extractor import ImageExtractor
    from utils.excel_writer import ExcelWriter
    from utils.pdf_like_writer import PDFLikeWriter
except ImportError as e:
    print(f"Warning: {e}")
    print("Some utility modules not found. They will be created.")


class PDFToExcelConverter:
    """
    Main class for converting PDF files to Excel format
    
    Features:
    - Text extraction with formatting
    - Table detection and extraction  
    - Image extraction and cataloging
    - Metadata extraction
    - Batch processing
    """
    
    def __init__(self, 
                 extract_images: bool = True,
                 extract_tables: bool = True,
                 extract_metadata: bool = True,
                 pages_range: Optional[Tuple[int, int]] = None,
                 table_detection_method: str = "tabula",
                 preserve_layout: bool = True,
                 optimize_for_speed: bool = True):
        """Initialize the PDF to Excel converter"""
        self.extract_images = extract_images
        self.extract_tables = extract_tables
        self.extract_metadata = extract_metadata
        self.pages_range = pages_range
        self.table_detection_method = table_detection_method
        self.preserve_layout = preserve_layout
        self.optimize_for_speed = optimize_for_speed
        self.password = None
        
        # Setup logging
        self._setup_logging()
        
        # Initialize utility classes
        try:
            self.pdf_reader = PDFReader()
            self.table_extractor = TableExtractor(method=table_detection_method)
            self.image_extractor = ImageExtractor() if extract_images else None
            
            # Choose writer based on layout preference
            if preserve_layout:
                self.excel_writer = PDFLikeWriter()  # Use new PDF-like writer
            else:
                self.excel_writer = ExcelWriter()
                
        except NameError:
            self.logger.warning("Utility classes not yet implemented")
        
    def _setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
    def convert_single_file(self, input_path: str, output_path: str) -> bool:
        """
        Convert a single PDF file to Excel
        
        Args:
            input_path (str): Path to input PDF file
            output_path (str): Path to output Excel file
            
        Returns:
            bool: True if conversion successful
        """
        try:
            self.logger.info(f"Starting conversion: {input_path} -> {output_path}")
            
            # Validate input
            if not os.path.exists(input_path):
                self.logger.error(f"Input file not found: {input_path}")
                return False
                
            if not input_path.lower().endswith('.pdf'):
                self.logger.error(f"Input file is not a PDF: {input_path}")
                return False
                
            # Create output directory if output_path contains directory
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # Extract PDF data
            pdf_data = self._extract_pdf_data(input_path)
            
            # Add PDF path for visual analysis
            if pdf_data:
                pdf_data['pdf_path'] = input_path
            
                # Convert text_data to pages format for PDFLikeWriter
                text_data = pdf_data.get('text_data', [])
                pages = []
                for page_data in text_data:
                    pages.append({
                        'text_content': page_data.get('content', ''),
                        'tables': pdf_data.get('tables', [])
                    })
                pdf_data['pages'] = pages
            
            if not pdf_data:
                self.logger.error("Failed to extract PDF data")
                return False
                
            # Write to Excel
            success = self._write_to_excel(pdf_data, output_path)
            
            if success:
                self.logger.info(f"Conversion completed: {output_path}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Error converting file: {str(e)}")
            return False
            
    def convert_multiple_files(self, input_files: List[str], output_dir: str) -> Dict[str, bool]:
        """Convert multiple PDF files to Excel"""
        results = {}
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            for input_file in tqdm(input_files, desc="Converting PDFs"):
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.xlsx")
                success = self.convert_single_file(input_file, output_file)
                results[input_file] = success
        except NameError:
            # tqdm not available, use simple loop
            for i, input_file in enumerate(input_files):
                print(f"Processing {i+1}/{len(input_files)}: {input_file}")
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.xlsx")
                success = self.convert_single_file(input_file, output_file)
                results[input_file] = success
                
        return results
        
    def _extract_pdf_data(self, pdf_path: str) -> Optional[Dict[str, Any]]:
        """Extract all data from PDF file"""
        try:
            data = {
                'filename': os.path.basename(pdf_path),
                'text_data': [],
                'tables': [],
                'images': [],
                'metadata': {}
            }
            
            self.logger.info(f"Extracting data from: {pdf_path}")
            
            # Extract text data
            try:
                text_data = self.pdf_reader.extract_text(pdf_path, self.pages_range)
                if text_data:
                    data['text_data'] = text_data
                    self.logger.info(f"Extracted text from {len(text_data)} pages")
            except Exception as e:
                self.logger.warning(f"Text extraction failed: {str(e)}")
                
            # Extract tables if enabled
            if self.extract_tables:
                try:
                    tables = self.table_extractor.extract_tables(pdf_path, self.pages_range)
                    if tables:
                        data['tables'] = tables
                        self.logger.info(f"Extracted {len(tables)} tables")
                except Exception as e:
                    self.logger.warning(f"Table extraction failed: {str(e)}")
                    
            # Extract images if enabled
            if self.extract_images and self.image_extractor:
                try:
                    images = self.image_extractor.extract_images(pdf_path, self.pages_range)
                    if images:
                        data['images'] = images
                        self.logger.info(f"Extracted {len(images)} images")
                except Exception as e:
                    self.logger.warning(f"Image extraction failed: {str(e)}")
                    
            # Extract metadata if enabled
            if self.extract_metadata:
                try:
                    metadata = self.pdf_reader.extract_metadata(pdf_path)
                    if metadata:
                        data['metadata'] = metadata
                        self.logger.info("Extracted metadata")
                except Exception as e:
                    self.logger.warning(f"Metadata extraction failed: {str(e)}")
            
            return data
            
        except Exception as e:
            self.logger.error(f"Error extracting PDF data: {str(e)}")
            return None
            
    def _write_to_excel(self, data: Dict[str, Any], output_path: str) -> bool:
        """Write extracted data to Excel file"""
        try:
            # Use the appropriate Excel writer
            if hasattr(self.excel_writer, 'write_to_excel'):
                success = self.excel_writer.write_to_excel(data, output_path)
            else:
                success = False
            
            if success:
                self.logger.info(f"Excel file created successfully: {output_path}")
            else:
                self.logger.error("Failed to create Excel file")
                
            return success
            
        except Exception as e:
            self.logger.error(f"Error writing to Excel: {str(e)}")
            return False


def main():
    """Command line interface"""
    parser = argparse.ArgumentParser(description='Convert PDF files to Excel format')
    
    parser.add_argument('--input', '-i', type=str, help='Input PDF file path')
    parser.add_argument('--output', '-o', type=str, help='Output Excel file path')
    parser.add_argument('--batch', action='store_true', help='Batch processing mode')
    parser.add_argument('--input-dir', type=str, help='Input directory for batch processing')
    parser.add_argument('--output-dir', type=str, help='Output directory for batch processing')
    
    # Layout and formatting options
    parser.add_argument('--preserve-layout', action='store_true', default=True, 
                       help='Preserve PDF layout in Excel (default: True)')
    parser.add_argument('--traditional-format', action='store_true', 
                       help='Use traditional separate sheets format')
    parser.add_argument('--pages', type=str, help='Page range (e.g., 1-5)')
    
    # Performance options
    parser.add_argument('--fast', action='store_true', 
                       help='Optimize for speed (may reduce accuracy)')
    parser.add_argument('--no-tables', action='store_true', 
                       help='Skip table extraction for faster processing')
    parser.add_argument('--no-images', action='store_true', 
                       help='Skip image extraction for faster processing')
    
    args = parser.parse_args()
    
    # Parse page range
    pages_range = None
    if args.pages:
        try:
            start, end = map(int, args.pages.split('-'))
            pages_range = (start, end)
        except ValueError:
            print(f"Invalid page range format: {args.pages}")
            sys.exit(1)
    
    # Determine layout preservation
    preserve_layout = not args.traditional_format and args.preserve_layout
    
    # Create converter with options
    converter = PDFToExcelConverter(
        extract_images=not args.no_images,
        extract_tables=not args.no_tables,
        extract_metadata=True,
        pages_range=pages_range,
        preserve_layout=preserve_layout,
        optimize_for_speed=args.fast
    )
    
    if args.batch:
        if not args.input_dir or not args.output_dir:
            print("Batch mode requires --input-dir and --output-dir")
            sys.exit(1)
            
        pdf_files = list(Path(args.input_dir).glob('*.pdf'))
        if not pdf_files:
            print(f"No PDF files found in {args.input_dir}")
            sys.exit(1)
            
        print(f"Found {len(pdf_files)} PDF files")
        results = converter.convert_multiple_files([str(f) for f in pdf_files], args.output_dir)
        
        successful = sum(1 for success in results.values() if success)
        print(f"\nConversion completed: {successful}/{len(results)} files successful")
        
    else:
        if not args.input or not args.output:
            print("Single file mode requires --input and --output")
            print("Example: python main.py --input document.pdf --output result.xlsx")
            sys.exit(1)
            
        success = converter.convert_single_file(args.input, args.output)
        if success:
            print(f"Conversion successful!")
        else:
            print("Conversion failed")
            sys.exit(1)


if __name__ == "__main__":
    main()
