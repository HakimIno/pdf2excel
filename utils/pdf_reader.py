#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Reader Module
================

Handles PDF file reading and text extraction using multiple libraries
for maximum compatibility and accuracy.
"""

import os
import logging
from typing import List, Dict, Any, Optional, Tuple

try:
    import PyPDF2
    import pdfplumber
except ImportError as e:
    print(f"Required PDF libraries not installed: {e}")
    print("Install with: pip install PyPDF2 pdfplumber")


class PDFReader:
    """
    PDF reading and text extraction class
    
    Features:
    - Text extraction with formatting preservation
    - Metadata extraction
    - Password-protected PDF support
    - Multiple extraction methods for compatibility
    """
    
    def __init__(self, optimize_for_speed: bool = False):
        """Initialize PDF reader"""
        self.logger = logging.getLogger(__name__)
        self.password = None
        self.optimize_for_speed = optimize_for_speed
        
    def set_password(self, password: str):
        """Set password for encrypted PDFs"""
        self.password = password
        
    def extract_text(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """
        Extract text from PDF using multiple methods
        
        Args:
            pdf_path (str): Path to PDF file
            pages_range (tuple): Optional page range (start, end)
            
        Returns:
            list: List of dictionaries containing page text and metadata
        """
        try:
            # Try pdfplumber first (better formatting)
            text_data = self._extract_with_pdfplumber(pdf_path, pages_range)
            
            if not text_data:
                # Fallback to PyPDF2
                self.logger.info("Falling back to PyPDF2")
                text_data = self._extract_with_pypdf2(pdf_path, pages_range)
                
            return text_data
            
        except Exception as e:
            self.logger.error(f"Error extracting text from PDF: {str(e)}")
            return []
            
    def _extract_with_pdfplumber(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Extract text using pdfplumber"""
        text_data = []
        
        try:
            with pdfplumber.open(pdf_path, password=self.password) as pdf:
                total_pages = len(pdf.pages)
                self.logger.info(f"PDF has {total_pages} pages")
                
                # Determine page range
                start_page = pages_range[0] - 1 if pages_range else 0
                end_page = min(pages_range[1], total_pages) if pages_range else total_pages
                
                for page_num in range(start_page, end_page):
                    try:
                        page = pdf.pages[page_num]
                        text = page.extract_text()
                        
                        if text:
                            # Extract additional information (skip for speed optimization)
                            if self.optimize_for_speed:
                                text_data.append({
                                    'page': page_num + 1,
                                    'text': text.strip(),
                                    'char_count': len(text),
                                    'word_count': len(text.split()),
                                    'fonts': [],
                                    'bbox': None
                                })
                            else:
                                chars = page.chars
                                words = page.extract_words()
                                
                                text_data.append({
                                    'page': page_num + 1,
                                    'text': text.strip(),
                                    'char_count': len(text),
                                    'word_count': len(words) if words else 0,
                                    'fonts': self._extract_font_info(chars) if chars else [],
                                    'bbox': page.bbox if hasattr(page, 'bbox') else None
                                })
                            
                        else:
                            # Handle pages with no extractable text
                            text_data.append({
                                'page': page_num + 1,
                                'text': '',
                                'char_count': 0,
                                'word_count': 0,
                                'fonts': [],
                                'bbox': None,
                                'note': 'No extractable text found'
                            })
                            
                    except Exception as e:
                        self.logger.warning(f"Error processing page {page_num + 1}: {str(e)}")
                        continue
                        
        except Exception as e:
            self.logger.error(f"Error with pdfplumber: {str(e)}")
            return []
            
        return text_data
        
    def _extract_with_pypdf2(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Extract text using PyPDF2 as fallback"""
        text_data = []
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Handle encrypted PDFs
                if pdf_reader.is_encrypted:
                    if self.password:
                        pdf_reader.decrypt(self.password)
                    else:
                        self.logger.error("PDF is encrypted but no password provided")
                        return []
                
                total_pages = len(pdf_reader.pages)
                
                # Determine page range
                start_page = pages_range[0] - 1 if pages_range else 0
                end_page = min(pages_range[1], total_pages) if pages_range else total_pages
                
                for page_num in range(start_page, end_page):
                    try:
                        page = pdf_reader.pages[page_num]
                        text = page.extract_text()
                        
                        text_data.append({
                            'page': page_num + 1,
                            'text': text.strip(),
                            'char_count': len(text),
                            'word_count': len(text.split()),
                            'fonts': [],  # PyPDF2 doesn't provide font info easily
                            'bbox': None,
                            'extraction_method': 'PyPDF2'
                        })
                        
                    except Exception as e:
                        self.logger.warning(f"Error processing page {page_num + 1}: {str(e)}")
                        continue
                        
        except Exception as e:
            self.logger.error(f"Error with PyPDF2: {str(e)}")
            return []
            
        return text_data
        
    def _extract_font_info(self, chars: List[Dict]) -> List[Dict[str, Any]]:
        """Extract font information from character data"""
        fonts = {}
        
        for char in chars:
            font_name = char.get('fontname', 'Unknown')
            font_size = char.get('size', 0)
            font_key = f"{font_name}_{font_size}"
            
            if font_key not in fonts:
                fonts[font_key] = {
                    'name': font_name,
                    'size': font_size,
                    'count': 0
                }
            fonts[font_key]['count'] += 1
            
        return list(fonts.values())
        
    def extract_metadata(self, pdf_path: str) -> Dict[str, Any]:
        """
        Extract metadata from PDF
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            dict: PDF metadata
        """
        metadata = {}
        
        try:
            # Try with PyPDF2 first for metadata
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                if pdf_reader.is_encrypted and self.password:
                    pdf_reader.decrypt(self.password)
                    
                if hasattr(pdf_reader, 'metadata') and pdf_reader.metadata:
                    metadata.update({
                        'title': pdf_reader.metadata.get('/Title', ''),
                        'author': pdf_reader.metadata.get('/Author', ''),
                        'subject': pdf_reader.metadata.get('/Subject', ''),
                        'creator': pdf_reader.metadata.get('/Creator', ''),
                        'producer': pdf_reader.metadata.get('/Producer', ''),
                        'creation_date': pdf_reader.metadata.get('/CreationDate', ''),
                        'modification_date': pdf_reader.metadata.get('/ModDate', ''),
                    })
                    
                metadata.update({
                    'page_count': len(pdf_reader.pages),
                    'encrypted': pdf_reader.is_encrypted,
                    'file_size': os.path.getsize(pdf_path),
                    'file_name': os.path.basename(pdf_path)
                })
                
        except Exception as e:
            self.logger.error(f"Error extracting metadata: {str(e)}")
            
        return metadata
        
    def get_page_count(self, pdf_path: str) -> int:
        """Get total number of pages in PDF"""
        try:
            with pdfplumber.open(pdf_path, password=self.password) as pdf:
                return len(pdf.pages)
        except:
            try:
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    if pdf_reader.is_encrypted and self.password:
                        pdf_reader.decrypt(self.password)
                    return len(pdf_reader.pages)
            except Exception as e:
                self.logger.error(f"Error getting page count: {str(e)}")
                return 0
                
    def validate_pdf(self, pdf_path: str) -> bool:
        """Validate if file is a readable PDF"""
        try:
            if not os.path.exists(pdf_path):
                return False
                
            if not pdf_path.lower().endswith('.pdf'):
                return False
                
            # Try to open and read first page
            with pdfplumber.open(pdf_path, password=self.password) as pdf:
                if len(pdf.pages) > 0:
                    return True
                    
        except Exception as e:
            self.logger.warning(f"PDF validation failed: {str(e)}")
            
        return False
