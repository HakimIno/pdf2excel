#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF-Like Excel Writer
====================

Creates Excel files that look exactly like the original PDF
using the Intelligent PDF Reader system.
"""

import os
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Fill, Border, Side, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    EXCEL_LIBS_AVAILABLE = True
except ImportError as e:
    EXCEL_LIBS_AVAILABLE = False
    print(f"Excel libraries not available: {e}")

try:
    from .table_filter import TableFilter
    TABLE_FILTER_AVAILABLE = True
except ImportError:
    TABLE_FILTER_AVAILABLE = False

try:
    from .intelligent_pdf_reader import IntelligentPDFReader
    INTELLIGENT_READER_AVAILABLE = True
except ImportError:
    INTELLIGENT_READER_AVAILABLE = False


class PDFLikeWriter:
    """
    Excel writer that creates output identical to PDF layout
    using intelligent PDF analysis without hard-coding
    """
    
    def __init__(self):
        """Initialize PDF-like writer"""
        self.logger = logging.getLogger(__name__)
        
        if not EXCEL_LIBS_AVAILABLE:
            self.logger.error("Excel libraries not available")
            
        # Initialize table filter
        if TABLE_FILTER_AVAILABLE:
            self.table_filter = TableFilter()
        else:
            self.table_filter = None
            
        # Initialize intelligent PDF reader (main system)
        if INTELLIGENT_READER_AVAILABLE:
            self.intelligent_reader = IntelligentPDFReader()
        else:
            self.intelligent_reader = None
            
    def write_to_excel(self, pdf_data: Dict[str, Any], output_path: str) -> bool:
        """
        Write PDF data to Excel with exact PDF layout
        
        Args:
            pdf_data (dict): Extracted PDF data
            output_path (str): Output Excel file path
            
        Returns:
            bool: True if successful
        """
        if not EXCEL_LIBS_AVAILABLE:
            return self._write_text_format(pdf_data, output_path)
            
        try:
            wb = Workbook()
            
            # Remove default sheet
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            
            # Get pages data
            pages_data = pdf_data.get('pages', [])
            if not pages_data:
                self.logger.warning("No page data found")
                return False
                
            # Create one sheet per page with intelligent layout
            for page_num, page_content in enumerate(pages_data, 1):
                self._create_intelligent_pdf_sheet(wb, page_num, page_content, pdf_data)
                
            # Save workbook
            wb.save(output_path)
            self.logger.info(f"PDF-like Excel file created: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating PDF-like Excel: {str(e)}")
            return False
            
    def _create_intelligent_pdf_sheet(self, wb: Workbook, page_num: int, page_content: Dict[str, Any], 
                                    pdf_data: Dict[str, Any]) -> None:
        """Create intelligent PDF sheet using PyMuPDF analysis"""
        ws = wb.create_sheet(f"หน้า_{page_num}")
        
        # Use intelligent PDF reader (primary method)
        if self.intelligent_reader:
            pdf_path = pdf_data.get('pdf_path')
            if pdf_path:
                layout_info = self.intelligent_reader.analyze_pdf_layout(pdf_path)
                if layout_info:
                    self.intelligent_reader.create_excel_from_layout(ws, layout_info)
                    return
                    
        # Fallback to basic method if intelligent reader fails
        self._create_basic_sheet(ws, page_content, pdf_data, page_num)
        
    def _create_basic_sheet(self, ws, page_content: Dict[str, Any], pdf_data: Dict[str, Any], page_num: int) -> None:
        """Basic fallback method"""
        text_content = page_content.get('text_content', '')
        tables = page_content.get('tables', [])
        
        # Filter real tables
        if self.table_filter:
            tables = self.table_filter.filter_real_tables(tables)
        
        # Create simple layout
        current_row = 1
        
        # Add text content
        if text_content:
            lines = text_content.split('\n')
            for line in lines:
                line = line.strip()
                if line:
                    ws[f'A{current_row}'] = line
                    ws[f'A{current_row}'].font = Font(size=10)
                    current_row += 1
                
        # Add tables
        for table in tables:
            table_data = table.get('data', [])
            if table_data:
                current_row += 1  # Space before table
                
                # Add table
                for row_idx, row_data in enumerate(table_data):
                    for col_idx, cell_value in enumerate(row_data[:8], 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=str(cell_value) if cell_value else "")
                        
                        if row_idx == 0:  # Header
                            cell.font = Font(size=10, bold=True, color='FFFFFFFF')
                            cell.fill = PatternFill(start_color='FF4472C4', end_color='FF4472C4', fill_type='solid')
                        else:
                            cell.font = Font(size=10)
                            
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = Border(
                            left=Side(style='thin', color='FFE0E0E0'),
                            right=Side(style='thin', color='FFE0E0E0'),
                            top=Side(style='thin', color='FFE0E0E0'),
                            bottom=Side(style='thin', color='FFE0E0E0')
                        )
                current_row += 1
                
        # Set column widths
        for i in range(1, 9):
            ws.column_dimensions[get_column_letter(i)].width = 15
            
    def _write_text_format(self, pdf_data: Dict[str, Any], output_path: str) -> bool:
        """Fallback to text format if Excel not available"""
        try:
            with open(output_path.replace('.xlsx', '.txt'), 'w', encoding='utf-8') as f:
                f.write("PDF to Excel Conversion Results\n")
                f.write("=" * 50 + "\n\n")
                
                pages_data = pdf_data.get('pages', [])
                for page_num, page_content in enumerate(pages_data, 1):
                    f.write(f"Page {page_num}\n")
                    f.write("-" * 20 + "\n")
                    
                    text_content = page_content.get('text_content', '')
                    if text_content:
                        f.write(text_content)
                    f.write("\n\n")
                    
                    tables = page_content.get('tables', [])
                    for table_idx, table in enumerate(tables, 1):
                        f.write(f"Table {table_idx}:\n")
                        table_data = table.get('data', [])
                        for row in table_data:
                            f.write(" | ".join(str(cell) for cell in row))
                            f.write("\n")
                        f.write("\n")
                        
            self.logger.info(f"Text format file created: {output_path.replace('.xlsx', '.txt')}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating text format: {str(e)}")
            return False