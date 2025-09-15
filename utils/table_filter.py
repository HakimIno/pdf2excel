#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Table Filter Module
==================

Filters out fake tables and keeps only meaningful table data.
"""

import logging
from typing import List, Dict, Any
import re


class TableFilter:
    """
    Filter for removing fake/meaningless tables
    """
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def filter_real_tables(self, tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Filter out fake tables and return only real ones"""
        real_tables = []
        
        for table in tables:
            if self._is_real_table(table):
                real_tables.append(table)
            else:
                self.logger.debug(f"Filtered out fake table")
                
        return real_tables
        
    def _is_real_table(self, table: Dict[str, Any]) -> bool:
        """Check if table is real and meaningful"""
        table_data = table.get('data', [])
        
        if not table_data or len(table_data) < 2:
            return False
            
        # Check if it has reasonable number of columns
        if len(table_data[0]) < 2:
            return False
            
        # Check data quality
        non_empty_cells = 0
        total_cells = 0
        numeric_cells = 0
        
        for row in table_data:
            for cell in row:
                total_cells += 1
                if cell and str(cell).strip():
                    non_empty_cells += 1
                    if self._is_numeric(str(cell)):
                        numeric_cells += 1
                        
        # Must have reasonable fill rate
        fill_rate = non_empty_cells / total_cells if total_cells > 0 else 0
        if fill_rate < 0.4:
            return False
            
        # Must have some structured data
        if numeric_cells == 0 and total_cells > 4:
            return False
            
        return True
        
    def _is_numeric(self, text: str) -> bool:
        """Check if text represents a numeric value"""
        text = text.strip().replace(',', '').replace('$', '').replace('%', '').replace('-', '')
        if not text:
            return False
            
        try:
            float(text)
            return True
        except ValueError:
            return False