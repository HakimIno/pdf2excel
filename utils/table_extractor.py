#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Table Extractor Module
=====================

Extracts tables from PDF files using multiple methods for maximum accuracy.
Supports both tabula-py and camelot-py for different table detection scenarios.
"""

import os
import logging
from typing import List, Dict, Any, Optional, Tuple

# For now, we'll implement basic table detection without pandas/tabula
PANDAS_AVAILABLE = False
TABULA_AVAILABLE = False
CAMELOT_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    print("pandas not available - using basic table extraction")

try:
    import tabula
    TABULA_AVAILABLE = True
except ImportError:
    print("tabula-py not available - using basic table extraction")

try:
    import camelot
    CAMELOT_AVAILABLE = True
except ImportError:
    print("camelot-py not available")


class TableExtractor:
    """
    Table extraction class supporting multiple methods
    
    Methods:
    - tabula-py: Good for simple tables with clear borders
    - camelot-py: Better for complex tables and mixed content
    """
    
    def __init__(self, method: str = "tabula"):
        """
        Initialize table extractor
        
        Args:
            method (str): Extraction method ('tabula' or 'camelot')
        """
        self.method = method.lower()
        self.logger = logging.getLogger(__name__)
        
        # Validate method availability
        if self.method == "camelot" and not CAMELOT_AVAILABLE:
            self.logger.warning("Camelot not available, falling back to tabula")
            self.method = "tabula"
            
    def extract_tables(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """
        Extract tables from PDF
        
        Args:
            pdf_path (str): Path to PDF file
            pages_range (tuple): Optional page range (start, end)
            
        Returns:
            list: List of extracted tables with metadata
        """
        try:
            if self.method == "camelot" and CAMELOT_AVAILABLE:
                return self._extract_with_camelot(pdf_path, pages_range)
            elif self.method == "tabula" and TABULA_AVAILABLE:
                return self._extract_with_tabula(pdf_path, pages_range)
            else:
                return self._extract_basic_tables(pdf_path, pages_range)
                
        except Exception as e:
            self.logger.error(f"Error extracting tables: {str(e)}")
            return []
            
    def _extract_with_tabula(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Extract tables using tabula-py"""
        tables_data = []
        
        try:
            # Determine pages to process
            pages = "all"
            if pages_range:
                if pages_range[0] == pages_range[1]:
                    pages = str(pages_range[0])
                else:
                    pages = f"{pages_range[0]}-{pages_range[1]}"
                    
            self.logger.info(f"Extracting tables from pages: {pages}")
            
            # Extract tables with multiple methods for better coverage
            extraction_methods = [
                {"method": "lattice", "multiple_tables": True},
                {"method": "stream", "multiple_tables": True},
                {"method": "lattice", "multiple_tables": False},
                {"method": "stream", "multiple_tables": False}
            ]
            
            all_tables = []
            
            for method_config in extraction_methods:
                try:
                    tables = tabula.read_pdf(
                        pdf_path,
                        pages=pages,
                        multiple_tables=method_config["multiple_tables"],
                        pandas_options={'header': None},
                        **{"method": method_config["method"]}
                    )
                    
                    if tables and len(tables) > 0:
                        for i, table in enumerate(tables):
                            if not table.empty:
                                table_info = {
                                    'table_id': len(all_tables) + 1,
                                    'extraction_method': f"tabula_{method_config['method']}",
                                    'page': self._estimate_page_number(pdf_path, pages),
                                    'rows': len(table),
                                    'columns': len(table.columns),
                                    'data': table,
                                    'csv_data': table.to_csv(index=False),
                                    'has_header': self._detect_header(table),
                                    'confidence': self._calculate_confidence(table, method_config["method"])
                                }
                                all_tables.append(table_info)
                                
                except Exception as e:
                    self.logger.warning(f"Method {method_config['method']} failed: {str(e)}")
                    continue
                    
            # Remove duplicates and select best tables
            tables_data = self._deduplicate_tables(all_tables)
            
            self.logger.info(f"Extracted {len(tables_data)} unique tables using tabula")
            
        except Exception as e:
            self.logger.error(f"Error with tabula extraction: {str(e)}")
            
        return tables_data
        
    def _extract_with_camelot(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Extract tables using camelot-py"""
        tables_data = []
        
        if not CAMELOT_AVAILABLE:
            return []
            
        try:
            # Determine pages to process
            pages = "all"
            if pages_range:
                if pages_range[0] == pages_range[1]:
                    pages = str(pages_range[0])
                else:
                    pages = f"{pages_range[0]}-{pages_range[1]}"
                    
            self.logger.info(f"Extracting tables with camelot from pages: {pages}")
            
            # Try both lattice and stream methods
            methods = ["lattice", "stream"]
            
            for method in methods:
                try:
                    tables = camelot.read_pdf(pdf_path, pages=pages, flavor=method)
                    
                    for i, table in enumerate(tables):
                        if len(table.df) > 0:
                            table_info = {
                                'table_id': len(tables_data) + 1,
                                'extraction_method': f"camelot_{method}",
                                'page': table.page,
                                'rows': len(table.df),
                                'columns': len(table.df.columns),
                                'data': table.df,
                                'csv_data': table.df.to_csv(index=False),
                                'has_header': self._detect_header(table.df),
                                'confidence': table.accuracy if hasattr(table, 'accuracy') else 0.0,
                                'parsing_report': table.parsing_report if hasattr(table, 'parsing_report') else {}
                            }
                            tables_data.append(table_info)
                            
                except Exception as e:
                    self.logger.warning(f"Camelot method {method} failed: {str(e)}")
                    continue
                    
            self.logger.info(f"Extracted {len(tables_data)} tables using camelot")
            
        except Exception as e:
            self.logger.error(f"Error with camelot extraction: {str(e)}")
            
        return tables_data
        
    def _extract_basic_tables(self, pdf_path: str, pages_range: Optional[Tuple[int, int]] = None) -> List[Dict[str, Any]]:
        """Basic table extraction using pdfplumber"""
        tables_data = []
        
        try:
            import pdfplumber
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                
                # Determine page range
                start_page = pages_range[0] - 1 if pages_range else 0
                end_page = min(pages_range[1], total_pages) if pages_range else total_pages
                
                for page_num in range(start_page, end_page):
                    try:
                        page = pdf.pages[page_num]
                        
                        # Try to find tables using pdfplumber with stricter settings
                        tables = page.find_tables(
                            table_settings={
                                "vertical_strategy": "lines",  # Only detect tables with actual lines
                                "horizontal_strategy": "lines",
                                "min_words_vertical": 3,  # Require at least 3 words vertically
                                "min_words_horizontal": 2,  # Require at least 2 words horizontally
                                "intersection_tolerance": 3,
                                "snap_tolerance": 3,
                                "join_tolerance": 3
                            }
                        )
                        
                        for i, table in enumerate(tables):
                            try:
                                # Extract table data
                                table_data = table.extract()
                                
                                if table_data and len(table_data) > 0:
                                    # Filter out fake tables
                                    if not self._is_valid_table(table_data, page):
                                        continue
                                        
                                    # Convert to simple format
                                    csv_data = ""
                                    for row in table_data:
                                        csv_data += ",".join([str(cell) if cell else "" for cell in row]) + "\n"
                                    
                                    confidence = self._calculate_table_confidence(table_data, table)
                                    
                                    # Only include tables with reasonable confidence
                                    if confidence >= 0.6:
                                        table_info = {
                                            'table_id': len(tables_data) + 1,
                                            'extraction_method': 'pdfplumber_basic',
                                            'page': page_num + 1,
                                            'rows': len(table_data),
                                            'columns': len(table_data[0]) if table_data else 0,
                                            'data': table_data,
                                            'csv_data': csv_data.strip(),
                                            'has_header': self._detect_header_basic(table_data),
                                            'confidence': confidence
                                        }
                                        tables_data.append(table_info)
                                    
                            except Exception as e:
                                self.logger.warning(f"Error extracting table {i} from page {page_num + 1}: {str(e)}")
                                continue
                                
                    except Exception as e:
                        self.logger.warning(f"Error processing page {page_num + 1}: {str(e)}")
                        continue
                        
            self.logger.info(f"Extracted {len(tables_data)} tables using basic method")
            
        except Exception as e:
            self.logger.error(f"Error with basic table extraction: {str(e)}")
            
        return tables_data
        
    def _detect_header_basic(self, table_data: List[List]) -> bool:
        """Basic header detection for simple table data"""
        if not table_data or len(table_data) < 2:
            return False
            
        try:
            first_row = table_data[0]
            second_row = table_data[1]
            
            # Count numeric values in each row
            first_row_numeric = sum(1 for cell in first_row if cell and str(cell).replace('.', '').replace('-', '').isdigit())
            second_row_numeric = sum(1 for cell in second_row if cell and str(cell).replace('.', '').replace('-', '').isdigit())
            
            # If first row has fewer numeric values, it might be a header
            if len(first_row) > 0 and first_row_numeric < second_row_numeric:
                return True
                
        except Exception:
            pass
            
        return False
        
    def _is_valid_table(self, table_data: List[List], page) -> bool:
        """Check if extracted table is a real table or just formatted text"""
        if not table_data or len(table_data) < 2:
            return False
            
        try:
            # Check 1: Table should have at least 2 rows and 2 columns
            if len(table_data) < 2 or len(table_data[0]) < 2:
                return False
                
            # Check 2: Count non-empty cells
            non_empty_cells = 0
            total_cells = 0
            
            for row in table_data:
                for cell in row:
                    total_cells += 1
                    if cell and str(cell).strip():
                        non_empty_cells += 1
                        
            # Table should have at least 30% non-empty cells
            if total_cells == 0 or (non_empty_cells / total_cells) < 0.3:
                return False
                
            # Check 3: Look for table-like patterns (numbers, consistent formatting)
            numeric_cells = 0
            for row in table_data[1:]:  # Skip header
                for cell in row:
                    if cell and self._is_numeric_like(str(cell)):
                        numeric_cells += 1
                        
            # Check 4: Avoid single-column "tables" (usually just formatted text)
            if len(table_data[0]) == 1:
                return False
                
            # Check 5: Avoid tables where all cells are very short (likely labels)
            avg_cell_length = 0
            cell_count = 0
            for row in table_data:
                for cell in row:
                    if cell:
                        avg_cell_length += len(str(cell))
                        cell_count += 1
                        
            if cell_count > 0:
                avg_cell_length = avg_cell_length / cell_count
                # If average cell length is too short, it might be just formatted text
                if avg_cell_length < 3:
                    return False
                    
            return True
            
        except Exception as e:
            self.logger.warning(f"Error validating table: {str(e)}")
            return False
            
    def _is_numeric_like(self, text: str) -> bool:
        """Check if text looks like a number or currency"""
        text = text.strip().replace(',', '').replace('$', '').replace('%', '')
        try:
            float(text)
            return True
        except ValueError:
            # Check for patterns like "1.23", "123.45", etc.
            import re
            return bool(re.match(r'^-?\d+\.?\d*$', text))
            
    def _calculate_table_confidence(self, table_data: List[List], table_obj) -> float:
        """Calculate confidence score for a table"""
        confidence = 0.5  # Base confidence
        
        try:
            # Factor 1: Size (bigger tables are more likely to be real)
            rows = len(table_data)
            cols = len(table_data[0]) if table_data else 0
            
            if rows >= 3 and cols >= 2:
                confidence += 0.2
            if rows >= 5:
                confidence += 0.1
            if cols >= 3:
                confidence += 0.1
                
            # Factor 2: Data variety (tables with mixed data types are more likely real)
            numeric_cells = 0
            text_cells = 0
            empty_cells = 0
            
            for row in table_data:
                for cell in row:
                    if not cell or not str(cell).strip():
                        empty_cells += 1
                    elif self._is_numeric_like(str(cell)):
                        numeric_cells += 1
                    else:
                        text_cells += 1
                        
            total_cells = rows * cols
            if total_cells > 0:
                # Good mix of data types
                if numeric_cells > 0 and text_cells > 0:
                    confidence += 0.2
                    
                # Not too many empty cells
                empty_ratio = empty_cells / total_cells
                if empty_ratio < 0.3:
                    confidence += 0.1
                elif empty_ratio > 0.7:
                    confidence -= 0.2
                    
            # Factor 3: Check if it looks like actual tabular data
            # Look for patterns that suggest real tables
            has_header_like_first_row = False
            if len(table_data) >= 2:
                first_row = table_data[0]
                second_row = table_data[1]
                
                # First row has more text, second row has more numbers
                first_row_text = sum(1 for cell in first_row if cell and not self._is_numeric_like(str(cell)))
                second_row_numeric = sum(1 for cell in second_row if cell and self._is_numeric_like(str(cell)))
                
                if first_row_text > 0 and second_row_numeric > 0:
                    has_header_like_first_row = True
                    confidence += 0.1
                    
            return min(1.0, max(0.0, confidence))
            
        except Exception as e:
            self.logger.warning(f"Error calculating table confidence: {str(e)}")
            return 0.5

    def _detect_header(self, table_df) -> bool:
        """Detect if table has a header row"""
        try:
            # If pandas is available, use pandas methods
            if PANDAS_AVAILABLE and hasattr(table_df, 'iloc'):
                if len(table_df) < 2:
                    return False
                    
                first_row = table_df.iloc[0]
                second_row = table_df.iloc[1]
                
                # Check if first row has different data types or patterns
                first_row_numeric = sum(1 for x in first_row if str(x).replace('.', '').replace('-', '').isdigit())
                second_row_numeric = sum(1 for x in second_row if str(x).replace('.', '').replace('-', '').isdigit())
                
                # If first row has significantly fewer numeric values, it might be a header
                if len(first_row) > 0 and first_row_numeric / len(first_row) < 0.5 and second_row_numeric / len(second_row) > 0.5:
                    return True
            else:
                # Fallback to basic detection
                return self._detect_header_basic(table_df)
                
            return False
            
        except Exception:
            return False
            
    def _calculate_confidence(self, table_df, method: str) -> float:
        """Calculate confidence score for extracted table"""
        try:
            confidence = 0.0
            
            # Handle pandas DataFrame
            if PANDAS_AVAILABLE and hasattr(table_df, 'columns'):
                # Basic metrics
                if len(table_df) > 0 and len(table_df.columns) > 0:
                    confidence += 0.3
                    
                # Check for empty cells (lower confidence with many empty cells)
                total_cells = len(table_df) * len(table_df.columns)
                empty_cells = table_df.isnull().sum().sum()
                if total_cells > 0:
                    confidence += 0.4 * (1 - empty_cells / total_cells)
                    
                # Structure consistency
                if len(table_df) > 1:
                    # Check if columns have consistent data types
                    consistent_cols = 0
                    for col in table_df.columns:
                        try:
                            if PANDAS_AVAILABLE:
                                pd.to_numeric(table_df[col], errors='raise')
                            consistent_cols += 1
                        except:
                            # Check if it's consistently text
                            if hasattr(table_df[col], 'dtype') and table_df[col].dtype == 'object':
                                consistent_cols += 0.5
                                
                    if len(table_df.columns) > 0:
                        confidence += 0.2 * (consistent_cols / len(table_df.columns))
                        
            # Handle basic list format
            elif isinstance(table_df, list) and table_df:
                confidence += 0.3  # Basic table found
                
                # Check for consistent row lengths
                if len(table_df) > 1:
                    first_row_len = len(table_df[0]) if table_df[0] else 0
                    consistent_rows = sum(1 for row in table_df if len(row) == first_row_len)
                    if len(table_df) > 0:
                        confidence += 0.4 * (consistent_rows / len(table_df))
                        
            # Method-specific adjustments
            if method == "lattice":
                confidence += 0.1  # Lattice generally more reliable for bordered tables
                
            return min(confidence, 1.0)
            
        except Exception:
            return 0.5  # Default confidence
            
    def _deduplicate_tables(self, tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Remove duplicate tables and select best versions"""
        if not tables:
            return []
            
        # Group tables by approximate size and position
        unique_tables = []
        
        for table in tables:
            is_duplicate = False
            
            for existing in unique_tables:
                # Check if tables are similar (same dimensions, similar content)
                if (abs(table['rows'] - existing['rows']) <= 1 and 
                    abs(table['columns'] - existing['columns']) <= 1 and
                    table.get('page', 0) == existing.get('page', 0)):
                    
                    # Compare first few cells for similarity
                    if self._tables_similar(table['data'], existing['data']):
                        is_duplicate = True
                        
                        # Keep the one with higher confidence
                        if table.get('confidence', 0) > existing.get('confidence', 0):
                            unique_tables.remove(existing)
                            unique_tables.append(table)
                        break
                        
            if not is_duplicate:
                unique_tables.append(table)
                
        # Sort by page and confidence
        unique_tables.sort(key=lambda x: (x.get('page', 0), -x.get('confidence', 0)))
        
        return unique_tables
        
    def _tables_similar(self, df1, df2, threshold: float = 0.8) -> bool:
        """Check if two tables are similar"""
        try:
            # Handle pandas DataFrames
            if PANDAS_AVAILABLE and hasattr(df1, 'shape') and hasattr(df2, 'shape'):
                if df1.shape != df2.shape:
                    return False
                    
                # Compare first few cells
                matches = 0
                total = 0
                
                for i in range(min(3, len(df1))):
                    for j in range(min(2, len(df1.columns))):
                        val1 = str(df1.iloc[i, j]).strip()
                        val2 = str(df2.iloc[i, j]).strip()
                        
                        if val1 == val2:
                            matches += 1
                        total += 1
                        
                return (matches / total) >= threshold if total > 0 else False
                
            # Handle basic list format
            elif isinstance(df1, list) and isinstance(df2, list):
                if len(df1) != len(df2):
                    return False
                    
                if not df1 or not df2:
                    return df1 == df2
                    
                if len(df1[0]) != len(df2[0]):
                    return False
                    
                matches = 0
                total = 0
                
                for i in range(min(3, len(df1))):
                    for j in range(min(2, len(df1[0]))):
                        val1 = str(df1[i][j]).strip() if df1[i][j] else ""
                        val2 = str(df2[i][j]).strip() if df2[i][j] else ""
                        
                        if val1 == val2:
                            matches += 1
                        total += 1
                        
                return (matches / total) >= threshold if total > 0 else False
            
        except Exception:
            pass
            
        return False
            
    def _estimate_page_number(self, pdf_path: str, pages_spec: str) -> int:
        """Estimate page number from pages specification"""
        try:
            if pages_spec == "all":
                return 1
            elif "-" in pages_spec:
                return int(pages_spec.split("-")[0])
            else:
                return int(pages_spec)
        except:
            return 1
            
    def get_table_summary(self, tables: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Get summary statistics for extracted tables"""
        if not tables:
            return {"total_tables": 0}
            
        summary = {
            "total_tables": len(tables),
            "pages_with_tables": len(set(t.get('page', 0) for t in tables)),
            "extraction_methods": list(set(t.get('extraction_method', 'unknown') for t in tables)),
            "average_confidence": sum(t.get('confidence', 0) for t in tables) / len(tables),
            "table_sizes": [(t['rows'], t['columns']) for t in tables],
            "total_rows": sum(t['rows'] for t in tables),
            "total_columns": sum(t['columns'] for t in tables)
        }
        
        return summary
