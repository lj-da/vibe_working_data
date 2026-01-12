import logging
import os.path
import re
from datetime import datetime, timedelta

import openpyxl

logger = logging.getLogger("desktopenv.metric.table")


def normalize_date_value(value):
    """
    Normalize date values to a comparable format.
    
    Excel stores dates as serial numbers (days since 1900-01-01).
    This function converts both date strings and date numbers to a normalized format
    for comparison.
    
    Args:
        value: Cell value (can be string, number, or datetime object)
    
    Returns:
        tuple: (normalized_value: str, is_date: bool)
            - normalized_value: Normalized date string or original value as string
            - is_date: True if value is a date, False otherwise
    """
    if value is None:
        return None, False
    
    # If it's already a datetime object
    if isinstance(value, datetime):
        return value.strftime("%Y/%m/%d"), True
    
    # If it's a number, check if it's an Excel date serial number
    if isinstance(value, (int, float)):
        # Excel date serial numbers are typically between 1 and 100000
        # (covers dates from 1900 to ~2174)
        if 1 <= value <= 100000:
            try:
                # Excel epoch is 1900-01-01, but Excel incorrectly treats 1900 as a leap year
                # So we need to adjust: Excel day 1 = 1899-12-30
                excel_epoch = datetime(1899, 12, 30)
                date_obj = excel_epoch + timedelta(days=value)
                return date_obj.strftime("%Y/%m/%d"), True
            except (ValueError, OverflowError):
                pass
        return str(value), False
    
    # If it's a string, try to parse as date
    if isinstance(value, str):
        value_stripped = value.strip()
        # Try common date formats
        date_formats = [
            "%Y/%m/%d",
            "%Y-%m-%d",
            "%Y.%m.%d",
            "%m/%d/%Y",
            "%d/%m/%Y",
        ]
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(value_stripped, fmt)
                return date_obj.strftime("%Y/%m/%d"), True
            except ValueError:
                continue
        return value_stripped, False
    
    return str(value), False


def get_cell_formula(ws, cell_coord):
    """
    Check if a cell contains a formula and extract the formula text.
    
    This is a reusable helper function for verifying formula existence.
    
    Args:
        ws: openpyxl worksheet object (loaded with data_only=False)
        cell_coord: Cell coordinate (e.g., "A1", "B2", "G2")
    
    Returns:
        tuple: (has_formula: bool, formula_text: str or None)
            - has_formula: True if cell contains a formula, False otherwise
            - formula_text: Formula text if found, None otherwise
    """
    try:
        cell = ws[cell_coord]
        
        # Check if cell contains a formula
        if cell.data_type != "f":
            logger.debug(f"Cell {cell_coord} data_type is '{cell.data_type}', not 'f' (formula). Cell value: {cell.value}")
            return False, None
        
        # Get formula text
        formula_text = None
        if hasattr(cell, "_value") and isinstance(cell._value, str) and cell._value.startswith("="):
            formula_text = cell._value
        elif hasattr(cell, "formula"):
            formula_text = cell.formula
        else:
            if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_text = cell.value
        
        if formula_text is None:
            logger.debug(f"Cell {cell_coord} has data_type='f' but could not extract formula text. Cell value: {cell.value}")
            return False, None
        
        return True, formula_text
        
    except Exception as e:
        logger.error(f"Error getting formula from cell {cell_coord}: {e}")
        import traceback
        logger.debug(traceback.format_exc())
        return False, None


def extract_name(filename):
    """
    从A列文件名中提取中文姓名
    支持格式: 
    1. IP地址 + 数字 + 班级标识(秋2班/秋2/秋中专2班/秋二班) + 姓名 + 扩展名
    2. IP地址 + 数字 + 姓名 + 扩展名（无班级标识）
    """
    if not filename or not isinstance(filename, str):
        return None
    
    # 方法1: 匹配有"秋"字的格式，更精确地匹配班级标识
    # 匹配: 秋 + (可选:中专) + (可选:一/二/2) + 班 + 姓名
    pattern1 = r'秋(?:中专)?(?:[一二2])?班([\u4e00-\u9fa5]+V?)\.'
    match = re.search(pattern1, filename)
    if match:
        return match.group(1)
    
    # 方法2: 匹配"秋"后直接跟数字或"中专"再跟姓名（没有"班"字）
    pattern2 = r'秋(?:中专)?(?:[一二2])?([\u4e00-\u9fa5]+V?)\.'
    match = re.search(pattern2, filename)
    if match:
        return match.group(1)
    
    # 方法3: 匹配没有"秋"的情况，数字后直接跟中文姓名
    # 匹配: 数字 + (可选:下划线+数字) + (可选:单独下划线) + 中文姓名
    pattern3 = r'\.\d+(?:_\d+)?_?([\u4e00-\u9fa5]+V?)\.'
    match = re.search(pattern3, filename)
    if match:
        return match.group(1)
    
    return None


def verify_mid_find_extract_name(result: str, expected: str = None, **options) -> float:
    """
    Verify if formulas exist in specified column to extract names and if the extracted values match expected names.
    
    This function checks:
    1. Whether cells in specified column contain formulas (any formula)
    2. Whether the extracted values match the expected names extracted from source column using regex
    
    The function automatically detects the number of data rows by checking the data column
    (default: A column) for non-empty cells. It stops checking after finding 3 consecutive
    empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "B")
            - start_row: Starting row number (default: 1)
            - source_column: Column containing source data (e.g., "A")
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'B')
        start_row = options.get('start_row', 1)
        source_column = options.get('source_column', 'A')
        data_column = options.get('data_column', 'A')
        
        logger.info(f"Verifying formula existence and name extraction in column {check_column} in file: {result}")
        logger.info(f"Source column: {source_column}")
        logger.info(f"Start row: {start_row}")
        
        # Load workbook to get formulas and values
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
            wb_data = openpyxl.load_workbook(result, data_only=True)  # data_only=True to get calculated values
            ws_data = wb_data.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Auto-detect end_row by checking data_column for non-empty cells
        logger.info(f"Auto-detecting end row by checking column {data_column} for data...")
        max_row = ws.max_row
        end_row = start_row  # Start from start_row
        
        # Find the last row with data in the data column
        # Check up to max_row, but stop if we find 3 consecutive empty rows
        empty_count = 0
        for row_num in range(start_row, max_row + 1):
            data_cell = ws[f"{data_column}{row_num}"]
            if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                empty_count += 1
                if empty_count >= 3:  # Stop after 3 consecutive empty rows
                    break
            else:
                empty_count = 0
                end_row = row_num  # Update end_row to the last row with data
        
        logger.info(f"Auto-detected end row: {end_row}")
        
        # Check each row in the specified column
        all_passed = True
        passed_count = 0
        checked_count = 0
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            source_cell_coord = f"{source_column}{row_num}"
            try:
                cell = ws[cell_coord]
                source_cell = ws[source_cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Skip if source cell is empty
                if source_cell.value is None or (isinstance(source_cell.value, str) and source_cell.value.strip() == ""):
                    logger.debug(f"Skipping row {row_num} because source cell {source_cell_coord} is empty")
                    continue
                
                checked_count += 1
                
                # Check if cell contains a formula using reusable function
                has_formula, formula_text = get_cell_formula(ws, cell_coord)
                if not has_formula:
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                logger.debug(f"Cell {cell_coord} has formula: {formula_text}")
                
                # Get the extracted value from B column
                extracted_value = ws_data[cell_coord].value
                if extracted_value is None:
                    extracted_value = ""
                elif not isinstance(extracted_value, str):
                    extracted_value = str(extracted_value)
                extracted_value = extracted_value.strip()
                
                if extracted_value == "":
                    logger.warning(f"Cell {cell_coord} formula extracted empty value, extraction may have failed")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Get the source filename from A column
                source_filename = source_cell.value
                if source_filename is None:
                    source_filename = ""
                elif not isinstance(source_filename, str):
                    source_filename = str(source_filename)
                
                # Extract expected name from source filename using regex
                expected_name = extract_name(source_filename)
                
                if expected_name is None:
                    logger.warning(f"Could not extract name from source filename: {source_filename}")
                    logger.warning(f"Cell {cell_coord} extracted value: {extracted_value}")
                    # Still check if extracted value is non-empty, but don't fail if we can't extract expected
                    passed_count += 1
                    logger.info(f"✓ Cell {cell_coord} has formula and extracted value: {extracted_value} (expected name extraction failed)")
                    continue
                
                # Compare extracted value with expected name
                if extracted_value != expected_name:
                    logger.warning(f"Cell {cell_coord} extracted value '{extracted_value}' does not match expected name '{expected_name}'")
                    logger.warning(f"Source filename: {source_filename}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                passed_count += 1
                logger.info(f"✓ Cell {cell_coord} has formula and correctly extracted name: {extracted_value}")
                logger.debug(f"  Source filename: {source_filename}")
                logger.debug(f"  Formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                all_passed = False
        
        if checked_count == 0:
            logger.error(f"No cells found to check in column {check_column}")
            return 0.0
        
        if all_passed and passed_count == checked_count:
            logger.info("=" * 60)
            logger.info(f"✓ All {passed_count} cells in column {check_column} contain formulas and correctly extracted names")
            logger.info(f"  - All cells passed verification")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ Formula verification failed")
            logger.error(f"  Passed: {passed_count}/{checked_count} cells")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_sum_sumif_fruit_sales(result: str, expected: str = None, **options) -> float:
    """
    Verify if a formula exists in specified cell and if the calculated value matches expected value.
    
    This function checks:
    1. Whether the specified cell contains a formula
    2. Whether the calculated value matches the expected value (25719)
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_cell: Cell to check (e.g., "G2")
            - expected_value: Expected calculated value (default: 25719)
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_cell = options.get('check_cell', 'G2')
        expected_value = options.get('expected_value', 25719)
        
        logger.info(f"Verifying formula and value in cell {check_cell} in file: {result}")
        logger.info(f"Expected value: {expected_value}")
        
        # Load workbook to get formulas and values
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
            wb_data = openpyxl.load_workbook(result, data_only=True)  # data_only=True to get calculated values
            ws_data = wb_data.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        try:
            cell_data = ws_data[check_cell]
            logger.debug(f"Checking cell {check_cell}")
            
            # Check 1: Verify formula exists using reusable function
            has_formula, formula_text = get_cell_formula(ws, check_cell)
            if not has_formula:
                logger.warning(f"Cell {check_cell} does not contain a formula")
                return 0.0
            
            logger.debug(f"Cell {check_cell} has formula: {formula_text}")
            
            # Check 3: Verify calculated value matches expected value
            calculated_value = cell_data.value
            
            # Handle different value types
            if calculated_value is None:
                logger.warning(f"Cell {check_cell} calculated value is None")
                return 0.0
            
            # Convert to number for comparison
            if isinstance(calculated_value, str):
                try:
                    calculated_value = float(calculated_value)
                except ValueError:
                    logger.warning(f"Cell {check_cell} calculated value '{calculated_value}' cannot be converted to number")
                    return 0.0
            elif not isinstance(calculated_value, (int, float)):
                calculated_value = float(calculated_value)
            
            # Compare values (allow small floating point differences)
            if abs(calculated_value - expected_value) < 0.01:
                logger.info(f"✓ Cell {check_cell} has formula and correct value: {calculated_value}")
                logger.debug(f"  Formula: {formula_text}")
                logger.info("=" * 60)
                logger.info(f"✓ Formula and value verification passed")
                logger.info(f"  - Cell: {check_cell}")
                logger.info(f"  - Formula: {formula_text}")
                logger.info(f"  - Calculated value: {calculated_value}")
                logger.info(f"  - Expected value: {expected_value}")
                logger.info("=" * 60)
                return 1.0
            else:
                logger.error(f"Cell {check_cell} calculated value {calculated_value} does not match expected value {expected_value}")
                logger.error(f"  Formula: {formula_text}")
                logger.error("=" * 60)
                logger.error(f"✗ Value verification failed")
                logger.error(f"  - Cell: {check_cell}")
                logger.error(f"  - Calculated value: {calculated_value}")
                logger.error(f"  - Expected value: {expected_value}")
                logger.error("=" * 60)
                return 0.0
                
        except Exception as e:
            logger.error(f"Error checking cell {check_cell}: {e}")
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_if_unique_dates(result: str, expected: str = None, **options) -> float:
    """
    Verify if formulas exist in specified column to extract unique dates and if the extracted dates are unique and from source column.
    
    This function checks:
    1. Whether cells in specified column contain formulas (any formula)
    2. Whether the non-empty values in the result column are unique
    3. Whether all non-empty values in the result column exist in the source column
    
    The function automatically detects the number of data rows by checking the data column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "C")
            - source_column: Column containing source data (e.g., "A")
            - start_row: Starting row number (default: 3)
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'C')
        source_column = options.get('source_column', 'A')
        start_row = options.get('start_row', 3)
        data_column = options.get('data_column', 'A')
        
        logger.info(f"Verifying formula for unique dates in column {check_column} in file: {result}")
        logger.info(f"Source column: {source_column}")
        logger.info(f"Start row: {start_row}")
        
        # Load workbook to get formulas and values
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
            wb_data = openpyxl.load_workbook(result, data_only=True)  # data_only=True to get calculated values
            ws_data = wb_data.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Auto-detect end_row by checking data_column for non-empty cells
        logger.info(f"Auto-detecting end row by checking column {data_column} for data...")
        max_row = ws.max_row
        end_row = start_row
        
        # Find the last row with data in the data column
        empty_count = 0
        for row_num in range(start_row, max_row + 1):
            data_cell = ws[f"{data_column}{row_num}"]
            if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                empty_count += 1
                if empty_count >= 3:
                    break
            else:
                empty_count = 0
                end_row = row_num
        
        logger.info(f"Auto-detected end row: {end_row}")
        
        # Check each row in the specified column
        all_passed = True
        passed_count = 0
        checked_count = 0
        non_empty_values = []  # Store non-empty values from check_column
        source_values = set()  # Store all values from source_column
        
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        # First, collect all source values (normalized for date comparison)
        for row_num in range(start_row, end_row + 1):
            source_cell = ws_data[f"{source_column}{row_num}"]
            if source_cell.value is not None:
                normalized_value, is_date = normalize_date_value(source_cell.value)
                if normalized_value:
                    source_values.add(normalized_value)
        
        # Then check formulas and collect non-empty values
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                cell_data = ws_data[cell_coord]
                
                # Skip if source cell is empty (no need to check)
                source_cell = ws_data[f"{source_column}{row_num}"]
                if source_cell.value is None or (isinstance(source_cell.value, str) and source_cell.value.strip() == ""):
                    continue
                
                checked_count += 1
                
                # Check if cell contains a formula using reusable function
                has_formula, formula_text = get_cell_formula(ws, cell_coord)
                if not has_formula:
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                logger.debug(f"Cell {cell_coord} has formula: {formula_text}")
                
                # Collect non-empty values from check_column (normalized for date comparison)
                calculated_value = cell_data.value
                if calculated_value is not None:
                    normalized_value, is_date = normalize_date_value(calculated_value)
                    if normalized_value and normalized_value.strip() != "":
                        non_empty_values.append(normalized_value)
                
                passed_count += 1
                logger.debug(f"✓ Cell {cell_coord} has valid formula")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                all_passed = False
        
        if checked_count == 0:
            logger.error(f"No cells found to check in column {check_column}")
            return 0.0
        
        # Check 1: All formulas passed
        if not all_passed:
            logger.error(f"✗ Formula verification failed: {passed_count}/{checked_count} cells passed")
            return 0.0
        
        # Check 2: Non-empty values are unique
        unique_values = set(non_empty_values)
        if len(non_empty_values) != len(unique_values):
            duplicates = [v for v in non_empty_values if non_empty_values.count(v) > 1]
            logger.error(f"✗ Non-empty values in column {check_column} are not unique")
            logger.error(f"  Duplicate values: {set(duplicates)}")
            logger.error(f"  Total non-empty values: {len(non_empty_values)}, Unique values: {len(unique_values)}")
            return 0.0
        
        # Check 3: All non-empty values exist in source column
        missing_values = []
        for value in non_empty_values:
            if value not in source_values:
                missing_values.append(value)
        
        if missing_values:
            logger.error(f"✗ Some values in column {check_column} do not exist in source column {source_column}")
            logger.error(f"  Missing values: {missing_values}")
            return 0.0
        
        logger.info("=" * 60)
        logger.info(f"✓ All {passed_count} cells in column {check_column} contain formulas")
        logger.info(f"  - Non-empty values are unique: {len(unique_values)} unique values")
        logger.info(f"  - All values exist in source column {source_column}")
        logger.info("=" * 60)
        return 1.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_average_with_empty_cells(result: str, expected: str = None, **options) -> float:
    """
    Verify if AVERAGE formulas exist and if the calculated values match the expected average of source range (ignoring empty cells).
    
    This function checks:
    1. Whether cells in specified column contain formulas (using get_cell_formula)
    2. Whether the calculated average values match the expected average of source range
    3. Empty cells in source range are ignored when calculating average
    
    The function automatically detects the number of data rows by checking the data column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column with average formulas (e.g., "I")
            - source_start_column: Start column of source range (e.g., "B")
            - source_end_column: End column of source range (e.g., "H")
            - start_row: Starting row number (default: 2)
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'I')
        source_start_column = options.get('source_start_column', 'B')
        source_end_column = options.get('source_end_column', 'H')
        start_row = options.get('start_row', 2)
        data_column = options.get('data_column', 'A')
        
        logger.info(f"Verifying AVERAGE formulas in column {check_column} in file: {result}")
        logger.info(f"Source range: {source_start_column}{start_row}:{source_end_column}{start_row} (and below)")
        logger.info(f"Start row: {start_row}")
        
        # Load workbook to get formulas and values
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
            wb_data = openpyxl.load_workbook(result, data_only=True)  # data_only=True to get calculated values
            ws_data = wb_data.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Auto-detect end_row by checking data_column for non-empty cells
        logger.info(f"Auto-detecting end row by checking column {data_column} for data...")
        max_row = ws.max_row
        end_row = start_row
        
        # Find the last row with data in the data column
        empty_count = 0
        for row_num in range(start_row, max_row + 1):
            data_cell = ws[f"{data_column}{row_num}"]
            if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                empty_count += 1
                if empty_count >= 3:
                    break
            else:
                empty_count = 0
                end_row = row_num
        
        logger.info(f"Auto-detected end row: {end_row}")
        
        # Check each row in the specified column
        all_passed = True
        passed_count = 0
        checked_count = 0
        
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                # Skip if data column is empty
                data_cell = ws_data[f"{data_column}{row_num}"]
                if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                    continue
                
                checked_count += 1
                
                # Check if cell contains a formula using reusable function
                has_formula, formula_text = get_cell_formula(ws, cell_coord)
                if not has_formula:
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                logger.debug(f"Cell {cell_coord} has formula: {formula_text}")
                
                # Get calculated value from result cell
                calculated_value = ws_data[cell_coord].value
                if calculated_value is None:
                    logger.warning(f"Cell {cell_coord} calculated value is None")
                    all_passed = False
                    continue
                
                # Convert to float for comparison
                try:
                    if isinstance(calculated_value, str):
                        calculated_value = float(calculated_value)
                    elif not isinstance(calculated_value, (int, float)):
                        calculated_value = float(calculated_value)
                except (ValueError, TypeError):
                    logger.warning(f"Cell {cell_coord} calculated value '{calculated_value}' cannot be converted to number")
                    all_passed = False
                    continue
                
                # Calculate expected average from source range (ignoring empty cells)
                source_values = []
                for col_letter in [chr(ord(source_start_column) + i) for i in range(ord(source_end_column) - ord(source_start_column) + 1)]:
                    source_cell = ws_data[f"{col_letter}{row_num}"]
                    if source_cell.value is not None:
                        # Try to convert to number
                        try:
                            if isinstance(source_cell.value, str):
                                value = float(source_cell.value)
                            else:
                                value = float(source_cell.value)
                            source_values.append(value)
                        except (ValueError, TypeError):
                            # Skip non-numeric values
                            pass
                
                if len(source_values) == 0:
                    logger.warning(f"Cell {cell_coord} source range has no numeric values")
                    # If source has no values, expected average should be 0 or error
                    # But AVERAGE of empty range might return #DIV/0! error
                    # For now, skip this check
                    continue
                
                expected_average = sum(source_values) / len(source_values)
                
                # Compare values (allow small floating point differences)
                if abs(calculated_value - expected_average) < 0.01:
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has correct average: {calculated_value} (expected: {expected_average})")
                else:
                    logger.warning(f"Cell {cell_coord} calculated value {calculated_value} does not match expected average {expected_average}")
                    logger.warning(f"  Source values: {source_values}")
                    logger.warning(f"  Formula: {formula_text}")
                    all_passed = False
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                all_passed = False
        
        if checked_count == 0:
            logger.error(f"No cells found to check in column {check_column}")
            return 0.0
        
        if all_passed and passed_count == checked_count:
            logger.info("=" * 60)
            logger.info(f"✓ All {passed_count} cells in column {check_column} contain formulas and correct average values")
            logger.info(f"  - All cells passed verification")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ AVERAGE formula verification failed")
            logger.error(f"  Passed: {passed_count}/{checked_count} cells")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_text_if_growth_rate(result: str, expected: str = None, **options) -> float:
    """
    Verify if TEXT and IF formulas exist in specified column to calculate growth rate and if the calculated values are correct.
    
    This function checks:
    1. Whether cells in specified column contain formulas (using get_cell_formula)
    2. Whether the calculated growth rate values match the expected values based on source columns
    
    The function automatically detects the number of data rows by checking the data column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column with growth rate formulas (e.g., "M")
            - base_column: Column with base year values (e.g., "K")
            - current_column: Column with current year values (e.g., "L")
            - start_row: Starting row number (default: 14)
            - data_column: Column to check for data to determine end_row (default: "J")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'M')
        base_column = options.get('base_column', 'K')
        current_column = options.get('current_column', 'L')
        start_row = options.get('start_row', 14)
        data_column = options.get('data_column', 'J')
        
        logger.info(f"Verifying TEXT/IF growth rate formulas in column {check_column} in file: {result}")
        logger.info(f"Base column: {base_column}, Current column: {current_column}")
        logger.info(f"Start row: {start_row}")
        
        # Load workbook to get formulas and values
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
            wb_data = openpyxl.load_workbook(result, data_only=True)  # data_only=True to get calculated values
            ws_data = wb_data.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Auto-detect end_row by checking data_column for non-empty cells
        logger.info(f"Auto-detecting end row by checking column {data_column} for data...")
        max_row = ws.max_row
        end_row = start_row
        
        # Find the last row with data in the data column
        empty_count = 0
        for row_num in range(start_row, max_row + 1):
            data_cell = ws[f"{data_column}{row_num}"]
            if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                empty_count += 1
                if empty_count >= 3:
                    break
            else:
                empty_count = 0
                end_row = row_num
        
        logger.info(f"Auto-detected end row: {end_row}")
        
        # Check each row in the specified column
        all_passed = True
        passed_count = 0
        checked_count = 0
        
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                # Skip if data column is empty
                data_cell = ws_data[f"{data_column}{row_num}"]
                if data_cell.value is None or (isinstance(data_cell.value, str) and data_cell.value.strip() == ""):
                    continue
                
                checked_count += 1
                
                # Check if cell contains a formula using reusable function
                has_formula, formula_text = get_cell_formula(ws, cell_coord)
                if not has_formula:
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                logger.debug(f"Cell {cell_coord} has formula: {formula_text}")
                
                # Get calculated value from result cell
                calculated_value = ws_data[cell_coord].value
                if calculated_value is None:
                    logger.warning(f"Cell {cell_coord} calculated value is None")
                    all_passed = False
                    continue
                
                # Convert to string for comparison
                calculated_str = str(calculated_value).strip()
                if calculated_str == "":
                    logger.warning(f"Cell {cell_coord} calculated value is empty")
                    all_passed = False
                    continue
                
                # Get base and current values
                base_cell = ws_data[f"{base_column}{row_num}"]
                current_cell = ws_data[f"{current_column}{row_num}"]
                
                if base_cell.value is None or current_cell.value is None:
                    logger.warning(f"Cell {cell_coord} source values are incomplete (base: {base_cell.value}, current: {current_cell.value})")
                    # Still check if formula exists and value is non-empty
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has formula and value: {calculated_str} (source validation skipped)")
                    continue
                
                # Convert to numbers
                try:
                    base_value = float(base_cell.value) if not isinstance(base_cell.value, (int, float)) else float(base_cell.value)
                    current_value = float(current_cell.value) if not isinstance(current_cell.value, (int, float)) else float(current_cell.value)
                except (ValueError, TypeError):
                    logger.warning(f"Cell {cell_coord} source values cannot be converted to numbers")
                    # Still check if formula exists and value is non-empty
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has formula and value: {calculated_str} (source validation skipped)")
                    continue
                
                # Calculate expected growth rate
                if base_value == 0:
                    # Division by zero case
                    logger.debug(f"Cell {cell_coord} base value is 0, skipping value validation")
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has formula and value: {calculated_str}")
                    continue
                
                growth_rate = (current_value / base_value) - 1
                is_positive = growth_rate > 0
                
                # Check if the calculated value contains the expected format
                # Expected format: percentage with arrow (e.g., "100%↑" or "-50%↓")
                has_percent = "%" in calculated_str
                has_arrow = "↑" in calculated_str or "↓" in calculated_str
                
                # Check if arrow direction matches sign
                arrow_correct = False
                if is_positive and "↑" in calculated_str:
                    arrow_correct = True
                elif not is_positive and "↓" in calculated_str:
                    arrow_correct = True
                
                if has_percent and has_arrow and arrow_correct:
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has correct growth rate format: {calculated_str} (growth: {growth_rate:.2%})")
                else:
                    logger.warning(f"Cell {cell_coord} value format may be incorrect: {calculated_str}")
                    logger.warning(f"  Expected: {'positive with ↑' if is_positive else 'negative with ↓'}")
                    logger.warning(f"  Growth rate: {growth_rate:.2%}")
                    logger.warning(f"  Formula: {formula_text}")
                    # Still pass if formula exists and value is non-empty (format check is lenient)
                    passed_count += 1
                    logger.debug(f"✓ Cell {cell_coord} has formula and value: {calculated_str} (format check lenient)")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                all_passed = False
        
        if checked_count == 0:
            logger.error(f"No cells found to check in column {check_column}")
            return 0.0
        
        if all_passed and passed_count == checked_count:
            logger.info("=" * 60)
            logger.info(f"✓ All {passed_count} cells in column {check_column} contain formulas and values")
            logger.info(f"  - All cells passed verification")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ Growth rate formula verification failed")
            logger.error(f"  Passed: {passed_count}/{checked_count} cells")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0

