import logging
import os.path
import re

import openpyxl

logger = logging.getLogger("desktopenv.metric.table")


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
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text to verify it exists
                formula_text = None
                if hasattr(cell, "_value") and isinstance(cell._value, str) and cell._value.startswith("="):
                    formula_text = cell._value
                elif hasattr(cell, "formula"):
                    formula_text = cell.formula
                else:
                    # Try to get from value attribute
                    if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_text = cell.value
                
                if formula_text is None:
                    logger.warning(f"Could not extract formula from cell {cell_coord}")
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

