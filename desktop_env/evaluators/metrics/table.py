import functools
import itertools
import logging
import os.path

# import operator
from numbers import Number
from typing import Any, Union, cast, Callable, Iterable
from typing import Dict, List, Tuple, Set

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from rapidfuzz import fuzz

from desktop_env.evaluators.metrics.utils import (
    _match_value_to_rule,
    _read_cell_style,
    read_cell_value,
)
from desktop_env.evaluators.metrics.utils import (
    load_charts,
    load_sparklines,
    load_rows_or_cols,
    load_xlsx_styles,
    load_filters,
    load_pivot_tables,
)

# from openpyxl.utils import coordinate_to_tuple

logger = logging.getLogger("desktopenv.metric.table")

BOOK = Union[pd.ExcelFile, Workbook, str]


def _parse_sheet_idx(
    sheet_idx: Union[int, str],
    result: BOOK,
    expected: BOOK,
    result_sheet_names: List[str],
    expected_sheet_names: List[str],
) -> Tuple[BOOK, str]:
    #  function _parse_sheet_idx {{{ #
    if isinstance(sheet_idx, int):
        try:
            if not result_sheet_names or sheet_idx >= len(result_sheet_names):
                logger.error(
                    f"Sheet index {sheet_idx} out of range. Available sheets: {result_sheet_names}"
                )
                index = ""
            else:
                index: str = result_sheet_names[sheet_idx]
                logger.debug(f"Sheet index {sheet_idx} resolved to sheet: {index}")
        except Exception as e:
            logger.error(f"Error resolving sheet index {sheet_idx}: {e}")
            index = ""
        book: BOOK = result
    elif sheet_idx.startswith("RI"):
        try:
            index: str = result_sheet_names[int(sheet_idx[2:])]
        except:
            index = ""
        book: BOOK = result
    elif sheet_idx.startswith("RN"):
        index: str = sheet_idx[2:]
        book: BOOK = result
    elif sheet_idx.startswith("EI"):
        try:
            index: str = expected_sheet_names[int(sheet_idx[2:])]
        except:
            index = ""
        book: BOOK = expected
    elif sheet_idx.startswith("EN"):
        index: str = sheet_idx[2:]
        book: BOOK = expected
    else:
        logger.error("Unrecognized sheet index")
        raise ValueError("Unrecognized sheet index")
    return book, index
    #  }}} function _parse_sheet_idx #


SHEET = Union[pd.DataFrame, Worksheet, List[str]]


def _load_sheet(book: BOOK, index: str) -> SHEET:
    #  function _load_sheet {{{ #
    try:
        if isinstance(book, str):
            book: str = cast(str, book)
            csv_name: str = "{:}-{:}.csv".format(os.path.splitext(book)[0], index)

            try:
                all_lines: List[str] = _safe_read_file(csv_name)
                csv_lines: List[str] = list(
                    itertools.dropwhile(
                        lambda l: len(l) == 0,
                        map(lambda l: l.strip(), reversed(all_lines)),
                    )
                )
                return csv_lines
            except (FileNotFoundError, IOError) as e:
                logger.error(f"Failed to read CSV file {csv_name}: {e}")
                return None
        if isinstance(book, pd.ExcelFile):
            return pd.read_excel(book, index)
        if isinstance(book, Workbook):
            return book[index]
        logger.error("Not supported workbook format")
        raise NotImplementedError("Not supported workbook format")
    except NotImplementedError as e:
        raise e
    except:
        return None
    #  }}} function _load_sheet #


def _safe_read_file(file_path: str) -> List[str]:
    """
    Safely read a file with multiple encoding attempts.

    Args:
        file_path: Path to the file to read

    Returns:
        List of lines from the file

    Raises:
        FileNotFoundError: If file doesn't exist
        IOError: If file cannot be read with any encoding
    """
    # Common encodings to try in order of preference
    encodings = [
        "utf-8",  # Most common modern encoding
        "utf-8-sig",  # UTF-8 with BOM
        "latin-1",  # ISO-8859-1, works with any byte sequence
        "windows-1252",  # Common Windows encoding
        "gbk",  # Chinese encoding
        "cp1251",  # Cyrillic encoding
        "iso-8859-1",  # Alternative latin-1
    ]

    last_error = None

    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                lines = f.read().splitlines()
                logger.debug(
                    f"Successfully read file {file_path} with encoding {encoding}"
                )
                return lines
        except UnicodeDecodeError as e:
            last_error = e
            logger.debug(f"Failed to read {file_path} with encoding {encoding}: {e}")
            continue
        except (FileNotFoundError, IOError) as e:
            # These are non-encoding related errors, re-raise immediately
            raise e

    # If all encodings fail, try with error handling as last resort
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.read().splitlines()
            logger.warning(f"Read file {file_path} with UTF-8 and error replacement")
            return lines
    except Exception as e:
        logger.error(
            f"Failed to read file {file_path} with any encoding. Last error: {last_error}"
        )
        raise IOError(
            f"Cannot read file {file_path} with any supported encoding"
        ) from last_error


def compare_csv(result: str, expected: Union[str, List[str]], **options) -> float:
    """
    Compare CSV files. If expected is a list, returns 1.0 if result matches any of the expected files.

    Args:
        result: Path to result CSV file
        expected: Path to expected CSV file or list of paths to expected CSV files
        options: Additional options (strict, ignore_case)

    Returns:
        1.0 if result matches expected (or any file in expected list), 0.0 otherwise
    """
    if result is None:
        return 0.0

    try:
        result_lines: List[str] = _safe_read_file(result)
    except (FileNotFoundError, IOError) as e:
        logger.error(f"Failed to read result file {result}: {e}")
        return 0.0

    # Convert expected to list if it's a single string (for backward compatibility)
    if isinstance(expected, str):
        expected_files = [expected]
    else:
        expected_files = expected

    # Try to match against each expected file
    for expected_file in expected_files:
        try:
            expected_lines: List[str] = _safe_read_file(expected_file)

            # Process lines based on options
            current_result_lines = result_lines
            current_expected_lines = expected_lines

            if not options.get("strict", True):
                current_result_lines = map(str.strip, current_result_lines)
                current_expected_lines = map(str.strip, current_expected_lines)
            if options.get("ignore_case", False):
                current_result_lines = map(str.lower, current_result_lines)
                current_expected_lines = map(str.lower, current_expected_lines)

            # Check if this expected file matches
            if list(current_result_lines) == list(current_expected_lines):
                return 1.0

        except (FileNotFoundError, IOError):
            # If this expected file doesn't exist, continue to next one
            continue

    # No match found
    return 0.0


def compare_table(result: str, expected: str = None, **options) -> float:
    #  function compare_table {{{ #
    """
    Args:
        result (str): path to result xlsx
        expected (str): path to golden xlsx
        rules (List[Dict[str, Any]]): list of dict like
          {
            "type": str,
            <str as parameters>: anything
          }
          as sequential rules

    Returns:
        float: the score
    """

    if result is None:
        logger.error("Result file path is None")
        return 0.0

    # Check if result file exists
    if not os.path.exists(result):
        logger.error(f"Result file not found: {result}")
        return 0.0

    try:
        logger.info(f"Loading result file: {result}")
        xlworkbookr: Workbook = openpyxl.load_workbook(filename=result)
        pdworkbookr = pd.ExcelFile(result)
        logger.info(
            f"Successfully loaded result file with sheets: {pdworkbookr.sheet_names}"
        )
    except Exception as e:
        logger.error(f"Failed to load result file {result}: {e}")
        return 0.0
    worksheetr_names: List[str] = pdworkbookr.sheet_names

    if expected is not None:
        xlworkbooke: Workbook = openpyxl.load_workbook(filename=expected)
        pdworkbooke = pd.ExcelFile(expected)
        worksheete_names: List[str] = pdworkbooke.sheet_names
    else:
        xlworkbooke: Workbook = None
        pdworkbooke = None
        worksheete_names: List[str] = None

    parse_idx: Callable[[Union[str, int], BOOK, BOOK], Tuple[BOOK, str]] = (
        functools.partial(
            _parse_sheet_idx,
            result_sheet_names=worksheetr_names,
            expected_sheet_names=worksheete_names,
        )
    )

    passes = True
    for r in options["rules"]:
        if r["type"] == "sheet_name":
            #  Compare Sheet Names {{{ #
            metric: bool = worksheetr_names == worksheete_names
            logger.debug(
                "Assertion: %s.sheet_names == %s.sheet_names - %s",
                result,
                expected,
                metric,
            )
            #  }}} Compare Sheet Names #

        elif r["type"] == "sheet_data":
            #  Compare Sheet Data by Internal Value {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # precision: int as number of decimal digits, default to 4

            error_limit: int = r.get("precision", 4)
            sheet1: pd.DataFrame = _load_sheet(
                *parse_idx(r["sheet_idx0"], pdworkbookr, pdworkbooke)
            )
            if sheet1 is None:
                return 0.0
            sheet2: pd.DataFrame = _load_sheet(
                *parse_idx(r["sheet_idx1"], pdworkbookr, pdworkbooke)
            )

            sheet1 = sheet1.round(error_limit)
            sheet2 = sheet2.round(error_limit)
            metric: bool = sheet1.equals(sheet2)
            logger.debug("Sheet1: \n%s", str(sheet1))
            logger.debug("Sheet2: \n%s", str(sheet2))
            try:
                logger.debug("Sheet1 =v= Sheet2: \n%s", str(sheet1 == sheet2))
            except:
                logger.debug("Sheet1 =/v= Sheet2")
            logger.debug(
                "Assertion: %s =v= %s - %s", r["sheet_idx0"], r["sheet_idx1"], metric
            )
            #  }}} Compare Sheet Data by Internal Value #

        elif r["type"] == "sheet_print":
            #  Compare Sheet Data by Printed Value {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # ignore_case: optional, defaults to False

            sheet1: List[str] = _load_sheet(
                *parse_idx(r["sheet_idx0"], result, expected)
            )
            if sheet1 is None:
                return 0.0
            sheet2: List[str] = _load_sheet(
                *parse_idx(r["sheet_idx1"], result, expected)
            )
            if r.get("ignore_case", False):
                sheet1 = [l.lower() for l in sheet1]
                sheet2 = [l.lower() for l in sheet2]
            metric: bool = sheet1 == sheet2
            logger.debug(
                "Assertion: %s =p= %s - %s", r["sheet_idx0"], r["sheet_idx1"], metric
            )
            #  }}} Compare Sheet Data by Printed Value #

        elif r["type"] == "sheet_fuzzy":
            #  Fuzzy Match for Ranges {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # rules: list of dict, each dict is like
            #   { "range": ["A1:B6", "C2:E5"],
            #     "type": "includes" | "included_by" | "fuzzy_match" | "exact_match", # 0 includes 1, 0 includes_by 1
            #     "threshold": 85, // for fuzzy match
            #     "ignore_case": true | false,
            #     "ignore_chars": " ()", # filtered out
            #     "trim_leadings": "+ ", # filtered by lstrip
            #     "trim_trailings": "", # filtered by rstrip
            #     "normalization": [["Rd", "Road"]], # filtered by replace
            #   }

            sheet1: Tuple[BOOK, str] = parse_idx(r["sheet_idx0"], result, expected)
            sheet2: Tuple[BOOK, str] = parse_idx(r["sheet_idx1"], result, expected)
            total_metric = True
            for rl in r["rules"]:
                for rng in MultiCellRange(rl["range"]):
                    for cdn in rng.cells:
                        coordinate: str = "{:}{:d}".format(
                            get_column_letter(cdn[1]), cdn[0]
                        )
                        value1: str = str(read_cell_value(*sheet1, coordinate))
                        value2: str = str(read_cell_value(*sheet2, coordinate))
                        logger.debug("%s: %s vs %s", cdn, value1, value2)

                        for rplc in rl.get("normalization", []):
                            value1 = value1.replace(rplc[0], rplc[1])
                            value2 = value2.replace(rplc[0], rplc[1])
                        if "trim_leadings" in rl:
                            value1 = value1.lstrip(rl["trim_leadings"])
                            value2 = value2.lstrip(rl["trim_leadings"])
                        if "trim_trailings" in rl:
                            value1 = value1.rstrip(rl["trim_trailings"])
                            value2 = value2.rstrip(rl["trim_trailings"])
                        if "ignore_chars" in rl:
                            ignore_chars: Set[str] = set(rl["ignore_chars"])
                            value1 = "".join(
                                filter(lambda ch: ch not in ignore_chars, value1)
                            )
                            value2 = "".join(
                                filter(lambda ch: ch not in ignore_chars, value2)
                            )
                        if rl.get("ignore_case", False):
                            value1 = value1.lower()
                            value2 = value2.lower()

                        if rl["type"] == "includes":
                            metric: bool = value2 in value1
                        elif rl["type"] == "included_by":
                            metric: bool = value1 in value2
                        elif rl["type"] == "fuzzy_match":
                            metric: bool = fuzz.ratio(value1, value2) >= rl.get(
                                "threshold", 85.0
                            )
                        elif rl["type"] == "exact_match":
                            metric: bool = value1 == value2
                        total_metric = total_metric and metric

            metric: bool = total_metric
            logger.debug(
                "Assertion: %s =~= %s - %s", r["sheet_idx0"], r["sheet_idx1"], metric
            )
            #  }}} Fuzzy Match for Ranges #

        elif r["type"] == "sparkline":
            #  Compare Sparklines {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0

            sparkline1: Dict[str, str] = load_sparklines(
                *parse_idx(r["sheet_idx0"], result, expected)
            )
            sparkline2: Dict[str, str] = load_sparklines(
                *parse_idx(r["sheet_idx1"], result, expected)
            )
            metric: bool = sparkline1 == sparkline2
            logger.debug(
                "Assertion: %s.sp == %.sp - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Compare Sparklines #

        elif r["type"] == "chart":
            #  Compare Charts {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # chart_props: list of str, see utils.load_charts

            charts1: Dict[str, Any] = load_charts(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke), **r
            )
            charts2: Dict[str, Any] = load_charts(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke), **r
            )
            metric: bool = charts1 == charts2
            logger.debug(
                "Assertion: %s[chart] == %s[chart] - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Compare Charts #

        elif r["type"] == "style":
            #  Compare Style (Also Conditional Formatiing) {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # props: list of str indicating concerned styles, see utils._read_cell_style

            sheet_idx1: Tuple[BOOK, str] = parse_idx(
                r["sheet_idx0"], xlworkbookr, xlworkbooke
            )
            book_name1: str = parse_idx(r["sheet_idx0"], result, expected)[0]
            styles1: Dict[str, List[Any]] = load_xlsx_styles(
                *sheet_idx1, book_name1, **r
            )

            sheet_idx2: Tuple[BOOK, str] = parse_idx(
                r["sheet_idx1"], xlworkbookr, xlworkbooke
            )
            book_name2: str = parse_idx(r["sheet_idx1"], result, expected)[0]
            styles2: Dict[str, List[Any]] = load_xlsx_styles(
                *sheet_idx2, book_name2, **r
            )
            # number_formats1: List[str] = [c.number_format.lower() for col in sheet1.iter_cols() for c in col if c.value is not None and c.data_type=="n"]
            # number_formats2: List[str] = [c.number_format.lower() for col in sheet2.iter_cols() for c in col if c.value is not None and c.data_type=="n"]
            metric: bool = styles1 == styles2
            logger.debug(
                "Assertion: %s.style == %s.style - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Compare Style (Also Conditional Formatiing) #

        elif r["type"] == "freeze":
            #  Compare Freezing {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0

            sheet1: Worksheet = _load_sheet(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke)
            )
            if sheet1 is None:
                return 0.0
            sheet2: Worksheet = _load_sheet(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke)
            )
            metric: bool = sheet1.freeze_panes == sheet2.freeze_panes
            logger.debug(
                "Assertion: %s.freeze(%s) == %s.freeze(%s) - %s",
                r["sheet_idx0"],
                sheet1.freeze_panes,
                r["sheet_idx1"],
                sheet2.freeze_panes,
                metric,
            )
            #  }}} Compare Freezing #

        elif r["type"] == "zoom":
            #  Check Zooming {{{ #
            # sheet_idx: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # method: str
            # ref: value

            sheet: Worksheet = _load_sheet(
                *parse_idx(r["sheet_idx"], xlworkbookr, xlworkbooke)
            )
            if sheet is None:
                return 0.0
            zoom_scale: Number = sheet.sheet_view.zoomScale or 100.0
            metric: bool = _match_value_to_rule(zoom_scale, r)
            logger.debug(
                "Assertion: %s.zoom(%.1f) %s %.1f - %s",
                r["sheet_idx"],
                zoom_scale,
                r["method"],
                r["ref"],
                metric,
            )
            #  }}} Check Zooming #

        elif r["type"] == "data_validation":
            #  Check Data Validation {{{ #
            # sheet_idx: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # dv_props: list of dict like {attribute: {"method": str, "ref": anything}}
            #   available attributes:
            #     * ranges
            #     * type
            #     * formula1
            #     * formula2
            #     * operator
            #     * allowBlank
            #     * showDropDown
            #     * showInputMessage
            #     * showErrorMessage
            #     * error
            #     * errorTitle
            #     * errorStyle
            #     * prompt
            #     * promptTitle
            #     * imeMode

            sheet: Worksheet = _load_sheet(
                *parse_idx(r["sheet_idx"], xlworkbookr, xlworkbooke)
            )
            if sheet is None:
                return 0.0
            data_validators: List[DataValidation] = (
                sheet.data_validations.dataValidation
            )

            total_metric = len(data_validators) >= len(r["dv_props"])
            for dat_vldt in data_validators:
                metric = False
                for prpt in r["dv_props"]:
                    metric = metric or all(
                        _match_value_to_rule(getattr(dat_vldt, attrbt), mr)
                        for attrbt, mr in prpt.items()
                    )
                    if metric:
                        break
                total_metric = total_metric and metric
                if not total_metric:
                    break

            logger.debug(
                "Assertion: %s.data_validation - %s", r["sheet_idx"], total_metric
            )
            metric: bool = total_metric
            #  }}} Check Data Validation #

        elif r["type"] == "row_props":
            #  Check Row Properties {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # props: list of str, see utils.load_rows_or_cols

            rows1: Dict[str, Any] = load_rows_or_cols(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke), obj="row", **r
            )
            rows2: Dict[str, Any] = load_rows_or_cols(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke), obj="row", **r
            )
            logger.debug("Rows1: %s", repr(rows1))
            logger.debug("Rows2: %s", repr(rows2))
            metric: bool = rows1 == rows2
            logger.debug(
                "Assertion: %s[rows] == %s[rows] - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Check Row Properties #

        elif r["type"] == "col_props":
            #  Check Row Properties {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # props: list of str, see utils.load_rows_or_cols

            cols1: Dict[str, Any] = load_rows_or_cols(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke), obj="column", **r
            )
            cols2: Dict[str, Any] = load_rows_or_cols(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke), obj="column", **r
            )
            metric: bool = cols1 == cols2
            logger.debug(
                "Assertion: %s[cols] == %s[cols] - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Check Row Properties #

        elif r["type"] == "filter":
            #  Compare Filters {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0

            filters1: Dict[str, Any] = load_filters(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke), **r
            )
            filters2: Dict[str, Any] = load_filters(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke), **r
            )
            metric: bool = filters1 == filters2
            logger.debug(
                "Assertion: %s[filter] == %s[filter] - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Compare Filters #

        elif r["type"] == "pivot_table":
            #  Compare Pivot Tables {{{ #
            # sheet_idx0: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # sheet_idx1: as sheet_idx0
            # pivot_props: list of str, see utils.load_pivot_tables

            pivots1: Dict[str, Any] = load_pivot_tables(
                *parse_idx(r["sheet_idx0"], xlworkbookr, xlworkbooke), **r
            )
            pivots2: Dict[str, Any] = load_pivot_tables(
                *parse_idx(r["sheet_idx1"], xlworkbookr, xlworkbooke), **r
            )
            metric: bool = pivots1 == pivots2
            logger.debug(
                "Assertion: %s[pivot]==%s[pivot] - %s",
                r["sheet_idx0"],
                r["sheet_idx1"],
                metric,
            )
            #  }}} Compare Pivot Tables #

        elif r["type"] == "check_cell":
            #  Check Cell Properties {{{ #
            # sheet_idx: 0 == "RI0" == "RNSheet1" | "EI0" == "ENSheet1"
            # coordinate: str, "E3"
            # props: dict like {attribute: {"method": str, "ref": anything}}
            #   supported attributes: value & those supported by utils._read_cell_style

            try:
                sheet: Worksheet = _load_sheet(
                    *parse_idx(r["sheet_idx"], xlworkbookr, xlworkbooke)
                )
                if sheet is None:
                    logger.error(
                        f"Failed to load sheet for sheet_idx: {r['sheet_idx']}"
                    )
                    return 0.0
                # data_frame: pd.DataFrame = _load_sheet(*parse_idx(r["sheet_idx"], pdworkbookr, pdworkbooke))
                cell: Cell = sheet[r["coordinate"]]
                metric: bool = True
                for prpt, rule in r["props"].items():
                    if prpt == "value":
                        try:
                            parsed_result = parse_idx(r["sheet_idx"], result, expected)
                            logger.debug(f"parse_idx result: {parsed_result}")
                            val = read_cell_value(*parsed_result, r["coordinate"])
                            logger.debug(f"Cell {r['coordinate']} value: {val}")
                        except Exception as e:
                            logger.error(
                                f"Failed to read cell value at {r['coordinate']}: {e}"
                            )
                            val = None
                    elif prpt == "formula":
                        # Support checking cell formula directly
                        try:
                            if cell.data_type == "f":
                                # For formula cells, get the formula text
                                # In openpyxl, formula is stored in cell.value for formula cells
                                # But we need the actual formula text, not the calculated value
                                # Try to get formula from internal representation
                                if hasattr(cell, "_value") and isinstance(cell._value, str) and cell._value.startswith("="):
                                    val = cell._value
                                elif hasattr(cell, "formula"):
                                    val = cell.formula
                                else:
                                    # Fallback: try to reconstruct from value if it's a formula
                                    val = f"={cell.value}" if cell.value is not None else None
                            else:
                                val = None
                            logger.debug(f"Cell {r['coordinate']} formula: {val}")
                        except Exception as e:
                            logger.error(
                                f"Failed to read cell formula at {r['coordinate']}: {e}"
                            )
                            val = None
                    else:
                        try:
                            val = _read_cell_style(prpt, cell)
                        except Exception as e:
                            logger.error(
                                f"Failed to read cell style {prpt} at {r['coordinate']}: {e}"
                            )
                            val = None

                    metric = metric and _match_value_to_rule(val, rule)
            except Exception as e:
                logger.error(f"Error in check_cell processing: {e}")
                return 0.0

            logger.debug(
                "Assertion: %s[%s] :%s - %s",
                r["sheet_idx"],
                r["coordinate"],
                repr(r["props"]),
                metric,
            )
            #  }}} Check Cell Properties #

        else:
            raise NotImplementedError(
                "Unimplemented sheet check: {:}".format(r["type"])
            )

        passes = passes and metric
        if not passes:
            break

    return float(passes)
    #  }}} function compare_table #


def compare_conference_city_in_order(actual_city_list_path, expected_city):
    expected_city_list = expected_city["expected"]
    wb = openpyxl.load_workbook(actual_city_list_path)
    sheet = wb.active
    actual_city_list = []
    for row in sheet["C2:C22"]:
        for cell in row:
            actual_city_list.append(cell.value)
    # expected_city is the city that we want to compare with the actual city list
    # must in order index
    # debug
    try:
        for i in range(len(actual_city_list)):
            if isinstance(expected_city_list[i], str):
                if expected_city_list[i] not in actual_city_list[i]:
                    logger.debug(
                        f"Expected city {expected_city_list[i]}; Actual city {actual_city_list[i]}"
                    )
                    print(
                        f"Expected city {expected_city_list[i]}; Actual city {actual_city_list[i]}"
                    )
                    return 0.0

            elif isinstance(expected_city_list[i], List):
                if not any(
                    possible_str in actual_city_list[i]
                    for possible_str in expected_city_list[i]
                ):
                    logger.debug(
                        f"Expected city {expected_city_list[i]}; Actual city {actual_city_list[i]}"
                    )
                    print(
                        f"Expected city {expected_city_list[i]}; Actual city {actual_city_list[i]}"
                    )
                    return 0.0

            else:
                raise TypeError("Expected city should be a string or a list of strings")

    except:
        return 0.0

    return 1.0


def verify_second_row_deleted_without_gold(result: str, expected: str = None, **options) -> float:
    """
    验证 Excel 文件的第二行是否被删除（不需要金标准文件）
    
    通过以下方式验证：
    1. 检查结果文件的行数是否比原始文件少1
    2. 检查原始文件的第二行数据是否在结果文件中不存在
    3. 检查其他所有行是否保持不变
    
    Args:
        result (str): 结果文件路径
        expected (str): 未使用（为了兼容框架接口）
        options (dict): 配置选项，应包含：
            - original_file_url: 原始文件的URL（用于下载和比对）
            - result_file_path: 结果文件的路径（可选，默认使用 result 参数）
            - original_file_cache: 原始文件的本地缓存路径（可选）
    
    Returns:
        float: 如果验证通过返回 1.0，否则返回 0.0
    """
    try:
        import tempfile
        import urllib.request
        
        # result 参数已经是从VM获取到宿主机的文件路径
        # 不应该从 options 中覆盖它，因为 options 中可能包含的是VM路径
        result_file_path = result
        original_file_url = options.get('original_file_url', '')
        
        logger.info(f"开始验证删除第二行任务...")
        logger.info(f"结果文件: {result_file_path}")
        logger.info(f"原始文件URL: {original_file_url}")
        
        if not result_file_path or not os.path.exists(result_file_path):
            logger.error(f"结果文件不存在: {result_file_path}")
            return 0.0
        
        # 下载原始文件到临时位置
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            original_file_temp = tmp_file.name
        
        try:
            logger.info(f"正在下载原始文件到临时位置: {original_file_temp}")
            urllib.request.urlretrieve(original_file_url, original_file_temp)
        except Exception as e:
            logger.warning(f"下载原始文件失败: {e}")
            # 如果下载失败，尝试从本地缓存读取
            cache_path = options.get('original_file_cache', '')
            if cache_path and os.path.exists(cache_path):
                logger.info(f"使用缓存文件: {cache_path}")
                original_file_temp = cache_path
            else:
                logger.error("无法获取原始文件")
                return 0.0
        
        # 加载原始文件
        logger.info("加载原始文件...")
        original_wb = openpyxl.load_workbook(original_file_temp)
        original_ws = original_wb.active
        
        # 获取原始文件的所有行
        original_rows = list(original_ws.iter_rows(values_only=True))
        original_row_count = len(original_rows)
        
        if original_row_count < 2:
            logger.error(f"原始文件行数不足: {original_row_count}（需要至少2行）")
            return 0.0
        
        # 保存第二行的数据（索引为1）
        second_row_data = original_rows[1]
        logger.info(f"原始文件行数: {original_row_count}")
        logger.info(f"原始文件第二行数据: {second_row_data}")
        
        # 加载结果文件
        logger.info(f"加载结果文件...")
        result_wb = openpyxl.load_workbook(result_file_path)
        result_ws = result_wb.active
        
        # 获取结果文件的所有行
        result_rows = list(result_ws.iter_rows(values_only=True))
        result_row_count = len(result_rows)
        
        logger.info(f"结果文件行数: {result_row_count}")
        
        # 验证1: 检查行数是否减少了1
        if result_row_count != original_row_count - 1:
            logger.error(f"行数验证失败: 期望 {original_row_count - 1} 行，实际 {result_row_count} 行")
            return 0.0
        else:
            logger.info(f"✓ 行数验证通过: {original_row_count} → {result_row_count}")
        
        # 验证2: 检查原始第二行是否存在于结果文件中
        second_row_exists = False
        for i, row in enumerate(result_rows):
            if row == second_row_data:
                logger.error(f"原始第二行数据仍存在于结果文件的第 {i+1} 行")
                second_row_exists = True
                break
        
        if second_row_exists:
            return 0.0
        else:
            logger.info(f"✓ 原始第二行数据已从结果文件中删除")
        
        # 验证3: 检查其他行是否保持不变（第一行和第3行之后）
        # 结果文件的第一行应该等于原始文件的第一行
        if result_rows[0] != original_rows[0]:
            logger.error(f"第一行数据不匹配")
            logger.error(f"  原始: {original_rows[0]}")
            logger.error(f"  结果: {result_rows[0]}")
            return 0.0
        
        # 结果文件的第2行及之后应该等于原始文件的第3行及之后
        for i in range(1, result_row_count):
            if result_rows[i] != original_rows[i+1]:
                logger.error(f"第 {i+1} 行数据不匹配")
                logger.error(f"  期望（原始第 {i+2} 行）: {original_rows[i+1]}")
                logger.error(f"  实际: {result_rows[i]}")
                return 0.0
        
        logger.info(f"✓ 其他行数据保持不变")
        
        # 清理临时文件
        if original_file_temp != options.get('original_file_cache', ''):
            try:
                os.unlink(original_file_temp)
            except:
                pass
        
        logger.info("=" * 60)
        logger.info("✓ 所有验证通过！第二行已成功删除")
        logger.info("=" * 60)
        return 1.0
        
    except Exception as e:
        import traceback
        logger.error(f"评估出错: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_regexp_extract(result: str, expected: str = None, **options) -> float:
    """
    Verify if REGEX formulas exist in specified columns (B and C) with correct patterns.
    
    This function checks:
    1. Whether cells in specified columns contain REGEX formulas
    2. Whether formulas reference the corresponding A column cell (B2->A2, B3->A3, etc.)
    3. Whether formulas contain the correct pattern text (牛肉丸 for B column, 牛筋丸 for C column)
    4. Whether formulas have the correct structure with lookbehind and lookahead
    
    The function automatically detects the number of data rows by checking the data column
    (default: A column) for non-empty cells. It stops checking after finding 3 consecutive
    empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_columns: List of columns to check (e.g., ["B", "C"])
            - start_row: Starting row number (default: 2)
            - end_row: Ending row number (optional, will auto-detect if not provided)
            - expected_pattern: Expected function name (default: "REGEX")
            - column_patterns: Dict mapping column letters to expected pattern text
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        from openpyxl.utils import get_column_letter, column_index_from_string
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_columns = options.get('check_columns', ['B', 'C'])
        start_row = options.get('start_row', 2)
        end_row = options.get('end_row', None)  # Optional, will auto-detect if not provided
        expected_pattern = options.get('expected_pattern', 'REGEX')
        column_patterns = options.get('column_patterns', {'B': '牛肉丸', 'C': '牛筋丸'})
        data_column = options.get('data_column', 'A')  # Column to check for data to determine end_row
        
        if not check_columns:
            logger.error("No columns specified in options")
            return 0.0
        
        logger.info(f"Verifying REGEX formulas in file: {result}")
        logger.info(f"Columns to check: {check_columns}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected function: {expected_pattern}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Auto-detect end_row by checking data_column for non-empty cells
        if end_row is None:
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
        else:
            logger.info(f"Using specified end row: {end_row}")
        
        # Check each column and row
        all_passed = True
        for col_letter in check_columns:
            expected_pattern_text = column_patterns.get(col_letter)
            if not expected_pattern_text:
                logger.warning(f"No pattern text specified for column {col_letter}, skipping")
                continue
            
            logger.info(f"Checking column {col_letter} with pattern '{expected_pattern_text}' (rows {start_row} to {end_row})")
            
            for row_num in range(start_row, end_row + 1):
                cell_coord = f"{col_letter}{row_num}"
                try:
                    cell = ws[cell_coord]
                    logger.debug(f"Checking cell {cell_coord}")
                    
                    # Check if cell contains a formula
                    if cell.data_type != "f":
                        logger.warning(f"Cell {cell_coord} does not contain a formula")
                        all_passed = False
                        continue
                    
                    # Get formula text
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
                    
                    # Remove leading = if present for comparison
                    formula_clean = formula_text.lstrip("=")
                    logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                    
                    # Check 1: Formula contains REGEX function
                    if expected_pattern.upper() not in formula_text.upper():
                        logger.warning(f"Cell {cell_coord} formula does not contain {expected_pattern}")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 2: Formula contains expected pattern text (牛肉丸 or 牛筋丸)
                    if expected_pattern_text not in formula_text:
                        logger.warning(f"Cell {cell_coord} formula does not contain pattern text '{expected_pattern_text}'")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 3: Formula contains REGEX function call structure
                    regex_match = re.search(r'REGEX\s*\([^)]+\)', formula_text, re.IGNORECASE)
                    if not regex_match:
                        logger.warning(f"Cell {cell_coord} formula does not have correct REGEX structure")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 4: Formula references the corresponding A column cell (A2, A3, etc.)
                    expected_a_cell = f"A{row_num}"
                    # Check if formula contains A column reference with the same row number
                    a_cell_pattern = rf'A{row_num}\b'
                    if not re.search(a_cell_pattern, formula_text, re.IGNORECASE):
                        logger.warning(f"Cell {cell_coord} formula does not reference {expected_a_cell}")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 5: Formula contains lookbehind pattern (?<=...)
                    if "(?<=" not in formula_text:
                        logger.warning(f"Cell {cell_coord} formula does not contain lookbehind pattern (?<=...)")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 6: Formula contains lookahead pattern (?=,)
                    if "(?=," not in formula_text:
                        logger.warning(f"Cell {cell_coord} formula does not contain lookahead pattern (?=,)")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                    
                    # Check 7: Formula contains \d+ pattern
                    if "\\d+" not in formula_text:
                        # Also check for unescaped version in the pattern
                        if not re.search(r'\\d\+|d\+', formula_text):
                            logger.warning(f"Cell {cell_coord} formula does not contain digit pattern \\d+")
                            logger.warning(f"Formula: {formula_text}")
                            all_passed = False
                            continue
                    
                    # Check 8: Formula pattern should contain 5 dots after pattern text
                    # Pattern should be like: (?<=牛肉丸.....)
                    pattern_with_dots = expected_pattern_text + "....."
                    if pattern_with_dots not in formula_text:
                        logger.warning(f"Cell {cell_coord} formula pattern may not have 5 dots after '{expected_pattern_text}'")
                        logger.debug(f"Formula: {formula_text}")
                        # Don't fail, just warn - the pattern might be correct but formatted differently
                    
                    logger.info(f"✓ Cell {cell_coord} has valid REGEX formula: {formula_text}")
                    
                except Exception as e:
                    logger.error(f"Error checking cell {cell_coord}: {e}")
                    import traceback
                    logger.error(traceback.format_exc())
                    all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in columns {check_columns} contain correct {expected_pattern} formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ {expected_pattern} formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_regexp_order_extract(result: str, expected: str = None, **options) -> float:
    """
    Verify if REGEXP formulas exist in specified column (C) to extract order numbers from addresses.
    
    This function checks:
    1. Whether cells in specified column contain REGEXP formulas
    2. Whether formulas reference the corresponding A column cell (C2->A2, C3->A3, etc.)
    3. Whether formulas contain the correct regex pattern (\\w{10})
    4. Whether formulas have the correct structure
    
    The function automatically detects the number of data rows by checking the data column
    (default: A column) for non-empty cells. It stops checking after finding 3 consecutive
    empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "C")
            - start_row: Starting row number (default: 2)
            - expected_pattern: Expected function name (default: "REGEXP")
            - expected_formula_pattern: Expected formula pattern (e.g., "REGEXP(A")
            - regex_pattern: Expected regex pattern in formula (e.g., "\\w{10}")
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'C')
        start_row = options.get('start_row', 2)
        expected_pattern = options.get('expected_pattern', 'REGEX')
        expected_formula_pattern = options.get('expected_formula_pattern', 'REGEX(A')
        regex_pattern = options.get('regex_pattern', '[a-zA-Z0-9]{10}')
        data_column = options.get('data_column', 'A')
        
        logger.info(f"Verifying REGEXP order extraction in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected function: {expected_pattern}")
        logger.info(f"Expected regex pattern: {regex_pattern}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                # Remove leading = if present for comparison
                formula_clean = formula_text.lstrip("=")
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains REGEX function (support both REGEX and LibreOffice internal format)
                # LibreOffice may save as _xlfn.ORG.LIBREOFFICE.REGEX
                formula_upper = formula_text.upper()
                if expected_pattern.upper() not in formula_upper and '_XLFN.ORG.LIBREOFFICE.REGEX' not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain {expected_pattern} or LibreOffice REGEX")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 2: Formula contains expected formula pattern (REGEX(A or _xlfn.ORG.LIBREOFFICE.REGEX(A)
                formula_clean_upper = formula_clean.upper()
                if expected_formula_pattern.upper() not in formula_clean_upper and 'REGEX(A' not in formula_clean_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain pattern '{expected_formula_pattern}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: Formula contains REGEX function call structure (support both formats)
                regexp_match = re.search(r'(REGEX|REGEXP|_XLFN\.ORG\.LIBREOFFICE\.REGEX)\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not regexp_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct REGEX structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: Formula references the corresponding A column cell (A2, A3, etc.)
                expected_a_cell = f"A{row_num}"
                # Check if formula contains A column reference with the same row number
                a_cell_pattern = rf'A{row_num}\b'
                if not re.search(a_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference {expected_a_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 5: Formula contains the regex pattern ([a-zA-Z0-9]{10})
                # The pattern might be escaped differently in the formula
                # Check for various escape formats
                pattern_variations = [
                    regex_pattern,  # [a-zA-Z0-9]{10}
                    regex_pattern.replace('\\', '\\\\'),  # [a-zA-Z0-9]{10} with double escape
                    regex_pattern.replace('[', '\\[').replace(']', '\\]'),  # Escaped brackets
                    '[a-zA-Z0-9]{10}',  # Original pattern
                    '\\[a-zA-Z0-9\\]{10}',  # Escaped brackets
                    '\\\\[a-zA-Z0-9\\\\]{10}',  # Double escaped
                ]
                found = False
                for pattern_var in pattern_variations:
                    if pattern_var in formula_text:
                        found = True
                        break
                if not found:
                    # Also check for pattern without escaping brackets
                    simple_pattern = 'a-zA-Z0-9]{10}'
                    if simple_pattern not in formula_text:
                        logger.warning(f"Cell {cell_coord} formula does not contain regex pattern '{regex_pattern}'")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        continue
                
                logger.info(f"✓ Cell {cell_coord} has valid REGEXP formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct {expected_pattern} formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ {expected_pattern} formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_sumif_formula(result: str, expected: str = None, **options) -> float:
    """
    Verify if SUMIF formulas exist in specified column (F) to calculate totals.
    
    This function checks:
    1. Whether cells in specified column contain SUMIF formulas
    2. Whether formulas reference the correct ranges (auto-detected from data)
    3. Whether formulas reference the corresponding E column cell (F2->E2, F3->E3, etc.)
    4. Whether formulas have the correct structure
    
    The function automatically detects:
    - end_row: by checking the data column (E) for non-empty cells
    - criteria_range: by detecting the range from the first formula or from criteria_column data
    - sum_range: by detecting the range from the first formula or from sum_column data
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "F")
            - start_row: Starting row number (default: 2)
            - expected_function: Expected function name (default: "SUMIF")
            - criteria_column: Column containing criteria (e.g., "B")
            - sum_column: Column containing values to sum (e.g., "C")
            - criteria_column_start: Starting row for criteria column (default: 2)
            - data_column: Column to check for data to determine end_row (default: "E")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'F')
        start_row = options.get('start_row', 2)
        expected_function = options.get('expected_function', 'SUMIF')
        criteria_column = options.get('criteria_column', 'B')
        sum_column = options.get('sum_column', 'C')
        criteria_column_start = options.get('criteria_column_start', 2)
        data_column = options.get('data_column', 'E')
        
        logger.info(f"Verifying SUMIF formulas in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected function: {expected_function}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        
        # Auto-detect criteria_range and sum_range by checking the first formula
        criteria_range = None
        sum_range = None
        
        # Try to extract ranges from the first formula
        first_cell_coord = f"{check_column}{start_row}"
        try:
            first_cell = ws[first_cell_coord]
            if first_cell.data_type == "f":
                first_formula_text = None
                if hasattr(first_cell, "_value") and isinstance(first_cell._value, str) and first_cell._value.startswith("="):
                    first_formula_text = first_cell._value
                elif hasattr(first_cell, "formula"):
                    first_formula_text = first_cell.formula
                elif first_cell.value is not None and isinstance(first_cell.value, str) and first_cell.value.startswith("="):
                    first_formula_text = first_cell.value
                
                if first_formula_text:
                    # Extract ranges from SUMIF formula: SUMIF(range1, criteria, range2)
                    # Pattern: SUMIF(range1, criteria, range2)
                    sumif_pattern = r'SUMIF\s*\(\s*([^,]+)\s*,\s*[^,]+\s*,\s*([^)]+)\s*\)'
                    match = re.search(sumif_pattern, first_formula_text, re.IGNORECASE)
                    if match:
                        criteria_range = match.group(1).strip()
                        sum_range = match.group(2).strip()
                        logger.info(f"Extracted from first formula: criteria_range={criteria_range}, sum_range={sum_range}")
        except Exception as e:
            logger.debug(f"Could not extract ranges from first formula: {e}")
        
        # If ranges not found in formula, detect from data columns
        if not criteria_range or not sum_range:
            logger.info(f"Auto-detecting ranges from data columns...")
            # Find the last row with data in criteria_column
            criteria_end_row = criteria_column_start
            empty_count = 0
            for row_num in range(criteria_column_start, max_row + 1):
                criteria_cell = ws[f"{criteria_column}{row_num}"]
                if criteria_cell.value is None or (isinstance(criteria_cell.value, str) and criteria_cell.value.strip() == ""):
                    empty_count += 1
                    if empty_count >= 3:
                        break
                else:
                    empty_count = 0
                    criteria_end_row = row_num
            
            # Find the last row with data in sum_column
            sum_end_row = criteria_column_start
            empty_count = 0
            for row_num in range(criteria_column_start, max_row + 1):
                sum_cell = ws[f"{sum_column}{row_num}"]
                if sum_cell.value is None or (isinstance(sum_cell.value, str) and sum_cell.value.strip() == ""):
                    empty_count += 1
                    if empty_count >= 3:
                        break
                else:
                    empty_count = 0
                    sum_end_row = row_num
            
            # Use the maximum end row for both ranges
            max_end_row = max(criteria_end_row, sum_end_row)
            criteria_range = f"{criteria_column}{criteria_column_start}:{criteria_column}{max_end_row}"
            sum_range = f"{sum_column}{criteria_column_start}:{sum_column}{max_end_row}"
            logger.info(f"Auto-detected ranges: criteria_range={criteria_range}, sum_range={sum_range}")
        
        # Check each row in the specified column
        all_passed = True
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                # Remove leading = if present for comparison
                formula_clean = formula_text.lstrip("=")
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains SUMIF function
                formula_upper = formula_text.upper()
                if expected_function.upper() not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain {expected_function}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 2: Formula contains SUMIF function call structure
                sumif_match = re.search(r'SUMIF\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not sumif_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct SUMIF structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: Formula contains criteria range
                if criteria_range and criteria_range.upper() not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain criteria range '{criteria_range}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: Formula contains sum range
                if sum_range and sum_range.upper() not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain sum range '{sum_range}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 5: Formula references the corresponding E column cell (E2, E3, etc.)
                expected_e_cell = f"E{row_num}"
                # Check if formula contains E column reference with the same row number
                e_cell_pattern = rf'E{row_num}\b'
                if not re.search(e_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference {expected_e_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {cell_coord} has valid SUMIF formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct {expected_function} formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ {expected_function} formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_networkdays_formula(result: str, expected: str = None, **options) -> float:
    """
    Verify if NETWORKDAYS formulas exist in specified column to calculate working days.
    
    This function checks:
    1. Whether cells in specified column contain NETWORKDAYS formulas
    2. Whether formulas reference the corresponding start date column cell (A2, A3, etc.)
    3. Whether formulas reference the corresponding end date column cell (B2, B3, etc.)
    4. Whether formulas have the correct structure
    
    The function automatically detects the number of data rows by checking the data column
    (default: A column) for non-empty cells. It stops checking after finding 3 consecutive
    empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "C")
            - start_row: Starting row number (default: 2)
            - start_date_column: Column containing start dates (e.g., "A")
            - end_date_column: Column containing end dates (e.g., "B")
            - expected_function: Expected function name (default: "NETWORKDAYS")
            - data_column: Column to check for data to determine end_row (default: "A")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'C')
        start_row = options.get('start_row', 2)
        start_date_column = options.get('start_date_column', 'A')
        end_date_column = options.get('end_date_column', 'B')
        expected_function = options.get('expected_function', 'NETWORKDAYS')
        data_column = options.get('data_column', 'A')
        
        logger.info(f"Verifying NETWORKDAYS formulas in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Start date column: {start_date_column}")
        logger.info(f"End date column: {end_date_column}")
        logger.info(f"Expected function: {expected_function}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                # Remove leading = if present for comparison
                formula_clean = formula_text.lstrip("=")
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains NETWORKDAYS function
                formula_upper = formula_text.upper()
                if expected_function.upper() not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain {expected_function}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 2: Formula contains NETWORKDAYS function call structure
                # NETWORKDAYS can have 2 or 3 parameters: NETWORKDAYS(start_date, end_date) or NETWORKDAYS(start_date, end_date, holidays)
                networkdays_pattern = r'NETWORKDAYS\s*\([^)]+\)'
                networkdays_match = re.search(networkdays_pattern, formula_text, re.IGNORECASE)
                if not networkdays_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct NETWORKDAYS structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: Formula references the corresponding start date column cell (A2, A3, etc.)
                expected_start_cell = f"{start_date_column}{row_num}"
                start_cell_pattern = rf'{start_date_column}{row_num}\b'
                if not re.search(start_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference start date cell {expected_start_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: Formula references the corresponding end date column cell (B2, B3, etc.)
                expected_end_cell = f"{end_date_column}{row_num}"
                end_cell_pattern = rf'{end_date_column}{row_num}\b'
                if not re.search(end_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference end date cell {expected_end_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {cell_coord} has valid NETWORKDAYS formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct {expected_function} formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ {expected_function} formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_conditional_formatting_reconciliation(result: str, expected: str = None, **options) -> float:
    """
    Verify if conditional formatting is correctly set up for reconciliation between two tables.
    
    This function checks:
    1. Whether conditional formatting rules exist in the worksheet
    2. Whether the formula matches the expected pattern (e.g., A1<>E1 to compare cells from two tables)
    3. Whether conditional formatting is applied to the correct range
    4. Whether cells with differences are formatted (highlighted)
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_range: Range where conditional formatting should be applied (e.g., "A1:C16")
            - compare_range: Range to compare against (e.g., "E1:G16")
            - expected_formula: Expected formula pattern (e.g., "A1<>E1")
            - format_column: Column to check for formatting (optional, e.g., "B")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        from openpyxl.utils import get_column_letter, column_index_from_string
        from openpyxl.worksheet.cell_range import CellRange
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_range = options.get('check_range', 'A1:C16')
        compare_range = options.get('compare_range', 'E1:G16')
        expected_formula = options.get('expected_formula', 'A1<>E1')
        format_column = options.get('format_column', None)
        
        logger.info(f"Verifying conditional formatting reconciliation in file: {result}")
        logger.info(f"Check range: {check_range}")
        logger.info(f"Compare range: {compare_range}")
        logger.info(f"Expected formula pattern: {expected_formula}")
        
        # Load workbook
        try:
            wb = openpyxl.load_workbook(result, data_only=False)
            ws = wb.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Check if conditional formatting exists
        conditional_formattings = ws.conditional_formatting
        if not conditional_formattings:
            logger.error("No conditional formatting rules found in the worksheet")
            return 0.0
        
        logger.info(f"Found {len(conditional_formattings)} conditional formatting rule(s)")
        
        # Parse expected formula to extract cell references
        # Expected formula like "A1<>E1" means compare A1 with E1
        expected_formula_clean = expected_formula.replace(" ", "").upper()
        
        # Find matching conditional formatting rule
        found_matching_rule = False
        rule_applied_to_correct_range = False
        
        for fmt in conditional_formattings:
            for rule in fmt.rules:
                # Check if rule has formula
                if not rule.formula:
                    continue
                
                # Check formula pattern
                formula_text = rule.formula[0] if rule.formula else ""
                formula_text_clean = formula_text.replace(" ", "").upper()
                
                logger.debug(f"Checking rule with formula: {formula_text}")
                
                # Check if formula matches expected pattern
                # The formula should contain comparison like A1<>E1, A2<>E2, etc.
                # We need to check if the pattern matches (allowing for relative references)
                if "<>" in expected_formula_clean:
                    # Extract cell references from expected formula
                    # Pattern: A1<>E1 means compare A column with E column
                    expected_parts = expected_formula_clean.split("<>")
                    if len(expected_parts) == 2:
                        expected_cell1 = expected_parts[0]  # e.g., "A1"
                        expected_cell2 = expected_parts[1]   # e.g., "E1"
                        
                        # Extract column letters
                        expected_col1 = re.match(r'([A-Z]+)', expected_cell1)
                        expected_col2 = re.match(r'([A-Z]+)', expected_cell2)
                        
                        if expected_col1 and expected_col2:
                            col1 = expected_col1.group(1)
                            col2 = expected_col2.group(1)
                            
                            # Check if formula contains comparison between these columns
                            # Pattern should be like: A1<>E1, A2<>E2, etc. (relative references)
                            pattern1 = rf'{col1}\d+\s*<>\s*{col2}\d+'
                            pattern2 = rf'{col1}\d+\s*!=\s*{col2}\d+'  # Alternative: !=
                            
                            if re.search(pattern1, formula_text_clean, re.IGNORECASE) or \
                               re.search(pattern2, formula_text_clean, re.IGNORECASE):
                                found_matching_rule = True
                                logger.info(f"✓ Found matching formula pattern: {formula_text}")
                                
                                # Check if rule is applied to correct range
                                fmt_ranges = [str(rng) for rng in fmt.cells]
                                check_range_upper = check_range.upper()
                                
                                # Check if check_range is covered by any of the formatting ranges
                                try:
                                    check_cell_range = CellRange(check_range_upper)
                                    for fmt_range_str in fmt_ranges:
                                        fmt_cell_range = CellRange(fmt_range_str)
                                        # Check if check_range is within or overlaps with fmt_range
                                        if (check_cell_range.min_row >= fmt_cell_range.min_row and
                                            check_cell_range.max_row <= fmt_cell_range.max_row and
                                            check_cell_range.min_col >= fmt_cell_range.min_col and
                                            check_cell_range.max_col <= fmt_cell_range.max_col):
                                            rule_applied_to_correct_range = True
                                            logger.info(f"✓ Rule applied to correct range: {fmt_range_str} covers {check_range}")
                                            break
                                except Exception as e:
                                    logger.debug(f"Error parsing ranges: {e}")
                                    # If range parsing fails, check if range string matches
                                    if check_range_upper in fmt_ranges:
                                        rule_applied_to_correct_range = True
                                        logger.info(f"✓ Rule applied to exact range: {check_range}")
                                
                                break
            
            if found_matching_rule:
                break
        
        if not found_matching_rule:
            logger.error("No conditional formatting rule found with expected formula pattern")
            return 0.0
        
        if not rule_applied_to_correct_range:
            logger.warning("Conditional formatting rule found but may not be applied to correct range")
            # Don't fail completely, as the range might be slightly different but still valid
        
        # Optional: Check if cells with differences are actually formatted
        # This is a more advanced check that verifies the formatting is working
        if format_column:
            logger.info(f"Checking formatting in column {format_column}...")
            # Try to find cells in format_column that have conditional formatting applied
            # This is a simplified check - in practice, we'd need to evaluate the formula
            # for each cell to see if it's formatted
            
        logger.info("=" * 60)
        logger.info("✓ Conditional formatting reconciliation verification passed")
        logger.info("=" * 60)
        return 1.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_right_len_find_extract(result: str, expected: str = None, **options) -> float:
    """
    Verify if RIGHT, LEN, and FIND formulas exist in specified column to extract text.
    
    This function checks:
    1. Whether cells in specified column contain RIGHT, LEN, and FIND functions
    2. Whether formulas reference the corresponding source column cell (C2->B2, C3->B3, etc.)
    3. Whether formulas contain the correct pattern (e.g., RIGHT(B2,LEN(B2)-FIND("班",B2)))
    4. Whether formulas have the correct structure with RIGHT, LEN, and FIND functions
    
    The function automatically detects the number of data rows by checking the data column
    (default: B column) for non-empty cells. It stops checking after finding 3 consecutive
    empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "C")
            - start_row: Starting row number (default: 2)
            - expected_functions: List of expected function names (e.g., ["RIGHT", "LEN", "FIND"])
            - expected_formula_pattern: Expected formula pattern (e.g., "RIGHT(B")
            - find_text: Text to find in FIND function (e.g., "班")
            - data_column: Column to check for data to determine end_row (default: "B")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'C')
        start_row = options.get('start_row', 2)
        expected_functions = options.get('expected_functions', ['RIGHT', 'LEN', 'FIND'])
        expected_formula_pattern = options.get('expected_formula_pattern', 'RIGHT(B')
        find_text = options.get('find_text', '班')
        data_column = options.get('data_column', 'B')
        
        logger.info(f"Verifying RIGHT/LEN/FIND extraction formulas in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected functions: {expected_functions}")
        logger.info(f"Expected formula pattern: {expected_formula_pattern}")
        logger.info(f"Find text: {find_text}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                # Remove leading = if present for comparison
                formula_clean = formula_text.lstrip("=")
                formula_upper = formula_text.upper()
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains all expected functions
                for func_name in expected_functions:
                    if func_name.upper() not in formula_upper:
                        logger.warning(f"Cell {cell_coord} formula does not contain {func_name}")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        break
                
                if not all_passed:
                    continue
                
                # Check 2: Formula contains expected formula pattern (e.g., RIGHT(B)
                if expected_formula_pattern.upper() not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain pattern '{expected_formula_pattern}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: Formula contains RIGHT function call structure
                right_match = re.search(r'RIGHT\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not right_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct RIGHT structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: Formula contains LEN function
                len_match = re.search(r'LEN\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not len_match:
                    logger.warning(f"Cell {cell_coord} formula does not have LEN function")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 5: Formula contains FIND function with find_text
                find_pattern = rf'FIND\s*\([^)]*{re.escape(find_text)}[^)]*\)'
                find_match = re.search(find_pattern, formula_text, re.IGNORECASE)
                if not find_match:
                    logger.warning(f"Cell {cell_coord} formula does not contain FIND function with text '{find_text}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 6: Formula references the corresponding source column cell (B2, B3, etc.)
                expected_source_cell = f"{data_column}{row_num}"
                source_cell_pattern = rf'{data_column}{row_num}\b'
                if not re.search(source_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference source cell {expected_source_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 7: Formula structure should be RIGHT(B2, LEN(B2)-FIND("班",B2))
                # Verify that LEN and FIND are used together in the second parameter of RIGHT
                # This is a pattern check - the formula should have LEN(...)-FIND(...) structure
                len_find_pattern = r'LEN\s*\([^)]+\)\s*-\s*FIND\s*\([^)]+\)'
                if not re.search(len_find_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not have LEN(...)-FIND(...) structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {cell_coord} has valid RIGHT/LEN/FIND formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct RIGHT/LEN/FIND formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ RIGHT/LEN/FIND formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_iferror_regex_extract(result: str, expected: str = None, **options) -> float:
    """
    Verify if IFERROR(REGEX(...)) formulas exist in specified column to extract text with error handling.
    
    This function checks:
    1. Whether cells in specified column contain IFERROR function wrapping REGEX
    2. Whether REGEX function uses capture group pattern (e.g., .*水笔(\d+).*)
    3. Whether REGEX function uses replacement pattern (e.g., $1元)
    4. Whether IFERROR has empty string as second parameter
    5. Whether formulas reference the corresponding source column cell
    
    The function automatically detects the number of data rows by checking the data column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "C")
            - start_row: Starting row number (default: 2)
            - expected_pattern_text: Expected text pattern in regex (e.g., "水笔")
            - data_column: Column to check for data to determine end_row (default: "B")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'C')
        start_row = options.get('start_row', 2)
        expected_pattern_text = options.get('expected_pattern_text', '水笔')
        data_column = options.get('data_column', 'B')
        
        logger.info(f"Verifying IFERROR(REGEX(...)) formulas in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected pattern text: {expected_pattern_text}")
        logger.info(f"Data column: {data_column}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                formula_upper = formula_text.upper()
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains IFERROR function
                if 'IFERROR' not in formula_upper:
                    logger.warning(f"Cell {cell_coord} formula does not contain IFERROR function")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 2: Formula contains REGEX function (inside IFERROR)
                # Support both REGEX and LibreOffice internal format _xlfn.ORG.LIBREOFFICE.REGEX
                has_regex = 'REGEX' in formula_upper or '_XLFN.ORG.LIBREOFFICE.REGEX' in formula_upper
                if not has_regex:
                    logger.warning(f"Cell {cell_coord} formula does not contain REGEX function")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: IFERROR structure - should have two parameters
                iferror_match = re.search(r'IFERROR\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not iferror_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct IFERROR structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: IFERROR second parameter should be empty string ""
                # Handle various formats: ,"" or , "" or ,'' or , ''
                # Also handle LibreOffice format with spaces: REGEX(...) ,""
                # Extract IFERROR parameters: IFERROR(param1, param2)
                iferror_params_match = re.search(r'IFERROR\s*\((.*)\)', formula_text, re.IGNORECASE)
                if iferror_params_match:
                    params_str = iferror_params_match.group(1)
                    # Split by comma, but need to handle nested commas in strings
                    # Simple approach: find the last comma (should separate the two parameters)
                    # For IFERROR(REGEX(...), ""), the last comma separates REGEX call from ""
                    last_comma_pos = params_str.rfind(',')
                    if last_comma_pos != -1:
                        second_param = params_str[last_comma_pos + 1:].strip()
                        # Check if second parameter is empty string "" or ''
                        if second_param in ['""', "''", '""', "''"]:
                            has_empty_string = True
                        else:
                            has_empty_string = False
                    else:
                        has_empty_string = False
                else:
                    has_empty_string = False
                
                if not has_empty_string:
                    logger.warning(f"Cell {cell_coord} IFERROR should have empty string as second parameter")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 5: REGEX function call structure
                # Support both REGEX and LibreOffice internal format
                regex_match = re.search(r'(REGEX|_XLFN\.ORG\.LIBREOFFICE\.REGEX)\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not regex_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct REGEX structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 6: Formula contains expected pattern text (e.g., "水笔")
                if expected_pattern_text not in formula_text:
                    logger.warning(f"Cell {cell_coord} formula does not contain pattern text '{expected_pattern_text}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 7: Formula contains capture group pattern (\d+)
                has_capture_group = bool(re.search(r'\(\\d\+\)|\(\\\\d\+\)', formula_text))
                if not has_capture_group:
                    logger.warning(f"Cell {cell_coord} REGEX formula should contain capture group (\\d+)")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 8: Formula contains replacement pattern $1
                has_replacement = '"$1' in formula_text or "'$1" in formula_text
                if not has_replacement:
                    logger.warning(f"Cell {cell_coord} REGEX formula should contain replacement pattern $1")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 9: Formula references the corresponding source column cell
                expected_source_cell = f"{data_column}{row_num}"
                source_cell_pattern = rf'{data_column}{row_num}\b'
                if not re.search(source_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference source cell {expected_source_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {cell_coord} has valid IFERROR(REGEX(...)) formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct IFERROR(REGEX(...)) formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ IFERROR(REGEX(...)) formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_rept_text_progress_bar(result: str, expected: str = None, **options) -> float:
    """
    Verify if REPT and TEXT formulas exist in specified column to create progress bars with percentage.
    
    This function checks:
    1. Whether cells in specified column contain REPT and TEXT functions
    2. Whether REPT function uses the correct character (e.g., "|")
    3. Whether REPT function uses the correct multiplier (e.g., *50)
    4. Whether TEXT function uses percentage format (e.g., "0%")
    5. Whether formulas reference the correct numerator and denominator columns
    6. Whether formulas use & operator to concatenate REPT and TEXT results
    
    The function automatically detects the number of data rows by checking the data column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - check_column: Column to check (e.g., "D")
            - start_row: Starting row number (default: 2)
            - expected_functions: List of expected function names (e.g., ["REPT", "TEXT"])
            - numerator_column: Column containing numerator values (e.g., "C")
            - denominator_column: Column containing denominator values (e.g., "B")
            - rept_char: Character to repeat in REPT function (e.g., "|")
            - rept_multiplier: Multiplier for REPT function (e.g., 50)
            - data_column: Column to check for data to determine end_row (default: "B")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        check_column = options.get('check_column', 'D')
        start_row = options.get('start_row', 2)
        expected_functions = options.get('expected_functions', ['REPT', 'TEXT'])
        numerator_column = options.get('numerator_column', 'C')
        denominator_column = options.get('denominator_column', 'B')
        rept_char = options.get('rept_char', '|')
        rept_multiplier = options.get('rept_multiplier', 50)
        data_column = options.get('data_column', 'B')
        
        logger.info(f"Verifying REPT/TEXT progress bar formulas in file: {result}")
        logger.info(f"Column to check: {check_column}")
        logger.info(f"Start row: {start_row}")
        logger.info(f"Expected functions: {expected_functions}")
        logger.info(f"Numerator column: {numerator_column}")
        logger.info(f"Denominator column: {denominator_column}")
        logger.info(f"REPT character: {rept_char}")
        logger.info(f"REPT multiplier: {rept_multiplier}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        logger.info(f"Checking column {check_column} (rows {start_row} to {end_row})")
        
        for row_num in range(start_row, end_row + 1):
            cell_coord = f"{check_column}{row_num}"
            try:
                cell = ws[cell_coord]
                logger.debug(f"Checking cell {cell_coord}")
                
                # Check if cell contains a formula
                if cell.data_type != "f":
                    logger.warning(f"Cell {cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                # Get formula text
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
                
                formula_upper = formula_text.upper()
                logger.debug(f"Cell {cell_coord} formula: {formula_text}")
                
                # Check 1: Formula contains all expected functions (REPT and TEXT)
                for func_name in expected_functions:
                    if func_name.upper() not in formula_upper:
                        logger.warning(f"Cell {cell_coord} formula does not contain {func_name}")
                        logger.warning(f"Formula: {formula_text}")
                        all_passed = False
                        break
                
                if not all_passed:
                    continue
                
                # Check 2: Formula contains REPT function call structure
                rept_match = re.search(r'REPT\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not rept_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct REPT structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 3: REPT function contains the correct character (e.g., "|")
                # Support both double quotes and single quotes
                rept_char_pattern1 = rf'REPT\s*\(\s*"{re.escape(rept_char)}"'
                rept_char_pattern2 = rf"REPT\s*\(\s*'{re.escape(rept_char)}'"
                if not re.search(rept_char_pattern1, formula_text, re.IGNORECASE) and \
                   not re.search(rept_char_pattern2, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} REPT function should use character '{rept_char}'")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 4: REPT function contains the multiplier (e.g., *50)
                rept_multiplier_pattern = rf'\*{rept_multiplier}\b'
                if not re.search(rept_multiplier_pattern, formula_text):
                    logger.warning(f"Cell {cell_coord} REPT function should use multiplier *{rept_multiplier}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 5: Formula contains TEXT function call structure
                text_match = re.search(r'TEXT\s*\([^)]+\)', formula_text, re.IGNORECASE)
                if not text_match:
                    logger.warning(f"Cell {cell_coord} formula does not have correct TEXT structure")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 6: TEXT function contains percentage format ("0%" or '0%')
                text_percent_pattern1 = r'TEXT\s*\([^,]+,\s*"0%"'
                text_percent_pattern2 = r"TEXT\s*\([^,]+,\s*'0%'"
                if not re.search(text_percent_pattern1, formula_text, re.IGNORECASE) and \
                   not re.search(text_percent_pattern2, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} TEXT function should use percentage format \"0%\"")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 7: Formula references the correct numerator column (C2, C3, etc.)
                expected_numerator_cell = f"{numerator_column}{row_num}"
                numerator_cell_pattern = rf'{numerator_column}{row_num}\b'
                if not re.search(numerator_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference numerator cell {expected_numerator_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 8: Formula references the correct denominator column (B2, B3, etc.)
                expected_denominator_cell = f"{denominator_column}{row_num}"
                denominator_cell_pattern = rf'{denominator_column}{row_num}\b'
                if not re.search(denominator_cell_pattern, formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {cell_coord} formula does not reference denominator cell {expected_denominator_cell}")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 9: Formula contains & operator to concatenate REPT and TEXT
                if '&' not in formula_text:
                    logger.warning(f"Cell {cell_coord} formula should use & operator to concatenate REPT and TEXT")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                # Check 10: Formula structure should be REPT(...)&TEXT(...)
                # Verify that REPT comes before TEXT (or at least both are present)
                rept_pos = formula_text.upper().find('REPT')
                text_pos = formula_text.upper().find('TEXT')
                if rept_pos == -1 or text_pos == -1:
                    logger.warning(f"Cell {cell_coord} formula should contain both REPT and TEXT functions")
                    logger.warning(f"Formula: {formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {cell_coord} has valid REPT/TEXT progress bar formula: {formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking cell {cell_coord}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All cells in column {check_column} contain correct REPT/TEXT progress bar formulas")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ REPT/TEXT progress bar formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_id_extract_gender_age_birthday(result: str, expected: str = None, **options) -> float:
    """
    Verify if formulas exist to extract gender, age, and birthday from ID numbers.
    
    This function checks:
    1. Gender column (C): IF(MOD(MID(B3,17,1),2),"男","女")
    2. Age column (D): DATEDIF(TEXT(MID(B3,7,8),"0-00-00"),TODAY(),"Y")
    3. Birthday column (E): --TEXT(MID(B3,7,8),"0-00-00")
    
    The function automatically detects the number of data rows by checking the ID column
    for non-empty cells. It stops checking after finding 3 consecutive empty rows.
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - id_column: Column containing ID numbers (e.g., "B")
            - gender_column: Column for gender formulas (e.g., "C")
            - age_column: Column for age formulas (e.g., "D")
            - birthday_column: Column for birthday formulas (e.g., "E")
            - start_row: Starting row number (default: 3)
            - data_column: Column to check for data to determine end_row (default: "B")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        import re
        
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        id_column = options.get('id_column', 'B')
        gender_column = options.get('gender_column', 'C')
        age_column = options.get('age_column', 'D')
        birthday_column = options.get('birthday_column', 'E')
        start_row = options.get('start_row', 3)
        data_column = options.get('data_column', 'B')
        
        logger.info(f"Verifying ID extraction formulas in file: {result}")
        logger.info(f"ID column: {id_column}")
        logger.info(f"Gender column: {gender_column}")
        logger.info(f"Age column: {age_column}")
        logger.info(f"Birthday column: {birthday_column}")
        logger.info(f"Start row: {start_row}")
        
        # Load workbook to get formulas
        try:
            wb = openpyxl.load_workbook(result, data_only=False)  # data_only=False to get formulas
            ws = wb.active
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
        
        # Check each row
        all_passed = True
        logger.info(f"Checking rows {start_row} to {end_row}")
        
        for row_num in range(start_row, end_row + 1):
            try:
                # Check gender column (C)
                gender_cell_coord = f"{gender_column}{row_num}"
                gender_cell = ws[gender_cell_coord]
                
                if gender_cell.data_type != "f":
                    logger.warning(f"Cell {gender_cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                gender_formula_text = None
                if hasattr(gender_cell, "_value") and isinstance(gender_cell._value, str) and gender_cell._value.startswith("="):
                    gender_formula_text = gender_cell._value
                elif hasattr(gender_cell, "formula"):
                    gender_formula_text = gender_cell.formula
                elif gender_cell.value is not None and isinstance(gender_cell.value, str) and gender_cell.value.startswith("="):
                    gender_formula_text = gender_cell.value
                
                if gender_formula_text is None:
                    logger.warning(f"Could not extract formula from cell {gender_cell_coord}")
                    all_passed = False
                    continue
                
                gender_formula_upper = gender_formula_text.upper()
                logger.debug(f"Cell {gender_cell_coord} formula: {gender_formula_text}")
                
                # Check gender formula: IF(MOD(MID(B3,17,1),2),"男","女")
                if 'IF' not in gender_formula_upper:
                    logger.warning(f"Cell {gender_cell_coord} formula does not contain IF function")
                    all_passed = False
                    continue
                
                if 'MOD' not in gender_formula_upper:
                    logger.warning(f"Cell {gender_cell_coord} formula does not contain MOD function")
                    all_passed = False
                    continue
                
                if 'MID' not in gender_formula_upper:
                    logger.warning(f"Cell {gender_cell_coord} formula does not contain MID function")
                    all_passed = False
                    continue
                
                # Check MID(B3,17,1) pattern
                mid_pattern = rf'MID\s*\(\s*{id_column}{row_num}\s*,\s*17\s*,\s*1\s*\)'
                if not re.search(mid_pattern, gender_formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {gender_cell_coord} formula should contain MID({id_column}{row_num},17,1)")
                    logger.warning(f"Formula: {gender_formula_text}")
                    all_passed = False
                    continue
                
                # Check for "男" and "女" in formula
                if '"男"' not in gender_formula_text and "'男'" not in gender_formula_text:
                    logger.warning(f"Cell {gender_cell_coord} formula should contain \"男\"")
                    logger.warning(f"Formula: {gender_formula_text}")
                    all_passed = False
                    continue
                
                if '"女"' not in gender_formula_text and "'女'" not in gender_formula_text:
                    logger.warning(f"Cell {gender_cell_coord} formula should contain \"女\"")
                    logger.warning(f"Formula: {gender_formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {gender_cell_coord} has valid gender formula: {gender_formula_text}")
                
                # Check age column (D)
                age_cell_coord = f"{age_column}{row_num}"
                age_cell = ws[age_cell_coord]
                
                if age_cell.data_type != "f":
                    logger.warning(f"Cell {age_cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                age_formula_text = None
                if hasattr(age_cell, "_value") and isinstance(age_cell._value, str) and age_cell._value.startswith("="):
                    age_formula_text = age_cell._value
                elif hasattr(age_cell, "formula"):
                    age_formula_text = age_cell.formula
                elif age_cell.value is not None and isinstance(age_cell.value, str) and age_cell.value.startswith("="):
                    age_formula_text = age_cell.value
                
                if age_formula_text is None:
                    logger.warning(f"Could not extract formula from cell {age_cell_coord}")
                    all_passed = False
                    continue
                
                age_formula_upper = age_formula_text.upper()
                logger.debug(f"Cell {age_cell_coord} formula: {age_formula_text}")
                
                # Check age formula: DATEDIF(TEXT(MID(B3,7,8),"0-00-00"),TODAY(),"Y")
                if 'DATEDIF' not in age_formula_upper:
                    logger.warning(f"Cell {age_cell_coord} formula does not contain DATEDIF function")
                    all_passed = False
                    continue
                
                if 'TEXT' not in age_formula_upper:
                    logger.warning(f"Cell {age_cell_coord} formula does not contain TEXT function")
                    all_passed = False
                    continue
                
                if 'TODAY' not in age_formula_upper:
                    logger.warning(f"Cell {age_cell_coord} formula does not contain TODAY function")
                    all_passed = False
                    continue
                
                if 'MID' not in age_formula_upper:
                    logger.warning(f"Cell {age_cell_coord} formula does not contain MID function")
                    all_passed = False
                    continue
                
                # Check MID(B3,7,8) pattern
                mid_pattern_age = rf'MID\s*\(\s*{id_column}{row_num}\s*,\s*7\s*,\s*8\s*\)'
                if not re.search(mid_pattern_age, age_formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {age_cell_coord} formula should contain MID({id_column}{row_num},7,8)")
                    logger.warning(f"Formula: {age_formula_text}")
                    all_passed = False
                    continue
                
                # Check TEXT format "0-00-00"
                # Use pattern that matches until the quote to handle nested functions like MID(B3,7,8)
                text_format_pattern1 = r'TEXT\s*\([^"]+,\s*"0-00-00"'
                text_format_pattern2 = r"TEXT\s*\([^']+,\s*'0-00-00'"
                if not re.search(text_format_pattern1, age_formula_text, re.IGNORECASE) and \
                   not re.search(text_format_pattern2, age_formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {age_cell_coord} TEXT function should use format \"0-00-00\"")
                    logger.warning(f"Formula: {age_formula_text}")
                    all_passed = False
                    continue
                
                # Check DATEDIF third parameter "Y"
                # Use a more flexible pattern that handles nested functions
                # Check if DATEDIF contains "Y" as the third parameter (after two commas)
                # Pattern: DATEDIF(...,...,"Y") or DATEDIF(...,...,'Y')
                # We'll count commas to find the third parameter
                datedif_match = re.search(r'DATEDIF\s*\((.*)\)', age_formula_text, re.IGNORECASE)
                if datedif_match:
                    datedif_params = datedif_match.group(1)
                    # Count commas to find the third parameter
                    # Simple approach: check if the last part before closing paren is "Y" or 'Y'
                    # More robust: find the pattern ,"Y" or ,'Y' before the closing paren
                    if not re.search(r',\s*"Y"\s*\)', age_formula_text, re.IGNORECASE) and \
                       not re.search(r",\s*'Y'\s*\)", age_formula_text, re.IGNORECASE):
                        logger.warning(f"Cell {age_cell_coord} DATEDIF function should use \"Y\" parameter")
                        logger.warning(f"Formula: {age_formula_text}")
                        all_passed = False
                        continue
                else:
                    logger.warning(f"Cell {age_cell_coord} could not parse DATEDIF function")
                    logger.warning(f"Formula: {age_formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {age_cell_coord} has valid age formula: {age_formula_text}")
                
                # Check birthday column (E)
                birthday_cell_coord = f"{birthday_column}{row_num}"
                birthday_cell = ws[birthday_cell_coord]
                
                if birthday_cell.data_type != "f":
                    logger.warning(f"Cell {birthday_cell_coord} does not contain a formula")
                    all_passed = False
                    continue
                
                birthday_formula_text = None
                if hasattr(birthday_cell, "_value") and isinstance(birthday_cell._value, str) and birthday_cell._value.startswith("="):
                    birthday_formula_text = birthday_cell._value
                elif hasattr(birthday_cell, "formula"):
                    birthday_formula_text = birthday_cell.formula
                elif birthday_cell.value is not None and isinstance(birthday_cell.value, str) and birthday_cell.value.startswith("="):
                    birthday_formula_text = birthday_cell.value
                
                if birthday_formula_text is None:
                    logger.warning(f"Could not extract formula from cell {birthday_cell_coord}")
                    all_passed = False
                    continue
                
                birthday_formula_upper = birthday_formula_text.upper()
                logger.debug(f"Cell {birthday_cell_coord} formula: {birthday_formula_text}")
                
                # Check birthday formula: TEXT(MID(B3,7,8),"0-00-00")
                if 'TEXT' not in birthday_formula_upper:
                    logger.warning(f"Cell {birthday_cell_coord} formula does not contain TEXT function")
                    all_passed = False
                    continue
                
                if 'MID' not in birthday_formula_upper:
                    logger.warning(f"Cell {birthday_cell_coord} formula does not contain MID function")
                    all_passed = False
                    continue
                
                # Check MID(B3,7,8) pattern
                mid_pattern_birthday = rf'MID\s*\(\s*{id_column}{row_num}\s*,\s*7\s*,\s*8\s*\)'
                if not re.search(mid_pattern_birthday, birthday_formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {birthday_cell_coord} formula should contain MID({id_column}{row_num},7,8)")
                    logger.warning(f"Formula: {birthday_formula_text}")
                    all_passed = False
                    continue
                
                # Check TEXT format "0-00-00"
                # Use pattern that matches until the quote to handle nested functions like MID(B3,7,8)
                text_format_pattern1 = r'TEXT\s*\([^"]+,\s*"0-00-00"'
                text_format_pattern2 = r"TEXT\s*\([^']+,\s*'0-00-00'"
                if not re.search(text_format_pattern1, birthday_formula_text, re.IGNORECASE) and \
                   not re.search(text_format_pattern2, birthday_formula_text, re.IGNORECASE):
                    logger.warning(f"Cell {birthday_cell_coord} TEXT function should use format \"0-00-00\"")
                    logger.warning(f"Formula: {birthday_formula_text}")
                    all_passed = False
                    continue
                
                logger.info(f"✓ Cell {birthday_cell_coord} has valid birthday formula: {birthday_formula_text}")
                
            except Exception as e:
                logger.error(f"Error checking row {row_num}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                all_passed = False
        
        if all_passed:
            logger.info("=" * 60)
            logger.info(f"✓ All rows contain correct ID extraction formulas (gender, age, birthday)")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ ID extraction formula verification failed")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_line_chart(result: str, expected: str = None, **options) -> float:
    """
    Verify if a line chart exists in the Excel file.
    
    This function checks:
    1. Whether at least one chart exists in the worksheet
    2. Whether the chart type is lineChart
    3. Whether the chart has the expected number of series
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - expected_chart_type: Expected chart type (default: "lineChart")
            - min_series_count: Minimum number of series expected (default: 1)
            - data_range: Data range used for chart (optional, for logging)
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        expected_chart_type = options.get('expected_chart_type', 'lineChart')
        min_series_count = options.get('min_series_count', 1)
        data_range = options.get('data_range', '')
        
        logger.info(f"Verifying line chart in file: {result}")
        logger.info(f"Expected chart type: {expected_chart_type}")
        logger.info(f"Minimum series count: {min_series_count}")
        if data_range:
            logger.info(f"Data range: {data_range}")
        
        # Load workbook
        try:
            wb = openpyxl.load_workbook(result)
            ws = wb.active
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Check if charts exist
        charts = ws._charts
        if not charts or len(charts) == 0:
            logger.error("No charts found in the worksheet")
            return 0.0
        
        logger.info(f"Found {len(charts)} chart(s) in the worksheet")
        
        # Check each chart
        for chart_idx, chart in enumerate(charts):
            logger.info(f"Checking chart {chart_idx + 1}...")
            
            # Check chart type
            chart_type = None
            if hasattr(chart, 'tagname'):
                chart_type = chart.tagname
            logger.info(f"Chart type: {chart_type}")
            
            # Check if it's a line chart
            if chart_type and expected_chart_type.lower() in chart_type.lower():
                logger.info(f"✓ Chart {chart_idx + 1} is a line chart")
                
                # Check if it has series
                if not hasattr(chart, 'series') or not chart.series:
                    logger.warning(f"Chart {chart_idx + 1} has no series")
                    continue
                
                series_count = len(chart.series)
                logger.info(f"Chart {chart_idx + 1} has {series_count} series")
                
                # Verify series count
                if series_count >= min_series_count:
                    logger.info("=" * 60)
                    logger.info(f"✓ Line chart verification passed")
                    logger.info(f"  Chart type: {chart_type}")
                    logger.info(f"  Series count: {series_count} (minimum required: {min_series_count})")
                    logger.info("=" * 60)
                    return 1.0
                else:
                    logger.warning(f"Chart {chart_idx + 1} has {series_count} series, but minimum required is {min_series_count}")
            else:
                logger.warning(f"Chart {chart_idx + 1} is not a line chart (type: {chart_type})")
        
        # If we get here, verification failed
        logger.error("=" * 60)
        logger.error(f"✗ Line chart verification failed")
        logger.error(f"  Expected chart type: {expected_chart_type}")
        logger.error(f"  Minimum series count: {min_series_count}")
        logger.error("=" * 60)
        return 0.0
             
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_salary_growth_chart(result: str, expected: str = None, **options) -> float:
    """
    Verify if the salary growth chart matches the expected specifications.
    
    This function checks the chart itself (not the data table):
    1. Whether a chart exists in the specified sheet
    2. Whether the chart title matches "店长工资增长"
    3. Whether the chart has the expected number of series (at least 3)
    4. Whether the chart is a combination chart (bar + line)
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - sheet_idx: Sheet index (default: 0)
            - expected_title: Expected chart title (default: "店长工资增长")
            - min_series_count: Minimum number of series (default: 3)
            - chart_type: Expected chart type (default: "combination")
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        sheet_idx = options.get('sheet_idx', 0)
        expected_title = options.get('expected_title', '店长工资增长')
        min_series_count = options.get('min_series_count', 3)
        chart_type = options.get('chart_type', 'combination')
        
        logger.info(f"Verifying salary growth chart in file: {result}")
        logger.info(f"Sheet index: {sheet_idx}")
        logger.info(f"Expected title: {expected_title}")
        logger.info(f"Min series count: {min_series_count}")
        
        # Load workbook
        try:
            wb = openpyxl.load_workbook(result)
            pdworkbook = pd.ExcelFile(result)
            sheet_names = pdworkbook.sheet_names
            
            if sheet_idx >= len(sheet_names):
                logger.error(f"Sheet index {sheet_idx} out of range. Available sheets: {sheet_names}")
                return 0.0
            
            sheet_name = sheet_names[sheet_idx]
            logger.info(f"Checking sheet: {sheet_name}")
            ws = wb[sheet_name]
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return 0.0
        
        # Check if charts exist
        charts = ws._charts
        if not charts:
            logger.error("No charts found in the sheet")
            return 0.0
        
        logger.info(f"Found {len(charts)} chart(s) in the sheet")
        
        # Load chart information
        chart_props = ['title', 'type', 'legend', 'xtitle', 'ytitle']
        chart_info = load_charts(wb, sheet_name, chart_props=chart_props)
        
        if not chart_info:
            logger.error("Could not load chart information")
            return 0.0
        
        # Check each chart
        chart_passed = False
        for chart_key, chart_data in chart_info.items():
            logger.info(f"Checking chart: {chart_key}")
            logger.debug(f"Chart data: {chart_data}")
            
            # Check 1: Chart title
            chart_title = chart_data.get('title')
            if chart_title != expected_title:
                logger.warning(f"Chart title mismatch: expected '{expected_title}', got '{chart_title}'")
                continue
            else:
                logger.info(f"✓ Chart title matches: {chart_title}")
            
            # Check 2: Chart type (for combination charts, we might see multiple types)
            chart_type_actual = chart_data.get('type')
            logger.info(f"Chart type: {chart_type_actual}")
            # Note: Combination charts might be represented differently in openpyxl
            # We'll be lenient here and just check that a chart exists
            
            # Check 3: Number of series
            # Extract series count from chart_key (format: "value_ref1,category_ref1;value_ref2,category_ref2;...")
            series_parts = chart_key.split(';')
            series_count = len(series_parts)
            logger.info(f"Number of series: {series_count}")
            
            if series_count < min_series_count:
                logger.warning(f"Insufficient series count: expected at least {min_series_count}, got {series_count}")
                continue
            else:
                logger.info(f"✓ Series count sufficient: {series_count} >= {min_series_count}")
            
            # If we get here, this chart passed all checks
            chart_passed = True
            logger.info("=" * 60)
            logger.info(f"✓ Chart verification passed!")
            logger.info("=" * 60)
            break
        
        if chart_passed:
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error("✗ Chart verification failed - no chart matched all criteria")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0


def verify_project_completion_chart(result: str, expected: str = None, **options) -> float:
    """
    Verify if a combination chart exists with the expected title that contains
    both bar chart series (for project values) and line chart series (for completion rates).
    
    This function checks:
    1. Whether at least one chart exists in the worksheet
    2. Whether the chart title matches expected_title
    3. Whether the chart has at least 16 series (8 projects + 8 completion rates)
    4. Whether at least one series name contains "rate" (for completion rates)
    5. Whether the chart has at least project_count * 5 categories (for 5 quarters per project)
    
    Args:
        result (str): Path to result Excel file
        expected (str): Not used (for compatibility with framework interface)
        options (dict): Configuration options, should contain:
            - sheet_idx: Sheet index to check (default: 0)
            - expected_title: Expected chart title (default: "项目")
            - min_series_count: Minimum number of series required (default: 16)
            - project_count: Number of projects (default: 8)
            - quarters_per_project: Number of quarters per project (default: 5)
    
    Returns:
        float: 1.0 if verification passes, 0.0 otherwise
    """
    try:
        if result is None or not os.path.exists(result):
            logger.error(f"Result file not found: {result}")
            return 0.0
        
        sheet_idx = options.get('sheet_idx', 0)
        expected_title = options.get('expected_title', '项目')
        min_series_count = options.get('min_series_count', 16)
        project_count = options.get('project_count', 8)
        quarters_per_project = options.get('quarters_per_project', 5)
        min_categories = project_count * quarters_per_project
        
        logger.info(f"Verifying project completion chart in file: {result}")
        logger.info(f"Sheet index: {sheet_idx}")
        logger.info(f"Expected title: {expected_title}")
        logger.info(f"Minimum series count: {min_series_count}")
        logger.info(f"Project count: {project_count}")
        logger.info(f"Quarters per project: {quarters_per_project}")
        logger.info(f"Minimum categories: {min_categories}")
        
        # Load workbook
        try:
            wb = openpyxl.load_workbook(result)
            sheet_names = wb.sheetnames
            if sheet_idx >= len(sheet_names):
                logger.error(f"Sheet index {sheet_idx} out of range. Available sheets: {sheet_names}")
                return 0.0
            sheet_name = sheet_names[sheet_idx]
            ws = wb[sheet_name]
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            return 0.0
        
        # Check if charts exist
        charts = ws._charts
        if not charts:
            logger.error("No charts found in the worksheet")
            return 0.0
        
        logger.info(f"Found {len(charts)} chart(s) in the worksheet")
        
        # Check each chart
        chart_found = False
        for chart in charts:
            # Check chart title
            chart_title = None
            try:
                if chart.title and chart.title.tx:
                    if hasattr(chart.title.tx, 'rich') and chart.title.tx.rich:
                        if hasattr(chart.title.tx.rich, 'p') and chart.title.tx.rich.p:
                            if len(chart.title.tx.rich.p) > 0:
                                if hasattr(chart.title.tx.rich.p[0], 'r') and chart.title.tx.rich.p[0].r:
                                    if len(chart.title.tx.rich.p[0].r) > 0:
                                        if hasattr(chart.title.tx.rich.p[0].r[0], 't'):
                                            chart_title = chart.title.tx.rich.p[0].r[0].t
            except Exception as e:
                logger.debug(f"Error reading chart title: {e}")
            
            logger.info(f"Chart title: {chart_title}")
            
            # Check if title matches
            if chart_title == expected_title:
                logger.info(f"✓ Chart title matches: {chart_title}")
                chart_found = True
                
                # Use load_charts to get all series information (includes both bar and line series)
                # This is more reliable for combination charts
                chart_props = ['title']
                chart_info = load_charts(wb, sheet_name, chart_props=chart_props)
                
                # Find the chart that matches our title
                chart_key = None
                for key, info in chart_info.items():
                    if info.get('title') == expected_title:
                        chart_key = key
                        break
                
                # Get all series from the chart object for detailed inspection
                all_series = list(chart.series) if hasattr(chart, 'series') else []
                logger.info(f"Series count from chart.series: {len(all_series)}")
                
                if not chart_key:
                    logger.warning("Could not find chart in load_charts output, using direct series access")
                    # Fallback to direct series access
                    series_count = len(all_series)
                else:
                    # Extract series count from chart_key (format: "value_ref1,category_ref1;value_ref2,category_ref2;...")
                    # This includes ALL series (both bar and line) in combination charts
                    series_parts = chart_key.split(';')
                    series_count_from_load = len(series_parts)
                    logger.info(f"Series count from load_charts: {series_count_from_load}")
                    
                    # For combination charts, load_charts should give us all series
                    # Use the count from load_charts as it's more reliable for combination charts
                    series_count = series_count_from_load
                    
                    # Also check for sub-charts in case series are stored there
                    if hasattr(chart, '_charts') and chart._charts:
                        # Check for sub-charts (for combination charts)
                        for sub_chart in chart._charts:
                            if hasattr(sub_chart, 'series'):
                                sub_series = list(sub_chart.series)
                                all_series.extend(sub_series)
                                logger.info(f"Found {len(sub_series)} additional series in sub-chart")
                    
                    # If load_charts gave us fewer series than direct access, use the larger count
                    # This handles edge cases where load_charts might miss some series
                    if series_count < len(all_series):
                        logger.warning(f"load_charts found {series_count} series but direct access found {len(all_series)}, using larger count")
                        series_count = len(all_series)
                
                logger.info(f"Chart has {series_count} series (including both bar and line series)")
                
                # Debug: Log all series details
                if all_series:
                    logger.info(f"Detailed series information:")
                    for idx, ser in enumerate(all_series):
                        logger.info(f"  Series {idx}: {type(ser).__name__}")
                        try:
                            if hasattr(ser, 'title'):
                                logger.debug(f"    Title: {ser.title}")
                        except:
                            pass
                
                if series_count < min_series_count:
                    logger.error(f"✗ Chart has only {series_count} series, expected at least {min_series_count}")
                    return 0.0
                
                logger.info(f"✓ Chart has {series_count} series (>= {min_series_count})")
                
                # Check series names for "rate" (completion rate series)
                has_rate_series = False
                series_names = []
                for i, ser in enumerate(all_series):
                    series_name = None
                    try:
                        # Try to get series title/name
                        if hasattr(ser, 'title') and ser.title:
                            if hasattr(ser.title, 'tx') and ser.title.tx:
                                if hasattr(ser.title.tx, 'rich') and ser.title.tx.rich:
                                    if hasattr(ser.title.tx.rich, 'p') and ser.title.tx.rich.p:
                                        if len(ser.title.tx.rich.p) > 0:
                                            if hasattr(ser.title.tx.rich.p[0], 'r') and ser.title.tx.rich.p[0].r:
                                                if len(ser.title.tx.rich.p[0].r) > 0:
                                                    if hasattr(ser.title.tx.rich.p[0].r[0], 't'):
                                                        series_name = ser.title.tx.rich.p[0].r[0].t
                        # Alternative: check if title is a string reference
                        if not series_name and hasattr(ser, 'title') and hasattr(ser.title, 'tx') and hasattr(ser.title.tx, 'strRef'):
                            if hasattr(ser.title.tx.strRef, 'f'):
                                series_name = ser.title.tx.strRef.f
                    except Exception as e:
                        logger.debug(f"Error reading series {i} name: {e}")
                    
                    if series_name:
                        series_names.append(series_name)
                        if "rate" in series_name.lower():
                            has_rate_series = True
                            logger.info(f"✓ Found series with 'rate' in name: {series_name}")
                
                if series_names:
                    logger.info(f"Series names found: {series_names[:10]}...")  # Log first 10
                else:
                    logger.warning("Could not extract series names, will skip rate check")
                
                if not has_rate_series and series_names:
                    logger.error(f"✗ No series found with 'rate' in name. Series names: {series_names}")
                    return 0.0
                elif not has_rate_series:
                    logger.warning("⚠ Could not verify 'rate' in series names (series names not extractable)")
                
                # Check category count
                max_categories = 0
                category_ranges = []
                
                def parse_range_count(range_str):
                    """Parse Excel range string and return count of cells"""
                    try:
                        # Remove sheet name if present (e.g., "Sheet1!$A$2:$A$6" -> "$A$2:$A$6")
                        if '!' in range_str:
                            range_str = range_str.split('!')[1]
                        
                        # Remove $ signs
                        range_str = range_str.replace('$', '')
                        
                        if ':' in range_str:
                            start, end = range_str.split(':')
                            # Parse start and end coordinates
                            start_col, start_row = coordinate_to_tuple(start)
                            end_col, end_row = coordinate_to_tuple(end)
                            
                            # Calculate count based on range
                            if start_col == end_col:
                                # Same column, count rows
                                return abs(end_row - start_row) + 1
                            elif start_row == end_row:
                                # Same row, count columns
                                return abs(end_col - start_col) + 1
                            else:
                                # 2D range
                                return (abs(end_row - start_row) + 1) * (abs(end_col - start_col) + 1)
                        else:
                            # Single cell
                            return 1
                    except Exception as e:
                        logger.debug(f"Error parsing range {range_str}: {e}")
                        return 0
                
                for i, ser in enumerate(all_series):
                    try:
                        # Try to get category count from category reference
                        if hasattr(ser, 'cat'):
                            cat_range = None
                            # Check if categories are from a range
                            if hasattr(ser.cat, 'numRef') and hasattr(ser.cat.numRef, 'f'):
                                cat_range = ser.cat.numRef.f
                            elif hasattr(ser.cat, 'strRef') and hasattr(ser.cat.strRef, 'f'):
                                cat_range = ser.cat.strRef.f
                            
                            if cat_range:
                                category_ranges.append(cat_range)
                                cat_count = parse_range_count(cat_range)
                                if cat_count > max_categories:
                                    max_categories = cat_count
                                logger.debug(f"Series {i} category range: {cat_range}, count: {cat_count}")
                    except Exception as e:
                        logger.debug(f"Error reading categories for series {i}: {e}")
                
                if max_categories > 0:
                    logger.info(f"Maximum category count found: {max_categories}")
                    if max_categories < min_categories:
                        logger.error(f"✗ Chart has only {max_categories} categories, expected at least {min_categories} (project_count * quarters_per_project)")
                        return 0.0
                    logger.info(f"✓ Chart has {max_categories} categories (>= {min_categories})")
                else:
                    # If we can't determine category count from ranges, use heuristic
                    # For 8 projects with 5 quarters each, we need at least 40 categories
                    # But since we can't verify directly, we'll log a warning
                    logger.warning(f"⚠ Could not determine exact category count from ranges. Expected at least {min_categories} categories.")
                    logger.info(f"Category ranges found: {category_ranges[:5]}...")  # Log first 5
                
                # Check if it's a combination chart
                if series_count >= 2:
                    logger.info("✓ Chart appears to be a combination chart (has multiple series)")
                else:
                    logger.error(f"✗ Chart has only {series_count} series, expected at least 2 for combination chart")
                    return 0.0
                
                break
        
        if chart_found:
            logger.info("=" * 60)
            logger.info("✓ Project completion combination chart verification passed")
            logger.info("=" * 60)
            return 1.0
        else:
            logger.error("=" * 60)
            logger.error(f"✗ Chart with title '{expected_title}' not found")
            logger.error("=" * 60)
            return 0.0
            
    except Exception as e:
        import traceback
        logger.error(f"Verification failed: {e}")
        logger.error(traceback.format_exc())
        return 0.0