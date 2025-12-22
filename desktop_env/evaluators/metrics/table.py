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
