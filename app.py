"""Streamlit interface for the AI SQL Data Assistant."""

from __future__ import annotations

import csv
import difflib
import html
import json
import math
import re
import sqlite3
import sys
import tempfile
import warnings
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import sqlglot
import speech_recognition as sr
from sqlglot import exp
from gtts import gTTS
from gtts.tts import gTTSError
from langdetect import DetectorFactory, LangDetectException, detect_langs
import streamlit as st
from streamlit.errors import StreamlitAPIException
from streamlit_mic_recorder import mic_recorder

from database import DatabaseError, execute_select_query, get_database_schema, test_database_connection
from groq_client import GroqRequestError, generate_response
from sql_generator import generate_answer, generate_sql
from sql_validator import apply_row_limit, validate_sql

MAX_HISTORY_MESSAGES = 8
SUPPORTED_LANGUAGES = {
    "en": {"name": "English", "tts": "en"},
    "mr": {"name": "Marathi", "tts": "mr"},
    "hi": {"name": "Hindi", "tts": "hi"},
}
DetectorFactory.seed = 0
SUPPORTED_UPLOAD_TYPES = {
    ".csv": "csv",
    ".xlsx": "excel",
    ".json": "json",
    ".parquet": "parquet",
}


@st.cache_data(ttl=20, show_spinner=False)
def _connection_status() -> tuple[bool, str]:
    return test_database_connection()


def _configure_dashboard_style() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: #f5f7fb; }
        [data-testid="stAppViewContainer"] > .main { background: #f5f7fb; }
        [data-testid="stMainBlockContainer"] {
            max-width: 1440px;
            padding-top: 2.5rem;
            padding-bottom: 2rem;
        }
        [data-testid="stMain"] [data-testid="stVerticalBlock"] { gap: 0.8rem; }
        h1 { color: #18324f; letter-spacing: -0.035em; margin-bottom: 0.35rem; }
        h2, h3 { color: #234765; letter-spacing: -0.02em; }
        h2 { margin-top: 1rem; margin-bottom: 0.55rem; }
        [data-testid="stVerticalBlockBorderWrapper"] {
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.82);
            box-shadow: 0 4px 16px rgba(28, 48, 74, 0.045);
            padding: 0.8rem 1rem;
        }
        [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlock"] {
            gap: 0.65rem;
        }
        .kpi-card {
            box-sizing: border-box;
            min-height: 118px;
            height: 100%;
            background: #ffffff;
            border: 1px solid #e3eaf2;
            border-top: 3px solid #537da5;
            border-radius: 14px;
            padding: 1rem 1.1rem;
            box-shadow: 0 4px 14px rgba(34, 60, 88, 0.055);
        }
        .kpi-label {
            font-size: 0.82rem;
            color: #62748a;
            margin-bottom: 0.55rem;
            font-weight: 600;
        }
        .kpi-value {
            font-size: clamp(1.35rem, 2vw, 1.8rem);
            font-weight: 700;
            color: #183b5b;
            line-height: 1.2;
            overflow-wrap: anywhere;
        }
        .summary-box {
            background: #f8fafc;
            border: 1px solid #e7edf4;
            border-radius: 12px;
            padding: 0.7rem 0.9rem;
            margin-bottom: 0.5rem;
        }
        [data-testid="stButton"] > button,
        [data-testid="stDownloadButton"] > button,
        [data-testid="stFormSubmitButton"] > button {
            border-radius: 10px;
            border-color: #d8e1eb;
            font-weight: 600;
            transition: border-color 120ms ease, background-color 120ms ease;
        }
        [data-testid="stButton"] > button:hover,
        [data-testid="stDownloadButton"] > button:hover,
        [data-testid="stFormSubmitButton"] > button:hover {
            border-color: #7795b2;
            color: #183b5b;
        }
        [data-testid="stSidebar"] {
            background: #f0f4f8;
            border-right: 1px solid #e1e8ef;
        }
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            margin-top: 0.7rem;
            margin-bottom: 0.45rem;
        }
        @media (max-width: 720px) {
            [data-testid="stMainBlockContainer"] {
                padding: 1.75rem 0.85rem 1.5rem;
            }
            .kpi-card { min-height: 100px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _style_plotly_figure(figure, title: str, x_title: str | None = None, y_title: str | None = None):
    figure.update_layout(
        title={"text": title, "x": 0.02, "xanchor": "left"},
        font={"family": "Arial, sans-serif", "color": "#334155", "size": 12},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        margin={"l": 24, "r": 20, "t": 58, "b": 24},
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
    )
    figure.update_xaxes(
        title=x_title,
        showgrid=False,
        automargin=True,
        linecolor="#dbe3ec",
        tickfont={"color": "#596b80"},
    )
    figure.update_yaxes(
        title=y_title,
        showgrid=True,
        gridcolor="#edf1f5",
        zeroline=False,
        automargin=True,
        tickformat=",~g",
        tickfont={"color": "#596b80"},
    )
    return figure


def _get_debug_info() -> dict[str, str]:
    try:
        import openpyxl
        openpyxl_version = openpyxl.__version__
    except Exception:
        openpyxl_version = "Not installed"

    return {
        "python_executable": sys.executable,
        "python_version": sys.version.split()[0],
        "pandas_version": pd.__version__,
        "openpyxl_version": openpyxl_version,
    }


def _ensure_excel_support() -> None:
    try:
        import openpyxl  # noqa: F401
    except Exception as exc:
        raise RuntimeError(
            "Excel support requires openpyxl. Install it in the same Python environment used by Streamlit: "
            "'" + sys.executable + "' -m pip install openpyxl"
        ) from exc


# --- NEW CODE: Universal File Upload Dashboard ---
def _detect_uploaded_file_type(file_name: str) -> str:
    file_extension = Path(file_name).suffix.lower()
    if file_extension not in SUPPORTED_UPLOAD_TYPES:
        supported_types = ", ".join(sorted(SUPPORTED_UPLOAD_TYPES.keys()))
        raise ValueError(f"Unsupported file type: {file_name}. Supported types are: {supported_types}.")
    return SUPPORTED_UPLOAD_TYPES[file_extension]


def _detect_csv_delimiter(file_bytes: bytes) -> str:
    sample = file_bytes[:65536].decode("utf-8-sig", errors="replace")
    candidates = [",", ";", "\t"]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        return dialect.delimiter
    except csv.Error:
        counts = {delimiter: sample.count(delimiter) for delimiter in candidates}
        best_delimiter = max(counts, key=counts.get)
        return best_delimiter if counts[best_delimiter] > 0 else ","


def _split_csv_like_single_column_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    if dataframe.shape[1] != 1:
        return dataframe

    header_text = str(dataframe.columns[0]).strip()
    values = dataframe.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    if not header_text or not values:
        return dataframe

    candidate_delimiters = [",", ";", "\t"]
    sample_values = [header_text, *values[:10]]
    delimiter_counts = {
        delimiter: sum(value.count(delimiter) for value in sample_values if isinstance(value, str))
        for delimiter in candidate_delimiters
    }
    chosen_delimiter = max(delimiter_counts, key=delimiter_counts.get)
    if delimiter_counts[chosen_delimiter] == 0 or chosen_delimiter not in header_text:
        return dataframe

    try:
        header = next(csv.reader([header_text], delimiter=chosen_delimiter))
    except csv.Error:
        return dataframe
    if len(header) <= 1:
        return dataframe

    data_rows: list[list[str]] = []
    for row in values:
        if not row:
            continue
        try:
            data_rows.append(next(csv.reader([row], delimiter=chosen_delimiter)))
        except csv.Error:
            return dataframe
    if not data_rows or not all(len(row) == len(header) for row in data_rows):
        return dataframe

    normalized_dataframe = pd.DataFrame(data_rows, columns=header)
    normalized_dataframe.columns = [str(column).strip() for column in normalized_dataframe.columns]
    return normalized_dataframe


def _normalize_uploaded_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    dataframe = _split_csv_like_single_column_dataframe(dataframe)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]

    for column in dataframe.select_dtypes(include=["object", "string"]).columns:
        values = dataframe[column].map(
            lambda value: (value.strip() or pd.NA) if isinstance(value, str) else value
        )
        numeric_values = pd.to_numeric(values, errors="coerce")
        non_missing_count = int(values.notna().sum())
        parsed_count = int(numeric_values.notna().sum())
        if non_missing_count and parsed_count == non_missing_count:
            dataframe[column] = numeric_values
        else:
            dataframe[column] = values

    return dataframe


def _detect_date_columns(dataframe: pd.DataFrame) -> list[object]:
    detected = []
    for column in dataframe.columns:
        series = dataframe[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            detected.append(column)
        elif not (
            pd.api.types.is_numeric_dtype(series)
            or pd.api.types.is_bool_dtype(series)
        ):
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)
                    parsed = pd.to_datetime(series, errors="coerce")
            except (TypeError, ValueError):
                continue
            non_missing = int(series.notna().sum())
            if non_missing and int(parsed.notna().sum()) / non_missing >= 0.8:
                detected.append(column)
    return detected


def _detect_numeric_columns(dataframe: pd.DataFrame) -> list[object]:
    numeric_columns = []
    for column in dataframe.columns:
        series = dataframe[column]
        if pd.api.types.is_bool_dtype(series):
            continue
        if pd.api.types.is_numeric_dtype(series):
            numeric_columns.append(column)
            continue
        if pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
            values = series.dropna().astype("string").str.strip()
            if not values.empty and pd.to_numeric(values, errors="coerce").notna().mean() >= 0.8:
                numeric_columns.append(column)
    return numeric_columns


def _detect_categorical_columns(dataframe: pd.DataFrame) -> list[object]:
    date_columns = set(_detect_date_columns(dataframe))
    numeric_columns = set(_detect_numeric_columns(dataframe))
    boolean_columns = set(_detect_boolean_columns(dataframe))
    return [
        column for column in dataframe.columns
        if column not in date_columns
        and column not in numeric_columns
        and column not in boolean_columns
        and not pd.api.types.is_bool_dtype(dataframe[column])
    ]


def _detect_boolean_columns(dataframe: pd.DataFrame) -> list[object]:
    boolean_columns = []
    for column in dataframe.columns:
        series = dataframe[column]
        if pd.api.types.is_bool_dtype(series):
            boolean_columns.append(column)
            continue
        if not (pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series)):
            continue
        values = set(series.dropna().astype(str).str.strip().str.casefold().unique())
        if values and values.issubset({"true", "false", "yes", "no", "1", "0"}):
            boolean_columns.append(column)
    return boolean_columns


def _detect_id_columns(dataframe: pd.DataFrame) -> list[object]:
    id_tokens = {"id", "index", "serial", "uuid", "code", "pk", "rownumber", "visitid"}
    business_tokens = {
        "visitor", "visitors", "tourist", "tourists", "count", "quantity", "booking", "bookings",
        "revenue", "sales", "sale", "income", "spending", "spend", "amount", "price",
        "rating", "score", "review", "stars", "star",
    }
    detected = []
    for column in dataframe.columns:
        spaced_name = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(column))
        tokens = set(re.findall(r"[a-z0-9]+", spaced_name.casefold()))
        compact_name = re.sub(r"[^a-z0-9]+", "", str(column).casefold())
        identifier_named = bool(tokens & id_tokens) or any(
            token in compact_name for token in ("id", "index", "serial", "rownumber")
        )
        generic_identifier_name = bool(tokens & {"no", "number"})
        meaningful_metric_name = bool(tokens & business_tokens) or any(
            compact_name.startswith(token)
            for token in ("visitors", "visitor", "tourists", "tourist", "revenue", "rating", "score")
        )
        if identifier_named or (generic_identifier_name and not meaningful_metric_name):
            detected.append(column)
    return detected


def _detect_column_types(dataframe: pd.DataFrame) -> dict[str, list[object]]:
    return {
        "numeric": _detect_numeric_columns(dataframe),
        "categorical": _detect_categorical_columns(dataframe),
        "date": _detect_date_columns(dataframe),
        "boolean": _detect_boolean_columns(dataframe),
        "id": _detect_id_columns(dataframe),
    }


def _detect_business_metrics(dataframe: pd.DataFrame) -> dict[str, list[object]]:
    detected_types = _detect_column_types(dataframe)
    numeric_columns = detected_types["numeric"]
    id_columns = set(detected_types["id"])

    def matching(tokens: set[str]) -> list[object]:
        return [
            column
            for column in numeric_columns
            if column not in id_columns
            and (
                set(re.findall(r"[a-z0-9]+", str(column).casefold())) & tokens
                or any(
                    re.search(rf"\b{re.escape(token)}", re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(column)).casefold())
                    for token in tokens
                )
            )
        ]

    return {
        "rating": matching({"rating", "score", "review", "stars", "star"}),
        "revenue": matching({"revenue", "sales", "sale", "income", "amount", "spending", "spend"}),
        "visitors": matching({"visitors", "visitor", "tourists", "tourist", "bookings", "booking", "quantity", "count"}),
    }


def _generate_dataset_profile(dataframe: pd.DataFrame, file_name: str) -> dict[str, object]:
    detected_types = _detect_column_types(dataframe)
    available_columns = [str(column) for column in dataframe.columns]
    date_columns = detected_types["date"]
    boolean_columns = detected_types["boolean"]
    numeric_columns = detected_types["numeric"]
    categorical_columns = detected_types["categorical"]
    id_columns = detected_types["id"]
    metrics = _detect_business_metrics(dataframe)
    metric_columns = list(dict.fromkeys(
        metrics["rating"] + metrics["revenue"] + metrics["visitors"]
    ))
    basic_statistics = {}
    for column in metric_columns:
        values = pd.to_numeric(dataframe[column], errors="coerce").dropna()
        if not values.empty:
            basic_statistics[str(column)] = {
                "count": int(values.count()),
                "sum": float(values.sum()),
                "mean": float(values.mean()),
                "min": float(values.min()),
                "max": float(values.max()),
            }
    dimensions = [
        column for column in dataframe.columns
        if column in categorical_columns and column not in id_columns
    ]
    try:
        duplicate_rows = int(dataframe.duplicated().sum())
    except TypeError:
        normalized = dataframe.copy()
        for column in normalized.columns:
            normalized[column] = normalized[column].map(
                lambda value: json.dumps(value, sort_keys=True, default=str)
                if isinstance(value, (dict, list, tuple, set))
                else value
            )
        duplicate_rows = int(normalized.duplicated().sum())

    return {
        "file_name": file_name,
        "file_type": _detect_uploaded_file_type(file_name),
        "rows": int(len(dataframe)),
        "columns_count": int(len(dataframe.columns)),
        "columns": available_columns,
        "dtypes": {str(column): str(dataframe[column].dtype) for column in dataframe.columns},
        "numeric_columns": [str(column) for column in numeric_columns],
        "categorical_columns": [str(column) for column in categorical_columns],
        "date_columns": [str(column) for column in date_columns],
        "boolean_columns": [str(column) for column in boolean_columns],
        "id_columns": [str(column) for column in id_columns],
        "business_metrics": {
            name: [str(column) for column in columns]
            for name, columns in metrics.items()
        },
        "important_dimensions": [str(column) for column in dimensions],
        "basic_statistics": basic_statistics,
        "missing_values": int(dataframe.isna().sum().sum()),
        "duplicate_rows": duplicate_rows,
    }


def _is_valid_column(dataframe: pd.DataFrame, column_name: object) -> bool:
    return column_name in dataframe.columns


def _generate_chart_suggestions(dataframe: pd.DataFrame) -> list[str]:
    if dataframe.empty or not len(dataframe.columns):
        return []

    available_columns = list(dataframe.columns)
    detected_types = _detect_column_types(dataframe)
    id_columns = _detect_id_columns(dataframe)
    date_columns = [
        column for column in detected_types["date"]
        if _is_valid_column(dataframe, column)
    ]
    categorical_columns = [
        column for column in detected_types["categorical"]
        if _is_valid_column(dataframe, column)
    ]
    numeric_columns = [
        column for column in detected_types["numeric"]
        if column not in id_columns
        and _is_valid_column(dataframe, column)
    ]
    suggestions: list[str] = []

    def add_suggestion(text: str, *columns: object) -> None:
        if columns and all(
            _is_valid_column(dataframe, column)
            and column in available_columns
            for column in columns
        ):
            suggestions.append(text)

    date_metrics = [(date_column, metric) for date_column in date_columns for metric in numeric_columns]
    if date_metrics:
        date_column, metric = date_metrics[0]
        add_suggestion(f"Line chart: {metric} over {date_column}", metric, date_column)

    categorical_metrics = [
        (category, metric)
        for category in categorical_columns
        for metric in numeric_columns
    ]
    if categorical_metrics:
        category, metric = categorical_metrics[0]
        add_suggestion(f"Bar chart: {metric} by {category}", metric, category)

    if categorical_columns:
        category = categorical_columns[0]
        add_suggestion(f"Category distribution: record count by {category}", category)

    if len(numeric_columns) >= 2:
        add_suggestion(
            f"Scatter chart: {numeric_columns[0]} versus {numeric_columns[1]}",
            numeric_columns[0],
            numeric_columns[1],
        )

    return list(dict.fromkeys(suggestions))[:4]


def _find_smart_chart_filter(
    dataframe: pd.DataFrame,
    question: str,
    categorical_columns: list[object],
) -> dict[str, object] | None:
    normalized_question = " ".join(re.findall(r"[a-z0-9]+", question.casefold()))
    question_tokens = _chart_request_tokens(question)
    matches: list[tuple[int, int, int, object, object]] = []

    for column in categorical_columns:
        if not _is_valid_column(dataframe, column):
            continue
        for value in dataframe[column].dropna().unique():
            value_text = " ".join(re.findall(r"[a-z0-9]+", str(value).casefold()))
            if len(value_text) < 2:
                continue
            pattern = rf"(?<![a-z0-9]){re.escape(value_text)}(?![a-z0-9])"
            if not re.search(pattern, normalized_question):
                continue
            column_tokens = set(_chart_name_tokens(column))
            semantic_score = len(column_tokens & question_tokens)
            preferred_column_score = next(
                (
                    score
                    for term, score in (
                        ("city", 5),
                        ("destination", 4),
                        ("state", 3),
                        ("region", 2),
                        ("country", 1),
                    )
                    if term in column_tokens
                ),
                0,
            )
            matches.append(
                (len(value_text), semantic_score, preferred_column_score, column, value)
            )

    if not matches:
        return None
    matches.sort(key=lambda match: match[:3], reverse=True)
    best_match = matches[0]
    tied_columns = {
        match[3]
        for match in matches
        if match[:3] == best_match[:3]
    }
    if len(tied_columns) > 1:
        return None
    return {"column": best_match[3], "value": best_match[4]}


def _match_column_from_text(
    dataframe: pd.DataFrame,
    text: str,
    candidates: list[object],
) -> object | None:
    available_candidates = [
        column for column in candidates
        if _is_valid_column(dataframe, column)
    ]
    if not available_candidates:
        return None

    normalized_text = _normalize_column_name(text)
    exact_matches = [
        column for column in available_candidates
        if _normalize_column_name(column) in normalized_text
    ]
    if len(exact_matches) == 1:
        return exact_matches[0]
    if exact_matches:
        exact_matches.sort(key=lambda column: len(_normalize_column_name(column)), reverse=True)
        if len(exact_matches) == 1 or len(_normalize_column_name(exact_matches[0])) > len(
            _normalize_column_name(exact_matches[1])
        ):
            return exact_matches[0]

    request_tokens = _chart_request_tokens(text)
    scored = [
        (column, _chart_column_score(column, request_tokens))
        for column in available_candidates
    ]
    scored = [(column, score) for column, score in scored if score > 0]
    if scored:
        scored.sort(key=lambda item: item[1], reverse=True)
        if len(scored) == 1 or scored[0][1] > scored[1][1]:
            return scored[0][0]

    request_words = [
        word for word in re.findall(r"[a-z0-9]+", normalized_text)
        if word not in {
            "show", "me", "top", "bottom", "by", "over", "time", "trend",
            "average", "avg", "mean", "total", "sum", "count", "minimum",
            "maximum", "highest", "lowest", "compare", "versus", "against",
            "distribution", "of", "the", "for", "in", "on", "and",
        }
    ]
    fuzzy_matches = []
    for column in available_candidates:
        normalized_column = _normalize_column_name(column)
        column_words = normalized_column.split()
        similarity = max(
            (
                difflib.SequenceMatcher(None, word, column_word).ratio()
                for word in request_words
                for column_word in column_words
            ),
            default=0.0,
        )
        if similarity >= 0.86:
            fuzzy_matches.append((column, similarity))
    fuzzy_matches.sort(key=lambda item: item[1], reverse=True)
    if fuzzy_matches and (
        len(fuzzy_matches) == 1 or fuzzy_matches[0][1] > fuzzy_matches[1][1]
    ):
        return fuzzy_matches[0][0]
    return None


def _parse_smart_chart_request(
    question: str,
    dataframe: pd.DataFrame,
) -> dict[str, object] | None:
    if dataframe.empty or not len(dataframe.columns) or not question.strip():
        return None

    columns = list(dataframe.columns)
    available_columns = [str(column) for column in dataframe.columns]
    detected_types = _detect_column_types(dataframe)
    id_columns = set(detected_types["id"])
    numeric_columns = [
        column for column in columns
        if str(column) in available_columns
        and column in detected_types["numeric"]
        and column not in id_columns
        and _is_valid_column(dataframe, column)
    ]
    category_columns = [
        column for column in columns
        if column in detected_types["categorical"]
        and column not in id_columns
        and _is_valid_column(dataframe, column)
    ]
    date_columns = [
        column for column in columns
        if column in detected_types["date"] and _is_valid_column(dataframe, column)
    ]
    matched_filter = _find_smart_chart_filter(dataframe, question, category_columns)
    matching_question = question
    if matched_filter is not None:
        value_text = str(matched_filter["value"])
        matching_question = re.sub(
            rf"(?<![a-z0-9]){re.escape(value_text)}(?![a-z0-9])",
            " ",
            matching_question,
            flags=re.IGNORECASE,
        )
    question_lower = matching_question.lower()
    has_trend_intent = any(
        term in question_lower
        for term in ("trend", "over time", "monthly", "weekly", "daily", "yearly", "by month", "by year")
    )
    is_comparison = bool(re.search(r"\b(vs|versus|against|compare)\b", question_lower))
    is_distribution = any(
        term in question_lower
        for term in ("distribution", "share", "proportion", "percentage", "percent", "pie")
    )

    if is_comparison:
        comparison_parts = re.split(
            r"\b(?:vs|versus|against|and|to)\b",
            question,
            maxsplit=1,
            flags=re.IGNORECASE,
        )
        if len(comparison_parts) != 2:
            return None
        left_metric = _match_column_from_text(dataframe, comparison_parts[0], numeric_columns)
        right_metric = _match_column_from_text(dataframe, comparison_parts[1], numeric_columns)
        if (
            left_metric is None
            or right_metric is None
            or left_metric == right_metric
            or not _is_valid_column(dataframe, left_metric)
            or not _is_valid_column(dataframe, right_metric)
        ):
            return None
        specification = {
            "x": left_metric,
            "y": right_metric,
            "aggregation": "none",
            "chart_type": "Scatter Chart",
            "top_n": None,
            "sort_ascending": False,
            "date_columns": _chart_date_columns(dataframe),
            "question": question,
        }
        if matched_filter is not None:
            specification["filter_column"] = matched_filter["column"]
            specification["filter_value"] = matched_filter["value"]
        return specification

    metric_column = _match_column_from_text(dataframe, matching_question, numeric_columns)
    date_column = _match_column_from_text(dataframe, matching_question, date_columns)
    category_column = _match_column_from_text(dataframe, matching_question, category_columns)
    explicit_count_request = bool(
        re.search(r"\b(count|number|records|all|each)\b", question_lower)
    )
    if (
        matched_filter is not None
        and metric_column is None
        and category_column is None
        and not explicit_count_request
    ):
        return None

    if has_trend_intent:
        if date_column is None and len(date_columns) == 1:
            date_column = date_columns[0]
        if date_column is None or metric_column is None:
            return None
        x_column = date_column
    else:
        x_column = category_column
        if x_column is None and matched_filter is not None:
            remaining_categories = [
                column for column in category_columns
                if column != matched_filter["column"]
            ]
            category_priority = (
                "destination", "city", "category", "region", "state",
                "country", "type",
            )
            preferred_categories = [
                (
                    next(
                        (
                            priority
                            for priority, token in enumerate(category_priority)
                            if token in set(_chart_name_tokens(column))
                        ),
                        len(category_priority),
                    ),
                    column,
                )
                for column in remaining_categories
            ]
            preferred_categories.sort(key=lambda item: item[0])
            x_column = (
                preferred_categories[0][1]
                if preferred_categories
                else remaining_categories[0]
                if remaining_categories
                else matched_filter["column"]
            )

    if x_column is None and is_distribution and metric_column is not None:
        specification = {
            "x": None,
            "y": metric_column,
            "aggregation": "none",
            "chart_type": "Histogram",
            "top_n": None,
            "sort_ascending": False,
            "date_columns": _chart_date_columns(dataframe),
            "question": question,
        }
        if matched_filter is not None:
            specification["filter_column"] = matched_filter["column"]
            specification["filter_value"] = matched_filter["value"]
        return specification

    if x_column is None:
        return None

    aggregation_match = re.search(
        r"\b(average|avg|mean|total|sum|count|minimum|min|maximum|max|highest|lowest)\b",
        question_lower,
    )
    aggregation_word = aggregation_match.group(1) if aggregation_match else ""
    if metric_column is None and aggregation_word not in {"count"} and (
        aggregation_word or re.search(r"\bby\b", question_lower)
    ):
        return None
    if aggregation_word in {"average", "avg", "mean"}:
        aggregation = "mean"
    elif aggregation_word in {"minimum", "min", "lowest"}:
        aggregation = "min"
    elif aggregation_word in {"maximum", "max", "highest"}:
        aggregation = "max"
    elif aggregation_word == "count":
        aggregation = "count"
    elif aggregation_word in {"total", "sum"}:
        aggregation = "sum"
    elif metric_column is None:
        aggregation = "count"
    else:
        aggregation = "sum"

    if metric_column is None and aggregation != "count":
        return None

    limit_match = re.search(r"\b(top|bottom)\s+(\d+)\b", matching_question, flags=re.IGNORECASE)
    ranking_intent = bool(re.search(r"\b(top|bottom)\b", question_lower))
    if limit_match:
        limit = int(limit_match.group(2))
        if limit <= 0:
            return None
        sort_ascending = limit_match.group(1).lower() == "bottom"
    else:
        sort_ascending = aggregation in {"min"}
        limit = 10 if ranking_intent else None
    if limit is None and dataframe[x_column].nunique(dropna=True) > 10:
        limit = 10

    if has_trend_intent:
        chart_type = "Line Chart"
    elif is_distribution and dataframe[x_column].nunique(dropna=True) <= 8:
        chart_type = "Pie Chart"
    else:
        chart_type = "Bar Chart"

    specification = {
        "x": x_column,
        "y": metric_column,
        "aggregation": aggregation,
        "chart_type": chart_type,
        "top_n": limit,
        "sort_ascending": sort_ascending,
        "date_columns": _chart_date_columns(dataframe),
        "question": question,
    }
    if matched_filter is not None:
        specification["filter_column"] = matched_filter["column"]
        specification["filter_value"] = matched_filter["value"]
    return specification


def _load_uploaded_dataframe(uploaded_file) -> pd.DataFrame:
    file_type = _detect_uploaded_file_type(uploaded_file.name)
    uploaded_file.seek(0)
    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        raise ValueError("The uploaded file is empty.")
    buffer = BytesIO(file_bytes)

    if file_type == "csv":
        delimiter = _detect_csv_delimiter(file_bytes)
        dataframe = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            buffer.seek(0)
            try:
                dataframe = pd.read_csv(buffer, sep=delimiter, engine="python", encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if dataframe is None:
            raise ValueError("The CSV encoding could not be detected. Save the file as UTF-8 and retry.")
        return _normalize_uploaded_dataframe(dataframe)

    if file_type == "excel":
        _ensure_excel_support()
        dataframe = pd.read_excel(buffer, engine="openpyxl")
        return _normalize_uploaded_dataframe(dataframe)

    if file_type == "json":
        dataframe = pd.read_json(buffer)
        return _normalize_uploaded_dataframe(dataframe)

    if file_type == "parquet":
        dataframe = pd.read_parquet(buffer)
        return _normalize_uploaded_dataframe(dataframe)

    raise ValueError(f"Unsupported file type: {file_type}")


def _apply_sidebar_filters(dataframe: pd.DataFrame, file_name: str) -> pd.DataFrame:
    filtered_dataframe = dataframe.copy()

    with st.sidebar:
        st.subheader("🎛️ Filters")

        if dataframe.empty or len(dataframe.columns) == 0:
            st.caption("No dataset values are available to filter.")
            return filtered_dataframe

        filter_key = f"sidebar_filters_{file_name}"
        if filter_key not in st.session_state:
            st.session_state[filter_key] = {}

        def filter_label(column_name: str) -> str:
            normalized = _normalize_column_name(column_name)
            if any(token in normalized for token in ("date", "time", "timestamp")):
                return "Date Range"
            label_map = {
                "state": "State",
                "city": "City",
                "destination": "Destination",
                "visitor": "Visitor Type",
                "tourist": "Visitor Type",
                "traveler": "Visitor Type",
                "traveller": "Visitor Type",
                "season": "Season",
                "category": "Category",
                "country": "Country",
                "region": "Region",
            }
            for alias, label in label_map.items():
                if alias in normalized:
                    return label
            return str(column_name)

        def sort_priority(column_name: str) -> tuple[int, str]:
            normalized = _normalize_column_name(column_name)
            priority = [
                "state",
                "city",
                "destination",
                "visitor",
                "tourist",
                "traveller",
                "traveler",
                "season",
                "category",
                "country",
                "region",
            ]
            for index, token in enumerate(priority):
                if token in normalized:
                    return (index, str(column_name).lower())
            if any(token in normalized for token in ("date", "time", "timestamp")):
                return (100, str(column_name).lower())
            return (200, str(column_name).lower())

        working_dataframe = dataframe.copy()
        ordered_columns = sorted(dataframe.columns, key=sort_priority)
        boolean_columns = set(_detect_boolean_columns(dataframe))
        numeric_columns = set(_detect_numeric_columns(dataframe))

        for column in ordered_columns:
            series = dataframe[column]
            widget_key = f"dataset_filter_{file_name}_{str(column)}"

            if column in boolean_columns:
                boolean_values = series.dropna().map(
                    lambda value: value
                    if isinstance(value, bool)
                    else str(value).strip().casefold() in {"true", "yes", "1"}
                )
                choices = [("All", None)]
                if bool(boolean_values.eq(True).any()):
                    choices.append(("Yes", True))
                if bool(boolean_values.eq(False).any()):
                    choices.append(("No", False))
                saved_value = st.session_state[filter_key].get(column)
                selected_index = next(
                    (index for index, (_, value) in enumerate(choices) if value == saved_value),
                    0,
                )
                selected_label = st.selectbox(
                    filter_label(column),
                    [label for label, _ in choices],
                    index=selected_index,
                    key=widget_key,
                )
                selected_value = dict(choices)[selected_label]
                if selected_value is not None:
                    boolean_mask = series.map(
                        lambda value: (
                            value if isinstance(value, bool)
                            else str(value).strip().casefold() in {"true", "yes", "1"}
                        )
                        if pd.notna(value)
                        else False
                    )
                    working_dataframe = working_dataframe.loc[
                        boolean_mask.reindex(working_dataframe.index, fill_value=False)
                        .eq(selected_value)
                    ]
                st.session_state[filter_key][column] = selected_value
                continue

            if pd.api.types.is_datetime64_any_dtype(series) or not pd.api.types.is_numeric_dtype(series):
                try:
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore", UserWarning)
                        parsed_dates = pd.to_datetime(series, errors="coerce")
                except (TypeError, ValueError):
                    parsed_dates = None
                non_missing_count = int(series.notna().sum())
                parsed_count = int(parsed_dates.notna().sum()) if parsed_dates is not None else 0
                if parsed_dates is not None and non_missing_count and parsed_count / non_missing_count >= 0.8:
                    date_values = parsed_dates.dropna()
                    if not date_values.empty:
                        start_date = date_values.min().date()
                        end_date = date_values.max().date()
                        saved_values = st.session_state[filter_key].get(column, (start_date, end_date))
                        if not isinstance(saved_values, (tuple, list)) or len(saved_values) != 2:
                            saved_values = (start_date, end_date)
                        selected_start, selected_end = saved_values
                        selected_dates = st.date_input(
                            filter_label(column),
                            value=(selected_start, selected_end),
                            min_value=start_date,
                            max_value=end_date,
                            key=widget_key,
                        )
                        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
                            selected_start, selected_end = selected_dates
                        else:
                            selected_start = selected_end = selected_dates
                        if selected_start is not None and selected_end is not None:
                            date_mask = pd.to_datetime(working_dataframe[column], errors="coerce").dt.date.between(
                                selected_start,
                                selected_end,
                            )
                            working_dataframe = working_dataframe.loc[date_mask.reindex(working_dataframe.index, fill_value=False)]
                            st.session_state[filter_key][column] = (selected_start, selected_end)
                        continue

            if column in numeric_columns:
                numeric_series = pd.to_numeric(series, errors="coerce")
                valid_values = numeric_series.dropna()
                valid_values = valid_values[valid_values.map(lambda value: math.isfinite(float(value)))]
                if valid_values.empty:
                    continue

                minimum = float(valid_values.min())
                maximum = float(valid_values.max())
                if minimum == maximum:
                    continue

                saved_range = st.session_state[filter_key].get(column, (minimum, maximum))
                if not isinstance(saved_range, (tuple, list)) or len(saved_range) != 2:
                    saved_range = (minimum, maximum)
                selected_range = st.slider(
                    filter_label(column),
                    min_value=minimum,
                    max_value=maximum,
                    value=(float(saved_range[0]), float(saved_range[1])),
                    key=widget_key,
                )
                numeric_mask = pd.to_numeric(working_dataframe[column], errors="coerce").between(
                    selected_range[0],
                    selected_range[1],
                )
                working_dataframe = working_dataframe.loc[numeric_mask.reindex(working_dataframe.index, fill_value=False)]
                st.session_state[filter_key][column] = selected_range
                continue

            try:
                options = working_dataframe[column].dropna().astype(str).str.strip().drop_duplicates().tolist()
                if not options:
                    continue
                saved_values = st.session_state[filter_key].get(column, [])
                if isinstance(saved_values, str):
                    saved_values = [saved_values]
                selected_values = st.multiselect(
                    filter_label(column),
                    options=options,
                    default=saved_values,
                    key=widget_key,
                )
                if selected_values:
                    category_mask = working_dataframe[column].astype(str).str.strip().isin(selected_values)
                    working_dataframe = working_dataframe.loc[category_mask]
                st.session_state[filter_key][column] = selected_values
            except (TypeError, ValueError):
                continue

        st.markdown("<div style='margin-top: 0.5rem; margin-bottom: 0.5rem;'></div>", unsafe_allow_html=True)
        st.caption(f"Showing {len(working_dataframe):,} of {len(dataframe):,} records")
        st.markdown("---")

        if st.button("🔄 Reset Filters", key=f"reset_filters_{file_name}", use_container_width=True):
            st.session_state[filter_key] = {}
            st.rerun()

    return working_dataframe


def _create_csv_download(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8")


def _create_excel_download(dataframe: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Filtered Data")
    return output.getvalue()


def _get_applied_filter_summary(dataframe: pd.DataFrame, file_name: str) -> list[str]:
    filters = st.session_state.get(f"sidebar_filters_{file_name}", {})
    applied: list[str] = []
    for column, selection in filters.items():
        if column not in dataframe.columns:
            continue
        if isinstance(selection, tuple) and len(selection) == 2:
            series = dataframe[column]
            if pd.api.types.is_numeric_dtype(series):
                values = pd.to_numeric(series, errors="coerce").dropna()
                if values.empty:
                    continue
                full_range = (float(values.min()), float(values.max()))
                if (float(selection[0]), float(selection[1])) != full_range:
                    applied.append(f"{column}: {selection[0]} to {selection[1]}")
            else:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)
                    dates = pd.to_datetime(series, errors="coerce")
                date_values = dates.dropna()
                if date_values.empty:
                    continue
                full_range = (date_values.min().date(), date_values.max().date())
                if (selection[0], selection[1]) != full_range:
                    applied.append(f"{column}: {selection[0]} to {selection[1]}")
        elif isinstance(selection, str):
            applied.append(f"{column}: {selection}")
        elif selection:
            selected_values = ", ".join(str(value) for value in selection)
            applied.append(f"{column}: {selected_values}")
    return applied


def _create_report_download(
    dataframe: pd.DataFrame,
    file_name: str,
    kpis: dict[str, str],
    insights: list[str],
    applied_filters: list[str],
) -> bytes:
    def list_items(items: list[str], empty_message: str) -> str:
        content = items or [empty_message]
        return "".join(f"<li>{html.escape(str(item))}</li>" for item in content)

    try:
        summary_html = dataframe.describe(include="all").transpose().to_html(
            classes="summary-table",
            border=0,
            na_rep="—",
            escape=True,
        )
    except (TypeError, ValueError):
        summary_html = "<p>Summary statistics are unavailable for these columns.</p>"

    kpi_items = [f"{label}: {value}" for label, value in kpis.items()]
    report_html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI SQL Data Assistant Report</title>
<style>
body {{ font-family: Arial, sans-serif; color: #243447; margin: 36px auto; max-width: 1000px; padding: 0 24px; }}
h1 {{ color: #123962; margin-bottom: 6px; }}
h2 {{ color: #234c73; margin-top: 28px; border-bottom: 1px solid #dbe5ef; padding-bottom: 8px; }}
.metadata {{ color: #5f6f87; }}
.summary-table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
.summary-table th, .summary-table td {{ border: 1px solid #dbe5ef; padding: 7px; text-align: left; }}
.summary-table th {{ background: #f0f5fa; }}
li {{ margin: 7px 0; }}
</style>
</head>
<body>
<h1>AI SQL Data Assistant Report</h1>
<p class="metadata">Dataset: {html.escape(file_name)}<br>
Filtered Records: {len(dataframe):,}<br>
Columns: {len(dataframe.columns):,}</p>
<h2>KPIs</h2><ul>{list_items(kpi_items, "No KPI values available.")}</ul>
<h2>AI Insights</h2><ul>{list_items(insights, "No insights available.")}</ul>
<h2>Applied Filters</h2><ul>{list_items(applied_filters, "No filters applied.")}</ul>
<h2>Data Summary</h2>{summary_html}
</body>
</html>"""
    return report_html.encode("utf-8")


def _remember_export_chart(figure, file_name: str) -> None:
    chart_key = f"export_chart_{file_name}"
    signature_key = f"export_chart_signature_{file_name}"
    signature = figure.to_json()
    st.session_state[chart_key] = figure
    st.session_state[signature_key] = signature


def _show_export_options(
    dataframe: pd.DataFrame,
    file_name: str,
    kpis: dict[str, str],
    insights: list[str],
) -> None:
    st.subheader("📤 Export Options")
    if dataframe.empty:
        st.info("The current filters returned no records. Export downloads are unavailable.")

    disabled = dataframe.empty
    applied_filters = _get_applied_filter_summary(dataframe, file_name)
    csv_data = _create_csv_download(dataframe) if not disabled else b""
    excel_data = _create_excel_download(dataframe) if not disabled else b""
    report_data = (
        _create_report_download(dataframe, file_name, kpis, insights, applied_filters)
        if not disabled else b""
    )

    columns = st.columns(4)
    columns[0].download_button(
        "📥 Download CSV",
        data=csv_data,
        file_name="filtered_data.csv",
        mime="text/csv",
        disabled=disabled,
        use_container_width=True,
    )
    columns[1].download_button(
        "📥 Download Excel",
        data=excel_data,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=disabled,
        use_container_width=True,
    )
    columns[2].download_button(
        "📄 Download Report",
        data=report_data,
        file_name="AI_SQL_Data_Assistant_Report.html",
        mime="text/html",
        disabled=disabled,
        use_container_width=True,
    )

    chart_key = f"export_chart_{file_name}"
    png_key = f"export_chart_png_{file_name}"
    png_signature_key = f"export_chart_png_signature_{file_name}"
    figure = st.session_state.get(chart_key)
    chart_png = st.session_state.get(png_key)
    current_signature = st.session_state.get(f"export_chart_signature_{file_name}")
    if chart_png and st.session_state.get(png_signature_key) == current_signature:
        columns[3].download_button(
            "📊 Download Chart",
            data=chart_png,
            file_name="chart.png",
            mime="image/png",
            use_container_width=True,
            key=f"download_chart_{file_name}",
        )
        return

    if not columns[3].button(
        "📊 Download Chart",
        disabled=disabled or figure is None,
        use_container_width=True,
        key=f"prepare_chart_download_{file_name}",
    ):
        return

    try:
        st.session_state[png_key] = figure.to_image(format="png")
        st.session_state[png_signature_key] = current_signature
        st.rerun()
    except (ImportError, ValueError) as exc:
        st.warning(
            "Chart image export requires Plotly's Kaleido package. Install it in the "
            f"Streamlit Python environment to enable PNG downloads. Details: {exc}"
        )
    except Exception as exc:
        st.warning(f"Unable to export the chart image: {exc}")


def _show_uploaded_file_dashboard(dataframe: pd.DataFrame, file_name: str, connected: bool) -> None:
    filtered_dataframe = _apply_sidebar_filters(dataframe, file_name)
    st.session_state.filtered_uploaded_dataframe = filtered_dataframe
    st.session_state[f"export_chart_{file_name}"] = None
    total_rows = len(filtered_dataframe)
    total_columns = len(dataframe.columns)
    missing_values = int(filtered_dataframe.isna().sum().sum())
    numeric_columns = len(dataframe.select_dtypes(include="number").columns)

    business_metrics = _detect_business_metrics(filtered_dataframe)
    preferred_numeric_columns = business_metrics["revenue"] or business_metrics["visitors"]
    main_numeric_column = None
    main_numeric_values = None
    for column in preferred_numeric_columns:
        parsed_values = pd.to_numeric(filtered_dataframe[column], errors="coerce")
        if parsed_values.notna().any():
            main_numeric_column = column
            main_numeric_values = parsed_values
            break
    if main_numeric_column is None:
        main_numeric_column, main_numeric_values = _pick_main_numeric_column(filtered_dataframe)
    total_numeric_value = float(main_numeric_values.sum()) if main_numeric_values is not None else None

    rating_column = _pick_rating_column(filtered_dataframe)
    average_rating = None
    if rating_column is not None:
        rating_values = pd.to_numeric(filtered_dataframe[rating_column], errors="coerce").dropna()
        if not rating_values.empty:
            average_rating = float(rating_values.mean())

    top_category_column = _pick_top_category_column(filtered_dataframe)
    top_category_value = "N/A"
    if top_category_column is not None:
        category_values = filtered_dataframe[top_category_column].dropna().astype(str).str.strip()
        if not category_values.empty:
            top_category_value = category_values.value_counts().idxmax()

    with st.container(border=True):
        status_columns = st.columns([2, 1, 1])
        status_columns[0].markdown(f"📁 **Dataset:** {file_name}")
        status_columns[1].markdown("🟢 **Connected**" if connected else "🔴 **Disconnected**")
        status_columns[2].markdown(f"📊 **{total_rows:,} Records**")

    profile = _generate_dataset_profile(dataframe, file_name)
    file_type_label = {
        "csv": "CSV",
        "excel": "Excel",
        "json": "JSON",
        "parquet": "Parquet",
    }.get(str(profile["file_type"]), str(profile["file_type"]).upper())
    with st.container(border=True):
        st.subheader("🔍 Dataset Auto Detection")
        st.success("Dataset detected successfully")
        profile_columns = st.columns(4)
        profile_columns[0].metric("File Type", file_type_label)
        profile_columns[1].metric("Rows", f"{profile['rows']:,}")
        profile_columns[2].metric("Columns", f"{profile['columns_count']:,}")
        profile_columns[3].metric("Missing Values", f"{profile['missing_values']:,}")
        type_columns = st.columns(4)
        type_columns[0].metric("Numeric", len(profile["numeric_columns"]))
        type_columns[1].metric("Categorical", len(profile["categorical_columns"]))
        type_columns[2].metric("Dates", len(profile["date_columns"]))
        type_columns[3].metric("Duplicates", f"{profile['duplicate_rows']:,}")
        with st.expander("Detected columns and data types"):
            st.write("**Columns and data types:**")
            for column in dataframe.columns:
                st.write(f"- `{column}`: `{dataframe[column].dtype}`")
            st.write(f"**Numeric:** {', '.join(profile['numeric_columns']) or 'None'}")
            st.write(f"**Categorical:** {', '.join(profile['categorical_columns']) or 'None'}")
            st.write(f"**Date / datetime:** {', '.join(profile['date_columns']) or 'None'}")
            st.write(f"**Boolean:** {', '.join(profile['boolean_columns']) or 'None'}")
            st.write(f"**ID / index-like:** {', '.join(profile['id_columns']) or 'None'}")
            st.write(f"**Business metrics:** {json.dumps(profile['business_metrics'], ensure_ascii=False)}")
            for date_column in profile["date_columns"]:
                parsed_dates = pd.to_datetime(dataframe[date_column], errors="coerce").dropna()
                if not parsed_dates.empty:
                    st.write(
                        f"**{date_column} range:** {parsed_dates.min().date()} to "
                        f"{parsed_dates.max().date()}"
                    )

    chart_suggestions = _generate_chart_suggestions(filtered_dataframe)
    if chart_suggestions:
        st.markdown("#### 📊 Recommended Visualizations")
        suggestion_columns = st.columns(min(2, len(chart_suggestions)))
        for index, suggestion in enumerate(chart_suggestions):
            with suggestion_columns[index % len(suggestion_columns)]:
                st.info(suggestion)

    _show_smart_ai_chart_generator(filtered_dataframe, file_name)

    st.subheader("📊 Dashboard Overview")
    with st.container(border=True):
        metric_columns = st.columns(4)
        metric_columns[0].markdown(
            f"<div class='kpi-card'><div class='kpi-label'>📊 Total Records</div><div class='kpi-value'>{total_rows:,}</div></div>",
            unsafe_allow_html=True,
        )
        metric_columns[1].markdown(
            f"<div class='kpi-card'><div class='kpi-label'>💰 Total Numeric Value</div><div class='kpi-value'>{_format_metric_value(total_numeric_value) if total_numeric_value is not None else 'N/A'}</div></div>",
            unsafe_allow_html=True,
        )
        metric_columns[2].markdown(
            f"<div class='kpi-card'><div class='kpi-label'>⭐ Average Rating</div><div class='kpi-value'>{_format_metric_value(average_rating) if average_rating is not None else 'N/A'}</div></div>",
            unsafe_allow_html=True,
        )
        metric_columns[3].markdown(
            f"<div class='kpi-card'><div class='kpi-label'>🏷️ Top Category / Destination</div><div class='kpi-value'>{str(top_category_value)}</div></div>",
            unsafe_allow_html=True,
        )

    insights = _generate_ai_insights(filtered_dataframe)
    with st.container(border=True):
        st.subheader("💡 AI Insights")
        for insight in insights:
            st.markdown(f"- {insight}")

    with st.container(border=True):
        st.subheader("📈 Analytics")
        numeric_columns_list = filtered_dataframe.select_dtypes(include="number").columns.tolist()
        category_columns_list = filtered_dataframe.select_dtypes(exclude="number").columns.tolist()
        if category_columns_list and numeric_columns_list and not filtered_dataframe.empty:
            chart_type = st.selectbox("Chart type", ["Bar Chart", "Line Chart", "Pie Chart"], key="dataset_chart_type")
            x_column = st.selectbox("X-axis", category_columns_list, key="dataset_chart_x")
            y_column = st.selectbox("Y-axis", numeric_columns_list, key="dataset_chart_y")

            try:
                title = f"{y_column} by {x_column}"
                if chart_type == "Bar Chart":
                    fig = px.bar(filtered_dataframe, x=x_column, y=y_column, title=title)
                elif chart_type == "Line Chart":
                    fig = px.line(filtered_dataframe, x=x_column, y=y_column, title=title, markers=True)
                else:
                    fig = px.pie(filtered_dataframe, names=x_column, values=y_column, title=title)
                _style_plotly_figure(fig, title, str(x_column), str(y_column))
                _remember_export_chart(fig, file_name)
                st.plotly_chart(fig, use_container_width=True, config={"responsive": True, "displayModeBar": False})
            except Exception as exc:
                st.warning(f"Chart generation failed: {exc}")
        else:
            st.info("📊 No chart generated yet. Use the Chart Request below to generate a visualization.")

        _show_ai_generated_chart(filtered_dataframe, file_name)
        smart_figure = st.session_state.get(f"smart_ai_chart_figure_{file_name}")
        if smart_figure is not None and not filtered_dataframe.empty:
            _remember_export_chart(smart_figure, file_name)

    with st.container(border=True):
        st.subheader("📋 Data Preview")
        st.markdown("<div class='summary-box'>", unsafe_allow_html=True)
        st.write("Dataset Summary")
        st.write(
            f"Filtered rows: {total_rows:,} of {len(dataframe):,} | Columns: {total_columns:,} | "
            f"Missing values: {missing_values:,} | Numeric columns: {numeric_columns:,}"
        )
        if not filtered_dataframe.empty:
            summary_table = filtered_dataframe.describe(include="all").T.head()
            for column in summary_table.select_dtypes(include="object").columns:
                summary_table[column] = summary_table[column].astype("string")
            st.write(summary_table)
        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("First 100 rows", expanded=True):
            st.dataframe(filtered_dataframe.head(100), use_container_width=True, hide_index=True)

    with st.container(border=True):
        _show_export_options(
            filtered_dataframe,
            file_name,
            {
                "Total Records": f"{total_rows:,}",
                "Total Numeric Value": _format_metric_value(total_numeric_value) if total_numeric_value is not None else "N/A",
                "Average Rating": _format_metric_value(average_rating) if average_rating is not None else "N/A",
                "Top Category / Destination": str(top_category_value),
            },
            insights,
        )

    with st.expander("🛠️ Developer / Runtime Debug", expanded=False):
        debug_info = _get_debug_info()
        st.code(
            f"Python executable: {debug_info['python_executable']}\n"
            f"Python version: {debug_info['python_version']}\n"
            f"Pandas version: {debug_info['pandas_version']}\n"
            f"Openpyxl version: {debug_info['openpyxl_version']}",
            language="text",
        )


# --- END NEW CODE: Universal File Upload Dashboard ---


def _chart_name_tokens(value: object) -> list[str]:
    name = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(value))
    tokens = re.findall(r"[a-z0-9]+", name.lower())
    normalized = []
    for token in tokens:
        if token.endswith("ies") and len(token) > 4:
            token = token[:-3] + "y"
        elif token.endswith("s") and not token.endswith(("ss", "us")) and len(token) > 3:
            token = token[:-1]
        normalized.append(token)
    return normalized


def _chart_alias_tokens(value: object) -> set[str]:
    text = _normalize_column_name(value)
    tokens = set(_chart_name_tokens(text))
    alias_map = {
        "destination": {"destination", "city", "town", "place", "location", "area"},
        "city": {"city", "town", "destination", "place", "location"},
        "state": {"state", "province", "region", "area"},
        "country": {"country", "nation", "region"},
        "category": {"category", "type", "segment", "group", "kind"},
        "customers": {"customer", "buyer", "client"},
        "tourist": {"tourist", "traveler", "traveller", "visitor", "visitors"},
        "visitors": {"visitor", "visitors", "tourist", "traveler", "traveller"},
        "revenue": {"revenue", "sales", "income", "amount", "value", "earnings"},
        "rating": {"rating", "score", "stars", "review"},
        "average": {"average", "avg", "mean", "meanvalue"},
        "count": {"count", "total", "number", "records", "rows"},
        "amount": {"amount", "total", "value", "sum"},
    }
    for key, expansions in alias_map.items():
        if key in text or any(word in text.split() for word in key.split()):
            tokens.update(expansions)
    return tokens


def _chart_column_score(column: object, question_tokens: set[str]) -> int:
    column_tokens = _chart_alias_tokens(column)
    if not column_tokens:
        return 0
    matches = column_tokens & question_tokens
    if not matches:
        return 0
    return len(matches) * 12 + (3 if set(_chart_name_tokens(column)) & question_tokens else 0)


def _chart_date_columns(dataframe: pd.DataFrame) -> dict[object, pd.Series]:
    date_columns: dict[object, pd.Series] = {}
    for column in dataframe.columns:
        series = dataframe[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            date_columns[column] = pd.to_datetime(series, errors="coerce")
        elif not pd.api.types.is_numeric_dtype(series):
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)
                    parsed = pd.to_datetime(series, errors="coerce")
            except (TypeError, ValueError):
                continue
            non_missing_count = int(series.notna().sum())
            if non_missing_count and int(parsed.notna().sum()) / non_missing_count >= 0.8:
                date_columns[column] = parsed
    return date_columns


def _chart_fallback_columns(dataframe: pd.DataFrame) -> tuple[list[str], list[str], list[str]]:
    date_columns = _chart_date_columns(dataframe)
    numeric_columns = [
        column
        for column in dataframe.select_dtypes(include="number").columns
        if not pd.api.types.is_bool_dtype(dataframe[column])
        and not any(token in {"id", "uuid", "index", "code", "pk"} for token in _chart_name_tokens(column))
    ]
    dimension_columns = [
        column for column in dataframe.columns
        if column not in numeric_columns
        and column not in date_columns
        and not any(token in {"id", "uuid", "index", "code", "pk"} for token in _chart_name_tokens(column))
    ]
    return numeric_columns, dimension_columns, list(date_columns.keys())


def _available_chart_examples(dataframe: pd.DataFrame) -> list[str]:
    examples: list[str] = []
    detected_types = _detect_column_types(dataframe)
    actual_columns = [column for column in dataframe.columns]
    numeric_columns = [
        column for column in detected_types["numeric"]
        if column in actual_columns
        and column not in detected_types["id"]
    ]
    category_columns = [
        column for column in detected_types["categorical"]
        if column in actual_columns and column not in detected_types["id"]
    ]
    date_columns = [
        column for column in detected_types["date"]
        if column in actual_columns
    ]
    category_priority = (
        "destination", "city", "category", "region", "state",
        "country", "type",
    )
    category_columns.sort(
        key=lambda column: next(
            (
                priority
                for priority, token in enumerate(category_priority)
                if token in set(_chart_name_tokens(column))
            ),
            len(category_priority),
        )
    )
    category_col = category_columns[0] if category_columns else None
    measure_col = numeric_columns[0] if numeric_columns else None

    if category_col is not None and measure_col is not None:
        examples.extend(
            [
                f"Show top 5 {category_col} by {measure_col}",
                f"Show average {measure_col} by {category_col}",
                f"Show {measure_col} by {category_col}",
            ]
        )
    elif category_col is not None:
        examples.extend(
            [
                f"Show top 10 {category_col}",
                f"Show count by {category_col}",
                f"Show {category_col} distribution",
            ]
        )
    elif measure_col is not None:
        examples.extend(
            [
                f"Show distribution of {measure_col}",
                f"Show average {measure_col}",
                f"Show total {measure_col}",
            ]
        )
    elif date_columns:
        date_column = date_columns[0]
        examples.append(f"Show records over {date_column}")
    else:
        return []
    seen: set[str] = set()
    deduped: list[str] = []
    for example in examples:
        if example.lower() not in seen:
            seen.add(example.lower())
            deduped.append(example)
    return deduped[:3]


def _chart_request_tokens(question: str) -> set[str]:
    words = set(_chart_name_tokens(question))
    stop_words = {
        "a", "an", "and", "by", "chart", "data", "display", "for", "from",
        "me", "of", "please", "show", "the", "to", "trend", "visualize", "wise",
        "all", "list", "please", "give", "top", "bottom",
    }
    return words - stop_words


def _largest_score_column(columns: list[str], question_tokens: set[str]) -> str | None:
    if not columns:
        return None
    scored = [
        (column, _chart_column_score(column, question_tokens))
        for column in columns
    ]
    scored = [item for item in scored if item[1] > 0]
    if not scored:
        return None
    return max(scored, key=lambda item: item[1])[0]


def _interpret_chart_request(
    question: str,
    dataframe: pd.DataFrame,
) -> dict[str, object] | None:
    if dataframe.empty or not len(dataframe.columns):
        return None

    question = str(question).strip()
    if not question:
        return None

    question_tokens = _chart_request_tokens(question)
    numeric_columns, dimension_columns, date_columns = _chart_fallback_columns(dataframe)
    if not numeric_columns and not dimension_columns and not date_columns:
        return None

    date_column = _largest_score_column(date_columns, question_tokens)
    dimension_column = _largest_score_column(dimension_columns, question_tokens)
    metric_column = _largest_score_column(numeric_columns, question_tokens)

    request_lower = question.lower()
    has_average_intent = any(keyword in request_lower for keyword in ("average", "avg", "mean"))
    has_total_intent = any(keyword in request_lower for keyword in ("total", "sum", "revenue", "sales", "amount"))
    has_count_intent = any(keyword in request_lower for keyword in ("count", "number", "how many", "all", "each"))
    has_explicit_measure = any(keyword in request_lower for keyword in ("by visitors", "by revenue", "by rating", "by sales", "by amount", "by total", "by count", "average", "avg", "mean", "sum", "total", "count"))
    has_top_intent = bool(re.search(r"\b(top|bottom)\s+(\d+)\b", question, flags=re.IGNORECASE))
    generic_count_request = (has_top_intent or has_count_intent) and not has_explicit_measure and not has_average_intent and not has_total_intent
    is_trend_request = any(keyword in request_lower for keyword in ("trend", "month", "monthly", "weekly", "daily", "year", "yearly", "over time"))
    is_pie_request = any(keyword in request_lower for keyword in ("pie", "share", "proportion", "percentage", "percent"))
    is_line_request = is_trend_request and bool(date_columns)

    if is_pie_request and dimension_column is not None:
        chart_type = "Pie Chart"
    elif is_line_request:
        chart_type = "Line Chart"
    else:
        chart_type = "Bar Chart"

    if generic_count_request:
        aggregation = "count"
        metric_column = None
    elif has_average_intent:
        aggregation = "mean"
    elif has_total_intent:
        aggregation = "sum"
    elif has_count_intent or (metric_column is None and dimension_column is not None):
        aggregation = "count"
    else:
        aggregation = "sum"

    if metric_column is not None and generic_count_request:
        metric_column = None

    if metric_column is None and not generic_count_request and not has_count_intent and not any(
        keyword in request_lower for keyword in ("count", "number", "all", "tourist", "tourists")
    ):
        if numeric_columns:
            metric_column = max(
                numeric_columns,
                key=lambda column: (
                    int(pd.to_numeric(dataframe[column], errors="coerce").notna().sum()),
                    not _is_likely_id_column(dataframe[column], column),
                ),
            )
            if metric_column and _chart_column_score(metric_column, question_tokens) == 0:
                metric_column = max(
                    numeric_columns,
                    key=lambda column: (
                        int(pd.to_numeric(dataframe[column], errors="coerce").notna().sum()),
                        1 if not _is_likely_id_column(dataframe[column], column) else 0,
                    ),
                )

    if dimension_column is None:
        if date_columns and is_trend_request:
            dimension_column = date_column or date_columns[0]
        elif dimension_columns:
            preferred_category_tokens = {"destination", "city", "state", "category", "region", "country", "type"}
            preferred = [
                column for column in dimension_columns
                if preferred_category_tokens & set(_chart_alias_tokens(column))
            ]
            dimension_column = preferred[0] if preferred else dimension_columns[0]

    if metric_column is None and aggregation == "count":
        metric_column = None
    elif metric_column is None and numeric_columns:
        metric_column = max(
            numeric_columns,
            key=lambda column: int(pd.to_numeric(dataframe[column], errors="coerce").notna().sum()),
        )

    if dimension_column is None and metric_column is not None and len(dimension_columns) == 1:
        dimension_column = dimension_columns[0]

    if dimension_column is None and metric_column is not None:
        return None

    if dimension_column is not None and dimension_column not in dataframe.columns:
        return None
    if metric_column is not None and metric_column not in dataframe.columns:
        return None

    top_n = None
    match = re.search(r"\b(top|bottom)\s+(\d+)\b", question, flags=re.IGNORECASE)
    if match:
        top_n = int(match.group(2))
        if match.group(1).lower() == "bottom":
            sort_ascending = True
        else:
            sort_ascending = False
    else:
        sort_ascending = bool(any(keyword in request_lower for keyword in ("lowest", "smallest", "ascending")))

    if chart_type == "Pie Chart" and dimension_column is not None and dataframe[dimension_column].nunique(dropna=True) > 8:
        chart_type = "Bar Chart"
    if top_n and top_n > 0 and chart_type in {"Bar Chart", "Line Chart"}:
        top_n = min(top_n, max(1, len(dataframe)))

    if metric_column is None and dimension_column is not None:
        aggregation = "count"

    return {
        "x": dimension_column,
        "y": metric_column,
        "aggregation": aggregation,
        "chart_type": chart_type,
        "top_n": top_n,
        "sort_ascending": sort_ascending,
        "date_columns": _chart_date_columns(dataframe),
        "question": question,
    }


def _chart_error_message(dataframe: pd.DataFrame) -> str:
    column_names = [str(column) for column in dataframe.columns]
    if not column_names:
        return (
            "I couldn't determine the chart fields because the current dataframe "
            "has no columns."
        )

    examples = _available_chart_examples(dataframe)
    example_text = (
        "; ".join(f"'{example}'" for example in examples[:3])
        if examples
        else "No chart examples can be generated because no chartable columns were detected"
    )
    return (
        f"I couldn't determine the chart fields. Available columns: {', '.join(column_names)}. "
        f"Example requests: {example_text}."
    )


def _show_ai_generated_chart(dataframe: pd.DataFrame, file_name: str) -> None:
    st.subheader("🤖 Ask AI → Generate Chart")
    st.caption("Describe the chart you want. The chart uses the current filtered dataset.")
    with st.form(f"ai_chart_form_{file_name}"):
        question = st.text_input(
            "Chart request",
            placeholder="e.g. Show me top 10 cities by visitors",
            key=f"ai_chart_question_{file_name}",
        )
        submitted = st.form_submit_button("✨ Generate AI Chart", type="primary")

    if submitted:
        st.session_state.ai_chart_request = question.strip()
        st.session_state.ai_chart_file = file_name

    saved_question = st.session_state.get("ai_chart_request", "")
    if not saved_question or st.session_state.get("ai_chart_file") != file_name:
        return

    specification = _interpret_chart_request(saved_question, dataframe)
    if specification is None:
        st.warning(_chart_error_message(dataframe))
        return

    try:
        figure, x_label, y_label, aggregation = _build_ai_chart(dataframe, specification)
    except (KeyError, TypeError, ValueError) as exc:
        st.warning(f"Unable to generate this chart: {exc}")
        return

    x_title = str(x_label)
    if aggregation == "mean":
        chart_title = f"Average {y_label} by {x_title}"
    elif aggregation == "sum":
        chart_title = f"Total {y_label} by {x_title}"
    elif aggregation == "count":
        chart_title = f"Record Count by {x_title}"
    else:
        chart_title = f"{y_label} by {x_title}"
    _style_plotly_figure(figure, chart_title, x_title, str(y_label))

    st.markdown("#### AI Generated Chart")
    _remember_export_chart(figure, file_name)
    st.plotly_chart(figure, use_container_width=True, config={"responsive": True, "displayModeBar": False})
    details = st.columns(4)
    details[0].caption(f"**X-axis:** {x_label}")
    details[1].caption(f"**Y-axis:** {y_label}")
    details[2].caption(f"**Aggregation:** {aggregation}")
    details[3].caption(f"**Chart type:** {specification['chart_type']}")


def _build_ai_chart(dataframe: pd.DataFrame, specification: dict[str, object]):
    chart_type = str(specification["chart_type"])
    x_column = specification["x"]
    y_column = specification["y"]
    title = str(specification["question"]).strip().rstrip("?.!")

    if chart_type == "Histogram":
        if y_column not in dataframe.columns:
            raise ValueError("The requested numeric column is not available.")
        values = pd.to_numeric(dataframe[y_column], errors="coerce").dropna()
        if values.empty:
            raise ValueError(f"No valid values are available in '{y_column}'.")
        return px.histogram(
            pd.DataFrame({str(y_column): values}),
            x=str(y_column),
            title=title,
        ), "—", str(y_column), "distribution"

    if chart_type == "Scatter Chart":
        if x_column not in dataframe.columns or y_column not in dataframe.columns:
            raise ValueError("The requested numeric columns are not available.")
        plotted = dataframe[[x_column, y_column]].copy()
        plotted[x_column] = pd.to_numeric(plotted[x_column], errors="coerce")
        plotted[y_column] = pd.to_numeric(plotted[y_column], errors="coerce")
        plotted = plotted.dropna()
        if plotted.empty:
            raise ValueError("No complete numeric rows are available for this chart.")
        return px.scatter(plotted, x=x_column, y=y_column, title=title), str(x_column), str(y_column), "none"

    if x_column not in dataframe.columns:
        raise ValueError(f"The requested category or date column '{x_column}' is not available.")

    working = dataframe.copy()
    date_columns = specification["date_columns"]
    if x_column in date_columns:
        working[x_column] = date_columns[x_column]
        if any(word in _chart_name_tokens(specification["question"]) for word in ("monthly", "month")):
            working[x_column] = working[x_column].dt.to_period("M").dt.to_timestamp()
        elif any(word in _chart_name_tokens(specification["question"]) for word in ("yearly", "year")):
            working[x_column] = working[x_column].dt.to_period("Y").dt.to_timestamp()

    working = working.loc[working[x_column].notna()]
    aggregation = str(specification["aggregation"])
    if y_column is None:
        grouped = working.groupby(x_column, dropna=True).size().rename("Record count").reset_index()
        y_label = "Record count"
        aggregation_label = "count"
    else:
        if y_column not in dataframe.columns:
            raise ValueError(f"The requested numeric column '{y_column}' is not available.")
        working[y_column] = pd.to_numeric(working[y_column], errors="coerce")
        working = working.dropna(subset=[y_column])
        if working.empty:
            raise ValueError(f"No valid values are available in '{y_column}'.")
        grouped = (
            working.groupby(x_column, dropna=True)[y_column]
            .agg(aggregation)
            .reset_index()
        )
        y_label = str(y_column)
        aggregation_label = aggregation

    if grouped.empty:
        raise ValueError("There are no non-missing rows to plot after applying the filters.")

    if chart_type == "Pie Chart":
        if grouped[x_column].nunique() > 8:
            raise ValueError("Pie charts need eight or fewer categories; try a bar chart request.")
        figure = px.pie(grouped, names=x_column, values=y_label, title=title)
    else:
        grouped = grouped.sort_values(y_label, ascending=bool(specification["sort_ascending"]))
        top_n = specification["top_n"]
        if top_n is not None:
            grouped = grouped.head(int(top_n))
        if chart_type == "Line Chart":
            grouped = grouped.sort_values(x_column)
            figure = px.line(grouped, x=x_column, y=y_label, title=title, markers=True)
        else:
            figure = px.bar(grouped, x=x_column, y=y_label, title=title)

    return figure, str(x_column), y_label, aggregation_label


def _show_smart_ai_chart_generator(dataframe: pd.DataFrame, file_name: str) -> None:
    request_key = f"smart_ai_chart_request_{file_name}"
    figure_key = f"smart_ai_chart_figure_{file_name}"
    st.subheader("🤖 Smart AI Chart Generator")
    st.caption("Describe a chart to generate from the current filtered dataset.")
    with st.form(f"smart_ai_chart_form_{file_name}"):
        question = st.text_input(
            "Describe the chart you want",
            placeholder="Example: Show me top 5 destinations by visitors",
            key=f"smart_ai_chart_input_{file_name}",
        )
        submitted = st.form_submit_button(
            "📊 Generate Chart",
            type="primary",
            disabled=dataframe.empty,
        )
    st.caption(
        "Examples: Show me top 5 destinations by visitors · "
        "Show average rating by city · Show revenue by state · "
        "Show visitor trend over time"
    )

    if submitted:
        st.session_state[request_key] = question.strip()
        st.session_state[figure_key] = None

    saved_question = st.session_state.get(request_key, "")
    if not saved_question:
        if dataframe.empty:
            st.info("Apply filters that return records to generate a chart.")
        return
    if dataframe.empty:
        st.session_state[figure_key] = None
        st.warning("The current filters returned no records. A chart cannot be generated.")
        return

    specification = _parse_smart_chart_request(saved_question, dataframe)
    if specification is None:
        st.session_state[figure_key] = None
        matched_filter = _find_smart_chart_filter(
            dataframe,
            saved_question,
            _detect_categorical_columns(dataframe),
        )
        if matched_filter is not None:
            st.warning(
                f"I found '{matched_filter['value']}' as a data value in "
                f"'{matched_filter['column']}', but I need a metric or chart type "
                "to create a useful visualization. Try: 'Show visitors for "
                f"{matched_filter['value']}', 'Show destinations for "
                f"{matched_filter['value']}', or 'Show top 5 destinations by visitors'."
            )
        else:
            st.warning(_chart_error_message(dataframe))
        return
    referenced_columns = [
        column for column in (
            specification["x"],
            specification["y"],
            specification.get("filter_column"),
        )
        if column is not None
    ]
    if not all(_is_valid_column(dataframe, column) for column in referenced_columns):
        st.session_state[figure_key] = None
        st.warning(_chart_error_message(dataframe))
        return

    chart_dataframe = dataframe
    filter_column = specification.get("filter_column")
    filter_value = specification.get("filter_value")
    if filter_column is not None:
        matching_values = dataframe[filter_column].astype("string").str.strip().str.casefold()
        matching_mask = matching_values == str(filter_value).strip().casefold()
        if not matching_mask.any():
            st.session_state[figure_key] = None
            st.warning(
                f"The value '{filter_value}' is no longer available in "
                f"'{filter_column}' for the current filtered data."
            )
            return
        chart_dataframe = dataframe.loc[matching_mask].copy()
        specification["date_columns"] = _chart_date_columns(chart_dataframe)

    try:
        figure, x_label, y_label, aggregation = _build_ai_chart(chart_dataframe, specification)
    except (KeyError, TypeError, ValueError) as exc:
        st.session_state[figure_key] = None
        st.warning(f"Unable to generate this chart: {exc}")
        return

    _style_plotly_figure(figure, saved_question, str(x_label), str(y_label))
    st.session_state[figure_key] = figure
    _remember_export_chart(figure, file_name)
    st.plotly_chart(figure, use_container_width=True, config={"responsive": True, "displayModeBar": False})
    details = st.columns(3)
    details[0].caption(f"**X-axis:** {x_label}")
    details[1].caption(f"**Y-axis:** {y_label}")
    details[2].caption(f"**Aggregation:** {aggregation}")


def _normalize_column_name(column_name: object) -> str:
    return " ".join(str(column_name).lower().replace("_", " ").replace("-", " ").split())


def _is_likely_id_column(series: pd.Series, column_name: object) -> bool:
    normalized_name = _normalize_column_name(column_name)
    if any(token in normalized_name.split() for token in ("id", "identifier", "uuid", "code", "index", "row", "pk")):
        return True

    non_null_values = series.dropna()
    if non_null_values.empty:
        return False

    unique_ratio = non_null_values.nunique() / len(non_null_values)
    return bool(unique_ratio > 0.9)


def _pick_main_numeric_column(dataframe: pd.DataFrame) -> tuple[str | None, pd.Series | None]:
    numeric_columns = dataframe.select_dtypes(include="number").copy()
    if numeric_columns.empty:
        return None, None

    best_column = None
    best_score = (-1, -1.0, -1.0)

    for column_name in numeric_columns.columns:
        numeric_series = pd.to_numeric(dataframe[column_name], errors="coerce")
        valid_values = numeric_series.dropna()
        if valid_values.empty:
            continue
        if _is_likely_id_column(valid_values, column_name):
            continue
        if pd.api.types.is_bool_dtype(dataframe[column_name]):
            continue

        score = (
            int(valid_values.count()),
            float(valid_values.std(ddof=0)) if len(valid_values) > 1 else 0.0,
            float(valid_values.abs().mean()) if len(valid_values) > 0 else 0.0,
        )
        if score > best_score:
            best_score = score
            best_column = column_name

    if best_column is None:
        return None, None

    return best_column, pd.to_numeric(dataframe[best_column], errors="coerce")


def _pick_rating_column(dataframe: pd.DataFrame) -> str | None:
    numeric_columns = dataframe.select_dtypes(include="number")
    if numeric_columns.empty:
        return None

    best_column = None
    best_score = -1

    for column_name in numeric_columns.columns:
        numeric_series = pd.to_numeric(dataframe[column_name], errors="coerce")
        valid_values = numeric_series.dropna()
        if valid_values.empty or _is_likely_id_column(valid_values, column_name):
            continue
        if pd.api.types.is_bool_dtype(dataframe[column_name]):
            continue

        normalized_name = _normalize_column_name(column_name)
        score = 0
        if any(token in normalized_name for token in ("rate", "rating", "score", "stars", "grade", "review", "feedback", "points")):
            score += 4
        if any(token in normalized_name for token in ("avg", "average")):
            score += 2
        if valid_values.min() >= 0 and valid_values.max() <= 5:
            score += 2
        if valid_values.min() >= 0 and valid_values.max() <= 10:
            score += 1
        if len(valid_values.unique()) <= 10:
            score += 1

        if score > best_score:
            best_score = score
            best_column = column_name

    return best_column


def _pick_top_category_column(dataframe: pd.DataFrame) -> str | None:
    non_numeric_columns = dataframe.select_dtypes(exclude="number")
    if non_numeric_columns.empty:
        return None

    best_column = None
    best_score = -1

    for column_name in non_numeric_columns.columns:
        series = non_numeric_columns[column_name].dropna().astype(str).str.strip()
        if series.empty:
            continue
        if _is_likely_id_column(non_numeric_columns[column_name], column_name):
            continue

        normalized_name = _normalize_column_name(column_name)
        if any(token in normalized_name for token in ("date", "time", "timestamp")):
            continue

        unique_count = series.nunique()
        top_frequency = series.value_counts().max()
        score = top_frequency - min(unique_count, 100) * 0.1
        if "category" in normalized_name or "type" in normalized_name or "segment" in normalized_name:
            score += 3
        if "destination" in normalized_name or "location" in normalized_name or "city" in normalized_name:
            score += 2

        if score > best_score and unique_count > 1:
            best_score = score
            best_column = column_name

    return best_column


@st.cache_data(ttl=300, show_spinner=False)
def _generate_ai_insights(dataframe: pd.DataFrame) -> list[str]:
    if dataframe.empty or not len(dataframe.columns):
        return ["No records are available in the current filtered dataset."]

    date_columns = _chart_date_columns(dataframe)

    def column_tokens(column: object) -> set[str]:
        return set(_chart_name_tokens(column))

    def is_identifier(column: object) -> bool:
        tokens = column_tokens(column)
        identifier_tokens = {"id", "index", "serial", "visitid", "uuid", "code", "pk", "no", "number"}
        return bool(tokens & identifier_tokens)

    def metric_columns(tokens: set[str]) -> list[object]:
        matches = []
        for column in dataframe.columns:
            if is_identifier(column) or pd.api.types.is_bool_dtype(dataframe[column]):
                continue
            if not column_tokens(column) & tokens:
                continue
            values = pd.to_numeric(dataframe[column], errors="coerce")
            if values.notna().any():
                matches.append(column)
        return matches

    rating_columns = metric_columns({"rating", "score", "review", "star"})
    visitor_columns = metric_columns({"visitor", "tourist", "count", "quantity", "booking"})
    revenue_columns = metric_columns({"revenue", "sale", "income", "spending", "spend", "amount", "price"})

    category_tokens = {"state", "city", "destination", "category", "visitor", "tourist", "season", "type"}
    category_priority = ("destination", "visitor", "tourist", "season", "city", "state", "category", "type")
    categorical_columns = [
        column
        for column in dataframe.columns
        if column not in date_columns
        and not is_identifier(column)
        and not pd.api.types.is_numeric_dtype(dataframe[column])
        and column_tokens(column) & category_tokens
    ]
    categorical_columns.sort(
        key=lambda column: min(
            (index for index, token in enumerate(category_priority) if token in column_tokens(column)),
            default=len(category_priority),
        )
    )

    # Include other non-ID text columns as categories for datasets with generic headers.
    for column in dataframe.columns:
        if (
            column not in date_columns
            and column not in categorical_columns
            and not is_identifier(column)
            and not pd.api.types.is_numeric_dtype(dataframe[column])
        ):
            non_missing = dataframe[column].dropna()
            unique_count = non_missing.nunique()
            if 1 < unique_count <= min(100, max(2, len(dataframe) // 2)):
                categorical_columns.append(column)

    insights = [f"The filtered dataset contains {len(dataframe):,} records."]

    def add_insight(text: str) -> None:
        normalized = re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()
        if len(insights) < 5 and all(
            re.sub(r"[^a-z0-9]+", " ", existing.lower()).strip() != normalized
            for existing in insights
        ):
            insights.append(text)

    category_insight_count = 0
    for column in categorical_columns:
        if category_insight_count >= 2:
            break
        values = dataframe[column].dropna().astype(str).str.strip()
        counts = values[values.ne("")].value_counts()
        if counts.empty:
            continue
        label = _normalize_column_name(column)
        value = counts.index[0]
        count = int(counts.iloc[0])
        record_label = "record" if count == 1 else "records"
        if "destination" in label:
            statement = f"{value} is the most frequent Destination with {count:,} {record_label}."
        elif "visitor" in label or "tourist" in label:
            statement = f"{value} is the largest visitor segment with {count:,} {record_label}."
        elif "season" in label:
            statement = f"{value} has the highest number of records by {column} with {count:,} {record_label}."
        else:
            statement = f"{value} has the highest number of records in {column} with {count:,} {record_label}."
        add_insight(statement)
        category_insight_count += 1

    if rating_columns:
        rating_values = pd.to_numeric(dataframe[rating_columns[0]], errors="coerce").dropna()
        if not rating_values.empty:
            add_insight(f"The average rating is {_format_metric_value(float(rating_values.mean()))}.")

    if visitor_columns:
        visitor_values = pd.to_numeric(dataframe[visitor_columns[0]], errors="coerce").dropna()
        if not visitor_values.empty:
            add_insight(f"Total visitors are {_format_metric_value(float(visitor_values.sum()))}.")

    if revenue_columns:
        revenue_values = pd.to_numeric(dataframe[revenue_columns[0]], errors="coerce").dropna()
        if not revenue_values.empty:
            add_insight(
                f"Average {revenue_columns[0]} is "
                f"{_format_metric_value(float(revenue_values.mean()))}."
            )

    return insights


def _format_metric_value(value: object) -> str:
    if pd.isna(value):
        return "N/A"
    if isinstance(value, (int, float)):
        if abs(float(value)) >= 1000:
            return f"{float(value):,.0f}"
        return f"{float(value):,.2f}".rstrip("0").rstrip(".")
    return str(value)


def _display_kpis(result: pd.DataFrame) -> None:
    filtered_result = result.copy()
    main_numeric_column, numeric_series = _pick_main_numeric_column(filtered_result)
    total_numeric_value = float(numeric_series.sum()) if numeric_series is not None else None

    rating_column = _pick_rating_column(filtered_result)
    average_rating = None
    if rating_column is not None:
        rating_values = pd.to_numeric(filtered_result[rating_column], errors="coerce").dropna()
        if not rating_values.empty:
            average_rating = float(rating_values.mean())

    top_category_column = _pick_top_category_column(filtered_result)
    top_category_value = "N/A"
    if top_category_column is not None:
        category_values = filtered_result[top_category_column].dropna().astype(str).str.strip()
        if not category_values.empty:
            top_category_value = category_values.value_counts().idxmax()

    st.caption("Query overview")
    columns = st.columns(4)
    columns[0].metric("📊 Total Records", f"{len(filtered_result):,}")
    columns[1].metric(
        "💰 Total Numeric Value",
        _format_metric_value(total_numeric_value) if total_numeric_value is not None else "N/A",
    )
    columns[2].metric(
        "⭐ Average Rating",
        _format_metric_value(average_rating) if average_rating is not None else "N/A",
    )
    columns[3].metric("🏷️ Top Category / Destination", str(top_category_value))


def _display_turns() -> None:
    for turn_index, turn in enumerate(st.session_state.turns):
        with st.chat_message("user"):
            st.markdown(turn["question"])
        with st.chat_message("assistant"):
            st.markdown(turn["answer"])
            if turn.get("sql"):
                with st.expander("Generated SQL"):
                    st.code(turn["sql"], language="sql")
            result = turn.get("result")
            if result is not None:
                if isinstance(result, pd.DataFrame):
                    _display_kpis(result)
                with st.expander("Query Result", expanded=True):
                    st.dataframe(result, use_container_width=True, hide_index=True)
                    _chart_controls(result, turn_index)


def _chart_controls(result: pd.DataFrame, turn_index: int) -> None:
    numeric_columns = result.select_dtypes(include="number").columns.tolist()
    category_columns = result.select_dtypes(exclude="number").columns.tolist()
    if not category_columns or not numeric_columns:
        return

    if result.empty:
        st.info("The query returned no rows to visualize.")
        return

    selected_chart = st.selectbox(
        "Chart type",
        ("Bar Chart", "Line Chart", "Pie Chart"),
        key=f"chart_type_{turn_index}",
    )
    x_column = st.selectbox(
        "X-axis column",
        category_columns,
        key=f"chart_x_{turn_index}",
    )
    y_column = st.selectbox(
        "Y-axis column",
        numeric_columns,
        key=f"chart_y_{turn_index}",
    )

    try:
        if selected_chart == "Pie Chart":
            title = f"{y_column} by {x_column}"
            figure = px.pie(result, names=x_column, values=y_column, title=title)
            _style_plotly_figure(figure, title, str(x_column), str(y_column))
            st.plotly_chart(
                figure,
                use_container_width=True,
                config={"responsive": True, "displayModeBar": False},
            )
        elif selected_chart == "Line Chart":
            st.line_chart(result, x=x_column, y=y_column, use_container_width=True)
        else:
            st.bar_chart(result, x=x_column, y=y_column, use_container_width=True)
    except (KeyError, StreamlitAPIException, TypeError, ValueError) as exc:
        st.warning(f"Unable to visualize these columns: {exc}")


def _sqlite_column_type(series: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(series) or pd.api.types.is_integer_dtype(series):
        return "INTEGER"
    if pd.api.types.is_numeric_dtype(series):
        return "REAL"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "TEXT"
    return "TEXT"


def _prepare_uploaded_dataframe(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[str, pd.Series]]]:
    sqlite_dataframe = dataframe.copy()
    unique_columns: list[str] = []
    used_names: set[str] = set()
    for column in sqlite_dataframe.columns:
        base_name = str(column).replace("\x00", "")
        candidate = base_name
        suffix = 2
        while candidate in used_names:
            candidate = f"{base_name}_{suffix}"
            suffix += 1
        used_names.add(candidate)
        unique_columns.append(candidate)
    sqlite_dataframe.columns = unique_columns

    for column in sqlite_dataframe.columns:
        if pd.api.types.is_object_dtype(sqlite_dataframe[column]):
            sqlite_dataframe[column] = sqlite_dataframe[column].map(
                lambda value: json.dumps(value, default=str) if isinstance(value, (dict, list, tuple, set)) else value
            )

    schema_columns = [
        (str(column), dataframe.iloc[:, position])
        for position, column in enumerate(sqlite_dataframe.columns)
    ]
    return sqlite_dataframe, schema_columns


def _uploaded_dataframe_schema(schema_columns: list[tuple[str, pd.Series]]) -> str:
    column_definitions = [
        f'- "{column_name}": pandas dtype {series.dtype}, SQLite type {_sqlite_column_type(series)}'
        for column_name, series in schema_columns
    ]
    if not column_definitions:
        return 'Table "uploaded_data" has no columns.'
    return 'Table "uploaded_data" (uploaded dataset):\n' + "\n".join(column_definitions)


def _validate_uploaded_table_scope(sql: str) -> bool:
    try:
        statement = sqlglot.parse_one(sql, read="sqlite")
    except sqlglot.errors.ParseError:
        return False

    cte_names = {
        cte.alias_or_name.casefold()
        for cte in statement.find_all(exp.CTE)
        if cte.alias_or_name
    }
    if "uploaded_data" in cte_names:
        return False
    table_names = [table.name.casefold() for table in statement.find_all(exp.Table)]
    return bool(table_names) and all(
        table_name == "uploaded_data" or table_name in cte_names
        for table_name in table_names
    ) and any(table_name == "uploaded_data" for table_name in table_names)


def _answer_uploaded_data_question(
    question: str,
    dataframe: pd.DataFrame,
    history: list[dict[str, str]],
    response_language: str = "en",
) -> dict[str, object]:
    if dataframe.empty or len(dataframe.columns) == 0:
        return {
            "question": question,
            "answer": "The uploaded dataset is empty or has no columns to query.",
            "sql": "",
            "result": None,
        }

    sqlite_dataframe, schema_columns = _prepare_uploaded_dataframe(dataframe)
    dataset_profile = _generate_dataset_profile(
        dataframe,
        str(st.session_state.get("uploaded_file_name") or "uploaded.csv"),
    )
    schema = (
        _uploaded_dataframe_schema(schema_columns)
        + "\n\nCurrent filtered dataset profile:\n"
        + json.dumps(dataset_profile, ensure_ascii=True, default=str)
    )
    generated = generate_sql(
        question,
        schema,
        history,
        dialect="sqlite",
        response_language=response_language,
    )
    if not generated["can_answer"]:
        explanation = generated["explanation"] or "That question cannot be answered from the uploaded dataset."
        return {"question": question, "answer": explanation, "sql": "", "result": None}

    sql = generated["sql"].strip()
    validation = validate_sql(sql, dialect="sqlite")
    if not validation["valid"] or not _validate_uploaded_table_scope(sql):
        return {
            "question": question,
            "answer": "I couldn't safely run a query against the uploaded dataset. Please try rephrasing your question.",
            "sql": sql,
            "result": None,
        }

    safe_sql = apply_row_limit(sql, dialect="sqlite")
    with sqlite3.connect(":memory:") as connection:
        sqlite_dataframe.to_sql("uploaded_data", connection, index=False, if_exists="replace")
        connection.execute("PRAGMA query_only = ON")
        result = pd.read_sql_query(safe_sql, connection)

    rows = result.to_dict(orient="records")
    answer = generate_answer(question, safe_sql, rows, history, response_language)
    return {"question": question, "answer": answer, "sql": safe_sql, "result": result}


def _answer_question(question: str, response_language: str = "en") -> dict[str, object]:
    history = st.session_state.history[-MAX_HISTORY_MESSAGES:]
    uploaded_dataframe = st.session_state.get("filtered_uploaded_dataframe")
    if not isinstance(uploaded_dataframe, pd.DataFrame):
        uploaded_dataframe = st.session_state.get("uploaded_dataframe")
    if isinstance(uploaded_dataframe, pd.DataFrame):
        return _answer_uploaded_data_question(
            question,
            uploaded_dataframe,
            history,
            response_language,
        )

    schema = get_database_schema()
    generated = generate_sql(
        question,
        schema,
        history,
        response_language=response_language,
    )

    if not generated["can_answer"]:
        explanation = generated["explanation"] or "The available database schema is not enough to answer that."
        return {"question": question, "answer": explanation, "sql": "", "result": None}

    sql = generated["sql"].strip()
    validation = validate_sql(sql)
    if not validation["valid"]:
        return {
            "question": question,
            "answer": "I couldn't safely run the generated query. Please rephrase the question and try again.",
            "sql": sql,
            "result": None,
        }

    safe_sql = apply_row_limit(sql)
    result = execute_select_query(safe_sql)
    rows = result.to_dict(orient="records")
    answer = generate_answer(question, safe_sql, rows, history, response_language)
    return {"question": question, "answer": answer, "sql": safe_sql, "result": result}


def _detect_language(text: str) -> str:
    try:
        detected = detect_langs(text.strip())
    except LangDetectException:
        return "en"
    if not detected or detected[0].prob < 0.6:
        return "en"
    language = detected[0].lang
    return language if language in SUPPORTED_LANGUAGES else "en"


def _queue_chat_question() -> None:
    question = st.session_state.get("chat_question_input", "").strip()
    if question:
        st.session_state.pending_chat_question = question
    st.session_state.chat_question_input = ""


def _transcribe_voice_recording(audio: dict[str, object]) -> str:
    recognizer = sr.Recognizer()
    audio_data = sr.AudioData(
        audio["bytes"],
        int(audio["sample_rate"]),
        int(audio["sample_width"]),
    )
    candidates: list[tuple[float, int, str]] = []
    recognition_errors: list[sr.RequestError] = []
    for locale in ("en-US", "mr-IN", "hi-IN"):
        try:
            result = recognizer.recognize_google(audio_data, language=locale, show_all=True)
        except sr.UnknownValueError:
            continue
        except sr.RequestError as exc:
            recognition_errors.append(exc)
            continue

        if not isinstance(result, dict):
            continue
        alternatives = result.get("alternative", [])
        if not isinstance(alternatives, list):
            continue
        expected_language = locale[:2]
        for alternative in alternatives:
            if not isinstance(alternative, dict):
                continue
            transcript = alternative.get("transcript")
            if not isinstance(transcript, str) or not transcript.strip():
                continue
            detected_language = _detect_language(transcript)
            confidence = alternative.get("confidence", 0.0)
            confidence_score = float(confidence) if isinstance(confidence, (int, float)) else 0.0
            language_match = int(detected_language == expected_language)
            candidates.append((confidence_score, language_match, transcript.strip()))

    if candidates:
        return max(candidates, key=lambda candidate: (candidate[0], candidate[1]))[2]
    if recognition_errors and len(recognition_errors) == 3:
        raise recognition_errors[-1]
    raise sr.UnknownValueError()


def _generate_speech_audio(text: str, language: str = "en") -> bytes:
    with tempfile.TemporaryFile(suffix=".mp3") as audio_file:
        tts_language = SUPPORTED_LANGUAGES.get(language, SUPPORTED_LANGUAGES["en"])["tts"]
        gTTS(text=text, lang=tts_language).write_to_fp(audio_file)
        audio_file.seek(0)
        return audio_file.read()


def main() -> None:
    st.set_page_config(page_title="AI SQL Data Assistant", page_icon=":speech_balloon:", layout="wide")
    st.title("🤖 AI SQL Data Assistant")
    st.caption("Explore your data, build visualizations, and ask questions in natural language.")
    _configure_dashboard_style()

    if "turns" not in st.session_state:
        st.session_state.turns = []
    if "history" not in st.session_state:
        st.session_state.history = []
    if "uploaded_dataframe" not in st.session_state:
        st.session_state.uploaded_dataframe = None
    if "uploaded_file_name" not in st.session_state:
        st.session_state.uploaded_file_name = None
    if "filtered_uploaded_dataframe" not in st.session_state:
        st.session_state.filtered_uploaded_dataframe = None
    if "pending_voice_transcript" in st.session_state:
        transcript = st.session_state.pop("pending_voice_transcript")
        st.session_state.chat_question_input = transcript
        st.session_state.pending_chat_question = transcript

    with st.sidebar:
        st.subheader("📁 Upload")
        uploaded_file = st.file_uploader(
            "Upload a dataset",
            type=["csv", "xlsx", "json", "parquet"],
            help="Supported files: CSV, Excel, JSON, Parquet",
        )
        if uploaded_file is not None:
            try:
                st.session_state.uploaded_dataframe = _load_uploaded_dataframe(uploaded_file)
                st.session_state.uploaded_file_name = uploaded_file.name
                st.success(f"Loaded {uploaded_file.name}")
            except Exception as exc:
                st.session_state.uploaded_dataframe = None
                st.session_state.uploaded_file_name = None
                st.session_state.filtered_uploaded_dataframe = None
                st.error(f"Unable to load file: {exc}")

        st.subheader("🗄️ Database")
        connected, status = _connection_status()
        if connected:
            st.success("Connected")
        else:
            st.error(status)
        if st.button("Clear Chat", icon=":material/delete:", use_container_width=True):
            st.session_state.turns = []
            st.session_state.history = []
            st.rerun()

        with st.expander("🛠️ Developer / Runtime Debug", expanded=False):
            debug_info = _get_debug_info()
            st.code(
                f"Python executable: {debug_info['python_executable']}\n"
                f"Python version: {debug_info['python_version']}\n"
                f"Pandas version: {debug_info['pandas_version']}\n"
                f"Openpyxl version: {debug_info['openpyxl_version']}",
                language="text",
            )

    if st.session_state.uploaded_dataframe is not None and st.session_state.uploaded_file_name:
        _show_uploaded_file_dashboard(
            st.session_state.uploaded_dataframe,
            st.session_state.uploaded_file_name,
            connected,
        )

    with st.container(border=True):
        st.subheader("🤖 Ask Your Data")
        _display_turns()
        example_questions = [
            "Which city has the highest number of visitors?",
            "Show top 5 destinations by visitors",
            "What is the average rating?",
            "Which state generates the highest revenue?",
        ]
        example_columns = st.columns(2)
        for index, example in enumerate(example_questions):
            if example_columns[index % 2].button(
                example,
                key=f"ask_data_example_{index}",
                use_container_width=True,
            ):
                st.session_state.pending_chat_question = example

        st.text_input(
            "Ask a question about your database",
            key="chat_question_input",
            on_change=_queue_chat_question,
        )
        st.caption("Allow microphone access in your browser. Recording uses the browser microphone.")
        try:
            voice_recording = mic_recorder(
                start_prompt="🎤",
                stop_prompt="⏹️",
                just_once=True,
                format="wav",
                key="chat_voice_input",
            )
        except (OSError, RuntimeError, StreamlitAPIException) as exc:
            st.error(f"Microphone recording is unavailable: {exc}")
            voice_recording = None
        if voice_recording is not None:
            try:
                voice_transcript = _transcribe_voice_recording(voice_recording)
            except sr.UnknownValueError:
                st.error("Speech was not recognized. Please speak clearly and try again.")
            except sr.RequestError as exc:
                st.error(f"Speech recognition service is unavailable: {exc}")
            except (OSError, ValueError, TypeError, KeyError) as exc:
                st.error(f"Unable to process the microphone recording: {exc}")
            else:
                if voice_transcript.strip():
                    st.session_state.pending_voice_transcript = voice_transcript.strip()
                    st.rerun()
                st.error("Speech was not recognized. Please speak clearly and try again.")
        st.button(
            "Ask",
            key="send_chat_question",
            type="primary",
            on_click=_queue_chat_question,
        )

        question = st.session_state.pop("pending_chat_question", None)
        if not question:
            return
        response_language = _detect_language(question)

        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            with st.spinner("Checking the schema and preparing a safe query..."):
                try:
                    turn = _answer_question(question, response_language)
                except DatabaseError as exc:
                    turn = {
                        "question": question,
                        "answer": str(exc),
                        "sql": "",
                        "result": None,
                    }
                except GroqRequestError as exc:
                    turn = {
                        "question": question,
                        "answer": str(exc),
                        "sql": "",
                        "result": None,
                    }
                except sqlite3.Error as exc:
                    turn = {
                        "question": question,
                        "answer": f"I couldn't run that query against the uploaded dataset: {exc}",
                        "sql": "",
                        "result": None,
                    }
                except (ValueError, TypeError):
                    turn = {
                        "question": question,
                        "answer": "Sorry, I couldn't interpret a safe query for that request. Please try rephrasing it.",
                        "sql": "",
                        "result": None,
                    }
                except Exception:
                    turn = {
                        "question": question,
                        "answer": "Something went wrong while processing that question. Please try again.",
                        "sql": "",
                        "result": None,
                    }

            answer_text = str(turn["answer"])
            st.markdown(answer_text)
            if answer_text.strip():
                try:
                    response_audio = _generate_speech_audio(answer_text, response_language)
                except (gTTSError, OSError, RuntimeError, ValueError) as exc:
                    st.warning(f"Unable to generate spoken audio for this response: {exc}")
                else:
                    try:
                        st.audio(response_audio, format="audio/mp3", autoplay=True)
                    except StreamlitAPIException as exc:
                        st.error(f"Unable to play the spoken response: {exc}")
                    else:
                        st.caption("If audio does not start automatically, use the audio player to play it.")
            if turn.get("sql"):
                with st.expander("Generated SQL"):
                    st.code(str(turn["sql"]), language="sql")
            result = turn.get("result")
            if isinstance(result, pd.DataFrame):
                _display_kpis(result)
                with st.expander("Query Result", expanded=True):
                    st.dataframe(result, use_container_width=True, hide_index=True)
                    _chart_controls(result, len(st.session_state.turns))

        st.session_state.turns.append(turn)
        st.session_state.history.extend(
            [
                {"role": "user", "content": question},
                {"role": "assistant", "content": str(turn["answer"])},
            ]
        )
        st.session_state.history = st.session_state.history[-MAX_HISTORY_MESSAGES:]


if __name__ == "__main__":
    main()
