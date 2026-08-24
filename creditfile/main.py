# Created 2023-10-20
# Updated: removed Google Colab / ipywidgets dependencies so the pipeline
# can be used from any Python environment (web server, CLI, tests, etc.)

from .utils import isna
from .parse import get_file_details, parse_credit_report
from .normalize import normalize_credit_data
from .featurize import prepare_features
from .score import make_credit_score

from collections.abc import Iterable
import io
import json
import pandas as pd


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ESSENTIAL_FIELDS = {
    'personal_data': [
        'name',
        'present_address',
        'present_address_tenure',
        'contact_no',
        'birthplace',
        'education',
        'parents_name',
        'parents_address',
        'date_applied',
        'unit_applied',
        'loan_amount',
        'loan_terms',
        'housing_status',
        'dob',
        'age',
        'marital_status',
        'n_children',
        'n_dependents',
        'dependent_ages',
    ],
    'income_analysis': {
        'summary': ['gross_income', 'monthly_amortization']
    },
    'officer_assessment': [
        'loan_purpose',
        'unit_payor',
        'unit_rider',
        'rider_license',
        'cell_signal_status',
        'prepared_by',
        'remarks',
    ],
}

MISSING_THRESHOLD = 5  # suppress score when >= this many features are null


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

def is_missing(x):
    'Return True when x is null/NaN or an iterable whose every element is.'
    return isna(x) or (isinstance(x, Iterable) and all(isna(_) for _ in x))


def validate_fields(data, essential_fields=ESSENTIAL_FIELDS):
    'Return a dict of section → [missing field names].'
    missing_fields = {}
    for k, v in essential_fields.items():
        subset = data.get(k, None)
        missing = []
        if isna(subset):
            missing_fields[k] = v
            continue
        if isinstance(v, dict):
            missing = validate_fields(subset, v)
        elif isinstance(v, list):
            missing = [f for f in v if isna(subset.get(f, None))]
        if missing:
            missing_fields[k] = missing
    return missing_fields


# ---------------------------------------------------------------------------
# Core analysis function — no I/O side-effects, works anywhere
# ---------------------------------------------------------------------------

def analyze_file(file_obj, filename: str, tmp_path: str = None) -> dict:
    """
    Run the full credit scoring pipeline on an already-opened file object.

    Parameters
    ----------
    file_obj : file-like or BytesIO
        The raw Excel file content.
    filename : str
        Original filename (used for metadata and the output JSON name).
    tmp_path : str, optional
        If provided, the real on-disk path of the file so that the actual
        last-modified timestamp can be read via os.path.getmtime.
        When omitted the current time (Asia/Manila, UTC+8) is used instead.

    Returns
    -------
    dict with keys:
        filename, credit_score, score_valid, missing_count,
        missing_fields, normalized
    """
    if isinstance(file_obj, (bytes, bytearray)):
        file_obj = io.BytesIO(file_obj)

    # File metadata — use real mtime when a temp path is supplied
    if tmp_path:
        file_details = get_file_details(tmp_path)
        file_details['filename'] = filename          # keep the original name
    else:
        file_details = {
            'filename': filename,
            'last_modified': str(
                pd.Timestamp.now() + pd.Timedelta(hours=8)
            ),
        }

    parsed = file_details | parse_credit_report(file_obj)
    normalized = normalize_credit_data(parsed)
    features = prepare_features(normalized)

    missing_count = sum(1 for f in features if isna(f) or f == -1)
    score = make_credit_score(features)
    score_valid = missing_count < MISSING_THRESHOLD
    missing_fields = validate_fields(normalized)

    return {
        'filename': filename,
        'credit_score': score if score_valid else None,
        'score_valid': score_valid,
        'missing_count': missing_count,
        'missing_fields': missing_fields,
        'normalized': normalized,
    }


def to_json(result: dict, indent: int = 4) -> str:
    """
    Serialize an analyze_file result to JSON matching the canonical format:
      filename, last_modified, personal_data, income_source_details,
      income_analysis, officer_assessment, credit_score
    """
    n = result['normalized']
    payload = {
        'filename':             n.get('filename'),
        'last_modified':        n.get('last_modified'),
        'personal_data':        n.get('personal_data', {}),
        'income_source_details': n.get('income_source_details', {}),
        'income_analysis':      n.get('income_analysis', {}),
        'officer_assessment':   n.get('officer_assessment', {}),
        'credit_score':         result['credit_score'],
    }
    return json.dumps(payload, indent=indent)