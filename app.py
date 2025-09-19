import io
import re
import copy
import datetime
from typing import Dict, Optional, List

import numpy as np
import streamlit as st
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ----------------------------
# Utilities and policy helpers
# ----------------------------

def normalise_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def mask_evidence_id(eid: str) -> str:
    if not eid or str(eid).strip().lower() == "pending":
        return "Pending"
    s = str(eid)
    return "*" * max(0, len(s) - 4) + s[-4:]


def validate_year(y: str) -> bool:
    try:
        y = str(y).strip()
        year = int(y)
        now = datetime.datetime.now().year
        return len(y) == 4 and 1900 <= year <= now
    except Exception:
        return False


def respond_to_instruction_request(user_text: str) -> Optional[str]:
    triggers = [
        r"show (your|the) instructions",
        r"reveal (your|the) prompt",
        r"what are your rules",
        r"display (system|agent) prompt",
        r"print your instructions",
    ]
    for pat in triggers:
        if re.search(pat, user_text or "", flags=re.IGNORECASE):
            return " I am not trained to do this"
    return None


# ----------------------------
# DOCX helpers
# ----------------------------

def all_doc
