"""core.py — پارس اکسل/HTML کدال با تطبیق کامل حروف فارسی/عربی"""
import re
import io
import pandas as pd
import numpy as np

PERSIAN = "۰۱۲۳۴۵۶۷۸۹"
ARABIC  = "٠١٢٣٤٥٦٧٨٩"

def to_en(s):
    for i, ch in enumerate(PERSIAN): s = s.replace(ch, str(i))
    for i, ch in enumerate(ARABIC):  s = s.replace(ch, str(i))
    return s

def normalize_chars(s):
    """یکسان‌سازی کامل حروف فارسی/عربی + حذف نویسه‌های نامرئی."""
    return (s.replace("ي", "ی")     # Arabic yeh
             .replace("ى", "ی")     # Alef maksura
             .replace("ك", "ک")     # Arabic kaf
             .replace("ة", "ه")     # Teh marbuta
             .replace("\u200c", "") # ZWNJ
             .replace("\u200f", "") # RLM
             .replace("\u200e", "") # LRM
             .replace("\u0640", ""))# Tatweel

def norm_label(s):
    s = to_en(str(s))
    s = normalize_chars(s)
    return re.sub(r"\s+", "", s).strip()

def parse_num(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return np.nan
    if isinstance(x, (int, float)): return float(x)
    s = to_en(str(x)).strip()
    if not s or s in {"-", "*", ""}: return np.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace(",", "").replace("،", "").replace(" ", "").replace("%", "").strip("()")
    try:
        v = float(s); return -v if neg else v
    except: return np.nan

def safe_div(a, b):
    try:
        if b in (0, None) or pd.isna(a) or pd.isna(b): return np.nan
        return a / b
    except: return np.nan

def safe_add(*vs): return np.nan if any(pd.isna(v) for v in vs) else float(sum(vs))
def safe_mul(a, b): return np.nan if pd.isna(a) or pd.isna(b) else a * b


# ═══════════════ تشخیص و پارس فرمت ═══════════════
def parse_any_format(content_bytes):
    if not content_bytes or len(content_bytes) < 8:
        return None, "empty", ["فایل خالی/خیلی کوچک"]

    magic = content_bytes[:8]
    if magic[:2] == b'PK':                     kind =
