"""
Common Data Formatting & Parsing Utilities
Sri Balaji Traders Automation System
"""

import re
import datetime

try:
    import xlrd
except ImportError:
    xlrd = None


def clean_str(val):
    """Safely converts any value to trimmed string, handling None, null, nan."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["none", "null", "nan"]:
        return ""
    return s


def parse_num(val):
    """Safely parses float or int from numbers, formatted strings with commas, etc."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date_to_obj(val):
    """Parses a date string or datetime into a datetime.date object."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if not s:
        return None
    for fmt in ["%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d\\%m\\%Y", "%d/%m/%y", "%d-%m-%y"]:
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.date()
        except ValueError:
            pass
    return None


def format_date_str(date_obj):
    """Formats date object or string to standard DD-MM-YYYY."""
    if not date_obj:
        return ""
    if isinstance(date_obj, (datetime.datetime, datetime.date)):
        return date_obj.strftime("%d-%m-%Y")
    d = parse_date_to_obj(date_obj)
    if d:
        return d.strftime("%d-%m-%Y")
    return str(date_obj).strip()


def format_date_val(val, delimiter="-"):
    """
    Normalizes string or date to DD-MM-YYYY (or DD/MM/YYYY if delimiter is '/').
    """
    if val is None:
        return ""
    target_fmt = f"%d{delimiter}%m{delimiter}%Y"
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime(target_fmt)
    s = str(val).strip()
    if not s:
        return ""
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d\\%m\\%Y", "%d/%m/%y", "%d-%m-%y"]:
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.strftime(target_fmt)
        except ValueError:
            pass
    return s


def parse_date_intelligent(val):
    """
    Intelligently parses dates from datetime, date, Excel serial numbers,
    or diverse string formats, and normalizes them to DD-MM-YYYY.
    """
    if val is None or str(val).strip() == "":
        return ""
    
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d-%m-%Y")
    
    s = str(val).strip()
    if not s:
        return ""

    # Check if numeric Excel serial date (e.g. 45447)
    try:
        num = float(s)
        if 35000 <= num <= 65000:
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(num))
            return dt.strftime("%d-%m-%Y")
    except ValueError:
        pass

    # Normalize delimiters
    s_norm = s.replace("\\\\", "/").replace("\\", "/").replace(".", "/").replace("-", "/").strip()

    # Try standard string formats
    for fmt in [
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y",
        "%d/%b/%Y", "%d/%B/%Y", "%d/%b/%y", "%d/%B/%y",
        "%b/%d/%Y", "%B/%d/%Y", "%b/%d/%y", "%B/%d/%y"
    ]:
        try:
            dt = datetime.datetime.strptime(s_norm, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            pass

    # Regex for day/month/year components
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", s_norm)
    if m:
        p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yr = p3 if p3 >= 100 else (2000 + p3)
        if p1 > 12:
            day, month = p1, p2
        elif p2 > 12:
            day, month = p2, p1
        else:
            day, month = p1, p2
        try:
            dt = datetime.datetime(yr, month, day)
            return dt.strftime("%d-%m-%Y")
        except Exception:
            pass

    return s


def format_date_xlrd(val, datemode):
    """Formats date from legacy xlrd numeric cell."""
    if val is None or val == "":
        return ""
    if isinstance(val, float) and xlrd is not None:
        try:
            dt = xlrd.xldate_as_datetime(val, datemode)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
    return format_date_val(val, delimiter="/")


def calculate_receivable_date(invoice_date_val, days_credit=45):
    """
    Calculates the receivable date by adding credit period (default 45 days) to invoice date.
    e.g. 09-04-2026 + 45 days -> 24-05-2026
    """
    d = parse_date_to_obj(invoice_date_val)
    if d:
        rec_d = d + datetime.timedelta(days=days_credit)
        return rec_d.strftime("%d-%m-%Y")
    return ""


def normalize_po(po_str):
    """Normalizes purchase order number by removing spaces, hyphens, and underscores."""
    if not po_str:
        return ""
    return re.sub(r'[\s\-_]', '', str(po_str)).upper()


def format_short_iv(invoice_num_str):
    """
    Extracts short display number: e.g. "SBT26270072" -> "72" or "SBT26270067" -> "67".
    """
    s = str(invoice_num_str).strip()
    m = re.search(r'SBT\d{4}(\d+)', s, re.I)
    if m:
        num_part = m.group(1).lstrip('0')
        return num_part if num_part else "0"
    m2 = re.search(r'(\d+)$', s)
    if m2:
        num_part = m2.group(1).lstrip('0')
        return num_part if num_part else "0"
    return s


def format_full_iv(invoice_num_str, fin_year="2627"):
    """
    Normalizes invoice number to canonical format: e.g. "72" -> "SBT26270072"
    """
    s = str(invoice_num_str).strip()
    m = re.search(r'SBT\d{4}(\d+)', s, re.I)
    if m:
        digits = int(m.group(1))
        return f"SBT{fin_year}{digits:04d}"
    digits_m = re.search(r'(\d+)', s)
    if digits_m:
        digits = int(digits_m.group(1))
        return f"SBT{fin_year}{digits:04d}"
    return s
