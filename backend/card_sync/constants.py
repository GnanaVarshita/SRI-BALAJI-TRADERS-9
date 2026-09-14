"""
Card Sync Engine Constants & Styling
Sri Balaji Traders Automation System
"""

import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Canonical activity normalization mapping
ACTIVITY_NORMALIZATION = {
    'FIELD DAYS': 'FD',
    'FIELD DAY': 'FD',
    'FIELDDAY': 'FD',
    'FIELDDAYS': 'FD',
    'CROP SHOW/FIELD DAYS': 'FD',
    'CROP SHOW / FIELD DAYS': 'FD',
    'VALUE DIFFERENTIATION DAY (FIELD DAYS)': 'FD',
    'VALUE DIFFERENTIATION DAY': 'FD',
    'VALUE DEMO DAYS (HARVEST DAYS)': 'HD',
    'FD': 'FD',
    'ORGANIZED FARMER MEETING/VILLAGE MEETING': 'OFM',
    'ORGANIZED FARMER MEETING / VILLAGE MEETING': 'OFM',
    'ORGANIZED FARMER MEETING': 'OFM',
    'VILLAGE MEETING': 'OFM',
    'OFM': 'OFM',
    'HARVEST DAYS': 'HD',
    'HARVEST DAY': 'HD',
    'HD': 'HD',
    'LARGE FARMER MEETING': 'LFM',
    'LARGE FARMER MEETING (RGL)': 'LFM',
    'LFM(RGL)': 'LFM',
    'LFM (RGL)': 'LFM',
    'L F M': 'LFM',
    'LFM': 'LFM',
    'MEGA FARMER MEETING': 'MFM',
    'MEGA FARMER MEETING / NEW PRODUCT LAUCHES(RGL)': 'MFM',
    'MEGA FARMER MEETING / NEW PRODUCT LAUNCHES(RGL)': 'MFM',
    'MFM(RGL)': 'MFM',
    'MFM (RGL)': 'MFM',
    'MFM': 'MFM',
    'GROUP MEETING (RGL)': 'GM_RGL',
    'G.M (RGL)': 'GM_RGL',
    'GM(RGL)': 'GM_RGL',
    'GM (RGL)': 'GM_RGL',
    'RGL': 'GM_RGL',
    'GROUP MEETING (REGULAR)': 'GM_REGULAR',
    'GROUP MEETING (RGULAR)': 'GM_REGULAR',
    'G.M (REGULAR)': 'GM_REGULAR',
    'GM(REGULAR)': 'GM_REGULAR',
    'GM (REGULAR)': 'GM_REGULAR',
    'G.M (REG)': 'GM_REGULAR',
    'GM (REG)': 'GM_REGULAR',
    'GROUP MEETING': 'GM_REGULAR',
    'GM': 'GM_REGULAR',
    'REGULAR': 'GM_REGULAR',
    'REG': 'GM_REGULAR',
    'DEMO ACTIVITY': 'DA',
    'DEMONSTRATION ACTIVITY': 'DA',
    'DA': 'DA',
    'OTHER BRANDING ACTIVITY': 'OBA',
    'OBA': 'OBA',
    'VIDEO SHOOT': 'VIDEO SHOOT',
    'VIDEO SHOOT / FARMER TESTIMONIAL': 'VIDEO SHOOT',
    'FARMER TESTIMONIAL': 'VIDEO SHOOT',
    'SKILL DEVELOPMENT TRAINING': 'SKILL DEVELOPMENT TRAINING',
    'SDT': 'SKILL DEVELOPMENT TRAINING',
    'BVC': 'BVC'
}

def clean_alphanumeric(text):
    """Strips all non-alphanumeric characters and converts to uppercase."""
    if not text:
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(text)).upper()

def normalize_activity(act):
    """
    Normalizes an activity string into its standard canonical code (e.g. 'FD', 'LFM', 'MFM').
    """
    if not act:
        return ""
    s = str(act).strip().upper()
    if s in ACTIVITY_NORMALIZATION:
        return ACTIVITY_NORMALIZATION[s]
    
    clean = clean_alphanumeric(s)
    for k, v in ACTIVITY_NORMALIZATION.items():
        if clean == clean_alphanumeric(k):
            return v

    # Fallback keyword matching
    if 'FIELD' in s or clean == 'FD':
        return 'FD'
    if 'MEGA' in s or 'MFM' in s:
        return 'MFM'
    if 'LARGE' in s or 'LFM' in s:
        return 'LFM'
    if 'HARVEST' in s or clean == 'HD':
        return 'HD'
    if 'GROUP' in s or 'G.M' in s or 'GM' in s or s in ['RGL', 'REGULAR', 'REG']:
        if ('REG' in s or 'REGULAR' in s) and 'RGL' not in s:
            return 'GM_REGULAR'
        return 'GM_RGL'
    if 'VILLAGE' in s or 'OFM' in s:
        return 'OFM'
    if 'DEMO' in s or clean == 'DA':
        return 'DA'
    if 'BVC' in s:
        return 'BVC'

    return s

def match_activity_to_column(tbm_activity, col_mapping):
    """
    Matches a TBM activity string to one of the columns in col_mapping (dict of norm_act -> col_idx).
    Returns target column index, or None if no match is found.
    """
    if not tbm_activity or not col_mapping:
        return None

    norm_target = normalize_activity(tbm_activity)
    if not norm_target:
        return None

    if norm_target in col_mapping:
        return col_mapping[norm_target]

    clean_target = clean_alphanumeric(norm_target)
    if not clean_target:
        return None

    # 1. Exact clean match
    for mapped_act, col_idx in col_mapping.items():
        if clean_target == clean_alphanumeric(mapped_act):
            return col_idx

    # 2. Substring match (require min 3 chars to prevent false positive short matches)
    if len(clean_target) >= 3:
        for mapped_act, col_idx in col_mapping.items():
            clean_mapped = clean_alphanumeric(mapped_act)
            if len(clean_mapped) >= 3:
                if clean_target in clean_mapped or clean_mapped in clean_target:
                    return col_idx

    return None

# Styling constants
REGULAR_FONT = Font(name='Calibri', size=10)
BOLD_FONT = Font(name='Calibri', size=10, bold=True)
GREEN_BOLD_FONT = Font(name='Calibri', size=11, bold=True, color='008000')
RED_BOLD_FONT = Font(name='Calibri', size=10, bold=True, color='FF0000')
STATUS_HDR_FONT = Font(name='Calibri', size=10, bold=True, color='006100')
STATUS_VAL_FONT = Font(name='Calibri', size=10, bold=True, color='006100')

THIN_SIDE = Side(style='thin', color='000000')
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

STATUS_SIDE = Side(style='thin', color='A6A6A6')
STATUS_BORDER = Border(left=STATUS_SIDE, right=STATUS_SIDE, top=STATUS_SIDE, bottom=STATUS_SIDE)

CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
