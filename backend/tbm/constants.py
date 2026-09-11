"""
TBM Constants & Field Keyword Mappings
Sri Balaji Traders Automation System
"""

FIELD_KEYWORDS = {
    'sl_no': ['sl no', 'sl.no', 's no', 's.no', 'sno', 'si no', 'si.no', 's. no', 'sl. no', 'slno'],
    'date': ['date'],
    'zdgm': ['zdgm', 'zdgl', 'area manager', 'adgl', 'adg', 'dm', 'zdsm', 'manager', 'rbm', 'am', 'area manager/am', 'area manager / am'],
    'tbm': ['tbm name', 'tbm', 'name of the tbm', 'tbm/sc/so', 'sc/so', 'tbm / sc / so', 'mie', 'mie name'],
    'mdo': ['mdo name', 'mdo', 'mdo/fa', 'fa'],
    'territory': ['territory', 'tbm territory', 'area', 'place', 'location'],
    'product': ['product', 'item', 'brand', 'product name'],
    'crop': ['crop', 'crops'],
    'activity': ['type of activity', 'activity name', 'activity', 'activities'],
    'village': ['village name', 'village', 'villages', 'town'],
    'farmers': ['no.of farmers', 'no of farmers', 'no. of farmers', 'n0 of farmers', 'no of rarmers', 'farmers attended', 'farmers', 'no of farmer', 'no.of farmer'],
    'tent': ['tent/hall /chairs expenses', 'tent/hall/chairs', 'tent/hall suppliers charges', 'tent/ hall', 'tent', 'chairs', 'chairs/ table/tent', 'suppliers/charges', 'suppliers charges', 'suppliers', 'hall charges', 'tent charges'],
    'food': ['food expense', 'food expenses', 'food/snacks', 'food', 'expenses', 'food expences', 'snacks', 'tiffin'],
    'transport': ['transport', 'trans port', 'auto charges', 'auto', 'travelling', 'travel'],
    'others': ['others/gifts', 'others / gifts', 'others / gift', 'others', 'gifts', 'saplaires', 'other', 'gift'],
    'total': ['total amount', 'total', 'amount'],
    'po_number': ['po number', 'po.no', 'ponumber', 'po #', 'po', 'po no']
}

STANDARD_TBM_HEADERS = [
    "SI No.", "Date", "ZDGM", "TBM", "MDO", "Territory", "Product", "Crop", "Activity", "Village",
    "No.of Farmers", "Tent/Hall Suppliers Charges", "Food Expenses", "Transport", "Others/Gifts", "Total", "PO Number"
]

STANDARD_TBM_COL_WIDTHS = {
    'A': 8,   # SI No.
    'B': 13,  # Date
    'C': 18,  # ZDGM
    'D': 18,  # TBM
    'E': 16,  # MDO
    'F': 14,  # Territory
    'G': 15,  # Product
    'H': 12,  # Crop
    'I': 14,  # Activity
    'J': 16,  # Village
    'K': 14,  # No.of Farmers
    'L': 24,  # Tent/Hall Suppliers Charges
    'M': 15,  # Food Expenses
    'N': 12,  # Transport
    'O': 20,  # Others/Gifts
    'P': 16,  # Total
    'Q': 18   # PO Number
}
