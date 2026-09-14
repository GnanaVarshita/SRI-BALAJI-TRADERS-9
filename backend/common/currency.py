"""
Indian Currency Words Converter
Sri Balaji Traders Automation System
"""

def num_to_indian_words(amount):
    """
    Converts a number (float or int) to Indian Currency Words format matching exact business style:
    e.g. 126843 -> "One Lakh Twenty Six Thousand Eight Hundred and Forty Three Rupees Only"
    e.g. 15488 -> "Fifteen Thousand Four Hundred and Eighty Eight Rupees Only"
    e.g. 8989 -> "Eight Thousand Nine Hundred and Eighty Nine Rupees Only"
    """
    try:
        amt_float = float(amount)
    except (ValueError, TypeError):
        return ""

    if amt_float == 0:
        return "Zero Rupees Only"

    is_negative = amt_float < 0
    amt_float = abs(amt_float)

    rupees = int(amt_float)
    paise = int(round((amt_float - rupees) * 100))

    units = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
             "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
             "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def two_digits_to_words(n):
        if n == 0:
            return ""
        elif n < 20:
            return units[n]
        else:
            t = tens[n // 10]
            u = units[n % 10]
            return f"{t} {u}".strip()

    def three_digits_to_words(n):
        h = n // 100
        rem = n % 100
        words = []
        if h > 0:
            words.append(f"{units[h]} Hundred")
        if rem > 0:
            if h > 0:
                words.append(f"and {two_digits_to_words(rem)}")
            else:
                words.append(two_digits_to_words(rem))
        return " ".join(words).strip()

    parts = []
    # Crores (10,000,000)
    crores = rupees // 10000000
    rupees %= 10000000
    if crores > 0:
        parts.append(f"{two_digits_to_words(crores)} Crore")

    # Lakhs (100,000)
    lakhs = rupees // 100000
    rupees %= 100000
    if lakhs > 0:
        parts.append(f"{two_digits_to_words(lakhs)} Lakh")

    # Thousands (1,000)
    thousands = rupees // 1000
    rupees %= 1000
    if thousands > 0:
        parts.append(f"{two_digits_to_words(thousands)} Thousand")

    # Hundreds and below
    if rupees > 0:
        parts.append(three_digits_to_words(rupees))

    rupees_str = " ".join(parts).strip()
    if not rupees_str:
        rupees_str = "Zero"

    res = f"{rupees_str} Rupees"
    if paise > 0:
        res += f" and {two_digits_to_words(paise)} Paise"
    res += " Only"

    if is_negative:
        res = "Minus " + res

    return res
