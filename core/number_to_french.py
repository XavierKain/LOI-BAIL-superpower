"""Convert numbers to French words in uppercase."""


def number_to_french_words(number: float) -> str:
    """Convert a number to French words in uppercase.

    Examples:
        5000 -> "CINQ MILLE"
        160000 -> "CENT SOIXANTE MILLE"
    """
    n = int(round(number))

    if n == 0:
        return "ZÉRO"

    if n < 0:
        return "MOINS " + number_to_french_words(abs(n))

    ones = ["", "UN", "DEUX", "TROIS", "QUATRE", "CINQ", "SIX", "SEPT", "HUIT", "NEUF"]
    teens = [
        "DIX", "ONZE", "DOUZE", "TREIZE", "QUATORZE", "QUINZE", "SEIZE",
        "DIX-SEPT", "DIX-HUIT", "DIX-NEUF",
    ]
    tens = [
        "", "DIX", "VINGT", "TRENTE", "QUARANTE", "CINQUANTE",
        "SOIXANTE", "SOIXANTE", "QUATRE-VINGT", "QUATRE-VINGT",
    ]

    def _below_thousand(n: int) -> str:
        if n == 0:
            return ""
        if n < 10:
            return ones[n]
        if n < 20:
            return teens[n - 10]
        if n < 100:
            t, o = n // 10, n % 10
            if t == 7:
                return "SOIXANTE-" + teens[o]
            if t == 9:
                return "QUATRE-VINGT-" + teens[o]
            if t == 8 and o == 0:
                return "QUATRE-VINGTS"
            if t == 8:
                return "QUATRE-VINGT-" + ones[o]
            if o == 0:
                return tens[t]
            if o == 1:
                return tens[t] + " ET " + ones[o]
            return tens[t] + "-" + ones[o]

        h, rest = n // 100, n % 100
        if h == 1:
            result = "CENT"
        else:
            result = ones[h] + " CENT"
            if rest == 0:
                result += "S"
        if rest > 0:
            result += " " + _below_thousand(rest)
        return result

    parts = []
    if n >= 1_000_000:
        millions = n // 1_000_000
        parts.append("UN MILLION" if millions == 1 else _below_thousand(millions) + " MILLIONS")
        n %= 1_000_000
    if n >= 1000:
        thousands = n // 1000
        parts.append("MILLE" if thousands == 1 else _below_thousand(thousands) + " MILLE")
        n %= 1000
    if n > 0:
        parts.append(_below_thousand(n))
    return " ".join(parts)
