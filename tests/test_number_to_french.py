from core.number_to_french import number_to_french_words


def test_zero():
    assert number_to_french_words(0) == "ZÉRO"

def test_ones():
    assert number_to_french_words(1) == "UN"
    assert number_to_french_words(5) == "CINQ"
    assert number_to_french_words(9) == "NEUF"

def test_teens():
    assert number_to_french_words(10) == "DIX"
    assert number_to_french_words(11) == "ONZE"
    assert number_to_french_words(16) == "SEIZE"
    assert number_to_french_words(19) == "DIX-NEUF"

def test_french_specific_tens():
    assert number_to_french_words(70) == "SOIXANTE-DIX"
    assert number_to_french_words(71) == "SOIXANTE-ONZE"
    assert number_to_french_words(79) == "SOIXANTE-DIX-NEUF"
    assert number_to_french_words(80) == "QUATRE-VINGTS"
    assert number_to_french_words(81) == "QUATRE-VINGT-UN"
    assert number_to_french_words(90) == "QUATRE-VINGT-DIX"
    assert number_to_french_words(99) == "QUATRE-VINGT-DIX-NEUF"

def test_et_linking():
    assert number_to_french_words(21) == "VINGT ET UN"
    assert number_to_french_words(31) == "TRENTE ET UN"
    assert number_to_french_words(61) == "SOIXANTE ET UN"

def test_hundreds():
    assert number_to_french_words(100) == "CENT"
    assert number_to_french_words(200) == "DEUX CENTS"
    assert number_to_french_words(201) == "DEUX CENT UN"

def test_thousands():
    assert number_to_french_words(1000) == "MILLE"
    assert number_to_french_words(2000) == "DEUX MILLE"
    assert number_to_french_words(40000) == "QUARANTE MILLE"

def test_real_amounts():
    assert number_to_french_words(160000) == "CENT SOIXANTE MILLE"
    assert number_to_french_words(1234) == "MILLE DEUX CENT TRENTE-QUATRE"

def test_millions():
    assert number_to_french_words(1000000) == "UN MILLION"
    assert number_to_french_words(2000000) == "DEUX MILLIONS"

def test_float_rounds():
    assert number_to_french_words(99.7) == "CENT"
    assert number_to_french_words(99.4) == "QUATRE-VINGT-DIX-NEUF"

def test_negative():
    assert number_to_french_words(-5) == "MOINS CINQ"
