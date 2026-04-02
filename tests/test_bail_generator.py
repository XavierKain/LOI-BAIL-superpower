import pytest
from generators.bail_generator import evaluer_condition, generer_conditions_suspensives


def test_condition_empty():
    assert evaluer_condition("", {}) is True
    assert evaluer_condition(None, {}) is True


def test_condition_equals_oui():
    assert evaluer_condition("Si [Actualisation] = 'Oui'", {"Actualisation": "Oui"}) is True
    assert evaluer_condition("Si [Actualisation] = 'Oui'", {"Actualisation": "Non"}) is False


def test_condition_equals_case_insensitive():
    assert evaluer_condition("Si [Actualisation] = 'oui'", {"Actualisation": "Oui"}) is True


def test_condition_greater_than():
    assert evaluer_condition("Si [Duree Bail] > 9", {"Duree Bail": "10"}) is True
    assert evaluer_condition("Si [Duree Bail] > 9", {"Duree Bail": "9"}) is False


def test_condition_superieur_a():
    assert evaluer_condition("Si [Duree Bail] superieur a 9", {"Duree Bail": "10"}) is True


def test_condition_not_equal():
    assert evaluer_condition("Si [Type] != 'SAS'", {"Type": "SARL"}) is True
    assert evaluer_condition("Si [Type] != 'SAS'", {"Type": "SAS"}) is False


def test_condition_non_vide():
    assert evaluer_condition("Si [Loyer année 1] non vide", {"Loyer année 1": "5000"}) is True
    assert evaluer_condition("Si [Loyer année 1] non vide", {"Loyer année 1": ""}) is False
    assert evaluer_condition("Si [Loyer année 1] non vide", {}) is False


def test_condition_non_vide_zero_is_empty():
    assert evaluer_condition("Si [Val] non vide", {"Val": "0"}) is False
    assert evaluer_condition("Si [Val] non vide", {"Val": 0}) is False


def test_condition_plusieurs_suspensives():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "Extraction",
        "Condition suspensive 3": "",
        "Condition suspensive 4": "",
    }
    assert evaluer_condition("Si plusieurs conditions suspensives", donnees) is True


def test_condition_plusieurs_suspensives_only_one():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "",
    }
    assert evaluer_condition("Si plusieurs conditions suspensives", donnees) is False


def test_condition_typographic_quotes():
    assert evaluer_condition("Si [Type] = \u2018Oui\u2019", {"Type": "Oui"}) is True


def test_condition_unrecognized():
    assert evaluer_condition("Something weird", {}) is False


def test_condition_less_than():
    assert evaluer_condition("Si [Duree Bail] < 10", {"Duree Bail": "9"}) is True
    assert evaluer_condition("Si [Duree Bail] < 10", {"Duree Bail": "10"}) is False


def test_condition_gte_lte():
    assert evaluer_condition("Si [Duree Bail] >= 9", {"Duree Bail": "9"}) is True
    assert evaluer_condition("Si [Duree Bail] <= 9", {"Duree Bail": "9"}) is True


def test_conditions_suspensives_single():
    donnees = {"Condition suspensive 1": "Financement"}
    result = generer_conditions_suspensives(donnees, "Option 1 text", "Option 2 text")
    assert result == "Option 1 text"


def test_conditions_suspensives_multiple():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "Extraction",
    }
    result = generer_conditions_suspensives(
        donnees, "Option 1", "Texte avec suivantes :\n\nblabla\n\nCi-après suite"
    )
    assert "a." in result
    assert "b." in result
    assert "Obtention" in result


def test_conditions_suspensives_none():
    donnees = {}
    result = generer_conditions_suspensives(donnees, "Option 1", "Option 2")
    assert result == ""
