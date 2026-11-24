from edutpl.config import LANGUAGES, load_labels_table


def _keys(d, prefix=""):
    out = set()
    for k, v in d.items():
        out |= _keys(v, f"{prefix}{k}.") if isinstance(v, dict) else {f"{prefix}{k}"}
    return out


def test_every_label_exists_in_every_language():
    table = load_labels_table()
    ref = _keys(table["en"])
    for lang in LANGUAGES:
        assert _keys(table[lang]) == ref, lang


def test_sheet_names_valid_for_excel():
    table = load_labels_table()
    for lang in LANGUAGES:
        names = list(table[lang]["sheet"].values())
        assert len(set(names)) == len(names)
        assert all(len(n) <= 31 and not set(n) & set("[]:*?/\\") for n in names)
