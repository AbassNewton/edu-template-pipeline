"""Load the template specification and the translated labels."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = ROOT / "config"
LANGUAGES = ("en", "fr", "es", "pt", "ru")


@dataclass
class Spec:
    template_version: str
    first_year: int
    last_year: int
    levels: list[str]
    sexes: list[str]
    age_min: int
    age_max: int
    age_levels: list[str]
    expenditure_items: list[str]
    thresholds: dict = field(default_factory=dict)

    @property
    def years(self) -> list[int]:
        return list(range(self.first_year, self.last_year + 1))

    @property
    def ages(self) -> list[int]:
        return list(range(self.age_min, self.age_max + 1))

    def with_years(self, first_year: int | None = None, last_year: int | None = None,
                   version: str | None = None) -> "Spec":
        """Return a copy covering another year range (used to prepare a new cycle)."""
        return Spec(
            template_version=version or self.template_version,
            first_year=first_year or self.first_year,
            last_year=last_year or self.last_year,
            levels=list(self.levels), sexes=list(self.sexes),
            age_min=self.age_min, age_max=self.age_max,
            age_levels=list(self.age_levels),
            expenditure_items=list(self.expenditure_items),
            thresholds=dict(self.thresholds),
        )


def load_spec(path: Path | str | None = None) -> Spec:
    data = yaml.safe_load(Path(path or CONFIG_DIR / "spec.yaml").read_text(encoding="utf-8"))
    return Spec(
        template_version=str(data["template_version"]),
        first_year=int(data["first_year"]),
        last_year=int(data["last_year"]),
        levels=list(data["levels"]),
        sexes=list(data["sexes"]),
        age_min=int(data["ages"]["min"]),
        age_max=int(data["ages"]["max"]),
        age_levels=list(data["age_levels"]),
        expenditure_items=list(data["expenditure_items"]),
        thresholds=dict(data.get("thresholds", {})),
    )


class Labels:
    """Dotted-key access to the labels of one language, e.g. labels('sheet.population')."""

    def __init__(self, lang: str, table: dict):
        if lang not in table:
            raise KeyError(f"language '{lang}' not in labels.yaml")
        self.lang = lang
        self._t = table[lang]

    def __call__(self, key: str) -> str:
        node = self._t
        for part in key.split("."):
            node = node[part]
        return str(node)


def load_labels_table(path: Path | str | None = None) -> dict:
    return yaml.safe_load(Path(path or CONFIG_DIR / "labels.yaml").read_text(encoding="utf-8"))


def load_labels(lang: str = "en", path: Path | str | None = None) -> Labels:
    return Labels(lang, load_labels_table(path))
