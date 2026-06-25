"""
Evaluador simple de publicaciones en revistas del grupo Matemáticas FONDECYT.

Uso por consola:
    python fondecyt_eval.py --listado listado.xlsx --bib cv.bib --since 2021 --out resultados.csv

Dependencias:
    pip install pandas openpyxl bibtexparser rapidfuzz
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import bibtexparser
from bibtexparser.bparser import BibTexParser
from rapidfuzz import fuzz, process

VALID_CATEGORIES = {"MB", "B", "R"}
JOURNAL_FIELDS = ("journal", "journaltitle", "journal-title", "shortjournal")
YEAR_FIELDS = ("year", "date")


@dataclass(frozen=True)
class JournalRecord:
    abbr: str
    full_name: str
    category: str

    @property
    def display_name(self) -> str:
        return self.full_name or self.abbr


def latexish_to_text(value: Any) -> str:
    """Limpieza liviana de texto BibTeX/LaTeX para comparar nombres de revistas."""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\\&", "&")
    text = text.replace("{", "").replace("}", "")
    # Convierte comandos LaTeX simples en el argumento, por ejemplo \emph{X} -> X.
    text = re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?", " ", text)
    text = text.replace("~", " ")
    return text


def normalize_journal_name(value: Any) -> str:
    """Normaliza acentos, puntuación, mayúsculas y espacios."""
    text = latexish_to_text(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_journal_catalog(xlsx_path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """
    Carga el Excel FONDECYT.

    Espera columnas:
      A: abreviatura de revista
      B: nombre completo
      C: categoria MB/B/R
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, usecols=[0, 1, 2], dtype=str)
    df = df.rename(columns={df.columns[0]: "abbr", df.columns[1]: "full_name", df.columns[2]: "category"})

    for col in ("abbr", "full_name", "category"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["category"] = df["category"].str.upper().str.replace("\u00a0", "", regex=False).str.strip()
    df = df[df["category"].isin(VALID_CATEGORIES)].copy()
    df["abbr_norm"] = df["abbr"].map(normalize_journal_name)
    df["full_norm"] = df["full_name"].map(normalize_journal_name)
    df = df.drop_duplicates(subset=["abbr_norm", "full_norm", "category"]).reset_index(drop=True)
    return df


def build_indices(catalog: pd.DataFrame) -> tuple[dict[str, JournalRecord], dict[str, JournalRecord]]:
    """
    Devuelve dos índices:
      - exact_index: abreviaturas y nombres completos normalizados.
      - fuzzy_index: solo nombres largos, para reducir falsos positivos.
    """
    exact_index: dict[str, JournalRecord] = {}
    fuzzy_index: dict[str, JournalRecord] = {}

    for row in catalog.itertuples(index=False):
        record = JournalRecord(abbr=row.abbr, full_name=row.full_name, category=row.category)
        for alias in (row.abbr_norm, row.full_norm):
            if alias:
                exact_index.setdefault(alias, record)
        if row.full_norm and len(row.full_norm) >= 12:
            fuzzy_index.setdefault(row.full_norm, record)

    return exact_index, fuzzy_index


def parse_bibtex(bib_text: str) -> list[dict[str, Any]]:
    parser = BibTexParser(common_strings=True)
    parser.ignore_nonstandard_types = False
    parser.homogenize_fields = False
    db = bibtexparser.loads(bib_text, parser=parser)
    return db.entries


def extract_year(entry: dict[str, Any]) -> int | None:
    for field in YEAR_FIELDS:
        value = entry.get(field)
        if value:
            match = re.search(r"(19|20)\d{2}", str(value))
            if match:
                return int(match.group(0))
    return None


def extract_journal(entry: dict[str, Any]) -> str:
    for field in JOURNAL_FIELDS:
        value = entry.get(field)
        if value:
            return str(value)
    return ""


def match_journal(
    journal_name: str,
    exact_index: dict[str, JournalRecord],
    fuzzy_index: dict[str, JournalRecord],
    fuzzy_threshold: int = 92,
) -> tuple[JournalRecord | None, str, int | None]:
    """
    Retorna (registro, tipo_de_match, score).

    tipo_de_match:
      - exact: coincidencia exacta después de normalización.
      - fuzzy: coincidencia aproximada sobre nombres completos.
      - no_match: no se encontró coincidencia confiable.
    """
    norm = normalize_journal_name(journal_name)
    if not norm:
        return None, "missing_journal", None

    if norm in exact_index:
        return exact_index[norm], "exact", 100

    # Evita fuzzy matching para abreviaturas muy cortas: ahí es más seguro exigir match exacto.
    if len(norm) < 12 or not fuzzy_index:
        return None, "no_match", None

    candidate = process.extractOne(norm, list(fuzzy_index.keys()), scorer=fuzz.WRatio)
    if candidate is None:
        return None, "no_match", None

    candidate_norm, score, _ = candidate
    score = int(round(score))
    if score >= fuzzy_threshold:
        return fuzzy_index[candidate_norm], "fuzzy", score

    return None, "no_match", score


def evaluate_bibtex(
    bib_text: str,
    catalog: pd.DataFrame,
    since_year: int = 2021,
    fuzzy_threshold: int = 92,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Evalúa un BibTeX contra el catálogo.

    Retorna:
      summary_df: conteos agregados.
      matched_df: publicaciones desde since_year con revista reconocida.
      review_df: publicaciones desde since_year no reconocidas o incompletas.
    """
    exact_index, fuzzy_index = build_indices(catalog)
    entries = parse_bibtex(bib_text)

    matched_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []

    for entry in entries:
        key = entry.get("ID", "")
        title = latexish_to_text(entry.get("title", "")).strip()
        year = extract_year(entry)
        journal_raw = extract_journal(entry)
        entry_type = entry.get("ENTRYTYPE", "")

        if year is None:
            review_rows.append({
                "key": key,
                "year": None,
                "type": entry_type,
                "title": title,
                "journal_bib": journal_raw,
                "status": "no_year",
                "best_score": None,
            })
            continue

        if year < since_year:
            continue

        record, match_type, score = match_journal(journal_raw, exact_index, fuzzy_index, fuzzy_threshold)
        if record is None:
            review_rows.append({
                "key": key,
                "year": year,
                "type": entry_type,
                "title": title,
                "journal_bib": journal_raw,
                "status": match_type,
                "best_score": score,
            })
            continue

        matched_rows.append({
            "key": key,
            "year": year,
            "type": entry_type,
            "title": title,
            "journal_bib": journal_raw,
            "category": record.category,
            "matched_to": record.display_name,
            "matched_abbr": record.abbr,
            "match_type": match_type,
            "score": score,
        })

    matched_df = pd.DataFrame(matched_rows)
    review_df = pd.DataFrame(review_rows)

    if not matched_df.empty:
        counts = matched_df["category"].value_counts().reindex(["MB", "B", "R"], fill_value=0)
    else:
        counts = pd.Series({"MB": 0, "B": 0, "R": 0})

    summary_df = pd.DataFrame([
        {"metric": "Desde año", "value": since_year},
        {"metric": "MB", "value": int(counts.get("MB", 0))},
        {"metric": "B", "value": int(counts.get("B", 0))},
        {"metric": "R", "value": int(counts.get("R", 0))},
        {"metric": "Total reconocido", "value": int(counts.sum())},
        {"metric": "Para revisar", "value": int(len(review_df))},
    ])

    sort_cols = [col for col in ["year", "category", "journal_bib", "title"] if col in matched_df.columns]
    if sort_cols:
        matched_df = matched_df.sort_values(sort_cols, ascending=[False] + [True] * (len(sort_cols) - 1))
    if not review_df.empty:
        review_df = review_df.sort_values(["year", "journal_bib", "title"], ascending=[False, True, True], na_position="last")

    return summary_df, matched_df.reset_index(drop=True), review_df.reset_index(drop=True)


def evaluate_bib_file(
    bib_path: str | Path,
    xlsx_path: str | Path,
    since_year: int = 2021,
    fuzzy_threshold: int = 92,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    catalog = load_journal_catalog(xlsx_path)
    bib_text = Path(bib_path).read_text(encoding="utf-8")
    return evaluate_bibtex(bib_text, catalog, since_year=since_year, fuzzy_threshold=fuzzy_threshold)


def main() -> None:
    argp = argparse.ArgumentParser(description="Evalúa publicaciones BibTeX contra listado FONDECYT Matemáticas.")
    argp.add_argument("--listado", required=True, help="Ruta al Excel con revistas FONDECYT, por ejemplo listado.xlsx")
    argp.add_argument("--bib", required=True, help="Ruta al archivo .bib")
    argp.add_argument("--since", type=int, default=2021, help="Año inicial, inclusivo. Default: 2021")
    argp.add_argument("--threshold", type=int, default=92, help="Umbral fuzzy matching. Default: 92")
    argp.add_argument("--out", default="resultados.csv", help="CSV con publicaciones reconocidas. Default: resultados.csv")
    argp.add_argument("--review-out", default="para_revisar.csv", help="CSV con publicaciones no reconocidas. Default: para_revisar.csv")
    args = argp.parse_args()

    summary, matched, review = evaluate_bib_file(
        bib_path=args.bib,
        xlsx_path=args.listado,
        since_year=args.since,
        fuzzy_threshold=args.threshold,
    )

    print("\n=== Resumen ===")
    print(summary.to_string(index=False))

    matched.to_csv(args.out, index=False)
    review.to_csv(args.review_out, index=False)
    print(f"\nPublicaciones reconocidas guardadas en: {args.out}")
    print(f"Publicaciones para revisar guardadas en: {args.review_out}")


if __name__ == "__main__":
    main()
