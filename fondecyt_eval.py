"""
Evaluador simple de publicaciones en revistas del grupo Matemáticas FONDECYT.

Uso por consola:
    python fondecyt_eval.py --listado listado.xlsx --bib cv.bib --since 2021

Dependencias:
    pip install pandas openpyxl rapidfuzz

Notas importantes:
  - Esta versión NO depende de bibtexparser.
  - El matching usa tres niveles:
      1. coincidencia exacta normalizada;
      2. coincidencia exacta por clave canónica, que tolera abreviaturas típicas
         como Theor./Theoret./Theoretical, Syst./System(s), J./Journal, etc.;
      3. coincidencia aproximada conservadora, dejando trazabilidad.
  - Siempre genera una tabla de diagnóstico para ver qué leyó desde el .bib.
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from rapidfuzz import fuzz, process

VERSION = "0.5.1-textarea-input"

VALID_CATEGORIES = {"MB", "B", "R"}
JOURNAL_FIELDS = ("journal", "journaltitle", "journal-title", "shortjournal", "journalabbr", "publication")
YEAR_FIELDS = ("year", "date")

# BibTeX month strings; useful if a .bib uses month = jan, etc.
DEFAULT_STRINGS = {
    "jan": "January",
    "feb": "February",
    "mar": "March",
    "apr": "April",
    "may": "May",
    "jun": "June",
    "jul": "July",
    "aug": "August",
    "sep": "September",
    "oct": "October",
    "nov": "November",
    "dec": "December",
}

# Palabras que no ayudan a identificar la revista.
CANONICAL_STOPWORDS = {
    "A", "AN", "AND", "THE", "OF", "ON", "IN", "FOR", "TO", "WITH", "BY",
    "DE", "DEL", "LA", "LE", "LES", "EL", "DER", "DIE", "DAS", "UND", "ET",
    "INTERNATIONAL", "INT", "NEW", "SER", "SERIES", "ENGLISH", "EDITION",
}

# Mapeo intencionalmente simple de abreviaturas frecuentes en revistas matemáticas/CS.
# La meta no es traducir todo, sino que variantes usuales caigan en la misma clave.
CANONICAL_TOKEN_MAP = {
    # journal / proceedings / transactions
    "J": "J",
    "JR": "J",
    "JOURNAL": "J",
    "JOURNALS": "J",
    "PROC": "PROC",
    "PROCEEDINGS": "PROC",
    "TRANS": "TRANS",
    "TRANSACTIONS": "TRANS",
   
    # mathematics
    "MATH": "MATH",
    "MATHS": "MATH",
    "MATHEMATICA": "MATH",
    "MATHEMATICAL": "MATH",
    "MATHEMATICS": "MATH",
    "MATHEMATIK": "MATH",
    "MATHEMATIQUES": "MATH",
   
    # computer science
    "COMPUT": "COMPUT",
    "COMP": "COMPUT",
    "COMPUTER": "COMPUT",
    "COMPUTERS": "COMPUT",
    "COMPUTING": "COMPUT",
    "COMPUTATION": "COMPUT",
    "COMPUTATIONAL": "COMPUT",
    "SCI": "SCI",
    "SC": "SCI",
    "SCIENCE": "SCI",
    "SCIENCES": "SCI",
    "SCIENTIFIC": "SCI",
    "SCIENTIA": "SCI",
    "THEOR": "THEOR",
    "THEORET": "THEOR",
    "THEORETICAL": "THEOR",
    "THEORY": "THEOR",
    "SYST": "SYSTEM",
    "SYSTEM": "SYSTEM",
    "SYSTEMS": "SYSTEM",
    "INFORM": "INFORM",
    "INF": "INFORM",
    "INFO": "INFORM",
    "INFORMATION": "INFORM",
    "PROCESS": "PROCESS",
    "PROCESSING": "PROCESS",
    "PROCESSOR": "PROCESS",
    "PROGRAM": "PROGRAM",
    "PROGRAMMING": "PROGRAM",
   
    # common math/OR words
    "APPL": "APPL",
    "APPLICABLE": "APPL",
    "APPLICANDAE": "APPL",
    "APPLIED": "APPL",
    "APPLICATION": "APPL",
    "APPLICATIONS": "APPL",
    "ANAL": "ANAL",
    "ANALYSIS": "ANAL",
    "ANALYTIC": "ANAL",
    "ALG": "ALG",
    "ALGEBRA": "ALG",
    "ALGEBRAIC": "ALG",
    "ALGORITHM": "ALGORITHM",
    "ALGORITHMS": "ALGORITHM",
    "ALGORITHMIC": "ALGORITHM",
    "ALGORITHMICA": "ALGORITHMICA",
    "ANN": "ANN",
    "ANNALS": "ANN",
    "ANNAL": "ANN",
    "BULL": "BULL",
    "BULLETIN": "BULL",
    "CALC": "CALC",
    "CALCULUS": "CALC",
    "COMBIN": "COMBIN",
    "COMBINATORICA": "COMBINATORICA",
    "COMBINATORIAL": "COMBIN",
    "COMBINATORICS": "COMBIN",
    "COMB": "COMBIN",
    "COMM": "COMM",
    "COMMUNICATIONS": "COMM",
    "COMMUNICATION": "COMM",
    "CONT": "CONT",
    "CONTEMP": "CONTEMP",
    "CONTROL": "CONTROL",
    "DISTRIB": "DISTRIB",
    "DISTRIBUTED": "DISTRIB",
    "DISTRIBUTION": "DISTRIB",
    "DISCRET": "DISCRETE",
    "DISCRETE": "DISCRETE",
    "DYN": "DYN",
    "DYNAMIC": "DYN",
    "DYNAMICAL": "DYN",
    "DYNAMICS": "DYN",
    "ELECTRON": "ELECTRON",
    "ELECTRONIC": "ELECTRON",
    "FINITE": "FINITE",
    "FIELDS": "FIELD",
    "FIELD": "FIELD",
    "GEOM": "GEOM",
    "GEOMETRY": "GEOM",
    "GEOMETRIC": "GEOM",
    "GRAPH": "GRAPH",
    "GRAPHS": "GRAPH",
    "GROUP": "GROUP",
    "GROUPS": "GROUP",
    "INTEGRAL": "INTEGRAL",
    "JPN": "JAPAN",
    "JAPANESE": "JAPAN",
    "LETT": "LETT",
    "LETTER": "LETT",
    "LETTERS": "LETT",
    "LINEAR": "LINEAR",
    "LOGIC": "LOGIC",
    "MECH": "MECH",
    "MECHANICS": "MECH",
    "METHOD": "METHOD",
    "METHODS": "METHOD",
    "MODEL": "MODEL",
    "MODELLING": "MODEL",
    "MODELING": "MODEL",
    "NUMER": "NUMER",
    "NUMERICAL": "NUMER",
    "OPER": "OPER",
    "OPERATIONAL": "OPER",
    "OPERATIONS": "OPER",
    "OPTIM": "OPTIM",
    "OPTIMAL": "OPTIM",
    "OPTIMIZATION": "OPTIM",
    "PROBAB": "PROB",
    "PROBABILITY": "PROB",
    "PROBL": "PROBL",
    "PROBLEMS": "PROBL",
    "RES": "RES",
    "RESEARCH": "RES",
    "REV": "REV",
    "REVIEW": "REV",
    "REVIEWS": "REV",
    "STAT": "STAT",
    "STATISTIC": "STAT",
    "STATISTICAL": "STAT",
    "STATISTICS": "STAT",
    "STOCH": "STOCH",
    "STOCHASTIC": "STOCH",
    "STRUCT": "STRUCT",
    "STRUCTURES": "STRUCT",
    "TOPOL": "TOPOL",
    "TOPOLOGY": "TOPOL",
    "ZEITSCHRIFT": "Z",
    "Z": "Z",
}

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
    replacements = {
        r"\&": "&",
        r"\'a": "á", r"\'e": "é", r"\'i": "í", r"\'o": "ó", r"\'u": "ú",
        r"\`a": "à", r"\`e": "è", r"\`i": "ì", r"\`o": "ò", r"\`u": "ù",
        r"\~n": "ñ", r"\"o": "ö", r"\"u": "ü",
        r"\o": "ø", r"\O": "Ø",
        r"\ss": "ss",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = text.replace("{", "").replace("}", "")
    # Convierte comandos LaTeX simples en espacio, por ejemplo \emph{X} -> X
    # después de remover llaves queda \emph X, y eliminamos el comando.
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


def canonical_journal_key(value: Any) -> str:
    """
    Clave canónica para comparar nombres abreviados y completos.

    Ejemplos que deberían caer juntos:
      - Theor. Comput. Sci. / Theoret. Comput. Sci. / Theoretical Computer Science
      - J. Comput. Syst. Sci. / Journal of Computer and System Sciences
      - Inf. Comput. / Inform. and Comput. / Information and Computation
      - SIAM J. Comput. / SIAM Journal on Computing
    """
    norm = normalize_journal_name(value)
    if not norm:
        return ""
    out: list[str] = []
    for token in norm.split():
        if token in CANONICAL_STOPWORDS:
            continue
        mapped = CANONICAL_TOKEN_MAP.get(token, token)
        if mapped in CANONICAL_STOPWORDS:
            continue
        out.append(mapped)
    return " ".join(out)


def load_journal_catalog(xlsx_path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """
    Carga el Excel FONDECYT.

    Espera columnas:
      A: abreviatura de revista
      B: nombre completo
      C: categoría MB/B/R
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, usecols=[0, 1, 2], dtype=str)
    df = df.rename(columns={df.columns[0]: "abbr", df.columns[1]: "full_name", df.columns[2]: "category"})

    for col in ("abbr", "full_name", "category"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["category"] = df["category"].str.upper().str.replace("\u00a0", "", regex=False).str.strip()
    df = df[df["category"].isin(VALID_CATEGORIES)].copy()
    df["abbr_norm"] = df["abbr"].map(normalize_journal_name)
    df["full_norm"] = df["full_name"].map(normalize_journal_name)
    df["abbr_key"] = df["abbr"].map(canonical_journal_key)
    df["full_key"] = df["full_name"].map(canonical_journal_key)
    df = df.drop_duplicates(subset=["abbr_norm", "full_norm", "category"]).reset_index(drop=True)
    return df


def build_indices(catalog: pd.DataFrame) -> tuple[
    dict[str, JournalRecord],
    dict[str, JournalRecord],
    dict[str, JournalRecord],
]:
    """
    Devuelve tres índices:
      - exact_index: abreviaturas y nombres completos normalizados.
      - canonical_index: abreviaturas y nombres completos reducidos a una clave de tokens.
      - fuzzy_index: aliases suficientemente largos, para diagnóstico y matching aproximado.
    """
    exact_index: dict[str, JournalRecord] = {}
    canonical_index: dict[str, JournalRecord] = {}
    fuzzy_index: dict[str, JournalRecord] = {}

    for row in catalog.itertuples(index=False):
        record = JournalRecord(abbr=row.abbr, full_name=row.full_name, category=row.category)
        for alias in (row.abbr_norm, row.full_norm):
            if alias:
                exact_index.setdefault(alias, record)
            if alias and len(alias) >= 8:
                fuzzy_index.setdefault(alias, record)
        for alias in (row.abbr_key, row.full_key):
            if alias:
                canonical_index.setdefault(alias, record)
            if alias and len(alias) >= 8:
                fuzzy_index.setdefault(alias, record)

    return exact_index, canonical_index, fuzzy_index


# ---------------------------------------------------------------------------
# Parser BibTeX liviano
# ---------------------------------------------------------------------------

def _find_matching_delimiter(text: str, start: int, open_ch: str, close_ch: str) -> int:
    """Retorna el índice del delimitador de cierre que balancea text[start]."""
    depth = 0
    in_quote = False
    escaped = False

    for i in range(start, len(text)):
        ch = text[i]
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            in_quote = not in_quote
            continue
        if in_quote:
            continue
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return i
    return -1


def _iter_raw_bib_entries(bib_text: str) -> list[tuple[str, str]]:
    """Extrae pares (tipo, cuerpo) desde un texto BibTeX."""
    entries: list[tuple[str, str]] = []
    i = 0
    n = len(bib_text)

    while i < n:
        at = bib_text.find("@", i)
        if at == -1:
            break

        j = at + 1
        while j < n and (bib_text[j].isalpha() or bib_text[j] in "_-:"):
            j += 1
        entry_type = bib_text[at + 1:j].strip().lower()

        while j < n and bib_text[j].isspace():
            j += 1
        if j >= n or bib_text[j] not in "{(":
            i = j + 1
            continue

        open_ch = bib_text[j]
        close_ch = "}" if open_ch == "{" else ")"
        end = _find_matching_delimiter(bib_text, j, open_ch, close_ch)
        if end == -1:
            # Entrada mal cerrada: salimos para no entrar en loop infinito.
            break

        body = bib_text[j + 1:end]
        entries.append((entry_type, body))
        i = end + 1

    return entries


def _split_top_level_commas(text: str) -> list[str]:
    """Separa por comas que no estén dentro de llaves o comillas."""
    parts: list[str] = []
    start = 0
    depth = 0
    in_quote = False
    escaped = False

    for i, ch in enumerate(text):
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            in_quote = not in_quote
            continue
        if in_quote:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
        elif ch == "," and depth == 0:
            parts.append(text[start:i].strip())
            start = i + 1

    last = text[start:].strip()
    if last:
        parts.append(last)
    return parts


def _strip_outer_delimiters(value: str) -> str:
    """Remueve un par exterior de llaves o comillas si cubre todo el valor."""
    value = value.strip()
    if len(value) >= 2 and value[0] == "{" and _find_matching_delimiter(value, 0, "{", "}") == len(value) - 1:
        return value[1:-1].strip()
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value[1:-1].strip()
    return value.strip()


def _split_hash_concat(text: str) -> list[str]:
    """Separa concatenaciones BibTeX con # fuera de llaves o comillas."""
    parts: list[str] = []
    start = 0
    depth = 0
    in_quote = False
    escaped = False

    for i, ch in enumerate(text):
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            in_quote = not in_quote
            continue
        if in_quote:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
        elif ch == "#" and depth == 0:
            parts.append(text[start:i].strip())
            start = i + 1

    parts.append(text[start:].strip())
    return parts


def _parse_bib_value(value: str, strings: dict[str, str]) -> str:
    """
    Parsea un valor BibTeX básico.

    Soporta:
      - journal = {Journal Name}
      - journal = "Journal Name"
      - journal = JACM
      - title = "A" # " " # "B"
    """
    pieces = _split_hash_concat(value)
    parsed_pieces: list[str] = []
    for piece in pieces:
        stripped_piece = piece.strip()
        p = _strip_outer_delimiters(stripped_piece)
        key = p.strip().lower()
        if key in strings and stripped_piece == p.strip():
            parsed_pieces.append(strings[key])
        else:
            parsed_pieces.append(p)
    return "".join(parsed_pieces).strip()


def _parse_assignment_list(text: str, strings: dict[str, str]) -> dict[str, str]:
    fields: dict[str, str] = {}
    for part in _split_top_level_commas(text):
        if "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = name.strip().lower()
        if not name:
            continue
        fields[name] = _parse_bib_value(value, strings)
    return fields


def parse_bibtex(bib_text: str) -> list[dict[str, Any]]:
    """
    Parser BibTeX liviano, suficiente para extraer ID, tipo, year/date, title y journal.
    """
    strings = dict(DEFAULT_STRINGS)
    parsed_entries: list[dict[str, Any]] = []

    for entry_type, body in _iter_raw_bib_entries(bib_text):
        if entry_type == "string":
            fields = _parse_assignment_list(body, strings)
            strings.update({k.lower(): v for k, v in fields.items()})
            continue
        if entry_type in {"comment", "preamble"}:
            continue

        parts = _split_top_level_commas(body)
        if not parts:
            continue

        key = parts[0].strip()
        fields_text = ",".join(parts[1:])
        fields = _parse_assignment_list(fields_text, strings)
        fields["ENTRYTYPE"] = entry_type
        fields["ID"] = key
        parsed_entries.append(fields)

    return parsed_entries


# ---------------------------------------------------------------------------
# Evaluación
# ---------------------------------------------------------------------------

def extract_year(entry: dict[str, Any]) -> int | None:
    for field in YEAR_FIELDS:
        value = entry.get(field)
        if value:
            match = re.search(r"(19|20)\d{2}", str(value))
            if match:
                return int(match.group(0))
    return None


def extract_journal(entry: dict[str, Any]) -> tuple[str, str]:
    """Retorna (campo_usado, valor)."""
    for field in JOURNAL_FIELDS:
        value = entry.get(field)
        if value:
            return field, str(value)
    return "", ""


def best_candidate_for_diagnostics(
    journal_name: str,
    fuzzy_index: dict[str, JournalRecord],
) -> tuple[JournalRecord | None, str, int | None]:
    norm = normalize_journal_name(journal_name)
    key = canonical_journal_key(journal_name)
    query = key or norm
    if not query or not fuzzy_index:
        return None, "", None
    candidate = process.extractOne(query, list(fuzzy_index.keys()), scorer=fuzz.WRatio)
    if candidate is None:
        return None, "", None
    candidate_alias, score, _ = candidate
    return fuzzy_index[candidate_alias], candidate_alias, int(round(score))


def match_journal(
    journal_name: str,
    exact_index: dict[str, JournalRecord],
    canonical_index: dict[str, JournalRecord],
    fuzzy_index: dict[str, JournalRecord],
    fuzzy_threshold: int = 86,
) -> tuple[JournalRecord | None, str, int | None, str]:
    """
    Retorna (registro, tipo_de_match, score, alias/candidato).

    tipo_de_match:
      - exact: coincidencia exacta después de normalización.
      - canonical: coincidencia exacta con clave canónica de abreviaturas.
      - fuzzy: coincidencia aproximada.
      - no_match: no se encontró coincidencia confiable.
    """
    norm = normalize_journal_name(journal_name)
    if not norm:
        return None, "missing_journal", None, ""

    if norm in exact_index:
        return exact_index[norm], "exact", 100, norm

    key = canonical_journal_key(journal_name)
    if key and key in canonical_index:
        return canonical_index[key], "canonical", 100, key

    if not fuzzy_index:
        return None, "no_match", None, ""

    query = key or norm
    if len(query) < 4:
        return None, "no_match", None, ""

    candidate = process.extractOne(query, list(fuzzy_index.keys()), scorer=fuzz.WRatio)
    if candidate is None:
        return None, "no_match", None, ""

    candidate_alias, score, _ = candidate
    score = int(round(score))
    if score >= fuzzy_threshold:
        return fuzzy_index[candidate_alias], "fuzzy", score, candidate_alias

    return None, "no_match", score, candidate_alias


def evaluate_bibtex(
    bib_text: str,
    catalog: pd.DataFrame,
    since_year: int = 2021,
    fuzzy_threshold: int = 86,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Evalúa un BibTeX contra el catálogo.

    Retorna:
      summary_df: conteos agregados.
      matched_df: publicaciones desde since_year con revista reconocida.
      review_df: publicaciones desde since_year no reconocidas o incompletas.
      diagnostic_df: todas las entradas leídas, con campos extraídos y candidato.
    """
    exact_index, canonical_index, fuzzy_index = build_indices(catalog)
    entries = parse_bibtex(bib_text)

    matched_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []
    diagnostic_rows: list[dict[str, Any]] = []

    for entry in entries:
        key = entry.get("ID", "")
        title = latexish_to_text(entry.get("title", "")).strip()
        authors = latexish_to_text(entry.get("author", "")).strip()
        year = extract_year(entry)
        journal_field, journal_raw = extract_journal(entry)
        entry_type = entry.get("ENTRYTYPE", "")
        journal_norm = normalize_journal_name(journal_raw)
        journal_key = canonical_journal_key(journal_raw)

        base_diag: dict[str, Any] = {
            "key": key,
            "year": year,
            "type": entry_type,
            "title": title,
            "authors": authors,
            "journal_field": journal_field,
            "journal_bib": journal_raw,
            "journal_norm": journal_norm,
            "journal_key": journal_key,
        }

        if year is None:
            cand_record, cand_alias, cand_score = best_candidate_for_diagnostics(journal_raw, fuzzy_index)
            row = {
                **base_diag,
                "status": "no_year",
                "category": "",
                "matched_to": "",
                "matched_abbr": "",
                "match_type": "",
                "score": "",
                "best_candidate": cand_record.display_name if cand_record else "",
                "best_candidate_abbr": cand_record.abbr if cand_record else "",
                "best_candidate_category": cand_record.category if cand_record else "",
                "best_candidate_alias": cand_alias,
                "best_score": cand_score,
            }
            review_rows.append(row)
            diagnostic_rows.append(row)
            continue

        if year < since_year:
            row = {
                **base_diag,
                "status": "before_since",
                "category": "",
                "matched_to": "",
                "matched_abbr": "",
                "match_type": "",
                "score": "",
                "best_candidate": "",
                "best_candidate_abbr": "",
                "best_candidate_category": "",
                "best_candidate_alias": "",
                "best_score": "",
            }
            diagnostic_rows.append(row)
            continue

        record, match_type, score, matched_alias = match_journal(
            journal_raw,
            exact_index,
            canonical_index,
            fuzzy_index,
            fuzzy_threshold,
        )
        cand_record, cand_alias, cand_score = best_candidate_for_diagnostics(journal_raw, fuzzy_index)

        if record is None:
            row = {
                **base_diag,
                "status": match_type,
                "category": "",
                "matched_to": "",
                "matched_abbr": "",
                "match_type": "",
                "score": "",
                "best_candidate": cand_record.display_name if cand_record else "",
                "best_candidate_abbr": cand_record.abbr if cand_record else "",
                "best_candidate_category": cand_record.category if cand_record else "",
                "best_candidate_alias": cand_alias,
                "best_score": cand_score,
            }
            review_rows.append(row)
            diagnostic_rows.append(row)
            continue

        row = {
            **base_diag,
            "status": "matched",
            "category": record.category,
            "matched_to": record.display_name,
            "matched_abbr": record.abbr,
            "match_type": match_type,
            "score": score,
            "matched_alias": matched_alias,
            # En filas ya reconocidas, el mejor candidato relevante ES el match aceptado.
            "best_candidate": record.display_name,
            "best_candidate_abbr": record.abbr,
            "best_candidate_category": record.category,
            "best_candidate_alias": matched_alias,
            "best_score": score,
        }
        matched_rows.append(row)
        diagnostic_rows.append(row)

    matched_df = pd.DataFrame(matched_rows)
    review_df = pd.DataFrame(review_rows)
    diagnostic_df = pd.DataFrame(diagnostic_rows)

    if not matched_df.empty:
        counts = matched_df["category"].value_counts().reindex(["MB", "B", "R"], fill_value=0)
    else:
        counts = pd.Series({"MB": 0, "B": 0, "R": 0})

    no_journal_count = 0
    if not diagnostic_df.empty:
        no_journal_count = int((diagnostic_df.get("journal_bib", pd.Series(dtype=str)).fillna("") == "").sum())

    summary_df = pd.DataFrame([
        {"metric": "Entradas BibTeX leídas", "value": int(len(entries))},
        {"metric": "Desde año", "value": since_year},
        {"metric": "MB", "value": int(counts.get("MB", 0))},
        {"metric": "B", "value": int(counts.get("B", 0))},
        {"metric": "R", "value": int(counts.get("R", 0))},
        {"metric": "Total reconocido", "value": int(counts.sum())},
        {"metric": "Para revisar", "value": int(len(review_df))},
        {"metric": "Sin campo journal", "value": no_journal_count},
    ])

    # Orden por defecto para evaluación: primero MB, luego B, luego R.
    if not matched_df.empty and "category" in matched_df.columns:
        category_order = {"MB": 0, "B": 1, "R": 2}
        matched_df = matched_df.copy()
        matched_df["_category_rank"] = matched_df["category"].map(category_order).fillna(99).astype(int)
        sort_cols = ["_category_rank"]
        ascending = [True]
        if "year" in matched_df.columns:
            sort_cols.append("year")
            ascending.append(False)
        if "journal_bib" in matched_df.columns:
            sort_cols.append("journal_bib")
            ascending.append(True)
        if "title" in matched_df.columns:
            sort_cols.append("title")
            ascending.append(True)
        matched_df = matched_df.sort_values(sort_cols, ascending=ascending, na_position="last")
        matched_df = matched_df.drop(columns=["_category_rank"])
    if not review_df.empty:
        review_df = review_df.sort_values(["year", "journal_bib", "title"], ascending=[False, True, True], na_position="last")
    if not diagnostic_df.empty:
        diagnostic_df = diagnostic_df.sort_values(["year", "journal_bib", "title"], ascending=[False, True, True], na_position="last")

    return summary_df, matched_df.reset_index(drop=True), review_df.reset_index(drop=True), diagnostic_df.reset_index(drop=True)


def evaluate_bib_file(
    bib_path: str | Path,
    xlsx_path: str | Path,
    since_year: int = 2021,
    fuzzy_threshold: int = 86,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    catalog = load_journal_catalog(xlsx_path)
    bib_text = Path(bib_path).read_text(encoding="utf-8", errors="replace")
    return evaluate_bibtex(bib_text, catalog, since_year=since_year, fuzzy_threshold=fuzzy_threshold)


def _print_compact_review(review: pd.DataFrame, max_rows: int = 12) -> None:
    if review.empty:
        return
    cols = [
        "key", "year", "type", "journal_field", "journal_bib", "status",
        "best_candidate_abbr", "best_candidate_category", "best_score",
    ]
    cols = [c for c in cols if c in review.columns]
    print("\n=== Para revisar: primeras filas ===")
    print(review[cols].head(max_rows).to_string(index=False))


def main() -> None:
    argp = argparse.ArgumentParser(description="Evalúa publicaciones BibTeX contra listado FONDECYT Matemáticas.")
    argp.add_argument("--version", action="version", version=f"fondecyt_eval {VERSION}")
    argp.add_argument("--listado", default="listado.xlsx", help="Ruta al Excel con revistas FONDECYT. Default: listado.xlsx")
    argp.add_argument("--bib", required=True, help="Ruta al archivo .bib")
    argp.add_argument("--since", type=int, default=2021, help="Año inicial, inclusivo. Default: 2021")
    argp.add_argument("--threshold", type=int, default=86, help="Umbral fuzzy matching. Default: 86")
    argp.add_argument("--out", default="resultados.csv", help="CSV con publicaciones reconocidas. Default: resultados.csv")
    argp.add_argument("--review-out", default="para_revisar.csv", help="CSV con publicaciones no reconocidas. Default: para_revisar.csv")
    argp.add_argument("--debug-out", default="diagnostico.csv", help="CSV con todas las entradas y candidatos. Default: diagnostico.csv")
    argp.add_argument("--show-review", action="store_true", help="Muestra en pantalla las primeras filas para revisar.")
    args = argp.parse_args()

    summary, matched, review, diagnostic = evaluate_bib_file(
        bib_path=args.bib,
        xlsx_path=args.listado,
        since_year=args.since,
        fuzzy_threshold=args.threshold,
    )

    print("\n=== Resumen ===")
    print(summary.to_string(index=False))

    matched.to_csv(args.out, index=False)
    review.to_csv(args.review_out, index=False)
    diagnostic.to_csv(args.debug_out, index=False)

    print(f"\nPublicaciones reconocidas guardadas en: {args.out}")
    print(f"Publicaciones para revisar guardadas en: {args.review_out}")
    print(f"Diagnóstico completo guardado en: {args.debug_out}")

    if args.show_review or matched.empty:
        _print_compact_review(review)


if __name__ == "__main__":
    main()
