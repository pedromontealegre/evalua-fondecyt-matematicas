"""
Interfaz web compacta con Streamlit.

Ejecutar:
    streamlit run app_streamlit.py

El listado FONDECYT se carga automáticamente desde listado.xlsx,
ubicado en la misma carpeta que esta app.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

from fondecyt_eval import evaluate_bibtex, load_journal_catalog


APP_DIR = Path(__file__).resolve().parent
DEFAULT_LISTADO_PATH = APP_DIR / "listado.xlsx"
CATEGORY_ORDER = {"MB": 0, "B": 1, "R": 2}


st.set_page_config(
    page_title="Evaluador FONDECYT Matemáticas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.25rem;
        padding-bottom: 1.1rem;
        max-width: 1280px;
    }
    h1 {
        font-size: 1.55rem !important;
        margin-bottom: 0.1rem !important;
    }
    h2, h3 {
        margin-top: 0.55rem !important;
        margin-bottom: 0.25rem !important;
    }
    div[data-testid="stMetric"] {
        padding: 0.1rem 0.25rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.2rem;
    }
    .stDataFrame {
        margin-top: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


BASE_COLUMNS = {
    "category": "Clasificación",
    "year": "Año",
    "journal_bib": "Revista en BibTeX",
    "match_in_list": "Match en listado",
}

OPTIONAL_COLUMNS = {
    "title": "Título",
    "authors": "Autores",
}


@st.cache_data(show_spinner=False)
def _load_default_catalog(path: str) -> pd.DataFrame:
    return load_journal_catalog(path)


def _counts_from_summary(summary: pd.DataFrame) -> dict[str, int]:
    out: dict[str, int] = {}
    if summary is None or summary.empty:
        return out
    for _, row in summary.iterrows():
        metric = str(row.get("metric", ""))
        value = row.get("value", 0)
        try:
            out[metric] = int(value)
        except Exception:
            out[metric] = 0
    return out


def _height_for_table(df: pd.DataFrame, max_height: int = 480) -> int:
    if df is None or df.empty:
        return 110
    return min(max_height, 38 + 35 * (len(df) + 1))


def _make_match_label(row: pd.Series) -> str:
    abbr = str(row.get("matched_abbr", "") or "").strip()
    full = str(row.get("matched_to", "") or "").strip()

    if abbr and full and abbr != full:
        return f"{abbr} — {full}"
    return full or abbr


def _sort_recognized(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "category" not in df.columns:
        return df

    out = df.copy()
    out["_category_rank"] = out["category"].map(CATEGORY_ORDER).fillna(99).astype(int)

    sort_cols = ["_category_rank"]
    ascending = [True]

    if "year" in out.columns:
        sort_cols.append("year")
        ascending.append(False)
    if "journal_bib" in out.columns:
        sort_cols.append("journal_bib")
        ascending.append(True)
    if "title" in out.columns:
        sort_cols.append("title")
        ascending.append(True)

    out = out.sort_values(sort_cols, ascending=ascending, na_position="last")
    return out.drop(columns=["_category_rank"])


def _compact_results_table(matched: pd.DataFrame, optional: Iterable[str]) -> pd.DataFrame:
    if matched is None or matched.empty:
        return pd.DataFrame(columns=list(BASE_COLUMNS.values()))

    df = _sort_recognized(matched)
    df = df.copy()
    df["match_in_list"] = df.apply(_make_match_label, axis=1)

    cols = ["category", "year", "journal_bib", "match_in_list"]
    for col in optional:
        if col in df.columns:
            cols.append(col)

    df = df[[col for col in cols if col in df.columns]].copy()
    df = df.rename(columns={**BASE_COLUMNS, **OPTIONAL_COLUMNS})
    return df


def _compact_review_table(review: pd.DataFrame) -> pd.DataFrame:
    if review is None or review.empty:
        return pd.DataFrame()

    cols = [
        "year",
        "type",
        "journal_bib",
        "status",
        "best_candidate_abbr",
        "best_candidate",
        "best_candidate_category",
        "best_score",
        "title",
        "authors",
    ]
    df = review[[c for c in cols if c in review.columns]].copy()
    return df.rename(
        columns={
            "year": "Año",
            "type": "Tipo",
            "journal_bib": "Revista en BibTeX",
            "status": "Estado",
            "best_candidate_abbr": "Mejor candidato abreviado",
            "best_candidate": "Mejor candidato",
            "best_candidate_category": "Categoría candidata",
            "best_score": "Score",
            "title": "Título",
            "authors": "Autores",
        }
    )


st.title("Evaluador de publicaciones FONDECYT — Matemáticas")
st.caption("El listado FONDECYT se carga automáticamente. Puedes subir un archivo BibTeX o pegar la bibliografía directamente.")

if not DEFAULT_LISTADO_PATH.exists():
    st.error(
        "No encontré el archivo listado.xlsx junto a app_streamlit.py. "
        "Copia listado.xlsx en la misma carpeta de la app y vuelve a ejecutar Streamlit."
    )
    st.stop()

with st.form("evaluation_form"):
    col_bib, col_params, col_optional = st.columns([1.45, 0.85, 1.1])

    with col_bib:
        bib_file = st.file_uploader("Archivo BibTeX (.bib)", type=["bib", "txt"])
        bib_text_input = st.text_area(
            "O pega aquí la bibliografía en formato BibTeX",
            height=180,
            placeholder="@article{...\n  title = {...},\n  author = {...},\n  journal = {...},\n  year = {2024}\n}",
        )

    with col_params:
        since_year = st.number_input("Desde el año", min_value=1900, max_value=2100, value=2021, step=1)
        fuzzy_threshold = st.slider("Umbral fuzzy", min_value=80, max_value=100, value=95, step=1)

    with col_optional:
        st.markdown("**Campos opcionales**")
        show_title = st.checkbox("Título", value=False)
        show_authors = st.checkbox("Autores", value=False)
        submitted = st.form_submit_button("Actualizar resultados", type="primary", use_container_width=True)

if submitted:
    if bib_file is not None:
        bib_text = bib_file.read().decode("utf-8", errors="replace")
    else:
        bib_text = (bib_text_input or "").strip()

    if not bib_text:
        st.error("Falta subir un archivo .bib o pegar la bibliografía en el campo de texto.")
        st.stop()

    try:
        catalog = _load_default_catalog(str(DEFAULT_LISTADO_PATH))
        summary, matched, review, diagnostic = evaluate_bibtex(
            bib_text,
            catalog,
            since_year=int(since_year),
            fuzzy_threshold=int(fuzzy_threshold),
        )
        matched = _sort_recognized(matched)
    except Exception as exc:
        st.exception(exc)
        st.stop()

    st.session_state["summary"] = summary
    st.session_state["matched"] = matched
    st.session_state["review"] = review
    st.session_state["diagnostic"] = diagnostic
    st.session_state["optional_columns"] = [
        col for col, enabled in [("title", show_title), ("authors", show_authors)] if enabled
    ]

if "summary" not in st.session_state:
    st.info("Sube un archivo .bib o pega la bibliografía en el campo de texto, y luego presiona **Actualizar resultados**.")
    st.stop()

summary = st.session_state["summary"]
matched = st.session_state["matched"]
review = st.session_state["review"]
diagnostic = st.session_state["diagnostic"]
optional_columns = st.session_state.get("optional_columns", [])

counts = _counts_from_summary(summary)
st.markdown(
    f"**MB:** {counts.get('MB', 0)} &nbsp;&nbsp;·&nbsp;&nbsp; "
    f"**B:** {counts.get('B', 0)} &nbsp;&nbsp;·&nbsp;&nbsp; "
    f"**R:** {counts.get('R', 0)} &nbsp;&nbsp;·&nbsp;&nbsp; "
    f"**Reconocidas:** {counts.get('Total reconocido', 0)} &nbsp;&nbsp;·&nbsp;&nbsp; "
    f"**Para revisar:** {counts.get('Para revisar', 0)}",
    unsafe_allow_html=True,
)

results_table = _compact_results_table(matched, optional_columns)
st.subheader("Resultados")
st.dataframe(results_table, use_container_width=True, hide_index=True, height=_height_for_table(results_table))

csv_left, csv_right, _ = st.columns([1.0, 1.0, 2.2])
with csv_left:
    st.download_button(
        "Descargar tabla",
        data=results_table.to_csv(index=False).encode("utf-8"),
        file_name="resultados_compactos.csv",
        mime="text/csv",
        use_container_width=True,
    )
with csv_right:
    st.download_button(
        "Descargar diagnóstico",
        data=diagnostic.to_csv(index=False).encode("utf-8"),
        file_name="diagnostico.csv",
        mime="text/csv",
        use_container_width=True,
    )

with st.expander(f"Para revisar ({len(review)})", expanded=False):
    st.caption(
        "Entradas sin match aceptado: pueden ser proceedings, revistas no listadas, nombres abreviados raros o entradas sin journal."
    )
    review_table = _compact_review_table(review)
    st.dataframe(review_table, use_container_width=True, hide_index=True, height=_height_for_table(review_table, 360))
    st.download_button(
        "Descargar para revisar",
        data=review.to_csv(index=False).encode("utf-8"),
        file_name="para_revisar.csv",
        mime="text/csv",
    )
