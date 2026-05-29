from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from io import BytesIO
import os
from pathlib import Path
import re
import subprocess
import textwrap
import unicodedata
from uuid import uuid4
from typing import Any

import pandas as pd
from jinja2 import Template
import plotly.graph_objects as go
import plotly.io as pio
from shutil import which
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


PILLAR_RANGES: list[tuple[int, int, str]] = [
    (1, 3, "Exigências Quantitativas"),
    (4, 4, "Ritmo de Trabalho"),
    (5, 7, "Exigências Cognitivas"),
    (8, 8, "Exigências Emocionais"),
    (9, 12, "Influência no Trabalho"),
    (13, 15, "Possibilidades de Desenvolvimento"),
    (16, 17, "Previsibilidade"),
    (18, 20, "Transparência no papel laboral desempenhado"),
    (21, 23, "Recompensas"),
    (24, 26, "Conflitos Laborais"),
    (27, 29, "Apoio Social de Colegas"),
    (30, 32, "Apoio Social de Superiores"),
    (33, 35, "Comunicado Social no Trabalho"),
    (36, 39, "Qualidade da Liderança"),
    (40, 42, "Confiança Horizontal"),
    (43, 45, "Confiança Vertical"),
    (46, 48, "Justiça e Respeito"),
    (49, 50, "Auto-eficácia"),
    (51, 53, "Significado do Trabalho"),
    (54, 55, "Compromisso Face ao Local de Trabalho"),
    (56, 59, "Satisfação no Trabalho"),
    (60, 60, "Insegurança no Trabalho"),
    (61, 61, "Saúde Geral"),
    (62, 64, "Conflito Trabalho/Família"),
    (65, 66, "Problemas em Dormir"),
    (67, 68, "Burnout"),
    (69, 70, "Stress"),
    (71, 72, "Sintomas Depressivos"),
    (73, 76, "Comportamentos Ofensivos"),
]

POSITIVE_STANDARD: dict[str, int] = {
    "nunca": 0,
    "quase nunca": 0,
    "nunca quase nunca": 0,
    "raramente": 25,
    "as vezes": 50,
    "às vezes": 50,
    "frequentemente": 75,
    "sempre": 100,
}

NEGATIVE_STANDARD: dict[str, int] = {
    "nunca": 100,
    "quase nunca": 100,
    "nunca quase nunca": 100,
    "raramente": 75,
    "as vezes": 50,
    "às vezes": 50,
    "frequentemente": 25,
    "sempre": 0,
}

POSITIVE_INTENSITY: dict[str, int] = {
    "nada quase nada": 0,
    "nada": 0,
    "quase nada": 0,
    "um pouco": 25,
    "moderadamente": 50,
    "muito": 75,
    "extremamente": 100,
}

NEGATIVE_INTENSITY: dict[str, int] = {
    "nada quase nada": 100,
    "nada": 100,
    "quase nada": 100,
    "um pouco": 75,
    "moderadamente": 50,
    "muito": 25,
    "extremamente": 0,
}

POSITIVE_HEALTH: dict[str, int] = {
    "excelente": 100,
    "exelente": 100,
    "muito boa": 75,
    "muito bom": 75,
    "boa": 75,
    "razoavel": 50,
    "deficitaria": 0,
}

NEGATIVE_HEALTH: dict[str, int] = {
    "excelente": 0,
    "exelente": 0,
    "muito boa": 25,
    "muito bom": 25,
    "boa": 25,
    "razoavel": 50,
    "deficitaria": 100,
}

ANSWER_ALIASES: dict[str, str] = {
    "extramamente": "extremamente",
    "muito boa": "muito boa",
    "muito bom": "muito bom",
    "razoavel": "razoavel",
    "exelente": "excelente",
}

METADATA_COLUMNS = {
    "carimbo de data/hora",
    "timestamp",
    "endereco de e-mail",
    "email",
    "nome",
    "id",
}


@dataclass(frozen=True)
class QuestionMatch:
    question_number: int
    question_text: str
    question_type: str
    pillar: str
    scale: str
    matched_column: str | None
    match_score: float
    score_column: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.lower().replace("/", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def identifier_from_text(value: Any) -> str:
    text = normalize_text(value)
    text = text.replace(" ", "_")
    text = re.sub(r"[^a-z0-9_]+", "", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "campo"


def pillar_from_number(question_number: int) -> str | None:
    for start, end, pillar in PILLAR_RANGES:
        if start <= question_number <= end:
            return pillar
    return None


def read_table(file_source: Any) -> pd.DataFrame:
    file_name = getattr(file_source, "name", str(file_source))
    suffix = Path(file_name).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_source, dtype=object)
    return pd.read_excel(file_source, dtype=object)


def find_column(columns: list[str], candidates: list[str]) -> str | None:
    normalized_map = {normalize_text(column): column for column in columns}
    for candidate in candidates:
        candidate_norm = normalize_text(candidate)
        if candidate_norm in normalized_map:
            return normalized_map[candidate_norm]
    for candidate in candidates:
        candidate_norm = normalize_text(candidate)
        for normalized_column, original_column in normalized_map.items():
            if candidate_norm and candidate_norm in normalized_column:
                return original_column
    return None


def load_question_catalog(file_source: Any) -> pd.DataFrame:
    raw = read_table(file_source)
    number_column = find_column(list(raw.columns), ["Nº", "Numero", "Número", "No"])
    question_column = find_column(list(raw.columns), ["Pergunta", "Questão", "Questao"])
    type_column = find_column(list(raw.columns), ["Tipo"])
    pillar_column = find_column(list(raw.columns), ["Pilar/Subescala", "Pilar", "Subescala"])
    scale_column = find_column(list(raw.columns), ["Escala"])

    if question_column is None or type_column is None or scale_column is None:
        missing = [
            name
            for name, column in {
                "Pergunta": question_column,
                "Tipo": type_column,
                "Escala": scale_column,
            }.items()
            if column is None
        ]
        raise ValueError(f"A planilha de perguntas não contém as colunas esperadas: {', '.join(missing)}")

    question_numbers = (
        pd.to_numeric(raw[number_column], errors="coerce")
        if number_column is not None
        else pd.Series(range(1, len(raw) + 1), index=raw.index)
    )
    question_numbers = question_numbers.fillna(pd.Series(range(1, len(raw) + 1), index=raw.index)).astype(int)

    question_catalog = pd.DataFrame(
        {
            "question_number": question_numbers,
            "question_text": raw[question_column].astype(str).str.strip(),
            "question_type": raw[type_column].astype(str).str.strip(),
            "pillar": raw[pillar_column].astype(str).str.strip() if pillar_column is not None else None,
            "scale": raw[scale_column].astype(str).str.strip(),
        }
    )

    question_catalog["pillar"] = question_catalog.apply(
        lambda row: row["pillar"] if isinstance(row["pillar"], str) and row["pillar"].strip() else pillar_from_number(int(row["question_number"])),
        axis=1,
    )
    question_catalog["score_column"] = question_catalog["question_number"].apply(lambda number: f"score_q{int(number):02d}")
    question_catalog["question_key"] = question_catalog["question_text"].map(normalize_text)
    question_catalog["pillar_key"] = question_catalog["pillar"].map(normalize_text)
    return question_catalog.sort_values("question_number").reset_index(drop=True)


def detect_scale_family(scale_text: Any) -> str:
    normalized_scale = normalize_text(scale_text)
    if any(keyword in normalized_scale for keyword in ["excelente", "exelente", "muito boa", "muito bom", "razoavel", "deficitaria"]):
        return "health"
    if "nada quase nada" in normalized_scale or "extremamente" in normalized_scale:
        return "intensity"
    return "standard"


def build_scale_mapping(scale_text: Any, question_type: Any) -> dict[str, int]:
    family = detect_scale_family(scale_text)
    is_positive = normalize_text(question_type).startswith("pos")
    if family == "health":
        return POSITIVE_HEALTH if is_positive else NEGATIVE_HEALTH
    if family == "intensity":
        return POSITIVE_INTENSITY if is_positive else NEGATIVE_INTENSITY
    return POSITIVE_STANDARD if is_positive else NEGATIVE_STANDARD


def score_answer(answer: Any, mapping: dict[str, int]) -> float | pd.NA:
    if answer is None or pd.isna(answer):
        return pd.NA
    normalized_answer = normalize_text(answer)
    if not normalized_answer:
        return pd.NA
    normalized_answer = ANSWER_ALIASES.get(normalized_answer, normalized_answer)
    if normalized_answer in mapping:
        return mapping[normalized_answer]
    for option, score in mapping.items():
        if normalized_answer == option or normalized_answer in option or option in normalized_answer:
            return score
    close_match = None
    best_ratio = 0.0
    for option in mapping:
        ratio = SequenceMatcher(None, normalized_answer, option).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            close_match = option
    if close_match is not None and best_ratio >= 0.88:
        return mapping[close_match]
    return pd.NA


def infer_segment_column(response_frame: pd.DataFrame) -> str | None:
    for column in response_frame.columns:
        normalized_column = normalize_text(column)
        if any(keyword in normalized_column for keyword in ["administrativa", "operacional", "area", "setor"]):
            return column
    for column in response_frame.columns:
        sample_values = [normalize_text(value) for value in response_frame[column].dropna().astype(str).head(50).tolist()]
        if any("administrativa" in value or "operacional" in value for value in sample_values):
            return column
    return None


def match_question_columns(question_catalog: pd.DataFrame, response_frame: pd.DataFrame) -> pd.DataFrame:
    available_columns = [column for column in response_frame.columns if normalize_text(column) not in METADATA_COLUMNS]
    matches: list[QuestionMatch] = []
    used_columns: set[str] = set()

    for row in question_catalog.itertuples(index=False):
        best_column: str | None = None
        best_score = 0.0
        question_text_norm = normalize_text(row.question_text)

        for column in available_columns:
            if column in used_columns:
                continue
            column_norm = normalize_text(column)
            similarity = SequenceMatcher(None, question_text_norm, column_norm).ratio()
            if question_text_norm == column_norm:
                similarity = 1.0
            elif question_text_norm in column_norm or column_norm in question_text_norm:
                similarity = max(similarity, 0.95)
            if similarity > best_score:
                best_score = similarity
                best_column = column

        if best_score < 0.86:
            best_column = None
            best_score = 0.0
        if best_column is not None:
            used_columns.add(best_column)

        matches.append(
            QuestionMatch(
                question_number=int(row.question_number),
                question_text=str(row.question_text),
                question_type=str(row.question_type),
                pillar=str(row.pillar),
                scale=str(row.scale),
                matched_column=best_column,
                match_score=float(best_score),
                score_column=str(row.score_column),
            )
        )

    return pd.DataFrame([match.__dict__ for match in matches])


def build_analysis(
    question_catalog: pd.DataFrame,
    response_frame: pd.DataFrame,
    matches: pd.DataFrame,
    segment_column: str | None = None,
    critical_threshold: float = 40.0,
    attention_threshold: float = 60.0,
) -> dict[str, pd.DataFrame | int | str | None]:
    scored_frame = response_frame.copy()
    question_matches = matches.dropna(subset=["matched_column"]).copy()
    matched_column_by_number = {
        int(row.question_number): row.matched_column
        for row in matches.itertuples(index=False)
        if pd.notna(row.matched_column)
    }

    for row in question_matches.itertuples(index=False):
        catalog_row = question_catalog.loc[question_catalog["question_number"] == int(row.question_number)].iloc[0]
        score_mapping = build_scale_mapping(catalog_row["scale"], catalog_row["question_type"])
        scored_frame[row.score_column] = scored_frame[row.matched_column].map(lambda value: score_answer(value, score_mapping))

    if segment_column is not None and segment_column in scored_frame.columns:
        scored_frame["segmento"] = scored_frame[segment_column].fillna("Não informado").astype(str).str.strip()
    else:
        scored_frame["segmento"] = "Não segmentado"

    pillar_rows: list[dict[str, Any]] = []
    question_rows: list[dict[str, Any]] = []
    pillar_score_columns: dict[str, str] = {}
    pillar_score_series: dict[str, pd.Series] = {}

    for pillar, pillar_questions in question_catalog.groupby("pillar", dropna=False):
        pillar_label = str(pillar) if pillar and str(pillar).strip() else "Sem pilar"
        score_column = f"pillar_{identifier_from_text(pillar_label)}"
        pillar_score_columns[pillar_label] = score_column
        pillar_question_score_columns = [
            row.score_column
            for row in pillar_questions.itertuples(index=False)
            if row.score_column in scored_frame.columns
        ]
        if pillar_question_score_columns:
            pillar_score_series[score_column] = scored_frame[pillar_question_score_columns].mean(axis=1)
        else:
            pillar_score_series[score_column] = pd.Series(pd.NA, index=scored_frame.index)

        pillar_series = pillar_score_series[score_column]
        pillar_mean = float(pillar_series.mean()) if pillar_series.notna().any() else float("nan")
        if pd.isna(pillar_mean):
            pillar_status = "Sem dados"
        elif pillar_mean < critical_threshold:
            pillar_status = "Crítico"
        elif pillar_mean < attention_threshold:
            pillar_status = "Atenção"
        else:
            pillar_status = "Saudável"

        pillar_rows.append(
            {
                "pilar": pillar_label,
                "perguntas": len(pillar_questions),
                "respondentes_validos": int(pillar_series.notna().sum()),
                "media": pillar_mean,
                "mediana": float(pillar_series.median()) if pillar_series.notna().any() else float("nan"),
                "status": pillar_status,
                "coluna_score": score_column,
            }
        )

    if pillar_score_series:
        scored_frame = pd.concat([scored_frame, pd.DataFrame(pillar_score_series, index=scored_frame.index)], axis=1)

    for row in question_catalog.itertuples(index=False):
        score_column = row.score_column
        series = scored_frame[score_column] if score_column in scored_frame.columns else pd.Series(dtype=float)
        mean_score = float(series.mean()) if series.notna().any() else float("nan")
        critical_rate = float((series < critical_threshold).mean()) if series.notna().any() else float("nan")
        attention_rate = float(((series >= critical_threshold) & (series < attention_threshold)).mean()) if series.notna().any() else float("nan")
        question_rows.append(
            {
                "numero": int(row.question_number),
                "pergunta": str(row.question_text),
                "pilar": str(row.pillar) if row.pillar else "Sem pilar",
                "tipo": str(row.question_type),
                "media": mean_score,
                "taxa_critica": critical_rate,
                "taxa_atencao": attention_rate,
                "respostas_validas": int(series.notna().sum()),
                "coluna_resposta": matched_column_by_number.get(int(row.question_number)),
            }
        )

    pillar_summary = pd.DataFrame(pillar_rows).sort_values("media", ascending=True, na_position="last").reset_index(drop=True)
    question_summary = pd.DataFrame(question_rows).sort_values(["media", "numero"], ascending=[True, True], na_position="last").reset_index(drop=True)

    segment_summary = pd.DataFrame()
    if "segmento" in scored_frame.columns:
        segment_columns = [column for column in pillar_score_columns.values() if column in scored_frame.columns]
        if segment_columns:
            segment_summary = (
                scored_frame.groupby("segmento", dropna=False)[segment_columns]
                .mean()
                .reset_index()
                .rename(columns={value: key for key, value in pillar_score_columns.items()})
            )

    response_count = len(scored_frame)
    matched_count = int(question_matches["matched_column"].notna().sum())
    unmatched_count = int(len(question_catalog) - matched_count)

    return {
        "scored_frame": scored_frame,
        "pillar_summary": pillar_summary,
        "question_summary": question_summary,
        "segment_summary": segment_summary,
        "response_count": response_count,
        "matched_count": matched_count,
        "unmatched_count": unmatched_count,
        "matches": matches,
        "pillar_score_columns": pillar_score_columns,
        "segment_column": segment_column,
    }


def dataframe_to_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8-sig")


def build_answer_distribution(response_frame: pd.DataFrame, response_column: str) -> pd.DataFrame:
    if response_column not in response_frame.columns:
        return pd.DataFrame(columns=["resposta", "quantidade", "percentual"])

    answer_series = response_frame[response_column].dropna().astype(str).str.strip()
    if answer_series.empty:
        return pd.DataFrame(columns=["resposta", "quantidade", "percentual"])

    distribution = answer_series.value_counts().reset_index()
    distribution.columns = ["resposta", "quantidade"]
    total = float(distribution["quantidade"].sum())
    distribution["percentual"] = distribution["quantidade"].div(total).mul(100)
    return distribution


def _format_markdown_table(dataframe: pd.DataFrame, columns: list[str], max_rows: int | None = None) -> str:
    subset = dataframe.loc[:, columns].copy()
    if max_rows is not None:
        subset = subset.head(max_rows)

    if subset.empty:
        return "_Sem dados para exibir._"

    headers = columns
    separator = ["---" for _ in columns]
    rows = ["| " + " | ".join(headers) + " |", "| " + " | ".join(separator) + " |"]

    for _, row in subset.iterrows():
        values: list[str] = []
        for column in columns:
            value = row[column]
            if pd.isna(value):
                text = "-"
            elif isinstance(value, float):
                text = f"{value:.1f}"
            elif isinstance(value, int):
                text = str(value)
            else:
                text = str(value)
            values.append(text.replace("|", "\\|"))
        rows.append("| " + " | ".join(values) + " |")

    return "\n".join(rows)


def build_recommendations(pillar_summary: pd.DataFrame, critical_threshold: float) -> list[dict[str, str]]:
    recommendations: list[dict[str, str]] = []
    critical_pillars = pillar_summary.loc[pillar_summary["media"] < critical_threshold, "pilar"].head(5).tolist()
    critical_text = " ".join(critical_pillars).lower()

    rules = [
        (
            ["liderança", "lideranca", "confiança vertical", "confianca vertical", "apoio social de superiores"],
            "Reforçar a qualidade da liderança, os feedbacks e a disponibilidade dos gestores para apoio ao time.",
            "Liderança e suporte gerencial",
        ),
        (
            ["conflitos laborais", "transparência no papel laboral desempenhado", "transparencia no papel laboral desempenhado", "recompensas"],
            "Revisar papéis, alinhamento de expectativas, reconhecimento e resolução de conflitos operacionais.",
            "Papéis, reconhecimento e conflitos",
        ),
        (
            ["exigências quantitativas", "exigencias quantitativas", "ritmo de trabalho", "exigências cognitivas", "exigencias cognitivas"],
            "Ajustar carga de trabalho, priorização e ritmo de execução para reduzir sobrecarga.",
            "Sobrecarga e ritmo",
        ),
        (
            ["burnout", "stress", "sintomas depressivos", "saúde geral", "saude geral", "conflito trabalho/família"],
            "Avaliar ações de saúde mental, descanso, equilíbrio entre vida pessoal e trabalho e suporte psicossocial.",
            "Saúde mental e equilíbrio",
        ),
    ]

    for keywords, recommendation, label in rules:
        if any(keyword in critical_text for keyword in keywords):
            recommendations.append({"label": label, "text": recommendation})

    if not recommendations:
        recommendations.append({"label": "Monitoramento contínuo", "text": "Manter o monitoramento periódico e acompanhar os pilares com tendência de queda."})

    return recommendations


def build_pillar_recommendations(
    pillar_summary: pd.DataFrame,
    critical_threshold: float,
    attention_threshold: float,
) -> list[dict[str, str]]:
    recommendations: list[dict[str, str]] = []

    for row in pillar_summary.itertuples(index=False):
        pillar = str(row.pilar)
        score = float(row.media) if row.media == row.media else float("nan")
        if score != score:
            continue

        status = str(row.status)
        if status == "Saudável":
            continue

        pillar_key = normalize_text(pillar)
        if any(keyword in pillar_key for keyword in ["lideranca", "liderança", "apoio social de superiores", "confianca vertical", "confiança vertical"]):
            text = "Reforçar a presença da liderança, ampliar feedbacks frequentes e garantir suporte direto ao time."
        elif any(keyword in pillar_key for keyword in ["conflitos laborais", "recompensas", "transparencia no papel laboral desempenhado", "transparência no papel laboral desempenhado"]):
            text = "Revisar alinhamento de expectativas, reconhecimento e clareza dos papéis para reduzir ruído operacional."
        elif any(keyword in pillar_key for keyword in ["exigencias quantitativas", "exigências quantitativas", "ritmo de trabalho", "exigencias cognitivas", "exigências cognitivas"]):
            text = "Revisar carga de trabalho, priorização e cadência de entrega para reduzir sobrecarga percebida."
        elif any(keyword in pillar_key for keyword in ["burnout", "stress", "sintomas depressivos", "saude geral", "saúde geral", "conflito trabalho/família"]):
            text = "Priorizar ações de saúde mental, pausas, equilíbrio vida-trabalho e monitoramento do bem-estar."
        elif any(keyword in pillar_key for keyword in ["apoio social de colegas", "colegas", "confiança horizontal", "confianca horizontal"]):
            text = "Estimular colaboração entre pares, rituais de equipe e canais claros de suporte entre colegas."
        elif any(keyword in pillar_key for keyword in ["justiça e respeito", "justica e respeito", "transparencia", "transparência"]):
            text = "Reforçar comunicação transparente, critérios de decisão e tratamento consistente entre áreas."
        else:
            text = "Definir uma ação corretiva específica, acompanhar a tendência mensal e reavaliar a percepção do grupo."

        recommendations.append(
            {
                "pilar": pillar,
                "status": status,
                "media": f"{score:.1f}",
                "label": f"{pillar} ({status})",
                "text": text,
            }
        )

    if not recommendations:
        recommendations.append(
            {
                "pilar": "Monitoramento contínuo",
                "status": "Saudável",
                "media": "-",
                "label": "Monitoramento contínuo",
                "text": "Nenhum pilar crítico ou em atenção no recorte atual. Manter acompanhamento periódico.",
            }
        )

    return recommendations


def build_recommendation_explanation(critical_threshold: float) -> str:
    return (
        "As recomendações são geradas por regras fixas do próprio BI: o sistema avalia cada pilar, destaca os que estão abaixo "
        f"de {critical_threshold:.0f} e associa cada um a uma ação padronizada de RH. Não há IA inferencial nem benchmark externo; "
        "o texto é uma camada de orientação baseada nos temas do COPSOQ/NR01 que apareceram como críticos ou em atenção."
    )


def _summary_status_counts(pillar_summary: pd.DataFrame, critical_threshold: float, attention_threshold: float) -> pd.DataFrame:
    rows = [
        {"status": "Crítico", "quantidade": int((pillar_summary["media"] < critical_threshold).sum())},
        {
            "status": "Atenção",
            "quantidade": int(((pillar_summary["media"] >= critical_threshold) & (pillar_summary["media"] < attention_threshold)).sum()),
        },
        {"status": "Saudável", "quantidade": int((pillar_summary["media"] >= attention_threshold).sum())},
    ]
    return pd.DataFrame(rows)


def _chart_vertical_bar(dataframe: pd.DataFrame, title: str, category_col: str, value_col: str, color: colors.Color) -> Drawing:
    drawing = Drawing(520, 240)
    chart = VerticalBarChart()
    chart.x = 50
    chart.y = 40
    chart.height = 150
    chart.width = 420
    chart.data = [list(dataframe[value_col].fillna(0).astype(float))]
    chart.categoryAxis.categoryNames = [str(value) for value in dataframe[category_col].tolist()]
    chart.categoryAxis.labels.angle = 30
    chart.categoryAxis.labels.dy = -12
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = 100
    chart.valueAxis.valueStep = 20
    chart.bars[0].fillColor = color
    chart.bars[0].strokeColor = color
    chart.barLabelFormat = "%.0f"
    chart.barLabels.nudge = 7
    chart.barLabels.fontSize = 8
    drawing.add(chart)
    drawing.add(String(50, 210, title, fontSize=12, fillColor=colors.HexColor("#0f172a")))
    return drawing


def _chart_horizontal_bar(dataframe: pd.DataFrame, title: str, category_col: str, value_col: str, color: colors.Color) -> Drawing:
    drawing = Drawing(520, 260)
    chart = HorizontalBarChart()
    chart.x = 170
    chart.y = 30
    chart.height = 170
    chart.width = 300
    chart.data = [list(dataframe[value_col].fillna(0).astype(float))]
    chart.categoryAxis.categoryNames = [str(value) for value in dataframe[category_col].tolist()]
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = 100
    chart.valueAxis.valueStep = 20
    chart.bars[0].fillColor = color
    chart.bars[0].strokeColor = color
    chart.barLabelFormat = "%.0f"
    chart.barLabels.nudge = 4
    chart.barLabels.fontSize = 8
    drawing.add(chart)
    drawing.add(String(20, 220, title, fontSize=12, fillColor=colors.HexColor("#0f172a")))
    return drawing


def _chart_pie(dataframe: pd.DataFrame, title: str, label_col: str, value_col: str) -> Drawing:
    drawing = Drawing(280, 220)
    pie = Pie()
    pie.x = 40
    pie.y = 10
    pie.width = 150
    pie.height = 150
    pie.data = list(dataframe[value_col].fillna(0).astype(float))
    pie.labels = [str(value) for value in dataframe[label_col].tolist()]
    pie.slices.strokeWidth = 0.5
    pie.slices[0].fillColor = colors.HexColor("#b91c1c")
    if len(pie.slices) > 1:
        pie.slices[1].fillColor = colors.HexColor("#d97706")
    if len(pie.slices) > 2:
        pie.slices[2].fillColor = colors.HexColor("#15803d")
    drawing.add(pie)
    drawing.add(String(10, 185, title, fontSize=11, fillColor=colors.HexColor("#0f172a")))
    return drawing


def _metric_card_table(title: str, value: str, accent: str) -> Table:
    card = Table([[Paragraph(f"<b>{title}</b>", ParagraphStyle("CardTitle", fontSize=9, textColor=colors.white, alignment=TA_CENTER)),
                   Paragraph(f"<font size='18'><b>{value}</b></font>", ParagraphStyle("CardValue", fontSize=18, textColor=colors.white, alignment=TA_CENTER))]],
                 colWidths=[3.0 * cm, 3.0 * cm], rowHeights=[1.35 * cm])
    card.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(accent)),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(accent)),
                ("INNERGRID", (0, 0), (-1, -1), 0.0, colors.HexColor(accent)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return card


def _page_decor(canvas, document) -> None:
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(colors.HexColor("#0f172a"))
    canvas.rect(0, height - 1.2 * cm, width, 1.2 * cm, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor("#e2e8f0"))
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawString(1.25 * cm, height - 0.8 * cm, "BI NR01 | COPSOQ")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(width - 1.25 * cm, height - 0.8 * cm, f"Página {document.page}")
    canvas.setStrokeColor(colors.HexColor("#cbd5e1"))
    canvas.line(1.25 * cm, 1.0 * cm, width - 1.25 * cm, 1.0 * cm)
    canvas.setFillColor(colors.HexColor("#64748b"))
    canvas.setFont("Helvetica", 7)
    canvas.drawString(1.25 * cm, 0.68 * cm, "Relatório gerado automaticamente a partir das respostas do Forms")
    canvas.restoreState()


def build_rh_report_pdf_bytes(
    *,
    response_count: int,
    critical_threshold: float,
    attention_threshold: float,
    pillar_summary: pd.DataFrame,
    question_summary: pd.DataFrame,
    segment_summary: pd.DataFrame,
    segment_column: str | None,
    selected_segment: str | None,
    application_start_date: str | None = None,
    total_collaborators: int | None = None,
    adherence_percentage: float | None = None,
) -> bytes:
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.3 * cm,
        rightMargin=1.3 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("ReportTitle", parent=styles["Title"], alignment=TA_CENTER, textColor=colors.HexColor("#0f172a"), fontSize=22, leading=26, spaceAfter=6)
    subtitle_style = ParagraphStyle("ReportSubtitle", parent=styles["Normal"], alignment=TA_CENTER, textColor=colors.HexColor("#475569"), fontSize=10, leading=13)
    section_style = ParagraphStyle("Section", parent=styles["Heading2"], textColor=colors.HexColor("#0f172a"), spaceBefore=8, spaceAfter=6)
    body_style = ParagraphStyle("Body", parent=styles["BodyText"], alignment=TA_LEFT, leading=13, spaceAfter=4)
    small_center = ParagraphStyle("SmallCenter", parent=styles["BodyText"], alignment=TA_CENTER, leading=12, textColor=colors.HexColor("#334155"))

    overall_average = float(pillar_summary["media"].mean()) if not pillar_summary.empty else float("nan")
    critical_pillars = int((pillar_summary["media"] < critical_threshold).sum()) if not pillar_summary.empty else 0
    attention_pillars = int(((pillar_summary["media"] >= critical_threshold) & (pillar_summary["media"] < attention_threshold)).sum()) if not pillar_summary.empty else 0
    healthy_pillars = int((pillar_summary["media"] >= attention_threshold).sum()) if not pillar_summary.empty else 0

    recommendations = build_pillar_recommendations(pillar_summary, critical_threshold, attention_threshold)
    recommendation_explanation = build_recommendation_explanation(critical_threshold)

    elements: list[Any] = []
    elements.append(Paragraph("Relatório executivo NR01 / COPSOQ", title_style))
    elements.append(Paragraph("Visão resumida e visual para suporte à leitura do RH", subtitle_style))
    elements.append(Spacer(1, 0.2 * cm))
    context_parts = [
        f"<b>Respostas analisadas:</b> {response_count}",
        f"<b>Segmentação:</b> {segment_column or 'não informada'}",
        f"<b>Filtro visual:</b> {selected_segment or 'todas as áreas'}",
    ]
    if application_start_date:
        context_parts.append(f"<b>Início da aplicação:</b> {application_start_date}")
    if total_collaborators is not None:
        context_parts.append(f"<b>Total de colaboradores:</b> {int(total_collaborators)}")
    if adherence_percentage is not None:
        context_parts.append(f"<b>Adesão:</b> {adherence_percentage:.1f}%")
    elements.append(Paragraph(" | ".join(context_parts), small_center))
    elements.append(Spacer(1, 0.25 * cm))

    metric_cards = Table([
        [
            _metric_card_table("Respondentes", str(response_count), "#0f766e"),
            _metric_card_table("Média geral", f"{overall_average:.1f}" if overall_average == overall_average else "-", "#2563eb"),
            _metric_card_table("Críticos", str(critical_pillars), "#b91c1c"),
            _metric_card_table("Atenção", str(attention_pillars), "#d97706"),
            _metric_card_table("Saudáveis", str(healthy_pillars), "#15803d"),
        ]
    ], colWidths=[3.45 * cm, 3.45 * cm, 3.45 * cm, 3.45 * cm, 3.45 * cm])
    metric_cards.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    elements.append(metric_cards)
    elements.append(Spacer(1, 0.28 * cm))

    elements.append(Paragraph("Resumo visual", section_style))

    if not pillar_summary.empty:
        top_pillars = pillar_summary.sort_values("media", ascending=False).head(8).sort_values("media", ascending=True)
        elements.append(_chart_vertical_bar(top_pillars, "Média dos pilares", "pilar", "media", colors.HexColor("#0f766e")))
        elements.append(Spacer(1, 0.25 * cm))

        status_counts = _summary_status_counts(pillar_summary, critical_threshold, attention_threshold)
        elements.append(_chart_pie(status_counts, "Status dos pilares", "status", "quantidade"))
        elements.append(Spacer(1, 0.25 * cm))

    if not question_summary.empty:
        top_questions = question_summary.sort_values("media", ascending=True).head(8).sort_values("media", ascending=True)
        elements.append(_chart_horizontal_bar(top_questions, "Menores médias por pergunta", "numero", "media", colors.HexColor("#b91c1c")))
        elements.append(Spacer(1, 0.2 * cm))

    if not segment_summary.empty and "segmento" in segment_summary.columns:
        segment_means = segment_summary.set_index("segmento").mean(axis=1).sort_values(ascending=True)
        segment_frame = pd.DataFrame({"segmento": segment_means.index.tolist(), "media": segment_means.values.tolist()})
        elements.append(_chart_horizontal_bar(segment_frame, "Média geral por área", "segmento", "media", colors.HexColor("#2563eb")))
        elements.append(Spacer(1, 0.2 * cm))

    elements.append(PageBreak())
    elements.append(Paragraph("Leitura executiva", section_style))
    elements.append(Paragraph(recommendation_explanation, body_style))
    elements.append(Spacer(1, 0.15 * cm))

    recommendation_rows = [[rec["label"], rec["status"], rec["media"], rec["text"]] for rec in recommendations]
    recommendation_table = Table(recommendation_rows, colWidths=[5.2 * cm, 2.0 * cm, 1.8 * cm, 8.0 * cm])
    recommendation_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cbd5e1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#f8fafc"), colors.white]),
                ("ALIGN", (1, 1), (2, -1), "CENTER"),
            ]
        )
    )
    elements.append(recommendation_table)
    elements.append(Spacer(1, 0.15 * cm))

    top_critical = pillar_summary.sort_values("media", ascending=True).head(5)
    critical_rows = [[row.pilar, f"{row.media:.1f}", row.status] for row in top_critical.itertuples(index=False)]
    critical_table = Table([["Pilar", "Média", "Status"]] + critical_rows, colWidths=[8.0 * cm, 2.5 * cm, 3.0 * cm])
    critical_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cbd5e1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#fff7ed"), colors.white]),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    elements.append(Spacer(1, 0.2 * cm))
    elements.append(Paragraph("Pilares com maior prioridade", section_style))
    elements.append(critical_table)

    document.build(elements, onFirstPage=_page_decor, onLaterPages=_page_decor)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def build_rh_report_markdown(
    *,
    response_count: int,
    critical_threshold: float,
    attention_threshold: float,
    pillar_summary: pd.DataFrame,
    question_summary: pd.DataFrame,
    segment_summary: pd.DataFrame,
    segment_column: str | None,
    selected_segment: str | None,
    application_start_date: str | None = None,
    total_collaborators: int | None = None,
    adherence_percentage: float | None = None,
) -> str:
    lines: list[str] = []
    lines.append("# Relatório executivo NR01 / COPSOQ")
    lines.append("")
    lines.append("## Escopo")
    lines.append(f"- Respostas analisadas: {response_count}")
    lines.append(f"- Segmentação disponível: {segment_column or 'não informada'}")
    lines.append(f"- Filtro de área aplicado na visualização: {selected_segment or 'todas as áreas'}")
    lines.append(f"- Data de início da aplicação: {application_start_date or 'não informada'}")
    lines.append(f"- Total de colaboradores: {total_collaborators if total_collaborators is not None else 'não informado'}")
    if adherence_percentage is not None:
        lines.append(f"- Percentual de adesão: {adherence_percentage:.1f}%")
    lines.append("")
    lines.append("## Metodologia")
    lines.append("- O score varia de 0 a 100.")
    lines.append("- Quanto maior o score, melhor o indicador.")
    lines.append(f"- Scores abaixo de {critical_threshold:.0f} indicam criticidade e entre {critical_threshold:.0f} e {attention_threshold:.0f} indicam atenção.")
    lines.append("")

    overall_average = float(pillar_summary["media"].mean()) if not pillar_summary.empty else float("nan")
    critical_pillars = int((pillar_summary["media"] < critical_threshold).sum()) if not pillar_summary.empty else 0
    attention_pillars = int(((pillar_summary["media"] >= critical_threshold) & (pillar_summary["media"] < attention_threshold)).sum()) if not pillar_summary.empty else 0
    critical_questions = int((question_summary["media"] < critical_threshold).sum()) if not question_summary.empty else 0

    lines.append("## Visão geral")
    lines.append(f"- Média geral dos pilares: {overall_average:.1f}" if overall_average == overall_average else "- Média geral dos pilares: sem dados")
    lines.append(f"- Pilares críticos: {critical_pillars}")
    lines.append(f"- Pilares em atenção: {attention_pillars}")
    lines.append(f"- Perguntas críticas: {critical_questions}")
    lines.append("")

    lines.append("## Pilares com menor score")
    lines.append(_format_markdown_table(pillar_summary, ["pilar", "media", "status", "perguntas"], max_rows=5))
    lines.append("")

    lines.append("## Todas as perguntas")
    lines.append(_format_markdown_table(question_summary, ["numero", "pilar", "media", "taxa_critica"], max_rows=None))
    lines.append("")

    if not segment_summary.empty:
        segment_means = segment_summary.set_index("segmento").mean(axis=1).sort_values(ascending=False)
        lines.append("## Comparativo por área")
        comparison = pd.DataFrame({"area": segment_means.index, "media_geral": segment_means.values})
        lines.append(_format_markdown_table(comparison, ["area", "media_geral"], max_rows=None))
        lines.append("")

    recommendations = build_pillar_recommendations(pillar_summary, critical_threshold, attention_threshold)

    lines.append("## Recomendações críticas e de atenção")
    for recommendation in recommendations:
        lines.append(f"- {recommendation['label']}: {recommendation['text']}")

    return "\n".join(lines)


def _plotly_div_from_figure(fig: go.Figure, *, include_plotlyjs: bool = False) -> str:
    return pio.to_html(
        fig,
        include_plotlyjs="inline" if include_plotlyjs else False,
        full_html=False,
        config={"displayModeBar": False},
    )


def _plotly_pillar_bar_div(pillar_summary: pd.DataFrame, *, include_plotlyjs: bool = False) -> str:
    if pillar_summary.empty:
        return "<div></div>"
    df = pillar_summary.sort_values("media", ascending=True).copy()
    color_map = {
        "Crítico": "#b91c1c",
        "Atenção": "#d97706",
        "Saudável": "#15803d",
        "Sem dados": "#64748b",
    }
    colors = [color_map.get(str(status), "#64748b") for status in df["status"].tolist()]
    fig = go.Figure(
        go.Bar(
            x=df["pilar"].tolist(),
            y=df["media"].tolist(),
            marker_color=colors,
            text=[f"{float(value):.0f}" if value == value else "-" for value in df["media"].tolist()],
            textposition="outside",
        )
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=20, b=165),
        height=380,
        yaxis=dict(range=[0, 100], title="Média"),
        xaxis=dict(title="Pilares", tickangle=-30, automargin=True),
    )
    return _plotly_div_from_figure(fig, include_plotlyjs=include_plotlyjs)


def _plotly_status_pie_div(pillar_summary: pd.DataFrame, critical_threshold: float, attention_threshold: float, *, include_plotlyjs: bool = False) -> str:
    if pillar_summary.empty:
        return "<div></div>"
    critical = int((pillar_summary["media"] < critical_threshold).sum())
    attention = int(((pillar_summary["media"] >= critical_threshold) & (pillar_summary["media"] < attention_threshold)).sum())
    healthy = int((pillar_summary["media"] >= attention_threshold).sum())
    labels = ["Crítico", "Atenção", "Saudável"]
    values = [critical, attention, healthy]
    fig = go.Figure(
        go.Pie(
            labels=labels,
            values=values,
            hole=0.58,
            marker_colors=["#b91c1c", "#d97706", "#15803d"],
            textinfo="percent",
            textposition="inside",
            showlegend=False,
        )
    )
    fig.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=300)
    return _plotly_div_from_figure(fig, include_plotlyjs=include_plotlyjs)


def _plotly_segment_heatmap_div(segment_summary: pd.DataFrame, *, include_plotlyjs: bool = False) -> str:
    if segment_summary.empty or "segmento" not in segment_summary.columns:
        return ""

    heatmap_df = segment_summary.set_index("segmento").select_dtypes(include=["number"])
    if heatmap_df.empty:
        return ""

    wrapped_y = ["<br>".join(textwrap.wrap(str(item), width=18)) for item in heatmap_df.index.tolist()]
    wrapped_x = ["<br>".join(textwrap.wrap(str(item), width=16)) for item in heatmap_df.columns.tolist()]

    fig = go.Figure(
        data=go.Heatmap(
            z=heatmap_df.values.tolist(),
            x=wrapped_x,
            y=wrapped_y,
            colorscale="RdYlGn",
            zmin=0,
            zmax=100,
            hoverongaps=False,
            colorbar=dict(title="Score", thickness=14, len=0.9),
        )
    )
    fig.update_layout(
        margin=dict(l=170, r=40, t=20, b=185),
        height=430,
        xaxis=dict(title="Pilares", tickangle=-25, automargin=True),
        yaxis=dict(title="Segmentos", automargin=True),
    )
    return _plotly_div_from_figure(fig, include_plotlyjs=include_plotlyjs)


def _plotly_top_questions_div(question_summary: pd.DataFrame, *, include_plotlyjs: bool = False) -> str:
    if question_summary.empty:
        return "<div></div>"
    df = question_summary.sort_values("media", ascending=True).head(12)
    fig = go.Figure(go.Bar(x=df["media"].tolist(), y=[str(n) + ". " + q for n, q in zip(df["numero"].tolist(), df["pergunta"].tolist())], orientation="h", marker_color="#b91c1c"))
    fig.update_layout(margin=dict(l=10, r=10, t=20, b=10), height=360, xaxis=dict(range=[0, 100], title="Média"))
    return _plotly_div_from_figure(fig, include_plotlyjs=include_plotlyjs)


def _preferred_answer_order(scale_text: Any) -> list[str]:
    family = detect_scale_family(scale_text)
    if family == "health":
        return ["deficitária", "razoável", "boa", "muito boa", "excelente"]
    if family == "intensity":
        return ["nada", "quase nada", "um pouco", "moderadamente", "muito", "extremamente"]
    return ["nunca/quase nunca", "quase nunca", "nunca", "raramente", "às vezes", "as vezes", "frequentemente", "sempre"]


def _normalize_answer_label(label: Any) -> str:
    return normalize_text(label)


def _ordered_distribution(distribution: pd.DataFrame, scale_text: Any) -> pd.DataFrame:
    ordered = distribution.copy()
    order_map = {normalize_text(item): index for index, item in enumerate(_preferred_answer_order(scale_text))}
    ordered["_ord"] = ordered["resposta"].map(lambda value: order_map.get(_normalize_answer_label(value), 999))
    return ordered.sort_values(["_ord", "quantidade"], ascending=[True, False]).drop(columns=["_ord"])


def _classify_segment_for_report(value: Any) -> str | None:
    normalized = normalize_text(value)
    if "admin" in normalized:
        return "Administrativo"
    if "operac" in normalized:
        return "Operacional"
    return None


def _segment_panel_html(segment_label: str, distribution: pd.DataFrame, scale_text: Any) -> str:
    if distribution.empty:
        return (
            "<div class='segment-panel'>"
            f"<div class='segment-title'>{segment_label}</div>"
            "<div class='segment-empty'>Sem respostas para este segmento.</div>"
            "</div>"
        )

    ordered = _ordered_distribution(distribution, scale_text)

    palette = ["#4f6bed", "#e91e8f", "#2ca6a4", "#8a63d2", "#2eb82e", "#f59e0b", "#ef4444"]
    fig = go.Figure(
        go.Pie(
            labels=ordered["resposta"].tolist(),
            values=ordered["quantidade"].fillna(0).astype(float).tolist(),
            hole=0.68,
            sort=False,
            direction="clockwise",
            marker=dict(colors=palette[: len(ordered)]),
            textinfo="none",
            showlegend=False,
            hoverinfo="skip",
        )
    )
    fig.update_layout(margin=dict(l=0, r=0, t=8, b=8), height=220, width=220)
    chart_html = _plotly_div_from_figure(fig)

    legend_items = []
    total = float(ordered["quantidade"].fillna(0).sum()) or 1.0
    for index, row in enumerate(ordered.itertuples(index=False)):
        color = palette[index % len(palette)]
        percentage = (float(row.quantidade) / total) * 100 if total else 0.0
        legend_items.append(
            f"<div class='legend-row'><span class='legend-dot' style='background:{color}'></span>"
            f"<span class='legend-label'>{row.resposta}</span>"
            f"<span class='legend-value'>{int(row.quantidade)}</span>"
            f"<span class='legend-pct'>{percentage:.0f}%</span></div>"
        )

    return f"""
    <div class='segment-panel'>
        <div class='segment-title'>{segment_label}</div>
        <div class='segment-body'>
        <div class='segment-legend'>
            {''.join(legend_items)}
        </div>
        <div class='segment-chart'>
            {chart_html}
        </div>
        </div>
    </div>
    """


def _question_card_html(
    question_number: int,
    question_text: str,
    pillar: str,
    administrative_distribution: pd.DataFrame,
    operational_distribution: pd.DataFrame,
    scale_text: Any,
) -> str:
    return f"""
    <div class='question-card'>
        <div class='question-title'>
        <div class='question-number'>{question_number}. {question_text}</div>
        <div class='question-pillar'>{pillar}</div>
        </div>
        <div class='segments-grid'>
            {_segment_panel_html('Administrativo', administrative_distribution, scale_text)}
            {_segment_panel_html('Operacional', operational_distribution, scale_text)}
        </div>
    </div>
    """


def _question_cards_html(
    response_frame: pd.DataFrame,
    matches: pd.DataFrame,
    question_summary: pd.DataFrame,
    question_catalog: pd.DataFrame,
    segment_column: str | None,
    limit: int | None = None,
) -> str:
    if question_summary.empty:
        return "<div class='empty-state'>Sem perguntas para exibir.</div>"

    selected = question_summary.sort_values("numero", ascending=True)
    if limit is not None:
        selected = selected.head(limit)
    cards: list[str] = []
    catalog_index = question_catalog.set_index("question_number")

    segmented_frame = response_frame.copy()
    if segment_column and segment_column in segmented_frame.columns:
        segmented_frame["_segmento_relatorio"] = segmented_frame[segment_column].map(_classify_segment_for_report)
    else:
        segmented_frame["_segmento_relatorio"] = None

    for row in selected.itertuples(index=False):
        match_row = matches.loc[matches["question_number"] == int(row.numero)]
        if match_row.empty:
            continue
        matched_column = match_row.iloc[0]["matched_column"]
        if not matched_column or matched_column not in response_frame.columns:
            continue
        administrative_distribution = build_answer_distribution(
            segmented_frame[segmented_frame["_segmento_relatorio"] == "Administrativo"],
            matched_column,
        )
        operational_distribution = build_answer_distribution(
            segmented_frame[segmented_frame["_segmento_relatorio"] == "Operacional"],
            matched_column,
        )

        if administrative_distribution.empty and operational_distribution.empty:
            overall_distribution = build_answer_distribution(response_frame, matched_column)
            administrative_distribution = overall_distribution.copy()
            operational_distribution = overall_distribution.copy()

        catalog_row = catalog_index.loc[int(row.numero)] if int(row.numero) in catalog_index.index else None
        scale_text = catalog_row["scale"] if catalog_row is not None else ""
        cards.append(
            _question_card_html(
                int(row.numero),
                str(row.pergunta),
                str(row.pilar),
                administrative_distribution,
                operational_distribution,
                scale_text,
            )
        )

    if not cards:
        return "<div class='empty-state'>Nenhuma distribuição de respostas disponível para as perguntas selecionadas.</div>"

    pages: list[str] = []
    for index in range(0, len(cards), 4):
        page_cards = "\n".join(cards[index:index + 4])
        page_class = "question-page question-page-break" if index + 4 < len(cards) else "question-page"
        pages.append(f"<div class='{page_class}'>{page_cards}</div>")

    return "\n".join(pages)


def build_rh_report_html(
        *,
        response_count: int,
        critical_threshold: float,
        attention_threshold: float,
        pillar_summary: pd.DataFrame,
        question_summary: pd.DataFrame,
        segment_summary: pd.DataFrame,
        segment_column: str | None,
        selected_segment: str | None,
    application_start_date: str | None = None,
    total_collaborators: int | None = None,
    adherence_percentage: float | None = None,
    response_frame: pd.DataFrame | None = None,
    matches: pd.DataFrame | None = None,
    question_catalog: pd.DataFrame | None = None,
) -> str:
    overall_average = float(pillar_summary["media"].mean()) if not pillar_summary.empty else float("nan")
    critical_pillars = int((pillar_summary["media"] < critical_threshold).sum()) if not pillar_summary.empty else 0
    attention_pillars = int(((pillar_summary["media"] >= critical_threshold) & (pillar_summary["media"] < attention_threshold)).sum()) if not pillar_summary.empty else 0
    healthy_pillars = int((pillar_summary["media"] >= attention_threshold).sum()) if not pillar_summary.empty else 0

    status_has_data = not pillar_summary.empty
    status_div = _plotly_status_pie_div(pillar_summary, critical_threshold, attention_threshold, include_plotlyjs=True)
    pillar_div = _plotly_pillar_bar_div(pillar_summary)
    heatmap_div = _plotly_segment_heatmap_div(segment_summary)
    question_cards = ""
    if response_frame is not None and matches is not None and question_catalog is not None:
        question_cards = _question_cards_html(
            response_frame,
            matches,
            question_summary,
            question_catalog,
            segment_column=segment_column,
            limit=None,
        )

    recommendations = build_pillar_recommendations(pillar_summary, critical_threshold, attention_threshold)

    template = Template("""
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
            body{font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial; color:#0f172a; margin:18px; background:#f8fafc}
            .header{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
            .title{font-size:20px;font-weight:700}
            .subtitle{color:#475569}
            .cover{margin-top:12px;background:linear-gradient(135deg,#0f172a 0%,#1e293b 45%,#0f766e 100%);color:#fff;border-radius:24px;padding:20px;box-shadow:0 18px 40px rgba(15,23,42,.18)}
            .cover-top{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:stretch}
            .cover-title{font-size:28px;line-height:1.1;font-weight:800;letter-spacing:-.03em;margin-bottom:8px}
            .cover-metrics{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-top:12px}
            .cover-metric{background:rgba(255,255,255,.09);border:1px solid rgba(255,255,255,.12);border-radius:16px;padding:12px 14px}
            .cover-metric-label{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#bfdbfe}
            .cover-metric-value{font-size:22px;font-weight:800;margin-top:4px}
            .cover-panel{background:#fff;color:#0f172a;border-radius:20px;padding:14px;box-shadow:0 10px 24px rgba(15,23,42,.14);min-width:0}
            .cover-panel-title{font-size:13px;font-weight:700;color:#334155;margin-bottom:10px}
            .cover-chart{overflow:hidden;border-radius:14px;background:#fff}
            .cover-panels{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}
            .section{margin-top:18px}
            .question-list{display:flex;flex-direction:column;gap:16px}
            .question-page{display:flex;flex-direction:column;gap:16px}
            .question-title{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:14px}
            .question-number{font-size:14px;line-height:1.4;font-weight:600;color:#334155}
            .question-pillar{font-size:12px;color:#64748b;background:#f1f5f9;border-radius:999px;padding:6px 10px;white-space:nowrap}
            .question-card{background:#fff;border:1px solid #e2e8f0;border-radius:18px;padding:18px;box-shadow:0 8px 24px rgba(15,23,42,.08);break-inside:avoid;page-break-inside:avoid;page-break-after:avoid;overflow:hidden}
            .segments-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:16px;align-items:start;break-inside:avoid;page-break-inside:avoid}
            .segment-panel{border:1px solid #e2e8f0;border-radius:14px;padding:12px;background:#fcfdff;min-width:0;break-inside:avoid;page-break-inside:avoid;overflow:hidden}
            .segment-title{font-size:12px;font-weight:700;letter-spacing:.02em;color:#0f172a;text-transform:uppercase;margin-bottom:8px}
            .segment-body{display:flex;gap:12px;align-items:center;min-width:0;flex-wrap:nowrap}
            .segment-legend{flex:1 1 auto;min-width:0;display:flex;flex-direction:column;gap:8px}
            .legend-row{display:grid;grid-template-columns:12px minmax(0,1fr) 28px 38px;gap:8px;align-items:center;font-size:12px;color:#334155;min-width:0}
            .legend-dot{width:10px;height:10px;border-radius:999px;display:inline-block}
            .legend-label{font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
            .legend-value,.legend-pct{color:#64748b;text-align:right;min-width:0;white-space:nowrap}
            .segment-chart{flex:0 0 220px;display:flex;justify-content:center;align-items:center;min-width:0;overflow:hidden}
            .empty-state{padding:18px;background:#fff;border:1px dashed #cbd5e1;border-radius:14px;color:#64748b}
            .segment-empty{padding:14px 10px;border:1px dashed #cbd5e1;border-radius:10px;background:#fff;color:#64748b;font-size:13px}
            .recommendation-card{background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:12px;margin-bottom:8px}
            .recommendation-meta{font-size:12px;color:#64748b;margin-top:4px}
            @media (max-width: 1200px){.cover-top,.cover-panels,.segments-grid{grid-template-columns:1fr}.segment-body{flex-direction:column}.segment-chart{flex:1 1 auto}}
            @media print{
                .question-list{break-before:page;page-break-before:always}
                .question-page.question-page-break{break-after:page;page-break-after:always}
                .question-card,.segments-grid,.segment-panel{break-inside:avoid;page-break-inside:avoid}
                .segments-grid{grid-template-columns:minmax(0,1fr) minmax(0,1fr) !important}
                .segment-body{flex-direction:row !important}
                .segment-chart{flex:0 0 210px !important}
            }
        </style>
    </head>
    <body>
        <div class="header">
            <div>
                <div class="title">Relatório executivo NR01 / COPSOQ</div>
                <div class="subtitle">Visão resumida e visual para suporte à leitura do RH</div>
            </div>
        </div>

        <div class="cover">
            <div class="cover-top">
                <div>
                    <div class="cover-title">Leitura executiva da pesquisa</div>
                    <div class="cover-metrics">
                        <div class="cover-metric">
                            <div class="cover-metric-label">Respondentes</div>
                            <div class="cover-metric-value">{{ response_count }}</div>
                        </div>
                        <div class="cover-metric">
                            <div class="cover-metric-label">Adesão</div>
                            <div class="cover-metric-value">{{ ('%.1f%%' % adherence_percentage) if adherence_percentage is not none else '-' }}</div>
                        </div>
                        <div class="cover-metric">
                            <div class="cover-metric-label">Pilares críticos</div>
                            <div class="cover-metric-value">{{ critical_pillars }}</div>
                        </div>
                        <div class="cover-metric">
                            <div class="cover-metric-label">Pilares em atenção</div>
                            <div class="cover-metric-value">{{ attention_pillars }}</div>
                        </div>
                    </div>
                </div>
                <div class="cover-panel">
                    <div class="cover-panel-title">Resumo por status</div>
                    <div class="cover-chart">{{ status_div | safe if status_has_data else '<div class="empty-state">Sem dados para o resumo de status.</div>' }}</div>
                </div>
            </div>
            <div class="cover-panels">
                <div class="cover-panel">
                    <div class="cover-panel-title">Pilares (todos) por status</div>
                    <div class="cover-chart">{{ pillar_div | safe }}</div>
                </div>
                <div class="cover-panel">
                    <div class="cover-panel-title">Heatmap por segmento</div>
                    <div class="cover-chart">{{ heatmap_div | safe if heatmap_div else '<div class="empty-state">Sem dados de segmentação para heatmap.</div>' }}</div>
                </div>
            </div>
        </div>

        <div class="section">
            <h3>Todas as perguntas</h3>
            <div class="question-list">
                {{ question_cards | safe }}
            </div>
        </div>

        <div class="section">
            <h3>Recomendações</h3>
            {% for rec in recommendations %}
                <div class="recommendation-card">
                    <div><b>{{ rec.label }}</b></div>
                    <div>{{ rec.text }}</div>
                    <div class="recommendation-meta">Status: {{ rec.status }} | Média: {{ rec.media }}</div>
                </div>
            {% endfor %}
        </div>
    </body>
    </html>
    """)

    html = template.render(
        response_count=response_count,
        overall_average=overall_average,
        critical_pillars=critical_pillars,
        attention_pillars=attention_pillars,
        healthy_pillars=healthy_pillars,
        segment_column=segment_column,
        selected_segment=selected_segment,
        application_start_date=application_start_date,
        total_collaborators=total_collaborators,
        adherence_percentage=adherence_percentage,
        pillar_div=pillar_div,
        status_div=status_div,
        status_has_data=status_has_data,
        heatmap_div=heatmap_div,
        question_cards=question_cards,
        recommendations=recommendations,
    )
    return html


def build_rh_report_pdf_via_playwright_bytes(html: str, *, wait_until: str = "domcontentloaded") -> bytes:
    def _brief_error(exc: Exception, max_len: int = 260) -> str:
        text = str(exc).replace("\n", " ").strip()
        return text if len(text) <= max_len else text[: max_len - 3] + "..."

    errors: list[str] = []

    # Attempt 1: Playwright (bundled or system browser)
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = None
            launch_error = None
            try:
                browser = p.chromium.launch(headless=True)
            except Exception as e:
                launch_error = e

            if browser is None:
                candidates = ["/usr/bin/chromium-browser", "/usr/bin/chromium", "/usr/bin/google-chrome-stable", "/usr/bin/google-chrome"]
                for candidate in candidates:
                    if which(candidate):
                        try:
                            browser = p.chromium.launch(headless=True, executable_path=candidate)
                            launch_error = None
                            break
                        except Exception as e:
                            launch_error = e

            if browser is None:
                raise RuntimeError(str(launch_error) if launch_error is not None else "falha ao iniciar browser")

            page = browser.new_page()
            page.set_content(html, wait_until=wait_until)
            page.wait_for_timeout(1200)
            pdf_bytes = page.pdf(format="A4", print_background=True)
            browser.close()
            return pdf_bytes
    except Exception as exc:
        errors.append(f"Playwright: {_brief_error(exc)}")

    # Attempt 2: Direct Chromium command (helps when Playwright+snap has compatibility issues)
    try:
        chromium_cmd = which("chromium-browser") or which("chromium") or which("google-chrome") or which("google-chrome-stable")
        if chromium_cmd:
            token = uuid4().hex
            html_path = Path(f"/tmp/nr1_report_{token}.html")
            pdf_path = Path(f"/tmp/nr1_report_{token}.pdf")
            user_data_dir = Path(f"/tmp/nr1_chrome_profile_{token}")
            try:
                html_path.write_text(html, encoding="utf-8")
                user_data_dir.mkdir(parents=True, exist_ok=True)
                cmd = [
                    chromium_cmd,
                    "--headless",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--allow-file-access-from-files",
                    "--virtual-time-budget=12000",
                    f"--user-data-dir={user_data_dir}",
                    f"--print-to-pdf={pdf_path}",
                    f"file://{html_path}",
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
                if result.returncode != 0:
                    stderr = (result.stderr or "").strip()
                    stdout = (result.stdout or "").strip()
                    raise RuntimeError(stderr or stdout or f"exit code {result.returncode}")
                if not pdf_path.exists() or pdf_path.stat().st_size == 0:
                    raise RuntimeError("arquivo PDF não foi gerado")
                return pdf_path.read_bytes()
            finally:
                if html_path.exists():
                    html_path.unlink()
                if pdf_path.exists():
                    pdf_path.unlink()
                if user_data_dir.exists():
                    try:
                        for path in sorted(user_data_dir.rglob("*"), reverse=True):
                            if path.is_file():
                                path.unlink(missing_ok=True)
                            elif path.is_dir():
                                path.rmdir()
                        user_data_dir.rmdir()
                    except Exception:
                        pass
    except Exception as exc:
        errors.append(f"Chromium direto: {_brief_error(exc)}")

    # Attempt 3: WeasyPrint (if installed)
    try:
        from weasyprint import HTML

        return HTML(string=html).write_pdf()
    except Exception as exc:
        errors.append(f"WeasyPrint: {_brief_error(exc)}")

    # Attempt 4: wkhtmltopdf binary (if installed)
    try:
        wkhtmltopdf_cmd = which("wkhtmltopdf")
        if wkhtmltopdf_cmd:
            result = subprocess.run(
                [wkhtmltopdf_cmd, "-q", "-", "-"],
                input=html.encode("utf-8"),
                capture_output=True,
                timeout=120,
            )
            if result.returncode == 0 and result.stdout:
                return result.stdout
            stderr = (result.stderr or "").strip()
            raise RuntimeError(stderr or f"exit code {result.returncode}")
    except Exception as exc:
        errors.append(f"wkhtmltopdf: {_brief_error(exc)}")

    # Attempt 5: xhtml2pdf (pure Python)
    try:
        from xhtml2pdf import pisa

        output = BytesIO()
        status = pisa.CreatePDF(src=html, dest=output, encoding="utf-8")
        if not status.err:
            pdf_bytes = output.getvalue()
            if pdf_bytes:
                return pdf_bytes
        raise RuntimeError(f"falha na conversão (err={status.err})")
    except Exception as exc:
        errors.append(f"xhtml2pdf: {_brief_error(exc)}")

    details = "\n- ".join(errors) if errors else "sem detalhes"
    raise RuntimeError(
        "Não foi possível gerar PDF automaticamente. Tentativas:\n"
        f"- {details}\n"
        "Sugestão: use o download em PDF (reportlab) ou instale Google Chrome (deb/rpm, fora do Snap), ou ainda weasyprint/wkhtmltopdf."
    )


def html_pdf_backend_available() -> tuple[bool, str]:
    """Checks if there is a viable HTML->PDF backend in current environment."""
    wkhtml = which("wkhtmltopdf")
    if wkhtml:
        return True, "wkhtmltopdf"

    chromium_cmd = which("chromium-browser") or which("chromium") or which("google-chrome") or which("google-chrome-stable")
    if chromium_cmd:
        return True, f"chromium:{chromium_cmd}"

    try:
        import weasyprint  # noqa: F401

        return True, "weasyprint"
    except Exception:
        pass

    try:
        from playwright.sync_api import sync_playwright  # noqa: F401

        # Playwright itself is present. Check for an available browser candidate.
        for candidate in ["/usr/bin/google-chrome-stable", "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser"]:
            if which(candidate):
                return True, f"playwright:{candidate}"
        return False, "playwright instalado, mas sem navegador encontrado"
    except Exception:
        return False, "sem backend HTML->PDF disponível"
