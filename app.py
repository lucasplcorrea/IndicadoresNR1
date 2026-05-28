from __future__ import annotations

from pathlib import Path

import plotly.express as px
import streamlit as st

from src.nr1_bi import (
    build_analysis,
    build_answer_distribution,
    build_rh_report_pdf_bytes,
    build_rh_report_html,
    build_rh_report_pdf_via_playwright_bytes,
    build_rh_report_markdown,
    dataframe_to_csv_bytes,
    html_pdf_backend_available,
    infer_segment_column,
    load_question_catalog,
    match_question_columns,
    read_table,
)


st.set_page_config(page_title="BI NR01 | COPSOQ", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        max-width: 1300px;
    }
    [data-testid="metric-container"] {
        border: 1px solid rgba(15, 23, 42, 0.12);
        border-radius: 16px;
        padding: 12px 14px;
        background: linear-gradient(180deg, rgba(248, 250, 252, 1) 0%, rgba(241, 245, 249, 1) 100%);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("BI NR01 | COPSOQ")
st.caption("Carregue as respostas do Forms para compilar os indicadores, localizar alertas e comparar a área administrativa com a operacional.")

default_questions_path = Path(__file__).resolve().parent / "perguntas" / "Perguntas.xlsx"

with st.sidebar:
    st.header("Entrada de dados")
    uploaded_questions = st.file_uploader("Planilha de perguntas (opcional)", type=["xlsx"], help="Use este arquivo se quiser substituir o catálogo padrão.")
    uploaded_responses = st.file_uploader("Respostas do Forms", type=["xlsx", "csv"], help="Arquivo exportado com as respostas da pesquisa.")
    critical_threshold = st.slider("Limite crítico", 0, 100, 40)
    attention_threshold = st.slider("Limite de atenção", 0, 100, 60)
    st.caption("A pontuação vai de 0 a 100. Quanto menor o score, maior o alerta.")

question_source = uploaded_questions if uploaded_questions is not None else default_questions_path

try:
    question_catalog = load_question_catalog(question_source)
except Exception as error:
    st.error(f"Não foi possível carregar a planilha de perguntas: {error}")
    st.stop()

if uploaded_responses is None:
    st.info("Carregue o arquivo de respostas para gerar o BI.")
    with st.expander("Visualizar catálogo de perguntas"):
        st.dataframe(question_catalog, use_container_width=True, hide_index=True)
    st.stop()

try:
    response_frame = read_table(uploaded_responses)
except Exception as error:
    st.error(f"Não foi possível ler o arquivo de respostas: {error}")
    st.stop()

segment_guess = infer_segment_column(response_frame)
segment_options = ["Sem segmentação"] + list(response_frame.columns)
segment_default_index = segment_options.index(segment_guess) if segment_guess in segment_options else 0

with st.sidebar:
    selected_segment = st.selectbox("Coluna de segmentação", segment_options, index=segment_default_index)
    segment_column = None if selected_segment == "Sem segmentação" else selected_segment

matches = match_question_columns(question_catalog, response_frame)
analysis = build_analysis(
    question_catalog=question_catalog,
    response_frame=response_frame,
    matches=matches,
    segment_column=segment_column,
    critical_threshold=float(critical_threshold),
    attention_threshold=float(attention_threshold),
)

pillar_summary = analysis["pillar_summary"]
question_summary = analysis["question_summary"]
segment_summary = analysis["segment_summary"]
scored_frame = analysis["scored_frame"]

overall_average = float(pillar_summary["media"].mean()) if not pillar_summary.empty else 0.0
critical_pillars = int((pillar_summary["media"] < critical_threshold).sum()) if not pillar_summary.empty else 0
critical_questions = int((question_summary["media"] < critical_threshold).sum()) if not question_summary.empty else 0

metric_a, metric_b, metric_c, metric_d = st.columns(4)
metric_a.metric("Respondentes", analysis["response_count"])
metric_b.metric("Pilares críticos", critical_pillars)
metric_c.metric("Perguntas críticas", critical_questions)
metric_d.metric("Média geral", f"{overall_average:.1f}")

tabs = st.tabs(["Visão geral", "Pilares", "Perguntas", "Segmentação", "Distribuição", "Relatório RH", "Dados"])

with tabs[0]:
    st.subheader("Resumo executivo")
    if not pillar_summary.empty:
        chart = px.bar(
            pillar_summary,
            x="pilar",
            y="media",
            color="status",
            text=pillar_summary["media"].map(lambda value: f"{value:.1f}" if value == value else ""),
            color_discrete_map={"Crítico": "#b91c1c", "Atenção": "#d97706", "Saudável": "#15803d", "Sem dados": "#64748b"},
        )
        chart.update_layout(
            yaxis=dict(range=[0, 100], title="Score médio"),
            xaxis_title="Pilar",
            legend_title_text="Status",
            height=520,
            margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(chart, use_container_width=True)

    left, right = st.columns(2)
    with left:
        st.markdown("### Pilares com menor score")
        st.dataframe(
            pillar_summary[["pilar", "media", "status", "perguntas", "respondentes_validos"]].head(10),
            use_container_width=True,
            hide_index=True,
        )
    with right:
        st.markdown("### Perguntas com maior alerta")
        st.dataframe(
            question_summary[["numero", "pilar", "media", "taxa_critica", "respostas_validas"]].head(10),
            use_container_width=True,
            hide_index=True,
        )

with tabs[1]:
    st.subheader("Análise por pilar")
    st.dataframe(pillar_summary, use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar pilares em CSV",
        data=dataframe_to_csv_bytes(pillar_summary),
        file_name="resumo_pilares_nr1.csv",
        mime="text/csv",
    )

with tabs[2]:
    st.subheader("Análise por pergunta")
    st.dataframe(question_summary, use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar perguntas em CSV",
        data=dataframe_to_csv_bytes(question_summary),
        file_name="resumo_perguntas_nr1.csv",
        mime="text/csv",
    )
    with st.expander("Perguntas sem correspondência"):
        unmatched = matches[matches["matched_column"].isna()][["question_number", "question_text", "pillar"]]
        if unmatched.empty:
            st.success("Todas as perguntas foram encontradas nas colunas do arquivo enviado.")
        else:
            st.dataframe(unmatched, use_container_width=True, hide_index=True)

with tabs[3]:
    st.subheader("Corte por área")
    if segment_summary.empty or segment_column is None:
        st.info("Nenhuma coluna de segmentação foi identificada ou selecionada.")
    else:
        st.dataframe(segment_summary, use_container_width=True, hide_index=True)
        heatmap = segment_summary.set_index("segmento")
        chart = px.imshow(
            heatmap,
            color_continuous_scale="RdYlGn",
            zmin=0,
            zmax=100,
            aspect="auto",
            labels=dict(color="Score médio"),
        )
        chart.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(chart, use_container_width=True)

with tabs[4]:
    st.subheader("Distribuição das respostas por pergunta")
    if segment_column is None:
        st.info("Selecione uma coluna de segmentação para filtrar por área.")
    else:
        segment_values = ["Todas as áreas"] + sorted(
            {
                value.strip()
                for value in response_frame[segment_column].dropna().astype(str).tolist()
                if value.strip()
            }
        )
        selected_area = st.selectbox("Filtrar por área", segment_values, index=0)
        if selected_area == "Todas as áreas":
            distribution_frame = response_frame.copy()
        else:
            distribution_frame = response_frame[
                response_frame[segment_column].astype(str).str.strip() == selected_area
            ].copy()

        available_pillars = [pillar for pillar in question_catalog["pillar"].dropna().unique().tolist() if str(pillar).strip()]
        selected_pillar = st.selectbox("Filtrar por pilar", ["Todos os pilares"] + available_pillars)
        if selected_pillar == "Todos os pilares":
            question_options = question_catalog.copy()
        else:
            question_options = question_catalog[question_catalog["pillar"] == selected_pillar].copy()

        question_options = question_options[question_options["question_number"].isin(matches[matches["matched_column"].notna()]["question_number"])]
        question_label_map = {
            f"{int(row.question_number):02d} - {row.question_text}": row
            for row in question_options.itertuples(index=False)
        }

        if not question_label_map:
            st.warning("Não há perguntas disponíveis para este filtro.")
        else:
            selected_question_label = st.selectbox("Pergunta", list(question_label_map.keys()))
            selected_question = question_label_map[selected_question_label]
            selected_match = matches.loc[matches["question_number"] == int(selected_question.question_number)].iloc[0]
            response_column = selected_match["matched_column"]

            filtered_distribution = build_answer_distribution(distribution_frame, response_column)

            st.caption(
                f"Pergunta selecionada: {selected_question_label} | Registros no filtro: {len(distribution_frame)}"
            )

            if filtered_distribution.empty:
                st.info("Nenhuma resposta disponível para esta pergunta no filtro escolhido.")
            else:
                chart = px.pie(
                    filtered_distribution,
                    names="resposta",
                    values="quantidade",
                    hole=0.4,
                )
                chart.update_traces(textposition="inside", textinfo="percent+label")
                chart.update_layout(height=560, margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(chart, use_container_width=True)

                st.dataframe(
                    filtered_distribution,
                    use_container_width=True,
                    hide_index=True,
                    column_config={"percentual": st.column_config.NumberColumn("Percentual", format="%.1f%%")},
                )

with tabs[5]:
    st.subheader("Relatório executivo para RH")
    st.caption("O relatório em PDF foi desenhado para priorizar gráficos e leitura rápida, mantendo um resumo textual no final.")
    report_markdown = build_rh_report_markdown(
        response_count=analysis["response_count"],
        critical_threshold=float(critical_threshold),
        attention_threshold=float(attention_threshold),
        pillar_summary=pillar_summary,
        question_summary=question_summary,
        segment_summary=segment_summary,
        segment_column=segment_column,
        selected_segment="todas as áreas",
    )
    st.markdown(report_markdown)
    report_pdf = build_rh_report_pdf_bytes(
        response_count=analysis["response_count"],
        critical_threshold=float(critical_threshold),
        attention_threshold=float(attention_threshold),
        pillar_summary=pillar_summary,
        question_summary=question_summary,
        segment_summary=segment_summary,
        segment_column=segment_column,
        selected_segment="todas as áreas",
    )
    st.download_button(
        "Baixar relatório em PDF (reportlab)",
        data=report_pdf,
        file_name="relatorio_rh_nr1.pdf",
        mime="application/pdf",
    )

    # HTML -> PDF via Playwright (preserva visual Plotly/HTML)
    html = build_rh_report_html(
        response_count=analysis["response_count"],
        critical_threshold=float(critical_threshold),
        attention_threshold=float(attention_threshold),
        pillar_summary=pillar_summary,
        question_summary=question_summary,
        segment_summary=segment_summary,
        segment_column=segment_column,
        selected_segment="todas as áreas",
        response_frame=response_frame,
        matches=matches,
        question_catalog=question_catalog,
    )

    backend_ok, backend_name = html_pdf_backend_available()
    if backend_ok:
        if st.button("Gerar PDF estilo Streamlit (HTML→PDF)"):
            with st.spinner("Gerando PDF estilo Streamlit..."):
                try:
                    pdf_bytes = build_rh_report_pdf_via_playwright_bytes(html)
                    st.success(f"PDF estilo Streamlit gerado ({backend_name})")
                    st.download_button(
                        "Baixar relatório (HTML→PDF)",
                        data=pdf_bytes,
                        file_name="relatorio_rh_nr1_playwright.pdf",
                        mime="application/pdf",
                    )
                except Exception as exc:
                    st.warning("Não foi possível gerar o PDF estilo Streamlit neste ambiente. Aplicando fallback para PDF padrão (reportlab).")
                    st.caption(f"Detalhe técnico: {str(exc)[:260]}...")
                    st.download_button(
                        "Baixar relatório (fallback reportlab)",
                        data=report_pdf,
                        file_name="relatorio_rh_nr1_fallback.pdf",
                        mime="application/pdf",
                    )
    else:
        st.info(f"PDF estilo Streamlit indisponível neste ambiente: {backend_name}. Use o PDF padrão abaixo.")
    st.download_button(
        "Baixar relatório em Markdown",
        data=report_markdown.encode("utf-8"),
        file_name="relatorio_rh_nr1.md",
        mime="text/markdown",
    )
    st.info(
        "As recomendações são automáticas e vêm de regras fixas: o sistema marca os pilares com média abaixo do limite crítico e associa cada tema a uma ação padrão de RH."
    )

with tabs[6]:
    st.subheader("Dados processados")
    st.write(f"Coluna de segmentação utilizada: {segment_column or 'não informada'}")
    st.dataframe(scored_frame, use_container_width=True, hide_index=True)
