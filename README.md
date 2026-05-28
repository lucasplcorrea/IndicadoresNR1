# IndicadoresNR1

BI em Streamlit para consolidar a pesquisa NR01/COPSOQ a partir de arquivos exportados do Forms.

## O que ele faz

Carrega a planilha de perguntas, cruza as respostas por texto da pergunta, converte as alternativas em score de 0 a 100 e exibe o resultado por pilar, por pergunta e por área administrativa/operacional.

A interface também inclui uma aba de distribuição por pergunta com filtro de área e uma aba com relatório executivo automático para o RH.

O relatório pode ser exportado em PDF com foco visual, priorizando gráficos e resumos executivos.

As recomendações do relatório são automáticas e rule-based: elas vêm dos pilares que ficam abaixo do limite crítico e de temas fixos do COPSOQ/NR01 associados a esses pilares.

## Arquivos esperados

O catálogo de perguntas fica em `perguntas/Perguntas.xlsx`.

O arquivo de respostas pode ser `.xlsx` ou `.csv`, desde que as colunas tragam o texto das perguntas do Forms.

O exemplo em `baseDeDados/Pesquisa de Ambiente de Trabalho (NR-01)(1-25).xlsx` pode ser usado como modelo de importação para validar o pipeline.

## Rodando localmente

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Se o comando de venv falhar por falta do módulo do sistema, instale o pacote `python3-venv` na sua distribuição Linux e tente novamente.

## Observações de leitura

O score já é invertido para perguntas negativas, então valores mais altos indicam uma condição mais favorável para a empresa.

O dashboard permite ajustar os limiares de alerta e de criticidade diretamente na lateral da tela.
