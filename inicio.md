Estou ajudando o RH da minha empresa a coletar dados sobre a NR01, a metodologia utilizada foi a COPSOQ, e separamos algumas perguntas que foram disparadas através de um forms, o objetivo e produzir um BI que aceite o arquivo xlsx com as respostas nas colunas e compile essas informações e ajude a trazer os pontos criticos da empresa que precisem de atenção.

A pesquisa foi totalmente anonima então não temos segmentação de setores, a unica segmentação da empresa foi uma pergunta inicial que questiona o colaborador se ele pertence a area administrativa ou a area operacional. 

As perguntas serão listadas no arquivo perguntas.xlsx, então recomendo já criar um venv e instalar as bibliotecas necessárias parra manipulação de arquivos xlsx e csv pois usaremos bastante eles no projeto.

Algumas considerações importantes sobre a tabela de perguntas:
Ela possui 5 colunas que são:
- Nº: Que indica o número da pergunta
- Pergunta: responsável pelo texto da pergunta no forms
- Tipo: Se é negativa ou positiva
- Pilar/Subescala: Que indica qual area da pesquisa está sendo avaliada
- Escala: Que mostra qual a escala utilizada na pesquisa

E porque isso é importante:
A coluna Nº serve para indicar a qual pilar aquela pergunta refere-se, a coluna de perguntas serve para fazer o match com a tabela que virá do forms, o tipo negativo ou positivo serve para indicar como a pergunta será avaliada, ex.: Se uma pergunta tiver o tipo "Positivo" e a Escala for "Nunca/Quase nunca - Raramente - Às vezes - Frequentemente - Sempre", caso o colaborador marque a opção "Sempre", é um sinal positivo, de que a empresa está saudável, entretanto se essa pergunta tiver o tipo "Negativo", é um sinal de alerta para aquela especificidade da empresa.

Vou deixar os exemplos de escala para utilizarmos e os respectivos valores:

Positivas:
Nunca: 0
Raramente: 25
Às vezes: 50
Frequentemente: 75
Sempre: 100

Nada/Quase nada: 0
Um pouco: 25
Moderadamente: 50
Muito: 75
Extremamente: 100

Negativas:
Nunca: 100
Raramente: 75
Às vezes: 50
Frequentemente: 25
Sempre: 0

Nada/Quase nada: 100
Um pouco: 75
Moderadamente: 50
Muito: 25
Extremamente: 0

Vou deixar também os pilares referentes a cada pergunta:
1 a 3:	Exigências Quantitativas
4:	Ritmo de Trabalho
5 a 7:	Exigências Cognitivas
8:	Exigências Emocionais
9 a 12:	Influência no Trabalho
13 a 15: Possibilidades de Desenvolvimento
16 a 17: Previsibilidade
18 a 20: Transparência no papel laboral desempenhado
21 a 23: Recompensas
24 a 26: Conflitos Laborais
27 a 29: Apoio Social de Colegas
30 a 32: Apoio Social de Superiores
33 a 35: Comunicado Social no Trabalho
36 a 39: Qualidade da Liderança
40 a 42: Confiança Horizontal
43 a 45: Confiança Vertical
46 a 48: Justiça e Respeito
49 a 50: Auto-eficácia
51 a 53: Significado do Trabalho
54 a 55: Compromisso Face ao Local de Trabalho
56 a 59: Satisfação no Trabalho
60: Insegurança no Trabalho
61: Saúde Geral
62 a 64: Conflito Trabalho/Família
65 a 66: Problemas em Dormir
67 a 68: Burnout
69 a 70: Stress
71 a 72: Sintomas Depressivos
73 a 76: Comportamentos Ofensivos