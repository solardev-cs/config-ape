# CONFIG-APE

O CONFIG-APE é um aplicativo desenvolvido em Python com Streamlit para facilitar a análise e o cálculo da viabilidade econômica de sistemas FV no mercado livre de energia. Ele integra uma interface intuitiva com cálculos avançados realizados no backend, utilizando planilhas Excel e dados armazenados em bancos de dados SQLite. Além disso, integra a API do setor elétrico fornecida pela Way2 Technologies para obter as tarifas de energia.

## Funcionalidades
Importação de Dados: Integração com arquivos Excel para cálculos e geração de relatórios.
Análises Personalizadas: Filtros por estado, cidade e modalidade tarifária.
Conexão com APIs: Consulta automática de tarifas de energia com base nos parâmetros fornecidos.
Banco de Dados: Gerenciamento eficiente de informações utilizando SQLite.
Interface Gráfica: Exibição de tabelas e gráficos interativos com Streamlit.

## Requisitos do Sistema
Python 3.10 ou superior
Bibliotecas (listadas em requirements.txt):
pandas
numpy
streamlit
requests
openpyxl
... (inclua todas as dependências necessárias)

## Como Instalar e Executar o Projeto
1. Clone o repositório
bash
Copiar código
git clone https://github.com/solardev-cs/CONFIG-APE.git
cd CONFIG-APE
2. Crie um ambiente virtual e instale as dependências
bash
Copiar código
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows
pip install -r requirements.txt
3. Execute o aplicativo
bash
Copiar código
streamlit run config_ui.py
Estrutura do Projeto
graphql
Copiar código

CONFIG-APE/
│
├── .venv/                  # Ambiente virtual (ignorar no Git)
├── data/                   # Arquivos de dados e banco de dados
│   ├── backend_ape.xlsx    # Planilha Excel usada para cálculos
│   ├── dados_irrad.db      # Banco de dados SQLite com informações de irradiação
│
├── images/                 # Imagens usadas no projeto
│   ├── icone.png
│   ├── logo.png
│
├── tarifas.py              # Módulo com funções para consulta de tarifas
├── config_ui.py            # Arquivo principal (interface Streamlit)
├── requirements.txt        # Arquivo de dependências
└── README.md               # Documentação do projeto

## Como Usar
Selecione o Estado e a Cidade:
Use os seletores na interface para filtrar os dados de acordo com a localização desejada.

Configure os Parâmetros:
Escolha o ano, modalidade tarifária, e subgrupo para gerar as tarifas personalizadas.

Visualize os Resultados:
O aplicativo exibe gráficos, tabelas interativas e métricas que facilitam a análise.

Exporte Relatórios:
Baixe os resultados em formatos amigáveis (se essa funcionalidade for implementada).

## Contribuindo
Contribuições são bem-vindas!

Faça um fork do projeto.
Crie uma branch para suas alterações: git checkout -b minha-feature.
Envie um pull request com uma descrição detalhada das mudanças.

## Contato
Se tiver dúvidas ou sugestões, entre em contato:

GitHub: solardev-cs
E-mail: kikisauer@gmail.com
