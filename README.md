## CONFIGURADOR APE

O Configurador APE é um aplicativo desenvolvido em Python com Streamlit para facilitar a análise e o cálculo da viabilidade econômica de sistemas FV no mercado livre de energia (ACL). Ele integra uma interface intuitiva com cálculos avançados realizados no backend, utilizando planilhas Excel e dados armazenados em bancos de dados SQLite. Além disso, integra a API do setor elétrico fornecida pela Way2 Technologies para obter as tarifas de energia.

### Funcionalidades
- Importação de Dados: Integração com arquivo Excel para cálculo econômico.
- Análises Personalizadas: Filtros por estado, cidade e modalidade tarifária.
- Conexão com API: Consulta automática de tarifas de energia com base nos parâmetros fornecidos.
- Banco de Dados: Consulta de dados de irradiação em SQLite.
- Interface Gráfica: Exibição de tabelas e gráficos com Streamlit.

### Requisitos do Sistema
- Python 3.12.3 ou superior
- Bibliotecas:
  - streamlit
  - pandas
  - openpyxl
  - matplotlib
  - millify
  - keyboard
  - xlwings
  - sqlite3
  - requests

### Como Instalar e Executar o Projeto
1. Clone o repositório
```
  git clone https://github.com/solardev-cs/config-ape.git
  cd config-ape
```
2. Crie um ambiente virtual e instale as dependências
```
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows
pip install -r requirements.txt
```
3. Execute o aplicativo
```
streamlit run app.py
```

Estrutura do Projeto
```
config-ape/
│
├── .venv/                  # Ambiente virtual
├── data/                   # Arquivos de dados
│   ├── backend_ape.xlsx    # Planilha Excel usada para cálculos
│   ├── dados_irrad.db      # Banco de dados SQLite com informações de irradiação
│
├── images/                 # Imagens
│   ├── icone.png
│   ├── logo.png
│
├── app.py                  # Arquivo principal (interface Streamlit)
├── excel.py                # Módulo com funções para acesso de leitura e escrita no Excel
├── irradiacao.py              # Módulo com funções para acesso ao banco de dados de irradiação
├── tarifas.py              # Módulo com funções para consulta de tarifas via API
├── requirements.txt        # Arquivo de dependências
└── README.md               # Documentação do projeto
```

### Como Usar
Dados Iniciais:
- Informe dados de localização, cliente e ano de análise.

Unidade Consumidora:
- Informe os dados da distribuidora, UC e consumos mensais.

Usina FV:
- Informe a potência de inversor desejada para o projeto.
  
Contrato de Energia:
- Informe as características do contrato ACL vigente.

Viabilidade:
- Informe os dados para simulação econômica.

Resumo:
- Em implementação.


### Contribuindo
Contribuições são bem-vindas!

Faça um fork do projeto.
Crie uma branch para suas alterações: git checkout -b minha-feature.
Envie um pull request com uma descrição detalhada das mudanças.

### Contato
Se tiver dúvidas ou sugestões, entre em contato:

GitHub: solardev-cs

E-mail: kikisauer@gmail.com
