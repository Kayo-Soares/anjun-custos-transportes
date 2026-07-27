# Indicador de Custos de Transporte | 运输成本指标

Dashboard interativo em **Streamlit** para acompanhamento de custos de transporte
secundário: custo total, custo por fornecedor, custo por pacote, custo por KM e
evolução mensal — construído a partir dos arquivos mensais "Controle de Envio".

## 📊 O que o dashboard mostra

- **KPIs**: Custo Total, Custo por Pacote, Custo por KM, Nº de Fretes, Total de KM
  Rodados, Total de Pacotes
- **Custo por Fornecedor**: ranking em barras horizontais
- **Custo por Pacote** e **Custo por KM**: distribuição por rota (donut)
- **Mês a mês**: evolução do custo total, custo/pacote e KM rodado, com variação
  percentual mês contra mês
- Filtros por mês, rota e fornecedor na barra lateral
- Todos os rótulos em português e chinês (PT | 中文)

## 🚀 Como rodar

### 1. Clonar o repositório
```bash
git clone https://github.com/Kayo-Soares/anjun-custos-transportes.git
cd anjun-custos-transportes
```

### 2. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 3. Colocar os dados
Coloque os arquivos mensais `.xlsx` ("Controle de Envio") na mesma pasta do
`streamlit_app.py` — **ou** use o botão de upload dentro do próprio app.

> Os arquivos `.xlsx` **não fazem parte deste repositório** (ver `.gitignore`) por
> conterem dados de custo da empresa. Guarde-os localmente ou em um repositório
> privado separado.

### 4. Rodar o app
```bash
streamlit run streamlit_app.py
```

Abre automaticamente em `http://localhost:8501`.

## 📁 Estrutura esperada da planilha de entrada

Cada arquivo mensal precisa ter uma aba chamada **`Custo Secundaria`**, com as
colunas originais bilíngues (chinês/português) exportadas do sistema de controle
de envios — o script já faz a limpeza e o mapeamento de nomes automaticamente.

## 🛠️ Stack

- [Streamlit](https://streamlit.io) — interface
- [Pandas](https://pandas.pydata.org) — consolidação e tratamento dos dados
- [Plotly](https://plotly.com/python/) — gráficos interativos
- [OpenPyXL](https://openpyxl.readthedocs.io) — leitura dos arquivos `.xlsx`

## 📌 Notas

- O app tem cache (`st.cache_data`) — se atualizar um `.xlsx` na pasta, use
  **R** (Rerun) ou o menu **⋮ → Clear cache** no Streamlit pra forçar releitura.
- "Custo por Pacote" usa a coluna `Carga real` da planilha original (nome
  historicamente confuso, mas é o campo correto de quantidade de pacotes).
