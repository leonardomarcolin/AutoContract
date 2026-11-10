# AutoContract

Aplicativo em Python com interface web (Streamlit) que gera contratos em PDF em lote. A partir de um modelo em Word (`.docx`) e de uma lista de nomes em CSV, ele cria um contrato para cada pessoa e entrega tudo em um único arquivo ZIP.

## Como funciona

1. Você anexa o modelo de contrato (`.docx`).
2. Você anexa o banco de dados (`.csv`) com uma coluna `nome`.
3. Ao clicar em **Gerar contratos**, o app substitui o primeiro parágrafo do modelo pelo nome de cada pessoa (em maiúsculas), converte o documento para PDF e salva em `contratos_gerados/`.
4. Todos os PDFs são compactados em `contratosPDF.zip`, disponível para download na própria tela.

## Requisitos

- **Windows** com **Microsoft Word instalado** (a conversão para PDF usa o próprio Word via `pywin32`)
- Python 3.9 ou superior

> O app não funciona em Linux, macOS ou em serviços de hospedagem como o Streamlit Cloud, porque depende do Word.

## Instalação

```bash
git clone https://github.com/leonardomarcolin/AutoContract.git
cd AutoContract
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Como rodar

```bash
python -m streamlit run app.py
```

## Arquivos esperados na pasta do projeto

| Arquivo | Para que serve |
|---|---|
| `modelo de contrato.docx` | Modelo de exemplo, oferecido para download na tela |
| `dados_contrato.csv` | Banco de dados de exemplo, oferecido para download na tela |

## Formato do CSV

O arquivo precisa ter uma coluna chamada `nome`, separada por vírgula:

```csv
nome
Maria Silva
João Souza
Ana Pereira
```

## Formato do modelo de contrato

O **primeiro parágrafo** do documento é substituído pelo nome da pessoa, seguido de vírgula (por exemplo, `MARIA SILVA,`). O restante do contrato é mantido como está.

## Observações

- Não abra o Word manualmente enquanto os contratos estão sendo gerados.
- Cada geração apaga os PDFs e o ZIP da execução anterior.
- Não envie ao repositório CSVs com dados pessoais reais. Use nomes fictícios nos arquivos de exemplo.

## Tecnologias

[Streamlit](https://streamlit.io), [pandas](https://pandas.pydata.org), [python-docx](https://python-docx.readthedocs.io) e [pywin32](https://github.com/mhammond/pywin32).