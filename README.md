# Gerador de Relatórios

Este repositório contém um gerador automático de relatórios em Word a partir de arquivos Excel e templates.

Requisitos mínimos

- Python 3.8+
- Bibliotecas: streamlit, pandas, python-docx

Instalação rápida

1. (Opcional) Crie e ative um ambiente virtual:

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1
```

2. Instale as dependências:

```powershell
pip install -r requirements.txt
```

Como executar

Execute o aplicativo Streamlit com o comando abaixo:

```powershell
python -m streamlit run app.py --server.address=0.0.0.0
```

Observações

- Coloque o template `modelo_relatorio.docx` na mesma pasta se quiser usar o template padrão.
- Verifique as colunas esperadas nas planilhas de controle e chamados conforme a documentação no código.
