# Relatório Dinâmico — Acervo PJe

App Streamlit que recebe um xlsx exportado do PJe ("Relação do Acervo da
Vara — Opção Completa") e gera um relatório com:

- planilha **Dados**: todas as colunas originais + `data_minima` (menor data,
  apenas a data, entre `dt_ultimo_movimento` e `data_entrada_tarefa`) + `anos`
  + `meses`, formatada como Tabela do Excel com autofiltro;
- planilha **Tabela Dinamica**: pivot com
  - **Filtros**: `cargo_judicial`, `situacao_atual` (aplicados na UI antes da geração);
  - **Linhas**: `tarefa_atual`;
  - **Valores**: contagem de `processo`;
  - **Colunas**: `anos` → `meses` → `data_minima`.

Nome do arquivo gerado: `relatorio_dinamico_dd_mm_aa_hh_mm_ss.xlsx`.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud

1. Crie um repositório no GitHub com `app.py` e `requirements.txt`.
2. Em [share.streamlit.io](https://share.streamlit.io), conecte sua conta
   GitHub, escolha o repositório e o arquivo `app.py`.
3. Deploy.

## Arquivos

| Arquivo            | Função                                |
|--------------------|---------------------------------------|
| `app.py`           | Aplicação Streamlit                   |
| `requirements.txt` | Dependências Python                   |
| `README.md`        | Este documento                        |
