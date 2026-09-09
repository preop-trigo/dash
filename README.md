# Dashboard Pre-Op 2026

Dashboard em Streamlit para acompanhar:

- Meta de inaugurações por marca x realizado.
- Meta total, realizado, atingimento e gap.
- Quantidade de lojas em projeto, em obra, em finalização e pausadas.
- Detalhamento das inaugurações previstas para 2026.

## Como executar no VS Code

1. Abra a pasta `dashboard_preop` no VS Code.
2. Deixe o arquivo `PROJETOS PREOP 2026 - DASH.xlsx` na mesma pasta do `app.py`.
3. Abra o terminal do VS Code e execute:

```bash
pip install -r requirements.txt
```

4. Depois execute:

```bash
streamlit run app.py
```

5. O navegador abrirá o dashboard automaticamente.

## Regras usadas

### Meta x realizado
Fonte: aba `Inaugurações 2026`.

O app soma os meses Jan-Jun e Jul-Dez das linhas `Orçado` e `Realizado` para cada marca. Assim ele não depende do cache das fórmulas do Excel.

### Status dos projetos
Fonte: aba `Geral`.

São considerados somente registros com:

- `Status = PREVISTO`
- `Tipo = INAUGURAÇÃO`
- `Data Prevista` no ano de 2026

Mapeamento:

- `PLANEJAMENTO` + `AGUARDANDO INÍCIO` -> `EM PROJETO`
- `EXECUÇÃO` -> `EM OBRA`
- `FINALIZAÇÃO` -> `EM FINALIZAÇÃO`
- `PAUSADO` -> `PAUSADAS`
