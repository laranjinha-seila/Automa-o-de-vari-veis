# Automa-o-de-vari-veis
Automatizar a atualização de variáveis econômicas em planilhas excel.

## Atualização da tabela do SIDRA (t/6318)

O arquivo base está em `Suporte_IE_PnadC_Mensal.xlsx`, aba `Tabela`. O script
`scripts/update_sidra_6318.R` baixa os dados da API do SIDRA, atualiza somente a
faixa de dados da aba e gera um CSV estável para uso no Power Query.

### Requisitos

- R 4+
- Pacotes: `jsonlite`, `openxlsx`

### Como executar

```bash
Rscript scripts/update_sidra_6318.R
```

### Variáveis de ambiente (opcional)

- `SIDRA_DATA_URL`: URL da API de valores. Ajuste filtros usando os descritores
  da tabela 6318: https://apisidra.ibge.gov.br/DescritoresTabela/t/6318
- `SIDRA_WORKBOOK`: caminho do arquivo Excel (padrão: `Suporte_IE_PnadC_Mensal.xlsx`)
- `SIDRA_SHEET`: nome da aba (padrão: `Tabela`)
- `SIDRA_OUTPUT_CSV`: caminho do CSV de saída (padrão: `sidra_6318_output.csv`)

### Power Query

1. No Excel: **Dados > Obter Dados > De Texto/CSV**.
2. Selecione `sidra_6318_output.csv` (gerado pelo script).
3. Carregue a tabela e configure **Atualizar ao abrir** ou uma atualização manual.
4. Quando quiser atualizar os números, execute o script novamente e clique em **Atualizar Tudo**.
