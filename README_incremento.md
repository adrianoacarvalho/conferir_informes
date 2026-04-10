# Conferência de informes + preenchimento da planilha consolidada

Este script faz 4 coisas ao mesmo tempo:

1. lê a planilha de referência na aba **GERAL**
2. extrai de cada PDF o **Nome**, **CPF** e o valor de **4. Rendimentos isentos e não tributáveis (Distribuição de lucros)**
3. renomeia/copia os PDFs para a pasta `renomeado`
4. preenche a planilha modelo do consolidado com as colunas:

- **Nome**
- **CPF**
- **Valor pago** (vem da planilha GERAL, coluna TOTAL)
- **informado** (vem do PDF)
- **diferença** (a planilha calcula)
- **observação** (preenchida quando o PDF não existe na planilha GERAL)

## Regras

- Se o PDF existir na planilha GERAL:
  - preenche `Nome`, `CPF`, `Valor pago` e `informado`
- Se o PDF **não** existir na planilha GERAL:
  - preenche `Nome`, `CPF` e `informado`
  - deixa `Valor pago` vazio
  - escreve em `observação`: `PDF não encontrado na planilha GERAL`

Além disso, o relatório `.txt` continua listando:
- divergências de valores
- PDFs sem cadastro na GERAL
- nomes da GERAL sem PDF correspondente

## Dependências

```bash
pip install openpyxl pypdf
```

## Uso

```bash
python conferir_informes_com_planilha.py "planilhaorigem.xlsx" "/caminho/pdfs" "consolidado.xlsx"
```

Com pasta de saída personalizada:

```bash
python conferir_informes_com_planilha.py "planilhaorigem.xlsx" "/caminho/pdfs" "consolidado.xlsx" --output-dir "/caminho/saida"
```

## Saídas geradas

Dentro da pasta de saída:

- `consolidado_preenchido.xlsx`
- `relatorio_conferencia.txt`
- `renomeado/`

## Exemplo de nome do PDF renomeado

```text
NOME_PESSOA_01234567890_lucrospdf.pdf
```

## Observações

- O script localiza automaticamente as colunas **Sócios** e **TOTAL** na aba `GERAL`
- A planilha modelo deve ter pelo menos os cabeçalhos:
  - `Nome`
  - `CPF`
  - `Valor pago`
  - `informado`
- Se já existir coluna `observação`, ela será usada
- Se não existir, o script cria essa coluna automaticamente
