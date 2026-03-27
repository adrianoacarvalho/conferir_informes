# Conferência de informes de rendimentos em PDF com planilha Excel

Este projeto automatiza a conferência de informes de rendimentos em PDF com uma planilha Excel, renomeia os arquivos encontrados e gera um relatório de inconsistências.

## O que o script faz

O script `conferir_informes.py`:

- lê a aba **GERAL** da planilha Excel;
- localiza as colunas **Sócios** e **TOTAL**;
- percorre todos os arquivos PDF de uma pasta;
- extrai do PDF:
  - o **Nome**;
  - o valor do campo **4. Rendimentos isentos e não tributáveis (Distribuição de lucros)**;
- compara o nome e o valor do PDF com a planilha;
- copia os PDFs para uma pasta de saída com o nome renomeado;
- gera um relatório `.txt` com:
  - arquivos conferidos;
  - valores divergentes;
  - nomes que existem na planilha, mas não possuem PDF correspondente;
  - PDFs cujo nome não foi encontrado na planilha;
  - erros de leitura ou extração.

## Estrutura esperada da planilha

A planilha deve conter uma aba chamada:

- `GERAL`

Nessa aba, o script procura automaticamente as colunas:

- `Sócios`
- `TOTAL`

Observações:

- o script ignora linhas vazias;
- a linha de fechamento com nome como `Total` ou `Totais` é ignorada;
- os nomes são normalizados para facilitar a comparação, removendo diferenças de maiúsculas/minúsculas e acentos.

## Estrutura esperada do PDF

O script busca no texto do PDF:

- o campo de **Nome**;
- o valor de **4. Rendimentos isentos e não tributáveis (Distribuição de lucros)**.

## Requisitos

- Python 3.10 ou superior
- bibliotecas Python:
  - `openpyxl`
  - `pypdf`

## Instalação

Crie e ative um ambiente virtual, se quiser:

```bash
python -m venv .venv
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install openpyxl pypdf
```

## Como usar

### Sintaxe

```bash
python conferir_informes.py <arquivo_excel.xlsx> <pasta_pdfs> [pasta_saida]
```

### Exemplo 1: usando a pasta padrão `renomeado`

```bash
python conferir_informes.py "relacao.xlsx" "/caminho/da/pasta/com/pdfs"
```

Nesse caso, o script criará automaticamente:

```text
/caminho/da/pasta/com/pdfs/renomeado
```

### Exemplo 2: informando uma pasta de saída personalizada

```bash
python conferir_informes.py "relacao.xlsx" "/caminho/da/pasta/com/pdfs" "/caminho/da/pasta/de/saida"
```

## Como os arquivos são renomeados

Cada PDF que tiver correspondência com um nome da planilha será copiado para a pasta de saída no formato:

```text
NOME_DO_SOCIO_nomeoriginal.pdf
```

Se já existir um arquivo com o mesmo nome, o script acrescenta um contador para evitar sobrescrita.

## Relatório gerado

O script cria um arquivo:

```text
relatorio_conferencia.txt
```

O relatório contém:

- total de PDFs lidos;
- quantidade de arquivos renomeados;
- quantidade de PDFs com status `OK`;
- quantidade de PDFs com divergência de valor;
- nomes da planilha sem PDF correspondente;
- detalhamento por arquivo.

### Status possíveis no relatório

- `OK` → nome e valor conferem;
- `VALOR_DIVERGENTE` → nome encontrado, mas valor diferente da planilha;
- `VALOR_NAO_EXTRAIDO` → nome encontrado, mas valor não foi localizado no PDF;
- `NAO_ENCONTRADO_EXCEL` → nome extraído do PDF não foi encontrado na planilha;
- `NAO_EXTRAIDO` → não foi possível extrair o nome do PDF;
- `ERRO_PDF` → erro ao abrir ou ler o PDF.

## Tolerância de valor

O script usa tolerância de `0,01` para diferenças de centavos.

Isso ajuda quando há pequenas diferenças de arredondamento entre o Excel e o PDF.

## Exemplo de fluxo de uso

1. coloque todos os PDFs em uma pasta;
2. deixe a planilha Excel pronta com a aba `GERAL`;
3. execute o script apontando para a planilha e para a pasta dos PDFs;
4. confira a pasta `renomeado` ou a pasta de saída escolhida;
5. abra o arquivo `relatorio_conferencia.txt` para ver faltantes e divergências.

## Limitações atuais

Este script depende da extração de texto do PDF.

Pode ser necessário ajustar a lógica se:

- os PDFs estiverem digitalizados como imagem, sem texto selecionável;
- o layout do informe mudar bastante;
- o campo do nome ou do valor aparecer com grafia muito diferente da esperada.

## Melhorias futuras sugeridas

Algumas melhorias úteis para a próxima versão:

- interface gráfica para selecionar:
  - planilha;
  - pasta dos PDFs;
  - pasta de saída;
- geração de relatório também em `.csv`;
- tentativa de correspondência aproximada para nomes com pequenas diferenças;
- suporte a OCR para PDFs escaneados;
- exportação de uma lista separada apenas com divergências.

## Arquivos do projeto

- `conferir_informes.py` → script principal
- `README.md` → este arquivo

## Observação importante

Antes de usar em produção, faça um teste com alguns PDFs primeiro.

Isso evita renomear em lote arquivos cujo layout possa estar diferente do modelo usado na extração.
