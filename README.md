# Excel Base64 to PNG Converter and Normalizer

Este projeto processa uma planilha Excel, realizando as seguintes operações:

1. **Normalização de Colunas**:
   - Remove acentos e caracteres especiais de strings na Coluna A.
   - Salva o resultado na Coluna C como "Nome Normalizado".

2. **Conversão de Base64 em Imagens PNG**:
   - Lê strings Base64 na Coluna C.
   - Gera imagens PNG com base nos nomes fornecidos na Coluna L.
   - Salva as imagens em uma subpasta chamada `ft`.
   - Adiciona o caminho das imagens geradas na Coluna N.

## Pré-requisitos

- Python 3.8 ou superior
- Bibliotecas:
  - `openpyxl`
  - `unidecode`

Instale as dependências com:

```bash
pip install openpyxl unidecode
