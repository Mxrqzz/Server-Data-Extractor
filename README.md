# Server Data Extractor

## Descrição
Este script automatiza a extração de informações de servidores a partir de uma planilha Excel, comparando com uma lista pré-definida de nomes e gerando uma nova planilha com informações de "Lotação" e "Exercício" dos servidores encontrados. Também exibe uma lista dos servidores não localizados.

## Funcionalidades
- Filtragem de servidores com base em uma lista de nomes.
- Exibição de nomes não encontrados.
- Geração de uma nova planilha (`dados_servidores.xlsx`) com as informações dos servidores encontrados.

## Dependências
- `openpyxl`

## Como usar
1. Coloque a planilha `data.xlsx` na pasta `planilhas`.
2. Adicione os nomes desejados à lista `lista_servidores`.
3. Execute o script.
4. A nova planilha será salva em `dados_servidores.xlsx`.

## Instalação
```bash
Execute `pip install openpyxl`
