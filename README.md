# Automação de Documentação de Power BI com IA

Este projeto utiliza Python para extrair metadados de um arquivo Power BI (`.pbit`) e usa a API do Google Gemini para gerar análises inteligentes sobre cada componente do relatório (páginas, tabelas, etc.).

## Funcionalidades
- Extrai dados de páginas, visuais, tabelas, medidas e relacionamentos.
- Gera uma análise por IA para cada página e tabela.
- Cria um documento Word (`.docx`) formatado com toda a documentação.

## Como Usar
1. Configure o arquivo `config.py`.
2. Adicione sua chave da API do Google no arquivo `.env`.
3. Rode o script principal.
