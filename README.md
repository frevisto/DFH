# Repo para trabalhos na DFH
---
## Intruções de uso:
### insert_cobertura_provedor.py
- Para **insert_cobertura_provedor.py**, certifique-se de criar seu arquivo .env a partir do arquivo .env.example e instalar as dependências do projeto (os imports).
- Os campos no **.env** são respectivamente: 
  - **URL_SDWAN**, a URL do portal;
  - **URL_ADD_COBERTURA**, a URL de Adição de Cobertura do Provedor em questão;
  - **SD_USER**, seu usuário no portal;
  - **SD_PASS**, sua senha no portal.
- A planilha com os dados de entrada deve estar armazenada no diretório **data** com o nome **Pasta1.xlsx**. Sua estrutura deve ser como segue:
  - Coluna A: Cidade
  - Coluna B: Estado
  - **Não deve haver cabeçalho**
- Execute com > py ./automation/insert_cobertura_provedor.py

### gerar_folhaderosto.py
- Para **gerar_folhasderosto.py**, certifique-se de instalar as dependências listadas nos imports.
- Execute com > py ./automation/gerar_folhaderosto.py
