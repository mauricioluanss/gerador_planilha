# Gerador de fichas de implantação
O script automatiza a coleta de dados de um cliente por meio da API do TomTicket. O programa solicita o IDs de um cliente e o ID de um chamado, faz duas requisições à API utilizando um token de autenticação, e organiza as informações relevantes, como razão social, nome fantasia, CNPJ, endereço e contatos em um único arquivo. Essas informações são usadas para gerar uma ficha de implantação em uma planilha Excel, que é salva em um diretório organizado por razão social no Google Drive. O script também cria pastas conforme necessário e insere um logo na planilha. Após a execução, os arquivos temporários são removidos.

Criei esse script pois uma das minhas atribuições como Analista de Suporte Junior era realizar a criação dessa ficha de implantação manualmente. Eu precisava ficar copiando e colando os dados do tom ticket para a planilha, e depois acessar o google drive web, aí criar uma pasta com o nome da razão social do cliente e, por fim, colar a planilha lá dentro. Era um trabalho repetitivo e demorado. Agora com esse programinha, consigo ser muito mais produtivo.


## Tecnologias

- **Bibliotecas Python**:
  - `requests` - Requisições HTTP à API.
  - `openpyxl` - Geração da planilha.
  - `python-dotenv` - Captura credenciais do .env.
  - `re` - Extração de dados via expressões regulares.
  - `datetime` - Formatação de datas.
  - `os` - Manipulação dos arquivos e pastas.

- **Ferramentas**:
  - Google Drive (armazenamento das fichas).
  - TomTicket API (fonte dos dados).
 
- ** Ver `requirements.txt` para instalação das dependências.**

## Progama executando:
![GIF de exemplo](gif.gif)

## Planilha que é gerada:
<img src="planilha-gerada.png" alt="Imagem responsiva" width="50%" />
