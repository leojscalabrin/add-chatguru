# add-chatguru
Automatiza inclusão de chats no chatguru processando uma tabela de contatos

### .env
inclua com essas infos:

SERVER=example.chatguru.app
KEY=sua_chave_API
ACCOUNT_ID=seu_account_ID

### Dependências
pip install python-dotenv pandas openpyxl requests


### Execução
rodar o script no mesmo diretório que o arquivo excel (nome padronizado: clients.xlsx)
a tabela deve conter na coluna A('Cadastrado'): "Nao" ou "Sim" indicando se o cliente já foi cadastrado; coluna B('Nome'): Nome do contato; coluna C('ID_do_telefone'); coluna D('ID_do_diálogo'): ID do dialogo a ser executado (opcional); coluna E('ID_do_usuario'): user_id a ser cadastrado(opcional); coluna F('Número'): número do contato e coluna G('Erro'): manter em branco, será incluso descrição caso ocorra erro na API.