# CLP---Controle-de-Lavagem-de-Placas

# Sistema de Controle de Lavagem de Placas (CLP)

## Descrição Geral
O Sistema de Controle de Lavagem de Placas (CLP) foi desenvolvido para otimizar e automatizar o controle e monitoramento do processo de lavagem de placas em ambientes industriais. Este sistema substitui métodos manuais baseados em Excel, facilitando o processo com automações que permitem a geração e o envio de relatórios com eficiência e precisão.

## Tecnologias Utilizadas
- **Python 3.x**: Linguagem de programação principal.
- **PyQt5**: Utilizado para criar a interface gráfica do usuário.
- **SQLite**: Banco de dados para armazenamento de dados do sistema.
- **openpyxl**: Biblioteca para manipulação de arquivos Excel.
- **win32**: Biblioteca utilizada para integração com o Outlook para envio de relatórios.

## Funcionalidades
- **Interface Principal**: Inclui campos de entrada para data, turno, hora, modelo da placa, responsável, linha solicitante e fase do processo.
- **Botões de Ação**: Permite a inserção de seriais, edição de registros, consulta de dados, geração e visualização de relatórios Excel, além de funcionalidades para envio de relatórios diários por e-mail.
- **Atualização Automática de Data e Hora**: Garante que os campos de data e hora sejam atualizados em tempo real.
- **Validação de Dados**: Confirma que todos os campos são preenchidos corretamente antes da inserção dos dados no banco de dados.

## Fluxo de Dados
O sistema inicia com a inserção de dados através da interface principal, passa pela validação desses dados, e termina com o armazenamento seguro no banco de dados SQLite. Relatórios podem ser gerados a partir desses dados e enviados diretamente via e-mail.

## Estrutura de Banco de Dados
O banco de dados é composto por várias tabelas que armazenam informações sobre registros de lavagem, modelos de placas, e relatórios gerados, facilitando a consulta e manipulação de dados conforme necessário. Importante: Como o banco de dados utilizado é o SQLite, para permitir que mais de uma pessoa utilize o sistema simultaneamente com o mesmo banco de dados, é necessário configurar o caminho no banco no código-fonte do sistema. Essa configuração assegura o acesso correto e a integridade dos dados entre múltiplos usuários.

## Requisitos de Sistema
- **Sistema Operacional**: Windows 10 ou superior.
- **Python 3.x**: Deve estar instalado no sistema.
- **Bibliotecas Python**: PyQt5, SQLite, openpyxl, e win32 devem estar instaladas.

## Conclusão
O Sistema de Controle de Lavagem de Placas é uma ferramenta essencial para garantir a qualidade e eficiência no monitoramento de processos de lavagem de placas, oferecendo uma interface amigável e funcionalidades robustas para a gestão completa do processo.
