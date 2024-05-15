# Sistema de Controle de Lavagem de Placas (CLP)
![image](https://github.com/HyperTechDevelopment/CLP---Controle-de-Lavagem-de-Placas/assets/155833544/d15bfb1e-272c-406a-874f-1053f3e0e517)

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
Lembre-se que, é um sistema desenvolvido para ambientes industriais, se sua empresa trabalha com placas eletrônicas e que tenha um setor dedicado para manutenção e lavagem de placas, este sistema é para você.
![image](https://github.com/HyperTechDevelopment/CLP-Controle-de-Lavagem-de-Placas/assets/155833544/214063b4-55af-405d-8154-5f24a72c4174)


# CLP - ADMIN: Módulo Administrativo do Sistema de Controle de Lavagem de Placas
![image](https://github.com/HyperTechDevelopment/CLP-Controle-de-Lavagem-de-Placas/assets/155833544/7acc1404-0925-48da-8771-17a73c6167d8)

## Descrição Geral
O CLP - ADMIN é um sistema complementar ao Sistema de Controle de Lavagem de Placas, desenvolvido para facilitar a administração de usuários e modelos. Ele permite a inserção, edição e exclusão de registros de forma eficiente e segura, garantindo que apenas coordenadores tenham acesso a essas funcionalidades.

## Funcionalidades
- **Gerenciamento de Usuários e Modelos**: Inclui funcionalidades para cadastro, edição e exclusão de usuários e modelos.
- **Exclusão de Registros**: Oferece a possibilidade de excluir registros específicos com justificativa, garantindo rastreabilidade e segurança.
- **Visualização de Registros**: Permite a visualização de listas completas de usuários e modelos cadastrados.
- **Interface Administrativa Intuitiva**: Utiliza Tkinter para fornecer uma interface gráfica clara e amigável ao usuário.

## Tecnologias Utilizadas
- **Python 3.x**: Linguagem de programação utilizada.
- **Tkinter**: Biblioteca para a criação da interface gráfica.
- **SQLite**: Sistema de gerenciamento de banco de dados.
- **Datetime**: Módulo para manipulação de datas e horas.

## Estrutura de Banco de Dados
- **Tabelas Principais**: `BDExclude` (armazena justificativas e seriais de registros excluídos), `ID_MODELO`, e `ID_USER` (armazenam informações sobre modelos e usuários).

## Fluxo de Dados
O sistema segue um fluxo linear, começando pela interface gráfica, onde diversas ações administrativas podem ser realizadas. Estas ações são processadas pela classe `DatabaseManager`, que interage diretamente com o banco de dados SQLite para atualizar, inserir ou excluir registros.

## Requisitos de Sistema
- **Sistema Operacional**: Windows 10 ou superior.
- **Bibliotecas Python necessárias**: Tkinter, SQLite, Datetime.

## Conclusão
O módulo CLP - ADMIN é uma ferramenta robusta que complementa o sistema principal de CLP, facilitando o gerenciamento eficiente de usuários e modelos e a exclusão de registros duplicados ou incorretos.
