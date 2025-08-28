SER - Sistema de Emissão de Recibos
Funcionamento do código

---------------------------------------------------------------------------------------------------------

Linguagem utilizada: Python
Bibliotecas utilizadas:tkinter, sqlite3, reportlab, pypdf2, io, os, smtplib, pandas e outras

---------------------------------------------------------------------------------------------------------

Funções Principais:
	* gerar_recibo(): função principal para criar a janela de gerar recibo, bem como implementar todas
	  as funcionalidades para a geração.

	* gerar_relatorio(): função para criar a janela de gerar relatório, fazer a leitura do banco de dados 
	  e realizar a geração dos relatórios

	* atualizar_produtos(): função para criar a janela de atualizar responsáveis com as instruções e o
	  código para atualização da tabela de produtos

	* atualizar_responsaveis(): função para criar a janela de atualizar responsáveis com as instruções e o
	  código para atualização da tabela de responsaveis

	* menu_principal(): menu principal que agrupa todas as janelas criadas anteriormente

----------------------------------------------------------------------------------------------------------

Implementação do sistema:

	1. Criar diretórios (SER > controle | emissao | modelos | tabelas)
	2. Criar e configurar os arquivos a serem utilizados pelo sistema 
		- modelos de recibos (controle e emissao)
		- tabelas de produtos, responsaveis e emails (planilhas)
	3. Executar os códigos para inserção de dados do sistema antigo no novo banco de dados
		- arquivos localizados em (_doc > config)
	4. Após essas etapas, o sistema estará em condições de ser executado, através do ser.exe

-------------------------------------------------------------------------------------------------------------
Organização das pastas:

SER

- _doc: documentação e configurações iniciais para implementação do sistema
- controle: pasta com os recibos de duas vias
- emissao: pasta com os recibos de uma via, utilizados nas notificações por email e para envio aos interessados
- modelos: modelos vazios dos recibos utilizados para impressão dos dados
- tabelas: tabelas utilizadas no sistema
* ser.db : banco de dados do sistema utilizado para armazenamento das informações
* ser.exe: arquivo para executar o sistema

--------------------------------------------------------------------------------------------------------------

Requisitos para utilização:

- Acesso ao google drive FINANÇAS;
- Google Drive instalado localmente;
- Fonte "Arial" (existe por padrão no windows)




