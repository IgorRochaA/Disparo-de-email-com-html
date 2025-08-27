🚀 Automatizador de E-mail Marketing para Clientes
📖 Tabela de Conteúdos
🎯 Objetivo do Projeto

✨ Funcionalidades

🛠️ Tecnologias Utilizadas

✅ Pré-requisitos

⚙️ Instalação e Configuração

▶️ Como Executar

🎯 Objetivo do Projeto
Este script foi desenvolvido para automatizar o processo de comunicação com clientes. Ele lê uma base de dados de uma planilha Excel, realiza uma limpeza e formatação completa dos dados, filtra os clientes elegíveis e envia um e-mail marketing com uma imagem embutida, utilizando uma caixa de correio compartilhada através do Microsoft Outlook.

O objetivo é otimizar o tempo e garantir uma comunicação padronizada e profissional, respeitando a privacidade dos clientes.

✨ Funcionalidades
Limpeza de Dados: Processa uma planilha Excel (.xlsx), limpando e formatando dados essenciais, como números de telefone.

Filtragem Inteligente: Seleciona apenas os clientes ativos para a comunicação, com base no status da venda (ex: ignora vendas canceladas).

Exportação de Dados: Salva a base de dados já limpa em um novo arquivo Excel (Dados_Limpos.xlsx) para auditoria e uso futuro.

Envio de E-mail em Massa Seguro: Envia e-mails para múltiplos destinatários de forma segura, utilizando o campo Cópia Oculta (CCO/BCC), garantindo que nenhum cliente veja o e-mail do outro.

Automação do Microsoft Outlook: Integra-se com o cliente de e-mail Outlook para desktop para realizar os envios.

Envio por Caixa de Correio Compartilhada: Permite que os e-mails sejam enviados "em nome de" uma caixa de correio compartilhada (ex: vendas@suaempresa.com), mantendo a comunicação centralizada.

Imagem Embutida: Incorpora uma imagem de marketing diretamente no corpo do e-mail (formato HTML), em vez de enviá-la como um anexo tradicional.

🛠️ Tecnologias Utilizadas
Python 3: Linguagem de programação principal.

Pandas: Para manipulação e limpeza de dados da planilha.

NumPy: Para operações numéricas e suporte ao Pandas.

PyWin32: Para a automação e controle do aplicativo Microsoft Outlook no Windows.

✅ Pré-requisitos
Para que o script funcione corretamente, o ambiente precisa atender aos seguintes requisitos:

Sistema Operacional: Windows.

Software: Microsoft Outlook para desktop instalado, configurado e em execução.

Python: Versão 3.8 ou superior.

Permissões:

Acesso de leitura e escrita na pasta onde o script e a planilha estão localizados.

Permissão "Enviar em Nome de" (Send on Behalf) configurada na sua conta de e-mail para a caixa de correio compartilhada que será utilizada.

Acesso programático ao Outlook habilitado (geralmente padrão, mas pode ser restrito em ambientes corporativos).

⚙️ Instalação e Configuração
Siga estes passos para preparar o ambiente:

1. Clone o Repositório:

Bash

git clone https://seu-repositorio-aqui.git
cd seu-repositorio-aqui 2. Crie um Ambiente Virtual (Recomendado):

Bash

python -m venv venv
venv\Scripts\activate 3. Instale as Dependências:
Crie um arquivo requirements.txt com o seguinte conteúdo:

Plaintext

pandas
numpy
openpyxl
pywin32
Em seguida, instale as bibliotecas:

Bash

pip install -r requirements.txt 4. Configure os Arquivos:

Coloque a planilha de dados dos clientes na mesma pasta do script e renomeie-a para 664 - Dados de Cliente! .xlsx.

Coloque a imagem que será enviada na mesma pasta do script com o nome COMUNICADO - MAX.png.

5. Ajuste as Variáveis no Código:
   Abra o arquivo do script (.py) e altere as seguintes variáveis conforme necessário:

Na função carregar_planilha():

caminho_arquivo: Verifique se o nome do arquivo corresponde ao seu.

Na função email_com_imagem():

email_caixa_compartilhada: Substitua pelo endereço de e-mail da sua caixa de correio compartilhada.

imagem_caminho: Verifique se o nome do arquivo de imagem corresponde ao seu.

▶️ Como Executar
Com o ambiente virtual ativado e as configurações ajustadas, execute o script principal pelo terminal:

Bash

python nome_do_seu_script.py
O script irá primeiro processar a planilha, salvar o arquivo Dados_Limpos.xlsx, e em seguida, iniciará o processo de envio de e-mails através do Outlook. Acompanhe o progresso pelo terminal.
