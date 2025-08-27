<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Automatizador de E-mail Marketing para Clientes</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif, "Apple Color Emoji", "Segoe UI Emoji";
            line-height: 1.6;
            color: #333;
            background-color: #f4f6f8;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
            background-color: #ffffff;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            padding: 30px 40px;
            border: 1px solid #e1e4e8;
        }
        header {
            text-align: center;
            border-bottom: 2px solid #e1e4e8;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        header h1 {
            font-size: 2.2em;
            color: #24292e;
            margin-bottom: 10px;
        }
        .badges img {
            margin: 0 5px;
        }
        nav {
            background-color: #f9f9f9;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 15px 20px;
            margin-bottom: 30px;
        }
        nav h2 {
            margin-top: 0;
            font-size: 1.2em;
            color: #586069;
        }
        nav ul {
            list-style-type: none;
            padding: 0;
            margin: 0;
        }
        nav li {
            margin-bottom: 8px;
        }
        nav a {
            text-decoration: none;
            color: #0366d6;
            font-weight: 500;
        }
        nav a:hover {
            text-decoration: underline;
        }
        section {
            margin-bottom: 40px;
        }
        h2 {
            font-size: 1.8em;
            border-bottom: 1px solid #eaecef;
            padding-bottom: 10px;
            margin-top: 0;
            color: #24292e;
        }
        ul {
            list-style-type: disc;
            padding-left: 20px;
        }
        li {
            margin-bottom: 10px;
        }
        code {
            font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, Courier, monospace;
            background-color: #f6f8fa;
            padding: .2em .4em;
            margin: 0;
            font-size: 85%;
            border-radius: 3px;
        }
        pre {
            background-color: #f6f8fa;
            border-radius: 5px;
            padding: 16px;
            overflow: auto;
            border: 1px solid #e1e4e8;
        }
        pre code {
            padding: 0;
            margin: 0;
            font-size: 100%;
            background: none;
        }
        .emoji {
            margin-right: 10px;
        }
        strong {
            color: #d73a49;
        }
        footer {
            text-align: center;
            margin-top: 40px;
            font-size: 0.9em;
            color: #6a737d;
            border-top: 1px solid #eaecef;
            padding-top: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1><span class="emoji">🚀</span> Automatizador de E-mail Marketing para Clientes</h1>
            <div class="badges">
                <img src="https://img.shields.io/badge/Python-3.8%2B-blue.svg" alt="Python 3.8+">
                <img src="https://img.shields.io/badge/Pandas-2.0-blue.svg" alt="Pandas">
                <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License MIT">
            </div>
        </header>

        <nav>
            <h2><span class="emoji">📖</span> Tabela de Conteúdos</h2>
            <ul>
                <li><a href="#objetivo">Objetivo do Projeto</a></li>
                <li><a href="#funcionalidades">Funcionalidades</a></li>
                <li><a href="#tecnologias">Tecnologias Utilizadas</a></li>
                <li><a href="#pre-requisitos">Pré-requisitos</a></li>
                <li><a href="#instalacao">Instalação e Configuração</a></li>
                <li><a href="#executar">Como Executar</a></li>
                <li><a href="#estrutura">Estrutura dos Arquivos</a></li>
                <li><a href="#licenca">Licença</a></li>
            </ul>
        </nav>

        <section id="objetivo">
            <h2><span class="emoji">🎯</span> Objetivo do Projeto</h2>
            <p>Este script foi desenvolvido para automatizar o processo de comunicação com clientes. Ele lê uma base de dados de uma planilha Excel, realiza uma limpeza e formatação completa dos dados, filtra os clientes elegíveis e envia um e-mail marketing com uma imagem embutida, utilizando uma caixa de correio compartilhada através do Microsoft Outlook.</p>
            <p>O objetivo é otimizar o tempo e garantir uma comunicação padronizada e profissional, respeitando a privacidade dos clientes.</p>
        </section>

        <section id="funcionalidades">
            <h2><span class="emoji">✨</span> Funcionalidades</h2>
            <ul>
                <li><strong>Limpeza de Dados:</strong> Processa uma planilha Excel (<code>.xlsx</code>), limpando e formatando dados essenciais, como números de telefone.</li>
                <li><strong>Filtragem Inteligente:</strong> Seleciona apenas os clientes ativos para a comunicação, com base no status da venda.</li>
                <li><strong>Exportação de Dados:</strong> Salva a base de dados já limpa em um novo arquivo Excel (<code>Dados_Limpos.xlsx</code>) para auditoria e uso futuro.</li>
                <li><strong>Envio de E-mail em Massa Seguro:</strong> Envia e-mails para múltiplos destinatários de forma segura, utilizando o campo <strong>Cópia Oculta (CCO/BCC)</strong>.</li>
                <li><strong>Automação do Microsoft Outlook:</strong> Integra-se com o cliente de e-mail Outlook para desktop para realizar os envios.</li>
                <li><strong>Envio por Caixa de Correio Compartilhada:</strong> Permite que os e-mails sejam enviados "em nome de" uma caixa de correio compartilhada.</li>
                <li><strong>Imagem Embutida:</strong> Incorpora uma imagem de marketing diretamente no corpo do e-mail em formato HTML.</li>
            </ul>
        </section>

        <section id="tecnologias">
            <h2><span class="emoji">🛠️</span> Tecnologias Utilizadas</h2>
            <ul>
                <li>Python 3</li>
                <li>Pandas</li>
                <li>NumPy</li>
                <li>PyWin32</li>
            </ul>
        </section>

        <section id="pre-requisitos">
            <h2><span class="emoji">✅</span> Pré-requisitos</h2>
            <p>Para que o script funcione corretamente, o ambiente precisa atender aos seguintes requisitos:</p>
            <ul>
                <li>Sistema Operacional: <strong>Windows</strong>.</li>
                <li>Software: <strong>Microsoft Outlook</strong> para desktop instalado, configurado e em execução.</li>
                <li>Python: Versão 3.8 ou superior.</li>
                <li><strong>Permissões:</strong>
                    <ul>
                        <li>Acesso de leitura e escrita na pasta onde o script está localizado.</li>
                        <li>Permissão "Enviar em Nome de" (Send on Behalf) para a caixa de correio compartilhada.</li>
                        <li>Acesso programático ao Outlook habilitado.</li>
                    </ul>
                </li>
            </ul>
        </section>

        <section id="instalacao">
            <h2><span class="emoji">⚙️</span> Instalação e Configuração</h2>
            <ol>
                <li><strong>Clone o Repositório:</strong><br>
                    <pre><code>git clone https://seu-repositorio-aqui.git
cd seu-repositorio-aqui</code></pre>
                </li>
                <li><strong>Crie um Ambiente Virtual (Recomendado):</strong><br>
                    <pre><code>python -m venv venv
venv\Scripts\activate</code></pre>
                </li>
                <li><strong>Instale as Dependências:</strong><br>
                    Crie um arquivo <code>requirements.txt</code> com o conteúdo abaixo e depois execute o comando <code>pip</code>.
                    <pre><code># requirements.txt
pandas
numpy
openpyxl
pywin32</code></pre>
                    <pre><code>pip install -r requirements.txt</code></pre>
                </li>
                <li><strong>Configure os Arquivos:</strong>
                    <ul>
                        <li>Coloque a planilha de dados na mesma pasta do script com o nome <code>664 - Dados de Cliente! .xlsx</code>.</li>
                        <li>Coloque a imagem a ser enviada na mesma pasta com o nome <code>COMUNICADO - MAX.png</code>.</li>
                    </ul>
                </li>
                <li><strong>Ajuste as Variáveis no Código:</strong>
                    <ul>
                        <li>Na função <code>carregar_planilha()</code>, verifique a variável <code>caminho_arquivo</code>.</li>
                        <li>Na função <code>email_com_imagem()</code>, substitua o valor da variável <code>email_caixa_compartilhada</code>.</li>
                    </ul>
                </li>
            </ol>
        </section>

        <section id="executar">
            <h2><span class="emoji">▶️</span> Como Executar</h2>
            <p>Com o ambiente virtual ativado e as configurações ajustadas, execute o script principal pelo terminal:</p>
            <pre><code>python nome_do_seu_script.py</code></pre>
            <p>O script irá processar a planilha, salvar o arquivo <code>Dados_Limpos.xlsx</code>, e em seguida, iniciará o envio de e-mails através do Outlook.</p>
        </section>

        <section id="estrutura">
            <h2><span class="emoji">📂</span> Estrutura dos Arquivos</h2>
            <pre><code>/SEU-PROJETO
|
|-- nome_do_seu_script.py
|-- Planilha_de_entrada .xlsx
|-- Imagem_a_ser_enviada.png
|-- Dados_Limpos.xlsx  (gerado pelo script)
|-- requirements.txt
|-- README.md
</code></pre>
        </section>

        

        <footer>
            <p>README gerado com assistência de IA. Adaptado para o projeto de Automação de E-mails.</p>
        </footer>
    </div>
</body>
</html>
