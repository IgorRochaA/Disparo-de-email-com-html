<h1 align="center">
  🚀 Automatizador de E-mail Marketing para Clientes 🚀
</h1>

<p align="center">
  Um script poderoso para limpar, filtrar e enviar e-mails em massa com imagens embutidas usando Python e Outlook.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue.svg" alt="Python 3.8+">
  <img src="https://img.shields.io/badge/Plataforma-Windows-informational.svg" alt="Plataforma Windows">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License MIT">
</p>

---

### 📖 Tabela de Conteúdos
- [🎯 Objetivo do Projeto](#-objetivo-do-projeto)
- [✨ Funcionalidades](#-funcionalidades)
- [✅ Pré-requisitos](#-pré-requisitos)
- [⚙️ Instalação e Configuração](#-instalação-e-configuração)
- [▶️ Como Executar](#-como-executar)
- [📂 Estrutura dos Arquivos](#-estrutura-dos-arquivos)

---

### 🎯 Objetivo do Projeto

Este script automatiza o processo de comunicação com clientes, lendo uma base de dados de uma planilha Excel, realizando uma limpeza completa, filtrando os clientes elegíveis e enviando um e-mail com uma imagem embutida a partir de uma caixa de correio compartilhada do **Microsoft Outlook**.

---

### ✨ Funcionalidades

- 🧹 **Limpeza de Dados:** Processa uma planilha Excel, limpando e formatando dados como números de telefone.
- 📊 **Filtragem Inteligente:** Seleciona apenas clientes com status de venda ativo para a comunicação.
- 📄 **Exportação de Dados:** Salva a base de dados limpa em um novo arquivo Excel (`Dados_Limpos.xlsx`) para auditoria.
- 🔒 **Envio Seguro em Massa:** Dispara e-mails usando **Cópia Oculta (CCO/BCC)** para proteger a privacidade dos destinatários.
- 📤 **Automação do Outlook:** Controla o aplicativo Microsoft Outlook para enviar e-mails de forma programática.
- 🏢 **Caixa de Correio Compartilhada:** Realiza o envio "em nome de" uma caixa de correio compartilhada, centralizando a comunicação.
- 🖼️ **Imagem Embutida:** Incorpora uma imagem diretamente no corpo do e-mail HTML, garantindo um visual profissional.

---

### ✅ Pré-requisitos

<p><strong>Atenção:</strong> O ambiente abaixo é <strong>obrigatório</strong> para o funcionamento do script.</p>

- **Sistema Operacional:** ⚠️ **Windows**
- **Software:** 💼 **Microsoft Outlook** para desktop instalado, configurado com uma conta de e-mail e em execução.
- **Python:** Versão 3.8 ou superior.
- **Permissões:**
  - Acesso de leitura e escrita na pasta do projeto.
  - Permissão **"Enviar em Nome de" (Send on Behalf)** na sua conta para a caixa de correio compartilhada.

---

### ⚙️ Instalação e Configuração

Siga os passos abaixo para preparar seu ambiente:

1.  **Clone o Repositório**
    ```bash
    git clone [https://seu-repositorio-aqui.git](https://seu-repositorio-aqui.git)
    cd seu-repositorio-aqui
    ```

2.  **Crie um Ambiente Virtual** (Recomendado)
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Instale as Dependências**
    Crie um arquivo `requirements.txt` com o conteúdo abaixo:
    ```txt
    pandas
    numpy
    openpyxl
    pywin32
    ```
    E instale com o comando:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure o Script**
    Abra o arquivo `.py` e altere as seguintes variáveis:
    - `caminho_arquivo` na função `carregar_planilha()`: Deve apontar para sua planilha de dados.
    - `email_caixa_compartilhada` na função `email_com_imagem()`: Deve ser o e-mail da sua caixa compartilhada.
    - `imagem_caminho` na função `email_com_imagem()`: Deve ser o nome do seu arquivo de imagem.

---

### ▶️ Como Executar

Com tudo configurado, basta executar o script principal pelo terminal:

```bash
python nome_do_seu_script.py
```
O script exibirá o progresso da limpeza e do envio dos e-mails no terminal.

---

<details>
<summary><b>📂 Clique para ver a Estrutura dos Arquivos</b></summary>
<br>

```
/SEU-PROJETO
|
|-- nome_do_seu_script.py      # O script principal de automação
|-- Planilha_de_entrada .xlsx  # A planilha com os dados brutos (ENTRADA)
|-- Imagem.png         # A imagem a ser enviada no e-mail
|-- Dados_Limpos.xlsx            # A planilha gerada após a limpeza (SAÍDA)
|-- requirements.txt             # Lista de dependências Python
|-- README.md                    # Este arquivo
```
</details>
