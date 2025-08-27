import pandas as pd
import numpy as np
import os
import win32com.client as win32

# Carrega a planilha em um DataFrame do Pandas
# Substitua 'contatos.xlsx' pelo nome do seu arquivo.
def carregar_planilha():
    """
    Carrega e limpa os dados de clientes de uma planilha Excel.
    Retorna um DataFrame do Pandas com os dados limpos.
    """
    try:
        # Carrega a planilha
        caminho_arquivo = r'664 - Dados de Cliente! .xlsx' # Caminho do arquivo
        df = pd.read_excel(caminho_arquivo) # Carrega o conteúdo da planilha em um DataFrame

        # --- VERIFICAÇÃO 1: Quantas linhas foram carregadas? ---
        print(f"Arquivo carregado com sucesso. Total de linhas antes da limpeza: {len(df)}")
        
        # --- Limpeza de Dados (Vetorizada) ---
        
        # Converte a coluna de telefone para string para garantir que os métodos de texto funcionem
        # .fillna('') previne erros se houver células vazias (NaN) na coluna
        df['Telefone'] = df['Telefone'].astype(str).fillna('') # Converte para string e preenche NaN

        # Aplica todas as limpezas de uma vez de forma encadeada
        df['Telefone'] = (df['Telefone']
                          .str.lstrip(';')                  # Remove ';' do início
                          .str.replace(r'[\(\)\-\s]', '', regex=True) # Remove (), -, e espaços com uma única expressão regular
                         ) # Aplica todas as limpezas de uma vez de forma encadeada

        # Substitui valores 'placeholder' por NaN (valor nulo padrão do Pandas)
        df.loc[df['Telefone'] == '00000000000', 'Telefone'] = np.nan # Transforma '00000000000' em NaN
        df.loc[df['Telefone'] == '', 'Telefone'] = np.nan # Também transforma strings vazias em NaN

        df = df.drop(columns=['Cod_Pes', 'Pai_PF', 'Mae_PF', 'Doc_Pes','Ender_Pes', 'DtNasc_Pes','SegundoProponente', 'Num_Ven', 'cod_Prod', 'Produto',
                              'Qtde_Unid','DescrTipo1', 'ValorTipo1', 'DescrTipo2', 'ValorTipo2', 'DescrTipo3','ValorTipo3','DescrTipo4','ValorTipo4',
                              'DescrTipo5','ValorTipo5','DescrTipo6','ValorTipo6', 'SaldoDevedorTotal_Vlp','EmpObraVen', 'Data_Venda','StatusEscritura',
                              'StatusCobranca','Empresa_Ven']) # Remove colunas desnecessárias
        df = df[df['Status de Venda'] != ''] # Remove linhas com status de venda vazio

        # --- VERIFICAÇÃO 2: Como ficaram os dados após a limpeza? ---
        print("\n--- Amostra dos dados após a limpeza (5 primeiras linhas): ---")
        print(df.head()) # Mostra as 5 primeiras linhas do DataFrame

        print("\n--- Amostra dos dados após a limpeza (5 últimas linhas): ---")
        print(df.tail()) # Mostra as 5 últimas linhas do DataFrame

        print(f"\nProcessamento concluído. Total de linhas final: {len(df)}")
        df.to_excel('Dados_Limpos.xlsx', index=False) # Salva o DataFrame limpo em um novo arquivo Excel

        return df # Retorna o DataFrame limpo

    except FileNotFoundError: # Tratamento de erro para arquivo não encontrado
        print(f"ERRO: O arquivo no caminho '{caminho_arquivo}' não foi encontrado.") # Tratamento de erro para arquivo não encontrado
        return None # Retorna None em caso de erro

    except KeyError as e: # Tratamento de erro para coluna não encontrada
        print(f"ERRO: A coluna {e} não foi encontrada na planilha. Verifique os nomes das colunas (ex: 'Nome_Pes', 'Telefone', 'Email_Pes').") # Tratamento de erro para coluna não encontrada
        return None # Retorna None em caso de erro

    except Exception as e: # Tratamento de erro genérico
        print(f"Ocorreu um erro inesperado: {e}") # Tratamento de erro genérico
        return None # Retorna None em caso de erro


def email_com_imagem(email_destinatario):

    try: # Tratamento de erro para envio de e-mail
        outlook = win32.Dispatch('outlook.application') # Cria uma instância do Outlook
        mail = outlook.CreateItem(0) # Cria um novo item de e-mail

        email_caixa_compartilhada = ' coloca_o_caixa_de_email_compartilhada@hotmail.com.br' # Caixa de e-mail compartilhada

        # Adiciona a lista de email em cópia oculta
        emails_em_cco = ";".join(email_destinatario) # Cria uma string com os e-mails separados por ponto e vírgula

        mail.SentOnBehalfOfName = email_caixa_compartilhada # Define a caixa de e-mail compartilhada como remetente

        # --- CONFIGURAÇÃO DA IMAGEM ---
        imagem_caminho = os.path.relpath(r'COMUNICADO - MAX.png') # Caminho relativo da imagem

        # Use um ID simples, sem espaços ou caracteres especiais, para evitar problemas
        id_da_imagem = 'COMUNICADO_MAX' # ID da imagem

        # --- PASSO 1: Anexar a imagem ao e-mail ---
        # Isso adiciona a imagem como um anexo normal (por enquanto)
        anexo = mail.Attachments.Add(imagem_caminho) # Anexa a imagem ao e-mail

        # --- PASSO 2: Definir o Content-ID (CID) para o anexo ---
        # Esta é a parte que conecta o anexo ao HTML.
        # O número '0x3712001F' é um código padrão da Microsoft para identificar o CID.
        anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", id_da_imagem) # Define o Content-ID (CID) para o anexo

        # --- PASSO 3: Definir o corpo do e-mail em HTML, referenciando o CID ---
        mail.To = email_caixa_compartilhada # Define o destinatário do e-mail
        mail.BCC = emails_em_cco # Define os destinatários em cópia oculta
        mail.Subject = 'Informativo - MAX BURITI' # Define o assunto do e-mail
        mail.HTMLBody = f"""
            <html>
            <body>
                <br>
                <img src="cid:{id_da_imagem}" alt="Informativo - MAX" width="450">
                <br>
            </body>
            </html>
        """ # Define o corpo do e-mail em HTML

        # Exibe o e-mail
        #mail.Display() # Exibe o e-mail antes de enviar, usado para teste antes do envio para verificar o corpo do e-mail

        mail.Send() # Envia o e-mail

        print("E-mail enviado com sucesso através do Outlook!")

    except Exception as e: # Tratamento de erro para envio de e-mail
        print(f"Ocorreu um erro ao controlar o Outlook: {e}") # Tratamento de erro para envio de e-mail
        print("Verifique se o Outlook está instalado e se o acesso programático está permitido.") # Tratamento de erro para envio de e-mail


def enviar_email():

    dados_limpos = pd.read_excel(r"Dados_Limpos.xlsx") # Carrega os dados limpos do arquivo Excel

    clientes_ativos = dados_limpos[dados_limpos['Status de Venda'] !='1 - Cancelada'] # Filtra os clientes que não estão cancelados

    lista_de_email = clientes_ativos['Email_Pes'].tolist()  # Cria uma lista com os e-mails dos clientes ativos

    print(f"Total de e-mails a serem enviados: {len(lista_de_email)}") # Exibe o total de e-mails listados que serão enviados

    email_com_imagem(lista_de_email) # chama a função para enviar os e-mails

    print("E-mails enviados com sucesso!") # Exibe mensagem de sucesso


if __name__ == "__main__": # Ponto de entrada do script

    # Carrega a planilha e envia os e-mails
    carregar_planilha()
    enviar_email()


    # chamada de função para teste e para envio de forma individual em caso de falhas.
    #email_Testes = ['teste@gmail.com']
    #email_com_imagem(email_Testes)