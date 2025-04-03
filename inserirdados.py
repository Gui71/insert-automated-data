import gspread
import pandas as pd
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
# Importe a exceção para ser mais explícito
from gspread.exceptions import WorksheetNotFound, APIError
from gspread.utils import column_letter_to_index # Útil para converter letras (A, B, C) em índices (1, 2, 3)
# Atenção: gspread usa índices baseados em 1 para funções como uptade_cell, mas a API(batch_update) usa índices baseados em 0.
# Vamos usar índices baseados em 0 para a API diretamente.

def conectar_planilha():
    # Conectar a API do google sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", scope)
    cliente = gspread.authorize(credentials)

    # Abrir a planilha pelo ID
    planilha = cliente.open_by_key("ID_SHEETS")
    return planilha

def ajustar_largura_colunas(planilha, aba, larguras):
    """
    Ajusta a largura das colunas da aba especificada usando batch_update.

    Args:
        planilha: O objeto spreadsheet do gspread.
        aba: O objeto worksheet (aba) do gspread.
        larguras: Um dicionário onde as chaves são os índices das colunas (base 0)
                  e os valores são a largura desejada em pixels.
                  Ex: {1: 150, 2: 250} para ajustar as colunas B (índice 1) e C (índice 2).
    """
    sheet_id = aba.id # Obtém o ID numérico da aba
    requests = []

    for col_index, pixel_size in larguras.items():
        requests.append({
            "updateDimensionProperties":{
                "range":{
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col_index, # Índice da coluna (base 0)
                    "endIndex": col_index + 1 # A API espera um intervalo, então +1 para afetar só uma coluna
                },
                "properties":{
                    "pixelSize": pixel_size # Largura em pixels
                },
                "fields": "pixelSize" # Campo que estamos atualizando
            }
        })
    if requests:
        body = {"requests": requests}
        try:
            planilha.batch_update(body)
            print(f"Largura das colunas ajustadas para a aba '{aba.title}'.")
        except APIError as e:
            print(f"Erro de API ao ajustar largura das colunas na aba '{aba.title}': {e}")
        except Exception as e:
            print(f"Erro inesperado ao ajustar largura das colunas: {e}")

def formatar_colunas_como_texto(planilha, aba, col_indices):
    """
    Formata colunas inteiras como Texto Simples usando batchUpdate.

    Args:
        planilha: O objeto spreadsheet do gspread.
        aba: O objeto worksheet (aba) do gspread.
        col_indices: Uma lista ou tupla de índices de colunas (base 0)
                     a serem formatadas como texto. Ex: [3, 4, 5]
    """
    sheet_id = aba.id
    requests = []

    for col_index in col_indices:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1
                },
                "cell": {                   
                    "userEnteredFormat": {      # Define o formato como o usuário o veria
                        "numberFormat": {
                            "type": "TEXT"      # Tipo de formato numérico -> TEXTO
                        }
                    }
                },
                # Campos que estamos atualizando dentro de 'cell'
                "fields": "userEnteredFormat.numberFormat"
            }
        })

    if requests:
        body = {"requests": requests}
        try:
            planilha.batch_update(body)
            print(f"Colunas {col_indices} formatadas como Texto Simples na aba '{aba.title}'")
        except APIError as e:
            print(f"Erro de API ao formatar colunas [{', '.join(map(str, col_indices))}] como texto: {e}")
        except Exception as e:
            print(f"Erro inesperado ao formatar colunas como texto: {e}")

def verificar_ou_criar_aba(planilha, empresa):
    try:
        aba = planilha.worksheet(empresa) # Tenta abrir a aba da empresa
        valores_existentes = aba.get_all_values()

        # Verifica se já há um cabeçalho definido (assumindo que o título está na linha 1)
        # Ajuste o índice se a estrutura for diferente
        cabecalho_esperado = ["Empresa", "Nome", "CPF", "Matrícula", "RG", "Email"]
        if len(valores_existentes) >= 2 and valores_existentes[1][1:7] == cabecalho_esperado: # Verifica da coluna B até G
            print(f"aba '{empresa}' encontrada e cabeçalho OK.")
            # <<<Aplica a formatação de texto imediatamente>>>
            colunas_para_texto = [1, 2, 3, 4, 5, 6] # Índices para texto simples
            formatar_colunas_como_texto(planilha, aba, colunas_para_texto)
            return aba # Se o cabeçalho já existir, retorna a aba sem recria-lo
        else:
            # Se a aba existe, mas o cabeçalho não está correto, recria (ou ajusta, dependendo da necessidade)
            print(f"Aba '{empresa}' encontrada, mas com cabeçalho inesperado. Recriando/ajustando...")
            # Aqui vamos recriar a formatação e o cabeçalho como se fosse nova
            pass # Continua para a criação/formatação abaixo
        
    except WorksheetNotFound:
        # Se não existir, cria uma nova aba com o nome da empresa
        print(f"Aba '{empresa}' não encontrada. Criando nova aba...")
        try:
            aba = planilha.add_worksheet(title=empresa, rows='1000', cols='23') # Cols=23 -> A a W
        except APIError as e:
            print(f"Erro de API ao CRIAR a aba '{empresa}': {e}")
            return None # Retorna None se não conseguiu criar a aba

    # Bloco de formatação/criação (executado se a aba foi criada ou se precisa ser reformatada)
    try:
        # Formatar título na linha 1
        aba.merge_cells('B1:G1')
        aba.update('B1', [[f"SECRASO - OPOSIÇÕES ({empresa}) 2025/2026"]]) # Valor envolvida em lista de listas para não ocasionar erro na criação
        aba.format('B1:G1', {
            "backgroundColor": {"red": 0.95, "green": 0.66, "blue": 0.52}, 
            "horizontalAlignment": "CENTER", 
            "textFormat": {"bold": True, "fontSize": 14, "fontFamily": "Arial"}
        })

        # Definir e Formatar cabeçalhos na linha 2 (B2:G2)
        cabecalhos = ["Empresa", "Nome", "CPF", "Matrícula", "RG", "Email"]
        aba.update('B2:G2', [cabecalhos]) # Garante a inserção de cabeçalhos na linha 2 (para não ocasionar conflito com a linha 1 (titulo)), colunas B a G       
        aba.format('B2:G2', {
            "backgroundColor": {"red": 0.96, "green": 0.77, "blue": 0.68}, 
            "horizontalAlignment": "CENTER", 
            "textFormat": {"bold": True, "fontSize": 12, "fontFamily": "Arial"}
        })

        # <<< AJUSTE DA LARGURA DE COLUNAS >>>
        # Defina aqui as larguras desejadas em PIXELS para cada coluna
        # Colunas: A=0, B=1, C=2, D=3, E=4, F=5, G=6
        # Ajuste como preferir!
        larguras_colunas = {
            # Índice da coluna: Largura em Pixels
            0: 50, # Coluna A (vazio)
            1: 380, # Coluna B (Empresa)
            2: 380, # Coluna C (Nome)
            3: 150, # Coluna D (CPF)
            4: 100, # Coluna E (Matrícula)
            5: 150, # Coluna F (RG)
            6: 380, # Coluna G (Email)
        }
        # Chama função para aplicar as larguras
        ajustar_largura_colunas(planilha, aba, larguras_colunas)
        
        print(f"Aba '{empresa}' configurada com sucesso!")
        return aba
    
    except APIError as e:
        print(f"Erro de API ao formatar/atualizar a aba '{empresa}': {e}")
        raise # Re-levanta o erro se não puder ser tratado aqui
    except Exception as e:
        print(f"Erro inesperado durante a formatação da aba '{empresa}': {e}")
        return None

def formatar_area_dados(aba):
    # Aplicar formatação as celulas preenchidas
    try:
        valores = aba.get_all_values()
        ultima_linha = len(valores)
        if ultima_linha > 2: # Formatar apenas se houver dados além do cabeçalho
            intervalo_dados = f'B3:G{ultima_linha}'
            aba.format(intervalo_dados, {
                "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.92}, 
                "textFormat": {"fontSize": 11, "fontFamily": "Arial"}
            })
    except APIError as e:
        print(f"Erro de API ao formatar área de dados: {e}")
    except Exception as e:
        print(f"Erro inesperado ao formatar área de dados: {e}")

def inserir_dados(aba, empresa):
    # Aplica a formatação de texto antes de inserir os dados
    colunas_para_texto = [1, 2, 3, 4, 5, 6]  # Índices para CPF, Matrícula, RG
    formatar_colunas_como_texto(planilha, aba, colunas_para_texto)

    # Solicita e insere dados do usuário na aba
    print("Digite os dados separados por espaço (Nome CPF Matrícula RG Email).")
    print("Use hífen '-' no lugar de espaços")
    print("Pressione Enter para inserir. Digite 'sair' para finalizar.")

    dados_lista = []
    while True:
        entrada = input("\u2192 ").strip()
        if entrada.lower() == "sair":
            if dados_lista:
                print("Inserindo dados antes de retornar...")
                try:
                    df = pd.DataFrame(dados_lista, columns=["Empresa", "Nome", "CPF", "Matrícula", "RG", "Email"])
                    valores_atuais = aba.get_all_values()
                    ultima_linha = len(valores_atuais) + 1
                    if ultima_linha < 3:
                        ultima_linha = 3

                    print(f"Inserido {len(df)} linha(s) a partir da linha {ultima_linha}...")
                    set_with_dataframe(aba, df, row=ultima_linha, col=2, include_column_header=False)
                    print("Dados inseridos. Formatando área de dados...")
                    formatar_area_dados(aba)
                    print("Formatação aplicada.")
                except APIError as e:
                    print(f"Erro de API ao inserir dados com set_with_dataframe ou formatar: {e}")
                except Exception as e:
                    print(f"Erro inesperado ao inserir dados: {e}")
            print("Retornando para inserção de novos dados em outra empresa...")
            return  # Sai da função inserir_dados e volta ao loop principal

        # Separa por espaço, mas permite nomes compostos com hífen
        dados = entrada.split(" ")
        if len(dados) == 5:
            # Adiciona empresa e troca hífens por espaço
            dados_formatados = [empresa] + [campo.replace("-", " ") for campo in dados]
            dados_lista.append(dados_formatados)
            print(f"Dados armazenados temporariamente ({len(dados_lista)} registros). Pressione Enter para mais dados ou digite 'sair' para gravar na planilha e ir aos próximos dados.")
        else:
            print(f"Formato inválido! Certifique-se de inserir 5 campos separados por espaço (use '-' para espaços internos). Recebido: {len(dados)}.")
            print(f"Exemplo: Joao-Silva 12345678900 98765 12345678 SSP-RJ joao-silva@email.com")

        # Verifica se o usuário pressionou Enter sem digitar nada
        if not entrada:
            if dados_lista:
                print("Processando dados para inserção...")
                try:
                    df = pd.DataFrame(dados_lista, columns=["Empresa", "Nome", "CPF", "Matrícula", "RG", "Email"])
                    valores_atuais = aba.get_all_values()
                    ultima_linha = len(valores_atuais) + 1
                    if ultima_linha < 3:
                        ultima_linha = 3

                    print(f"Inserido {len(df)} linha(s) a partir da linha {ultima_linha}...")
                    set_with_dataframe(aba, df, row=ultima_linha, col=2, include_column_header=False)
                    print("Dados inseridos. Formatando área de dados...")
                    formatar_area_dados(aba)
                    print("Formatação aplicada.")
                    dados_lista = []
                except APIError as e:
                    print(f"Erro de API ao inserir dados com set_with_dataframe ou formatar: {e}")
                except Exception as e:
                    print(f"Erro inesperado ao inserir dados: {e}")
            else:
                print("Nenhum dado para inserir.")

# --- Bloco Principal ---
if __name__ == "__main__":
    try:
        planilha = conectar_planilha()
        while True:  # Adicionado loop principal
            empresa_input = input("Nome da empresa: (use '-' para espaços): ").strip()
            empresa = empresa_input.replace("-", " ")
            aba_selecionada = verificar_ou_criar_aba(planilha, empresa)
            if aba_selecionada:
                inserir_dados(aba_selecionada, empresa)
            else:
                print("Não foi possível obter ou criar a aba da empresa. Encerrando.")
            # Remover o break daqui, para que o loop continue
    except FileNotFoundError:
        print("Erro: Arquivo 'credenciais.json' não encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado no script: {e}")
