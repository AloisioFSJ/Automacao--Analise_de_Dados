import pandas as pd
import os
import glob
from IPython.display import display
from openpyxl import load_workbook
from openpyxl.styles import numbers
from pandas import ExcelWriter

#============================================================================================================================================================
# CRIAÇÃO DA PASTA PARA COLOCAR O ARQUIVO ===================================================================================================================
# AJUSTA O CAMINHO DO ARQUIVO A DEPENDER SE FOR '.py.' OU '.exe'
def resource_path():
    import sys
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

caminho_base = resource_path()

pasta_destino = "Arquivos"
caminho_pasta = os.path.join(caminho_base, pasta_destino)

# 🔍 Verifica se a pasta existe =============================================================================================================================
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)
    print(f'📂  Pasta "{pasta_destino}" criada com sucesso.')
    print(f'⚠️  Coloque o arquivo .xls na pasta "{pasta_destino}" e execute o script novamente.')
    input('Pressione "Enter" para finalizar ')
    exit()

# 🔎 Procura arquivos Excel dentro da pasta criada ==========================================================================================================
arquivos_excel = glob.glob(os.path.join(caminho_pasta, "*.xlsx")) + glob.glob(os.path.join(caminho_pasta, "*.xls"))

# PRINTA O CAMINHO ONDE A PASTA ESTA SENDO CRIADA PARA VERIFICACAO ===========================================================================================
print(f'📍 Caminho completo verificado: {caminho_pasta}')

# ❌ Se não encontrar nenhum, avisa e para o script =========================================================================================================
if not arquivos_excel:
    print(f'⚠️  Nenhum arquivo Excel encontrado na pasta "{pasta_destino}".')
    print(f'➡️  Por favor, coloque o arquivo desejado na pasta e rode novamente.')
    input('Pressione "Enter" para finalizar ')
    exit()

# ✅ Lê o primeiro arquivo encontrado =======================================================================================================================
arquivo = arquivos_excel[0]
print(f'✅ Arquivo encontrado: {os.path.basename(arquivo)}')


# 📊 Leitura do arquivo com pandas ===========================================================================================================================
try:
    vendas_df = pd.read_excel(arquivo)
    print('✅  Arquivo carregado com sucesso.')
except Exception as e:
    print(f'❌  Erro ao ler o arquivo: {e}')
    exit()

# LEITURA DO ARQUIVO EM '.xls' E PORTE DELE PARA '.xlsx' =====================================================================================================
vendas_df.to_excel(os.path.join(caminho_pasta, "Pedidos_conversao.xlsx"), index=False)
vendas_df = pd.read_excel(os.path.join(caminho_pasta, "Pedidos_conversao.xlsx"), skiprows=6)
vendas_df = vendas_df.iloc[:, 1:]

# ORIGEM ======================================================================================================================================================
# APLICA O TEXTO 'PRÓPRIO' EM TODOS OS CAMPOS VAZIOS DE 'Origem' ==============================================================================================
vendas_df['Origem'] = vendas_df['Origem'].fillna('PRÓPRIO')

# CLASSIFICACAO
# IF PARA DEFINIR A CLASSIFICACAO DA DISTANCIA POR ROTA =======================================================================================================
vendas_df['Distancia por rota (Km)'] = (vendas_df['Distancia por rota (Km)'].astype(str).str.replace(',', '.', regex=False))
vendas_df['Distancia por rota (Km)'] = pd.to_numeric(vendas_df['Distancia por rota (Km)'], errors='coerce')
col_distancia_rota = vendas_df.columns.get_loc('Distancia por rota (Km)')

def classificar(km):
    if pd.isna(km):
        return '0 a 2.5 km'
    elif km <= 2.499:
        return '0 a 2.5 km'
    elif km <= 4.99:
        return '2.5 a 5 km'
    elif km <= 8.99:
        return '5 a 9 km'
    elif km <= 30:
        return '9 a 22 km'

# CRIAR A COLUNA 'Classificação' LOGO APOS 'Distancia por rota (Km)' ===========================================================================================
vendas_df.insert(col_distancia_rota + 1, 'Classificação', vendas_df['Distancia por rota (Km)'].apply(classificar))

# ALTERAR TODOS OS ITENS QUE FICARAM COMO 'Diária' POREM QUE NAO SAO UMA DIARIA ================================================================================
vendas_df.loc[
    vendas_df['Código'] == 'DIARIA',
    'Classificação'
] = 'Diária'

# TAXA DE ENTREGA ==============================================================================================================================================
# CRIAR A COLUNA 'Taxa de Entrega' COM O VALOR DA 'Taxa Total' - 'TAXA EXTRA' ==================================================================================
vendas_df['Taxa total cobrada'] = (
    vendas_df['Taxa total cobrada'].astype(str).str.replace(',', '.', regex=False)
)
vendas_df['Taxa extra cobrada'] = (
    vendas_df['Taxa extra cobrada'].astype(str).str.replace(',', '.', regex=False)
)
vendas_df['Taxa total cobrada'] = pd.to_numeric(vendas_df['Taxa total cobrada'], errors='coerce')
vendas_df['Taxa extra cobrada'] = pd.to_numeric(vendas_df['Taxa extra cobrada'], errors='coerce')


col_entregador = vendas_df.columns.get_loc('Entregador')
vendas_df.insert(col_entregador + 1, 'Taxa de Entrega', vendas_df['Taxa total cobrada'] - vendas_df['Taxa extra cobrada'])

vendas_df['Taxa de Entrega'] = vendas_df['Taxa total cobrada'] - vendas_df['Taxa extra cobrada']

# TAXA DE ENTREGA ENTREGADOR ==================================================================================================================================
# CRIA A COLUNA 'Taxa de entrega entregador' COM O VALOR DA 'Taxa total' - 'Taxa extra' =======================================================================
vendas_df['Taxa total entregador'] = (
    vendas_df['Taxa total entregador'].astype(str).str.replace(',', '.', regex=False)
)
vendas_df['Taxa extra entregador'] = (
    vendas_df['Taxa extra entregador'].astype(str).str.replace(',', '.', regex=False)
)
vendas_df['Taxa total entregador'] = pd.to_numeric(vendas_df['Taxa total entregador'], errors='coerce')
vendas_df['Taxa extra entregador'] = pd.to_numeric(vendas_df['Taxa extra entregador'], errors='coerce')


col_taxa_extra = vendas_df.columns.get_loc('Taxa extra cobrada')
vendas_df.insert(col_taxa_extra + 1, 'Taxa de entrega entregador', vendas_df['Taxa total entregador'] - vendas_df['Taxa extra entregador'])

vendas_df['Taxa de entrega entregador'] = vendas_df['Taxa total entregador'] - vendas_df['Taxa extra entregador']

# DATA DE CADASTRO =============================================================================================================================================
# DIVIDE A TABELA 'Data de cadastro' EM DUAS, SEPARANDO A DATA DO HORARIO ======================================================================================
vendas_df['Data de cadastro'] = vendas_df['Data de cadastro'].astype(str)

# CRIAR AS COLUNAS
col_data_cadastro = vendas_df.columns.get_loc('Data de cadastro')
vendas_df.insert(col_data_cadastro + 1, 'Data', '0')

col_datas = vendas_df.columns.get_loc('Data')
vendas_df.insert(col_datas + 1, 'Dia da semana', '0')

col_dia_semana = vendas_df.columns.get_loc('Dia da semana')
vendas_df.insert(col_dia_semana + 1, 'Turno', '0')

col_turno = vendas_df.columns.get_loc('Turno')
vendas_df.insert(col_turno + 1, 'Hora do cadastro', '0')

# DATA
vendas_df['Data'] = vendas_df['Data de cadastro'].str.split().str[0]
vendas_df['Data'] = pd.to_datetime(vendas_df['Data'], dayfirst=True).dt.normalize()

# DIA DA SEMANA
dias_pt = {
    0: 'seg',
    1: 'ter',
    2: 'qua',
    3: 'qui',
    4: 'sex',
    5: 'sáb',
    6: 'dom'
}

vendas_df['Dia da semana'] = vendas_df['Data'].dt.weekday.map(dias_pt)

# HORA DO CADASTRO & TURNO
vendas_df['Hora do cadastro'] = vendas_df['Data de cadastro'].str.split().str[1]
vendas_df['Hora do cadastro'] = pd.to_datetime(vendas_df['Hora do cadastro'], format='%H:%M:%S').dt.time

def definir_turno(hora):
    if hora <= pd.to_datetime('17:29:00').time():
        return 1
    elif hora >= pd.to_datetime('17:30:00').time() and hora <= pd.to_datetime('23:00:00').time():
        return 2
    else:
        return 0 
    
vendas_df['Turno'] = vendas_df['Hora do cadastro'].apply(definir_turno)

# ALTERAR TODOS OS ITENS QUE FICARAM COMO TURNO SENDO DIÁRIA ===================================================================================================
col_data_agendamento = vendas_df.columns.get_loc('Data de agendamento')
vendas_df.insert(col_data_agendamento + 1, 'Hora de agendamento', '0')
vendas_df['Data de agendamento'] = vendas_df['Data de agendamento'].astype(str)
vendas_df['Hora de agendamento'] = vendas_df['Data de agendamento'].str.split().str[1]
vendas_df['Hora de agendamento'] = pd.to_datetime(vendas_df['Hora de agendamento'], format='%H:%M:%S').dt.time

mask = vendas_df['Código'] == 'DIARIA'
vendas_df.loc[mask, 'Turno'] = (vendas_df.loc[mask, 'Hora de agendamento'].apply(definir_turno))

# HORA DE PRONTO
# CRIAR A COLUNA 'Hora de Pronto' UTILIZANDO APENAS O HORARIO DA COLUNA 'Data pronto' ==========================================================================
col_hora_cadastro = vendas_df.columns.get_loc('Hora do cadastro')
vendas_df.insert(col_hora_cadastro + 1, 'Hora de Pronto', '0')
vendas_df['Data pronto'] = vendas_df['Data pronto'].astype(str)
vendas_df['Hora de Pronto'] = vendas_df['Data pronto'].str.split().str[1]
vendas_df['Hora de Pronto'] = pd.to_datetime(vendas_df['Hora de Pronto'], format='%H:%M:%S').dt.time

# MEDIA DE PREPARO
# CRIA A COLUNA 'Média Preparo' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA ==================================================================================
col_hora_pronto = vendas_df.columns.get_loc('Hora de Pronto')
vendas_df.insert(col_hora_pronto + 1, 'Média Preparo', '')

# HORA DE DESPACHADO
# CRIAR A COLUNA 'Hora de despachado' PEGANDO APENAS O HORARIO DA COLUNA 'Data despachado' =====================================================================
col_data_despachado = vendas_df.columns.get_loc('Data despachado')
vendas_df.insert(col_data_despachado + 1, 'Hora de despachado', '0')
vendas_df['Data despachado'] = vendas_df['Data despachado'].astype(str)
vendas_df['Hora de despachado'] = vendas_df['Data despachado'].str.split().str[1]
vendas_df['Hora de despachado'] = pd.to_datetime(vendas_df['Hora de despachado'], format='%H:%M:%S').dt.time

# MEDIA DE ATRIBUICAO
# CRIA A COLUNA 'Média de atribuição' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA =============================================================================
col_hora_despachado = vendas_df.columns.get_loc('Hora de despachado')
vendas_df.insert(col_hora_despachado + 1, 'Média de atribuição', '')

# HORA DE ROTA
# CRIA A COLUNA 'Hora de rota' APENAS COM O HORARIO DA COLUNA 'Data em rota' ====================================================================================
col_media_atribuicao = vendas_df.columns.get_loc('Média de atribuição')
vendas_df.insert(col_media_atribuicao + 1, 'Hora de rota', '0')
vendas_df['Hora de rota'] = vendas_df['Data em rota'].astype(str).str.split().str[1]
vendas_df['Hora de rota'] = pd.to_datetime(vendas_df['Hora de rota']).dt.time

# MEDIA DE DESPACHO
# CRIAR A COLUNA 'Média de despacho' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA ==============================================================================
col_hora_rota = vendas_df.columns.get_loc('Hora de rota')
vendas_df.insert(col_hora_rota + 1, 'Média de despacho', '')

# HORA DE FNALIZACAO
# CRIAR A COLUNA 'Hora de Finalização' COM APENAS O HORARIO DA COLUNA 'Data finalização' ========================================================================
col_media_despacho = vendas_df.columns.get_loc('Média de despacho')
vendas_df.insert(col_media_despacho + 1, 'Hora de Finalização', '0')

vendas_df['Hora de Finalização'] = vendas_df['Data finalização'].astype(str).str.split().str[1]
vendas_df['Hora de Finalização'] = pd.to_datetime(vendas_df['Hora de Finalização']).dt.time

# MEDIA DE ENTREGA
# CRIAR A COLUNA 'Média de entrega' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA ===============================================================================
col_hora_finalizacao = vendas_df.columns.get_loc('Hora de Finalização')
vendas_df.insert(col_hora_finalizacao + 1, 'Média de entrega', '')

# PERIODO
# CRIAR A COLUNA 'periodo' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA ===============================================================================
col_media_entrega = vendas_df.columns.get_loc('Média de entrega')
vendas_df.insert(col_media_entrega + 1, 'Período', '')

# OPERACAO
# CRIAR A COLUNA 'operacao' VAZIA PARA TER A FORMULA DA PLANILHA APLICADA ===============================================================================
col_periodo = vendas_df.columns.get_loc('Período')
vendas_df.insert(col_periodo + 1, 'Operação', '')

# APAGAR COLUNAS NAO UTILIZADAS ================================================================================================================================
vendas_df = vendas_df.drop('Situação', axis=1)
vendas_df = vendas_df.drop('CNPJ', axis=1)
vendas_df = vendas_df.drop('ID', axis=1)
vendas_df = vendas_df.drop('Detalhes', axis=1)
vendas_df = vendas_df.drop('Forma de pagamento', axis=1)
vendas_df = vendas_df.drop('Nome do cliente', axis=1)
vendas_df = vendas_df.drop('Tem retorno', axis=1)
vendas_df = vendas_df.drop('Distancia por raio (Km)', axis=1)
vendas_df = vendas_df.drop('CPF', axis=1)
vendas_df = vendas_df.drop('Tipo veiculo', axis=1)
vendas_df = vendas_df.drop('Código entregador', axis=1)
vendas_df = vendas_df.drop('Data de cadastro', axis=1)
vendas_df = vendas_df.drop('Data de agendamento', axis=1)
vendas_df = vendas_df.drop('Hora de agendamento', axis=1)
vendas_df = vendas_df.drop('Data pronto', axis=1)
vendas_df = vendas_df.drop('Data despachado', axis=1)
vendas_df = vendas_df.drop('Data de aceite', axis=1)
vendas_df = vendas_df.drop('Data chegou no estabelecimento', axis=1)
vendas_df = vendas_df.drop('Data em rota', axis=1)
vendas_df = vendas_df.drop('Data retornando', axis=1)
vendas_df = vendas_df.drop('Data chegou no destino', axis=1)
vendas_df = vendas_df.drop('Data finalização', axis=1)
vendas_df = vendas_df.drop('Data de conclusão', axis=1)
vendas_df = vendas_df.drop('Data ETA Entrega', axis=1)
vendas_df = vendas_df.drop('Tipo de fatura', axis=1)
vendas_df = vendas_df.drop('Tipo de despacho', axis=1)
vendas_df = vendas_df.drop('Nome do operador', axis=1)
vendas_df = vendas_df.drop('Endereço origem', axis=1)
vendas_df = vendas_df.drop('Endereço entrega', axis=1)


vendas_df = vendas_df.dropna(how='all', axis=1)

# MOSTRAR (NO CODIGO) A PLANILHA POS MANUTENCOES ===============================================================================================================
print('PLANILHA INTEIRA ==========================================================')
display(vendas_df)
print('COLUNAS TABELA ============================================================')
print(vendas_df.columns)
print('COLUNAS ESPECÍFICAS =======================================================')
display(vendas_df[['Média de entrega', 'Período', 'Hora do cadastro', 'Operação']])

# TRANSFORMA EM UM ARQUIVO NOVO NA PASTA DA AUTOMACAO ==========================================================================================================
arquivo_final = os.path.join(caminho_base, "Pedidos.xlsx")

# Salva com o ExcelWriter para aplicar formatações
with ExcelWriter(arquivo_final, engine='openpyxl') as writer:
    vendas_df.to_excel(writer, index=False, sheet_name='Planilha')

    # Pega a planilha ativa
    ws = writer.book['Planilha']

    # Lista de colunas que devem ser formatadas como contábil
    colunas_formatar = ['Taxa total cobrada', 'Taxa extra cobrada', 'Taxa de Entrega', 
                        'Taxa total entregador', 'Taxa extra entregador', 'Taxa de entrega entregador']

    # Aplica formatação contábil (estilo brasileiro) nessas colunas
    for col in colunas_formatar:
        if col in vendas_df.columns:
            idx = vendas_df.columns.get_loc(col) + 1  # 1-based index pro Excel
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):  # pula o cabeçalho
                for cell in row:
                    cell.number_format = 'R$ #,##0.00'

print(print('✅  Formatação e Padronização concluídas com sucesso.'))
