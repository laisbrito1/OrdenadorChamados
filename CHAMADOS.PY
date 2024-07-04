import pandas as pd

file_path = 'relatorio.xlsx'
df = pd.read_excel(file_path)

columns_to_keep = ['ID', 'Tipo', 'SLA', 'Assunto', 'Descrição assunto', 'Protocolo', 'Endereço']
df = df[columns_to_keep]

# Ordenar pelo campo 'Descrição assunto'
df = df.sort_values(by='Descrição assunto')

# Definir listas de palavras-chave
flextel_keywords = ['BDQ', 'MSBDQ', '[FALHA] INFRAESTRUTURA', '[TROCA', 'LICENÇA', 'SERVIDOR', 'CPE', 'MSCPE', 'DIB', 'MSDIB', 'RRR', 'MSRRR', 'BND', 'MSBND', 'SDN', 'MSSDN', 'MJC', 'MSMJC', 'SOND', 'SATÉLITE DOWM', 'TRN','ANC','MSANC', 'MSTRN','ID0', '[INSTALAÇÃO]', '[VIABILIDADE]', "ALTERAÇÃO EQUIPAMENTO",'[SUPORTE AVANÇADO]', '[FALHA] DATACENTER INFRAESTRUTURA', '[FALHA] ALARMES' ]
sonda_keywords = ['BDQ', 'MSBDQ', 'CPE', 'MSCPE', 'DIB', 'MSDIB', 'RRR', 'MSRRR', 'BND', 'MSBND', 'SDN', 'MSSDN', 'MCJ', 'MSMCJ', 'SOND', 'SATÉLITE DOWM', 'TRN', 'MSTRN','ID0', 'CGX', 'ANC', 'MSANC']
engenharia_keywords = ['[INSTALAÇÃO]', '[VIABILIDADE]', "ALTERAÇÃO EQUIPAMENTO" ]
interno_keywords = ['[SUPORTE AVANÇADO]', '[FALHA] DATACENTER INFRAESTRUTURA', '[FALHA] ALARMES' ]


df_engenharia = df[df['Assunto'].str.startswith(tuple(engenharia_keywords))]
df_internos = df[df['Assunto'].str.startswith(tuple(interno_keywords))]
df_sonda = df[df['Descrição assunto'].str.startswith(tuple(sonda_keywords))]
df_flextel = df[~df['Descrição assunto'].str.startswith(tuple(flextel_keywords)) & ~df['Assunto'].str.startswith(tuple(flextel_keywords))]


internos_file = 'chamados_internos.xlsx'
engenharia_file = 'chamados_engenharia.xlsx'
sonda_file = 'chamados_sonda.xlsx'
flextel_file = 'chamados_flextel.xlsx'


# Salvar em arquivos Excel
with pd.ExcelWriter(internos_file, engine='openpyxl') as writer:
    df_internos.to_excel(writer, index=False)

with pd.ExcelWriter(engenharia_file, engine='openpyxl') as writer:
    df_engenharia.to_excel(writer, index=False)

with pd.ExcelWriter(sonda_file, engine='openpyxl') as writer:
    df_sonda.to_excel(writer, index=False)

with pd.ExcelWriter(flextel_file, engine='openpyxl') as writer:
    df_flextel.to_excel(writer, index=False)
