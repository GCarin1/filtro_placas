import pandas as pd

# Leitura da planilha em um DataFrame do pandas
df = pd.read_excel('dados.xlsx')

# Define os títulos das colunas que você deseja usar
colunas_desejadas = ['Placa', 'Desempenho', 'Arquitetura', 'Ano de fabricação', 'Preço atual', 'TDP']

# Filtra as colunas desejadas
df = df[colunas_desejadas]

# Converte as colunas 'Desempenho' e 'Preço atual' para números (caso não estejam no formato correto)
df['Desempenho'] = pd.to_numeric(df['Desempenho'], errors='coerce')
df['Preço atual'] = pd.to_numeric(df['Preço atual'], errors='coerce')

# Filtra linhas com valores válidos nas colunas
df = df.dropna(subset=['Desempenho', 'Preço atual', 'Ano de fabricação'])

# Calcula a relação de desempenho/preço
df['Relação Desempenho/Preço'] = df['Desempenho'] / df['Preço atual']

# Encontra as linhas com a maior relação desempenho/preço e ano de fabricação mais recente
melhores_placas = df.loc[df.groupby('Ano de fabricação')['Relação Desempenho/Preço'].idxmax()]

# Ordena o DataFrame com base na relação de desempenho/preço
melhores_placas = melhores_placas.sort_values(by='Relação Desempenho/Preço', ascending=False)

# Exibe a lista das melhores placas de vídeo
print("Melhores placas de vídeo:")
print(melhores_placas)

# Cria uma planilha com as melhores placas
melhores_placas.to_excel('melhores_placas.xlsx', index=False)
