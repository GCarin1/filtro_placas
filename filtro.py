import pandas as pd
from datetime import datetime
import os

class Filtro:
    def __init__(self):
        self.df = None

    def ler_convercao_dado_planilha(self):
        self.df = pd.read_excel('dados.xlsx')

        # Define os títulos das colunas que você deseja usar
        colunas_desejadas = ['Placa', 'Desempenho', 'Arquitetura', 'Ano de fabricação', 'Preço atual', 'TDP']

        # Filtra as colunas desejadas
        self.df = self.df[colunas_desejadas]

        # Converte as colunas 'Desempenho' e 'Preço atual' para números (caso não estejam no formato correto)
        self.df['Desempenho'] = pd.to_numeric(self.df['Desempenho'], errors='coerce')
        self.df['Preço atual'] = pd.to_numeric(self.df['Preço atual'], errors='coerce')

        # Filtra linhas com valores válidos nas colunas
        self.df = self.df.dropna(subset=['Desempenho', 'Preço atual', 'Ano de fabricação'])

    def calculo_planilha(self):
        # Calcula a relação de desempenho/preço
        self.df['Relação Desempenho/Preço'] = self.df['Desempenho'] / self.df['Preço atual']

        # Ordena o DataFrame com base na relação de desempenho/preço, ano de fabricação e preço
        self.df = self.df.sort_values(by=['Relação Desempenho/Preço', 'Ano de fabricação', 'Preço atual'], ascending=[False, False, True])

        # Obtém as 10 melhores placas
        top_10_placas = self.df.head(10)
        self.df = top_10_placas

    def resultado(self):
        # Exibe a lista das melhores placas de vídeo
        print("Melhores placas de vídeo:")
        print(self.df)

        # Cria uma planilha com as melhores placas
        self.df.to_excel('melhores_placas.xlsx', index=False)

    def salvar_planilha(self):
        # Verifica se a pasta 'planilhas' existe e cria se não existir
        if not os.path.exists('planilhas'):
            os.makedirs('planilhas')

        # Gere o nome do arquivo com a data e hora atual
        agora = datetime.now()
        nome_arquivo = f'planilhas/melhores_placas_{agora.strftime("%Y%m%d_%H%M%S")}.xlsx'

        # Salve a planilha no arquivo
        self.df.to_excel(nome_arquivo, index=False)

        print(f"Planilha salva em {nome_arquivo}")

    def execute(self):
        self.ler_convercao_dado_planilha()
        self.calculo_planilha()
        
        self.resultado()
        self.salvar_planilha()
