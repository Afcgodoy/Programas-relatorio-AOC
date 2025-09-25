import pandas as pd

#Para carregar a planilha:
df = pd.read_excel('gastos.xlsx', header=None, names=['Gasto', 'Descrição', 'Valor', 'Tipo de gasto', 'Total gasto'])

#Para eliminar as duplicatas:
df = df.drop_duplicates()

#Para converter os valores da coluna D para numeros:
df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')

#Para somar todos os valores da coluna D:
total_gastos = df['Valor'].sum()

#Para escrever o total no final da coluna:
df.loc[len(df)] = ['', 'Total', total_gastos, '', '']

#Para somar todos os valores gastos com cada tipo de descrição:
total_alimentacao = df[df['Descrição'] == 'Alimentação']['Valor'].sum()
total_lazer = df[df['Descrição'] == 'Lazer']['Valor'].sum()
total_transporte = df[df['Descrição'] == 'Transporte']['Valor'].sum()

#Para inserir o totais gastos com cada tipo de descrição:
df.loc[1, 'Total gasto'] = total_alimentacao
df.loc[2, 'Total gasto'] = total_lazer
df.loc[3, 'Total gasto'] = total_transporte

#Para salvar a planilha  atualizada:
df.to_excel('gastos_atualizado.xlsx', index=False)