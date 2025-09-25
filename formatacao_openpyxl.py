from openpyxl import load_workbook

# Função para converter valores para float
def parse_value(val):
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        val = val.replace('R$', '').replace(',', '.').strip()
        try:
            return float(val)
        except ValueError:
            print(f"Valor não numérico ignorado: {val}")
            return 0
    return 0

#Para carregar a planilha
wb = load_workbook('gastos.xlsx')
ws = wb.active

#Para ler os dados da planilha
data = []
for row in ws.iter_rows(min_row=2, max_row=21, values_only=True):
    if len(row) >= 5 and row[1] is not None and row[2] is not None:
        data.append([
            row[0] if len(row) > 0 else '',  # Gasto
            row[1],                          # Descrição
            row[2],                          # Valor
            row[3] if len(row) > 3 else '',  # Tipo de gasto
            row[4] if len(row) > 4 else ''   # Total gasto
        ])
    else:
        print(" ")

#Para eliminar as duplicatas:
conferido = set()
dados = []
for row in data:
    row_tuple = tuple(row)
    if row_tuple not in conferido:
        conferido.add(row_tuple)
        dados.append(row)

#Para somar todos os valores da coluna C
total_gastos = sum(parse_value(row[2]) for row in dados)

#Para escrever o total ao final da coluna C:
dados.append(['', 'Total', total_gastos, '', ''])

#Para somar valores por tipo de descrição
total_alimentacao = sum(parse_value(row[2]) for row in dados if row[1] == 'Alimentação')
total_lazer = sum(parse_value(row[2]) for row in dados if row[1] == 'Lazer')
total_transporte = sum(parse_value(row[2]) for row in dados if row[1] == 'Transporte')

#Para limpar a planilha existente (a partir da linha 2, para não perder os títulos das colunas):
ws.delete_rows(2, ws.max_row)

Para reescrever os dados, já manipulados:
for i, row in enumerate(dados, start=2):
    ws[f'A{i}'] = row[0]
    ws[f'B{i}'] = row[1]
    ws[f'C{i}'] = row[2]
    ws[f'D{i}'] = row[3]
    ws[f'E{i}'] = row[4]

#Para escrever os totais por tipo de descrição na coluna E
if len(dados) >= 4:
    ws['E2'] = total_alimentacao
    ws['E3'] = total_lazer
    ws['E4'] = total_transporte
else:
    print("Aviso: Não há linhas suficientes para escrever os totais nas linhas 3, 4 e 5.")

#Para salvar a planilha atualizada:
wb.save('gastos_atualizados_2.xlsx')
