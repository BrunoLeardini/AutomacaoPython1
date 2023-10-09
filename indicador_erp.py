import pandas as pd
from openpyxl import load_workbook

try:
    # Loading do arquivo .csv em um DataFrame
    df_power_bi = pd.read_csv(r'C:\Users\bruno.ribeiro\Desktop\Task Sprint ERP\ERP Tasks.csv')

    # Divisão das colunas "Estimativa original" e "Estimativa de trabalho restante" por 3600, pois no arquivo .csv os números são dados em segundos
    df_power_bi['Estimativa original'] = df_power_bi['Estimativa original'] / 3600
    df_power_bi['Estimativa de trabalho restante'] = df_power_bi['Estimativa de trabalho restante'] / 3600
    df_power_bi['Tempo gasto'] = df_power_bi['Tempo gasto'] / 3600

    # Seleção das colunas desejadas
    colunas_selecionadas = df_power_bi[['Nome do projeto', 'Responsável', 'Chave da item', 'Categoria do status', 'Resolvido', 'Atualizado(a)', 
                                        'Estimativa original', 'Estimativa de trabalho restante', 'Tempo gasto']]

    # Loading da planilha Excel existente
    workbook = load_workbook(r'C:\Users\bruno.ribeiro\Desktop\Scripts Indicadores\Indicadores-ERP.xlsx')
    # Selecão da página onde terá os Appends
    worksheet = workbook['Tasks ERP']

    # Append para o Nome das colunas (Cabeçalho) aparecer
    worksheet.append(colunas_selecionadas.columns.tolist())
    #worksheet.append(df_power_bi.columns.tolist())

    # Append das informações referentes às colunas_selecionadas
    for index, row in colunas_selecionadas.iterrows():
        worksheet.append(row.tolist())

    """Append das informações de todas as colunas existentes no arquivo .csv
    for index, row in df_power_bi.iterrows():
        #worksheet.append(row.tolist())"""

    # Save das alterações nesse arquivo .xlsx
    workbook.save(r'C:\Users\bruno.ribeiro\Desktop\Scripts Indicadores\Indicadores-ERP.xlsx')

    # Um print apenas para conferir se o programa chegou ao final corretamente
    print("Feito!")

except Exception as e:
    print(f"Ocorreu um erro: {e}")
 