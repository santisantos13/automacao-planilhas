import pandas as pd
def gerar_relatorio(arquivo):
  """
  Lê um arquivo Excel com abas por mês (formato Controle Uber),
  Limpa os dados e retorna um resumo mensal com ganhos, gastos, liquido e R$/km.  
  """

  lista_mes = []

  df = pd.read_excel(arquivo)
  abas = pd.ExcelFile(arquivo).sheet_names

  for mes in abas:
    mes_atual = pd.read_excel(arquivo, sheet_name=mes, header=1)
    mes_atual["Data"] = pd.to_datetime(mes_atual["Data"], errors="coerce")
    mes_atual = mes_atual.dropna(subset=["Data"])
    mes_atual = mes_atual[mes_atual["Data"].dt.year >= 2025]
    mes_atual = mes_atual.dropna(subset=["KM Rodados"])

    lista_mes.append(mes_atual)

  df_complete = pd.concat(lista_mes)
  df_complete["Mes"] = df_complete["Data"].dt.month


  total_ganhos = df_complete["TOTAL GANHOS"].sum()
  total_gastos = df_complete["TOTAL GASTOS"].sum()
  liquido = total_ganhos - total_gastos
  media_km = total_ganhos / df_complete['KM Rodados'].sum()
  print(200 * '-')

  print(f'GANHOS: {total_ganhos:.2f}')
  print(f'GASTOS: {total_gastos:.2f}')
  print(f'LIQUIDO: {liquido:.2f}')
  print(f'R$/KM MÉDIO: {media_km:.2f}km')
  print(200* "-")

  resumo_mensal = df_complete.groupby("Mes")[["TOTAL GANHOS", "TOTAL GASTOS", "KM Rodados"]].sum()
  resumo_mensal["Liquido"] = (resumo_mensal["TOTAL GANHOS"] - resumo_mensal["TOTAL GASTOS"]).round(2)
  resumo_mensal["R$/KM"] = (resumo_mensal["TOTAL GANHOS"] / resumo_mensal["KM Rodados"]).round(2)
  return resumo_mensal

def exportar_planilha(relatorio, nome):
  relatorio.to_excel(nome)




resultado = gerar_relatorio('Controle_Uber.xlsx')
exportar_planilha(resultado, "Relatorio_Mensal.xlsx")

print(resultado)

