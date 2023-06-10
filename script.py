import pandas as pd
from pathlib import Path
import win32com.client as win32

gerentes = pd.read_excel(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\Email_Automatico\Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\Email_Automatico\Bases de Dados\Lojas.csv', sep=';', encoding='latin-1')
vendas = pd.read_excel(r'C:\Users\mobishopgamer\Documents\RepositorioGitHub\Email_Automatico\Bases de Dados\Vendas.xlsx')

vendas_lojas = vendas.merge(lojas, on='ID Loja')

dic_lojas = {}

for loja in lojas['Loja']:
    dic_lojas[loja] = vendas_lojas.loc[vendas_lojas['Loja']==loja, :]

ultimo_dia = vendas_lojas['Data'].max()
dia_indicador = f'{ultimo_dia.day}-{ultimo_dia.month}'

pasta_backup = Path(r'Backup do Dia')

for loja in dic_lojas:
    if not (pasta_backup/loja).exists():
        (pasta_backup/loja).mkdir()

    nome_arquivo = f'{loja}({dia_indicador}).xlsx'
    local_save = pasta_backup / loja / nome_arquivo
    dic_lojas[loja].to_excel(local_save)

meta_faturamento = (1000, 1650000)
mf_diario, mf_anual = meta_faturamento
meta_produtos = (4, 120)
mp_diario, mp_anual = meta_produtos
meta_media = 500

dic_indicadores = {}

for loja in lojas['Loja']:

    loja_df = dic_lojas[loja]

    faturamento_diario = loja_df.loc[loja_df['Data'] == ultimo_dia, :]
    faturamento_diario = faturamento_diario['Valor Final'].sum()
    faturamento_diario = f'{faturamento_diario:.2f}'

    faturamento_anual = dic_lojas[loja]['Valor Final'].sum()
    faturamento_anual = f'{faturamento_anual:.2f}'

    media_diario = loja_df.loc[loja_df['Data'] == ultimo_dia, :]
    media_diario = loja_df['Valor Final'].mean()
    media_diario = f'{media_diario:.2f}'

    media_anual = loja_df['Valor Final'].mean()
    media_anual = f'{media_anual:.2f}'

    produtos_diario = loja_df.loc[loja_df['Data'] == ultimo_dia, :]
    produtos_diario = produtos_diario['Produto'].unique()
    produtos_diario = f'{len(produtos_diario)}'

    produtos_anual = loja_df['Produto'].unique()
    produtos_anual = f'{len(produtos_anual)}'

    dic_indicadores.update({loja: ([float(faturamento_diario), float(media_diario),float(produtos_diario)],[float(faturamento_anual), float(media_anual), float(produtos_anual)])})

meta_txt = ('🟢', '🔴')
positivo, negativo = meta_txt

dic_metas = {}

for loja in dic_indicadores:

    result_diario, result_anual = dic_indicadores[loja]

    if float(result_diario[0]) >= mf_diario: 
        fd_result = positivo
    else:
        fd_result = negativo

    if float(result_anual[0]) >= mf_anual:
        fa_result = positivo
    else:
        fa_result = negativo

    if float(result_diario[1]) >= meta_media:
        md_result = positivo
    else:
        md_result = negativo

    if float(result_anual[1]) >= meta_media:
        ma_result = positivo
    else:
        ma_result = negativo

    if float(result_diario[2]) >= mp_diario:
        pd_result = positivo
    else:
        pd_result = negativo

    if float(result_anual[2]) >= mp_anual:
        pa_result = positivo
    else:
        pa_result = negativo
    
    dic_metas.update({loja: ([fd_result, md_result, pd_result], [fa_result, ma_result, pa_result])})

    outlook = win32.Dispatch('outlook.application')

for i, loja in enumerate(gerentes['Loja']):
  if loja == 'Diretoria':
    break
  
  mail = outlook.CreateItem(0)
  mail.To = gerentes['E-mail'][i]
  mail.Subject = f'Feedback - {loja}'
  mail.HTMLBody = f'''<h1>Bom dia, {gerentes['Gerente'][i]}</h1>
  <p>O resultado de ontem (dia {ultimo_dia.day}/{ultimo_dia.month}) referênte a loja {loja} foi:</p>
  <br>
  <table>
    <tr>
      <th>Indicador</th>
      <th>Valor Dia</th>
      <th>Meta Dia</th>
      <th>Cenário Dia</th>
    </tr>
    <tr>
      <td>Faturamento</td>
      <td style="text-align: center">R${result_diario[0]:,.2f}</td>
      <td style="text-align: center">R${mf_diario:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][0][0]}</td>
    </tr>
    <tr>
      <td>Diversidade de Produtos</td>
      <td style="text-align: center">{result_diario[2]:,.2f}</td>
      <td style="text-align: center">{mp_diario:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][0][2]}</td>
    </tr>
    <tr>
      <td>Ticket Médio</td>
      <td style="text-align: center">R${result_diario[1]:,.2f}</td>
      <td style="text-align: center">R${meta_media:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][0][1]}</td>
    </tr>
  </table>
  <br>
  <table>
    <tr>
      <th>Indicador</th>
      <th>Valor Ano</th>
      <th>Meta Ano</th>
      <th>Cenário Ano</th>
    </tr>
    <tr>
      <td>Faturamento</td>
      <td style="text-align: center">R${result_anual[0]:,.2f}</td>
      <td style="text-align: center">R${mf_anual:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][1][0]}</td>
    </tr>
    <tr>
      <td>Diversidade de Produtos</td>
      <td style="text-align: center">{result_anual[2]:,.2f}</td>
      <td style="text-align: center">{mp_anual:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][1][2]}</td>
    </tr>
    <tr>
      <td>Ticket Médio</td>
      <td style="text-align: center">R${result_anual[1]:,.2f}</td>
      <td style="text-align: center">R${meta_media:,.2f}</td>
      <td style="text-align: center">{dic_metas[loja][1][1]}</td>
    </tr>
  </table>
  <br>
  <p>Segue em anexo a planilha com todos os dados para mais detalhes. <br>Qualquer Dúvida estou à disposição.</p>
  <br>
  <p>Att., <br>Nicolas Arthur</p>
  '''

  attachment = rf'C:\Users\mobishopgamer\Documents\Estudo\Hashtag\Python\AulasAplicaçoes\EmailProgramado\Backup do Dia\{loja}\{loja}({dia_indicador}).xlsx'
  mail.Attachments.Add(str(attachment))

  mail.Send()
  print(f'Email da loja {loja} enviado à {gerentes["Gerente"][i]}!')

ranking_anual = vendas_lojas[['Loja', 'Valor Final']].groupby('Loja').sum()
ranking_anual['Valor Final'] = ranking_anual

ranking_anual['Valor Final'] = [f'R${valor:,.2f}' for valor in ranking_anual['Valor Final']]

ranking_anual = ranking_anual.sort_values(by='Valor Final', ascending=False)

ranking_anual.to_excel(r'Backup do Dia/Diretoria/Rankind do Ano.xlsx')

ranking_diario = vendas_lojas.loc[vendas_lojas['Data']==ultimo_dia, :]
ranking_diario = ranking_diario[['Loja', 'Valor Final']].groupby('Loja').sum()
ranking_diario['Valor Final'] = [f'R${valor:,.2f}' for valor in ranking_diario['Valor Final']]
ranking_diario.sort_values(by='Valor Final', ascending=False)

ranking_diario.to_excel(r'Backup do Dia/Diretoria/Rankind do Dia.xlsx')

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.to = gerentes.loc[gerentes['Gerente']=='Diretoria', 'E-mail'].values[0]
mail.subject = f'Ranking das lojas - {ultimo_dia.day}/{ultimo_dia.month}'
mail.body = f'''
Bom dia,

Melhor faturamento do dia: {ranking_diario.index[0]} - {ranking_diario.iloc[0,0]}
Pior faturamento do dia: {ranking_diario.index[-1]} - {ranking_diario.iloc[-1,0]}

Melhor faturamento do ano: {ranking_anual.index[0]} - {ranking_anual.iloc[0,0]}
Pior faturamento do ano: {ranking_anual.index[-1]} - {ranking_anual.iloc[-1,0]}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Nico Arth
'''

attachment = r'C:\Users\mobishopgamer\Documents\Estudo\Hashtag\Python\AulasAplicaçoes\EmailProgramado\Backup do Dia\Diretoria\Rankind do Dia.xlsx'
mail.Attachments.Add(str(attachment))
attachment = r'C:\Users\mobishopgamer\Documents\Estudo\Hashtag\Python\AulasAplicaçoes\EmailProgramado\Backup do Dia\Diretoria\Rankind do Ano.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('Email Ranking enviado à Diretoria!')