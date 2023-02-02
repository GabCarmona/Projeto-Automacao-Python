# Importar arquivos e bibliotecas
import pandas as pd
import pathlib
import win32com.client as win32
import time

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', sep=';', encoding= 'latin1')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

#Criar uma tabela para cada loja e Definir o dia do indicador
vendas = vendas.merge(lojas, on='ID Loja')
dict_lojas = {}
for loja in lojas["Loja"]:
    dict_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

dia_indicador = vendas['Data'].max()

#Salvar a planilha na pasta de backup
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_backup = caminho_backup.iterdir()

lista_nomes_backup = []
for arquivo in arquivos_backup:
    lista_nomes_backup.append(arquivo.name) 

for loja in dict_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo =  caminho_backup / loja / nome_arquivo 
    dict_lojas[loja].to_excel(local_arquivo)

#Definição de Metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_tm_ano= 500
meta_tm_dia = 500


################## Cálculo do indicador ############################
for  loja in dict_lojas:
    #Faturamento
    vendas_loja = dict_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]
    faturamento_anual = vendas_loja['Valor Final'].sum()
    faturamento_diario = vendas_loja_dia['Valor Final'].sum()

    #Diversidade de Produtos
    qtde_produtos_ano = vendas_loja['Produto'].drop_duplicates().count()
    qtde_produtos_dia = vendas_loja_dia['Produto'].drop_duplicates().count()

    #Ticket Médio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    tm_ano = valor_venda['Valor Final'].mean()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    tm_dia = valor_venda_dia['Valor Final'].mean()

    #Enviar Email para o gerente
    outlook = win32.Dispatch('outlook.application')

    nome= emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = 'OnePage dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)

    if faturamento_diario >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_anual >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtde_produtos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if tm_dia >= meta_tm_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if tm_ano >= meta_tm_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p> Bom dia, {nome} </p>
    <p> O Resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi </p>

    <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_diario:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_dia}</td>
            <td style="text-align: center">{meta_qtde_produtos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${tm_dia:.2f}</td>
            <td style="text-align: center">R${meta_tm_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
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
            <td style="text-align: center">R${faturamento_anual:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_ano}</td>
            <td style="text-align: center">{meta_qtde_produtos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${tm_ano:.2f}</td>
            <td style="text-align: center">R${meta_tm_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
        </tr>
        </table>

    <p> Segue em anexo a planilha com todo os dados para mais detalhes. </p>
    <p> Qualquer dúvida estou na disposição. </p>
    <p> Att., Gabriel </p>

    '''
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx' 
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f'E-mail da loja {loja} enviado')
    time.sleep(3.5) ######## SEM O SLEEP SÓ ESTAVA CHEGANDO O PRIMEIRO EMAIL

#Criar Ranking para diretoria
faturamento_lojas = vendas.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']== dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

#Enviando email para diretoria
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = 'Ranking dia {}/{}'.format(dia_indicador.day, dia_indicador.month)
mail.Body = f'''
Prezados, bom dia

Melhor Loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior Loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor Loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior Loja do ano em faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas lojas.
Qualquer dúvida estou à disposição.
Att., Gabriel
'''

attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx' 
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx' 
mail.Attachments.Add(str(attachment))
mail.Send()

print('E-mail da diretoria enviado')