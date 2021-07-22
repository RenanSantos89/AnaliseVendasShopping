import pandas as pd
import win32com.client as win32

# importar base de dados

tb_vendas = pd.read_excel('Vendas_Shopping.xlsx')

# Visualizar a base de dados

pd.set_option('display.max_columns', None)
#print(tb_vendas)

# Faturamento por Loja

Faturamento_loja = tb_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(Faturamento_loja)
print('-'* 50)

# Quantidade de produtos vendidos por loja

QTD_prod_loja = tb_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(QTD_prod_loja)
print('-'* 50)

# Ticket medio por produto em cada loja

Tic_medio = (Faturamento_loja['Valor Final'] / QTD_prod_loja['Quantidade']).to_frame()
Tic_medio = Tic_medio.rename(columns={0: 'Ticket Medio'})
print(Tic_medio)

# enviar por email
# deixa o outlook do computador configurado para o email de envio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'renan.3@gmail.com'
mail.Subject = 'Relatorio de Vendas por Loja'
#: formatando numero, , separador de milhar , . separador dedecimal , 2f duas casas decimais
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatorio de Vednas por cada loja.</p>

<p>Faturamento:</p>
{Faturamento_loja.to_html(formatter={'Valor Final': 'R${:,.2f}'.format })}  

<p>Quantidade Vendida: </p>
{QTD_prod_loja.to_html()}

<p>Ticket medio dos produtos em cada Loja: </p>
{Tic_medio.to_html(formatter={'Ticket Medio': 'R${:,.2f}'.format })}

<p>Att, </p>
<p>Renan Silva Santos </p>
'''

mail.Send()

print('Email Enviado')