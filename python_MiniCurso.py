#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd

#importar a db
tabela_vendas = pd.read_excel('Vendas.xlsx')


#visualizar a db
pd.set_option('display.max_columns',None)
print(tabela_vendas)

# Lógica do Programa


# In[8]:


print("=" * 50)
# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)


# In[10]:


print("=" * 50)
# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)


# In[20]:


print("=" * 50)
# Ticket médio (faturamento / qtd)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() #to_frame transforma essa série de dados em tabela #as operacoes entre tabelas retornam listas de valores e não  uma nova tabela 
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)


# In[19]:


# enviar email com relatório
# enviar email com relatório
import win32com.cliente as win32

outlook = win32.Dipatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = u'luizcarlos_bsb2006@hotmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<h1>Prezados, </h1>
<p>Segue o realatório de Vendas por Loja</p>
<p>Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:_.2f}'.format})}

<p>Quantidade: </p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:_.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>

<p>Atenciosamente,</p>
<p>Luiz</p>
'''

mail.send()
print("Email enviado.)


# In[ ]:




