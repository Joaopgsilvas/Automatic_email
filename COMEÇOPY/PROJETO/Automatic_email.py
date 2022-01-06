import pandas as pd
import smtplib
import email.message
tabela_vendas = pd.read_excel('Vendas.xlsx') #importando a base de dados

#vizualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-'*50)
#produtos vendidos
produtos_vendidos = tabela_vendas[['ID Loja','Quantidade','Valor Final']].groupby('ID Loja').sum()
print(produtos_vendidos)
print('-'*50)
#ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / produtos_vendidos['Quantidade']).to_frame()
print(ticket_medio)


#enviar o email com o relatório
def enviar_email():
    corpo_email = f"""
    <p>enviando o email com o faturamento de todas as lojas 
    
    faturamento:
    {faturamento.to_html()}
    produtos vendidos:
    {produtos_vendidos.to_html()}
    tickets medios :
    {ticket_medio.to_html()}</p> 
    """
    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'remetente'
    password = 'senha'
    msg['To'] = 'destinatario'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email foi enviado com sucesso ! ')

enviar_email()