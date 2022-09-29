from typing import Counter
from nbformat import write
import pyodbc 
from mysql.connector.errors import ProgrammingError
import win32com.client as win32
from datetime import datetime

# CONECTANDO BANCO DE DADOS
con = pyodbc.connect("Driver={SQL Server};Server=172.16.60.3;Database=SGMC_COS;UID=sa;PWD=*******;") 

# PESQUISA BANCO DE DADOS
triagem = "SELECT EMISSOR_DESCRICAO, ITEM,  COUNT(*) FROM VWDATASETRELATORIO WHERE TICKET_ESTADO = 'Pre cadastro' AND TICKET_SEQUENCIAL <> 0 GROUP BY EMISSOR_DESCRICAO, ITEM"
desc_interna = "SELECT EMISSOR_DESCRICAO, ITEM,  COUNT(*) FROM VWDATASETRELATORIO WHERE TICKET_ESTADO = 'Pesagem inicial' AND EMISSOR_DESCRICAO = 'BR - PAULINIA' GROUP BY EMISSOR_DESCRICAO, ITEM"
carr_interno = "SELECT EMISSOR_DESCRICAO, ITEM,  COUNT(*) FROM VWDATASETRELATORIO WHERE TICKET_ESTADO = 'Pesagem inicial' AND EMISSOR_DESCRICAO <> 'BR - PAULINIA' AND EMISSOR_DESCRICAO <> 'STOLLER' GROUP BY EMISSOR_DESCRICAO, ITEM"
# rec_final = "SELECT EMISSOR_DESCRICAO, ITEM, COUNT(EMISSOR_DESCRICAO), SUM(OPERACAO_PESO_LIQUIDO) FROM VWDATASETRELATORIO WHERE CONVERT(DATE,TICKET_DATA) = CONVERT (date, GETDATE()) AND TICKET_ESTADO = 'Pesagem Final' AND EMISSOR_DESCRICAO = 'BR - PAULINIA' GROUP BY EMISSOR_DESCRICAO, ITEM"
exp_final = "SELECT EMISSOR_DESCRICAO, ITEM, COUNT(EMISSOR_DESCRICAO), SUM(OPERACAO_PESO_LIQUIDO), FORMAT(SYSDATETIME(), '%H') AS HORA_FORMATADA FROM VWDATASETRELATORIO WHERE CONVERT(DATE,FINAL_OPERACAO_DATA) = CONVERT (date, GETDATE()) AND TICKET_ESTADO = 'Pesagem Final' AND EMISSOR_DESCRICAO <> 'BR - PAULÍNIA' GROUP BY EMISSOR_DESCRICAO, ITEM"
# exp_total = "SELECT SUM(OPERACAO_PESO_LIQUIDO), COUNT(*) FROM VWDATASETRELATORIO WHERE CONVERT(DATE,TICKET_DATA) = CONVERT (date, GETDATE()) AND TICKET_ESTADO = 'Pesagem Final' AND EMISSOR_DESCRICAO <> 'BR - PAULÍNIA'"
exp_total = "SELECT SUM(OPERACAO_PESO_LIQUIDO), COUNT(*), SUM(DATEDIFF(MINUTE, INICIAL_OPERACAO_DATA, FINAL_OPERACAO_DATA)) FROM VWDATASETRELATORIO WHERE CONVERT(DATE,FINAL_OPERACAO_DATA) = CONVERT (date, GETDATE()) AND TICKET_ESTADO = 'Pesagem Final' AND EMISSOR_DESCRICAO <> 'BR - PAULÍNIA'"
rec_final = "SELECT EMISSOR_DESCRICAO, ITEM, COUNT(EMISSOR_DESCRICAO), SUM(OPERACAO_PESO_LIQUIDO), SUM(DATEDIFF(MINUTE, INICIAL_OPERACAO_DATA, FINAL_OPERACAO_DATA)) FROM VWDATASETRELATORIO WHERE CONVERT(DATE,TICKET_DATA) = CONVERT (date, GETDATE()) AND TICKET_ESTADO = 'Pesagem Final' AND EMISSOR_DESCRICAO = 'BR - PAULINIA' GROUP BY EMISSOR_DESCRICAO, ITEM"

# CONECTANDO OUTLOOK
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = "ta.dacio@grupounimetal.com.br; ls.freitas@grupounimetal.com.br; sa.bianchin@grupounimetal.com.br; cco.cosmopolis@grupounimetal.com.br; faturamento.cos@grupounimetal.com.br; logistica.cos@grupounimetal.com.br"
email.Subject = "CENÁRIO - UNIDADE COSMÓPOLIS"



imagem = "https://www.grupounimetal.com.br/wp-content/uploads/sites/1060/2022/05/Logo.png"
pi_1 = ''
pi_2 = ''
pi_3 = ''
pi_4 = ''
carr_1 = ''
carr_2 = ''
carr_3 = ''
carr_4 = ''
exp_1 = ''
exp_2 = ''
exp_3 = ''
exp_4 = ''
tria_1 = ''
tria_2 = ''
tria_3 = ''
tria_4 = ''
rec_1 = ''
rec_2 = ''
rec_3 = ''
rec_4 = ''


# CONFIGURAÇÃO TRIGEM
with con:
    try:
        cursor = con.cursor()
        cursor.execute(triagem)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            tria_1 += (f'{contato[0]}<br><br>')
            tria_2 += (f'{contato[1]}<br><br>')
            tria_3 += (f'{contato[2]}<br><br>')
            tria_4 += (f"")


# CONFIGURAÇÃO CARREGAMENTO INTERNO
with con:
    try:
        cursor = con.cursor()
        cursor.execute(carr_interno)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            carr_1 += (f'{contato[0]}<br><br>')
            carr_2 += (f'{contato[1]}<br><br>')
            carr_3 += (f'{contato[2]}<br><br>')
            carr_4 += (f">{''}<br><br>")
            

# CONFIGURAÇÃO DESCARGA INTERNA
with con:
    try:
        cursor = con.cursor()
        cursor.execute(desc_interna)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            pi_1 += (f'{contato[0]}<br><br>')
            pi_2 += (f'{contato[1]}<br><br>')
            pi_3 += (f'{contato[2]}<br><br>')
            pi_4 += (f"SEM DADOS<br>")


# CONFIGURAÇÃO RECEBIMENTO FINAL
with con:
    try:
        cursor = con.cursor()
        cursor.execute(rec_final)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            rec_1 += (f'{contato[0]}<br><br>')
            rec_2 += (f'{contato[1]}<br><br>')
            rec_3 += (f'{contato[2] + 0}')
            rec_4 += (f'{contato[3] / 1000}')
            rec_5 = (contato[4] / contato[2])
    


# CONFIGURAÇÃO EXPEDICAO FINAL
with con:
    try:
        cursor = con.cursor()
        cursor.execute(exp_final)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            exp_1 += (f'{contato[0]}<br><br>')
            exp_2 += (f'{contato[1]}<br><br>')
            exp_3 += (f'{contato[2]}<br><br>')
            exp_4 += (f"{contato[3] / 1000} T<br><br>")
            exp_5 = (f'{contato[4]}<br><br>')


# CONFIGURAÇÃO EXPEDICAO FINAL TOTAL
with con:
    try:
        cursor = con.cursor()
        cursor.execute(exp_total)
        contatos = cursor.fetchall()
    except ProgrammingError as e:
        print(f'Erro: {e.msg}')
    else:
        for contato in contatos:
            # teste += (f'<table border="2"><tr><td>{contato[1]} | {contato[4]} | {contato[6]}</td></tr></table>') + '<br>'
            exp_total1 = (f'{contato[0] / 1000} Toneladas')
            exp_total2 = (contato[1])
            exp_total3 = (contato[0] / 1000)
            exp_total4 = (contato[2] / contato[1])

# RITMO EXPEDICAO
data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime('%H')

int_pesototal = int(float(exp_total3))
int_hora = int(data_e_hora_em_texto)
ritmo_hora = int_pesototal / int_hora
ritmo_dia = ritmo_hora * 24

# RITMO RECEBIMENTO
int_rectotal = int(float(rec_4))
ritmo_rechora = int_rectotal / int_hora
ritmo_recdia = ritmo_rechora * 24


email.HTMLBody ="""

<!DOCTYPE html>
<html>
<head>
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@300;400&display=swap');


* {margin:0; 
padding: 0; 
box-sizing: 
border-box;}

.content {
    display:flex; 
    margin: auto;
}

.rTable{
    width: 100%; 
    text-align: center;}

    .rTable thead{
        background: black; 
        font-weight: bold; 
        color:#fff;
    }

    .rTable tbody tr:nth-child(2n){
        background: #ccc;
    }
    .rTable th , .rTable td{
        padding: 7px 0;
    }

@media screen and (max-width: 480px){
    .content{
        width: 94%;
    }

    .rTable thead{
        display:none;
    }
    .rTable tbody td{
        display: flex; 
        flex-direction: column; 
    }
}

@media only screen and (min-width: 1200px){
    .content{
        width:100%;
    }
    .rTable tbody tr td:nth-child(1){
        width:10%;
    }
    .rTable tbody tr td:nth-child(2){
        width:30%;
    }
    .rTable tbody tr td:nth-child(3){
        width:20%;
    }
    .rTable tbody tr td:nth-child(4){
        width:10%;
    }
    .rTable tbody tr td:nth-child(5){
        width:30%;
    }
}


table {
    border-collapse: separate;
    width: 55%;
    table-layout: fixed;
    border: 2px solid rgb(255, 255, 255);
}

.titulo1{
    background-color: #1a1a1a;
    color: rgb(255, 255, 255);
    text-align: center;
    padding: 12px;
    font-family: 'Nunito', sans-serif;
    font-size: 18px;
}

.titulo2 {
    background-color: #e9e9e9;
    color: rgb(0, 0, 0);
    font-family: 'Nunito', sans-serif;
    padding: 3px;
    font-size: 12px;
}

.titulo3 {
    background-color: #ebf3d5;
    text-align: right;
    color: rgb(0, 0, 0);
    font-family: 'Nunito', sans-serif;
    padding: 10px;
    font-size: 16px;
}

.titulo4 {
    background-color: #e9e9e9;
    color: rgb(0, 0, 0);
    font-family: 'Nunito', sans-serif;
    padding: 16px;
    font-size: 12px;
    border: 3px solid #a8a8a8;
    border-radius: 30px;
    text-align: center;
    width: 55%;
}

.titulo5 {
    background-color: #6e6e6e;
    color: rgb(218, 218, 218);
    font-family: 'Nunito', sans-serif;
    padding: 8px;
    font-size: 15px;
}


h1 {
    font-family: 'Nunito', sans-serif;
    color: rgb(25, 0, 255);
}

th, td {
    color: rgb(0, 0, 0);
    text-align: center;
    padding: 20px;
    font-family: 'Nunito', sans-serif;
    font-size: 10px;
    border: 2px solid rgb(255, 255, 255);
}

h2 {
    font-family: 'Nunito', sans-serif;
}

tr:nth-child(even){background-color: #ffffff}

th {
    background-color: #e9e9e9;
    color: rgb(0, 0, 0);
    font-family: 'Nunito', sans-serif; 
    padding: 8px;   
    border: 2px solid rgb(255, 255, 255);    
}

</style>
</head>
"""f'''
<body>

    <table>
                <tr>
                    <td class="titulo5" colspan="2"><img src={imagem}><br></td>
                    <td class="titulo5" colspan="2">UNIDADE COSMÓPOLIS<br></td>
                </tr>
                <tr>
                    <td class="titulo1" colspan="4">CENÁRIO ATUAL</td>
                </tr>
                <tr>
                    <th class="titulo5" colspan="4">TRIAGEM</th>
                </tr>
                <tr>
                    <th colspan="2">CLIENTE</th>
                    <th>PRODUTO</th>
                    <th>QTD</th>
                </tr>
                <tr>
                    <th colspan="2">{tria_1}</th>
                    <th>{tria_2}</th>
                    <th>{tria_3}</th>
                </tr>
                <tr>
                    <th class="titulo5" colspan="4">DESCARGA INTERNA</th>
                </tr>
                <tr>
                    <th colspan="2">CLIENTE</th>
                    <th>PRODUTO</th>
                    <th>QTD</th>
                </tr>
                <tr>
                    <th colspan="2">{pi_1}</th>
                    <th>{pi_2}</th>
                    <th>{pi_3}</th>
                </tr>
            
                <tr>
                    <th class="titulo5" colspan="4">CARREGAMENTO INTERNO</th>
                </tr>
                <tr>
                    <th colspan="2">CLIENTE</th>
                    <th>PRODUTO</th>
                    <th>QTD</th>
                </tr>
                <tr>
                    <th colspan="2">{carr_1}</th>
                    <th>{carr_2}</th>
                    <th>{carr_3}</th>
                </tr>
    
                <tr>
                    <th class="titulo1" colspan="4">DADOS DO DIA</th>
                </tr>
    
                <tr>
                    <th class="titulo5" colspan="4">EXPEDIÇÃO</th>
                </tr>
    
                <tr>
                    <th>CLIENTE</th>
                    <th>PRODUTO</th>
                    <th>QTD</th>
                    <th>VOLUME</th>
                </tr>
                <tr>
                    <th>{exp_1}</th>
                    <th>{exp_2}</th>
                    <th>{exp_3}</th>
                    <th>{exp_4}</th>
                </tr>
    
                <tr>
                    <th class="titulo2" colspan="3">CARRETAS EXPEDIDAS<br></th>
                    <th class="titulo2">{exp_total2}<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">VOLUME EXPEDIDO<br></th>
                    <th class="titulo2">{exp_total1}<br></th>        
                </tr>
    
                <tr>
                    <th class="titulo2" colspan="3">RITMO P/ HORA<br></th>
                    <th class="titulo2">{ritmo_hora:.2f} Toneladas<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">PREVISÃO DE CARREGAMENTO DO DIA<br></th>
                    <th class="titulo2">{ritmo_dia:.2f} Toneladas<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">TEMPO MÉDIO CARREGAMENTO<br></th>
                    <th class="titulo2">{exp_total4:.0f} minutos<br></th>        
                </tr>
    
                <tr>
                    <th class="titulo5" colspan="4">RECEBIMENTO</th>
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">CARRETAS RECEBIDAS<br></th>
                    <th class="titulo2">{rec_3}<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">VOLUME RECEBIDO<br></th>
                    <th class="titulo2">{rec_4}<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">RITMO P/ HORA<br></th>
                    <th class="titulo2">{ritmo_rechora:.2f} Toneladas<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">PREVISÃO DE DESCARGA DO DIA<br></th>
                    <th class="titulo2">{ritmo_recdia:.2f} Toneladas<br></th>        
                </tr>
                <tr>
                    <th class="titulo2" colspan="3">TEMPO MÉDIO DESCARGA<br></th>
                    <th class="titulo2">{rec_5:.0f} minutos<br></th>        
                </tr>
    
     </table>
    
    </body>
    </html>
'''

email.send
print("Email enviado")