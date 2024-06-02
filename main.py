#1. Scaricamento 4 Excel
import tableau_details as td
import restAPI as restAPI
#2. Import, Scaricamento Excel unico, Cancellazione vecchi Excel
import pandas as pd
import os 
from datetime import datetime, timedelta
#3. Invio Mail
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

now = datetime.now()
print(now)
today = yesterday = str(datetime.strftime(datetime.now(), '%Y-%m-%d'))
yesterday = str(datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d'))

with open("logs.txt", "w") as o: # ",w" write, tutti gli altri "a", append
    o.write('*** logs file del {} ***'.format(now))
    o.writelines('\n')
    o.writelines('\n')
    o.write('Pacchetti importati')
    o.writelines('\n')
    o.close()

### PARTE 1 ###


###  PARAMETERS

# Tableau Server site
site_name = '' # none for TECH

server = td.server
username = td.username
password = td.password

view_name_list = ["file1", "file2", "file3", "file4"]
file_temp = "Output"
filename = 'Report.xlsx'

directory = os.getcwd()

## SCRIPT TASK

# Authentication
auth_token, site_id, user_id = restAPI.sign_in(server, username, password, site_name)
with open("logs.txt", "a") as o:
    o.write('Login a Tableau Server')
    o.writelines('\n')
    o.close()

#per le 4 viste che vogliamo scaricare
for view_name in view_name_list: 

    # Get the views
    for i in range(1, 1000):
        try:
            view_id = restAPI.get_view_id(server, auth_token, site_id, view_name, 100, i)  
            print('vista {}'.format(view_id))
            break
        except: 
            continue
              
    # Download the view
    response = restAPI.download_excel_view(server, auth_token, site_id, view_id)
    print('scarico vista {}'.format(response))
    with open("logs.txt", "a") as o:
        o.write('Download vista "{0}" (id: {1})'.format(view_name, view_id))
        o.writelines('\n')
        o.close()

    # Write in the file
    with open(file_temp + ' - ' + view_name + '.xlsx', "wb") as file:
        file.write(response.content)

# Sign out
restAPI.sign_out(server, auth_token)
with open("logs.txt", "a") as o:
        o.write('Sign out')
        o.writelines('\n')
        o.close()


### PARTE 2: Caricare i 4 excel, salvarli in un unico file formato da fogli diversi, cancellare i vecchi file ###

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(filename, engine='xlsxwriter')
with open("logs.txt", "a") as o:
        o.write('Creazione del file giornaliero da inviare')
        o.writelines('\n')
        o.close()

#Leggi i 3 Excel che scarichi
for i in range(len(view_name_list)):
    #Import 
    df = pd.read_excel(file_temp + ' - ' + view_name_list[i] + '.xlsx')
    #non so perché, ma salta il nome delle colonne nella prima view
    if i == 0:
        df.columns = df.iloc[0]
        df.drop([0], inplace=True)
        #conversione colonne da testo a int
        field_to_convert_into_int = ['field1', 'field2', ... , 'Year  ', 'Year N Ser  ']
        for field_to_convert in field_to_convert_into_int:
            df[field_to_convert] = df[field_to_convert].astype(int)
    if i == 2 or i == 3:
        #devo riempire le prime due colonne "nan" con la cella precedente affiché scarichi l'excel con le celle piene (non come le viste in tableau)
        df.iloc[:, 0].fillna(method='ffill', inplace=True)
        df.iloc[:, 1].fillna(method='ffill', inplace=True)
    df.to_excel(writer, sheet_name = view_name_list[i], index = False)
    with open("logs.txt", "a") as o:
        o.write('Finalizzazione file giornaliero finale ' + str(i+1) + '/4 ...')
        o.writelines('\n')
 
# Close the Pandas Excel writer and output the Excel file.
writer._save()
with open("logs.txt", "a") as o:
        o.write('Salvataggio file')
        o.writelines('\n')


### PARTE 3: Invio Mail ###

#from
me = "tableau-tech@tech.it"
#to
file_path_to = os.path.join(directory, 'mailing_list', 'list_to.txt')
with open(file_path_to) as file:
    to = file.read().rstrip()
#cc
file_path_cc = os.path.join(directory, 'mailing_list', 'list_cc.txt')
with open(file_path_cc) as file:
    cc = file.read().rstrip()

msg = MIMEMultipart()

msg['From'] = me
msg['To'] = to
msg['Cc'] = cc
msg['Subject'] = "Report del {}".format(yesterday)

body = "Ciao, in allegato trovi il Report del {}".format(yesterday)

#allegato inizio
msg.attach(MIMEText(body, 'plain'))

attachment_t = os.path.join(directory, filename)
attachment = open(attachment_t, "rb")

part = MIMEBase('application', 'octet-stream')
part.set_load((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)
with open("logs.txt", "a") as o:
        o.write('File Excel allegato alla mail')
        o.writelines('\n')
#fine allegato

server = smtplib.SMTP('smtprelay.gruppo.net', 25)
text = msg.as_string()
server.sendmail(me, msg["To"].split(",") + msg["Cc"].split(","), msg.as_string())
#server.sendmail(me, rcpt, msg.as_string())
server.quit()
with open("logs.txt", "a") as o:
        o.write('Email inviata')
        o.writelines('\n')

#Tutto è andato bene, meno male :)
with open("logs.txt", "a") as o:
        o.write('Programma eseguito correttamente')
        o.writelines('\n')
        