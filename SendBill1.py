# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 20:37:46 2020

@author: Corkidi
"""


from tkinter import *
from tkinter import Menu
from tkinter import messagebox
from tkinter import filedialog
import csv
import pickle
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
from prettytable import PrettyTable

window = Tk() #tkinter window instance
window.title("Estados de Cuenta Epcot")
window.geometry('510x410')

var1 = StringVar()
sentEmails = StringVar()

statement = PrettyTable(['Fecha', 'Numero','Factura', 'Abono',  'Saldo'])

frameCSV = Frame(window, width = 500, height = 200, bd = 5, relief = RAISED)
frameCSV.grid(row = 0, column = 0, sticky = N+E+W, padx = 5)
frameCSV.grid_propagate(False)
frameSend = Frame(window, width = 500, height = 200, bd = 5, relief = RAISED)
frameSend.grid(row = 1, column = 0, sticky = N+S+E+W, padx = 5)
frameSend.grid_propagate(False)

with open('config.dat', 'rb') as f:
            outgoingServer, port, myEmail, myPassword = pickle.load(f)
server = smtplib.SMTP_SSL(outgoingServer, port)
server.ehlo()
user = myEmail
password = myPassword







#load whole Estados de Cuenta file procedure
saldos_file = 'saldos.csv'

def loadSaldos():
    try:
        reader = csv.reader(open(saldos_file, 'r'), delimiter = ';')
        d = []  #empty list
        for row in reader: 

           d.append(row)
#        print(len(d))   
        return d
        
        reader.close()
    except:
        messagebox.showwarning('Error', 'Tu lista de saldos no fue cargada.')

saldos = loadSaldos()

#show on screen that CSV file has loaded correctly
def postSaldos():

    var1.set("Se cargaron " + str(len(saldos)) + " registros exitosamente.")   

   

#create account statement per account and prepare to send

def buildStatement(client):  
    #verify data on every line and find values and create 
    outgoingEmail = ''  
    saldoAnterior = ''
    totalAPagar = ''
    statement.clear_rows()
   
    for i in range(len(saldos)):
        if saldos[i][0] == client:
            date = ''
            transactionNumber = ''
            facturaAmount = ''
            reciboAmount = ''
            ongoingSaldo =''
             #find and verify email, if no email, skip to next client
            for j in range(len(saldos[i])):
                if outgoingEmail == '':
                    if saldos[i][j] == 'E-mail :':
                        outgoingEmail = saldos[i][j-1] 
#                        outgoingEmail = ''
     # find and verify saldo
                if saldoAnterior == '':
                    if saldos[i][j] == 'Saldo  Anterior:  ========>>>':
                        saldoAnterior = saldos[i][j+1]
    #find and verify Total a Pagar
                if totalAPagar == '':
                    if saldos[i][j] == 'Total a pagar :  ==========>>>':
                        totalAPagar = saldos[i][j-1]
    #Find and verify date
                if date == '':
                    if saldos[i][j] == 'Observaciones':
                        date = saldos[i][j+2]
        #find and verify nunmber of Factura or Recibo
                        transactionNumber = saldos[i][j+5]
        #find and verify amount of Factura
                if saldos[i][j] == 'FACTURAS':
                    facturaAmount = ' '
                    facturaAmount = saldos[i][j+3]
    #find and verify amount of Recibo
                if saldos[i][j] == 'RECIBOS':
                    reciboAmount = ' '
                    reciboAmount = saldos[i][j+4]
    #find and verify ongoing saldo
                if saldos[i][j] == 'Observaciones':
                    ongoingSaldo = saldos[i][j+8]  
    #build the table for email                     
            statement.add_row([date, transactionNumber, facturaAmount, reciboAmount,ongoingSaldo])             
    textIn = client + '\n' + 'Saldo  Anterior:  ========>>> ' + saldoAnterior + \
        '\n' + statement.get_string() + '\n' + 'Total a pagar :  ==========>>>' + totalAPagar
    buildEmail(textIn, outgoingEmail)
 

def buildEmail(textIn, outgoingEmail):
            
    try:           
        text = textIn   
        part1 = MIMEText(text, 'plain')
#        html = """\
#                <html>
#                  <head></head>
#                  <body>
#                    <p>Hi!<br>
#                       How are you?<br>
#                       Here is the <a href="https://www.python.org">link</a> you wanted.
#                    </p>
#                  </body>
#                </html>
#                """
#        part2 = MIMEText(html, 'html')         
        message = MIMEMultipart()
        message['From'] = myEmail       
        message['To'] = outgoingEmail         
        message['Subject'] = 'Tu Estado de Cuenta de Epcot'
        message.attach(part1)
#        message.attach(part2)                
        server.sendmail(user, outgoingEmail, message.as_string()) 
        time.sleep(1)

    except:
        messagebox.showwarning('Error', 'Problema enviando los correos.')
        
        
def sendEmails():
    #create loop to verufy that we are working on one account at a time
    try:
       
        server.login(user, password)
           
        for i in range(len(saldos)):
            if i == 0:
                buildStatement(saldos[i][0])
                
            elif saldos[i][0] != saldos[i - 1][0]:
                buildStatement(saldos[i][0])
        
        mailStat = Toplevel(window)
        mailStat.title("Email Status")
        mailStat.geometry('180x180')
        status = Label(mailStat, text ='Enviando los estados.')
        status.grid(column=0, row=0)
        def closeWindow():
            mailStat.destroy()
            
        sentMail=Label(mailStat, text='Estados enviados correctamente.')        
        sentMail.grid(column=0, row=2) 
        mailOk = Button(mailStat, text="OK", command=closeWindow)
        mailOk.grid(column=0, row=4)
        server.quit()
    except: 
        errorMail=Label(mailStat, text='Error, contacta a tu administrador.')
        errorMail.grid(column=0, row=3)        
        mailOk = Button(mailStat, text="OK", command=closeWindow)
        mailOk.grid(column=0, row=4)
        
def configWin():  #configuration window for setting up email account, headers, clients file
    configWindow = Toplevel(window)
    configWindow.title('Configuraciones')
    configWindow.geometry('350x250')
    emailConfigFrame = Frame(configWindow, width = 400, height = 200, bd = 5, relief = RAISED)
    emailConfigFrame.grid(column = 0, row = 0)
    configLbl = Label(emailConfigFrame, text = 'Email Configuration')
    configLbl.grid(column = 0, row = 0)
    
    configLbl1 = Label(emailConfigFrame, text = 'Outgoing eMail Server: ')
    configLbl1.grid(column = 0, row = 1)
    configEntry1 = Entry(emailConfigFrame, width = 30)
    configEntry1.grid(column = 1, row = 1)
    
    configLbl2 = Label(emailConfigFrame, text = 'Port: ')
    configLbl2.grid(column = 0, row = 2)
    configEntry2 = Entry(emailConfigFrame, width = 30)
    configEntry2.grid(column = 1, row = 2)
    
    configLbl3 = Label(emailConfigFrame, text = 'Password: ')
    configLbl3.grid(column = 0, row = 3)
    configEntry3 = Entry(emailConfigFrame, width = 30)
    configEntry3.grid(column = 1, row = 3)
    
    clientConfigFrame = Frame(configWindow, width = 400, height = 200, bd = 5, relief = RAISED)
    clientConfigFrame.grid(column = 0, row = 1)
    
    configLbl4 = Label(clientConfigFrame, text = 'Mi correo: ')
    configLbl4.grid(column = 0, row = 0)
    configEntry4 = Entry(clientConfigFrame, width = 30)
    configEntry4.grid(column = 1, row = 0)    
    
    
    try:
        with open('config.dat', 'rb') as f:
            outgoingServer, port, myEmail, myPassword = pickle.load(f)
        f.close()
        configEntry1.insert(0, outgoingServer)
        configEntry2.insert(0, port)
        configEntry3.insert(0, myPassword)
        configEntry4.insert(0, myEmail)
    except:
        pass
    
    def enterConfig():
        outgoingServer = configEntry1.get()
        port = configEntry2.get()
        myEmail = configEntry4.get()
        myPassword = configEntry3.get()
        with open('config.dat', 'wb') as f:
            pickle.dump((outgoingServer, port, myEmail, myPassword), f)
        f.close()
        configWindow.destroy()
    
    enterConfigBtn = Button(clientConfigFrame, text = 'Enter', command = enterConfig)
    enterConfigBtn.grid(row = 4, column = 3, padx = 15, pady = 10)    
 

#tkinter window design


selectCSVLbl = Label(frameCSV, text="Selecciona tu CSV", font=("Helvetica", 16, 'underline'))
selectCSVLbl.grid(column=0, row=0)
btnLoadSaldos = Button(frameCSV, text = "Cargar Saldos", command = postSaldos)
btnLoadSaldos.grid(column = 0, row = 1)
loadedCSVLbl = Label(frameCSV, textvariable= var1, font=("Helvetica", 16, 'underline'))
loadedCSVLbl.grid(column=0, row=2)

btnConfigEmail = Button(frameSend, text = "Configuracion de correo",\
                        font=("Helvetica", 16, 'underline'), command = configWin)
btnConfigEmail.grid(column = 0, row = 0)

btnSendEmails = Button(frameSend, text = "Enviar Estados", \
                       font=("Helvetica", 16, 'underline'), command = sendEmails)
btnSendEmails.grid(column=0, row=1)



window.mainloop()
