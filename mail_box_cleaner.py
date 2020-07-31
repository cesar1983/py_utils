#!/usr/bin/env python
'''
Programa executado via linha de comando

Efetua a limpeza de alguma caixa de emails. Ajuda o php a processar os emails corretamente.

1 - Remove emails com To ou Cc contendo mais de 50 destinatarios (normalmente emails de correntes)
2 - Remove qualquer email que nao contenha ao menos um arquivo xml ou zip
3 - Efetua um expunge para remover definitivamente as mensagens flegadas com DELETED

Cesar - 12/2014
'''

import time
import sys
import imaplib
import email
from email.parser import HeaderParser
imaplib.IMAP4.debug = imaplib.IMAP4_SSL.debug = 1
 
 
def mail_box_cleaner(host, port, login, password):
    
    print '\n==========================', login ,'==========================\n'
    
    # connect
    mail_server = imaplib.IMAP4(host, port)
 
    # authenticate
    mail_server.login(login, password)
    
    # seleciona a pasta 
    #mail_server.select('APAGAR') 
    #mail_server.select('teste') 

    mail_server.select('INBOX') 
    
    # le as pastas disponiveis 
    #folders = mail_server.list()
    #print '\nfolders\n'
    #print(folders)
    #print '\n'
    
    msg_to_remove = []
    
    typ, data = mail_server.search(None, 'ALL')
    for num in data[0].split():
        
        remove = 1
        
        typ, data = mail_server.fetch(num, '(RFC822)')
        
        mail = email.message_from_string(data[0][1])
        
        print '\n================================================================================\n'
        print 'NUM email ',num, '\n'
        
        # TODO QUANDO houverem mais de 50 destinatarios, deleta o email
        print 'From: '   , mail['From'], '\n'
        #print 'To : '    , mail['To'], '\n'
        #print 'Cc : '    , mail['Cc'], '\n'
        #print 'Titulo : ', mail['Subject'], '\n'
        

        # Se tiver mais de 50 emails em copia elimina esse email
        recipent_to_list = str(mail['To']).split(',')
        count_to = len(recipent_to_list)
        print 'Total TO LIST : ' , count_to
        if int(count_to) > 50 :
            msg_to_remove.append( num )
            continue 
            
        
        # Se tiver mais de 50 emails em copia elimina esse email
        recipent_cc_list = str(mail['Cc']).split(',')
        count_cc = len(recipent_cc_list)
        print 'Total CC LIST : ' , count_cc
        if int(count_cc) > 50 :
            msg_to_remove.append( num )
            continue 

        print '\n'
        
        #msg = HeaderParser().parsestr(data[0][1])
        #print msg['From']
        #print msg['To']
        #print msg['Subject']
        
        #print mail
        #print 'Message %s\n %s\n' % (num, data[0][1])
 
        # PERCORRE AS PARTES DO EMAIL, SE NAO HOUVER ALGUM ANEXO FORMATO XML, o EXCLUI 
        for part in mail.walk():
            
            filename = part.get_filename()
            
            try:
                #filename = filename.decode('ascii', 'ignore')
                filename = filename.encode('utf-8')
                filename_str = str(filename)
                print 'Anexo: ', filename_str, ' - ', part.get_content_subtype()

                #print 'Content-Type:', part.get_content_type(), '\n'
                #print 'Main Content:', part.get_content_maintype(), '\n'
                #print 'Sub Content:', part.get_content_subtype(), '\n'
                #print '\n\n'
            
                # Somente deixar as mensagens que contiverem um XML OU UM ZIP             
                if part.get_content_subtype() == 'xml' or part.get_content_subtype() == 'XML' or filename_str.find('.xml') > 0 or filename_str.find('.XML') > 0 or part.get_content_subtype() == 'zip' or part.get_content_subtype() == 'ZIP'  or filename_str.find('.zip') > 0 or filename_str.find('.ZIP') > 0 :
                    remove = 0
 
                #payload = part.get_payload()
                #print payload
                
            except: # catch *all* exceptions
                 e = sys.exc_info()[0]
                 remove = 0
                 print "Erro de charset na conversao do nome do arquivo para string : ", filename
                 print "\n Error: ", e
                                                               
        if remove == 1 :
            msg_to_remove.append( num )
            
              
    print '\n\n==================================== END ============================================\n'
    
    
    if len(msg_to_remove)>0 :
        
        print 'MSGS TO REMOVE \n'
        print msg_to_remove
        print '\n\n'
    
        msg_ids = ','.join(msg_to_remove)
        print 'Matching messages:', msg_ids
        
        # What are the current flags?
        typ, response = mail_server.fetch(msg_ids, '(FLAGS)')
        #print 'Flags before:', response

        # Copy the messages
        print 'COPYING:', msg_ids, ' to removidas_py folder'
        typ, response = mail_server.copy(msg_ids, 'removidas_py')
    
        # Change the Deleted flag
        typ, response = mail_server.store(msg_ids, '+FLAGS', r'(\Deleted)')
    
        # What are the flags now?
        typ, response = mail_server.fetch(msg_ids, '(FLAGS)')
        #print 'Flags after:', response
        
        # Really delete the message.
        typ, response = mail_server.expunge()
        print 'Eliminados:', response
       
        #msg_ids = ','.join(msg_to_remove)
        #typ, create_response = mail_server.create('APAGAR')
        #print 'CREATED APAGAR:', create_response
         
        # Look at the results
        #mail_server.select('APAGAR')
        #typ, [response] = mail_server.search(None, 'ALL')
        #print 'COPIED:', response
    
        # imap.store(num, '+FLAGS', '\\Deleted')
 
# usage example

host = ""
port = ""
username = ""
passowrd = ""

# LOOP INFINITO 
while(1):

    mail_box_cleaner( host, port, login, password ) 
     
     
    time.sleep(60) # dorme alguns segundos e verifica novamente
     

