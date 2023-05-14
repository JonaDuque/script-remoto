import sys
print(sys.path.append('/pysftp'))

from pathlib import Path
import win32com.client
import hashlib
import datetime
import os.path
import pysftp

# CARPETA DE SALIDA s
out_folder = Path.cwd() / "Output"
print(Path.cwd())
out_folder.mkdir(parents=True, exist_ok=True) # VALIDAR QUE EXISTA

# CONECTAR A OUTLOOK
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

# CONECTAR A BANDEJA DE ENTRADA
inbox = outlook.GetDefaultFolder(6).Items.Restrict("[LastModificationTime] > '02/04/2023'" and "[LastModificationTime] < '12/05/2023'")

# GET MESSAGES ELEMENTS
date = datetime.datetime.now()
messages = inbox

for message in messages:
    subject = message.Subject
    body = message.Body
    attachments = message.Attachments
    mail_id = hashlib.md5(message.EntryID.encode())

    print(mail_id.hexdigest())
    
    # CREAR CARPETAS SEPARADAS PARA CADA MENSAJE
    target_folder = out_folder / str(date.strftime('%Y')) / str(date.strftime('%m')) / 'prueba'
    target_folder.mkdir(parents=True, exist_ok=True) # VALIDAR QUE EXISTA
    
    # ADJUNTOS
    for attachment in attachments:
        print(attachment)
        if not os.path.isfile(str(target_folder)+str(attachment)):
            if ".pdf" in str(attachment) or ".xml" in str(attachment):
                attachment.SaveAsFile(target_folder / str(attachment))
                print(target_folder)
                
# DATE
print(date.strftime('%Y-%m-%d'))

# CONECTION TO SFTP & UPLOAD FILE
#cnopts = pysftp.CnOpts()
#cnopts.hostkeys = None
#with pysftp.Connection(host='my.uxlabs.mx', username='jonathan-uxlabs@api.uxlabs.mx', password='Tuxedo0827', cnopts = cnopts, port=222) as sftp:
#    print("Connection successfully established ... ")
#    sftp.put_r(target_folder, target_folder, preserve_mtime=True)
#    print("Upload file ")
#    sftp.close()
#f = ftpretty('my.uxlabs.mx', 'jonathan-uxlabs@api.uxlabs.mx', 'Tuxedo0827' )
#f.put(str(target_folder)+'Duque.pdf', 'prueba/Duque.pdf')
