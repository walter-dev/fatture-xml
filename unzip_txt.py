# Directory di import INVB2B - Ver. 1.9.4 07-09-19
import zipfile, re, os, shutil, time, pyodbc 
from datetime import date, timedelta
import datetime, win32api
import print_job # Funzione verifica stampe in coda

from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import Encoders
from email.utils import formatdate
import smtplib

path_import = "."
ko = "zip_errore/"
ok = "zip_backup/"
path_pdf = "pdf_file/"
ERR_PASS = 'riep_errori.txt'    # File di riepilogo
FILE_LOG = 'LOG.txt'            # Lista dei file elaborati
#DESTINATARIO = ''

# Calcola orario corrente e lo formatta
def Time_Stamp():
    global timestamp
    timestamp = time.strftime("%d-%m-%Y %H:%M:%S")
    return (timestamp)

#Cancella il file lista dei file elaborati se presente
if os.path.isfile(FILE_LOG):
    os.remove(FILE_LOG)
    open(FILE_LOG, 'w').close()

log = open(FILE_LOG, 'a')
Time_Stamp()
log.write ('%s - Start programma\n' %timestamp)

#Cancella il file di riepilogo esistente se presente
if os.path.isfile(ERR_PASS):
    os.remove(ERR_PASS)
    open(ERR_PASS, 'w').close()

#Sposta files pdf rimasti sulla directory
def sposta_pdf():
    dirs = os.listdir( path_import )
    for file_presenti in dirs:
        if file_presenti.endswith(".pdf" or ".PDF"):
            shutil.move(file_presenti, path_pdf+file_presenti)

#Funzione invio email del file
def sendmail():
    print '\n\nInvio e-mail\n'
    #Parametri server
    smtp_server = "xxxxxxxxxxxxxxxxxxx"
    smtp_login = "cexxxxxxxxxxxxxxxxxxxxxxx"
    smtp_password = "xxxxxxxxxxxxxxxxxx"
    smtp_account = "xxxxxxxxxx"
    smtp_to = "walter.palumbo@xxxxxxxxxxxxxxxxxx.it"
    smtp_from = "xxxxxxxxxxxxxxx"
    server = smtplib.SMTP(smtp_server)
    server.login(smtp_login,smtp_password)

    #E-mail
    attach=(FILE_LOG)
    msg = MIMEMultipart()
    msg.attach(MIMEText(file(FILE_LOG).read()))
    msg["From"] = smtp_from
    msg["To"] = smtp_to
    msg["Subject"] = 'Report Stampe FATT/NC'
    msg['Date'] = formatdate(localtime=True)

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(attach, 'rb').read())
    Encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="%s"' % os.path.basename(attach))
    msg.attach(part)

    #Invio e-mail
    server.sendmail(smtp_from, smtp_to, msg.as_string())
    server.close()
    print ('Fine invio, inviata al seguente indirizzo: %s \n\n\n' %(smtp_to))
#Fine funzione invio email
    

def elabora_file_TXT(FILE_TXT): # Estrae da txt i dati
    global PIVA
    global DATADOC
    global NUMDOC
    global DENOMINAZIONE
    global TIPODOC
    global FILE_XML
    global SDI_ID
    global DESTINATARIO
    
    filename = open(FILE_TXT, 'r')
    print ('Lettura del file TXT: '+FILE_TXT)
    righe = filename.readlines()
    for riga in righe:
        if len(riga) >= 1:
            contenuto = (riga.split('#'))
            indice = contenuto[1]
            Indice = "".join(c for c in indice if c not in ('\n',']','[','\'','+'))
            if (Indice == '001'):
                Cliente = contenuto[2]
                Nome_file = contenuto[3]
                Mimetype = contenuto[4]
                Id_univoco = contenuto[5]
                Numero_documento = contenuto[6]
                Data_documento = contenuto[7]
                Sezionale = contenuto[8]
                P_IVA_mittente = contenuto[9]
                CF_mittente = contenuto[10]
                Rag_soc_mitt = contenuto[11]
                P_IVA_destinat = contenuto[12]
                CF_destinat = contenuto[13]
                Rag_soc_dest = contenuto[14]
                Sequenza = contenuto[15]
                Codice_ufficio = contenuto[16]
                Codice_canale_inv = contenuto[17]
                Identificativo_sdi = contenuto[18]
                Campo_tecnico = contenuto[19]
                Campo_tecnico = contenuto[20]
                Data_ricez_doc = contenuto[21]
                Ora_ricez_doc = contenuto[22]
                PIVA = P_IVA_mittente
                DATADOC = Data_documento
                NUMDOC = "".join(c for c in Numero_documento if c not in ('\n',']','[','\'','+',' ','/')) # Num. docum.
                DENOMINAZIONE = "".join(c for c in Rag_soc_mitt if c not in ('\n',']','[','\'','+','/',"'",'(',')')) # Denominazione
                FILE_XML = "".join(c for c in Nome_file if c not in ('\n',']','[','\'','+',' ','/','(',')'))
                TIPODOC = "".join(c for c in Id_univoco if c not in ('\n',']','[','\'','+',' ','/'))
                SDI_ID = Identificativo_sdi
                DESTINATARIO = "".join(c for c in Rag_soc_dest if c not in ('\n',']','[','\'','+','/',"'",'(',')')) # Rag_soc_dest
                filename.close()
                da_cancellare.append(FILE_TXT) # Aggiunge da cancella il file TXT elaborato
                return (PIVA, DENOMINAZIONE, NUMDOC, TIPODOC, DATADOC, FILE_XML, SDI_ID,DESTINATARIO)
        

def insert_db(PIVA, DENOMINAZIONE, TIPODOC, NUMDOC, DATADOC, FILE_PDF1, FILE_PDF2, FILE_XML, STAMPATO, SDI_ID, DESTINATARIO): # Insert su DB
    DATREG = time.strftime("%Y-%m-%d")
    #Time_Stamp()
    #log.write('%s - Insert su DB \n' % timestamp)
    FILE_PDF1 = FILE_PDF1.replace('$','')
    FILE_PDF2 = FILE_PDF2.replace('$','')
    DATA = str(DATADOC)
    DATA = ("%s-%s-%s" % (DATADOC[0:4], DATADOC[4:6], DATADOC[6:8]))
    DATADOC = DATA
    # Apre connessione al DB
    con = pyodbc.connect(driver="{SQL Server}",server='SVFSV001',database='FatturaElettronica',uid='sa',pwd='xxxxxxxxxxxxxx')
    cur = con.cursor()
    # Crea la query
    db_cmd = (("SELECT COUNT(*) from dbo.Fatture_importate where PIVA = '%s' and NUMDOC = '%s' and DATADOC = '%s' COLLATE SQL_Latin1_General_CP1_CS_AS") % (PIVA, NUMDOC, DATADOC))
    cur.execute(db_cmd)
    result = cur.fetchone()[0]
    if (result > 0):
        db_cmd = (("DELETE from dbo.Fatture_importate where PIVA = '%s' and NUMDOC = '%s' and DATADOC = '%s' COLLATE SQL_Latin1_General_CP1_CS_AS") % (PIVA, NUMDOC, DATADOC))
        cur.execute(db_cmd)

    db_cmd = ("INSERT INTO dbo.Fatture_importate ([PIVA],[NUMDOC],[DATADOC],[DENOMINAZIONE],[TIPODOC],[FILE_PDF1],[DATREG],[FILE_PDF2],[FILE_XML],[STAMPATO],[SDI_ID],[DESTINATARIO])\
              VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PIVA,NUMDOC,DATADOC,DENOMINAZIONE,TIPODOC,FILE_PDF1,DATREG,FILE_PDF2, FILE_XML, STAMPATO, SDI_ID, DESTINATARIO))
    cur.execute(db_cmd)
    cur.commit()
    con.close()  

def stampa_pdf(FILE_PDF): # Stampa pdf trovati
    log.write('%s - Stampa n. %s pdf file %s\n' % (timestamp,n_file,FILE_PDF))
    pdf_file_name = FILE_PDF
    win32api.ShellExecute(0, "print", pdf_file_name, None, ".", 0)
    time.sleep(3)

def cercazip(): # Cerca i file zippati
    '''
    Funzione principale che lista i file ZIP
    '''
    Time_Stamp()
    log.write('%s - Verifica presenza zip files\n' % timestamp)
    if os.listdir( path_import ) == []:
        print 'Directory vuota, nessun file da elaborare.\n'
    else:
        elabora_zip()
    
def elabora_zip(): # Inizio elaborazione ZIP
    print 'Inizio elaborazione'
    zip_elab_ok = []    # Lista dei file zip elaborati con successo
    pdf_elab_ok = []    # Lista dei file pdf elaborati con successo
    lista_zip = []      # Lista dei file zip contenuti nella directory
    global da_cancellare
    da_cancellare = []  # Lista dei file da cencallare
    FILE_PDF1 = ''
    FILE_PDF2 = ''
    FILE_TXT = ''
    dirs = os.listdir( path_import )
    global n_file
    n_file = 0
    for filezippato in dirs:
        if filezippato.endswith(".zip" or ".ZIP"):
            n_file = n_file + 1
            inizio_elab = 0                                 
            print ('- Elaborazione n. %s ZIP in corso: %s' % (n_file,filezippato))
            Time_Stamp()
            log.write('%s - Elaborazione n. %s zip file %s\n' % (timestamp,n_file,filezippato))
            lista_zip.append(filezippato)
            zz = zipfile.ZipFile(filezippato)
            lista = zz.namelist()
            pdf_trovato = 'NO'
            txt_trovato = 'NO'
            STAMPATO = 'NO'
            conta_pdf = 0
            conta_txt = 0
            for f_cont in lista:
                if not (f_cont.endswith(".txt") or f_cont.endswith(".pdf") or f_cont.endswith(".py")):
                    da_cancellare.append(f_cont)
                if txt_trovato == 'NO' and conta_txt == 0:                      # Verifica esistenza TXT dentro lo ZIP
                    for f_cont in lista:
                        if f_cont.endswith(".txt"): 
                            if f_cont[0:3] == 'imp':
                                FILE_TXT = f_cont
                                txt_trovato = 'SI'
                                conta_txt =  1
                                zz.extract(FILE_TXT)
                                elabora_file_TXT(FILE_TXT)  # Estrae le info x archiviare la fattura da TXT
                                log.write('%s - Prima elaborazione TXT %s\n' % (timestamp,FILE_TXT))
                                if (PIVA == '9999999999999999') or (PIVA == '9999999999999') or PIVA == ('99999999999') or PIVA == ('9999999999999'):
                                    STAMPATO = 'SI'
                                    txt_trovato = 'NO'
                            else:
                                da_cancellare.append(FILE_TXT)                       
                    
                if f_cont.endswith(".pdf"):                 # Verifica esistenza PDF dentro il ZIP
                    if STAMPATO == 'NO' and conta_pdf == 0:
                        conta_pdf = conta_pdf + 1
                        FILE_PDF1 = f_cont
                        pdf_trovato = 'SI'
                        if FILE_PDF1 in pdf_elab_ok:    # Verifica se il pdf si trova nella lista di quelli stampati
                            print ('Il PDF %s - Fornitore %s - Doc. n. %s del %s risulta stampato\n' % (FILE_PDF,DENOMINAZIONE,NUMDOC,DATADOC))
                            log.write('%s - Il PDF %s - Fornitore %s - Doc. n. %s del %s - Destinatario %s risulta stampato\n' % (timestamp,FILE_PDF,DENOMINAZIONE,NUMDOC,DATADOC,DESTINATARIO))
                        else:
                            FILE_PDF = FILE_PDF1
                            zz.extract(FILE_PDF1)               # Estrae pdf
                            stampa_pdf(FILE_PDF)                # Stampa PDF
                            print ('Stampa PDF %s - Fornitore %s - Doc. n. %s del %s\n' % (FILE_PDF,DENOMINAZIONE,NUMDOC,DATADOC))
                            log.write('%s - Stampa PDF %s - Fornitore %s - Doc. n. %s del %s - Destinatario %s\n' % (timestamp,FILE_PDF,DENOMINAZIONE,NUMDOC,DATADOC,DESTINATARIO))
                            STAMPATO = 'SI'
                            pdf_elab_ok.append(FILE_PDF1)
                                        
                    if conta_pdf >= 1 and f_cont != FILE_PDF1:
                        FILE_PDF2 = f_cont
                        zz.extract(FILE_PDF2)               # Estrae secondo pdf
                        if FILE_PDF2 in pdf_elab_ok:   # Verifica se il secondo pdf si trova nella lista di quelli stampati
                            print ('Il PDF %s risulta stampato\n' % FILE_PDF)
                            log.write('%s - Il secondo PDF %s risulta stampato\n' % (timestamp,FILE_PDF))
                        else:
                            FILE_PDF = FILE_PDF2
                            print ('Stampa PDF %s\n' % FILE_PDF)
                            log.write('%s - Stampa secondo PDF %s\n' % (timestamp,FILE_PDF))
                            stampa_pdf(FILE_PDF)                # Stampa secondo PDF
                            pdf_elab_ok.append(FILE_PDF2)
 
                if STAMPATO == 'SI' and txt_trovato == 'SI':
                    insert_db(PIVA, DENOMINAZIONE, TIPODOC, NUMDOC, DATADOC, FILE_PDF1, FILE_PDF2, FILE_XML, STAMPATO,SDI_ID,DESTINATARIO) # INSERT su DB
                                          
            inizio_elab = 1
            zz.close()
            print ('Fine elaborazione n. %s ZIP: %s'% (n_file,filezippato))  # Fine elaborazione zip file
            Time_Stamp()
            log.write('%s - Fine elaborazione n. %s ZIP %s\n' % (timestamp,n_file,filezippato))
            log.write('___________________________________________________________________ \n')
            if inizio_elab == 1:
                zip_elab_ok.append(filezippato)             # Zip elaborati OK li aggiungo alla lista
                
    if zip_elab_ok >= 1:                                    # Operazioni post elaborazione ZIP
        print_job.print_job_checker()                       # Verifica se ci sono stampe in coda
        for i in zip_elab_ok:
            shutil.move(i, ok+i)                            # Sposto i file zip OK sulla directory zip_backup
        for i in da_cancellare:                             # Cancella file non necessari
            if os.path.isfile(i):
                os.remove(i)
        #for i in pdf_elab_ok:                               # Sposta file pdf elaborati
        #    shutil.move(i, path_pdf+i)
                        
    print 'Fine elab ZIP'
    Time_Stamp()
    log.write('%s - Fine elaborazione\n' % timestamp)
    log.close()
    #if DESTINATARIO == '' or DESTINATARIO == None: DESTINATARIO = 'Nessuna elaborazione compiuta'
    sendmail()                                  # Inoltro email di riepilogo giornaliero

sposta_pdf()
cercazip()
