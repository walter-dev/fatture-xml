# Importa i file presenti all'interno della directory da_importare sul DB - Ver. 1.0.1 del 28-05-19
# FatturaElettronica.Fatture_M

import pyodbc, shutil, time, os, re
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import Encoders
from email.utils import formatdate
import smtplib

path_import = "T:/da_importare"
#print path_import
stato_import = 0
NOCODE = 'NO'
filei = "fatture.txt"
FILE_RIEP = ('riep_efatt.txt') #File di riepilogo

#Cancella il file esistente se presente di riepilogo
if os.path.isfile(FILE_RIEP):
    os.remove(FILE_RIEP)
    open(FILE_RIEP, 'w').close()

#Funzione invio email del file
def sendmail(importate):
    print '\n\nInvio e-mail\n'
    #Parametri server
    smtp_server = "XXXXXXXX"
    smtp_login = "XXXXXXX"
    smtp_password = "XXXXXXX"
    smtp_account = "XXXXX"
    smtp_to = "walter.palumbo@XXXXXXXX.it"
    smtp_from = "XXXXXXX"
    server = smtplib.SMTP(smtp_server)
    server.login(smtp_login,smtp_password)

    #E-mail
    attach=(FILE_RIEP)
    msg = MIMEMultipart()
    msg.attach(MIMEText(file(FILE_RIEP).read()))
    msg["From"] = smtp_from
    msg["To"] = smtp_to
    msg["Subject"] = 'Riepilogo Import FATT/NC - Importati %s documenti' % importate
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

# Verifica CODFIS
def controllaCODFIS(CODFIS, CLIFOR, REGDOC, NUMDOC):
    codice_fiscale = CODFIS.strip()
    if ((codice_fiscale == '') or (codice_fiscale == 'null')):
        NOCODE = 'SI'
        print "Campo Vuoto!"
        ERRCODE = 8
        comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)    
    elif (codice_fiscale[0] == '9' or codice_fiscale[0] == '8'): # IdCodice per Onlus e Associazioni che non hanno PI commerciale
        PARTIV = codice_fiscale
        controllaPIVA(PARTIV,CLIFOR,REGDOC,NUMDOC)
    else:
        #print codice_fiscale, CLIFOR, REGDOC, NUMDOC
        CODICE_REGEXP = "^[0-9A-Z]{16}$"
        SETDISP = [1, 0, 5, 7, 9, 13, 15, 17, 19, 21, 2, 4, 18, 20,
            11, 3, 6, 8, 12, 14, 16, 10, 22, 25, 24, 23 ]
        ORD_ZERO = ord('0')
        ORD_A = ord('A')

        if(16 != len(codice_fiscale)):
            ERRCODE = 7
            print ("La lunghezza del codice fiscale errata contiene %s caratteri: %s\n" % (len(codice_fiscale), codice_fiscale))
        codice_fiscale = codice_fiscale.upper()
        match = re.match(CODICE_REGEXP, codice_fiscale)
        if not match:
            ERRCODE = 6
            print ("Il codice fiscale contiene dei caratteri non ammessi")
            comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
        s = 0
        for i in range(1,14,2):
            c = codice_fiscale[i]
            if c.isdigit():                
                s += ord(c) - ORD_ZERO
            else:
                s += ord(c) - ORD_A
        for i in range(0,15,2):
            c = codice_fiscale[i]
            if c.isdigit():
                c = ord(c) - ORD_ZERO
            else:   
                c = ord(c) - ORD_A
            s += SETDISP[c]
        if (s % 26 + ORD_A != ord(codice_fiscale[15])):
            print "Codice fiscale NON valido\n"
            ERRCODE = 5
            comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
    # print 'CODFIS OK'
    
# Verifica partita IVA.
def controllaPIVA(PARTIV,CLIFOR,REGDOC,NUMDOC):
    IVA_REGEXP = "^[0-9]{11}$"
    ORD_ZERO = ord('0')
    partita_iva = PARTIV.strip()
    #print partita_iva, CLIFOR, REGDOC, NUMDOC
    if len(partita_iva) < 0:
        print "Campo Vuoto!"
    if 11 != len(partita_iva):
        ERRCODE = 4
        print ("La lunghezza della partita IVA errata contiene %s caratteri: %s\n" % (len(partita_iva), partita_iva))
        comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
    match = re.match(IVA_REGEXP, partita_iva)
    if not match:
        ERRCODE = 3
        print ("La partita IVA contiene dei caratteri non ammessi")
        comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
    s = 0
    for i in range(0, 10, 2):
        s += ord(partita_iva[i]) - ORD_ZERO
    for i in range(1, 10, 2):
        c = 2 * (ord(partita_iva[i]) - ORD_ZERO)
        if c > 9:
            c -= 9
        s += c
    if (10 - s%10)%10 != ord(partita_iva[10]) - ORD_ZERO:
        print "Partita IVA NON valida\n"
        ERRCODE = 2
        comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
    #print 'P.IVA OK'
        
def comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE):
        #print CLIFOR, REGDOC, NUMDOC, ERRCODE
        #Inizio file riepilogo
        FILE_R = open(FILE_RIEP, 'a+')
        if ERRCODE == 1:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s RISULTA REGISTRATO SUL DB!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s RISULTA REGISTRATO SUL DB!!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 2:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s CONTIENE P.IVA ERRATA!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s CONTIENE P.IVA ERRATA!!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 3:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La partita IVA contiene dei caratteri non ammessi!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La partita IVA contiene dei caratteri non ammessi!!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 4:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La lunghezza della partita IVA errata!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La lunghezza della partita IVA errata!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 5:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s CONTIENE COD.FIS ERRATO!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s CONTIENE COD.FIS ERRATO!!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 6:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s Il cod. fisc. contiene dei caratteri non ammessi!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s Il cod. fisc. contiene dei caratteri non ammessi!!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 7:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La lunghezza del cod. fis. errata!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s La lunghezza del cod. fis. errata!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 8:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s Sprovvisto di P.I. e Cod. Fis.!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s Sprovvisto di P.I. e Cod. Fis.!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 9:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s contiene partita iva estera!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s contiene partita iva estera!!" % (CLIFOR, REGDOC, NUMDOC))
        if ERRCODE == 10:
                FILE_R.write("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s contiene cliente generico!!!" % (CLIFOR, REGDOC, NUMDOC))
                print ("ATTENZIONE!! Il documento intestato a: %s, numero: %s/%s contiene cliente generico!!" % (CLIFOR, REGDOC, NUMDOC))
        FILE_R.write('\n________________________________\n\n')
        FILE_R.close() #Chiusura file riepilogo

                
def sposta_file(stato_import, filei):
        DATA = time.strftime("%d%m%Y_%H%M%S")
        filen = ('fatt_%s.csv' % DATA)
        if stato_import == 0:
                sorg = ("%s/%s" % (path_import, filei))
                dest = ("%s/ko/%s" % (path_import, filen))
                shutil.move(sorg, dest)
        if stato_import == 1:
                sorg = ("%s/%s" % (path_import, filei))
                dest = ("%s/importate/%s" % (path_import, filen))
                shutil.move(sorg, dest)
        
def importa(filei):
        global stato_import
        stato_import = 1
        # Legge il file
        nfile = open((path_import+"/%s" % filei))
        lines = nfile.readlines()
        # Apre connessione al DB
        con = pyodbc.connect(driver="{SQL Server}",server='SVFSV001',database='FatturaElettronica',uid='sa',pwd='xxxxxxxxxxx')
        cur = con.cursor()
        global importate
        importate = 0
        partitaiva = 'SI'
        CODFIS = 'null'
        COFIESC = 'null'
        IMPUNI = '0.00'
        for line in lines: # Splitto il file
                line = line.replace("'", " ")
                if len(line) > 0:
                        record = line.split("|")
                        RAGSO1A = record[0]
                        RAGSO2A = record[1]
                        INDIRI2 = record[2]
                        CITTAC  = record[3]
                        CAPCLIC = record[4]
                        PROVINC = record[5]
                        PARTIVC = record[6]
                        CODFISC = record[7]
                        COFIESC = record[8]
                        CODNAZC = record[9]
                        FIMPEXC = record[10]
                        EMAILC  = record[11]
                        DATDOCC = record[12]
                        REGDOCC = record[13]
                        NUMDOCC = record[14]
                        CODPAGC = record[15]
                        CODVENC = record[16]
                        CFOTRAC = record[17]
                        CLIFORC = record[18]
                        ACCONTC = record[19]
                        CODARTC = record[20]
                        DESCR1C = record[21]
                        QTAC    = record[22]
                        IMPUNIC = record[23]
                        CODIVAC = record[24]
                        TIPDOCC = record[25]
                        CAUTRAC = record[26]
                        CODFINC = record[27]
                        TOTMERC = record[28]
                        IMPOR2C = record[29] 
                        ANNOC   = record[30]
                        NUMASSC = record[31]
                        RIGAC   = record[32]
                        CODALTC = record[33]
                        SCONTIC = record[34]
                        INETTOC = record[35] # Imponibile netto tutti sconti
                        IMCTP1C = record[36] # Importo totale su 966
                        NUMFAX = record[37] # COD identificativo Sdi
                        TOTIMPC = record[38] # Tot. Imponibile
                        PIVAAZIC = record[39] # PIVA AZI Mittente

                        # Verifica campi vuoti
                        RAGSO1 = "".join(c for c in RAGSO1A if c not in ('\n',']','[','\'','+','&'))
                        #print 'Rag. Soc.1: ',RAGSO1 #RAGSO1
                        if RAGSO2A == '': RAGSO2 = 'null'
                        else: RAGSO2 = "".join(c for c in RAGSO2A if c not in ('\n',']','[','\'','+','&'))
                        #print 'Tag. Soc2: ',RAGSO2 
                        if INDIRI2 == '': INDIRI = 'null'
                        else:
                            INDIRI3 = re.sub('[^A-Za-z0-9]+', ' ', INDIRI2)
                            INDIRI = "".join(c for c in INDIRI3 if c not in ('\n','\'','/'))
                        #print 'Indirizzo: ',INDIRI
                        if CITTAC == '': CITTA = 'null'
                        else: CITTA = "".join(c for c in CITTAC if c not in ('\n',']','[','\'','+'))
                        #print 'Citta: ',CITTA
                        if CAPCLIC == '': CAPCLI = "null"
                        else: CAPCLI = "".join(c for c in CAPCLIC if c not in ('\n',']','[','\'','+'))
                        #print 'CAP: ',CAPCLI
                        if PROVINC == '': PROVIN = 'null'
                        else: PROVIN = "".join(c for c in PROVINC if c not in ('\n',']','[','\'','+'))
                        #print 'Provincia: ',PROVIN
                        PARTIV = "".join(c for c in PARTIVC if c not in ('\n',''))        # P. IVA Cli.
                        if PARTIV == '' or PARTIV == 'null': partitaiva = 'NO' # Verifica se Cliente con P.I o solo COD.FIS.
                        else: partitaiva = 'SI'
                        if CODFISC == '' or CODFISC == 'null': CODFIS = 'null'
                        CODFIS = "".join(c for c in CODFISC if c not in ('\n',']','[','\'','+',''))        # Cod. Fis. Cli.
                        if COFIESC == '': COFIES = 'null'
                        else: COFIES = "".join(c for c in COFIESC if c not in ('\n',']','[','\'','+',''))
                        #print 'Cod. Fis. Estero: ',COFIES
                        if CODNAZC == '' or CODNAZC == 'null': CODNAZ = 'IT'
                        else: CODNAZ = "".join(c for c in CODNAZC if c not in ('\n',']','[','\'','+'))
                        #print 'Nazione: ',CODNAZ
                        if FIMPEXC == '': FIMPEX = 'I'
                        else: FIMPEX = "".join(c for c in FIMPEXC if c not in ('\n',']','[','\'','+'))
                        #print 'Mazionalita: ',FIMPEX
                        if EMAILC == '': EMAIL = 'null'
                        else: EMAIL = "".join(c for c in EMAILC if c not in ('\n',']','[','\'','+',''))
                        #print 'Email: ',EMAIL
                        #if DATDOCC == '': DATDOC = 'null'
                        DATDOC = "".join(c for c in DATDOCC if c not in ('\n',']','[','\'','+'))
                        #if REGDOCC == '': REGDOC = 'null'
                        REGDOC = "".join(c for c in REGDOCC if c not in ('\n',']','[','\'','+')) # Reg docum.
                        #if NUMDOCC == '': NUMDOC = 'null'
                        NUMDOC = "".join(c for c in NUMDOCC if c not in ('\n',']','[','\'','+')) # Num. docum.
                        if CODPAGC == '': CODPAG = 'null'
                        else: CODPAG = "".join(c for c in CODPAGC if c not in ('\n',']','[','\'','+'))
                        #print 'Codice Pagam.: ',CODPAG
                        if CODVENC == '': CODVEN = 'null'
                        else: CODVEN = "".join(c for c in CODVENC if c not in ('\n',']','[','\'','+'))
                        #print 'Cod. Conv: ',CODVEN
                        if CFOTRAC == '': CFOTRA = 'null'
                        else: CFOTRA = "".join(c for c in CFOTRAC if c not in ('\n',']','[','\'','+'))
                        #print 'Cod. Forn. TRrasp: ',CFOTRA
                        #if CLIFORC == '': CLIFOR = 'null'
                        CLIFOR = "".join(c for c in CLIFORC if c not in ('\n',']','[','\'','+')) # Codice cliente
                        if ACCONTC == '': ACCONT = 'null'
                        else: ACCONT = "".join(c for c in ACCONTC if c not in ('\n',']','[','\'','+'))
                        #print 'Acconto: ',ACCONT
                        CODART = "".join(c for c in CODARTC if c not in ('\n',']','[','\'','+','*','&'))
                        #print 'cod. ART: ',CODART
                        if DESCR1C == '' or DESCR1C == None: DESCR1 = '...'
                        DESCR1 = "".join(c for c in DESCR1C if c not in ('\n',']','[','\'','+','&','\''))
                        QTA = "".join(c for c in QTAC if c not in ('\n',']','[','\'','+'))
                        IMPUNI = IMPUNIC
                        #print len(IMPUNI)
                        if len(IMPUNI) < 1: IMPUNI = '0.00'
                        #print IMPUNI
                        #raw_input("Press Enter to continue...")
                        #IMPUNI = IMPUNIC.replace('\n','')
                        if CODIVAC == '' or CODIVAC == None: CODIVA = '22'
                        else:CODIVA = "".join(c for c in CODIVAC if c not in ('\n',']','[','\'','+'))
                        #print 'Codice IVA: ',CODIVA
                        if TIPDOCC == '': TIPDOC = 'null'
                        else: TIPDOC = "".join(c for c in TIPDOCC if c not in ('\n',']','[','\'','+'))
                        #print 'Tipo docum.: ',TIPDOC
                        if CAUTRAC == '': CAUTRA = 'null'
                        else: CAUTRA = "".join(c for c in CAUTRAC if c not in ('\n',']','[','\'','+'))
                        #print 'Causale T.: ',CAUTRA
                        if CODFINC == '': CODFIN = 'null'
                        else: CODFIN = "".join(c for c in CODFINC if c not in ('\n',']','[','\'','+'))
                        #print 'Codcie finanz.: ',CODFIN
                        if TOTMERC == '' or TOTMERC == None: TOTMER = '0.00'
                        else: TOTMER = "".join(c for c in TOTMERC if c not in ('\n',']','[','\'','+'))
                        #print 'Totale merce: ',TOTMER
                        if IMPOR2C == '' or IMPOR2C == None: IMPORTO2 = '0.00'
                        else: IMPORTO2 = "".join(c for c in IMPOR2C if c not in ('\n',']','[','\'','+'))
                        #print 'Importo2: ',IMPORTO2
                        if ANNOC == '': ANNO = 'null'
                        else: ANNO = "".join(c for c in ANNOC if c not in ('\n',']','[','\'','+'))
                        #print 'Anno: ',ANNO
                        #if NUMASSC == '': NUMASS = 'null'
                        NUMASS = "".join(c for c in NUMASSC if c not in ('\n',']','[','\'','+'))
                        #print 'Num Ass.: ',NUMASS
                        #if RIGAC == '': RIGA = 'null'
                        RIGA = "".join(c for c in RIGAC if c not in ('\n',']','[','\'','+'))
                        #print 'Riga: ',RIGA
                        CODALT = "".join(c for c in CODALTC if c not in ('\n',']','[','\'','+'))
                        #print 'EAN: ',CODALT
                        if SCONTIC == '' or SCONTIC == 'null': SCONTI = '0'
                        else: SCONTI = "".join(c for c in SCONTIC if c not in ('\n',']','[','\'','+'))
                        #print 'Sconti: ',SCONTI
                        if INETTOC == '' or INETTOC == None: INETTO = '0.00'
                        else: INETTO = "".join(c for c in INETTOC if c not in ('\n',']','[','\'','+'))
                        if IMCTP1C == '': IMCTP1 = 'null'
                        else: IMCTP1 = "".join(c for c in IMCTP1C if c not in ('\n',']','[','\'','+'))
                        NUMFAXC = "".join(c for c in NUMFAX if c not in ('\n',']','[','\'','+',''))
                        if NUMFAXC == '': CODIDES = '0000000'
                        if len(NUMFAXC) != 7: CODIDES = '0000000'
                        else: CODIDES = "".join(c for c in NUMFAX if c not in ('\n',']','[','\'','+',''))
                        if TOTIMPC == '' or TOTIMPC == None: TOTIMP = '0.00'
                        else: TOTIMP = "".join(c for c in TOTIMPC if c not in ('\n',']','[','\'','+'))
                        PIVAAZI = "".join(c for c in PIVAAZIC if c not in ('\n',']','[','\'','+'))
                        #print 'Netto: ',INETTO
                        # Esclude clienti generici
                        lista_generici = 'listagenerici.txt'
                        listagen = open(lista_generici, 'r')
                        for ccgenerico in listagen:
                            GENERICO = "".join(c for c in ccgenerico if c not in ('\n',''))
                            #print 'Confronto il mio CC: ', CLIFOR
                            #print 'con il CC: ', ccgenerico
                            #raw_input("Press Enter to continue...")
                            
                            if CLIFOR == GENERICO:
                                #print 'Confronto il mio CC: ', CLIFOR
                                #print 'con il CC: ', ccgenerico (LO USO ANCHE PER EVENTUALI ESCLUZIONI)
                                #print 'Trovato UGUALE'
                                ERRCODE = 10
                                comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
                                #raw_input("Press Enter to continue...")
                                CODFIS = ''
                                partitaiva = 'NO'
                                
                
                        #print CLIFOR, REGDOC, NUMDOC, NUMASS, RIGA, CITTA, DESCR1, RAGSO1, RAGSO2
                        if partitaiva == 'NO':
                            #if COFIESC != None or COFIESC != 'null': # Verifica se presente P.IVA estera
                            if ((COFIESC != None) or (COFIESC != 'null')) and ((CODFIS == None) or (CODFIS == 'null')): # Verifica se presente P.IVA estera
                                print ('P. IVA Estera: %s - Num. doc.: %s/%s ' %(COFIESC, NUMDOC, REGDOC))
                                ERRCODE = 9
                                comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
                                #raw_input("Press Enter to continue...")
                            else:
                                codice_fiscale = "".join(c for c in CODFIS if c not in ('\n','')) # Verifica lunghezza CODFIS
                                if(16 != len(codice_fiscale)):
                                    ERRCODE = 7
                                    comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
                                else:
                                    controllaCODFIS(CODFIS, CLIFOR, REGDOC, NUMDOC) # Verifica PIVA
                                    partitaiva = 'SI'
                        else:
                            controllaPIVA(PARTIV, CLIFOR, REGDOC, NUMDOC) # Verifica PIVA
                        # Crea la query
                        db_cmd = (("SELECT COUNT(*) from dbo.Fatture_Metropolis where [NUMASS] = '%s' and [RIGA] = '%s' and [NUMDOC] = '%s' and [ANNO] = '%s' and [REGDOC] = '%s' and [PIVAAZI] = '%s' COLLATE SQL_Latin1_General_CP1_CS_AS") % (NUMASS, RIGA, NUMDOC, ANNO, REGDOC, PIVAAZI))
                        cur.execute(db_cmd)
                        result = cur.fetchone()[0]
                        comunicata = 'NO'
                        if NOCODE == 'NO':
                            if (result > 0 and comunicata == 'NO'):
                                    comunicata = 'SI'
                                    ERRCODE = 1
                                    comunica_doc(CLIFOR, REGDOC, NUMDOC, ERRCODE)
                                    db_cmd = (("DELETE from dbo.Fatture_Metropolis where [NUMASS] = '%s' and [RIGA] = '%s' and [NUMDOC] = '%s' and [ANNO] = '%s' and [REGDOC] = '%s' COLLATE SQL_Latin1_General_CP1_CS_AS") % (NUMASS, RIGA, NUMDOC, ANNO, REGDOC))
                                    cur.execute(db_cmd)
                                    stato_import = 0
                            importate = importate + 1
                            print 'Importaz. n.: %s' % importate
                            db_cmd = ("INSERT INTO dbo.Fatture_Metropolis ([RAGSO1],[RAGSO2],[INDIRI],[CITTAC],[CAPCLI],[PROVIN],[PARTIV],[CODFIS],[COFIES],\
                            [CODNAZ],[FIMPEX],[EMAIL],[DATDOC],[REGDOC],[NUMDOC],[CODPAG],[CODVEN],[CFOTRA],[CLIFOR],[ACCONT],[CODART],[DESCR1],[QTA],\
                            [IMPUNI],[CODIVA],[TIPDOC],[CAUTRA],[CODFIN],[TOTMER],[IMPORTO2],[ANNO],[NUMASS],[RIGA],[CODALT],[SCONTI],[INETTO],[IMCTP1],[CODIDES],[TOTIMP],[PIVAAZI]) VALUES ('%s','%s','%s','%s','%s','%s','%s',\
                            '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (RAGSO1,\
                            RAGSO2,INDIRI,CITTA,CAPCLI,PROVIN,PARTIV,CODFIS,COFIES,CODNAZ,FIMPEX,EMAIL,DATDOC,REGDOC,NUMDOC,CODPAG,CODVEN,CFOTRA,CLIFOR,ACCONT,\
                            CODART,DESCR1,QTA,IMPUNI,CODIVA,TIPDOC,CAUTRA,CODFIN,TOTMER,IMPORTO2,ANNO,NUMASS,RIGA,CODALT,SCONTI,INETTO,IMCTP1,CODIDES,TOTIMP,PIVAAZI))
                            cur.execute(db_cmd)
                            cur.commit()
        con.close()        
        nfile.close()                        
                        
def lista_files():
        # Crea lista dei file da importare
        if os.listdir ( path_import ) == []:
                print 'Directory vuota, nessun file da importare.\n'
        else:
                global filei
                print 'Inizio import'
                dirs = os.listdir ( path_import )
                #n_files = len(dirs)
                for filei in dirs:
                        print ('- ')+filei
                        importa(filei)
                        time.sleep(1)
                        #print "Importati %s di %s files"
                print 'Fine Import'

# Verifica ed importa il file di xxxx
if os.path.isfile("%s/%s" %(path_import, filei)):
        importa(filei)
        sposta_file(stato_import, filei)
        sendmail(importate) # Comunica segnalazione
else: print 'File %s non presente, verificare << Export manuale Fatt./N.C. >>' % filei
