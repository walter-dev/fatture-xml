# Fattura Elettronica - Crea XML Ver. 1.3.4 del 25-06-19
# -*- coding: cp1252 -*-
# Seleziona dal DB e crea XML
#Struttura XML fattura elettronica Ver. 1.2.1
#(verso PA e privati in forma ordinaria)
import pyodbc, datetime, shutil, os, decimal
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import Encoders
from email.utils import formatdate
import smtplib

path_import = "T:/da_importare" # preleva da questa direMIory il file lista.txt che contiene i NUMASS da elaborare
path_inoltro = "T:/ediel"
errori = 'errori_fatture.txt'
FILE_RIEP = ('errori_fatture.txt') #File di riepilogo

#Cancella il file esistente se presente di riepilogo errori
if os.path.isfile(errori):
    os.remove(errori)
    open(errori, 'w').close()

#Funzione invio email del file
def sendmail():
    print '\n\nInvio e-mail\n'
    #Parametri server
    smtp_server = "mail.xxxxxx.it"
    smtp_login = "xxxx@xxxxx.it"
    smtp_password = "xxxxxxxx"
    smtp_account = "xxx"
    smtp_to = "walter.palumbo@xxxxxxx.it"
    smtp_from = "xxx@xxxxxxxxx.it"
    server = smtplib.SMTP(smtp_server)
    server.login(smtp_login,smtp_password)

    #E-mail
    attach=(FILE_RIEP)
    msg = MIMEMultipart()
    msg.attach(MIMEText(file(FILE_RIEP).read()))
    msg["From"] = smtp_from
    msg["To"] = smtp_to
    msg["SubjeMI"] = 'Riepilogo FATT/NC inoltrate'
    msg['Date'] = formatdate(localtime=True)

    part = MIMEBase('application', 'oMIet-stream')
    part.set_payload(open(attach, 'rb').read())
    Encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="%s"' % attach)
    msg.attach(part)

    #Invio e-mail
    server.sendmail(smtp_from, smtp_to, msg.as_string())
    server.close()
    print ('Fine invio, inviata al seguente indirizzo: %s \n\n\n' %(smtp_to))
#Fine funzione invio email
    
def sposta_file(PERCORSO, filexml_nopath):
    '''
    Funzione che sposta gli XML prodotti sulle direMIory di riferimento
    '''
    if PERCORSO == 'C00': dir_inoltro = 'fatt_xxxxxxxxx_attivo'        # xxxxxxxxx attivo
    if PERCORSO == 'CN0': dir_inoltro = 'fatt_xxxxxxxxx2_attivo'   # xxxxx attivo
    if PERCORSO == 'CM0': dir_inoltro = 'fatt_xxxxxxxxxx3_attivo'           # xxxx
    if PERCORSO == 'CG0': dir_inoltro = 'fatt_xxxxx4_attivo'      # xxx
    if PERCORSO == 'MI0': dir_inoltro = 'fatt_xxxxxxx5_attivo'            # xxxx

    sorgente = ("%s/%s" % (path_import, filexml_nopath))
    destinazione = ("%s/%s/%s" % (path_inoltro, dir_inoltro, filexml_nopath))
    print sorgente
    print destinazione
    shutil.move(sorgente, destinazione)    

def crea_lista():
    '''
    Funzione principale che elabora i NUMASS e genera gli XML
    '''
    errore_f = open(errori, 'a+')
    EsigibilitaIVA = 'I' # Di default - Immediata
    linea = 0
    FILEASS = ('%s/lista.txt' % path_import) #Lista NUMASS delle fatture da selezionare
    #if os.path.isfile(NOMEFILE): print 'File presente.'
    nfile = open(FILEASS, "r")
    lines = nfile.readlines()
    for line in lines:
        linea = linea + 1
        NUMASS1 = str(line)
        NUMASS2 = NUMASS1.strip()
        print ('Export n. %s - NUMASS: %s' % (linea, NUMASS2))
        # Apre connessione al DB
        DatiRitenuta = 'NO'
        BolloVirtuale = 'NO'
        con = pyodbc.conneMI(driver="{SQL Server}",server='SVFSV001',database='FatturaElettronica',uid='sa',pwd='xxxxxxxxx')
        cur = con.cursor()

        # Verifica presenza partita IVA e Codcie Fiscale o P. IVA Estera
        db_cmd = ("SELEMI TOP 1 [NUMASS],[REGDOC],[NUMDOC],[CLIFOR],[PARTIV],[CODFIS],[COFIES],[CODVEN] FROM dbo.Fatture_M where NUMASS = '%s'" % NUMASS2)
        res = cur.execute(db_cmd)

        for row in res:
            NUMASS=row.NUMASS # N. Assoluto
            REGDOC=row.REGDOC # Sulla base del REGDOC cambiano i dati AZI CedentePrestatore
            NUMDOC=row.NUMDOC # La somma di questo + il REGDOC compone il progressivo
            CLIFOR=row.CLIFOR # Codice cliente interno
            PARTIV=row.PARTIV # Partita IVA
            CODFIS=row.CODFIS
            COFIES=row.COFIES # Codice fiscale estero
            CODVEN=row.CODVEN # Codice convenzione x 17T (Split payement)
            CODICEC = CODVEN.strip()
            if CODICEC == '17T': EsigibilitaIVA = 'S' # Split payement
            
            # Escludo CC generici
            lista_generici = 'listagenerici.txt'
            listagen = open(lista_generici, 'r')
            for ccgenerico in listagen:
                GENERICO = "".join(c for c in ccgenerico if c not in ('\n',''))
                if CLIFOR == GENERICO:
                    errore_f.write('Il documento: %s - %s - %s Num. Assoluto: %s contiene CC generico\n' % (NUMDOC, REGDOC, CLIFOR, NUMASS))
                    NUMASS2 = '9999999999'

            if PARTIV == '' or PARTIV == 'null':
                if CODFIS == '' or CODFIS == 'null':
                    if COFIES == '' or COFIES == 'null':
                        print 'Cliente sprovvisto di P.IVA e Cod. Fis.'
                        errore_f.write('Il documento: %s - %s - %s Num. Assoluto: %s manca di P.IVA e Cod.Fis.\n' % (NUMDOC, REGDOC, CLIFOR, NUMASS))
                        NUMASS2 = '9999999999' # Cambio il valore di NUMASS per evitare di creare il file
        
        db_cmd = ("SELEMI COUNT(*) from dbo.Fatture_M where [NUMASS] = '%s' " % NUMASS2)
        cur.execute(db_cmd)
        result = cur.fetchone()[0]
        if (result == 0):
            print 'MUNASS NON TROVATO'
            errore_f.write('MUNASS NON TROVATO: %s\n' % NUMASS1)
            continue

        db_cmd = ("SELEMI COUNT(*) from dbo.Fatture_M where [NUMASS] = '%s' and CODART = 'BOLLI'" % NUMASS2) # Verifica bolli virtuali
        cur.execute(db_cmd)
        result = cur.fetchone()[0]
        if (result >= 1):
            BolloVirtuale = 'SI'
            ImportoBollo = '2.00'

        db_cmd = ("SELEMI COUNT(*) from dbo.Fatture_M where [NUMASS] = '%s' and [ACCONT] != '' and [ACCONT] != '0,00' and REGDOC = 'V01'" % NUMASS2) # Verifica ritenuta acconto
        cur.execute(db_cmd)
        result = cur.fetchone()[0]
        if (result >= 1):
            DatiRitenuta = 'SI'
            db_cmd = ("SELEMI TOP 1 [ACCONT] FROM dbo.Fatture_M where NUMASS = '%s'" % NUMASS2)
            res = cur.execute(db_cmd)
            for row in res:
                ACCONTO = row.ACCONT
            TipoRitenuta = 'RT02' # Persona giuridica
            ImportoRitenuta = ACCONTO.replace(',','.') # Importo della ritenuta
            AliquotaRitenuta = '23.00'
            CausalePagamento = 'Z' # Causale del pagamento: (Titolo diverso dai precedenti.)
            
        db_cmd = ("SELEMI TOP 1 [NUMASS],[REGDOC],[NUMDOC],[CODIVA],[RAGSO1],[RAGSO2],[INDIRI],[CAPCLI],[CITTAC],[PROVIN],[CODNAZ],[CLIFOR],[DATDOC],[TIPDOC],[CAUTRA],[PARTIV],[CODFIS],[CODVEN],[COFIES],[TOTMER],[IMMIP1],[TOTIMP],[CODIDES],[EMAIL],[PIVAAZI] FROM dbo.Fatture_M where NUMASS = '%s'" % NUMASS2)
        res = cur.execute(db_cmd)
        
        for row in res:
            NUMASS=row.NUMASS # N. Assoluto
            REGDOC=row.REGDOC # Sulla base del REGDOC cambiano i dati AZI CedentePrestatore
            NUMDOC=row.NUMDOC # La somma di questo + il REGDOC compone il progressivo
            CODIVA=row.CODIVA # Aliquota IVA
            RAGSO1=row.RAGSO1 # Denominazione CLI
            RAGSO2=row.RAGSO2 # Denominazione da sommare alla 1
            INDIRI=row.INDIRI # Indirizzo CLI
            CAPCLI=row.CAPCLI # CAP CLI
            CITTAC=row.CITTAC # Comune CLI
            PROVIN=row.PROVIN # Provincia CLI
            CODNAZ=row.CODNAZ # Nazione CLI
            CLIFOR=row.CLIFOR # Codice cliente interno
            DATDOC=row.DATDOC # Data emissione documento
            TIPDOC=row.TIPDOC # Tipo documento
            CAUTRA=row.CAUTRA # Causale di vendita
            PARTIV=row.PARTIV # Partita IVA
            CODFIS=row.CODFIS
            CODVEN=row.CODVEN # Convenzione
            COFIES=row.COFIES # Codice fiscale estero
            TOTMER=row.TOTMER # Totale Merce
            IMMIP1=row.IMMIP1 # Totale Merce per V01
            TOTIMP=row.TOTIMP # Totale Imponibile
            CODIDES=row.CODIDES # Codice Identificatico Destinatario
            EMAIL=row.EMAIL # E-mail destinatario
            PIVAAZI = row.PIVAAZI # P IVA AZIENDA MITTENTE

        # Verifica campi vuoti
        if REGDOC == '': print 'N. registro VUOTO'
        if NUMDOC == '': print 'N. docum. VUOTO'
        if (RAGSO2 == '' or RAGSO2 == 'null'): RAGSO2 = ''
        partitaiva = 'SI'
        if PARTIV == '' or PARTIV == 'null': partitaiva = 'NO' # Verifica se Cliente con P.I o solo COD.FIS.
        if len(PARTIV) > 0:
            if PARTIV[0] == '9': partitaiva = 'NO' # Per le Associazioni e le Onlus la P.I. e CODFIS, ma va messo su CODFIS
            if PARTIV[0] == '8': partitaiva = 'NO' # Per le Associazioni e le Onlus la P.I. e CODFIS, ma va messo su CODFIS
        COFIES = str(COFIES)
            # Codice destinatario
        if CODIDES == '' or CODIDES == 'null':
            CodiceDestinatario = "0000000" # Generico IT
        else:
            CodiceDestinatario = str(CODIDES)
        PECDestinatario = 'NO'
        if EMAIL == 'null' or len(EMAIL) < 5:
            PECDestinatario = 'NO'
        else:
            PECDestinatario = str(EMAIL)
        if (COFIES != 'null'):
            print 'Codice fiscale Estero'
            CodiceDestinatario = "XXXXXXX" # Generico estero
            partitaiva = 'ES'
            PECDestinatario = 'NO'
                        
        con.close() # Chiusura DB
        
        # Valori Fissi
            # Dati Anagrafici Trasmittente (TESI)
        IdPaese = "IT"
        IdCodice = "999999999" # P. IVA Tesi
            # Dati trasmissione
        ProgressivoInvio = str(NUMDOC+REGDOC)
        FormatoTrasmissione = "FPR12"
        
            # Dati Anagrafici cedente/prestatore
        IdPaeseCP = "IT"
        RegimeFiscale = "RF01" # Ordinario

        # Dati cedente/prestatore (xxxxxxxxx)
        if (REGDOC == "CN0" or REGDOC == "RN0") or (REGDOC == 'V01' and PIVAAZI == '9999999999'): # xxxxxxx Srl
            IdCodiceCP =  "999999999999"
            DenominazioneCP = "xxxxxxxx S.R.L. "
            IndirizzoCP = "xxxxx xxxx xxxx A3"
            CAPCP = "xxxx"
            ComuneCP = "xxxx"
            ProvinciaCP = "MI"
            NazioneCP = "IT"
            PERCORSO = 'CN0'
        elif (REGDOC == "CM0" or REGDOC == "RM0") or (REGDOC == 'V01' and PIVAAZI == '9999999999'): # xxxxxx Srl
            IdCodiceCP =  "999999999999"
            DenominazioneCP = "xxxxxxxxxxxxx"
            IndirizzoCP = "xxxxxxxxxxxx"
            CAPCP = "99999"
            ComuneCP = "MILANO"
            ProvinciaCP = "MI"
            NazioneCP = "IT"
            PERCORSO = 'CM0'
        elif (REGDOC == "CG0" or REGDOC == "RG0") or (REGDOC == "CC0" or REGDOC == "RC0") or (REGDOC == 'V01' and PIVAAZI == '9999999999'): # xxxxxxxxx Srl
            IdCodiceCP =  "999999999999"
            DenominazioneCP = "xxxxxxxxx"
            IndirizzoCP = "xxxxxxxxxx"
            CAPCP = "99999"
            ComuneCP = "MILANO"
            ProvinciaCP = "MI"
            NazioneCP = "IT"
            PERCORSO = 'CG0'
        elif (REGDOC == "MI0" or REGDOC == "RT0") or (REGDOC == 'V01' and PIVAAZI == '99999999999'): # XXXX Srl
            IdCodiceCP =  "999999999"
            DenominazioneCP = "xxxxxxxxxxx"
            IndirizzoCP = "xxxxxxxxxxxxxxxxxxxx"
            CAPCP = "99999"
            ComuneCP = "MILANO"
            ProvinciaCP = "MI"
            NazioneCP = "IT"
            PERCORSO = 'MI0'
        else: # xxxxxxxxx xxxxx Spa
            IdCodiceCP =  "99999999"
            DenominazioneCP = "xxxxxxxxx"
            IndirizzoCP = "xxxxxxxxxxxxxxxx"
            CAPCP = "99999"
            ComuneCP = "MILANO"
            ProvinciaCP = "MI"
            NazioneCP = "IT"
            PERCORSO = 'C00'

        # Dati anagrafici cessionario/committente (CLIENTE/DEST.)
        if partitaiva == 'SI': CodiceFiscale = str(PARTIV)
        elif partitaiva == 'ES': CodiceFiscale = str(COFIES)
        else: CodiceFiscale = str(CODFIS)
        Denominazione = ('%s %s' %(RAGSO1,RAGSO2))
        Indirizzo = str(INDIRI)
        CAP = str(CAPCLI)
        Comune = str(CITTAC)
        Provincia = str(PROVIN)
        Nazione = str(CODNAZ)

        # Dati da inserire sul Body
        CODCLI = str(CLIFOR)
        TipoDocumento = str(TIPDOC)
        #print TipoDocumento
        if (TipoDocumento == "CO") or (TipoDocumento == "FD") or (TipoDocumento == "FI") or (TipoDocumento == "FO"): # Fattura e rifatturazione - fattura affiliato
            TipoDocumento = "TD01"
        if (TipoDocumento == "BC") or (TipoDocumento == "AC") or (TipoDocumento == "MS") or (TipoDocumento == "NA"): # Nota credito
            TipoDocumento = "TD04"
        if (TipoDocumento == "FI" and CODCLI == "CC9999999"):
            TipoDocumento = "TD20" # Autofattura
        # Valori fissi
        Divisa = "EUR"
        DATA = str(DATDOC)
        Data = ("%s-%s-%s" % (DATA[0:4], DATA[4:6], DATA[6:8]))
        REGISTRO = str(REGDOC)
        if (REGISTRO == 'V01') or (REGISTRO == 'NAC') or (REGISTRO == 'V04') or (REGISTRO == 'R04'):
            IMMIP = IMMIP1.replace(',','.')
            ImportoTotaleDocumento = "{:.2f}".format(float(IMMIP)) # Import totale documento
        else:
            TOTMER1 = TOTMER.replace(',','.')
            ImportoTotaleDocumento = "{:.2f}".format(float(TOTMER1)) # Import totale documento
        
        # Verifica se si tratta di fattura PDV o Sede
        VPREZZO = 'SIIVA' # I prezzi di default sono tutti ivati, quindi da scorporare
        if (REGISTRO == 'V01') or (REGISTRO == 'NAC') or (REGISTRO == 'V04') or (REGISTRO == 'R04'): VPREZZO = 'NOIVA' # Solo per la sede i prezzi sono al netto
        Numero = str(NUMDOC+REGDOC)
        Causale = str(CAUTRA)
        if CAUTRA == "VE":
            Causale = "Fattura di vendita"
        if CAUTRA == "RE":
            Causale = "Nota di credito a cliente"
        # Crea file XML - NOME DEL FILE SINTASSI (codice paese | identificativo univoco del soggetto trasmittente | progressivo univoco del file)
        filexml = ('%s/%s%s%s.xml' % (path_import, NazioneCP, IdCodiceCP, Numero))
        filexml_nopath = ('%s%s%s.xml' % (NazioneCP, IdCodiceCP, Numero))
        # Scrittura Header
        NOME_FILE = filexml # Scrivo su un file temporaneo da rinomiare a fine lavoro
        FILE = open(NOME_FILE, 'a+')
        FILE.write('<?xml version="1.0" encoding="UTF-8"?>')
        FILE.write('\r<p:FatturaElettronica versione="FPR12" xmlns:ds="http://www.w3.org/2000/09/xmldsig#" xmlns:p="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2 http://www.fatturapa.gov.it/export/fatturazione/sdi/fatturapa/v1.2/Schema_del_file_xml_FatturaPA_versione_1.2.xsd">')
        FILE.write('\r<FatturaElettronicaHeader>')
        FILE.write('\r<DatiTrasmissione>')
        FILE.write('\r<IdTrasmittente>')
        FILE.write('\r<IdPaese>%s</IdPaese>' % IdPaese)
        FILE.write('\r<IdCodice>%s</IdCodice>' % IdCodice)
        FILE.write('\r</IdTrasmittente>')
        FILE.write('\r<ProgressivoInvio>%s</ProgressivoInvio>' % ProgressivoInvio)
        FILE.write('\r<FormatoTrasmissione>%s</FormatoTrasmissione>' % FormatoTrasmissione)
        FILE.write('\r<CodiceDestinatario>%s</CodiceDestinatario>' % CodiceDestinatario)
        if PECDestinatario != 'NO':
            FILE.write('\r<PECDestinatario>%s</PECDestinatario>' % PECDestinatario)
        FILE.write('\r</DatiTrasmissione>')
        #--- CEDENTE / PRESTATORE
        FILE.write('\r<CedentePrestatore>')
        FILE.write('\r<DatiAnagrafici>')
        FILE.write('\r<IdFiscaleIVA>')
        FILE.write('\r<IdPaese>%s</IdPaese>' % IdPaeseCP)
        FILE.write('\r<IdCodice>%s</IdCodice>' % IdCodiceCP)
        FILE.write('\r</IdFiscaleIVA>')
        FILE.write('\r<Anagrafica>')
        FILE.write('\r<Denominazione>%s</Denominazione>' % DenominazioneCP)
        FILE.write('\r</Anagrafica>')
        FILE.write('\r<RegimeFiscale>%s</RegimeFiscale>' % RegimeFiscale)
        FILE.write('\r</DatiAnagrafici>')
        FILE.write('\r<Sede>')
        FILE.write('\r<Indirizzo>%s</Indirizzo>' % IndirizzoCP)
        FILE.write('\r<CAP>%s</CAP>' % CAPCP)
        FILE.write('\r<Comune>%s</Comune>' % ComuneCP)
        FILE.write('\r<Provincia>%s</Provincia>' % ProvinciaCP)
        FILE.write('\r<Nazione>%s</Nazione>' % NazioneCP)
        FILE.write('\r</Sede>')
        FILE.write('\r</CedentePrestatore>')
        #--- CESSIONARIO / COMMITTENTE (Dati del Cliente/Destinatario)
        FILE.write('\r<CessionarioCommittente>')
        FILE.write('\r<DatiAnagrafici>')
        if partitaiva == 'SI' or partitaiva == 'ES':
            FILE.write('\r<IdFiscaleIVA>')
            FILE.write('\r<IdPaese>%s</IdPaese>' % Nazione)
            FILE.write('\r<IdCodice>%s</IdCodice>' % CodiceFiscale) # Partita IVA Cli.
            FILE.write('\r</IdFiscaleIVA>')
        else:
            FILE.write('\r<CodiceFiscale>%s</CodiceFiscale>' % CodiceFiscale)
        FILE.write('\r<Anagrafica>')
        FILE.write('\r<Denominazione>%s</Denominazione>' % Denominazione)
        FILE.write('\r</Anagrafica>')
        FILE.write('\r</DatiAnagrafici>')
        FILE.write('\r<Sede>')
        FILE.write('\r<Indirizzo>%s</Indirizzo>' % Indirizzo)
        FILE.write('\r<CAP>%s</CAP>' % CAP)
        FILE.write('\r<Comune>%s</Comune>' % Comune)
        FILE.write('\r<Provincia>%s</Provincia>' % Provincia)
        FILE.write('\r<Nazione>%s</Nazione>' % Nazione)
        FILE.write('\r</Sede>')
        FILE.write('\r</CessionarioCommittente>')
        FILE.write('\r</FatturaElettronicaHeader>')
        # FINE HEADER
        # Inizio Body
        FILE.write('\r<FatturaElettronicaBody>')
        FILE.write('\r<DatiGenerali>')
        FILE.write('\r<DatiGeneraliDocumento>')
        FILE.write('\r<TipoDocumento>%s</TipoDocumento>' % TipoDocumento)
        FILE.write('\r<Divisa>%s</Divisa>' % Divisa)
        FILE.write('\r<Data>%s</Data>' % Data)
        FILE.write('\r<Numero>%s</Numero>' % Numero)
        if DatiRitenuta == 'SI': # Dati Ritenuta acconto
            FILE.write('\r<DatiRitenuta>')
            FILE.write('\r<TipoRitenuta>%s</TipoRitenuta>' % TipoRitenuta)
            FILE.write('\r<ImportoRitenuta>%s</ImportoRitenuta>' % ImportoRitenuta)
            FILE.write('\r<AliquotaRitenuta>%s</AliquotaRitenuta>' % AliquotaRitenuta)
            FILE.write('\r<CausalePagamento>%s</CausalePagamento>' % CausalePagamento)
            FILE.write('\r</DatiRitenuta>')
            
        if BolloVirtuale == 'SI': # DatiBollo
            FILE.write('\r<DatiBollo>')
            FILE.write('\r<BolloVirtuale>%s</BolloVirtuale>' % BolloVirtuale)
            FILE.write('\r<ImportoBollo>%s</ImportoBollo>' % ImportoBollo)
            FILE.write('\r</DatiBollo>')
        FILE.write('\r<ImportoTotaleDocumento>%s</ImportoTotaleDocumento>' % ImportoTotaleDocumento)
        FILE.write('\r<Causale>%s</Causale>' % Causale)
        FILE.write('\r</DatiGeneraliDocumento>')
        FILE.write('\r</DatiGenerali>')
        FILE.write('\r<DatiBeniServizi>')       
        FILE.close()
        
        # Inizio del Dettaglio Linee
        # Connessione al DB
        con = pyodbc.conneMI(driver="{SQL Server}",server='SVFSV001',database='FatturaElettronica',uid='sa',pwd='xxxxxxxxx')
        cur = con.cursor()
        db_cmd = ("SELEMI TOP (100) PERCENT [CODALT], [DESCR1], [QTA], [IMPUNI], [TOTMER], [CODIVA], [SCONTI], [INETTO] FROM dbo.Fatture_M where NUMASS = '%s'" % NUMASS2)
        res = cur.execute(db_cmd)
        NumeroLinea = 0
        
        # Estraggo dati restituiti dalla seleMI
        listaImporti = [] # lista importi da sommare per riepilogo IVA 22
        listaImporti04 = [] # lista importi da sommare per riepilogo IVA 4
        listaImporti00 = [] # lista importi da sommare per riepilogo IVA Esente
        listaImporti01 = [] # lista importi da sommare per riepilogo IVA Reverse X17
        listaImporti05 = [] # lista importi da sommare per riepilogo IVA 10
        
        for row in res.fetchall():
            CODALT = row.CODALT     # Cod. EAN dell'articolo
            DESCR1 = row.DESCR1     # Descrizione
            QTA = row.QTA           # Quantita
            IMPUNI = row.IMPUNI     # Importo unitario
            TOTMER = row.TOTMER     # Prezzo totale
            CODIVA = row.CODIVA     # Aliquota IVA
            SCONTI = row.SCONTI     # Sconti automatici VALORE %
            INETTOT = row.INETTO    # Importo totale a netto degli sconti
            
            IMPUNI = IMPUNI.replace(',','.')
            if IMPUNI == '0': IMPUNI = '0.00'
            QTA = QTA.replace(',','.')
             
            NumeroLinea = NumeroLinea + 1 # N. progressivo linea

            CodiceTipo = 'EAN'
            CodiceValore = str(CODALT)
            Descrizione = str(DESCR1)
            
            if (QTA == ('null' or '')): QTA = '0.00'

            Quantita = "{:.2f}".format(float(QTA))
 
            if Quantita < '0.00':
 
                Quantita =  "{:.2f}".format(float(Quantita) * -1)
                QTA_Negativa = 'SI'
            else:
                QTA_Negativa = 'NO'

            if ((INETTOT == '') or (INETTOT == None) or (INETTOT == 'null') or (IMPUNI == '0.00')):
                INETTO = IMPUNI
            else:
                INETTOT2 = INETTOT.replace(',','.')
                INETTO = (float(INETTOT2)/ float(Quantita)) # Importo netto unitario
            if CODVEN == 'IV4':
                INETTO = IMPUNI
            if ((TOTMER == '') or (TOTMER == None) or (TOTMER == 'null')):
                TOTMER = '0.00'
            else:
                TOTMER = TOTMER.replace(',','.')

            AliquotaIVA = str(CODIVA)
            AliquotaIVA = AliquotaIVA.strip()
            CODVEN = str(CODVEN)
            CODVEN = CODVEN.strip()

             # Calcolo del prezzo imponibile (al netto di IVA)
            if VPREZZO == 'SIIVA': # Per le vendite fatte da PDV
                if AliquotaIVA == '22': INETTO = float(INETTO)/1.22 # Scorporo IVA 22
                if (AliquotaIVA == '04' and CODVEN != 'IV4'): INETTO = float(INETTO)/1.04 # Scorporo IVA 04
                if (AliquotaIVA == '4' and CODVEN != 'IV4'): INETTO = float(INETTO)/1.04 # Scorporo IVA 04
                if (AliquotaIVA == '04' and CODVEN == 'IV4'):
                    INETTO = float(INETTO)/1.22 # Scorporo IVA 22 SU CODVEN 'IV4'
                if AliquotaIVA == '10': INETTO = float(INETTO)/1.10 # Scorporo IVA 10
                if AliquotaIVA == '05': INETTO = float(INETTO)/1.05 # Scorporo IVA 05

            PrezzoUnitario = decimal.Decimal(INETTO, 2)  # Valorizzazione del PrezzoUnitario
            
            '''

            if CODALT == '0000000000000':
                print float(Quantita)
                print float(INETTO)
                print float(INETTO)/float(Quantita)
                print float(PrezzoUnitario) # Importo netto unitario
                print "{:.2f}".format(float(PrezzoUnitario))
                print decimal.Decimal(PrezzoUnitario, 2)
                raw_input("\nPress Enter to continue...")
            '''

            # Trascodifica aliquote
            if (AliquotaIVA == "22") or (AliquotaIVA == "22T"):
                ALIQUOTA = "22.00"
                AliquotaIVA1 = "22.00"
                PrezzoTotale1 = (float(PrezzoUnitario) * float(Quantita))
                PrezzoTotale = "{:.2f}".format(float(PrezzoTotale1))
                listaImporti.append(PrezzoTotale1)

            if (AliquotaIVA == "04" or AliquotaIVA == "4"):
                AliquotaIVA2 = "04.00"
                ALIQUOTA = "04.00"
                PrezzoTotale1 = (float(PrezzoUnitario) * float(Quantita))
                PrezzoTotale = "{:.2f}".format(float(PrezzoTotale1))
                listaImporti04.append(PrezzoTotale1)
                
            if (AliquotaIVA == "10"):
                AliquotaIVA2 = "10.00"
                ALIQUOTA = "10.00"
                PrezzoTotale1 = (float(PrezzoUnitario) * float(Quantita))
                PrezzoTotale = "{:.2f}".format(float(PrezzoTotale1))
                listaImporti05.append(PrezzoTotale1)
                
            if (AliquotaIVA == "X17"):
                AliquotaIVA4 = "00.00"
                ALIQUOTA = "00.00"
                PrezzoTotale1 = (float(PrezzoUnitario) * float(Quantita))
                PrezzoTotale = "{:.2f}".format(float(PrezzoTotale1))
                listaImporti01.append(PrezzoTotale1)
                # Valori ammessi campo Natura
                if AliquotaIVA == "X17": Natura = 'N6' # Reverse charge

            if ((AliquotaIVA == "E74") or (AliquotaIVA == "E02") or (AliquotaIVA == "FCI") or (AliquotaIVA == "E26") or (AliquotaIVA == "N8B") or (AliquotaIVA == "X8") or (AliquotaIVA == "N08") or (AliquotaIVA == "N8C") or (AliquotaIVA == "N07") or (AliquotaIVA == "E15") or (AliquotaIVA == "P10") or (AliquotaIVA == "E10")):
                AliquotaIVA3 = "00.00"
                ALIQUOTA = "00.00"
                PrezzoTotale1 = (float(PrezzoUnitario) * float(Quantita))
                PrezzoTotale = "{:.2f}".format(float(PrezzoTotale1))
                listaImporti00.append(PrezzoTotale1)
                # Valori ammessi campo Natura
                if AliquotaIVA == "E74": Natura = 'N4' # Esenzione
                if AliquotaIVA == "E02": Natura = 'N4' # Esenzione
                if AliquotaIVA == "E10": Natura = 'N4' # Esenzione
                if AliquotaIVA == "N8B": Natura = 'N3' # N.IMP. ART.8/C
                if AliquotaIVA == "N8C": Natura = 'N3' # N.IMP. ART.8/C
                if AliquotaIVA == "N08": Natura = 'N3' # NON IMP.ART.8 1 COMMA L.A/B
                if AliquotaIVA == "N07": Natura = 'N2' # NON SOGGETTO A IVA ART.7 TER
                if AliquotaIVA == "FCI": Natura = 'N2' # NON SOGGETTO A IVA ART.7 TER - FUORI CAMPO IVA
                if AliquotaIVA == "E15": Natura = 'N1' # ESCLUSO ART. 15
                if AliquotaIVA == "P10": Natura = 'N4' # ESCLUSO ART. 15
                if AliquotaIVA == "E26": Natura = 'N4' # ESCLUSO ART. 26
                if AliquotaIVA == "X8": Natura = 'N3' # ESCLUSO ART. 15
                                                
            if (AliquotaIVA == ('null' or '')): continue #AliquotaIVA = '22.00'

            # Scrittura Dettaglio Linee
            NOME_FILE = filexml
            FILE = open(NOME_FILE, 'a+')
            #--- ELEMENTO DATI DEI BENI/SERVIZI
            FILE.write('\r<DettaglioLinee>')
            FILE.write('\r<NumeroLinea>%s</NumeroLinea>' % NumeroLinea)
            FILE.write('\r<CodiceArticolo>')
            FILE.write('\r<CodiceTipo>%s</CodiceTipo>' % CodiceTipo)
            FILE.write('\r<CodiceValore>%s</CodiceValore>' % CodiceValore)
            FILE.write('\r</CodiceArticolo>')
            FILE.write('\r<Descrizione>%s</Descrizione>' % Descrizione)
            FILE.write('\r<Quantita>%s</Quantita>' % Quantita)

            PrezzoU = "{:.4f}".format(float(PrezzoUnitario))

            FILE.write('\r<PrezzoUnitario>%s</PrezzoUnitario>' % PrezzoU)
            FILE.write('\r<PrezzoTotale>%s</PrezzoTotale>' % PrezzoTotale)
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % ALIQUOTA)
            if (ALIQUOTA == "00.00"):
                 FILE.write('\r<Natura>%s</Natura>' % Natura)
            FILE.write('\r</DettaglioLinee>')
            #print AliquotaIVA
            #raw_input("Press Enter to continue...")
        
        # Inizio campi Riepilogo
        #Aliquota "22.00"
        totale22 = '0.00'
        if len(listaImporti) >= 1:
            Riep_AliquotaIVA = '22.00'
            somma_lista = 0
            for i in range(0, len(listaImporti)):
                somma_lista += listaImporti[i]
                ImponibileImporto = "{:.2f}".format(float(somma_lista))
                Importo_Imposta = (float(ImponibileImporto) * 22) /100
                Imposta = "{:.2f}".format(float(Importo_Imposta))
            FILE.write('\r<DatiRiepilogo>')
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % Riep_AliquotaIVA)
            FILE.write('\r<ImponibileImporto>%s</ImponibileImporto>' % ImponibileImporto)
            FILE.write('\r<Imposta>%s</Imposta>' % Imposta)
            FILE.write('\r<EsigibilitaIVA>%s</EsigibilitaIVA>' % EsigibilitaIVA)
            FILE.write('\r</DatiRiepilogo>')
            totale22 = ImponibileImporto
        
        #Aliquota "04.00"
        totale04 = '0.00'
        if len(listaImporti04) >= 1:
            Riep_AliquotaIVA = '04.00'
            somma_lista = 0
            for i in range(0, len(listaImporti04)):
                somma_lista += listaImporti04[i]
                ImponibileImporto = "{:.2f}".format(float(somma_lista))
                Importo_Imposta = (float(ImponibileImporto) * 4) /100
                Imposta = "{:.2f}".format(float(Importo_Imposta))
            FILE.write('\r<DatiRiepilogo>')
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % Riep_AliquotaIVA)
            FILE.write('\r<ImponibileImporto>%s</ImponibileImporto>' % ImponibileImporto)
            FILE.write('\r<Imposta>%s</Imposta>' % Imposta)
            FILE.write('\r<EsigibilitaIVA>%s</EsigibilitaIVA>' % EsigibilitaIVA)
            FILE.write('\r</DatiRiepilogo>')
            totale04 = ImponibileImporto
            
        #Aliquota "10.00"
        totale10 = '0.00'
        if len(listaImporti05) >= 1:
            Riep_AliquotaIVA = '10.00'
            somma_lista = 0
            for i in range(0, len(listaImporti05)):
                somma_lista += listaImporti05[i]
                ImponibileImporto = "{:.2f}".format(float(somma_lista))
                Importo_Imposta = (float(ImponibileImporto) * 10) /100
                Imposta = "{:.2f}".format(float(Importo_Imposta))
            FILE.write('\r<DatiRiepilogo>')
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % Riep_AliquotaIVA)
            FILE.write('\r<ImponibileImporto>%s</ImponibileImporto>' % ImponibileImporto)
            FILE.write('\r<Imposta>%s</Imposta>' % Imposta)
            FILE.write('\r<EsigibilitaIVA>%s</EsigibilitaIVA>' % EsigibilitaIVA)
            FILE.write('\r</DatiRiepilogo>')
            totale10 = ImponibileImporto

        #Aliquota "00.00"
        totale00 = '0.00'
        if len(listaImporti00) >= 1:
            NATURA = 'N4'
            Riep_AliquotaIVA = '00.00'
            somma_lista = 0
            for i in range(0, len(listaImporti00)): # IVA ESENTE
                somma_lista += listaImporti00[i]
                ImponibileImporto = "{:.2f}".format(float(somma_lista))
                Importo_Imposta = (float(ImponibileImporto) * 0) /100
                Imposta = "{:.2f}".format(float(Importo_Imposta))
            FILE.write('\r<DatiRiepilogo>')
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % Riep_AliquotaIVA)
            FILE.write('\r<Natura>%s</Natura>' % NATURA)
            FILE.write('\r<ImponibileImporto>%s</ImponibileImporto>' % ImponibileImporto)
            FILE.write('\r<Imposta>%s</Imposta>' % Imposta)
            FILE.write('\r<EsigibilitaIVA>%s</EsigibilitaIVA>' % EsigibilitaIVA)
            FILE.write('\r</DatiRiepilogo>')
            totale00 = ImponibileImporto
            
        #Aliquota = X17
        totalex17 = '0.00'
        if len(listaImporti01) >= 1: # IVA X17 REVERSE
            NATURA = 'N6'
            Riep_AliquotaIVA = '00.00'
            somma_lista = 0
            for i in range(0, len(listaImporti01)):
                somma_lista += listaImporti01[i]
                ImponibileImporto = "{:.2f}".format(float(somma_lista))
                Importo_Imposta = (float(ImponibileImporto) * 0) /100
                Imposta = "{:.2f}".format(float(Importo_Imposta))
            FILE.write('\r<DatiRiepilogo>')
            FILE.write('\r<AliquotaIVA>%s</AliquotaIVA>' % Riep_AliquotaIVA)
            FILE.write('\r<Natura>%s</Natura>' % NATURA)
            FILE.write('\r<ImponibileImporto>%s</ImponibileImporto>' % ImponibileImporto)
            FILE.write('\r<Imposta>%s</Imposta>' % Imposta)
            FILE.write('\r<EsigibilitaIVA>%s</EsigibilitaIVA>' % EsigibilitaIVA)
            FILE.write('\r</DatiRiepilogo>')
            totalex17 = ImponibileImporto
         
        FILE.write('\r</DatiBeniServizi>')
        if DatiRitenuta == 'SI': # Dati pagamento da aggiungere in caso di Ritenuta fiscale
            CondizioniPagamento = 'TP02'
            Beneficiario = 'xxxxxxxxx'
            ModalitaPagamento = 'MP05'
            DataScadenzaPagamento = Data
            ImportoPagamento = ''
            IstitutoFinanziario = 'XXXXXXXXXXXXX'
            IBAN = 'IT52H00000000000000000'
            ABI = '00000'
            CAB = '00000'
            Totaleapagare = float(ImportoTotaleDocumento) - float(ImportoRitenuta)
            ImportoPagamento = "{:.2f}".format(float(Totaleapagare))
            FILE.write('\r<DatiPagamento>')
            FILE.write('\r<CondizioniPagamento>%s</CondizioniPagamento>' % CondizioniPagamento)
            FILE.write('\r<DettaglioPagamento>')
            FILE.write('\r<Beneficiario>%s</Beneficiario>' % Beneficiario)
            FILE.write('\r<ModalitaPagamento>%s</ModalitaPagamento>' % ModalitaPagamento)
            FILE.write('\r<DataScadenzaPagamento>%s</DataScadenzaPagamento>' % DataScadenzaPagamento)
            FILE.write('\r<ImportoPagamento>%s</ImportoPagamento>' % ImportoPagamento)
            FILE.write('\r<IstitutoFinanziario>%s</IstitutoFinanziario>' % IstitutoFinanziario)
            FILE.write('\r<IBAN>%s</IBAN>' % IBAN)
            FILE.write('\r<ABI>%s</ABI>' % ABI)
            FILE.write('\r<CAB>%s</CAB>' % CAB)
            FILE.write('\r</DettaglioPagamento>')
            FILE.write('\r</DatiPagamento>')
        FILE.write('\r</FatturaElettronicaBody>')
        FILE.write('\r</p:FatturaElettronica>')
        FILE.close()
        # FINE BODY
        if DatiRitenuta == 'SI': # Sottraggo eventuale ritenuta acconto dal totale documento
            totale_doc = float(totale22) + float(totale04) + float(totale10) + float(totale00) + float(totalex17) - float(ImportoRitenuta)
        else:
            totale_doc = float(totale22) + float(totale04) + float(totale10) + float(totale00) + float(totalex17)
        
        if REGISTRO != 'V01' or REGISTRO != 'NAC':
            TOTIMP = TOTIMP.replace(',','.')
            TOTMER = float(TOTIMP)
        
        if ("{:.2f}".format(float(totale_doc))) != ("{:.2f}".format(float(TOTMER))):
            print 'Attenzione, Totale documento differisce con la somma dei totali: %s - %s --> %s' % (totale_doc, TOTMER, filexml_nopath)
            errore_f.write('Attenzione, Totale documento differisce con la somma dei totali: %s - %s --> %s\n' % (totale_doc, TOTMER, filexml_nopath))
        print ('Scrittura XML del n. %s Num. Ass. %s effettuata' % (linea, NUMASS2))
        nfile.close()
        BolloVirtuale = 'NO'
        sposta_file(PERCORSO, filexml_nopath) # Richiama la funzione sposta_file
        
    errore_f.close()    
    nfile.close()    
    os.remove(FILEASS) #Cancella file numass.txt    
    #Fine creazione XML
            

crea_lista()
sendmail() # Comunica segnalazione
