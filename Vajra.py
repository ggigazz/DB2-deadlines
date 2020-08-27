import pyodbc
import pandas
import os
import datetime as dt
import getpass

utente = input('Inserisci il nome utente: ')
password = getpass.getpass()


class Azienda:

    def __init__(self, codage, desage, anacod, ragsoc, codpag, despag, totale, fido):
        self.codage = codage
        self.desage = desage.rstrip()
        self.anacod = anacod
        self.ragsoc = ragsoc.rstrip()
        self.codpag = codpag
        self.despag = despag.rstrip()
        self.totale = totale
        self.fido = fido

    def getAnacod(self):
        return self.anacod

    def __repr__(self):
        return "{0} + {1} + {2} + {3} + {4} + {5} + {6} + {7}".format(self.codage, self.desage, self.anacod, self.ragsoc, self.codpag, self.despag, self.totale, self.fido)


# get cwd
cwd = os.getcwd()
# fetch dati da database
bludatft = pyodbc.connect(driver='{iSeries Access ODBC Driver}',
                          system='asp.blusys.it',
                          uid=utente,
                          pwd=password)

sql_totali = "SELECT T2.BGCODAGE, T2.BGDESAGE AS AGENTE, T1.ESANACOD AS CODICE_AZIENDA, \
T2.BARAGSOC AS RAGIONE_SOCIALE, T2.PACODPAG AS CODICE_PAGAMENTO, T2.PADESPAG AS PAGAMENTO, T1.TOTALE, T2.BAFIDO AS FIDO \
FROM (SELECT ESANACOD, SUM(ESRESVAL) AS TOTALE FROM BLUDATFT.CEXSC00F WHERE ESDATSCA <= CURRENT DATE GROUP BY ESANACOD) AS T1 \
INNER JOIN (SELECT DISTINCT C.ESANACOD, B.BARAGSOC, P.PADESPAG, P.PACODPAG, A.BGCODAGE, A.BGDESAGE, B.BAFIDO FROM BLUDATFT.CEXSC00F C \
JOIN BLUDATFT.BANAG00F B ON C.ESANACOD = B.BAANACOD JOIN BLUDATFT.BAGEN40F A ON B.BACODAGE = A.BGCODAGE \
JOIN BLUDATFT.BPAGA00F P ON B.BACODPAG = P.PACODPAG WHERE BAANATIP = '1') AS T2 \
USING (ESANACOD) WHERE T1.TOTALE > 0 ORDER BY T2.BARAGSOC ASC"

# trasporto in csv
# with open("SCADENZE.csv", "w", newline="") as f:
#    writer = csv.writer(f)
#    writer.writerows(results)

totali = pandas.read_sql(sql_totali, bludatft)

totali.columns = ["CODICE_AGENTE", "NOME_AGENTE", "CODICE_AZIENDA", "RAGIONE_SOCIALE", "CODICE_PAGAMENTO", "PAGAMENTO", "TOTALE", "FIDO"]

listaDiAzienda=[Azienda(row.CODICE_AGENTE, row.NOME_AGENTE, row.CODICE_AZIENDA, row.RAGIONE_SOCIALE, row.CODICE_PAGAMENTO, row.PAGAMENTO, row.TOTALE, row.FIDO) for index, row in totali.iterrows()]

for i in listaDiAzienda:
    print(repr(i))

newcolumns = ["CODICE AGENTE", "NOME AGENTE", "CODICE AZIENDA", "RAGIONE SOCIALE", "FIDO", "PAGAMENTO", "TOTALE", "CODICE PAGAMENTO"]
totali = totali[newcolumns].loc[:, 'CODICE AGENTE':'TOTALE']
totali.to_excel("Riassunto_Scadenze.xlsx", index=False)
# os.system("pause")

totali = totali.sort_values("NOME AGENTE", ascending=True)

sql_agenti = "SELECT C.ESNUMPRO AS PROTOCOLLO, C.ESCODAGE AS CODICE_AGENTE, A.BGDESAGE AS AGENTE, C.ESANACOD AS CODICE_AZIENDA, \
B.BARAGSOC AS RAGIONE_SOCIALE, C.ESDESCAU AS MOVIMENTO, C.ESNUMDOC AS NUMERO_DOCUMENTO, C.ESDATDOC AS DATA_DOCUMENTO, \
C.ESDATSCA AS DATA_SCADENZA, B.BACODPAG AS CODICE_PAGAMENTO, C.ESDESPAG AS PAGAMENTO, C.ESRESVAL AS SCADUTO, NOTE \
FROM BLUDATFT.PEXSC00F C \
JOIN BLUDATFT.BANAG00F B ON C.ESANACOD = B.BAANACOD \
JOIN BLUDATFT.BAGEN40F A ON B.BACODAGE = A.BGCODAGE \
LEFT JOIN (SELECT NSNUMPRO, LISTAGG(TRIM(NSNOTE), ' ') AS NOTE FROM BLUDATFT.BASCA00F GROUP BY NSNUMPRO) AS N ON C.ESNUMPRO = N.NSNUMPRO \
WHERE B.BAANATIP = '1' AND C.ESDATSCA <= CURRENT DATE AND B.BAANACOD IN (SELECT DISTINCT BA.BAANACOD \
                                                                         FROM BLUDATFT.PEXSC00F SC JOIN BLUDATFT.BANAG00F BA \
                                                                         ON SC.ESANACOD = BA.BAANACOD \
                                                                         WHERE SC.ESRESVAL > '0') \
ORDER BY B.BARAGSOC ASC"  # occhio al formato data /MM/GG/YYYY !
# JOIN BLUDATFT.BPAGA00F P ON B.BACODPAG = P.PACODPAG \ rimosso per avere pagamento da fattura

sql_aziende = "SELECT BAANACOD, BARAGSOC, BACODAGE, BACODPAG, BAFIDO FROM BLUDATFT.BANAG00F WHERE BAANATIP = '1'"
sql_pagamenti = "SELECT PACODPAG, PADESPAG FROM BLUDATFT.BPAGA00F"
sql_agente = "SELECT BGCODAGE, BGDESAGE FROM BLUDATFT.BAGEN40F"
sql_pexsc00f = "SELECT ESNUMPRO, ESCODAGE, ESANACOD, ESDESCAU, ESNUMDOC, ESDATDOC, ESDATSCA, ESRESVAL \
FROM BLUDATFT.PEXSC00F ORDER BY ESDESAGE ASC, ESANACOD ASC"
sql_note = "SELECT NSNUMPRO, TRIM(NSNOTE) AS NSNOTE FROM BLUDATFT.BASCA00F"

aziende = pandas.read_sql(sql_aziende, bludatft)
aziende.columns = ['CODICE AZIENDA', 'RAGIONE SOCIALE', 'CODICE AGENTE', 'CODICE PAGAMENTO', 'FIDO']
pagamenti = pandas.read_sql(sql_pagamenti, bludatft)
pagamenti.columns = ['CODICE PAGAMENTO', 'PAGAMENTO']
agente = pandas.read_sql(sql_agente, bludatft)
agente.columns = ['CODICE AGENTE', 'NOME AGENTE']
pexsc00f = pandas.read_sql(sql_pexsc00f, bludatft)
pexsc00f.columns = ['PROTOCOLLO', 'CODICE AGENTE', 'CODICE AZIENDA', 'MOVIMENTO', 'NUMERO DOCUMENTO', 'DATA DOCUMENTO', 'DATA SCADENZA', 'SCADUTO']
note = pandas.read_sql(sql_note, bludatft).groupby('NSNUMPRO')['NSNOTE'].apply(lambda x: "%s" % ' '.join(x))
note.columns = ['PROTOCOLLO', 'NOTE']

agents = pandas.read_sql(sql_agenti, bludatft)

agents.columns = ["PROTOCOLLO", "CODICE AGENTE", "AGENTE", "CODICE AZIENDA", "RAGIONE SOCIALE", "MOVIMENTO",
                  "NUMERO_DOCUMENTO", "DATA DOCUMENTO", "DATA SCADENZA", "CODICE PAGAMENTO", "PAGAMENTO", "SCADUTO", "NOTE"]

labels = ["AGENTE", 'CODICE AZIENDA', "RAGIONE SOCIALE", "MOVIMENTO", "NUMERO_DOCUMENTO",
"DATA DOCUMENTO", "DATA SCADENZA", "PAGAMENTO", "SCADUTO", 'NOTE', "PROTOCOLLO",
"CODICE AGENTE", "CODICE PAGAMENTO"]

agents[labels].loc[:, 'AGENTE':'NOTE'].to_excel("Scadenze.xlsx", index=False)

today = dt.datetime.now()
columns_pexsc = ['PROTOCOLLO', 'CODICE AZIENDA', 'DATA SCADENZA', 'SCADUTO', 'CODICE AGENTE', 'MOVIMENTO', 'NUMERO DOCUMENTO', 'DATA DOCUMENTO']
residui_data = pexsc00f[columns_pexsc].loc[:, 'PROTOCOLLO':'SCADUTO']

residui_data['-OLTRE'] = 0
residui_data['-120'] = 0
residui_data['-90'] = 0
residui_data['-60'] = 0
residui_data['-30'] = 0
residui_data['SCADUTO'] = 0
residui_data['+30'] = 0
residui_data['+60'] = 0
residui_data['+90'] = 0
residui_data['+120'] = 0
residui_data['+OLTRE'] = 0

deltadate_v = []
for i in residui_data['DATA SCADENZA']:
#   print(today.date() - i)
    deltadate_v.append(today.date() - i)

# qui stampa gli excel divisi per agente
codici_agenti = agents.loc[:, 'AGENTE'].unique()

# totalone = agents.merge(totali.loc[:, ('CODICE AZIENDA', 'TOTALE')], on='CODICE AZIENDA', suffixes=('','_t'))
# totalone=totalone[labels]

pivot = agents.pivot_table(index=['CODICE AGENTE', 'AGENTE', 'CODICE AZIENDA', 'RAGIONE SOCIALE'], values='SCADUTO', aggfunc='sum')

# Creazione cartella Agenti
if os.path.isdir(os.path.join(cwd, 'Agenti')):
    cartella = os.path.join(cwd, 'Agenti')
else:
    cartella = os.makedirs(os.path.join(cwd, 'Agenti'))

# prepara files agente
for i in codici_agenti:
    dett = agents[labels].loc[agents['AGENTE'] == i, 'AGENTE':'NOTE']
    riass = pivot.filter(like=i, axis=0)
    nomefile = str(i).rstrip()
    writer = pandas.ExcelWriter(os.path.join(cartella, nomefile + '.xlsx'), engine='openpyxl')
    dett.to_excel(writer, sheet_name='DETTAGLIO', index=False)
    riass.to_excel(writer, sheet_name='RIASSUNTO')
    writer.save()

# file per elisa
writer = pandas.ExcelWriter("Elisa_FTS.xlsx", engine='openpyxl')
totali.to_excel(writer, sheet_name="RIASSUNTO", index=False)
for i in codici_agenti:
    dett = agents[labels].loc[agents['AGENTE'] == i, 'AGENTE':'NOTE']
    dett.to_excel(writer, sheet_name=str(i).rstrip(), index=False)
writer.save()
