from MSPyBentley import * 
from MSPyBentleyGeom import *
from MSPyECObjects import *
from MSPyDgnPlatform import *
from MSPyDgnView import *
from MSPyMstnPlatform import *

import os
from openpyxl import Workbook

# DEFINIZIONE PERCORSO CARTELLA "Excel"
cartella_excel = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Excel'))

# CREAZIONE CARTELLA NEL CASO NON ESISTA
os.makedirs(cartella_excel, exist_ok=True)

# CREAZIONE DEL FILE EXCEL
wb = Workbook()
ws = wb.active
ws.title = "Etichette Formattate"

# FUNZIONE CHE AGGIUNGE UN ETICHETTA IN MODO FORMATTATO SE LA STRINGA NON CONTIENE LA PAROLA 'SOCKET'
def aggiungi_etichetta(stringa, ws, riga_corrente):
    if "SOCKET" in stringa.upper():
        print(f"Etichetta ignorata: {stringa}")
        return riga_corrente
    
    parti = stringa.split()
    print(f"Inserisco etichetta: {parti}")
    for colonna, parte in enumerate(parti, start=1):
        ws.cell(row=riga_corrente, column=colonna, value=parte)
    riga_corrente += 1
    return riga_corrente

# FUNZIONE CHE SALVA IL FILE EXCEL CREATO
def salva_excel(nome_file="etichette_formattate.xlsx"):
    percorso_file = os.path.join(cartella_excel, nome_file)
    wb.save(percorso_file)
    print(f"File salvato in: {percorso_file}")

def main():
    riga_corrente = 1
    # DEFINIZIONE DEL MODELLO
    model = ISessionMgr.GetActiveDgnModel()

    # DEFINIZIONE DELL'HANDLER
    handler = TextNodeHandler.GetInstance()

    # CICLO SUI GRAPHIC ELEMENTS DEL MODELLO (el -> elemento grafico)
    for el in model.GetGraphicElements():

        # CONTROLLA SE L'ELEMENTO È UN TEXT NODE (tipo 2)
        if el.GetElementType() == 2:
            
            try:
                # CONVERSIONE A ELEMENT HANDLE
                # Questo è necessario per poter utilizzare i metodi dell'handler e quindi estrarre il testo
                eh = ElementHandle(el)

                # ESTRAZIONE DEL TESTO
                text_str = handler.GetFirstTextPartValue(eh).ToString()
                text_str = str(text_str)

                # OUTPUT 
                # print(f"\nElemento ID={el.GetElementId()}, Testo: {text_str}")

                # AGGIUNTA DELL'ETICHETTA AL FILE EXCEL
                riga_corrente = aggiungi_etichetta(text_str, ws, riga_corrente)

            # GESTIONE DEGLI ERRORI
            # Se si verifica un errore durante l'estrazione del testo, lo ignora
            # (le note che prendiamo in considerazione non presentano errori, ma è sempre buona norma gestire le eccezioni)
            except Exception as e:
                continue
    
    # SALVATAGGIO DEL FILE EXCEL
    salva_excel()

if __name__ == "__main__":
    main()

