from tkinter import *
from tkinter import filedialog
import sys
import xml.etree.ElementTree as ET
import pandas as pd
import itertools
import os
from datetime import datetime




def close_window():
    window.destroy()
    sys.exit()
    
    
def browse_file():
    window.filename = filedialog.askopenfilename(title="Seleziona File")
    file_text = "File Selezionato: " + window.filename
    Label(window, text=file_text, bg="grey", fg="white").grid(row=5, column=2, sticky="W")
   

def browse_directory():
    global files_dir
    files_dir = filedialog.askdirectory()
    Label(window, text="Cartella Selezionata: " +files_dir, bg="grey", fg="white").grid(row=11, column=2, sticky="W")  


def convert_file():
    df_invoice = parseXML(window.filename)
    df_invoice.to_excel(window.filename[:-4] + ".xlsx", index=False)
    Label(window, text="File convertito correttamente", bg="green", fg="white").grid(row=6, column=2, sticky="W")
    
def convert_directory():
    os.chdir(files_dir)
    files_list = os.listdir(files_dir)
    
    for item in files_list:     
        if item[-4:] == ".xml":
            df_invoice = parseXML(item).copy()
            
            try:       
                frames = [all_df_invoice, df_invoice]
                all_df_invoice = pd.concat(frames, ignore_index=True)

            except:
                all_df_invoice = df_invoice
           
    now = datetime.now()
    dt_string = now.strftime("%Y%m%d_%H%M%S")
    all_df_invoice.to_excel("ConversioneXML"+ dt_string + ".xlsx", index=False)
    Label(window, text="Cartella convertita correttamente", bg="green", fg="white").grid(row=12, column=2, sticky="W")
     
    
def parseXML(xmlfile):
    
    global header_data_dic,body_data_dic, invoice_element_dic
    
    fields = ["Denominazione Fornitore",
               "CF Fornitore",        
               "ID Paese Fornitore",   
               "ID Codice Fornitore",                           
               "Denominazione Cliente", 
               "CF Cliente",            
               "ID Paese Cliente",     
               "ID Codice Cliente",
               "Numero Documento",
               "Data Fattura",
               "Tipo Documento",
               "Body"  
               ]
    
    fields_structure_header = { fields[0]: ["./FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione"],
                                fields[1]: ["./FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/CodiceFiscale"],
                                fields[2]: ["./FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdPaese"],
                                fields[3]: ["./FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice"],                         
                                fields[4]: ["./FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione"],
                                fields[5]: ["./FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/CodiceFiscale"],
                                fields[6]: ["./FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/IdFiscaleIVA/IdPaese"],
                                fields[7]: ["./FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/IdFiscaleIVA/IdCodice"],
                                fields[8]: ["./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero"],
                                fields[9]: ["./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data"],
                                fields[10]: ["./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/TipoDocumento"],
                         }
       
   
    #create element tree object
    tree = ET.parse(xmlfile)
    
    #parse the single elements
    for field in fields_structure_header:
        item_path = fields_structure_header[field][0]
        
        for item in tree.findall(item_path):
            fields_structure_header[field].append(item.text.encode('utf8').decode('utf8'))
                                                  
   
    
    #parse multiple elements
    item_path = "./FatturaElettronicaBody/DatiBeniServizi/DettaglioLinee"
    i=0
    invoice_element_dic={}
    for invoice_element in tree.findall(item_path):
        item_keys = []
        item_values = []
        for item in list(invoice_element):
            item_keys.append(item.tag)
            try:
                item_values.append(item.text.encode('utf8').decode('utf8'))
            except:
                item_values.append(None)
        invoice_element_dic[i] = {item_keys[i]: item_values[i] for i in range(len(item_keys))}  
        i+=1
    
    #cercare tutti i campi 
    body_keys = ['AliquotaIVA', 'AltriDatiGestionali', 'Descrizione', 'NumeroLinea', 'PrezzoTotale', 'PrezzoUnitario', 'Quantita', 'RiferimentoAmministrazione', 'TipoCessionePrestazione', 'UnitaMisura']
    # body_keys = []
    # for entry in invoice_element_dic:
    #     for item in invoice_element_dic[entry]:
    #         if item not in body_keys:
    #             body_keys.append(item)
    # print(sorted(body_keys))
    
    #creare il dizionario per il body     
    body_data_dic = {k: [] for k in body_keys}       
    for entry in invoice_element_dic:
        for item in body_keys:
            try:
                body_data_dic[item].append(invoice_element_dic[entry][item])
            except:
                body_data_dic[item].append(None)
    
    #concluding the single field
    header_data_dic = {}
    for field in fields_structure_header: 
        try: 
            header_data_dic[field] = list(itertools.repeat(fields_structure_header[field][1], len(invoice_element_dic)))
        except:
            header_data_dic[field] = list(itertools.repeat(None, len(invoice_element_dic)))
           
   
    #merge the 2 dictionaries
    global df_invoice
    data_dic = {**header_data_dic, **body_data_dic}
    df_invoice = pd.DataFrame(data_dic)
    df_invoice["P.Iva Cliente"] = df_invoice["ID Paese Cliente"] + df_invoice["ID Codice Cliente"]
    df_invoice["P.Iva Fornitore"] = df_invoice["ID Paese Fornitore"] + df_invoice["ID Codice Fornitore"]
    
    #formatting
    
    
    df_invoice_filtered = df_invoice[["Denominazione Fornitore",
                                      "CF Fornitore", 
                                      "P.Iva Fornitore",
                                      "Denominazione Cliente", 
                                      "CF Cliente", 
                                      "P.Iva Cliente",
                                      "Numero Documento",
                                      "Data Fattura",
                                      "Tipo Documento",
                                      "NumeroLinea",
                                      'TipoCessionePrestazione',
                                      'Descrizione',
                                      "Quantita",
                                      "UnitaMisura",
                                      "PrezzoUnitario",
                                      "PrezzoTotale",
                                      "AliquotaIVA"
                                      ]]
    
    body_keys_map = {"Denominazione Fornitore": "Denominazione Fornitore",
                     "CF Fornitore"           : "CF Fornitore",
                     "P.Iva Fornitore"        : "P.Iva Fornitore", 
                     "Denominazione Cliente"  : "Denominazione Cliente" ,
                     "CF Cliente"             : "CF Cliente",
                     "P.Iva Cliente"          : "P.Iva Cliente",
                     "Numero Documento"       : "Numero Documento",
                     "Data Fattura"           : "Data Fattura",
                     "Tipo Documento"         : "Tipo Documento",
                     "NumeroLinea"            : 'Numero Linea',
                     'TipoCessionePrestazione': 'Tipo Cessione Prestazione', 
                     'Descrizione'            : 'Descrizione', 
                     "Quantita"               : 'Quantità', 
                     "UnitaMisura"            : 'Unità Misura',
                     "PrezzoUnitario"         : 'Prezzo Unitario',
                     "PrezzoTotale"           : 'Prezzo Totale',
                     "AliquotaIVA"            : 'Aliquota IVA'
                     }
                                      
    df_invoice_filtered.rename(columns=body_keys_map, inplace=True)
                                    
    return df_invoice_filtered
       
    
    
  
if __name__ == "__main__":
    
    #creating the window
    window = Tk()
    window.title("XML Converter")
    window.configure(background= '#49A')
    window.geometry("900x600")
    
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=1, column=1, sticky=W)    #space
    
    
    # --> single file conversion
    #select a file
    Label(window, text="Conversione singolo File  ", bg="black", fg="white", width=30).grid(row=3, column=1, sticky="W") 
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=4, column=1, sticky=W) #space
    
    Button(window, text="Seleziona File", width=14, command=browse_file).grid(row=5, column=1, sticky="W")
    
    #convert a file
    Button(window, text="Converti File", width=14, command=convert_file).grid(row=6, column=1, sticky="W")
    
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=7, column=1, sticky=W) #space
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=8, column=1, sticky=W) #space
    
    # --> folder conversion
    #select a folder
    Label(window, text="Conversione Intera Cartella", bg="black", fg="white", width=30).grid(row=9, column=1, sticky="W")
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=10, column=1, sticky=W) #space
    
    Button(window, text="Seleziona Cartella", width=14, command=browse_directory).grid(row=11, column=1, sticky="W")
    
    #convert an entire folder
    Button(window, text="Converti Cartella", width=14, command=convert_directory).grid(row=12, column=1, sticky="W")
    
    
    #exit 
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=13, column=1, sticky=W) #space
    Label(window, text="", bg="#49A", fg="#49A", width=20).grid(row=14, column=1, sticky=W) #space
    Button(window, text="Exit", width=14, command=close_window).grid(row=15, column=2, sticky="W")
    
    
    #running the GUI
    window.mainloop()
    
    