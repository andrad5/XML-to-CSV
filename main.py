import xmltodict
import os
import json
import pandas as pd

def info(archive_name, values):
   # with open(f'namepatch/{archive_name}', "rb") as archive_xml:
    with open(f'nfs/{archive_name}', "rb") as archive_xml:
        dic_archive = xmltodict.parse(archive_xml)
        
        # -- Parametros do arquivo XML 
        try:
            if "NFe" in dic_archive:
                infos_xml = dic_archive ["NFe"]["infNFe"]
            elif "nfeProc" in dic_archive:
                infos_xml = dic_archive ["nfeProc"]["NFe"]["infNFe"]
            else:
                infos_xml = dic_archive ["NFeProc"]["nfeProc"]["NFe"]["infNFe"]

            Info_I =  infos_xml["@Id"]
            info_II = infos_xml["emit"]["xNome"]
            info_III = infos_xml["dest"]["xNome"]
            info_IV = infos_xml["dest"]["enderDest"]

            if "vol" in infos_xml["transp"]:
               info_V = infos_xml["transp"]["vol"]["pesoB"]
            else:
                info_V = "Null"

            values.append([Info_I, info_II, info_III, info_IV, info_V])
            
        except Exception as e:
            print("~~ E R R O ~~")
            print(e)
            print(json.dumps(dic_archive, indent=4))
#-------------------------------------------------------------------------------------------


list_archive = os.listdir("nfs")

columns = ["Info_I", "Info_II", "Info_III", "Info_IV", "Info_V"]
values = []

for archive_name in list_archive:
    info(archive_name, values)

table = pd.DataFrame(columns=columns, data=values)
table.to_excel("NameArchive.xlsx", index=False)