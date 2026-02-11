from sap import Sap
import pandas as pd

def get_project_by_component(part:str,connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('MM03')
    sap.input_text('wnd[0]/usr/ctxtRMMG1-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() != '':
        print(part, 'Peça não encontrada')
        return False
    sap.send_enter_key(2)
    sap.select("wnd[0]/usr/tabsTABSPR1/tabpSP05")
    print(part, sap.get_text('wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLZMM_GF_MTRO_MAT_CUST:9156/ctxtMVKE-MVGR1'))

def get_stock_by_component(part:str,connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('ZQLOT')
    sap.input_text('wnd[0]/usr/ctxtMATNR-LOW',part)
    sap.send_f8_key()
    sap.set_focus('wnd[0]/usr/lbl[54,0]')
    sap.send_key(2)
    csv_file = sap.get_text('wnd[0]/usr/lbl[0,2]')
    df = pd.read_csv(csv_file,sep='%',encoding='ISO-8859-1')

def get_description_by_component(part:str,connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    if sap.get_transaction_name() != 'MM03':
        sap.enter_transaction('MM03')
    sap.input_text('wnd[0]/usr/ctxtRMMG1-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() != '':
        print(part, 'Peça não encontrada')
        return False
    sap.send_enter_key(2)
    sap.select("wnd[0]/usr/tabsTABSPR1/tabpSP01")
    text = sap.get_text('wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX')
    sap.back()
    return text
