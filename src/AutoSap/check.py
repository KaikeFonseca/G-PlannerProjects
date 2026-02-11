from sap import Sap

def check_mm03(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('MM03')
    sap.input_text('wnd[0]/usr/ctxtRMMG1-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() == '305':
        return False
    elif sap.get_status_mesage_number() != '':
        raise Exception(sap.get_status_mesage_number())
    sap.send_enter_key(2)
    sap.back()
    return True

def check_standard(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('MM03')
    sap.input_text('wnd[0]/usr/ctxtRMMG1-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() == '305':
        return False
    elif sap.get_status_mesage_number() != '':
        raise Exception(sap.get_status_mesage_number())
    sap.send_enter_key(2)
    sap.select("wnd[0]/usr/tabsTABSPR1/tabpSP24")
    price = sap.get_text('wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/subSUBCURR:SAPLCKMMAT:0200/txtCKMMAT_DISPLAY-STPRS_1')
    sap.back()
    if price == '0,00':
        return False
    else:
        return True
    
def check_cs03(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('CS03')
    sap.input_text('wnd[0]/usr/ctxtRC29N-MATNR',part)
    sap.input_text('wnd[0]/usr/ctxtRC29N-STLAN','3')
    sap.send_enter_key()
    if sap.get_status_mesage_number() != '':
        return False
    sap.back()
    return True

def check_ca23(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('CA23')
    sap.input_text('wnd[0]/usr/ctxtRC27M-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() != '':
        return False
    if sap.get_active_window_name() == 'wnd[1]':
        sap.send_esc_key()
        return False
    sap.back(4)
    return True

def check_c223(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('C223')
    sap.input_text('wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MATNR',part)
    sap.send_enter_key()
    if sap.get_status_mesage_number() == '058':
        return False
    elif sap.get_status_mesage_number() == '068':
        return False
    elif sap.get_status_mesage_number() == '':
        return True

def check_kkf6n(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('KKF6N')
    sap.input_text('wnd[0]/usr/subSUB_SELECT:SAPMKOSA_46:0110/ctxtMAT-LOW',part)
    sap.send_enter_key()
    try:
        sap.get_path('wnd[0]/shellcont/shell/shellcont[1]/shell[1]').GetItemText('          2','&Hierarchy')
    except:
        pass
    if sap.get_status_mesage_number() == '100':
        return False
    try:
        sap.get_path('wnd[0]/shellcont/shell/shellcont[1]/shell[1]').GetItemText('          2','&Hierarchy')
        return True
    except:
        return False
    

def check_pop3(part:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('POP3')
    sap.input_text('wnd[0]/usr/ctxtPIKP-PIID',f'{part}-001')
    sap.send_enter_key()
    if sap.get_status_mesage_number() == '002':
        return False
    elif sap.get_status_mesage_number() == '':
        sap.back()
        return True
    
def check_pof3(part:str, pack:str, connection_num:int=-1):
    sap = Sap()
    sap.get_existing_connection(connection_num)
    sap.enter_transaction('POF3')
    sap.input_text('wnd[0]/usr/ctxtP000-KSCHL',pack)
    sap.send_enter_key()
    if pack.lower() == 'zshi':
        # IMPLEMENTAR RECEBEDOR DE MERCADORIA
        sap.select('wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[0,0]')
        sap.send_enter_key()
        sap.input_text('wnd[0]/usr/ctxtF002',part)
    elif pack.lower() == 'zsto':
        sap.select('wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]')
        sap.send_enter_key()
        sap.input_text('wnd[0]/usr/ctxtF002-LOW',part)
    else:
        raise Exception(f'pack {pack} não identificado')
    sap.send_f8_key()
    # return sap.get_status_mesage_number()
    if sap.get_status_mesage_number() == '058' or sap.get_status_mesage_number() == '021':
        sap.back()
        return False
    elif sap.get_status_mesage_number() == '':
        sap.back()
        return True