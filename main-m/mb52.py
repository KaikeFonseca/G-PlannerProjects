import sys
sys.path.append(r'\\sb2-fs\4_GESTAO_DA_QUALIDADE$\00_KPI_&_GSD\Melhoria Contínua\Fluxos\014. Planner Project\src')  # Substitua pelo caminho correto

from time import sleep

from AutoSap.sap import Sap # type: ignore
from GetDate.get_date import get_period # type: ignore

from TerminateProcess.close_process import encerrar_processos # type: ignore

terminate = encerrar_processos()

sap = Sap()
get_date = get_period()
dataFormatada = get_date.today_ymd_period()

def updateStock():
    sap = Sap()
    if sap.sap_aberto(0) == 0:
        print("SAP NÃO ABERTO")
        sap.open_and_login('SORIP002','Masterlog.2025@', 'PT')
    else:
        print("SAP JÁ ABERTO")

    sap.get_existing_connection(0)
    sap.enter_transaction('mb52')

    sap.selectField("wnd[0]/usr/chkPA_SOND") #-SET TAMBÉM ESTOQUE ESPECIAL

    sap.notSelectField("wnd[0]/usr/chkNEGATIV") #-EXIBIR SÓ ESTOQUE NEGATIVOS 
    sap.notSelectField("wnd[0]/usr/chkXMCHB") #-EXIBIR ESTOQUE DE LOTES
    sap.notSelectField("wnd[0]/usr/chkNOZERO") #-SEM LINHAS ESTOQUE ZERO
    sap.notSelectField("wnd[0]/usr/chkNOVALUES") #-NÃO EXIBIR VALORES

    #sap.select("wnd[0]/usr/radPA_HSQ") #-REPRESENTAÇÃO HIERARQUICA
    sap.select("wnd[0]/usr/radPA_FLT") #-REPRESENTAÇÃO NÃO HIERARQUICA
    sap.notSelectField("wnd[0]/usr/chkB_FILE") #-B-FILE

    sap.input_text('wnd[0]/usr/ctxtP_VARI','sorpcp') #-LAYOUT

    sap.send_f8_key()

    sap.select("wnd[0]/mbar/menu[0]/menu[3]/menu[1]")
    sap.press_button("wnd[1]/tbar[0]/btn[0]")
    sap.input_text("wnd[1]/usr/ctxtDY_PATH",r'\\sb2-fs\4_GESTAO_DA_QUALIDADE$\00_KPI_&_GSD\Melhoria Contínua\APPs\Automação Planner\Estoque' )
    sap.input_text("wnd[1]/usr/ctxtDY_FILENAME", 'estoque.xlsx')
    sap.press_button('wnd[1]/tbar[0]/btn[11]')

    sap.send_esc_key(2)
    sleep(3)
    terminate.sap_logon()
    terminate.excel()
    #sleep(2)