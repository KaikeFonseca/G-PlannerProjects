import win32com.client
import sys
import subprocess
import time
from pandas import read_excel

class Sap:
    def __init__(self) -> None:
        self.__session = None
        self.__application = None
        self.__connection = None
        self.__SapGuiAuto = None

    def open_and_login(self, user:str, pwd:str, lang:str='EN'):
        application_name = "PS1 (Base de Produção)"
        try:
            path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            subprocess.Popen(path)
            time.sleep(3)

            self.__SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(self.__SapGuiAuto) == win32com.client.CDispatch:
                return

            self.__application = self.__SapGuiAuto.GetScriptingEngine
            if not type(self.__application) == win32com.client.CDispatch:
                self.__SapGuiAuto = None
                return
            self.__connection = self.__application.OpenConnection(application_name, True)

            if not type(self.__connection) == win32com.client.CDispatch:
                self.__application = None
                self.__SapGuiAuto = None
                return

            self.__session = self.__connection.Children(0)
            if not type(self.__session) == win32com.client.CDispatch:
                self.__connection = None
                self.__application = None
                self.__SapGuiAuto = None
                return

            self.__session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
            self.__session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pwd
            self.__session.findById("wnd[0]/usr/txtRSYST-LANGU").text = lang
            self.__session.findById("wnd[0]").sendVKey(0)
            
            try:
                self.__session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                self.__session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus()
                self.__session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                pass

        except:
            print('erro')
            print(sys.exc_info()[0])

    def sap_aberto(self, connection_num:int):
        try:
            self.__SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(self.__SapGuiAuto) == win32com.client.CDispatch:
                return
            self.__application = self.__SapGuiAuto.GetScriptingEngine
            if not type(self.__application) == win32com.client.CDispatch:
                self.__SapGuiAuto = None
                return
            if self.__application.connections.count == 0:
                return 0
        except:
            return 0
        
    def pd_excel(self, username:str):

        config ={
            "arquivo_pd": rf"\\sb2-fs\4_GESTAO_DA_QUALIDADE$\00_KPI_&_GSD\Melhoria Contínua\Fluxos\pwd\sap_pd.xlsx",
            "planilha_pd": "Plan1",
            "user_sap": username 
        }

        df = read_excel(config["arquivo_pd"], sheet_name=config["planilha_pd"])

        print(df)

        # Encontrar a linha onde a coluna "user" é igual a '[user_sap]'
        resultado = df[df['user'] == config["user_sap"]]

        # Verificar se a palavra foi encontrada e obter o valor da coluna pd
        if not resultado.empty:
            valor_pd = resultado['pd'].values[0]  # Obtém o primeiro valor correspondente
        #    print(f'A pd do usuário_SAP {username} é: {valor_pd}')
        #else:
        #    print(f'O usuário {username} não foi encontrado.')
        return valor_pd

    def get_existing_connection(self, connection_num:int):
        self.__SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.__SapGuiAuto) == win32com.client.CDispatch:
            return

        self.__application = self.__SapGuiAuto.GetScriptingEngine
        if not type(self.__application) == win32com.client.CDispatch:
            self.__SapGuiAuto = None
            return
        
        if self.__application.connections.count == 0:
            raise Exception('Nenhuma conexão aberta')
        if connection_num > self.__application.connections.count -1:
            raise Exception(f'Conexão {connection_num} não existe')

        if connection_num == -1:
            self.__connection = self.__application.Children(self.__application.connections.count -1)
        else:
            self.__connection = self.__application.Children(connection_num)

        if not type(self.__connection) == win32com.client.CDispatch:
            self.__application = None
            self.__SapGuiAuto = None
            return

        self.__session = self.__connection.Children(0)

    def enter_transaction(self, t_code:str):
        self.__session = self.__connection.Children(0)
        self.__session.findById("wnd[0]/tbar[0]/okcd").Text = f'/n{t_code}'
        self.__session.findById("wnd[0]").sendVKey(0)

    def input_text(self, path:str, text:str):
        self.__session.findById(path).Text = text
    
    def get_text(self, path:str) -> str:
        return self.__session.findById(path).Text

    def get_path(self, path:str):
        return self.__session.findById(path)

    def set_focus(self, path:str):
        self.__session.findById(path).setFocus()

    def change_checkbox(self, path:str, param1:str, param2:str, value:bool):
        self.__session.findById(path).changeCheckbox(param1,param2,value)

    def select(self, path:str):
        self.__session.findById(path).select()

    def press_button(self, path:str):
        self.__session.findById(path).Press()

    def send_key(self,key:int, repeat:int=1):
        for _ in range(repeat):
            self.__session.findById("wnd[0]").sendVKey(key)

    def send_enter_key(self,repeat:int=1):
        for _ in range(repeat):
            self.send_key(0)
            # if self.get_status_mesage_type == 'E':
            #     raise Exception('Error')

    def send_f8_key(self,repeat:int=1):
        for _ in range(repeat):
            self.send_key(8)

    def send_esc_key(self,repeat:int=1):
        for _ in range(repeat):
            self.send_key(12)

    def back(self,repeat:int=1):
        for _ in range(repeat):
            self.__session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    def get_status_mesage_type(self):
        return self.__session.findById("wnd[0]/sbar").MessageType
    
    def get_status_mesage_number(self):
        return self.__session.findById("wnd[0]/sbar").MessageNumber
    
    def get_status_mesage_id(self):
        return self.__session.findById("wnd[0]/sbar").MessageId
    
    def get_status_mesage(self):
        return self.__session.findById("wnd[0]/sbar").Text
    
    def get_status_mesage_textTbar_w1(self):
        return self.__session.findById("wnd[1]/usr").MessageId

    def get_status_mesage_textUsr_w1(self):
        return self.__session.findById("wnd[1]/usr").Text

    def get_transaction_name(self):
        return self.__session.Info.Transaction

    def get_children(self):
        return len(self.__session.children)

    def get_active_window_name(self):
        return self.__session.ActiveWindow.Name

    def get_cell_value(self, path:str, line:int, column_name:str):
        return self.__session.findById(path).getCellValue(line,column_name)

    def get_element(self, path:str):
        return self.__session.findById(path)
    
    def current_cell_row(self, path:str, value:int):
        self.__session.findById(path).currentCellRow = value

    def selected_rows(self, path:str, value:str):
        self.__session.findById(path).selectedRows = value

    def caret_position(self, path:str, value:int):
        self.__session.findById(path).caretPosition = value

    def press_toolbar_context_button(self):
        self.__session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")

    def press_toolbar_context_button_view(self):
        self.__session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton ("&MB_VIEW")

    def select_context_menu_item(self):
        self.__session.findById("wnd[0]/shellcont/shell").selectContextMenuItem ("&PC")

    def press_ctn_context_button(self):
        self.__session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")

    def select_context_menu_item_print(self):
        self.__session.findById("wnd[0]/shellcont/shell").selectContextMenuItem ("&PRINT_BACK_PREVIEW")

    def select_ctn_context_menu_item_xxl(self):
        self.__session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem ("&XXL")
    
    def selectField(self, path:str):
        self.__session.findById(path).Selected = True
    
    def notSelectField(self, path:str):
        try:
            self.__session.findById(path).Selected = False
        except:
            return True

if __name__ == '__main__':
    sap = Sap()
    # sap.open_and_login('AA3IP001','Eng!2028')
    # sap.enter_transaction('MM03')
    parts = [
        'K234000B6'
    ]
    sap.get_existing_connection(1)
    for p in parts:
        sap.input_text('wnd[0]/usr/ctxtRMMG1-MATNR',p)
        sap.send_enter_key()
        print(sap.get_status_mesage_id())
        # sap.send_enter_key(3)
        # sap.select("wnd[0]/usr/tabsTABSPR1/tabpSP24")
        # print(p,sap.get_text('wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/subSUBCURR:SAPLCKMMAT:0200/txtCKMMAT_DISPLAY-STPRS_1'))
        # sap.back()

