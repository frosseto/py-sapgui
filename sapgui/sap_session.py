import win32com.client
import os
import time
import log
import traceback
from functools import wraps
from constants import sap_conn_str


def retry(exceptions, tries=4, delay=3, backoff=2, logger=None):
    def deco_retry(f):
        @wraps(f)
        def f_retry(*args, **kwargs):
            logger = log.setup_custom_logger('sap')
            mtries, mdelay = tries, delay
            while mtries > 1:
                try:
                    return f(*args, **kwargs)
                except exceptions as e:
                    msg = '{}, Nova tentativa em {} s...'.format(e, mdelay)
                    if logger:
                        logger.warning(msg)
                    else:
                        print(msg)
                    time.sleep(mdelay)
                    mtries -= 1
                    mdelay *= backoff
            return f(*args, **kwargs)

        return f_retry  # true decorator

    return deco_retry


class SapGui:
    '''Classe SapGui(self, pmode=False, numero_conexao=0, historico=True)
    * pmode = True - uses the first sap session available
            False - open a new sap window (using a .sap file) and select last opened connection
    * numero_conexao - if pmode = True,  is used to select the connection
    * historico - history sap app flag 
    connections = []'''


    def sap_connections(self):
        return self.connections

    def __init__(self, pmode=False, numero_conexao=0, historico=True):
        self.__sap_connect(pmode, numero_conexao, historico)


    @retry(Exception, tries=6)
    def __sap_connect(self, pmode, numero_conexao, historico):
        logger = log.setup_custom_logger('sap')
        if not pmode:

            dir_path = os.path.dirname(os.path.realpath(__file__))
            path = os.path.join(dir_path, 'tx.sap')
            sap_gui_auto = self.__get_sap_gui(path)
            appl = sap_gui_auto.GetScriptingEngine

            while appl.Connections.Count == 0:
                time.sleep(5)
            con = appl.Connections.Count - 1
            connection = appl.Children(con)
            self.connections.append(con)
            self.connection_current = con

            if connection.sessions.Count > 0:
                self.session = connection.Children(0)

        else:
            sap_gui_auto = self.__get_sap_gui() 
            appl = sap_gui_auto.GetScriptingEngine
            appl.historyEnabled = historico

            if appl.Connections.Count > 0:
                connection = appl.Children(0)
                self.connections.append(0)
                self.connection_current = 0
            else:
                try:   
                    connection = appl.openConnectionByConnectionString(
                        sap_conn_str, 
                        True, True)
                except Exception as e:
                    logger.error(traceback.format_exc())

            if connection.sessions.count > 0:
                self.session = connection.sessions(numero_conexao)
            else:
                self.session_close(sap_kill=True)
                self.__sap_connect(self, pmode, numero_conexao, historico)


    @retry(Exception, tries=6)
    def executar(self):
        self.session.findById("wnd[0]").sendVKey(8)


    @retry(Exception, tries=6)
    def session_findby_text(self, sap_path: str, value: str = None):
        if value:
            self.session.findById(sap_path).text = str(value)
        else:
            text = str(self.session.findById(sap_path).Text)
            return text


    @retry(Exception, tries=6)
    def send_vkey(self, sap_path: str = "wnd[0]", key: int = 0):
        self.session.findById(sap_path).sendVKey(key)


    @retry(Exception, tries=6)
    def press(self, sap_path: str):
        self.session.findById(sap_path).press()


    def __get_sap_gui(self, path=None, first=True):
        sap_gui_auto = None
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if path and first:
                n_conexoes = sap_gui_auto.GetScriptingEngine.Connections.Count
                os.startfile(path)
                while n_conexoes == sap_gui_auto.GetScriptingEngine.Connections.Count:
                    time.sleep(5)
        except:
            if first:
                os.startfile(path if path else r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
            time.sleep(5)
            sap_gui_auto = self.__get_sap_gui(first=False)
        finally:
            return sap_gui_auto


    def session_close(self, sap_kill=False):
        if not sap_kill:
            try:
                self.session.findById("wnd[0]").Close()
                try:
                    print('Sessão sap finalizada.')
                    self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                except:
                    None
            except:
                None
        else:
            os.system('TASKKILL /F /IM saplogon.exe')


    @retry(Exception, tries=6)
    def status_text(self):
        try:
            return self.session.findById("wnd[0]/sbar").Text
        except:
            return ''


    @retry(Exception, tries=6)
    def status_type(self):
        # E - error, S - sucess, W - warning, A - abort, I - information
        try:
            return self.session.findById("wnd[0]/sbar").messageType
        except:
            return 'F'


    @retry(Exception, tries=6)
    def has_popup(self):
        try:
            return self.sessionsession.findById("wnd[0]/sbar").messageAsPopup
        except:
            return False


    @retry(Exception, tries=6)
    def start_transaction(self, transaction: str):
        try:
            self.session.startTransaction(transaction)
            return True
        except:
            return False


    @retry(Exception, tries=6)
    def enter_no_warnings(self):
        try:
            while True:
                self.session.findById("wnd[0]").sendVKey(0)
        except:
            return False


    def chamar_variante(self, variante: str):
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/txtV-LOW").text = variante
        self.session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        self.session.findById("wnd[1]/usr/ctxtENVIR-LOW").Text = ""
        self.session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
        self.session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()


    def exportar_txt(self, file_path: str, file_name: str):
        menus = [1, 2, 3, 10, 11]
        for idx, menu in enumerate(menus):
            try:
                menu_path = f"wnd[0]/mbar/menu[0]/menu[{menu}]/menu[2]"
                menu_texto = self.session.findById(menu_path).Text
                if menu_texto.find('File') >= 0:
                    self.session.findById(menu_path).select()
                    print('sucesso...')
                    break
            except Exception as e:
                None
        try:
            menu_path = "wnd[0]/tbar[1]/btn[45]"
            menu_tooltip = self.session.findById(menu_path).Tooltip
            if menu_tooltip.find('File') >= 0:
                self.session.findById(menu_path).press()
                print('sucesso...')
        except Exception as e:
            None

        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()


    def chamar_variante_exibicao(self, variante: str, tcode: str):
        menus = [4, 3]
        menu_texto = ''
        for idx, menu in enumerate(menus):
            try:
                # print(f'tentativa {idx}...')
                menu_path = f"wnd[0]/mbar/menu[{menu}]/menu[0]/menu[1]"
                menu_texto = self.session.findById(menu_path).Text
                if menu_texto.find('Selecionar') >= 0:
                    self.session.findById(menu_path).select()
                    print('sucesso...')
                    break
            except Exception as e:
                None
        if menu_texto.find('Selecionar') < 0:
            try:
                tooltip_menu_text = self.session.findById("wnd[0]/tbar[1]/btn[33]").Tooltip
                if tooltip_menu_text.find('Selecionar') >= 0:
                    self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
            except Exception as e:
                None
        if tcode == "yspm_listacentrab" or tcode == "mrp" or tcode == "yspm_textos":
            self.session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectColumn("VARIANT")
            self.session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").contextMenu()
            self.session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectContextMenuItem("&FIND")
            self.session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/" + variante
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[2]").close()
            self.session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell()
        else:
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn(
                "VARIANT")
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem(
                "&FIND")
            self.session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/" + variante
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[2]").close()
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        try:
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e:
            pass

            
    def copy(self, texto: str):
        import pandas as pd
        df = pd.DataFrame([texto])
        import csv
        df.to_clipboard(index=False, header=False, quoting=csv.QUOTE_NONE, quotechar="",  escapechar="\\")


    def open_url(self, url: str):
        import urllib
        return urllib.urlopen(url)


    def curr_user(self):
        return os.environ("UserName")