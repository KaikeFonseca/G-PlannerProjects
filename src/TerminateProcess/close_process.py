from psutil import process_iter, Process, NoSuchProcess, AccessDenied, ZombieProcess # type: ignore
import threading

class encerrar_processos:
    def excel(self): # repeat = True -> usar quando é certeza que o processo sempre será executado
        class aguardar_processo:
            def check_excel(event):
                for proc in process_iter(['pid', 'name']):
                    if 'EXCEL.EXE' in proc.info['name']:
                        event.set()
            def main():
                event = threading.Event()
                excel_started = False

                # Inicia uma thread para verificar se o Excel foi iniciado
                excel_checker_thread = threading.Thread(target=aguardar_processo.check_excel, args=(event,))
                excel_checker_thread.start()

                while not excel_started:
                    if event.wait(timeout=1):
                        print("O Excel foi iniciado!")
                        excel_started = True
                encerrar_processos.excel
                
            if __name__ == "__main__":
                main()

        # Lista todos os processos em execução
        for proc in process_iter(['pid', 'name']):
            try:
                # Verifica se o nome do processo é o SAP Logon
                if 'EXCEL.EXE' in proc.info['name'].upper():
                    # Encerra o processo
                    process = Process(proc.info['pid'])
                    process.terminate()
                    print("Processo do Excel encerrado com sucesso.")
                    return True
            except (NoSuchProcess, AccessDenied, ZombieProcess):
                pass
        print("Processo do Excel não encontrado.")
        return False
    
    def sap_logon(self):
        # Lista todos os processos em execução
        for proc in process_iter(['pid', 'name']):
            try:
                # Verifica se o nome do processo é o SAP Logon
                if 'saplogon.exe' in proc.info['name'].lower():
                    # Encerra o processo
                    process = Process(proc.info['pid'])
                    process.terminate()
                    print("Processo do SAP Logon encerrado com sucesso.")
                    return True
            except (NoSuchProcess, AccessDenied, ZombieProcess):
                pass
        print("Processo do SAP Logon não encontrado.")
        return False