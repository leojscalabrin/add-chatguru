import os
import time
import requests
import pandas as pd
import shutil
import signal
from dotenv import load_dotenv
from typing import Optional

load_dotenv()

stop_processing = False

def signal_handler(sig, frame):
    global stop_processing
    print("\nInterrupção detectada (Ctrl+C). Salvando progresso e parando...")
    stop_processing = True

signal.signal(signal.SIGINT, signal_handler)

def load_config() -> dict:
    config = {
        'server': os.getenv('SERVER'),
        'api_key': os.getenv('KEY'),
        'account_id': os.getenv('ACCOUNT_ID'),
        'excel_file': 'clients.xlsx'
    }
    print("Loaded config:", {k: v if k != 'api_key' else '****' for k, v in config.items()}) 
    return config

def read_excel(file_path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, header=0)
        columns_to_str = range(15) 
        for col in columns_to_str:
            if len(df.columns) > col:
                df.iloc[:, col] = df.iloc[:, col].astype('object').fillna('')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def write_excel(df: pd.DataFrame, file_path: str):
    temp_path = file_path.replace('.xlsx', '_temp.xlsx')
    try:
        with pd.ExcelWriter(temp_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        shutil.move(temp_path, file_path)
        print(f"Excel file updated: {file_path}")
    except KeyboardInterrupt:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        print(f"Error writing to Excel: {e}")

def add_contact(config: dict, name: str, phone_id: str, dialog_id: str, user_id: str, chat_number: str, tags: str = "") -> Optional[str]:
    if stop_processing:
        return "Interrompido pelo usuário"
    
    url = f"https://{config['server']}/api/v1"
    print(f"Sending request to URL: {url}")
    
    payload = {
        "action": "chat_add",
        "key": config['api_key'],
        "account_id": config['account_id'],
        "phone_id": phone_id,
        "name": name,
        "chat_number": chat_number,
        "text": " " # Espaço em branco para não enviar mensagem ao cliente
    }
    
    if dialog_id.strip():
        payload["dialog_id"] = dialog_id
    
    if user_id.strip():
        payload["user_id"] = user_id
        
    # Lógica para incluir tags se existirem
    if tags.strip():
        payload["tags"] = tags.strip()
    
    print("Payload (key hidden):", {k: v if k != 'key' else '****' for k, v in payload.items()})
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    try:
        response = requests.post(url, data=payload, headers=headers)
        print(f"Response status code: {response.status_code}")
        # print(f"Response content: {response.text}") # debug
            
        if response.status_code in [200, 201]:
            data = response.json()
            chat_add_id = data.get("chat_add_id")
            chat_status = data.get("chat_add_status")

            if not chat_add_id:
                # Às vezes o retorno é sucesso mas sem ID se for duplicado imediato
                print(f"Aviso: Sem chat_add_id no retorno para {chat_number}")
                return "Sem chat_add_id (Verificar)"

            print(f"Chat enviado ({chat_status}) - chat_add_id: {chat_add_id}")
            return chat_add_id
        elif response.status_code == 400:
            error_data = response.json()
            error_desc = error_data.get('description', 'Unknown error')
            print(f"Error adding contact {chat_number}: {error_desc}")
            return error_desc
        else:
            print(f"Unexpected status code {response.status_code} for {chat_number}")
            return f"HTTP {response.status_code}"
    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"Request failed for {chat_number}: {e}")
        return str(e)
   

def process_contacts(config: dict):
    if not all([config['server'], config['api_key'], config['account_id']]):
        missing = [k for k, v in config.items() if not v and k != 'excel_file']
        print(f"Missing required config from .env: {', '.join(missing)}")
        return
    
    df = read_excel(config['excel_file'])
    if df.empty:
        print("No data found in Excel file 'clients.xlsx'.")
        return
    
    global stop_processing
    stop_processing = False
    
    for idx in range(len(df)):
        if stop_processing:
            print("Parando processamento...")
            break

        row = df.iloc[idx]
        cadastrado = str(row.iloc[0]).strip().lower() 

        if cadastrado == 'nao':
            name = str(row.iloc[1]).strip()
            if not name: name = "Sem Nome"
            phone_id = str(row.iloc[2]).strip()
            dialog_id = str(row.iloc[3]).strip()
            user_id = str(row.iloc[4]).strip()
            chat_number = str(row.iloc[5]).strip()
            
            # Ler tags da coluna J (index 9) se existir
            tags = ""
            if len(df.columns) > 9:
                val = str(row.iloc[9]).strip()
                if val.lower() not in ['nan', 'none', '']:
                    tags = val

            # Passando tags para a função
            result = add_contact(config, name, phone_id, dialog_id, user_id, chat_number, tags)

            # Lógica de gravação do resultado (Coluna H = index 7)
            if result and not result.startswith("HTTP") and not result.startswith("Sem chat_add_id"):
                # Sucesso (retornou ID)
                df.iloc[idx, 7] = result # Salva o ID na coluna H
                df.iloc[idx, 0] = 'Sim (pendente)'
            else:
                # Erro
                df.iloc[idx, 7] = result # Salva o erro na coluna H para debug
                df.iloc[idx, 0] = 'Erro'

            write_excel(df, config['excel_file'])

            if idx < len(df) - 1 and not stop_processing:
                print("Waiting 1 second before next attempt...")
                time.sleep(1)
        elif cadastrado == 'erro':
            print(f"Skipping row {idx + 1} due to previous error.")
        else:
            print(f"Skipping row {idx + 1}: already processed ({cadastrado}).")
    
    if stop_processing:
        print("Processamento interrompido.")
    else:
        print("Processing complete.")
        
def check_chat_status(config: dict, phone_id: str, chat_add_id: str, max_attempts: int = 10, wait_seconds: int = 2) -> str:
    """Verifica o status de inclusão de um chat no ChatGuru."""
    base_url = f"https://{config['server']}/api/v1"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}

    payload = {
        "action": "chat_add_status",
        "key": config['api_key'],
        "account_id": config['account_id'],
        "phone_id": phone_id,
        "chat_add_id": chat_add_id
    }

    for attempt in range(max_attempts):
        if stop_processing:
            return "Interrompido pelo usuário"

        try:
            resp = requests.post(base_url, data=payload, headers=headers)
            data = resp.json()
            status = data.get("chat_add_status")
            desc = data.get("chat_add_status_description", "")
            print(f"[{attempt+1}] Status: {status} - {desc}")

            if status in ("done", "error"):
                return f"{status} - {desc}"

        except Exception as e:
            print(f"Erro ao checar status: {e}")
            return f"Erro de requisição: {e}"

        time.sleep(wait_seconds)

    return "timeout - sem resposta final"

def check_pending_chats(config: dict):
    df = read_excel(config['excel_file'])
    if df.empty:
        print("Planilha vazia.")
        return

    for idx in range(len(df)):
        if stop_processing:
            break

        # Verifica se foi processado e tem ID (Coluna H = index 7)
        chat_add_id = str(df.iloc[idx, 7]).strip()
        
        # Só checa se tiver ID válido e não for mensagem de erro
        if not chat_add_id or chat_add_id.lower() in ("nan", "none") or " " in chat_add_id:
             # Verifica se é um ID puro (sem espaços, geralmente numérico ou hash)
             # Se for mensagem de erro (ex: "HTTP 400"), pula
             if " " in chat_add_id: continue 

        phone_id = str(df.iloc[idx, 2]).strip()
        name = str(df.iloc[idx, 1]).strip()
        
        # Coluna A precisa estar 'Sim (pendente)' para checar
        status_controle = str(df.iloc[idx, 0]).strip()
        
        if status_controle == 'Sim (pendente)':
            print(f"Checando {name} ({chat_add_id})...")
            result = check_chat_status(config, phone_id, chat_add_id)
            df.iloc[idx, 8] = result  # Coluna I (Status ChatGuru)

            if "done" in result:
                df.iloc[idx, 0] = 'Sim' # Concluído

            write_excel(df, config['excel_file'])
            time.sleep(1)


if __name__ == "__main__":
    config = load_config()
    import sys
    
    # Modo de verificação (segundo passo)
    if len(sys.argv) > 1 and sys.argv[1] == "check":
        check_pending_chats(config)
    else:
        # Modo de envio (padrão)
        try:
            process_contacts(config)
        except KeyboardInterrupt:
            print("\nProcesso interrompido pelo usuário.")
        finally:
            print("Script finalizado.")