import os
import time
import requests
import pandas as pd
from dotenv import load_dotenv
from typing import Optional

load_dotenv()

def load_config() -> dict:
    config = {
        'server': os.getenv('SERVER'),
        'api_key': os.getenv('KEY'),
        'account_id': os.getenv('ACCOUNT_ID'),
        'phone_id': os.getenv('PHONE_ID'),
        'excel_file': 'clients.xlsx'
    }
    print("Loaded config:", {k: v if k != 'api_key' else '****' for k, v in config.items()})  # Hide api_key for security
    return config

def read_excel(file_path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, header=0)
        if len(df.columns) > 2:
            df.iloc[:, 2] = df.iloc[:, 2].astype('object')
        if len(df.columns) > 4:
            df.iloc[:, 4] = df.iloc[:, 4].astype('object')
        df.iloc[:, 2] = df.iloc[:, 2].fillna('')
        df.iloc[:, 4] = df.iloc[:, 4].fillna('')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def write_excel(df: pd.DataFrame, file_path: str):
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        print(f"Excel file updated: {file_path}")
    except Exception as e:
        print(f"Error writing to Excel: {e}")

def add_contact(config: dict, name: str, dialog_id: str, chat_number: str) -> Optional[str]:
    url = f"https://{config['server']}/api/v1"
    print(f"Sending request to URL: {url}")
    
    payload = {
        "action": "chat_add",
        "key": config['api_key'],
        "account_id": config['account_id'],
        "phone_id": config['phone_id'],
        "name": name,
        "chat_number": chat_number,
        "text": " "
    }
    
    if dialog_id.strip():
        payload["dialog_id"] = dialog_id
    
    print("Payload (key hidden):", {k: v if k != 'key' else '****' for k, v in payload.items()})
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    try:
        response = requests.post(url, data=payload, headers=headers)
        print(f"Response status code: {response.status_code}")
        print(f"Response content: {response.text}")
        if response.status_code in [200, 201]:
            print(f"Contact added successfully: {chat_number}")
            return None
        elif response.status_code == 400:
            error_data = response.json()
            error_desc = error_data.get('description', 'Unknown error')
            print(f"Error adding contact {chat_number}: {error_desc}")
            return error_desc
        else:
            print(f"Unexpected status code {response.status_code} for {chat_number}")
            return f"HTTP {response.status_code}"
    except Exception as e:
        print(f"Request failed for {chat_number}: {e}")
        return str(e)

def process_contacts(config: dict):
    if not all([config['server'], config['api_key'], config['account_id'], config['phone_id']]):
        missing = [k for k, v in config.items() if not v and k != 'excel_file']
        print(f"Missing required config from .env: {', '.join(missing)}")
        return
    
    df = read_excel(config['excel_file'])
    if df.empty:
        print("No data found in Excel file 'clients.xlsx'.")
        return
    
    for idx in range(len(df)):
        row = df.iloc[idx]
        cadastrado = str(row.iloc[0]).strip().lower()  # Coluna A (Cadastrado)
        
        if cadastrado == 'nao':
            name = str(row.iloc[1])  # Coluna B (Nome)
            dialog_id = str(row.iloc[2])  # Coluna C (ID do diálogo)
            chat_number = str(row.iloc[3])  # Coluna D (Número)
            
            error_desc = add_contact(config, name, dialog_id, chat_number)
            
            if error_desc is None:
                df.at[idx, df.columns[0]] = 'Sim'
            else:
                df.at[idx, df.columns[0]] = 'Erro' 
                df.at[idx, df.columns[4]] = error_desc 

            write_excel(df, config['excel_file'])
            
            # Wait 5 seconds before next attempt
            if idx < len(df) - 1:
                print("Waiting 5 seconds before next attempt...")
                time.sleep(5)
        elif cadastrado == 'erro':
            print(f"Skipping row {idx + 1} due to previous error.")
        else:
            print(f"Skipping row {idx + 1}: already processed ({cadastrado}).")
    
    print("Processing complete.")

if __name__ == "__main__":
    config = load_config()
    process_contacts(config)