import os
import time
import requests
import pandas as pd
from dotenv import load_dotenv
from typing import Optional

load_dotenv()

def load_config() -> dict:
    return {
        'server': os.getenv('SERVER'),
        'api_key': os.getenv('KEY'),
        'account_id': os.getenv('ACCOUNT_ID'),
        'phone_id': os.getenv('PHONE_ID'),
        'excel_file': 'clients.xlsx' 
    }

def read_excel(file_path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, header=0)
        df.columns = ['Cadastrado', 'Nome', 'ID_do_diálogo', 'Número', 'Erro'][:len(df.columns)]
        df['Erro'] = df['Erro'].fillna('')
        df['ID_do_diálogo'] = df['ID_do_diálogo'].fillna('')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def write_excel(df: pd.DataFrame, file_path: str):
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        print(f"Excel file updated: {file_path}")
    except Exception as e:
        print(f"Error writing to Excel: {e}")

def add_contact(config: dict, name: str, dialog_id: str, chat_number: str) -> Optional[str]:
    url = f"https://{config['server']}" 
    
    payload = {
        "key": config['api_key'],
        "account_id": config['account_id'],
        "phone_id": config['phone_id'],
        "name": name,
        "chat_number": chat_number,
        "text": " "
    }
    
    if dialog_id.strip():
        payload["dialog_id"] = dialog_id
    
    headers = {
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.post(url, json=payload, headers=headers)
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
    df = read_excel(config['excel_file'])
    if df.empty:
        print("No data found in Excel file 'clients.xlsx'.")
        return
    
    for idx in range(len(df)):
        row = df.iloc[idx]
        cadastrado = str(row['Cadastrado']).strip().lower()
        
        if cadastrado == 'nao':
            name = str(row['Nome'])
            dialog_id = str(row['ID_do_diálogo'])
            chat_number = str(row['Número'])
            
            error_desc = add_contact(config, name, dialog_id, chat_number)
            
            if error_desc is None:
                df.at[idx, 'Cadastrado'] = 'Sim'
            else:
                df.at[idx, 'Cadastrado'] = 'Erro'
                df.at[idx, 'Erro'] = error_desc
            
            # Save after each update
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
    if not all([config['api_key'], config['account_id'], config['phone_id'], config['server']]):
        print("Missing required config from .env: SERVER, KEY, ACCOUNT_ID, PHONE_ID")
    else:
        process_contacts(config)