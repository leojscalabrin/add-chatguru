import os
import requests
from dotenv import load_dotenv

load_dotenv()

SERVER = os.getenv('SERVER')
API_KEY = os.getenv('KEY')
ACCOUNT_ID = os.getenv('ACCOUNT_ID')

PHONE_ID = "PHONE_ID" 

def get_chatguru_tags():
    if not all([SERVER, API_KEY, ACCOUNT_ID, PHONE_ID]):
        print("Erro: Verifique se o .env e o PHONE_ID no script estão preenchidos.")
        return

    url = f"https://{SERVER}/api/v1"
    
    payload = {
        "action": "chat_tags_list",
        "key": API_KEY,
        "account_id": ACCOUNT_ID,
        "phone_id": PHONE_ID
    }
    
    print(f"Consultando tags no servidor {SERVER} usando o aparelho {PHONE_ID}...")
    
    try:
        response = requests.post(url, data=payload)
        
        if response.status_code == 200:
            data = response.json()
            
            # Tenta localizar a lista de tags dentro da resposta
            tags = []
            if isinstance(data, list):
                tags = data
            elif isinstance(data, dict):
                tags = data.get('data', data.get('tags', []))

            if tags:
                print("\n=== SUCESSO! COPIE OS IDs ABAIXO PARA A COLUNA J ===")
                print(f"{'NOME DA TAG':<30} | {'ID (USE NO EXCEL)'}")
                print("-" * 70)
                
                for tag in tags:
                    t_id = tag.get('id') or tag.get('_id') or tag.get('tag_id')
                    t_name = tag.get('name') or tag.get('tag_name') or "Sem Nome"
                    t_color = tag.get('color', '')
                    
                    if t_id:
                        print(f"{t_name:<30} | {t_id}")
            else:
                print("\nNenhuma tag encontrada na lista.")
                print(f"Resposta bruta: {data}")
        else:
            print(f"Erro na requisição: Status {response.status_code}")
            print(f"Mensagem: {response.text}")
            
    except Exception as e:
        print(f"Erro ao conectar: {e}")

if __name__ == "__main__":
    if PHONE_ID == "PHONE_ID":
        print("PHONE_ID invalido")
    else:
        get_chatguru_tags()