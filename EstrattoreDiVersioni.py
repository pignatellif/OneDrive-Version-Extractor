import requests
import json

# Definisci le costanti per l'endpoint dell'API di Microsoft Graph e il token di accesso
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0/"
ACCESS_TOKEN = "il_tuo_token_di_accesso"

def get_user_files():
    url = f"{GRAPH_API_ENDPOINT}/me/drive/root/children"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        files = response.json()["value"]
        print("File disponibili:")
        for file in files:
            print(file["name"])
        return files
    else:
        print(f"Errore durante la richiesta dei file: {response.status_code} - {response.text}")
        return None

def get_drive_item_versions(item_id):
    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{item_id}/versions"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()["value"]
    else:
        print(f"Errore durante la richiesta delle versioni: {response.status_code} - {response.text}")
        return None

def main():
    # Ottieni l'elenco dei file disponibili
    files = get_user_files()
    
    if files:
        # Chiedi all'utente di inserire il titolo del file di cui desidera trovare le versioni
        file_title = input("Inserisci il titolo del file di cui desideri trovare le versioni: ")
        
        # Cerca il file con il titolo inserito dall'utente
        selected_file = next((file for file in files if file["name"] == file_title), None)
        
        if selected_file:
            item_id = selected_file["id"]
            versions = get_drive_item_versions(item_id)
            
            if versions:
                print("Versioni del file:")
                for version in versions:
                    print(f"Versione: {version['id']}, Modificato da: {version['lastModifiedBy']['user']['displayName']}, Data e Ora: {version['lastModifiedDateTime']}, Dimensione: {version['size']} bytes")
            else:
                print("Non è stato possibile ottenere le versioni del file.")
        else:
            print("Il file specificato non è stato trovato.")
    else:
        print("Non sono stati trovati file disponibili.")

if __name__ == "__main__":
    main()
