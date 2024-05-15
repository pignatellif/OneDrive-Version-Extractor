import requests
from datetime import datetime

class OneDriveVersionExtractor:
    def __init__(self, access_token):
        self.access_token = access_token

    def list_versions(self, file_id):
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions"
        headers = {
            "Authorization": f"Bearer {self.access_token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            versions = response.json()["value"]
            return versions
        else:
            print("Failed to retrieve versions.")
            return []

    def download_version(self, file_id, version_id, save_path):
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content"
        headers = {
            "Authorization": f"Bearer {self.access_token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            with open(save_path, "wb") as f:
                f.write(response.content)
            print(f"Version {version_id} downloaded successfully.")
        else:
            print(f"Failed to download version {version_id}.")

if __name__ == "__main__":
    access_token = "YOUR_ACCESS_TOKEN_HERE"
    file_id = "YOUR_FILE_ID_HERE"

    extractor = OneDriveVersionExtractor(access_token)
    versions = extractor.list_versions(file_id)
    print("Available versions:")
    for version in versions:
        version_id = version["id"]
        last_modified = datetime.strptime(version["lastModifiedDateTime"], "%Y-%m-%dT%H:%M:%SZ")
        print(f"Version: {version_id}, Last Modified: {last_modified}")

    # Seleziona una versione per scaricarla
    version_to_download = versions[0]  # Modifica questa riga in base alla tua logica di selezione
    version_id_to_download = version_to_download["id"]
    save_path = "downloaded_version.txt"  # Modifica il percorso in cui salvare il file scaricato
    extractor.download_version(file_id, version_id_to_download, save_path)
