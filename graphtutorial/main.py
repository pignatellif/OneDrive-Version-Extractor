import asyncio
import configparser
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog, ttk
import os
import webbrowser
import requests
from graph import Graph
from datetime import datetime
from pytz import timezone
import pytz
import csv
from tzlocal import get_localzone
from dateutil.parser import parse

class OneDriveApp:
    def __init__(self, root, graph, token):
        self.root = root
        self.graph = graph
        self.token = token
        self.current_folder_id = 'root'  # Inizialmente siamo nella radice
        self.folder_stack = []  # Stack per tenere traccia del percorso

        self.root.title("OneDrive Management")
        self.root.geometry("800x600")

        self.back_button = ttk.Button(root, text="Back", command=self.go_back)
        self.back_button.pack(side=tk.TOP, pady=10)

        self.file_frame = ttk.Frame(root)
        self.file_frame.pack(fill=tk.BOTH, expand=True)

        self.load_files_button = ttk.Button(root, text="Load Files", command=self.display_list_files)
        self.load_files_button.pack(side=tk.BOTTOM, pady=10)

    async def greet_user(self):
        user = await self.graph.get_user()
        if user:
            messagebox.showinfo("Welcome", f'Hello, {user.display_name}\nEmail: {user.mail or user.user_principal_name}')

    def display_list_files(self):
        headers = {
            'Authorization': f'Bearer {self.token}'
        }

        try:
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.current_folder_id}/children'
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            root_items = response.json()

            for widget in self.file_frame.winfo_children():
                widget.destroy()

            for item in root_items['value']:
                frame = ttk.Frame(self.file_frame)
                frame.pack(fill=tk.X, pady=2)

                label = ttk.Label(frame, text=item['name'])
                label.pack(side=tk.LEFT, padx=5, pady=5)

                if item.get('folder'):  # Se è una cartella, aggiungi il pulsante "Apri"
                    button = ttk.Button(frame, text="Open", command=lambda item=item: self.open_folder(item['id']))
                    button.pack(side=tk.RIGHT, padx=5, pady=5)
                else:  # Se è un file, aggiungi il pulsante "Visualizza versioni"
                    button = ttk.Button(frame, text="View Versions", command=lambda item=item: self.open_versions_window(item['id']))
                    button.pack(side=tk.RIGHT, padx=5, pady=5)

        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except requests.exceptions.RequestException as err:
            messagebox.showerror("Request Error", f'Request Error: {err}')

    def open_folder(self, folder_id):
        self.folder_stack.append(self.current_folder_id)  # Aggiungi la cartella corrente allo stack
        self.current_folder_id = folder_id
        self.display_list_files()

    def go_back(self):
        if self.folder_stack:
            self.current_folder_id = self.folder_stack.pop()
            self.display_list_files()
        else:
            messagebox.showinfo("Info", "You are at the root folder.")

    def open_versions_window(self, file_id):
        VersionsWindow(self, file_id, self.graph, self.token)

 
class VersionsWindow(tk.Toplevel):
    def __init__(self, parent, file_id, graph, token):
        super().__init__(parent.root)
        self.file_id = file_id
        self.graph = graph
        self.token = token
 
        self.title("File Versions")
        self.geometry("800x600")
 
        self.versions_frame = ttk.Frame(self)
        self.versions_frame.pack(fill=tk.BOTH, expand=True)
 
        self.export_csv_button = ttk.Button(self, text="Export to CSV", command=self.export_to_csv)
        self.export_csv_button.pack(side=tk.LEFT, padx=5, pady=5)
 
        self.download_all_versions_button = ttk.Button(self, text="Download All Versions", command=self.download_all_versions)
        self.download_all_versions_button.pack(side=tk.RIGHT, padx=5, pady=5)
 
        self.display_file_versions()
 
    def display_file_versions(self):
        headers = {
            'Authorization': f'Bearer {self.token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions'
 
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            versions = response.json()
 
            for widget in self.versions_frame.winfo_children():
                widget.destroy()
 
            local_tz = get_localzone()
 
            for version in versions['value']:
                frame = ttk.Frame(self.versions_frame)
                frame.pack(fill=tk.X, pady=2)
 
                modified_time_str = version.get('lastModifiedDateTime', 'N/A')
                try:
                    modified_time = parse(modified_time_str)
                except ValueError:
                    continue
 
                modified_time_localized = modified_time.astimezone(local_tz)
                modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')
 
                label = ttk.Label(frame, text=f"Version ID: {version['id']}, Last Modified: {modified_time_str_with_tz}, Size: {version['size']} bytes")
                label.pack(side=tk.LEFT, padx=5, pady=5)
 
                download_button = ttk.Button(frame, text="Download", command=lambda version_id=version['id']: self.download_version(version_id))
                download_button.pack(side=tk.RIGHT, padx=5, pady=5)
 
        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except Exception as err:
            messagebox.showerror("Error", f'An error occurred: {err}')
 
    def export_to_csv(self):
        save_directory = filedialog.askdirectory()
        if not save_directory:
            messagebox.showinfo("Cancelled", "Export cancelled.")
            return
 
        headers = {
            'Authorization': f'Bearer {self.token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions'
 
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            versions = response.json()
 
            csv_filename = f"{self.file_id}_versions.csv"
            csv_path = os.path.join(save_directory, csv_filename)
 
            with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(['Version ID', 'Download URL', 'Last Modified', 'Size', 'Last Modified By'])
 
                for version in versions['value']:
                    version_id = version.get('id', 'N/A')
                    download_url = version.get('@microsoft.graph.downloadUrl', 'N/A')
                    modified_time_str = version.get('lastModifiedDateTime', 'N/A')
                    modified_time = parse(modified_time_str).strftime('%Y-%m-%d %H:%M:%S %Z') if modified_time_str != 'N/A' else 'N/A'
                    file_size = str(version.get('size', 'N/A')) + " bytes"
                    last_modified_by = version.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'N/A')
 
                    writer.writerow([version_id, download_url, modified_time, file_size, last_modified_by])
 
            messagebox.showinfo("Success", f"CSV exported successfully to {csv_path}")
 
        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except Exception as err:
            messagebox.showerror("Error", f'An error occurred: {err}')
 
    def download_version(self, version_id):
        save_directory = filedialog.askdirectory()
        if not save_directory:
            messagebox.showinfo("Cancelled", "Download cancelled.")
            return
 
        headers = {
            'Authorization': f'Bearer {self.token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions/{version_id}/content'
 
        try:
            response = requests.get(url, headers=headers, allow_redirects=True)
            response.raise_for_status()
 
            original_name_response = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}', headers=headers)
            original_name_response.raise_for_status()
            original_name = original_name_response.json()['name']
            original_extension = os.path.splitext(original_name)[1]
            file_name = f"{os.path.splitext(original_name)[0]}_{version_id}{original_extension}"
            save_path = os.path.join(save_directory, file_name)
 
            with open(save_path, 'wb') as file:
                file.write(response.content)
 
            messagebox.showinfo("Success", f"File downloaded successfully as {save_path}")
 
        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except Exception as err:
            messagebox.showerror("Error", f'An error occurred: {err}')
 
    def download_all_versions(self):
        save_directory = filedialog.askdirectory()
        if not save_directory:
            messagebox.showinfo("Cancelled", "Download cancelled.")
            return
 
        headers = {
            'Authorization': f'Bearer {self.token}'
        }
        url_versions = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions'
 
        try:
            response_versions = requests.get(url_versions, headers=headers)
            response_versions.raise_for_status()
            versions = response_versions.json()
 
            original_name_response = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}', headers=headers)
            original_name_response.raise_for_status()
            original_name = original_name_response.json()['name']
            original_extension = os.path.splitext(original_name)[1]
 
            for version in versions['value']:
                version_id = version['id']
                url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions/{version_id}/content'
                file_name = f"{os.path.splitext(original_name)[0]}_{version_id}{original_extension}"
                save_path = os.path.join(save_directory, file_name)
 
                response = requests.get(url, headers=headers, allow_redirects=True)
                response.raise_for_status()
 
                with open(save_path, 'wb') as file:
                    file.write(response.content)
 
            messagebox.showinfo("Success", "All versions downloaded successfully.")
 
        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except Exception as err:
            messagebox.showerror("Error", f'An error occurred: {err}')
 
async def main():
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
 
    azure_settings = config['azure']
    graph = Graph(azure_settings)
    token = await graph.get_user_token()
 
    if not token:
        return
 
    root = tk.Tk()
    app = OneDriveApp(root, graph, token)
 
    await app.greet_user()
 
    root.mainloop()
 
if __name__ == "__main__":
    asyncio.run(main())
