import asyncio
import configparser
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog, ttk
import os
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

        # Impostazione di uno stile per i widget
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Helvetica', 12), padding=5)
        style.configure('TLabel', font=('Helvetica', 12))
        style.configure('TTreeview', font=('Helvetica', 10))

        # Creazione di un frame principale
        main_frame = ttk.Frame(root, padding="10 10 10 10")
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Aggiunta di pulsanti con icone e stile migliorato
        self.back_button = ttk.Button(main_frame, text="◀ Back", command=self.go_back)
        self.back_button.grid(row=0, column=0, sticky=tk.W, pady=10)

        self.request_activity_button = ttk.Button(main_frame, text="🔄 Request Activity", command=self.request_activity)
        self.request_activity_button.grid(row=0, column=1, sticky=tk.E, pady=10)

        self.file_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=2)
        self.file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.N, tk.S, tk.E, tk.W))
        main_frame.rowconfigure(1, weight=1)

        self.load_files_button = ttk.Button(main_frame, text="📁 Load Files", command=self.display_list_files)
        self.load_files_button.grid(row=2, column=0, columnspan=2, pady=10)

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
                frame = ttk.Frame(self.file_frame, padding="5 5 5 5", relief=tk.GROOVE, borderwidth=1)
                frame.pack(fill=tk.X, pady=2)

                label = ttk.Label(frame, text=item['name'], anchor='w')
                label.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

                if item.get('folder'):  # Se è una cartella, aggiungi il pulsante "Apri"
                    button = ttk.Button(frame, text="📂 Open", command=lambda item=item: self.open_folder(item['id']))
                    button.pack(side=tk.RIGHT, padx=5, pady=5)
                else:  # Se è un file, aggiungi il pulsante "Visualizza versioni"
                    button = ttk.Button(frame, text="📝 View Versions", command=lambda item=item: self.open_versions_window(item['id']))
                    button.pack(side=tk.RIGHT, padx=5, pady=5)

        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except requests.exceptions.RequestException as err:
            messagebox.showerror("Request Error", f'HTTP Error: {err}')

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
        
    def request_activity(self):
        ActivityWindow(self.graph, self.token)

 
class VersionsWindow(tk.Toplevel):
    def __init__(self, parent, file_id, graph, token):
        super().__init__(parent.root)
        self.file_id = file_id
        self.graph = graph
        self.token = token

        self.title("File Versions")
        self.geometry("800x600")

        self.versions_frame = ttk.Frame(self, padding="10 10 10 10")
        self.versions_frame.pack(fill=tk.BOTH, expand=True)

        self.button_frame = ttk.Frame(self)
        self.button_frame.pack(fill=tk.X, pady=5)

        self.export_csv_button = ttk.Button(self.button_frame, text="💾 Export to CSV", command=self.export_to_csv)
        self.export_csv_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.download_all_versions_button = ttk.Button(self.button_frame, text="📥 Download All Versions", command=self.download_all_versions)
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
                frame = ttk.Frame(self.versions_frame, padding="5 5 5 5", relief=tk.GROOVE, borderwidth=1)
                frame.pack(fill=tk.X, pady=2)

                modified_time_str = version.get('lastModifiedDateTime', 'N/A')
                try:
                    modified_time = parse(modified_time_str)
                except ValueError:
                    continue

                modified_time_localized = modified_time.astimezone(local_tz)
                modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

                label = ttk.Label(frame, text=f"Version ID: {version['id']}, Last Modified: {modified_time_str_with_tz}, Size: {version['size']} bytes", anchor='w')
                label.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

                download_button = ttk.Button(frame, text="⬇ Download", command=lambda version_id=version['id']: self.download_version(version_id))
                download_button.pack(side=tk.RIGHT, padx=5, pady=5)
                
                restore_button = ttk.Button(frame, text="🔄 Restore", command=lambda version_id=version['id']: self.restore_version(version_id))
                restore_button.pack(side=tk.RIGHT, padx=5, pady=5)

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
        
    def restore_version(self, version_id):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{self.file_id}/versions/{version_id}/restoreVersion'

        try:
            response = requests.post(url, headers=headers)
            response.raise_for_status()

            messagebox.showinfo("Success", f"Version {version_id} restored successfully.")

        except requests.exceptions.HTTPError as err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {err}')
        except Exception as err:
            messagebox.showerror("Error", f'An error occurred: {err}')

class ActivityWindow(tk.Toplevel):
    def __init__(self, graph, token):
        super().__init__()
        self.graph = graph
        self.token = token
        self.title("OneDrive Activities")
        self.geometry("800x600")

        self.activities_frame = ttk.Frame(self, padding="10 10 10 10")
        self.activities_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.create_table()
        self.load_activities()

    def create_table(self):
        self.tree = ttk.Treeview(self.activities_frame, columns=('Name', 'Action', 'Time'), show='headings', height=15)
        self.tree.heading('Name', text='File/Folder Name')
        self.tree.heading('Action', text='Action')
        self.tree.heading('Time', text='Time (CEST)')
        
        self.tree.column('Name', anchor=tk.W, width=300)
        self.tree.column('Action', anchor=tk.CENTER, width=100)
        self.tree.column('Time', anchor=tk.CENTER, width=150)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Aggiunta di una scrollbar verticale
        scrollbar = ttk.Scrollbar(self.activities_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=scrollbar.set)

    def load_activities(self):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Accept': 'application/json'
        }
        
        url = 'https://graph.microsoft.com/v1.0/me/drive/root/delta'
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            changes = response.json()

            # Clear previous entries
            for row in self.tree.get_children():
                self.tree.delete(row)

            if 'value' in changes:
                cest_tz = pytz.timezone('Europe/Berlin')  # CEST timezone (Europe/Berlin includes CET/CEST adjustments)
                for change in changes['value']:
                    name = change.get('name', 'N/A')
                    change_action = "Modified" if change.get('lastModifiedDateTime') else "Created"
                    if change.get('deleted'):
                        change_action = "Deleted"

                    modified_time_str = change.get('lastModifiedDateTime', 'N/A')
                    try:
                        modified_time = parse(modified_time_str).astimezone(cest_tz).strftime('%Y-%m-%d %H:%M:%S %Z')
                    except ValueError:
                        modified_time = 'N/A'

                    self.tree.insert('', 'end', values=(name, change_action, modified_time))
            else:
                self.tree.insert('', 'end', values=("No activities found", "", ""))

        except requests.exceptions.HTTPError as http_err:
            messagebox.showerror("HTTP Error", f'HTTP Error: {http_err}\nResponse: {response.text}')
        except requests.exceptions.RequestException as req_err:
            messagebox.showerror("Request Error", f'Request Error: {req_err}')
        except Exception as err:
            messagebox.showerror("Error", f'An unexpected error occurred: {err}')

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
