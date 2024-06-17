import asyncio
import configparser
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
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

# Greet user with their name and email
async def greet_user(graph: Graph, output_text):
    user = await graph.get_user()
    if user:
        output_text.insert(tk.END, f'Hello, {user.display_name}\n')
        output_text.insert(tk.END, f'Email: {user.mail or user.user_principal_name}\n\n')

# Recursive function to print items in a hierarchical structure
def print_items(token, items, level=0, output_text=None):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    for item in items['value']:
        message = ' ' * level * 4 + '- ' + f"[{item['id']}] {item['name']}\n"
        if output_text:
            output_text.insert(tk.END, message)
        else:
            print(message)

        if 'folder' in item:
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item['id']}/children"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            sub_items = response.json()
            print_items(token, sub_items, level + 1, output_text)

# Display a list of files and folders in OneDrive
def display_list_files(graph: Graph, token, output_text):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    try:
        response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
        response.raise_for_status()
        root_items = response.json()

        print_items(token, root_items, output_text=output_text)

    except requests.exceptions.HTTPError as err:
        message = f'HTTP Error: {err}\n'
        output_text.insert(tk.END, message)
    except requests.exceptions.RequestException as err:
        message = f'Request Error: {err}\n'
        output_text.insert(tk.END, message)

# Validate if a file ID exists in OneDrive
def validate_file_id(token, file_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

# Validate if a version ID exists for a given file in OneDrive
def validate_version_id(token, file_id, version_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

# Display versions of a file in OneDrive with option to export to CSV
def display_file_versions(graph: Graph, token, file_id, output_text):
    try:
        headers = {
            'Authorization': f'Bearer {token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions'

        response = requests.get(url, headers=headers)
        response.raise_for_status()
        versions = response.json()

        # Get local timezone of the device
        local_tz = get_localzone()

        # Print file versions information to console
        for version in versions['value']:
            modified_time_str = version.get('lastModifiedDateTime', 'N/A')
            try:
                modified_time = parse(modified_time_str)
            except ValueError as e:
                message = f"Errore nel parsing della data: {e}\n"
                output_text.insert(tk.END, message)
                continue

            # Directly convert the datetime object to the local timezone
            modified_time_localized = modified_time.astimezone(local_tz)
            modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

            output_text.insert(tk.END, f"Version ID: {version['id']}\n")
            output_text.insert(tk.END, f"Last Modified: {modified_time_str_with_tz}\n")
            output_text.insert(tk.END, f"Size: {version['size']} bytes\n\n")

        # Ask user if they want to export to CSV
        export_csv = messagebox.askquestion("Export to CSV", "Do you want to export the versions to a CSV file?")
        if export_csv == 'yes':
            csv_filename = f"{file_id}_versions.csv"
            directory = filedialog.askdirectory()

            if directory is None or directory == '':
                output_text.insert(tk.END, "Canceled\n")
                return

            if not os.path.exists(directory):
                os.makedirs(directory)

            csv_path = os.path.join(directory, csv_filename)
            with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(['Version ID', 'Download URL', 'Last Modified', 'Size', 'Last Modified By'])

                for version in versions['value']:
                    version_id = version.get('id', 'N/A')
                    download_url = version.get('@microsoft.graph.downloadUrl', 'N/A')
                    modified_time_str = version.get('lastModifiedDateTime', 'N/A')
                    modified_time = parse(modified_time_str).strftime('%Y-%m-%d %H:%M:%S %Z') if modified_time_str != 'N/A' else 'N/A'
                    file_size = str(version.get('size', 'N/A')) + "bytes"
                    last_modified_by = version.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'N/A')

                    writer.writerow([version_id, download_url, modified_time, file_size, last_modified_by])

            output_text.insert(tk.END, f"File {csv_path} containing versions metadata exported successfully.\n")
        elif export_csv == 'no':
            output_text.insert(tk.END, "Not exporting to CSV.\n")
        else:
            output_text.insert(tk.END, "Invalid choice. Not exporting to CSV.\n")
    except requests.exceptions.HTTPError as e:
        output_text.insert(tk.END, f"HTTP Error: {e}\n")
    except Exception as e:
        output_text.insert(tk.END, f"An unexpected error occurred: {e}\n")

def download_file_version(token, file_id, version_id, save_directory, output_text):
    try:
        if save_directory is None or save_directory == '':
            output_text.insert(tk.END, "Canceled\n")
            return

        headers = {
            'Authorization': f'Bearer {token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content'

        # Get the original file name and extension
        url_item = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}'
        response_item = requests.get(url_item, headers=headers)
        response_item.raise_for_status()
        item = response_item.json()
        original_name = item['name']
        original_extension = os.path.splitext(original_name)[1]

        # Create directory if it does not exist
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)

        # Construct the file name with version id
        file_name = f"{os.path.splitext(original_name)[0]}_{version_id}{original_extension}"
        save_path = os.path.join(save_directory, file_name)

        response = requests.get(url, headers=headers, allow_redirects=True)
        response.raise_for_status()

        with open(save_path, 'wb') as file:
            file.write(response.content)

        output_text.insert(tk.END, f"File version downloaded successfully as {save_path}\n")

    except requests.exceptions.HTTPError as e:
        output_text.insert(tk.END, f"HTTP Error occurred: {e}\n")
    except requests.exceptions.RequestException as e:
        output_text.insert(tk.END, f"Request Exception occurred: {e}\n")
    except Exception as e:
        output_text.insert(tk.END, f"An unexpected error occurred: {e}\n")

# Download all versions of a file from OneDrive
def download_all_file_versions(token, file_id, save_directory, output_text):
    try:
        if save_directory is None or save_directory == '':
            output_text.insert(tk.END, "Canceled\n")
            return

        headers = {
            'Authorization': f'Bearer {token}'
        }
        url_versions = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions'
        response_versions = requests.get(url_versions, headers=headers)
        response_versions.raise_for_status()
        versions = response_versions.json()

        # Get the original file name and extension
        url_item = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}'
        response_item = requests.get(url_item, headers=headers)
        response_item.raise_for_status()
        item = response_item.json()
        original_name = item['name']
        original_extension = os.path.splitext(original_name)[1]

        # Create directory if it does not exist
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)

        for version in versions['value']:
            version_id = version['id']
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content'
            file_name = f"{os.path.splitext(original_name)[0]}_{version_id}{original_extension}"
            save_path = os.path.join(save_directory, file_name)

            response = requests.get(url, headers=headers, allow_redirects=True)
            response.raise_for_status()

            with open(save_path, 'wb') as file:
                file.write(response.content)

            output_text.insert(tk.END, f"Downloaded version {version_id} as {save_path}\n")
    except requests.exceptions.HTTPError as e:
        output_text.insert(tk.END, f"HTTP Error occurred: {e}\n")
    except requests.exceptions.RequestException as e:
        output_text.insert(tk.END, f"Request Exception occurred: {e}\n")
    except ValueError as e:
        output_text.insert(tk.END, f"Value Error: {e}\n")
    except Exception as e:
        output_text.insert(tk.END, f"An unexpected error occurred: {e}\n")

# Monitor activities and changes in OneDrive
def monitor_onedrive_activities(token, output_text):
    url_delta = 'https://graph.microsoft.com/v1.0/me/drive/root/delta'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    try:
        response_delta = requests.get(url_delta, headers=headers)
        response_delta.raise_for_status()
        changes = response_delta.json()

        local_tz = get_localzone()

        output_text.insert(tk.END, "Changes in OneDrive:\n")
        output_text.insert(tk.END, "=" * 171 + "\n")
        output_text.insert(tk.END, "| {:<40} | {:<40} | {:<20} | {:<25} | {:<30} |\n".format("Filename", "File ID", "Action", "Modified By", "Modified Time"))
        output_text.insert(tk.END, "-" * 171 + "\n")

        for item in changes.get('value', []):
            filename = item.get('name', 'N/A')
            file_id = item.get('id', 'N/A')
            modified_time_str = item.get('lastModifiedDateTime', 'N/A')
            
            try:
                modified_time = parse(modified_time_str)
            except ValueError as e:
                message = f"Errore nel parsing della data: {e}\n"
                output_text.insert(tk.END, message)
                continue

            # Directly convert the datetime object to the local timezone
            modified_time_localized = modified_time.astimezone(local_tz)
            modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

            modified_by = item.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'N/A')
            action = 'Deleted' if 'deleted' in item else 'Modified'
            output_text.insert(tk.END, "| {:<40} | {:<40} | {:<20} | {:<25} | {:<30} |\n".format(filename[:40], file_id[:40], action, modified_by[:25], modified_time_str_with_tz))

        output_text.insert(tk.END, "=" * 171 + "\n")

    except requests.exceptions.HTTPError as err:
        message = f'HTTP Error: {err}\n'
        output_text.insert(tk.END, message)
    except Exception as err:
        message = f'An error occurred: {err}\n'
        output_text.insert(tk.END, message)

# Restore a specific version of a file in OneDrive
def restore_file_version(graph: Graph, token, file_id, version_id, output_text):
    try:
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/restoreVersion'

        response = requests.post(url, headers=headers)
        response.raise_for_status()

        output_text.insert(tk.END, f"Version {version_id} of file {file_id} restored successfully.\n")

    except requests.exceptions.HTTPError as e:
        output_text.insert(tk.END, f"HTTP Error: {e}\n")
    except requests.exceptions.RequestException as e:
        output_text.insert(tk.END, f"Request Exception: {e}\n")
    except Exception as e:
        output_text.insert(tk.END, f"An unexpected error occurred: {e}\n")

async def main():
    print('OneDrive-Version-Extractor\n')

    # Load settings from configuration files
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    # Initialize Graph object with Azure settings
    graph = Graph(azure_settings)

    # Open a web page in the browser
    webbrowser.open('https://microsoft.com/devicelogin')

    # Initialize the Tkinter root window
    root = tk.Tk()
    root.title("OneDrive Management")
    root.geometry("1550x700")  

    # Create a Text widget to display output with increased width and use a fixed-width font
    output_text = tk.Text(root, width=256, height=30, font=('Courier', 15))
    output_text.pack()

    # Function to close the program
    def exit_program():
        root.destroy()

    # Greet the user and get the token for API requests
    await greet_user(graph, output_text)
    token = await graph.get_user_token()

    # Function to display list of files
    def display_list_files_wrapper():
        display_list_files(graph, token, output_text)

    # Function to display file versions
    def display_file_versions_wrapper():
        file_id = simpledialog.askstring("File ID", "Enter File ID:")
        if file_id is None:
            output_text.insert(tk.END, "Canceled\n")
        else:
            display_file_versions(graph, token, file_id, output_text)

    # Function to download file version
    def download_file_version_wrapper():
        file_id = simpledialog.askstring("File ID", "Enter File ID:")
        if file_id is None:
            output_text.insert(tk.END, "Canceled\n")
            return

        version_id = simpledialog.askstring("Version ID", "Enter Version ID:")
        if version_id is None:
            output_text.insert(tk.END, "Canceled\n")
            return

        save_directory = filedialog.askdirectory()
        if save_directory is None or save_directory == '':
            output_text.insert(tk.END, "Canceled\n")
            return

        if version_id == 'all':
            download_all_file_versions(token, file_id, save_directory, output_text)
        else:
            download_file_version(token, file_id, version_id, save_directory, output_text)

    # Function to monitor OneDrive activities
    def monitor_onedrive_activities_wrapper():
        monitor_onedrive_activities(token, output_text)

    # Function to restore file version
    def restore_file_version_wrapper():
        file_id = simpledialog.askstring("File ID", "Enter File ID:")
        if file_id is None:
            output_text.insert(tk.END, "Canceled\n")
            return

        version_id = simpledialog.askstring("Version ID", "Enter Version ID:")
        if version_id is None:
            output_text.insert(tk.END, "Canceled\n")
            return

        restore_file_version(graph, token, file_id, version_id, output_text)

    # Create buttons for each function
    button_exit = tk.Button(root, text="Exit", command=exit_program)
    button_exit.pack()

    button_display_list_files = tk.Button(root, text="Display List Files", command=display_list_files_wrapper)
    button_display_list_files.pack()

    button_display_file_versions = tk.Button(root, text="Display File Versions", command=display_file_versions_wrapper)
    button_display_file_versions.pack()

    button_download_file_version = tk.Button(root, text="Download File Version", command=download_file_version_wrapper)
    button_download_file_version.pack()

    button_monitor_onedrive_activities = tk.Button(root, text="Monitor OneDrive Activities", command=monitor_onedrive_activities_wrapper)
    button_monitor_onedrive_activities.pack()

    button_restore_file_version = tk.Button(root, text="Restore File Version", command=restore_file_version_wrapper)
    button_restore_file_version.pack()

    root.mainloop()

# Entry point to run the main function
asyncio.run(main())