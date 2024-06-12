import asyncio
import requests
import configparser
import os
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph

async def main():
    print('OneDrive-Version-Extractor\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph = Graph(azure_settings)

    await greet_user(graph)

    # Get the token once
    token = await graph.get_user_token()

    choice = -1

    while choice != 0:
        print('Please choose one of the following options:')
        print('0. Exit')
        print('1. Display list files')
        print('2. Display file versions')
        print('3. Download a file version')
        print('4. Updates on OneDrive')
        print('5. Restore a version')

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            match choice:
                case 0:
                    #Exit
                    print('Goodbye...')
                case 1:
                    #Display list files
                    await display_list_files(graph, token)
                case 2:
                    #Display file versions
                    file_id = input('Enter the file ID of the file you want to see the versions (type cancel for exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        await display_file_versions(graph, token, file_id)
                case 3:
                    #Download a file version
                    file_id = input('Enter the file ID of the file you want to download (type cancel for exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        if not await validate_file_id(token, file_id):
                            print("Invalid file ID. Please try again.")
                        else:
                            choice_2 = input('Do you want to download all the versions of the file? (yes/no/cancel for exit): ')
                            match choice_2:
                                case "no":
                                    version_id = input('Enter the version ID of the file you want to download (type cancel for exit): ')
                                    if version_id == "cancel":
                                        continue
                                    else:
                                        if not await validate_version_id(token, file_id, version_id):
                                            print("Invalid version ID. Please try again.")
                                        else:
                                            save_directory = input('Enter the directory path to save the file: ')
                                            file_name = input('Enter the file name with extension: ')
                                            await download_file_version(token, file_id, version_id, save_directory, file_name)
                                case "yes":
                                    save_directory = input('Enter the directory path to save the file: ')
                                    await download_all_file_versions(token, file_id, save_directory)
                                case "cancel":
                                    continue
                                case _:
                                    print('Invalid choice!\n')
                case 4:
                    #Updates on OneDrive
                    await monitor_onedrive_activities(token)
                case 5:
                    #Restore a version
                    file_id = input('Enter the file ID (type cancel for exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        if not await validate_file_id(token, file_id):
                            print("Invalid file ID. Please try again.")
                        else:
                            version_id = input('Enter the version ID of the file you wnat to restore (type cancel for exit): ')
                            if version_id == "cancel":
                                continue
                            else:
                                if not await validate_version_id(token, file_id, version_id):
                                            print("Invalid version ID. Please try again.")
                                else:
                                    await restore_file_version(graph, token, file_id, version_id)
                case _:
                    print('Invalid choice!\n')
        except ODataError as odata_error:
            print('Error:')
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)

# Greet user
async def greet_user(graph: Graph):
    user = await graph.get_user()
    if user:
        print('Hello,', user.display_name)
        print('Email:', user.mail or user.user_principal_name, '\n')

# Helper function to print items in a hierarchical format
async def print_items(token, items, level=0):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    for item in items['value']:
        print(' ' * level * 4 + '- ' + f"[{item['id']}] {item['name']}")
        if 'folder' in item:
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{item["id"]}/children'
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            sub_items = response.json()
            await print_items(token, sub_items, level + 1)

# Display list files
async def display_list_files(graph: Graph, token):
    
    headers = {
        'Authorization': f'Bearer {token}'
    }

    try:
        response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
        response.raise_for_status()  # Verifica se la richiesta ha avuto successo o solleva un'eccezione
        root_items = response.json()

        await print_items(token, root_items)

    except requests.exceptions.HTTPError as err:
        print(f'Errore HTTP: {err}')  # Gestione dell'errore HTTP

    except requests.exceptions.RequestException as err:
        print(f'Errore di richiesta: {err}')  # Gestione di altri tipi di eccezioni di richiesta

async def validate_file_id(token, file_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

async def validate_version_id(token, file_id, version_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

async def display_file_versions(graph: Graph, token, file_id):
    try:
        headers = {
            'Authorization': f'Bearer {token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions'

        response = requests.get(url, headers=headers)
        response.raise_for_status()
        versions = response.json()

        for version in versions['value']:
            print(f"Version ID: {version['id']}")
            print(f"Last Modified: {version['lastModifiedDateTime']}")
            print(f"Size: {version['size']} bytes\n")
    
    except requests.exceptions.HTTPError as e:
        print(f"Make sure you have entered a valid File ID {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

async def download_file_version(token, file_id, version_id, save_directory, file_name):
    try:
        headers = {
            'Authorization': f'Bearer {token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content'

        response = requests.get(url, headers=headers, allow_redirects=True)
        response.raise_for_status()

        # Create directory if it does not exist
        if save_directory and not os.path.exists(save_directory):
            os.makedirs(save_directory)

        save_path = os.path.join(save_directory, file_name)

        with open(save_path, 'wb') as file:
            file.write(response.content)

        print(f"File version downloaded successfully as {save_path}")
    
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error occurred: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Request Exception occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

async def download_all_file_versions(token, file_id, save_directory):
    try:
        headers = {
            'Authorization': f'Bearer {token}'
        }
        url_versions = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions'

        response_versions = requests.get(url_versions, headers=headers)
        response_versions.raise_for_status()
        versions = response_versions.json()

        if not versions['value']:
            print("No versions found for this file.")
            return

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
            file_name = f"{os.path.splitext(original_name)[0]}_{version_id}{original_extension}"
            await download_file_version(token, file_id, version_id, save_directory, file_name)

    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error occurred: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Request Exception occurred: {e}")
    except ValueError as e:
        print(f"Value Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Funzione asincrona per monitorare le attività su OneDrive
async def monitor_onedrive_activities(token):
    url_delta = 'https://graph.microsoft.com/v1.0/me/drive/root/delta'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    try:
        # Monitorare le modifiche su OneDrive
        response_delta = requests.get(url_delta, headers=headers)
        response_delta.raise_for_status()
        changes = response_delta.json()

        print("Changes in OneDrive:")
        print("=" * 140)
        print("| {:<30} | {:<40} | {:<20} | {:<20} | {:<20} |".format("Filename", "File ID", "Action", "Modified By", "Modified Time"))
        print("-" * 140)

        for item in changes.get('value', []):
            filename = item.get('name', 'N/A')
            file_id = item.get('id', 'N/A')
            modified_time = item.get('lastModifiedDateTime', 'N/A')
            modified_by = item.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'N/A')
            action = 'Deleted' if 'deleted' in item else 'Modified'
            print("| {:<30} | {:<40} | {:<20} | {:<20} | {:<20} |".format(filename[:30], file_id[:40], action, modified_by[:20], modified_time[:20]))

        print("=" * 140)

    except requests.exceptions.HTTPError as err:
        print(f"HTTP error occurred: {err}")
    except Exception as err:
        print(f"An error occurred: {err}")

async def restore_file_version(graph: Graph, token, file_id, version_id):
    try:
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/restoreVersion'

        response = requests.post(url, headers=headers)
        response.raise_for_status()

        print(f"Version {version_id} of file {file_id} restored successfully.")
    
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error occurred: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Request Exception occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Run main
asyncio.run(main())