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

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print('Goodbye...')
            elif choice == 1:
                await display_list_files(graph, token)
            elif choice == 2:
                file_id = input('Enter the file ID (cancel for exit): ')
                if file_id == "cancel":
                    continue
                else:
                    await display_file_versions(graph, token, file_id)
            elif choice == 3:
                file_id = input('Enter the file ID (cancel for exit): ')
                if file_id == "cancel":
                    continue
                else:
                    choice_2 = "temp"
                    while choice_2 != "cancel":
                        choice_2 = input('Do you want to download all versions? (yes/no/cancel for exit): ')
                        if choice_2 == "no":
                            version_id = input('Enter the version ID (cancel for exit): ')
                            if version_id == "cancel":
                                continue
                            else:
                                save_path = input('Enter the full file path to save (including filename and extension): ')
                                await download_file_version(token, file_id, version_id, save_path)
                        elif choice_2 == "yes":
                            #Codice per scaricare tutte le versioni.
                            continue
                        else:
                            continue
            else:
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

# Helper function to download the file version
async def download_file_version(token, file_id, version_id, save_path):
    try:
        headers = {
            'Authorization': f'Bearer {token}'
        }
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content'

        response = requests.get(url, headers=headers, allow_redirects=False)
        response.raise_for_status()

        # The actual download URL is in the 'Location' header
        download_url = response.headers.get('Location')
        if not download_url:
            raise ValueError("Download URL not found")

        response = requests.get(download_url)
        response.raise_for_status()

        # Create directory if it does not exist
        if not os.path.exists(os.path.dirname(save_path)):
            os.makedirs(os.path.dirname(save_path))

        with open(save_path, 'wb') as file:
            file.write(response.content)

        print(f"File version downloaded successfully as {save_path}")
    
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error occurred: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Request Exception occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Run main
asyncio.run(main())
