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
                file_id = input('Enter the file ID: ')
                await display_file_versions(graph, token, file_id)
            elif choice == 3:
                file_id = input('Enter the file ID: ')
                version_id = input('Enter the version ID: ')
                save_path = input('Enter the full file path to save (including filename): ')
                await download_file_version(token, file_id, version_id, save_path)
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

# Helper function to get drive items
async def get_drive_items(token, folder_id=None):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    if folder_id:
        url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children'
    else:
        url = 'https://graph.microsoft.com/v1.0/me/drive/root/children'

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Helper function to print items in a hierarchical format
async def print_items(token, items, level=0):
    for item in items['value']:
        print(' ' * level * 4 + '- ' + f"[{item['id']}] {item['name']}")
        if 'folder' in item:
            sub_items = await get_drive_items(token, item['id'])
            await print_items(token, sub_items, level + 1)

# Display list files
async def display_list_files(graph: Graph, token):
    root_items = await get_drive_items(token)
    await print_items(token, root_items)

# Helper function to get file versions
async def get_file_versions(token, file_id):
    headers = {
        'Authorization': f'Bearer {token}'
    }
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions'

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Helper function to print file versions
async def print_file_versions(versions):
    for version in versions['value']:
        print(f"Version ID: {version['id']}")
        print(f"Last Modified: {version['lastModifiedDateTime']}")
        print(f"Size: {version['size']} bytes\n")

# Display file versions
async def display_file_versions(graph: Graph, token, file_id):
    versions = await get_file_versions(token, file_id)
    await print_file_versions(versions)

# Helper function to download the file version
async def download_file_version(token, file_id, version_id, save_path):
    headers = {
        'Authorization': f'Bearer {token}'
    }
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}/content'

    response = requests.get(url, headers=headers, allow_redirects=False)
    response.raise_for_status()

    # The actual download URL is in the 'Location' header
    download_url = response.headers['Location']
    
    response = requests.get(download_url)
    response.raise_for_status()
    
    # Create directory if it does not exist
    if not os.path.exists(os.path.dirname(save_path)):
        os.makedirs(os.path.dirname(save_path))
    
    with open(save_path, 'wb') as file:
        file.write(response.content)
    
    print(f"File version downloaded successfully as {save_path}")

# Run main
asyncio.run(main())
