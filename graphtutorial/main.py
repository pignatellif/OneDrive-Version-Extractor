import asyncio
import requests
import configparser
import os
import pytz
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph
from datetime import datetime
from dateutil.parser import parse
from tzlocal import get_localzone
import csv

async def main():
    print('OneDrive-Version-Extractor\n')

    # Load settings from configuration files
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    # Initialize Graph object with Azure settings
    graph = Graph(azure_settings)

    # Greet the user and get the token for API requests
    await greet_user(graph)
    token = await graph.get_user_token()

    choice = -1

    # Main menu loop
    while choice != 0:
        print('Please choose one of the following options:')
        print('0. Exit')
        print('1. Display list of files')
        print('2. Display file versions')
        print('3. Download a file version')
        print('4. Monitor updates on OneDrive')
        print('5. Restore a file version')

        try:
            choice = int(input())  # Take user input for menu choice
        except ValueError:
            choice = -1

        try:
            # Handle user's choice using match-case (Python 3.10+ feature)
            match choice:
                case 0:
                    print('Goodbye...')
                case 1:
                    await display_list_files(graph, token)
                case 2:
                    file_id = input('Enter the file ID to view its versions (type cancel to exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        await display_file_versions(graph, token, file_id)
                case 3:
                    file_id = input('Enter the file ID to download (type cancel to exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        if not await validate_file_id(token, file_id):
                            print("Invalid file ID. Please try again.")
                        else:
                            choice_2 = input('Do you want to download all versions of the file? (yes/no/cancel to exit): ')
                            match choice_2:
                                case "no":
                                    version_id = input('Enter the version ID to download (type cancel to exit): ')
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
                                    save_directory = input('Enter the directory path to save the files: ')
                                    await download_all_file_versions(token, file_id, save_directory)
                                case "cancel":
                                    continue
                                case _:
                                    print('Invalid choice!\n')
                case 4:
                    await monitor_onedrive_activities(token)
                case 5:
                    file_id = input('Enter the file ID to restore (type cancel to exit): ')
                    if file_id == "cancel":
                        continue
                    else:
                        if not await validate_file_id(token, file_id):
                            print("Invalid file ID. Please try again.")
                        else:
                            version_id = input('Enter the version ID to restore (type cancel to exit): ')
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

# Greet user with their name and email
async def greet_user(graph: Graph):
    user = await graph.get_user()
    if user:
        print('Hello,', user.display_name)
        print('Email:', user.mail or user.user_principal_name, '\n')

# Recursive function to print items in a hierarchical structure
async def print_items(token, items, level=0):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    for item in items['value']:
        print(' ' * level * 4 + '- ' + f"[{item['id']}] {item['name']}")
        if 'folder' in item:
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item['id']}/children"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            sub_items = response.json()
            await print_items(token, sub_items, level + 1)

# Display a list of files and folders in OneDrive
async def display_list_files(graph: Graph, token):
    headers = {
        'Authorization': f'Bearer {token}'
    }

    try:
        response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
        response.raise_for_status()
        root_items = response.json()

        await print_items(token, root_items)

    except requests.exceptions.HTTPError as err:
        print(f'HTTP Error: {err}')

    except requests.exceptions.RequestException as err:
        print(f'Request Error: {err}')

# Validate if a file ID exists in OneDrive
async def validate_file_id(token, file_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

# Validate if a version ID exists for a given file in OneDrive
async def validate_version_id(token, file_id, version_id):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/versions/{version_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200

# Display versions of a file in OneDrive with option to export to CSV
async def display_file_versions(graph: Graph, token, file_id):
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
                print(f"Errore nel parsing della data: {e}")
                continue

            # Directly convert the datetime object to the local timezone
            modified_time_localized = modified_time.astimezone(local_tz)
            modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

            print(f"Version ID: {version['id']}")
            print(f"Last Modified: {modified_time_str_with_tz}")
            print(f"Size: {version['size']} bytes\n")

        # Ask user if they want to export to CSV
        export_csv = input('Do you want to export the versions to a CSV file? (yes/no): ')
        if export_csv.lower() == 'yes':
            csv_filename = f"{file_id}_versions.csv"
            directory = input('Enter the directory path to save the CSV file: ')

            if not os.path.exists(directory):
                os.makedirs(directory)

            csv_path = os.path.join(directory, csv_filename)

            with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(['Version ID', 'Last Modified', 'Size'])

                for version in versions['value']:
                    modified_time_str = version.get('lastModifiedDateTime', 'N/A')
                    modified_time = parse(modified_time_str)

                    modified_time_localized = modified_time.astimezone(local_tz)
                    modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

                    writer.writerow([version['id'], modified_time_str_with_tz, version['size']])

            print(f"File {csv_path} containing versions metadata exported successfully.")
        elif export_csv.lower() == 'no':
            print("Not exporting to CSV.")
        else:
            print("Invalid choice. Not exporting to CSV.")

    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Download a specific version of a file from OneDrive
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

# Download all versions of a file from OneDrive
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

# Monitor activities and changes in OneDrive
async def monitor_onedrive_activities(token):
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

        print("Changes in OneDrive:")
        print("=" * 171)
        print("| {:<40} | {:<40} | {:<20} | {:<25} | {:<30} |".format("Filename", "File ID", "Action", "Modified By", "Modified Time"))
        print("-" * 171)

        for item in changes.get('value', []):
            filename = item.get('name', 'N/A')
            file_id = item.get('id', 'N/A')
            modified_time_str = item.get('lastModifiedDateTime', 'N/A')
            
            try:
                modified_time = parse(modified_time_str)
            except ValueError as e:
                print(f"Errore nel parsing della data: {e}")
                continue

            # Directly convert the datetime object to the local timezone
            modified_time_localized = modified_time.astimezone(local_tz)
            modified_time_str_with_tz = modified_time_localized.strftime('%Y-%m-%d %H:%M:%S %Z')

            modified_by = item.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'N/A')
            action = 'Deleted' if 'deleted' in item else 'Modified'
            print("| {:<40} | {:<40} | {:<20} | {:<25} | {:<30} |".format(filename[:40], file_id[:40], action, modified_by[:25], modified_time_str_with_tz))

        print("=" * 171)

    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
    except Exception as err:
        print(f"An error occurred: {err}")


# Restore a specific version of a file in OneDrive
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
        print(f"HTTP Error: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Request Exception: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Entry point to run the main function
asyncio.run(main())

