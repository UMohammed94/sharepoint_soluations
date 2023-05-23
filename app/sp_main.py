# APPLICATION NAME: SHAREPOINT HELPER LIB
# VERSION: 5
# FEATURES: UPLOAD FILES, DOWNLOAD FILES, DELETE FILES, MOVE FILES, CREATE FOLDER 

import requests, os, shutil, json
import datetime as dt
from os import listdir
from os.path import isfile, join
from sp_config import SHAREPOINT_URLS
from sp_helpers import generate_headers, generate_sharepoint_url, generate_delete_headers, get_file_extensions, check_extension_and_download, find_sharepoint_files, find_desired_filename

def archive_local_files(local_source_directory_path: str, local_destination_directory_path: str) -> str:
    try:
        all_files=os.listdir(local_source_directory_path)
        for file in all_files:
            src_path=os.path.join(local_source_directory_path, file)
            dst_path=os.path.join(local_destination_directory_path, file)

            if os.path.exists(os.path.join(local_destination_directory_path, file)):

                    get_original_file_name_extension = get_file_extensions(file)

                    original_file_name = file.split(get_original_file_name_extension)

                    # Add the split part back to the list
                    original_file_name[-1] += get_original_file_name_extension

                    new_file_name = original_file_name[0]+"_"+str(dt.datetime.now().strftime("%Y%m%d%H%M%S"))+original_file_name[1]

                    # rename the file
                    os.rename(local_source_directory_path+'/'+file, local_source_directory_path+'/'+new_file_name)

                    print(f"File '{file}' already exists. Renaming the file to '{new_file_name}' and moving into archive directory on local machine '{local_source_directory_path}'")
                    shutil.move(local_source_directory_path+'/'+new_file_name, local_destination_directory_path)  
            else:
                shutil.move(src_path,dst_path)
                print(f"File '{file}' successfully moved to '{dst_path}'")

        return True
    except Exception as err:
        print("ERROR: " + str(err))

def delete_all_local_files(local_directory_path: str) -> None:
    try:
        for filename in os.listdir(local_directory_path):
            file_path = os.path.join(local_directory_path, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
                print('Deleted the following file from local machine: %s' % file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
                print('Deleted the following directory from local machine: %s' % file_path)
    except Exception as e:
        print('Failed to delete %s. Reason: %s' % (file_path, e))

def delete_specific_local_file(local_file_path: str) -> None:
    try:
        if os.path.isfile(local_file_path) or os.path.islink(local_file_path):
            os.unlink(local_file_path)
            print('Deleted the following file from local machine: %s' % local_file_path)
        elif os.path.isdir(local_file_path):
            shutil.rmtree(local_file_path)
            print('Deleted the following directory from local machine: %s' % local_file_path)
    except Exception as e:
        print('Failed to delete %s. Reason: %s' % (local_file_path, e))

def list_all_files(sharepoint_directory_path: str) -> list:
    try:
        list_of_files = []
        url = generate_sharepoint_url('make_url_all_files', False, sharepoint_directory_path)
        response = requests.get(url, headers=generate_headers())
        files = json.loads(response.content)['d']['results']
        if len(files) > 0:
            for file in files:
                file_url = file['__metadata']['uri']
                file_name = file['Name']
                if response.status_code == 200:
                    print(f"File '{file_name}' Found")
                else:
                    print(f"Error moving file '{file_name}': " + response.content.decode("utf-8"))
                list_of_files.append(file_name)
        else:
            print(f'No files found in sharepoint directory: "{sharepoint_directory_path}"')

        return list_of_files   
    except Exception as err:
            print("ERROR: " + str(err)) 

def rename_and_move_sharepoint_file(sharepoint_source_directory_path: str, sharepoint_destination_directory_path: str, file_name: str) -> None:
    try:
        old_file_relative_url = SHAREPOINT_URLS["group_url"] +'/'+ sharepoint_source_directory_path +'/'+ file_name

        get_old_file_name_extension = get_file_extensions(file_name)

        old_file_name = file_name.split(get_old_file_name_extension)

        # Add the split part back to the list
        old_file_name[-1] += get_old_file_name_extension

        new_file_name = old_file_name[0]+"_"+str(dt.datetime.now().strftime("%Y%m%d%H%M%S"))+old_file_name[1]

        destination_folder_url = SHAREPOINT_URLS["group_url"] +'/'+ sharepoint_destination_directory_path +'/'+ new_file_name

        move_url = 'https://' + SHAREPOINT_URLS['tenant_url'] + SHAREPOINT_URLS['group_url'] + SHAREPOINT_URLS['api_file_func'] + "('" + old_file_relative_url + "')/MoveTo"
        move_request_body = {'newUrl': destination_folder_url}

        response = requests.post(move_url, headers=generate_headers(), json=move_request_body)

        if response.status_code == 200:
            print(f"File '{file_name}' renamed successfully to '{new_file_name}'")
        else:
            print(f"Failed to rename the '{file_name}' file. Error: '{response.content}'")
    except Exception as err:
            print("ERROR: " + str(err))

def move_all_files(sharepoint_source_directory_path: str, sharepoint_destination_directory_path: str) -> None:
    try:
        all_file_names = list_all_files(sharepoint_source_directory_path)

        if len(all_file_names) > 0:
            source_folder_url = SHAREPOINT_URLS['group_url'] + '/' + sharepoint_source_directory_path
            destination_folder_url = SHAREPOINT_URLS['group_url'] + '/' + sharepoint_destination_directory_path

            for file_name in all_file_names:
                file_url = source_folder_url + '/' + file_name
                move_url = 'https://' + SHAREPOINT_URLS['tenant_url'] + SHAREPOINT_URLS['group_url'] + SHAREPOINT_URLS['api_file_func'] + "('" + file_url + "')/MoveTo"
                move_request_body = {'newUrl': destination_folder_url + "/" + file_name}

                response = requests.post(move_url, headers=generate_headers(), json=move_request_body)
                
                if response.status_code == 200:
                    print(f"File '{file_name}' moved successfully to {destination_folder_url}")
                else:
                    print("Error moving file: " + response.content.decode("utf-8"))

                    error_response_content = response.content.decode("utf-8")
                    error_response_json = json.loads(error_response_content)
                    split_error_code = str(error_response_json['error']['code']).split(',')
                    get_error_code = str(split_error_code[0])
                    # checking if the same file name exist in the destination directory. If it's true then we will rename and move the file to destination directory.
                    # error code in sharepoint api if file name already exists in destination directory: "-2130575257, Microsoft.SharePoint.SPException"
                    if str(-2130575257) == get_error_code:
                        rename_and_move_sharepoint_file(sharepoint_source_directory_path, sharepoint_destination_directory_path, file_name)
    except Exception as err:
        print("ERROR: " + str(err))

def delete_all_files(sharepoint_directory_path: str) -> None:
    try:
        all_file_names = list_all_files(sharepoint_directory_path)

        if len(all_file_names) > 0:
            for file_name in all_file_names:
                url = generate_sharepoint_url('make_url_delete_file', file_name, sharepoint_directory_path)
                print(f"Deleting file '{file_name}'")
                response = requests.post(url, headers=generate_delete_headers())
                if response.status_code == 200:
                    print(f"Deleted file '{file_name}' successfully")
                else:
                    print(f"Error deleting file: '{response.text}'") 
    except Exception as err:
        print("ERROR: " + str(err))
    
def delete_one_file(sharepoint_directory_path: str, file_name: str) -> None:
    try:
        url = generate_sharepoint_url('make_url_delete_file', file_name, sharepoint_directory_path)
        print(f"Deleting file '{file_name}'")
        response = requests.post(url, headers=generate_delete_headers())
        if response.status_code == 200:
            print(f"Deleted file '{file_name}' successfully")
        else:
            print(f"Error deleting file: '{response.text}'") 
    except Exception as err:
        print("ERROR: " + str(err))

def make_upload_request(sharepoint_directory_path: str, file_name: str, file_path: str) -> None:     
    with open(file_path, "rb") as file_input:
        try:
            url = generate_sharepoint_url('make_url_upload_files', file_name, sharepoint_directory_path)
            print(f"Uploading '{file_name}' from '{file_path}'")
            response = requests.post(url, headers=generate_headers(), data=file_input)

            if response.status_code == 200:
                print(f"Uploaded {file_name} file successfully")
        except Exception as err:
            print("ERROR: "+ str(err))

def upload_to_sharepoint(sharepoint_directory_path: str, local_directory_path: str) -> str:
    try:
        all_files = listdir(local_directory_path)
        for file_name in all_files:
            if isfile(join(local_directory_path, file_name)):
                file_path=os.path.join(local_directory_path, ''.join((file_name)))
                make_upload_request(sharepoint_directory_path, file_name, file_path)
        return True
    except Exception as err:
            print("ERROR: " + str(err))

def make_download_request(sharepoint_directory_path: str, file_name: str) -> None:    
    try:
        print(f"Downloading '{file_name}' from sharepoint")
        url = generate_sharepoint_url('make_url_download_file', file_name, sharepoint_directory_path)
        response = requests.get(url, headers = generate_headers())
        return response
    except Exception as err:
        print("ERROR: "+ str(err))

def make_download_all_files_request(sharepoint_directory_path: str, file_names: list, download_local_file_path: str) -> None:
    try:
        for file in file_names:
            print(f"Downloading '{file}' from sharepoint")
            url = generate_sharepoint_url('make_url_download_file', file, sharepoint_directory_path)
            response = requests.get(url, headers = generate_headers())
            file_extension = get_file_extensions(file)
            check_extension_and_download(file_extension, response, file, download_local_file_path)
    except Exception as err:
        print("ERROR: "+ str(err))

def download_sharepoint_file(sharepoint_directory_path: str, file_name: str, download_local_file_path: str) -> None:
    try:
        file_names_list = find_sharepoint_files(sharepoint_directory_path)
        sp_desired_file = find_desired_filename(file_names_list, file_name)
        response = make_download_request(sharepoint_directory_path, sp_desired_file)
        file_extension = get_file_extensions(sp_desired_file)
        check_extension_and_download(file_extension, response, sp_desired_file, download_local_file_path)
    except Exception as err:
        print("ERROR: " + str(err))

def download_all_sharepoint_files(sharepoint_directory_path: str, download_local_file_path: str) -> None:
    try:
        file_names = find_sharepoint_files(sharepoint_directory_path)
        make_download_all_files_request(sharepoint_directory_path, file_names, download_local_file_path)
    except Exception as err:
        print("ERROR: " + str(err))

def create_folder(existing_parent_directory_name: str, new_folder_name: str) -> str:
    try:
        print(f"Checking if folder '{new_folder_name}' exists in directory: '{existing_parent_directory_name}' on sharepoint")

        # Send GET request to SharePoint to check if folder exists
        check_folder_response = requests.get(url=generate_sharepoint_url('make_url_check_folders', new_folder_name, existing_parent_directory_name), headers=generate_headers())

        # Check if folder already exists
        if check_folder_response.status_code == 200:
            print(f"Folder '{new_folder_name}' already exists in directory: '{existing_parent_directory_name}' on sharepoint")
            return True
        else:
            print(f"Creating a folder: '{new_folder_name}' in directory: '{existing_parent_directory_name}' on sharepoint")
            # SharePoint request body
            body = {
                "__metadata": {
                    "type": "SP.Folder"
                },
                "ServerRelativeUrl": generate_sharepoint_url('make_url_desired_folder_name', new_folder_name, existing_parent_directory_name)
            }

            # Create folder in SharePoint
            response = requests.post(url=generate_sharepoint_url('make_url_directory_path', False, False), headers=generate_headers(), data=json.dumps(body), verify=False)

            # Check if folder was created successfully
            if response.status_code == 201:
                print(f"Created a folder '{new_folder_name}' in directory: '{existing_parent_directory_name}' on sharepoint")
                return True
            else:
                print(f"Error creating folder: {response.text}")
                return f"Error creating folder: {response.text}"           
    except Exception as err:
        print("ERROR: " + str(err))