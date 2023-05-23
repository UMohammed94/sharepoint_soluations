import requests, json, pathlib, io
from openpyxl import load_workbook
from sp_config import CREDS, SHAREPOINT_URLS

def make_sharepoint_creds():
    try:
        sharepoint_creds= {
        'client_id':CREDS['client_id'],
        'client_secret':CREDS['client_secret'],
        'tenant':CREDS['tenant'],
        'tenant_id':CREDS['tenant_id']
        }
        sharepoint_creds['client_tenant_id']= sharepoint_creds['client_id'] + '@' + sharepoint_creds['tenant_id']
        return sharepoint_creds
    except Exception as err:
        print("ERROR: " + str(err))

def generate_sp_api_json():
    try:
        sharepoint_creds= make_sharepoint_creds()
        req_info= {
            'grant_type':'client_credentials',
            'resource': "00000003-0000-0ff1-ce00-000000000000/" + sharepoint_creds['tenant'] + ".sharepoint.com@" + sharepoint_creds['tenant_id'], 
            'client_id': sharepoint_creds['client_tenant_id'],
            'client_secret': sharepoint_creds['client_secret']
        }

        auth_headers= {
            'Content-Type':'application/x-www-form-urlencoded'
        }

        url = "https://accounts.accesscontrol.windows.net/{}/tokens/OAuth/2".format(sharepoint_creds['tenant_id'])
        r = requests.post(url, data=req_info, headers=auth_headers)
        json_data= json.loads(r.text)

        return json_data
    except Exception as err:
        print("ERROR: " + str(err))

def generate_headers():
    try:
        json_data=generate_sp_api_json()
        post_req_headers= {
            'Authorization': "Bearer " + json_data['access_token'],
            'Accept':'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }
        return post_req_headers
    except Exception as err:
        print("ERROR: " + str(err))

def generate_delete_headers():
    try:
        json_data=generate_sp_api_json()
        sharepoint_creds= make_sharepoint_creds()

        site_url = "https://{}{}".format(
            SHAREPOINT_URLS["tenant_url"],
            SHAREPOINT_URLS["group_url"])
        
        endpoint_url = site_url + '/_api/contextinfo'

        response = requests.post(endpoint_url, headers=generate_headers())
        response.raise_for_status()

        response_json = response.json()
        form_digest_value = response_json['d']['GetContextWebInformation']['FormDigestValue']

        delete_req_headers={
            'Authorization': "Bearer " + json_data['access_token'], 
            'Accept':'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': form_digest_value,
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*'
            }
        return delete_req_headers
    except Exception as err:
        print("ERROR: " + str(err))

def find_sharepoint_files(sharepoint_directory_path: str) -> list:
    try:
        filenames=[]
        url = generate_sharepoint_url('make_url_all_files', False, sharepoint_directory_path)
        response = requests.get(url, headers=generate_headers())
        response_data = response.text
        json_data = json.loads(response_data)
        for index in range(len(json_data['d']['results'])):
            sp_file_names = str(json_data['d']['results'][index]['Name'])
            filenames.append(sp_file_names)
        return filenames
    except Exception as err:
        print("ERROR: " + str(err))

def find_desired_filename(filenames: list, user_file: str) -> str:
    try:
        desired_filename=''

        for file in filenames:
            if file.lower() == user_file.lower():
                desired_filename=file

        if len(desired_filename) == 0:
            desired_filename = 'file_not_found'

        return desired_filename
    except Exception as err:
        print("ERROR: " + str(err))

def get_file_extensions(desired_file: str) -> str:
    try:
        sp_file_extension=pathlib.Path(desired_file).suffix
        return sp_file_extension
    except Exception as err:
        print("ERROR: " + str(err))

def check_extension_and_download(file_extension: str, response, file_name: str, download_file_path: str) -> None:
    try:
        if file_extension=='.csv':
            data=response.content
            with open(download_file_path + '/' + file_name + ".csv", "wt", encoding="utf-8", newline="") as file_out:
                file_out.writelines(data.decode())

        elif file_extension=='.xlsx':
            data=response.content.strip()
            work_book = load_workbook(filename=(io.BytesIO(data)), data_only=True)
            work_book.save(download_file_path + '/'+ file_name)

        else:
            print("file type not supported: ", ' ' , file_extension)
    except Exception as err:
        print("ERROR: " + str(err))

def generate_sharepoint_url(identifier: str, file_or_folder_name: str, sharepoint_directory_path: str) -> str:
    try:
        if identifier == 'make_url_all_files':
                url="https://{}{}{}('{}'){}".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["api_func"],
                    sharepoint_directory_path,
                    SHAREPOINT_URLS["all_files_suffix"]
                )
        elif identifier == 'make_url_download_file':
                url="https://{}{}{}('{}'){}('{}')/$value".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["api_func"],
                    sharepoint_directory_path,
                    SHAREPOINT_URLS["all_files_suffix"],
                    file_or_folder_name
                )
        elif identifier == 'make_url_upload_files':
                url="https://{}{}{}('{}')/Files/add(url='{}',overwrite=true)".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["api_func"],
                    sharepoint_directory_path,
                    file_or_folder_name
                )
        # elif identifier == 'make_url_list_all_files':
        #         url="https://{}{}{}('{}/{}')".format(
        #             SHAREPOINT_URLS["tenant_url"],
        #             SHAREPOINT_URLS["group_url"],
        #             SHAREPOINT_URLS["api_func"],
        #             SHAREPOINT_URLS["group_url"],
        #             sharepoint_directory_path
        #         )
        elif identifier == 'make_url_delete_file':
                url="https://{}{}{}('{}/{}/{}')".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["api_func"],
                    SHAREPOINT_URLS["group_url"],
                    sharepoint_directory_path,
                    file_or_folder_name
                )
        elif identifier == 'make_url_check_folders':
                url = "https://{}{}{}('{}')".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["api_func"],
                    sharepoint_directory_path+'/'+file_or_folder_name
                )
        elif identifier == 'make_url_desired_folder_name':
                url = "https://{}{}/{}/{}".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    sharepoint_directory_path,
                    file_or_folder_name
                )
        elif identifier == 'make_url_directory_path':
                url = "https://{}{}{}".format(
                    SHAREPOINT_URLS["tenant_url"],
                    SHAREPOINT_URLS["group_url"],
                    SHAREPOINT_URLS["create_directory_url"]
                )
        else:
            return False

        return url
    except Exception as err:
        print("ERROR: " + str(err))