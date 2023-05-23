import os
from dotenv import load_dotenv

load_dotenv()

SHAREPOINT_URLS={
    'tenant_url': "interpublic.sharepoint.com",
    'group_url': "/sites/T-MobileINITeam",
    'api_func': "/_api/web/GetFolderByServerRelativeUrl",
    'api_file_func' : "/_api/web/GetFileByServerRelativeUrl",
    'create_directory_url': "/_api/web/folders",
    'all_files_suffix':"/Files",
}

CREDS={
    'user_name':os.getenv('user_name'),
    'password':os.getenv('password'),
    'db_name':os.getenv('db_name'),
    'server':os.getenv('server'),
    'client_id':os.getenv('client_id'),
    'client_secret':os.getenv('client_secret'),
    'tenant':os.getenv('tenant'),
    'tenant_id':os.getenv('tenant_id')
}