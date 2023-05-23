Hello Developer! 

This Sharepoint Helper LIB is a library that should sit in your project folder to assist with uploading and downloading files with sharepoint. 

Currently it only uploads and downloads however, in future iterations we will add other CRUD functionalities. 

What you need to use this:

1) Create an ENV file in the app folder with the following information, which you can retrieve by creating an app on microsoft Sharepoint using the link bellow and following the steps

    Link to site for creating an app: 
    https://[your_tenant].sharepoint.com/sites/TestCommunication/_layouts/15/appregnew.aspx
    
    ENV REQUIREMENTS:
    client_id=***
    client_secret=***
    tenant=***
    tenant_id=***

    !REMEMBER! Always print your creds to confirm they are being imported in properly

2) Go into sp_config.py and go to the dictionary SHAREPOINT_URLS and change the upload and download path to desired location following the same convention

3) For downloads, in your project main.py, call download_sharepoint_file(file_name) and the parameter should be a string that contains characters in your desired download file, or create a variable file_name with characters in your desired download file

4) For uploads, in your project main.py, call upload_to_sharepoint() and place your folders in the upload_files folder or point to a folder of your choice using the UPLOAD_FILE_PATH variable. 