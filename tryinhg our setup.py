import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from httplib2.error import ServerNotFoundError
from apiclient.http import MediaFileUpload
import os
import win32com.client
from datetime import datetime



def get_file_metadata(path, filename, metadata):
    sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
    ns = sh.NameSpace(path)

    file_metadata = dict()
    item = ns.ParseName(str(filename))
    for ind, attribute in enumerate(metadata):
        attr_value = ns.GetDetailsOf(item, ind)
        if attr_value:
            file_metadata[attribute] = attr_value

    return file_metadata

def time_extract():
    folder = "C:\\Users\\hp\\Desktop\\backup"
    dir_list = os.listdir(folder)
    lst = []
    metadata = ['Name', 'Size', 'Item type', 'Date modified', 'Date created']
    for filename in dir_list:
        fileinfo = get_file_metadata(folder, filename, metadata)
        all_file_metadata.append(fileinfo)
        second_time=fileinfo['Date modified'].strip()
        if fileinfo['Date modified'].strip()[-1].upper()=='M':
            second_time=fileinfo['Date modified'][:-2].strip()
        if (datetime.strptime(dt_string,"%d-%m-%Y %H:%M")-datetime.strptime(second_time,"%d-%m-%Y %H:%M")).total_seconds()/3600 <24:
            lst.append(fileinfo)
    return lst


class MyDrive():
    service = None
    def __init__(self):
        # If modifying these scopes, delete the file token.pickle.
        SCOPES = ['https://www.googleapis.com/auth/drive']
        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
            # The file token.pickle stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('drive', 'v3', credentials=creds)

        
    def list_files(self, page_size=10):
    # Call the Drive v3 API
        results = self.service.files().list(
        pageSize=page_size, fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])

    def upload_file(self, filename, path,cmplist,modified_time_list):
        folder_id = "1dmGQwAGGjRb4V47rrKmCY_SD8un-nJfl"
        media = MediaFileUpload(f"{path}{filename}")

        response = self.service.files().list(
                                        q=f"name='{filename}' and parents='{folder_id}'",
                                        spaces='drive',
                                        fields='nextPageToken, files(name,id)',
                                        pageToken=None).execute()
        
        x=response['files']

        if len(response['files']) == 0:
            file_metadata = {
                'name': filename,
                'parents': [folder_id]
            }
            file = self.service.files().create(body=file_metadata, media_body=media, fields='name').execute()
            print(file)
            with open('log.txt','a') as log_file:
                print(f"{dt_string} backup_successful A new file was created {path}{file.get('name')}",file=log_file)
            print("new file was created")
                
        elif len(response['files']) != 0:
            x=x[0]
            if x['name'] not in cmplist:
                file_metadata = {
                    'name': filename,
                    'parents': [folder_id]
                }
                file = self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                with open('log.txt','a') as log_file:
                    print(f"{dt_string} backup_successful A new file was created {path}{x['name']}",file=log_file)
                print("new file was created")

            else:
                filenames_list = []
                for d in modified_time_list:
                    filenames_list.append(d['Name'])
                i=0
                for name_check in modified_time_list:
                    for file in response.get('files', []):
                        # Process change
                        t=file['name']
                        ind=t.index('.')
                        t=t[:ind]
                        if(name_check["Name"]==t):
                            update_file = self.service.files().update(
                                fileId=file.get('id'),
                                media_body=media,
                            ).execute()
                            with open('log.txt','a') as log_file:
                                print(f"{dt_string} backup_successful Updated file {path}{file.get('name')}",file=log_file)
                            print("updated file")
                            break

    def list_files_drfls(self, q , nextPageToken=None ,page_size=10):
        # Call the Drive v3 API
        drive_results = self.service.files().list(
                q=q,
                pageToken=nextPageToken,
            pageSize=page_size, fields="nextPageToken, files(*)").execute()
        drive_files = drive_results.get('files', [])
        nextPageToken = drive_results.get('nextPageToken')
        if nextPageToken is not None:
            drive_files += self.list_files_drfls(q=q, nextPageToken=nextPageToken)
        return drive_files
            

    def get_files_from_folder(self ,drive_updated,cmplist, folder_id=None):
        if folder_id is not None:
            drive_files = self.list_files_drfls(q=f"'{folder_id}' in parents and trashed=false")
            if drive_files is not None:
                for drfls in drive_files:
                    cmplist.append(drfls['name'])
                    drive_updated.append(drfls['modifiedTime'])
            else:
                print(f'no items found in {folder_id}')

        else:
            print('folder id is not specified')


def main():
    modified_time_list = []
    modified_time_list = time_extract()
    cmplist=[]
    drive_updated=[]
    path = "C:/Users/hp/Desktop/backup/"
    files = os.listdir(path)
    my_drive = MyDrive()
    drive_files = my_drive.list_files_drfls(q="'root' in parents and trashed=false")
    #from drive
    for drfls in drive_files:
        if(drfls.get('mimeType')=='application/vnd.google-apps.folder'):
            my_drive.get_files_from_folder(drive_updated,cmplist,folder_id=drfls.get('id'))
            continue

    #from local folder
    my_drive.list_files()
    time_list = []
    for item in files:
        my_drive.upload_file(item, path,cmplist,modified_time_list)
    
    drive_updated=[]
    for drfls in drive_files:
        if(drfls.get('mimeType')=='application/vnd.google-apps.folder'):
            my_drive.get_files_from_folder(drive_updated,cmplist,folder_id=drfls.get('id'))
            continue

if __name__ == '__main__':
    now = datetime.now()
    all_file_metadata = []
    dt_string = now.strftime("%d-%m-%Y %H:%M")
    try:
        main()
    except Exception as e:
        with open('log.txt','a') as log_file:
            print(dt_string, "backup_failed", e, file=log_file)
