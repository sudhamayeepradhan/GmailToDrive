from datetime import datetime
from Google import Create_Service
import time
import io
import base64 #download attachments
from googleapiclient.http import MediaIoBaseUpload
import DatePicker
files_count = 0
def ExecuteMain():
    def construct_service(api_service):
        CLIENT_SERVICE_FILE = 'C:/Users/dv914285/Desktop/work/GmailToDrive_2/client_secret_974510245116.json'
        try:
            if api_service =='drive':
                API_NAME ='drive'
                API_VERSION = 'v3'
                SCOPES = ['https://www.googleapis.com/auth/drive']
                return Create_Service(CLIENT_SERVICE_FILE,API_NAME,API_VERSION,SCOPES)
            elif api_service=='gmail':
                API_NAME ='gmail'
                API_VERSION = 'v1'
                SCOPES = ['https://mail.google.com/',
                        'https://www.googleapis.com/auth/gmail.settings.basic',
                        'https://www.googleapis.com/auth/gmail.readonly']
                return Create_Service(CLIENT_SERVICE_FILE,API_NAME,API_VERSION,SCOPES)
        except Exception as e:
            print(e)
            return None

    def search_email(service,query_string,label_ids=[]):
        try:
            message_list_response = service.users().messages().list(
                userId = 'me',
                labelIds=label_ids,
                q=query_string
            ).execute()

            message_items=message_list_response.get('messages')
            nextPageToken=message_list_response.get('nextPageToken')

            while nextPageToken:
                message_list_response=service.users().messages().list(
                    userId='me',
                    labelIds=label_ids,
                    q=query_string,
                    pageToken=nextPageToken
                ).execute()

                message_items.extend(message_list_response.get('messages'))
                nextPageToken= message_list_response.get('nextPageToken')
            return message_items

        except Exception as e:
            return None

    def init_service():
        gmail_service = construct_service('gmail')
        time.sleep(5)
        drive_service = construct_service('drive')
        time.sleep(5)
        return gmail_service, drive_service

    def get_message_detail(service,message_id,format='metadata',metadata_headers=[]):
        try:
            message_detail=service.users().messages().get(
                userId='me',
                id=message_id,
                format=format,
                metadataHeaders=metadata_headers
            ).execute()
            return message_detail
        except Exception as e:
            print(e)
            return None

    def create_folder_in_drive(service,sub_folder,folder_name,parent_folder=[]):
        folder_exist = False
        sub_folder_exist = False
        response = service.files().list(q="'root' in parents and trashed = false and mimeType='application/vnd.google-apps.folder'", fields='files(id,name)', spaces='drive').execute()
        parent_folder = [file.get('id') for file in response.get('files',[]) if file.get('name') == 'GmailToDrive' ]
        # set_trace()
        response = service.files().list(q="'"+parent_folder[0]+"' in parents and trashed = false and mimeType='application/vnd.google-apps.folder'", fields='files(id,name)').execute()
        for file in response.get('files', []):
            if file.get('name') == folder_name:
                folder_exist = True
                file_id = {'id':file.get('id')}
                break
        # print ("Folder_exist: "+ str(folder_exist))
        # print ("Folder name: ", folder_name)
        if not folder_exist:
            file_metadata={
                'name':folder_name,
                'parents': parent_folder,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            file_id=service.files().create(body=file_metadata,fields='id').execute()
        
        sub_response = service.files().list(q="'"+file_id['id']+"' in parents and trashed = false and mimeType='application/vnd.google-apps.folder'", fields='files(id,name)', spaces='drive').execute()
        for file in sub_response.get('files', []):
            if file.get('name') == sub_folder:
                sub_folder_exist = True
                file_id = {'id':file.get('id')}
                break
        if not sub_folder_exist:
            sub_folder_metadata={
                'name':sub_folder,
                'parents': [file_id['id']],
                'mimeType': 'application/vnd.google-apps.folder'
            }
            file_id=service.files().create(body=sub_folder_metadata,fields='id').execute()
        
        return sub_folder_exist, file_id

    gmail_service, drive_service = init_service()
    if not gmail_service:
        raise Exception("Gmail Service could not be created")
    if not drive_service:
        raise Exception("Drive Service could not be created")
    try:
        frmDate, toDate = DatePicker.get_date()
    except Exception as e:
        print(e)
        raise Exception(e)
    # import pdb;pdb.set_trace()
    print("Saving attachments from ", frmDate, "till ", toDate)
    newToDate = int(datetime.combine(toDate, datetime.max.time()).timestamp())
    query_string ='(before:'+str(newToDate)+' after:'+str(frmDate)+' (has:attachment OR has:drive))'
    email_messages = search_email(gmail_service,query_string,['INBOX'])
    if not email_messages:
        raise Exception("No message could be retrieved")
    folder_to_be_updated = []
    for email_message in email_messages:
        messageId = email_message['threadId']
    
        messageDetail = get_message_detail(
        gmail_service,email_message['id'],format='full',
        metadata_headers=['parts'])
        messageDetailPayload = messageDetail.get('payload')
    
        for item in messageDetailPayload['headers']:
            if item['name']=='Subject':
                if item['value']:
                    messageSubject= '{0}'.format(item['value'])
                else:
                    messageSubject = '(No Subject) ({0})'.format(messageId)
            elif item['name'] == 'From':
                fromValue = item['value']
    
        # import pdb;pdb.set_trace()
        folderTobeCreated = fromValue.split('@',1)[1].split('.')[0].title()
        folder_to_be_updated.append(folderTobeCreated)
        folder_to_be_updated = list(set(folder_to_be_updated))
        global files_count
        folder_exists, folder=create_folder_in_drive(drive_service, messageSubject, folderTobeCreated)
        folder_id = folder['id']

        if 'parts' in messageDetailPayload:
            for msgPayload in messageDetailPayload['parts']:
                mime_type = msgPayload['mimeType']
                file_name=msgPayload['filename']
                body=msgPayload['body']

                if 'attachmentId' in body:
                    attachment_id=body['attachmentId']
                    response=gmail_service.users().messages().attachments().get(
                    userId='me',
                    messageId=email_message['id'],
                    id=attachment_id
                ).execute()
            
                    file_data=base64.urlsafe_b64decode(
                    response.get('data').encode('UTF-8'))
                    fh=io.BytesIO(file_data)
                    file_metadata={
                    'name':file_name,
                    'parents': [folder_id]
                }
                
                    if not folder_exists:
                        files_count+=1
                        media_body=MediaIoBaseUpload(fh,mimetype=mime_type,chunksize=1024*1024,resumable=True)
                        file=drive_service.files().create(
                        body=file_metadata,
                        media_body=media_body,
                        fields='id'
                    ).execute()
    print("Folders Updated are: ", folder_to_be_updated)
    print(files_count, " number(s) of files created\n")
if __name__ == "__main__":
    import time
    # get the start time
    st = time.time()
    ExecuteMain()
    # get the end time
    et = time.time()
    # get the execution time
    elapsed_time = et - st
    final_res = elapsed_time / 60
    if final_res > 1:
        print('Execution time:', final_res, 'minutes')
    else:
        print('Execution time:', elapsed_time, 'seconds')
    k=input("press Enter to exit")
