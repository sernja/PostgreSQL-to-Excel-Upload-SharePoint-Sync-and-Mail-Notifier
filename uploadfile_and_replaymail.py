import requests

async def upload_large_file(token, file_name, site_id, item_path, file_path, file_size):
    # Step 1: Create the upload session request body
    upload_session_request = {
        "item": {
            "@microsoft.graph.conflictBehavior": "replace",
            "name": file_name
        }
    }

    # Step 2: Create the upload session
    upload_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{item_path}:/createUploadSession"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.post(upload_session_url, json=upload_session_request, headers=headers)
    
    if response.status_code != 200:
        raise Exception(f"Failed to create upload session: {response.status_code} - {response.text}")
    
    upload_session = response.json()

    # Step 3: Upload the file in chunks
    max_slice_size = 320 * 1024  # 320KB
    max_attempts = 5

    async def upload_chunks(upload_url, file_stream, stream_size, chunk_size, max_attempts):
        start = 0
        while start < stream_size:
            end = min(start + chunk_size, stream_size) - 1
            file_stream.seek(start)
            chunk_data = file_stream.read(chunk_size)

            headers = {
                "Content-Range": f"bytes {start}-{end}/{stream_size}"
            }

            for attempt in range(max_attempts):
                response = requests.put(upload_url, headers=headers, data=chunk_data)
                
                if response.status_code in [200, 201, 202]:  # Success status codes
                    print(f"Uploaded {end + 1} bytes out of {stream_size}")
                    break
                else:
                    print(f"Error during upload attempt {attempt + 1}: {response.status_code} - {response.text}")
                    if attempt == max_attempts - 1:
                        raise Exception(f"Failed to upload chunk: {response.status_code} - {response.text}")

            start = end + 1

    with open(file_path, 'rb') as file_stream:
        await upload_chunks(upload_session['uploadUrl'], file_stream, file_size, max_slice_size, max_attempts)

    print("Upload complete")

def sendmail(token, message_mail, link_share, year_month, message_id):
    # Define the request body
    request_body = {
    "message": {
        "body": {
            "contentType": "HTML",
            "content": f"""{message_mail}<br><br>
            <span style="background-color:rgb(244,244,244)">
            <a href="{link_share}" id="OLK_Beautified_a9e43898-8292-cc57-f6c3-0936079a7668" 
            class="OWAAutoLink eScj0 none" rel="noopener noreferrer" 
            data-ogsb="rgb(244,244,244)" data-ogsc="" style="margin:0px; 
            padding-right:1px; padding-left:1px; background-color:rgb(244,244,244); 
            border-radius:2px; text-align:left">
            <img width="16" height="16" 
            src="https://res.cdn.office.net/assets/mail/file-icon/png/folder_16x16.png" 
            style="width:16px; height:16px; padding-top:1px; padding-right:2px; 
            padding-bottom:2px; vertical-align:middle">{year_month}</a></span><br>"""
            }
        }
    }

    # Define the API endpoint URL
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/replyAll"

    # Define the headers
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Make the POST request to reply all
    response = requests.post(url, headers=headers, json=request_body)

    # Check the response
    if response.status_code == 202:
        print("Reply sent successfully.")
    else:
        print(f"Failed to send reply: {response.status_code} - {response.text}")
