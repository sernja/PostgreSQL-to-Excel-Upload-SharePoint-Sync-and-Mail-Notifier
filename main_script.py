import query_and_save_to_excel
import uploadfile_and_replaymail
import asyncio
import os
from azure.identity import InteractiveBrowserCredential
from datetime import date

#set date
today_date = date.today()
date_from = today_date.replace(day=1)
date_to = today_date

# set Output excel file path
output_excel_file = f"report from {date_from} to {date_to}.xlsx"

def get_query():
    # Database connection parameters
    host = "your_IP"
    dbname = "your_name"
    user = "your_user"
    password = "your_password"
    port = "5432"

    query_and_save_to_excel.query(host, dbname, user, password, port, date_from, date_to, output_excel_file)

def get_token():
    client_id = "your_client_id"
    tenant_id = "your_tenant_id"

    # Initialize the InteractiveBrowserCredential
    credential = InteractiveBrowserCredential(
        client_id=client_id,
        tenant_id=tenant_id,
    )

    scopes = [
        'User.Read',
        'Mail.Read',
        'Mail.Send',
        'Mail.ReadWrite',
        'Files.ReadWrite',
        'Files.ReadWrite.All',
        'Sites.ReadWrite.All'
    ]

    # Get the access token
    return credential.get_token(*scopes).token

async def put_file_sendmail(token):
    # Define file path and read the file stream
    file_path = os.path.abspath(output_excel_file)
    year_month = today_date.strftime("%Y-%m")
    link_share = "your_link_share_folder"
    file_name = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)
    site_id = "your_site_id"
    item_path = f"your_path_to_save_file/{file_name}"
    message_mail = f"get data from {date_from} to {date_to}"
    message_id = "your_message_id_in_mail"
    

    await uploadfile_and_replaymail.upload_large_file(token, file_name, site_id, item_path, file_path, file_size)
    uploadfile_and_replaymail.sendmail(token, message_mail, link_share, year_month, message_id)

get_query()
token = get_token()
asyncio.run(put_file_sendmail(token))
