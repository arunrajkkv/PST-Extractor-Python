from datetime import datetime 
import flask
from flask import jsonify
from flask_cors import CORS
from aspose.email.storage.pst import PersonalStorage
import struct
from flask import request
from bs4 import BeautifulSoup

app = flask.Flask(__name__)
app.config["DEBUG"] = True
CORS(app)

UPLOAD_FOLDER = 'D:/others/pst files/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def decodeTime(time_bytes):
    timeValues = struct.unpack('<4I', time_bytes) # Decode time bytes to integer values
    # Convert to human-readable timestamp
    timestamp = (timeValues[0] * 60 * 60 * 24 * 365) + (timeValues[1] * 60 * 60 * 24 * 30) + (timeValues[2] * 60 * 60 * 24) + (timeValues[3])
    return datetime.fromtimestamp(timestamp) # Convert timestamp to datetime object

def getBasicDataFromPstHeader(file_path):
    with open(file_path, 'rb') as f:
        header_data = f.read(56)  # Read the first 40 bytes of the file (PST header size)
        if len(header_data) < 56:
            raise ValueError("Invalid PST file: Header size is less than expected")
        # Parse the header fields
        signature, version, file_format, root_folder_id, creation_time_bytes, modification_time_bytes = struct.unpack('<4sHH16s16s16s', header_data)
        signature = signature.decode('ascii') # Convert binary data to human-readable format
        creation_time = creation_time_bytes.rstrip(b'\x00').decode('utf-8', errors='ignore')  # Remove null bytes and decode using 'utf-8'
        modification_time = modification_time_bytes.rstrip(b'\x00').decode('utf-8', errors='ignore')  # Remove null bytes and decode using 'utf-8'

        return {
            'signature': signature,
            'version': version,
            'file_format': file_format,
            'root_folder_id': root_folder_id,
            'creation_time': creation_time,
            'modification_time': modification_time
        }
        
def extractAttachmentsFromMessage(message):
    attachments = []
    for attachment in message.attachments:
        attachment_data = {
            'name': attachment.name,
            'size': attachment.size,
            'content': attachment.data
        }
        attachments.append(attachment_data)
    return attachments


def checkForProperEmailDelivery(mapi):
    dsn = mapi.headers.get('Disposition-Notification-To')
    if dsn:
        return 'Successful'
    else:
        return 'Unsuccessful'

# def extractImagesFromHtml(html_content):
#     images = []
#     soup = BeautifulSoup(html_content, 'html.parser')
#     img_tags = soup.find_all('img')
#     for img_tag in img_tags:
#         src = img_tag.get('src')
#         if src:
#             images.append(src)
#     return images
def extractImagesFromHtml(html_content):
    images = []
    soup = BeautifulSoup(html_content, 'html.parser')
    img_tags = soup.find_all('img')
    for img_tag in img_tags:
        src = img_tag.get('src')
        if src:
            # If the src is a URL, you can directly add it to the images list
            if src.startswith('http') or src.startswith('https'):
                images.append(src)
            # If the src is a base64 encoded image, you can extract and decode it
            elif src.startswith('data:image'):
                # Extract the base64 encoded image data
                src_parts = src.split(',')
                if len(src_parts) > 1:
                    image_data = src_parts[1]
                    images.append(image_data)  # Add the decoded image data to the images list
    return images


@app.route('/getExtractedData', methods=['GET'])
def getExtractedData():
    file_path = request.args.get('file')
    with PersonalStorage.from_file(file_path) as pst:
        folderList = pst.root_folder.get_sub_folders()
        result = []
        headerInfo = getBasicDataFromPstHeader(file_path)
        for folder in folderList:
            folderData = {
                'folder_name': folder.display_name,
                'total_items': folder.content_count,
                'total_unread_items': folder.content_unread_count,
                'messages': [],
                'contacts': [],
                'attachments': [],
                'header_data': [headerInfo],
                'message_delivery_data': []
            }
            folder = pst.root_folder.get_sub_folder(folder.display_name)
            if folder.content_count:
                messages = folder.get_contents(0, folder.content_count)
                for message_info in messages:
                    mapi = pst.extract_message(message_info)
                    print(mapi.headers)
                    received_spf = mapi.headers.get('Received-SPF')
                    received = mapi.headers.get('Received')
                    message_id = mapi.headers.get('Message-Id')
                    received_line = f"Received: {received}\n\n"
                    message_id_line = f"Message-Id: {message_id}\n\n"
                    headers = message_id_line + received_line
                    if received_spf is not None:
                        sender_email_server = mapi.headers.get(received_spf)
                    else:
                        sender_email_server = None

                    if received is not None:
                        way_to_recipient_server = mapi.headers.get(received)
                    else:
                        way_to_recipient_server = None
                    
                    sender_ip_address = received.split('[')[-1].split(']')[0] if received else None
                    receiver_ip_address = received.split('from ')[-1].split(' ')[0] if received else None
                    authenticated = True if received_spf else False
                    
                    proper_delivery = checkForProperEmailDelivery(mapi)
                    message_delivery_data = {
                        'subject': mapi.subject,
                        'proper_delivery': proper_delivery,
                        'headers': headers
                    }
                    folderData['message_delivery_data'].append(message_delivery_data)

                    message_data = {
                        'subject': mapi.subject,
                        'sender_name': mapi.sender_name,
                        'sender_email': mapi.sender_email_address,
                        'to': mapi.display_to,
                        'cc': mapi.display_cc,
                        'bcc': mapi.display_bcc,
                        'delivery_time': str(mapi.delivery_time),
                        'body': mapi.body_html,
                        'client_submit_time': mapi.client_submit_time,
                        'sender_address_type': mapi.sender_address_type,
                        'sender_smtp_address': mapi.sender_smtp_address,
                        'conversation_topic': mapi.conversation_topic,
                        'sender_email_server': sender_email_server,
                        'way_to_recipient_server': way_to_recipient_server,
                        'images': extractImagesFromHtml(mapi.body_html),
                        'sender_ip_address': sender_ip_address,
                        'receiver_ip_address': receiver_ip_address,
                        'authenticated': authenticated
                    }
                    folderData['messages'].append(message_data)
                    if mapi.message_class.startswith('IPM.Contact'):
                        contact_data = {
                            'name': mapi.display_name,
                            'email': mapi.sender_email_address,
                            'phone': '',  # Add phone number if available
                        }
                        folderData['contacts'].append(contact_data)
                    for attachment in mapi.attachments:
                        if hasattr(attachment, 'name'):
                            attachment_data = {
                                'name': attachment.name,
                                'size': attachment.size
                            }
                            folderData['attachments'].append(attachment_data)
            result.append(folderData)
        headerInfo_serializable = {
            'signature': headerInfo['signature'],
            'version': headerInfo['version'],
            'file_format': headerInfo['file_format'],
            'root_folder_id': headerInfo['root_folder_id'].decode('utf-8', errors='ignore'),
            'creation_time': headerInfo['creation_time'],
            'modification_time': headerInfo['modification_time']
        }
        for folderData in result:
            folderData['header_data'] = headerInfo_serializable
        return jsonify(result)


if __name__ == '__main__':
    app.run(host='localhost', port=5000)
    
