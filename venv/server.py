from datetime import datetime 
import flask
from flask import jsonify
from flask_cors import CORS
from aspose.email.storage.pst import PersonalStorage
import struct
from flask import request
import json

app = flask.Flask(__name__)
app.config["DEBUG"] = True
CORS(app)

UPLOAD_FOLDER = 'D:/others/pst files/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def decode_time(time_bytes):
    # Decode time bytes to integer values
    time_values = struct.unpack('<4I', time_bytes)
    
    # Convert to human-readable timestamp
    timestamp = (time_values[0] * 60 * 60 * 24 * 365) + (time_values[1] * 60 * 60 * 24 * 30) + (time_values[2] * 60 * 60 * 24) + (time_values[3])
    
    # Convert timestamp to datetime object
    return datetime.fromtimestamp(timestamp)

def read_pst_header(file_path):
    with open(file_path, 'rb') as f:
        header_data = f.read(56)  # Read the first 40 bytes of the file (PST header size)
        if len(header_data) < 56:
            raise ValueError("Invalid PST file: Header size is less than expected")

        # Parse the header fields
        signature, version, file_format, root_folder_id, creation_time_bytes, modification_time_bytes = struct.unpack('<4sHH16s16s16s', header_data)

        # Convert binary data to human-readable format
        signature = signature.decode('ascii')
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


@app.route('/getExtractedData', methods=['GET'])
def getExtractedData():
    file_path = request.args.get('file')
    with PersonalStorage.from_file(file_path) as pst:
        folder_info_collection = pst.root_folder.get_sub_folders()
        result = []
        header_info = read_pst_header(file_path)
        for folder_info in folder_info_collection:
            folder_data = {
                'folder_name': folder_info.display_name,
                'total_items': folder_info.content_count,
                'total_unread_items': folder_info.content_unread_count,
                'messages': [],
                'contacts': [],
                'attachments': [],
                'header_data': [header_info]
            }
            folder = pst.root_folder.get_sub_folder(folder_info.display_name)
            if folder_info.content_count:
                messages = folder.get_contents(0, folder_info.content_count)
                for message_info in messages:
                    mapi = pst.extract_message(message_info)
                    # for prop in dir(mapi):
                    #     print(f"{prop}: {getattr(mapi, prop)}")
                    received_spf = mapi.headers.get('Received-SPF')
                    received = mapi.headers.get('Received')
                    if received_spf is not None:
                        sender_email_server = mapi.headers.get(received_spf)
                    else:
                        sender_email_server = None

                    if received is not None:
                        way_to_recipient_server = mapi.headers.get(received)
                    else:
                        way_to_recipient_server = None

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
                        'way_to_recipient_server': way_to_recipient_server
                    }
                    folder_data['messages'].append(message_data)
                    if mapi.message_class.startswith('IPM.Contact'):
                        contact_data = {
                            'name': mapi.display_name,
                            'email': mapi.sender_email_address,
                            'phone': '',  # Add phone number if available
                        }
                        folder_data['contacts'].append(contact_data)
                    for attachment in mapi.attachments:
                        if hasattr(attachment, 'name'):
                            attachment_data = {
                                'name': attachment.name,
                                'size': attachment.size
                            }
                            folder_data['attachments'].append(attachment_data)
            result.append(folder_data)
        header_info_serializable = {
            'signature': header_info['signature'],
            'version': header_info['version'],
            'file_format': header_info['file_format'],
            'root_folder_id': header_info['root_folder_id'].decode('utf-8', errors='ignore'),
            'creation_time': header_info['creation_time'],
            'modification_time': header_info['modification_time']
        }
        for folder_data in result:
            folder_data['header_data'] = header_info_serializable
        return jsonify(result)


if __name__ == '__main__':
    app.run(host='localhost', port=5000)
    
