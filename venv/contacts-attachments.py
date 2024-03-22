from types import MethodDescriptorType
import flask
from flask import jsonify
from flask import json
from flask_cors import CORS
from flask import Flask
from aspose.email.storage.pst import PersonalStorage
from werkzeug.utils import secure_filename
from aspose.email.mapi import MapiMessage

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'D:/others/pst files/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/getExtractedData', methods=['GET'])
def getExtractedData():
    pst = PersonalStorage.from_file("D:\others\pst files\sample.pst")
    folder_info_collection = pst.root_folder.get_sub_folders()
    result = []

    for folder_info in folder_info_collection:
        folder_data = {
            'folder_name': folder_info.display_name,
            'total_items': folder_info.content_count,
            'total_unread_items': folder_info.content_unread_count,
            'contacts': [],  # Initialize contacts list
            'attachments': []  # Initialize attachments list
        }

        folder = pst.root_folder.get_sub_folder(folder_info.display_name)

        if folder_info.content_count:
            messages = folder.get_contents(0, folder_info.content_count)

            for message_info in messages:
                mapi = pst.extract_message(message_info)

                # Extract contact information
                if mapi.message_class.startswith('IPM.Contact'):
                    contact_data = {
                        'name': mapi.display_name,
                        'email': mapi.sender_email_address,
                        'phone': '',  # Add phone number if available
                        # Add additional contact properties as needed
                    }
                    folder_data['contacts'].append(contact_data)

                # Extract attachment information
                for attachment in mapi.attachments:
                    attachment_data = {
                        'name': attachment.name,
                        'size': attachment.size,
                        # Add additional attachment properties as needed
                    }
                    folder_data['attachments'].append(attachment_data)

        result.append(folder_data)

    return jsonify(result)


if __name__ == '__main__':
    app.run(debug=True)
