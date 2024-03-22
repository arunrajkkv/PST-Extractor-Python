@app.route('/getExtractedData', methods=['GET'])
def getExtractedData():
    pst = PersonalStorage.from_file("D:\others\pst files\source.pst")
    folder_info_collection = pst.root_folder.get_sub_folders()
    result = []
    file_path = "D:\others\pst files\source.pst"
    header_info = read_pst_header(file_path)
    print(header_info)
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
        # folder_data['header_data'].append(header_info)
        folder = pst.root_folder.get_sub_folder(folder_info.display_name)
        if folder_info.content_count:
            messages = folder.get_contents(0, folder_info.content_count)
            for message_info in messages:
                mapi = pst.extract_message(message_info)
                message_data = {
                    'subject': mapi.subject,
                    'sender_name': mapi.sender_name,
                    'sender_email': mapi.sender_email_address,
                    'to': mapi.display_to,
                    'cc': mapi.display_cc,
                    'bcc': mapi.display_bcc,
                    'delivery_time': str(mapi.delivery_time),
                    'body': mapi.body
                }
                folder_data['messages'].append(message_data)
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
                    if hasattr(attachment, 'name'):
                        attachment_data = {
                            'name': attachment.name,
                            'size': attachment.size,
                            # Add additional attachment properties as needed
                        }
                        folder_data['attachments'].append(attachment_data)
        result.append(folder_data)
        # Convert header_info to a JSON serializable format
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