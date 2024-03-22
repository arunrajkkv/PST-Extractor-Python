from aspose.email.storage.pst import PersonalStorage

# pst = PersonalStorage.from_file("D:\others\pst files\sample.pst")
# pst = PersonalStorage.from_file("D:\others\pst files\example.pst")
pst = PersonalStorage.from_file("D:\others\pst files\sample.pst")


folderInfoCollection = pst.root_folder.get_sub_folders()
for folderInfo in folderInfoCollection:
	print("Folder: " + folderInfo.display_name)
	print("Total Items: " + str(folderInfo.content_count))
	print("Total Unread Items: " + str(folderInfo.content_unread_count))
	print("----------------------")
	folder = pst.root_folder.get_sub_folder(f"{folderInfo.display_name}")
	# Extracts messages starting from 10th index top and extract total 100 messages
	if folderInfo.content_count:
		messages = folder.get_contents(0, folderInfo.content_count)
		print(messages[0])
		for messageInfo in messages:
			mapi = pst.extract_message(messageInfo)
			print("Subject: " + mapi.subject)
			print("Sender name: " + mapi.sender_name)
			print("Sender email address: " + mapi.sender_email_address)
			print("To: ", mapi.display_to)
			print("Cc: ", mapi.display_cc)
			print("Bcc: ", mapi.display_bcc)
			print("Delivery time: ", str(mapi.delivery_time))
			print("Body: " + mapi.body)
