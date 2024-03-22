import datetime
import struct

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

# Usage
file_path = "D:\others\pst files\exportedPST.pst"
header_info = read_pst_header(file_path)
print(header_info)
