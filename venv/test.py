from datetime import datetime
import re 
import flask
from flask import jsonify, send_file
from flask_cors import CORS
from aspose.email.storage.pst import PersonalStorage
import struct
from flask import request, Flask
from bs4 import BeautifulSoup
import whois
from docx import Document
import io
import dns.resolver
from ipwhois import IPWhois
import socket
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
import geoip2.database
import concurrent.futures

dns_cache = {}
ip_cache = {}

app = flask.Flask(__name__)
app.config["DEBUG"] = True
CORS(app)

UPLOAD_FOLDER = 'D:/others/pst files/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def decodeTime(time_bytes):
    timeValues = struct.unpack('<4I', time_bytes)
    timestamp = (timeValues[0] * 60 * 60 * 24 * 365) + (timeValues[1] * 60 * 60 * 24 * 30) + (timeValues[2] * 60 * 60 * 24) + (timeValues[3])
    return datetime.fromtimestamp(timestamp)

def getBasicDataFromPstHeader(file_path):
    with open(file_path, 'rb') as f:
        header_data = f.read(56)
        if len(header_data) < 56:
            raise ValueError("Invalid PST file: Header size is less than expected")
        signature, version, file_format, root_folder_id, creation_time_bytes, modification_time_bytes = struct.unpack('<4sHH16s16s16s', header_data)
        signature = signature.decode('ascii')
        creation_time = creation_time_bytes.rstrip(b'\x00').decode('utf-8', errors='ignore')
        modification_time = modification_time_bytes.rstrip(b'\x00').decode('utf-8', errors='ignore')

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

def extractImagesFromHtml(html_content):
    images = []
    soup = BeautifulSoup(html_content, 'html.parser')
    img_tags = soup.find_all('img')
    for img_tag in img_tags:
        src = img_tag.get('src')
        if src:
            images.append(src)
    return images

def extractMessagePreview(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    preview_text = soup.get_text()[:50]
    return preview_text

def extractLabelsFromHeaders(headers):
    label_patterns = [
        r"X-Label: (.*?)\r?\n",  # Example: X-Label: Important
        r"Keywords: (.*?)\r?\n",  # Example: Keywords: Work, Personal
        # Add more patterns as needed
    ]
    labels = []
    for pattern in label_patterns:
        match = re.search(pattern, headers, re.IGNORECASE)
        if match:
            labels.extend(match.group(1).split(','))
    return labels

def extractLabelsFromBody(body_html):
    label_pattern = r'<span class="label">(.*?)</span>'
    labels = re.findall(label_pattern, body_html)
    return labels

def isReply(headers):
    # Define patterns for reply headers
    reply_patterns = [
        r"In-Reply-To:.*\n",  # Example: In-Reply-To: <unique_id>
        r"References:.*\n",    # Example: References: <unique_id>
        r"X-MS-TNEF-Correlator:.*\n"  # Example: X-MS-TNEF-Correlator: <unique_id>
        # Add more patterns as needed
    ]
    
    for pattern in reply_patterns:
        if re.search(pattern, headers):
            return True
    return False

def isForward(headers):
    # Define patterns for forward headers
    forward_patterns = [
        r"X-Forwarded-Message-Id:.*\n",  # Example: X-Forwarded-Message-Id: <unique_id>
        r"X-MS-Exchange-Inbox-Rules-Loop:.*\n"  # Example: X-MS-Exchange-Inbox-Rules-Loop: <unique_id>
        # Add more patterns as needed
    ]
    
    for pattern in forward_patterns:
        if re.search(pattern, headers):
            return True
    return False

def isEncrypted(headers):
    # encryption_patterns = [r"Content-Type: multipart/encrypted", r"Content-Transfer-Encoding: base64"]
    encryption_patterns = [
        r"Content-Type: application/pkcs7-mime",  # S/MIME encryption
        r"Content-Type: application/x-pkcs7-mime",  # S/MIME encryption
        r"Content-Type: application/x-pkcs7-mime; smime-type=enveloped-data",  # S/MIME encryption
        r"Content-Type: application/octet-stream; name=\".*?\.p7m\"",  # S/MIME encryption
        r"Content-Type: application/x-msdownload; name=\".*?\.p7m\"",  # S/MIME encryption
        r"Content-Type: application/pkcs7-signature",  # S/MIME digital signature
        r"Content-Type: multipart/signed",  # S/MIME digital signature
        r"Content-Type: multipart/encrypted",  # PGP/MIME encryption
        r"Content-Type: application/pgp-encrypted",  # PGP/MIME encryption
        r"Content-Type: application/pgp-signature",  # PGP digital signature
        r"Content-Type: application/pgp-keys",  # PGP public key
        r"Content-Transfer-Encoding: base64",  # Base64-encoded content
        r"X-Encrypted: true",  # Custom header indicating encryption
        # Add more patterns as needed
    ]
    for pattern in encryption_patterns:
        if re.search(pattern, headers):
            return True
    return False

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
                    
                    arc_seal = mapi.headers.get('ARC-Seal')
                    arc_message_signature = mapi.headers.get('ARC-Message-Signature')
                    x_google_smtp_source = mapi.headers.get('X-Google-Smtp-Source')
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
                    is_reply = isReply(headers)
                    is_forward = isForward(headers)
                    is_encrypted = isEncrypted(headers)
                    
                    message_delivery_data = {
                        'subject': mapi.subject,
                        'proper_delivery': proper_delivery,
                        'headers': headers,
                        'is_reply': is_reply,
                        'is_forward': is_forward,
                        'is_encrypted': is_encrypted
                    }
                    folderData['message_delivery_data'].append(message_delivery_data)
# Appending Message Data
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
                        'authenticated': authenticated,
                        'arc_seal': arc_seal,
                        'arc_message_signature': arc_message_signature,
                        'x_google_smtp_source': x_google_smtp_source,
                        'preview_text': extractMessagePreview(mapi.body_html),
                        'labels': extractLabelsFromHeaders(headers) + extractLabelsFromBody(mapi.body_html),
                        'encryption_status': is_encrypted,
                        'message_delivery_data': message_delivery_data
                    }
                    folderData['messages'].append(message_data)
# Appending Contacts                    
                    if mapi.message_class.startswith('IPM.Contact'):
                        contact_data = {
                            'name': mapi.display_name,
                            'email': mapi.sender_email_address,
                            'phone': '',  # Add phone number if available
                        }
                        folderData['contacts'].append(contact_data)
# Appending Attachments                        
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

# ------------------------------------------WHOIS lookup Data-------------------------------------------------------
    
@app.route('/whoIs', methods=['GET'])
def getWhoIsDetails():
    domain = request.args.get('domain')
    if domain:
        try:
            whoIsInfo = whois.whois(domain)
            rawWhoIs = whoIsInfo.text
            parsedWhoIs = whoIsInfo
            return jsonify({'rawWhoIs': rawWhoIs, 'parsedWhoIs': parsedWhoIs})
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Domain parameter is required'}), 400

# ------------------------------------------PDF Generation-------------------------------------------------------

def draw_card_header(canvas, doc):
    canvas.saveState()
    width, height = letter
    # Draw the card header rectangle with background color
    canvas.setFillColorRGB(0.8, 0.8, 0.8)
    canvas.rect(doc.leftMargin, height - doc.topMargin, doc.width, 0.5 * inch, stroke=0, fill=1)

    # Draw the title text
    canvas.setFillColorRGB(0, 0, 0)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawString(doc.leftMargin + 10, height - doc.topMargin + 0.25 * inch - 10, "Report")
    canvas.restoreState()

def generatePDF(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=72)
    elements = []

    # Define styles for the document
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CardHeader', fontSize=14, leading=14, alignment=1, spaceAfter=20, backColor=colors.lightgrey))
    styles.add(ParagraphStyle(name='Heading2Custom', parent=styles['Heading1'], fontSize=12))
    styles.add(ParagraphStyle(name='BodyTextCustom', parent=styles['Normal'], spaceAfter=12))

    # General Information
    for key, value in data.items():
        if key in ["Folders List", "Servers Involved"]:
            continue  # Skip printing these here

        text = f"<b>{key}:</b> {value}"
        elements.append(Paragraph(text, styles['BodyTextCustom']))

    # Print Folders List table
    folders_list = data.get("Folders List")
    if folders_list:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("<b>Folders List</b>", styles['Heading2Custom']))
        folders_data = [["Folder Name", "Count"]]
        for item in folders_list.split(", "):
            parts = item.rsplit(" ", 1)  # Split by the last space to separate name and count
            if len(parts) == 2:
                folders_data.append([parts[0], parts[1]])
        table = Table(folders_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)

    # Print Servers Involved table
    servers_list = data.get("Servers Involved")
    if servers_list:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("<b>Servers Involved</b>", styles['Heading2Custom']))
        servers_data = [["Server"]]
        for i, server in enumerate(servers_list.split(";")):
            servers_data.append([server.strip()])
        table = Table(servers_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)

    # Build and save the PDF
    doc.build(elements, onFirstPage=draw_card_header)
    buffer.seek(0)
    return buffer

@app.route('/generate/report', methods=['POST'])
def generateReport():
    try:
        data = request.get_json()
        pdf_buffer = generatePDF(data)
        return send_file(pdf_buffer, as_attachment=True, download_name='report.pdf', mimetype='application/pdf')
    except Exception as e:
        print("Error:", str(e))  # Log the error message
        return jsonify({'error': str(e)}), 500
    
# ------------------------------------------DNS-------------------------------------------------------
    
def fetchDnsRecords(domain, record_type):
    cache_key = f"{domain}_{record_type}"
    if cache_key in dns_cache:
        return dns_cache[cache_key]
    try:
        resolver = dns.resolver.Resolver()
        answers = resolver.resolve(domain, record_type)
        records = [r.to_text() for r in answers]
        dns_cache[cache_key] = records
        return records
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.resolver.NoNameservers):
        return []
    except Exception as e:
        return str(e)

def getIpInfo(ip):
    if ip in ip_cache:
        return ip_cache[ip]
    try:
        obj = IPWhois(ip)
        res = obj.lookup_rdap()
        asn_info = {
            'asn': res.get('asn'),
            'asn_cidr': res.get('asn_cidr'),
            'asn_country_code': res.get('asn_country_code'),
            'asn_description': res.get('asn_description'),
            'asn_date': res.get('asn_date'),
            'network': res.get('network', {}).get('name'),
            'country': res.get('network', {}).get('country'),
        }

        # Geolocation information
        with geoip2.database.Reader('GeoLite2-City.mmdb') as reader:
            geo_info = reader.city(ip)
            location_info = {
                'city': geo_info.city.name,
                'country': geo_info.country.name,
                'latitude': geo_info.location.latitude,
                'longitude': geo_info.location.longitude
            }

        ip_info = {**asn_info, **location_info}
        ip_cache[ip] = ip_info
        return ip_info
    except Exception as e:
        return {'error': str(e)}

def getDetailedRecordInfo(records):
    detailed_records = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_to_record = {executor.submit(getIpInfo, record): record for record in records}
        for future in concurrent.futures.as_completed(future_to_record):
            record = future_to_record[future]
            try:
                ip_info = future.result()
                ip = socket.gethostbyname(record)
                detailed_records.append({
                    'record': record,
                    'ip': ip,
                    'details': ip_info
                })
            except Exception as e:
                detailed_records.append({
                    'record': record,
                    'error': str(e)
                })
    return detailed_records

@app.route('/nslookup', methods=['GET'])
def getNsLookupDetails():
    domain = request.args.get('domain')
    if domain:
        try:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                future_dns = {
                    'A': executor.submit(fetchDnsRecords, domain, 'A'),
                    'AAAA': executor.submit(fetchDnsRecords, domain, 'AAAA'),
                    'CNAME': executor.submit(fetchDnsRecords, domain, 'CNAME'),
                    'TXT': executor.submit(fetchDnsRecords, domain, 'TXT'),
                    'SPF': executor.submit(fetchDnsRecords, domain, 'SPF'),
                    'NS': executor.submit(fetchDnsRecords, domain, 'NS'),
                    'MX': executor.submit(fetchDnsRecords, domain, 'MX')
                }
                dns_data = {}
                for record_type, future in future_dns.items():
                    records = future.result()
                    if record_type in ['A', 'AAAA', 'MX']:
                        dns_data[record_type] = getDetailedRecordInfo(records)
                    else:
                        dns_data[record_type] = records

            return jsonify(dns_data)
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Domain parameter is required'}), 400



if __name__ == '__main__':
    app.run(host='localhost', port=5000)
    
