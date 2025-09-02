import zipfile
import xml.etree.ElementTree as ET

def read_word_document(file_path):
    with zipfile.ZipFile(file_path, 'r') as docx:
        # Read document.xml
        with docx.open('word/document.xml') as xml_file:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            
            # Extract text
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }
            
            paragraphs = root.findall('.//w:p', namespaces)
            for p in paragraphs:
                texts = p.findall('.//w:t', namespaces)
                paragraph_text = ''.join(t.text for t in texts if t.text)
                if paragraph_text.strip():
                    print(paragraph_text)

read_word_document('Matthias.docx')