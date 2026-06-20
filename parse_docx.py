import zipfile
import xml.etree.ElementTree as ET

def get_docx_text(path):
    """Extracts text from a docx file using zipfile and xml parsing."""
    try:
        # docx is a zip file, text is in word/document.xml
        with zipfile.ZipFile(path) as docx:
            xml_content = docx.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            # XML namespaces
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            paragraphs = []
            # Find all paragraph elements
            for p in root.findall('.//w:p', ns):
                text_runs = []
                # Find all text elements in the paragraph
                for r in p.findall('.//w:r', ns):
                    t = r.find('.//w:t', ns)
                    if t is not None and t.text:
                        text_runs.append(t.text)
                
                if text_runs:
                    paragraphs.append("".join(text_runs))
            
            return "\n".join(paragraphs)
    except Exception as e:
        return f"Error reading docx: {str(e)}"

# Read the document and write it to text file
text = get_docx_text("ros_resume_new.docx")
with open("docx_content.txt", "w", encoding="utf-8") as f:
    f.write(text)

print("Extraction completed. File size:", len(text))
