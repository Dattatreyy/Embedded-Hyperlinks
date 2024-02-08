from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def extract_hyperlinks(docx_file):
    doc = Document(docx_file)
    hyperlinks = []

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for element in run._r:
                if element.tag.endswith('hyperlink'):
                    href = element.get('{%s}href' % nsdecls['r'])
                    if href is not None:
                        hyperlinks.append(href)

    return hyperlinks

# Full path to the DOCX file
docx_file = "C:\\Users\\Dattatrey\\Desktop\\abcd.docx"

# Extract hyperlinks from the DOCX file
hyperlinks = extract_hyperlinks(docx_file)

# Print the hyperlinks found in the document
print("Hyperlinks found in the document:")
for hyperlink in hyperlinks:
    print(hyperlink)
