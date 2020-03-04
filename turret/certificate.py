import docx

attendees = ['CJ Salgo', 'Mac Aliwanag', 'Warren Cedro', 'Gerbert Reyes', 'Emman', 'Jefferze Briggs']

def create_certificate(name):
    doc = docx.Document('certificate.docx')
    doc.paragraphs[7].runs[4].text = name
    doc.save('files/' + name + '.docx')

for name in attendees:
    create_certificate(name)
