import win32com.client as win32

# Input and output file paths
input_file = r'C:\Users\sairaj.thikkisetty\Downloads\un.doc'
output_file = r'C:\Users\sairaj.thikkisetty\Downloads\un_modified.docx'

# Open Word application
word = win32.Dispatch('Word.Application')
word.Visible = False

# Open the document
doc = word.Documents.Open(input_file)

# Unhide text
def unhide_text_in_range(text_range):
    if text_range.Font.Hidden:
        text_range.Font.Hidden = False

# Unhide text in the whole document
for paragraph in doc.Paragraphs:
    unhide_text_in_range(paragraph.Range)

for table in doc.Tables:
    for row in table.Rows:
        for cell in row.Cells:
            for paragraph in cell.Range.Paragraphs:
                unhide_text_in_range(paragraph.Range)

# Save the document as .docx
doc.SaveAs(output_file, FileFormat=12)  # FileFormat=12 is for .docx
doc.Close()
word.Quit()

print(f"Document saved to {output_file}")
