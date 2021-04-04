import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
word = win32.gencache.EnsureDispatch('Word.Application')

excel.Visible = True
word.Visible = True
folder_path = r'C:\Users\LSO\Documents\GitHub\Replace-word'

book = excel.Workbooks.Open(folder_path + '\\dictionary.xlsx')


last_row = book.ActiveSheet.Cells(book.ActiveSheet.Rows.Count,1).End(-4162).Row
last_col = book.ActiveSheet.Cells(1,book.ActiveSheet.Columns.Count).End(-4159).Column

for item in range(2, last_row+1):
    doc = word.Documents.Open(folder_path + '\\template.docx')
    for p2r in range(1,len(doc.Paragraphs)+1):
        for i in range(2,last_col+1):
            print(i)
            doc.Paragraphs(p2r).Range.Text = doc.Paragraphs(p2r).Range.Text.replace(book.ActiveSheet.Cells(1,i).Text,
                                         book.ActiveSheet.Cells(item,i).Text)
    word.ActiveDocument.SaveAs(folder_path + "\\output\\document" + str(item-1)+".docx")
