Dim excel, book, sheet, fso, path
' Excel���N�� --- (*1)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' �u�b�N���J�� --- (*2)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(path & "\hello.xlsx")
' �V�[�g�𓾂� --- (*3)
Set sheet = book.Sheets(1)
' �V�[�g��A1�̒l���擾 --- (*4)
MsgBox sheet.Range("A1").Value
' Excel����� --- (*5)
excel.Quit



