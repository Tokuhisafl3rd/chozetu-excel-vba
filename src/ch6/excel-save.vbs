Dim excel, book, sheet, fso, path
' Excel���N�� --- (*1)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' �V�K�u�b�N�̃V�[�g�𓾂� --- (*2)
Set book = excel.Workbooks.Add()
Set sheet = book.Sheets(1)
' �V�[�g��A1�ɒl���� --- (*3)
sheet.Range("A1").Value = "����ɂ���"
' �X�N���v�g�̃t�H���_�Ƀu�b�N��ۑ� --- (*4)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)
book.SaveAs(path & "\hello.xlsx")
' Excel����� --- (*5)
excel.Quit



