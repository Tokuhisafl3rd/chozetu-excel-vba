Dim excel, book, fso, path
' VBScript�̂���p�X�𒲂ׂ� --- (*1)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)

' Excel���N�����ău�b�N���J�� --- (*2)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
Set book = excel.Workbooks.Open(path & "\macro.xlsm")

' �}�N�������s --- (*3)
excel.Application.Run "Sheet1.��{�\��", 30

' Excel����� --- (*4)
excel.Quit
