' Excel���N��
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' �����u�b�N��ǂ�
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(ThisPath & "\test.xlsx")
' ���삷��V�[�g��I��
Set sheet = book.Worksheets("Sheet1")
' �Z��A1�̒l��ǂݎ��
v = sheet.Range("A1").Value
MsgBox v
book.Close
excel.Quit
