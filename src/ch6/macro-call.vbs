' Excel���N��
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' �}�N�����`�����u�b�N��ǂ�
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(ThisPath & "\macro-message-call.xlsm")
' �}�N�����N��
excel.Application.Run "Sheet1.ShowMessage"
' Excel���I��
excel.Quit
Set excel = Nothing
