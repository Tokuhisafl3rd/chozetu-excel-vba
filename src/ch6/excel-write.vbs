' Excel���N��
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' �V�K�u�b�N���쐬
Set book = excel.Workbooks.Add()
' ���삷��V�[�g��I��
Set sheet = book.Worksheets("Sheet1")
' �Z��A1�Ɋi����l������
sheet.Range("A1").Value = "�ӂ��҂͉��������ΕׂȐl�͖��������"
' ���O��t���ĕۑ�
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
book.SaveAs ThisPath & "\test.xlsx"
' Excel�����
excel.Quit
