Attribute VB_Name = "TextInClipboardOption"
Option Explicit

Sub TextInClipboardOption()
Attribute TextInClipboardOption.VB_Description = "qwe"
Attribute TextInClipboardOption.VB_ProcData.VB_Invoke_Func = "�\n14"
'Ctrl+w ��������� ����� � ������� ��� ����������� ������ �
'���������� ������ � ����� ������
    Application.ScreenUpdating = False
    
    If CorrectCell Then frmTextOption.Show
    
    Application.ScreenUpdating = True
End Sub
