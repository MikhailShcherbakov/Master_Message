Attribute VB_Name = "TextInClipboardOption"
Option Explicit

Sub TextInClipboardOption()
Attribute TextInClipboardOption.VB_Description = "����� ������ Ctrl + w"
Attribute TextInClipboardOption.VB_ProcData.VB_Invoke_Func = "w\n14"
'Ctrl+w ��������� ����� � ������� ��� ����������� ������ �
'���������� ������ � ����� ������
    Application.OnKey "^{w}"
    Application.ScreenUpdating = False
    
    If CorrectCell And Correct.CorrectSheet Then frmTextOption.Show
    
    Application.ScreenUpdating = True
End Sub
