Attribute VB_Name = "TextInClipboard"
Option Explicit

Sub TextInClipboard()
Attribute TextInClipboard.VB_Description = "���������� � ����� ������ ������ � ������ � ��������� ������-��� �� ���������� ������  Ctrl + q"
Attribute TextInClipboard.VB_ProcData.VB_Invoke_Func = "q\n14"
'Ctrl+q ���������� � ����� ������ ������ ����� ������
'� �������� ������� ��� ����������� ���������
    Application.ScreenUpdating = False
    
    If Correct.CorrectCell Then
        
        Dim infWord As New Collection
        infWord.Add "���"
        Dim FileName As Collection
        FileName.Add "AllDocuments"
        MessageInClipBoard.SetMessage FileName, infLabel:=infWord
    
    End If
    
    Application.ScreenUpdating = True
End Sub
