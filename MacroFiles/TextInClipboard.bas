Attribute VB_Name = "TextInClipboard"
Option Explicit

Sub TextInClipboard()
Attribute TextInClipboard.VB_Description = "Записывает в буфер обмена письмо с именем и отчеством абитур-нта из выделенной ячейки  Ctrl + q"
Attribute TextInClipboard.VB_ProcData.VB_Invoke_Func = "q\n14"
'Ctrl+q Записывает в Буфер обмена полный текст письма
'с просьбой выслать все необходимые документы
    Application.ScreenUpdating = False
    
    If Correct.CorrectCell Then
        
        Dim infWord As New Collection
        infWord.Add "Все"
        Dim FileName As Collection
        FileName.Add "AllDocuments"
        MessageInClipBoard.SetMessage FileName, infLabel:=infWord
    
    End If
    
    Application.ScreenUpdating = True
End Sub
