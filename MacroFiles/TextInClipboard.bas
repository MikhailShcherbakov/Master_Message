Attribute VB_Name = "TextInClipboard"
Option Explicit

Sub TextInClipboard()
Attribute TextInClipboard.VB_Description = "Записывает в буфер обмена письмо с именем и отчеством абитур-нта из выделенной ячейки  Ctrl + q"
Attribute TextInClipboard.VB_ProcData.VB_Invoke_Func = "q\n14"
'Ctrl+q Записывает в Буфер обмена полный текст письма
'с просьбой выслать все необходимые документы
    Application.OnKey "^{q}"
    Application.ScreenUpdating = False
    
    If Correct.CorrectCell And Correct.CorrectSheet Then
        
        Dim obj As New clsMessage
        
        obj.Caption = "Все"
        obj.Name = "AllDocuments"
        
        Dim c As New Collection
        c.Add obj
        MessageInClipBoard.SetMessage c
    End If
    
    Application.ScreenUpdating = True
End Sub
