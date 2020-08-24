Attribute VB_Name = "TextInClipboard"
Option Explicit

Sub TextInClipboard()
Attribute TextInClipboard.VB_Description = "Çàïèñûâàåò â áóôåð îáìåíà ïèñüìî ñ èìåíåì è îò÷åñòâîì àáèòóð-íòà èç âûäåëåííîé ÿ÷åéêè  Ctrl + q"
Attribute TextInClipboard.VB_ProcData.VB_Invoke_Func = "q\n14"
'Ctrl+q Çàïèñûâàåò â Áóôåð îáìåíà ïîëíûé òåêñò ïèñüìà
'ñ ïðîñüáîé âûñëàòü âñå íåîáõîäèìûå äîêóìåíòû
    Application.ScreenUpdating = False
    
    If Correct.CorrectCell Then
        
        obj.Caption = "Âñå"
        obj.Name = "AllDocuments"
        
        Dim c As New Collection
        c.Add obj
        MessageInClipBoard.SetMessage c
    
    End If
    
    Application.ScreenUpdating = True
End Sub
