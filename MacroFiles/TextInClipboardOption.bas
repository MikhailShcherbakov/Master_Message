Attribute VB_Name = "TextInClipboardOption"
Option Explicit

Sub TextInClipboardOption()
Attribute TextInClipboardOption.VB_Description = "qwe"
Attribute TextInClipboardOption.VB_ProcData.VB_Invoke_Func = "у\n14"
'Ctrl+w Загружает форму с опциями для конструкции письма и
'дальнейшей записи в буфер обмена
    Application.ScreenUpdating = False
    
    If CorrectCell Then frmTextOption.Show
    
    Application.ScreenUpdating = True
End Sub
