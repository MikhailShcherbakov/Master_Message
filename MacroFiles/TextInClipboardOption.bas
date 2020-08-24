Attribute VB_Name = "TextInClipboardOption"
Option Explicit

Sub TextInClipboardOption()
Attribute TextInClipboardOption.VB_Description = "Выбор письма Ctrl + w"
Attribute TextInClipboardOption.VB_ProcData.VB_Invoke_Func = "w\n14"
'Ctrl+w Загружает форму с опциями для конструкции письма и
'дальнейшей записи в буфер обмена
    Application.OnKey "^{w}"
    Application.ScreenUpdating = False
    
    If CorrectCell And Correct.CorrectSheet Then frmTextOption.Show
    
    Application.ScreenUpdating = True
End Sub
