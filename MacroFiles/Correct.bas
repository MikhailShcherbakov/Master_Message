Attribute VB_Name = "Correct"
Option Explicit

Function CorrectCell() As Boolean

    If ActiveCell.Column <> 1 Or ActiveCell.value = "" Then
    
        MsgBox "Выделите коректную ячейку"
        CorrectCell = False
    
    Else
        
        CorrectCell = True
        
    End If
    
End Function

Function CorrectSheets() As Boolean

    On Error Resume Next
    Sheets("SavedPersons").Name
    
    If Err.number = 9 Then
        
        Sheets.Add(after:=Sheets(Sheets.Count)).Name = "SavedPersons"
        ActiveSheet.Cells(1, 1).value = "ФИО"
        ActiveSheet.Cells(1, 2).value = "Доп Информ"
        ActiveSheet.Cells(1, 3).value = "Дата добавления"
        
    End If
    
    CorrectSheets = True
    
End Function

Function CorrectSheet() As Boolean

    If ActiveSheet.Name = "WorkList" Then
        CorrectSheet = True
    Else
        CorrectSheet = False
    End If

End Function
