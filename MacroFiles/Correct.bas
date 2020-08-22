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
        
    End If
    
    CorrectSheets = True
    
End Function
