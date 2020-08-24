Attribute VB_Name = "SavedPerson"
Option Explicit

Sub SavedPerson()
Attribute SavedPerson.VB_Description = "—охран€ет абитуриента в одтельный лист и помечает зеленым цветом"
Attribute SavedPerson.VB_ProcData.VB_Invoke_Func = "e\n14"
'добавл€ет абитуриента в лист сохранненных,
' помеча€ его зеленым цветом
    
    Application.ScreenUpdating = False
    
    If Correct.CorrectCell And Correct.CorrectSheets Then
        
        
        Dim lLastRow As Long, c As Range
        lLastRow = Worksheets("SavedPersons").Cells(rows.Count, 1).End(xlUp).Row + 1
        Set c = Worksheets("SavedPersons").Cells(lLastRow, 1)
        
        If c.Offset(-1, 2).value <> Date Then _
            lLastRow = lLastRow + 1

        With ActiveCell
            .Copy _
                Worksheets("SavedPersons").Cells(lLastRow, 1)
            .Font.Bold = True
            .Font.Color = 5287936 'Green
        End With
        Worksheets("SavedPersons").Cells(lLastRow, 3).value = Date
        
    End If
    
    Application.ScreenUpdating = True
End Sub
