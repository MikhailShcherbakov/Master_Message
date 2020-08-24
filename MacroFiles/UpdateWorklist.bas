Attribute VB_Name = "UpdateWorklist"
Option Explicit

Sub Update_WorkList()
Attribute Update_WorkList.VB_Description = "Обновляет рабочий лист на основе нового скачаного файла"
Attribute Update_WorkList.VB_ProcData.VB_Invoke_Func = " \n14"
'   Процедура обновляет рабочий лист
    Application.ScreenUpdating = False
    
    Dim objWorkList As Worksheet, objExport As Worksheet
    
    Set objWorkList = ActiveWorkbook.Worksheets("WorkList")
    Set objExport = Worksheets(CopyNewSheetInWork(ActiveWorkbook))
    
    Dim strName As String, objFindRange As Range, objColumn As Range, objExColumn As Range
    
    Dim i&, strExport As String, lLastExRow As Long
    objExport.Activate
    lLastExRow = Range("A1").CurrentRegion.Rows.Count
    
    For i = 2 To lLastExRow
    
        strName = Cells(i, 1).value
        Set objFindRange = FindRangeInWork(strName, i)
        
        If Not objFindRange Is Nothing Then
        
            Call CorrectingDate(objExport, objWorkList, i:=i, FindRange:=objFindRange)
            
            With objFindRange.Font
            
                If .Bold = True And .Color = 5287936 Then 'Green
                    
                    .Color = 12611584 'Blue
                
                End If
                
            End With
        Else
        End If
        
    Next i
    
    Sheets("WorkList").Activate
    
    Application.ScreenUpdating = True
End Sub


Function CopyNewSheetInWork(GeneralBook As Workbook) As String
' копирует из нового файла информацию в рабочую книгу
' в новый созданный лист и возвращает имя листа
    Dim strPath As String, strFile As String, strDateTime As String
    strPath = ThisWorkbook.Path
    strFile = CStr(Date) & ".xlsx"
    
    On Error Resume Next
    Workbooks.Open FileName:=strPath & "\" & strFile

    If Err.number = 1004 Then 'File not found
        
            Dim vbFile
            vbFile = Application.GetOpenFilename _
                ("Excel files(*.xls*), *.xls*", 1, "Выбрать Excel файл", , MultiSelect:=False)
            
            If VarType(vbFile) = vbBoolean Then
                Exit Function
            Else
                
                Workbooks.Open FileName:=vbFile
            
            End If
    End If
    
    strDateTime = CStr(Date) & " " & CStr(Hour(Time)) & "." & CStr(Minute(Time))
    ActiveSheet.Name = strDateTime
    Sheets(strDateTime).Copy Before:=GeneralBook.Sheets("WorkList")
    Windows(strFile).Activate
    ActiveWorkbook.Close savechanges:=True, FileName:=strDateTime & ".xlsx"
    GeneralBook.Activate
    CopyNewSheetInWork = strDateTime
End Function

Function FindRangeInWork(strName As String, i As Long) As Range
    Dim lLastRow As Long, objFindRange As Range
    
    If IsMyFistLetter(strName) Then
        With Sheets("WorkList")
            
            lLastRow = .Cells(1, 1).CurrentRegion.Rows.Count
            Set objFindRange = .Range("A1:A" & CStr(lLastRow)).Find(strName)
            
            If objFindRange Is Nothing Then
                
                CopyRow (i)
                
            Else
            
                Set FindRangeInWork = objFindRange
            
            End If
        End With
    Else
        
        Set FindRangeInWork = Nothing
    
    End If
End Function

Function IsMyFistLetter(str As String) As Boolean
    Dim IsMyLetter As Boolean
    IsMyFistLetter = (Left(str, 1) = "С") Or (Left(str, 1) = "Т") Or (Left(str, 1) = "У")
End Function

Sub CopyRow(i As Long)
'   процедура копирует строку из Export in WorkList

    Dim LastRow As Long
    LastRow = Sheets("WorkList").Range("A1").End(xlDown).Row + 1
    'objExport.Range(objExCell.Address, "K" & CStr(i)).Copy
'    Range("A" & CStr(i), "K" & CStr(i)).Copy _
'        Sheets("WorkList").Range("A" & strLastWorkRow)
    Range(Cells(i, 1), Cells(i, 11)).Copy _
        Sheets("WorkList").Cells(LastRow, 1)
    
End Sub
'objExp As Worksheet DELETE
Sub CorrectingDate(objExp As Worksheet, objWork As Worksheet, i As Long, FindRange As Range)
    Dim arrColumnLett(1) As String, e As Variant
    arrColumnLett(0) = "H"
    arrColumnLett(1) = "K"
    For Each e In arrColumnLett
        Range(e & CStr(i)).Copy _
            objWork.Range(e & CStr(FindRange.Row))
    Next
End Sub

