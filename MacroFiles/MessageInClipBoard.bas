Attribute VB_Name = "MessageInClipBoard"
Option Explicit
Const subscription As String = "С уважением," + vbCrLf + "Информационный центр магистратуры СПбГУ по направлениям ""Филология"" и ""Лингвистика"""

Sub SetMessage(FileName As Collection, infLabel As Collection)
'Собирает письмо, записывает информацию и дату в ячейки об отправке писем
    Dim intStr As Integer, strName As String
    intStr = InStr(ActiveCell.value, " ")
    strName = Mid(ActiveCell.value, intStr)
    
    Dim strLetter As String
    strLetter = SexPerson(strName) & strName & "!" & vbCrLf & vbCrLf & _
        CorpusMassege(FileName) & vbCrLf & vbCrLf & subscription
    
    PutTxtInClp strLetter
    Inform FileName, ActiveCell
End Sub

Function SexPerson(NamePerson As String) As String
    'Функция возвращает обращение в необходимом роде
    Dim Response As Integer
    
    Response = MsgBox("Уважаемая " & NamePerson & vbCrLf & vbCrLf & "Пол мужской?", vbYesNo)
    If Response = vbYes Then
        SexPerson = "Уважаемый"
    ElseIf Response = vbNo Then
        SexPerson = "Уважаемая"
    End If
    
End Function

Function CorpusMassege(FileName As Collection)
        
    Dim Path As String, strMessage As String, strLine As String
    Path = "C:\Users\Abbadon\Documents\Магистратура-прием 2020\Макросы\Files\"
    
    If Parent(FileName(1), frmTextOption.Frame1) And Not InStr(FileName(1), "1st") Then
    
    Else
    Open Path & FileName(1) & ".txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strLine
            strMessage = strMessage & vbCrLf & strLine
        Loop
    Close #1
    End If
    
    Dim strReduct As String
    If FileName.Count = 1 Or InStr(FileName(1), "1st") Then
        'only LastMessage1st or LastMessage1st is fist in collection or any collection with fixed order
        
    ElseIf Parent(FileName(1), frmTextOption.Frame1) Then ' was Or InStr(FileName(1), "1st") > 0
        'without LastMessage1st

        strReduct = Reduct_Text(strReduct, FileName)

    Else
    End If
    
    CorpusMassege = strMessage & vbCrLf & vbCrLf & strReduct
End Function


Function Parent(FileName As String, frame As frame) As Boolean
    If ControlExist(FileName) Then
        With frmTextOption
            If .Controls(FileName).Parent Is frame Then
                
                Parent = True
                Exit Function
                
            End If
        End With
    Else: Parent = False
    End If
End Function

Function ControlExist(strName As String) As Boolean
    Dim er As String
    On Error Resume Next
        er = frmTextOption(strName).Name
        ControlExist = (Err.number = 0)
End Function

Function Reduct_Text(text As String, FileName As Collection) As String
    
    Dim Reduct As String, FinText As String
    Dim FName As New Collection
    FName.Add "TextOption"
    Reduct = CorpusMassege(FName)
    
    Dim start As Integer, intro As String
    start = InStr(Reduct, "intro")
    Reduct = Replace(Reduct, "intro", "")
    intro = Left(Reduct, start)
    
    Dim lines As String
    
    If InStr(FileName(1), "1st") > 0 Then
    
        lines = Complate_Fragments(Reduct, FileName, 2)
        'вставляем необходимые фрагменты текста
        Reduct_Text = intro & vbCrLf & vbCrLf & lines
        Exit Function
        
    Else
        lines = Complate_Fragments(Reduct, FileName, 1)
        Reduct_Text = intro & vbCrLf & vbCrLf & lines
        Exit Function
    End If
    
End Function

Function Complate_Fragments(Reduct As String, checkName As Collection, position As Integer) As String
    'Возвращает отредактированный текст
    
    Dim Fragment As String, i&, FinText As String
    Dim number As Integer
    If frmTextOption.ChApplication And frmTextOption.ChConsent Then
    
        Dim start As Integer
        start = InStr(Reduct, "ChApplicationChConsent")
        FinText = Mid(Reduct, start, InStr(start, Reduct, ";") - start + 1) '-------
        FinText = Replace(FinText, "ChApplicationChConsent", "")
    
    End If
    For i = position To checkName.Count
        If checkName(i) = "ChApplication" Or checkName(i) = "ChConsent" Then
        Else
            number = InStr(Reduct, checkName(i))
            number = Len(Reduct) - number
            Fragment = Right(Reduct, number + 1)
            Fragment = Left(Fragment, InStr(Fragment, ";"))
            Fragment = Replace(Fragment, checkName(i), "")
            FinText = FinText & vbCrLf & Fragment
        End If
    Next i
    Complate_Fragments = FinText
End Function


Sub PutTxtInClp(txt As String)
'Принимает txt и записывает текст в буфер обмена
    
    Dim myObject As New DataObject
    myObject.SetText (txt)
    myObject.PutInClipboard
    
End Sub


Sub Inform(FileName As Collection, rng As Range)
'Записывает в ячейки информацию и дату об отправки письма
    Dim i&
    For i = 1 To FileName.Count
        Dim cntr As Controls

        If Not frmTextOption.Controls(FileName(i)).Parent Is frmTextOption.Frame1 Or FileName(i) = "LastMessage1st" Then
            RecordCellInform FileName(i), 13, rng
        Else
            RecordCellInform FileName(i), 11, rng
        End If
    Next i
    
End Sub

Sub RecordCellInform(Name As String, Column As Integer, rng As Range)
    rng.Offset(0, Column).Select
    Dim Word As String
    Word = frmTextOption.Controls(Name).Caption
    
    If InStr(Selection.value, Word) = 0 Then
    
        Selection.value = Selection.value & " " & Word
        
    End If
    
    Selection.Offset(0, 1).value = CStr(Date) & " " & CStr(Time)
End Sub

