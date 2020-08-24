Attribute VB_Name = "MessageInClipBoard"
Option Explicit
Const Subscription As String = "С уважением," + vbCrLf + "Информационный центр магистратуры СПбГУ по направлениям ""Филология"" и ""Лингвистика"""

Sub SetMessage(objects As Collection)
'Собирает письмо, записывает информацию и дату в ячейки об отправке писем
    Dim intStr As Integer, strName As String
    intStr = InStr(ActiveCell.value, " ")
    strName = Mid(ActiveCell.value, intStr)
    
    Dim strLetter As String
    strLetter = SexPerson(strName) & strName & "!" & vbCrLf & vbCrLf
    strLetter = strLetter & CorpusMassege(objects)
    strLetter = strLetter & vbCrLf & vbCrLf & Subscription
        
    PutTxtInClp strLetter
    Inform objects, ActiveCell
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

Function CorpusMassege(objects As Collection) As String
        
    Dim strReduct As String, strMessage As String
     
    If objects.Count = 1 And InStr(objects(1).Name, "1st") Then
        'only LastMessage1st or LastMessage1st is fist in collection or any collection with fixed order
        
        strMessage = Open_File(objects)
        
    ElseIf Parent(objects(1).Name, frmTextOption.Frame1) Then
        ' (LastMessage1st первый в коллекции)

        strReduct = Reduct_Text(objects)

    Else
        'collection with one element
        strMessage = Open_File(objects)
            
    End If
    
    CorpusMassege = strMessage & vbCrLf & vbCrLf & strReduct
End Function

Function Open_File(objects As Collection) As String
    
    Dim Path As String, strMessage As String, strLine As String
    Path = ThisWorkbook.Path & "\Files\"

    Open Path & objects(1).Name & ".txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strLine
            strMessage = strMessage & vbCrLf & strLine
        Loop
    Close #1
        Open_File = strMessage
End Function

Function Parent(ControlName As String, myFrame As frame) As Boolean

    If ControlExist(ControlName) Then
        With frmTextOption
            If .Controls(ControlName).Parent Is myFrame Then
                
                Parent = True
                Exit Function
                
            End If
        End With
    Else: Parent = False
    End If

End Function

Function ControlExist(ControlName As String) As Boolean
    Dim er As String
    On Error Resume Next
        er = frmTextOption(ControlName).Name
        ControlExist = (Err.number = 0)
End Function

Function Reduct_Text(objects As Collection) As String
    
    Dim Reduct As String, FinText As String
    Dim FName As New Collection, obj As New clsMessage
    
    obj.Name = "TextOption"
    FName.Add obj
    Reduct = Open_File(FName)
    
    Dim start As Integer, intro As String
    start = InStr(Reduct, "intro")
    Reduct = Replace(Reduct, "intro", "")
    intro = Left(Reduct, start)
    
    Dim lines As String
    
    If InStr(objects(1).Name, "1st") > 0 Then
        
        FinText = Open_File(objects)
        lines = Complate_Fragments(Reduct, objects, position:=2)
        'вставляем необходимые фрагменты текста
        Reduct_Text = FinText & vbCrLf & vbCrLf & intro & vbCrLf & vbCrLf & lines
        Exit Function
        
    Else
    
        lines = Complate_Fragments(Reduct, objects, 1)
        Reduct_Text = intro & vbCrLf & vbCrLf & lines
        Exit Function
        
    End If
    
End Function

Function Complate_Fragments(Reduct As String, objects As Collection, position As Integer) As String
    'Возвращает отредактированный текст
    
    Dim Fragment As String, i&, FinText As String
    Dim start As Integer, strAppConsent As String
    start = InStr(Reduct, "ChApplicationChConsent")
    strAppConsent = Mid(Reduct, start, InStr(start, Reduct, ";") - start + 1)

    If frmTextOption.ChApplication And frmTextOption.ChConsent Then
    
        FinText = strAppConsent
        FinText = Replace(FinText, "ChApplicationChConsent", "")
        objects("ChApplication").Name = ""
        objects("ChConsent").Name = ""
        
    Else
        
        Reduct = Replace(Reduct, strAppConsent, "")

    End If
    For i = position To objects.Count
        If objects(i).Name <> "" Then 'обходим ситуацию одновременно включенных заяв и согл
            
            start = InStr(Reduct, objects(i).Name)
            start = Len(Reduct) - start
            Fragment = Right(Reduct, start + 1)
            Fragment = Left(Fragment, InStr(Fragment, ";"))
            Fragment = Replace(Fragment, objects(i).Name, "")
            FinText = FinText & vbCrLf & Fragment
        
        End If
    Next i
    
    If frmTextOption.ChApplication And frmTextOption.ChConsent Then
        'снова втавляем чтобы правильно добавлялась информация об отправке
        objects("ChApplication").Name = "ChApplication"
        objects("ChConsent").Name = "ChConsent"
    
    End If
    
    Complate_Fragments = FinText
End Function


Sub PutTxtInClp(txt As String)
'Принимает txt и записывает текст в буфер обмена
    
    Dim myObject As New DataObject
    myObject.SetText (txt)
    myObject.PutInClipboard
    
End Sub


Sub Inform(objects As Collection, rng As Range)
'Записывает в ячейки информацию и дату об отправки письма
    Dim i&
    For i = 1 To objects.Count
        Dim cntr As Controls

        'переделать это условие!
        If objects(i).Caption = "Все" Or (Parent(objects(i).Name, frmTextOption.Frame1) And Not (objects(i).Name = "LastMessage1st")) Then
            
            RecordCellInform objects(i).Caption, 11, rng
        
        Else
            
            RecordCellInform objects(i).Caption, 13, rng

        End If
    Next i
    
End Sub

Sub RecordCellInform(Name As String, Column As Integer, rng As Range)
    rng.Offset(0, Column).Select

    If InStr(Selection.value, Name) = 0 Then
    
        Selection.value = Selection.value & " " & Name
        
    End If
    
    Selection.Offset(0, 1).value = CStr(Date) & " " & CStr(Time)
End Sub

