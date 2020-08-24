VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTextOption 
   Caption         =   "Содержание письма"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   OleObjectBlob   =   "frmTextOption.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTextOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim number As Integer
Const Subscription As String = "С уважением," + vbCrLf + "Информационный центр магистратуры СПбГУ по направлениям ""Филология"" и ""Лингвистика"""


Private Sub ChPassport_Click()

End Sub

Private Sub CommandButton1_Click()
'подготавливает данные для сборки письма

    Dim Frames As New Collection, item As Control
    For Each item In frmTextOption.Controls
    
        If TypeOf item Is frame Then Frames.Add item
        
    Next
    
    Dim Fr As frame
    Set Fr = Get_Frame(Frames)
    If Fr Is Nothing Then
    
        MsgBox "Неверные данные"
        Exit Sub
        
    Else
        
        MessageInClipBoard.SetMessage Get_Objects(Fr)
        
    End If
    
    Unload frmTextOption
End Sub

Function Get_Frame(Frames As Collection) As frame
'возвращает корректную рамку
    
    Dim i&, j&
    For i = 1 To Frames.Count
        If Value_Controls(Frames(i)) And (i <> Frames.Count) Then
            For j = i + 1 To Frames.Count
                If Value_Controls(Frames(j)) Then
                
                    Set Get_Frame = Nothing
                    Cls_frames Frames
                    Exit Function
                    
                End If
            Next j
            
            Set Get_Frame = Frames(i)
            Exit Function
            
        ElseIf i = Frames.Count Then
            
            Set Get_Frame = Frames(i)
            Exit Function
            
        End If
    Next i
    
End Function

Function Value_Controls(frame As frame) As Boolean
    '
    Dim item As Control
    For Each item In frame.Controls
        If item.value Then
        
            Value_Controls = True
            Exit Function
            
        End If
    Next
    
    Value_Controls = False
End Function

Sub Cls_frames(Frames As Collection)
'очищает все рамки
    Dim frame As frame
    For Each frame In Frames
        Dim contr As Control
        For Each contr In frame.Controls
        
            contr.value = False
            
        Next
    Next
End Sub


Function Get_Objects(f As frame) As Collection
'return clsMessage

    Dim item As Control, objColl As New Collection
    Dim obj As clsMessage
    
    For Each item In f.Controls
        If item.value And InStr(item.Name, "1st") > 0 Then
            If objColl.Count = 0 Then
                
                Set obj = New clsMessage
                obj.Init item
                objColl.Add obj, Key:=obj.Name
            
            Else
                
                Set obj = New clsMessage
                obj.Init item
                objColl.Add obj, Key:=obj.Name, Before:=1
            
            End If
        ElseIf item.value Then
            
            Set obj = New clsMessage
            obj.Init item
            objColl.Add obj, Key:=obj.Name
        
        End If
    Next
    
    Set Get_Objects = objColl
End Function

Private Sub LastMessage1st_Click()
    If LastMessage1st.value Then
        ChApplication.value = False
        ChApplication.Enabled = False
        
        ChConsent.value = False
        ChConsent.Enabled = False
    Else
        ChApplication.Enabled = True
        
        ChConsent.Enabled = True
    End If
End Sub


Private Sub Send_Mail_Click()

End Sub
Sub Send_Massege(txt As String)
    Const CDO_Cnf = "http://schemas.microsoft.com/cdo/configuration/"
    Dim oCDOCnf As New CDO.Configuration, oCDOMsg As Object
    Dim SMTPserver As String, sUsername As String, sPass As String, sMsg As String
    Dim sTo As String, sFrom As String, sSubject As String, sBody As String
'    On Error Resume Next
    
    Dim objCDOmas As New CDO.Message
    
    
    
    SMTPserver = ""
    sUsername = ""
    sPass = ""

    sTo = ActiveCell.Offset(0, 3).value
    sFrom = ""
    sSubject = "Документы в магистратуру СПбГУ"

    Set oCDOCnf = CreateObject("CDO.Configuration")
    With oCDOCnf.Fields
        .item(CDO_Cnf & "sendusing") = 2
        .item(CDO_Cnf & "smtpauthenticate") = 1
        .item(CDO_Cnf & "smtpserver") = SMTPserver
        '???? ?????????? ??????? SSL
        '.Item(CDO_Cnf & "smtpserverport") = 465 '??? ??????? ? Gmail 465
        '.Item(CDO_Cnf & "smtpusessl") = True
        '=====================================
        .item(CDO_Cnf & "sendusername") = sUsername
        .item(CDO_Cnf & "sendpassword") = sPass
        .Update
    End With

    Set objCDOmas = CreateObject("CDO.Message")
    With objCDOmas
        Set .Configuration = oCDOCnf
        .BodyPart.Charset = "koi8-r"
        .From = sFrom
        .To = sTo
        .Subject = sSubject
        .TextBody = txt
        .Send
    End With

    Dim strAdress As String
    strAdress = InputBox("Дополнительный e-mail")
    If strAdress <> "" Then
    With objCDOmas
        Set .Configuration = oCDOCnf
        .BodyPart.Charset = "koi8-r"
        .From = sFrom
        .To = strAdress
        .Subject = sSubject
        .TextBody = txt
        .Send
    End With
    End If
    
'    Select Case Err.number
'        Case -2147220973: sMsg = "Нет доступа к сети Интернет"
'        Case -2147220975: sMsg = "Отказано в доступе SMTP"
'        Case 0: sMsg = "Письмо отправлен"
'        Case Else: sMsg = "Ошибка номер: " & Err.number & vbNewLine & "Опиание ошибки: " & Err.Description
'    End Select
'    MsgBox sMsg, vbInformation, "www.Excel-VBA.ru"
    Set oCDOMsg = Nothing: Set oCDOCnf = Nothing
End Sub

Private Sub UserForm_Click()

End Sub
