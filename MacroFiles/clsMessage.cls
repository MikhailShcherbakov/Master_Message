VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strName As String
Private strCaption As String

Public Property Get Name() As String
    Name = strName
End Property

Public Property Let Name(value As String)
    strName = value
End Property

Public Property Get Caption() As String
    Caption = strCaption
End Property

Private Property Let Caption(value As String)
    strCaption = value
End Property

Public Sub Init(cntr As Control)
    
    Name = cntr.Name
    Caption = cntr.Caption
    
End Sub
