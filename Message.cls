VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pstrFrom As String
Private pstrFromName As String
Private pstrTo As String
Private pstrMessage As String
Private pstrFromComputer As String
Private pstrMessageType As String * 1
Private pstrOtherData As String

Public Property Get From() As String
    From = pstrFrom
End Property
Public Property Let From(vData As String)
    pstrFrom = vData
    'For winpopup messages, where no from computer is specified
    If pstrFromComputer = "" Then pstrFromComputer = pstrFrom
End Property

Public Property Get msgTo() As String
    msgTo = pstrTo
End Property
Public Property Let msgTo(vData As String)
    pstrTo = vData
End Property

Public Property Get Message() As String
    Message = pstrMessage
End Property
Public Property Let Message(vData As String)
    pstrMessage = vData
End Property

Public Property Get fromName() As String
    If pstrFromName = "" Then
        pstrFromName = getRealName(frmMain.umUsers.Server, pstrFrom)
    End If
    fromName = pstrFromName
End Property

Public Property Get FromComputer() As String
    FromComputer = pstrFromComputer
End Property
Public Property Let FromComputer(vData As String)
    pstrFromComputer = vData
End Property

Public Property Get MessageType() As String
    MessageType = pstrMessageType
End Property
Public Property Let MessageType(vData As String)
    pstrMessageType = Left(vData, 1)
End Property

'Otherdata (Additional Data field) usually contains either
'a password (for admin functions), or modifiers for the message
'-eg. This message is a ping reply, this message is a
'confirmation message, or this message requires confirmation.
Public Property Get OtherData() As String
    OtherData = pstrOtherData
End Property
Public Property Let OtherData(vData As String)
    pstrOtherData = vData
End Property

Public Property Get MessageSource() As String
    MessageSource = pstrTo & vbNullChar & _
                    pstrFrom & vbNullChar & _
                    pstrMessage & vbNullChar & _
                    pstrFromComputer & vbNullChar & _
                    pstrMessageType & pstrOtherData
End Property

Public Property Let MessageSource(vData As String)
    Dim colSource As Collection
    
    Set colSource = TokeniseString(vData, vbNullChar)
    On Error Resume Next
    pstrTo = colSource(1)
    pstrFrom = colSource(2)
    pstrMessage = colSource(3)
    pstrFromComputer = colSource(4)
    pstrMessageType = Left(colSource(5), 1)
    pstrOtherData = Right(colSource(5), Len(colSource(5)) - 1)
    Err.Clear
End Property

Private Function TokeniseString(ByVal strData As String, strToken As String) As Collection
    Set TokeniseString = New Collection
    
    Dim lngNextTokenPos As Long
    
    lngNextTokenPos = InStr(1, strData, strToken)
    While lngNextTokenPos <> 0
        TokeniseString.Add Left(strData, lngNextTokenPos - 1)
        strData = Right(strData, Len(strData) - lngNextTokenPos)
        lngNextTokenPos = InStr(1, strData, strToken)
    Wend
End Function

Private Sub Class_Initialize()
    'Default settings for message
    pstrMessageType = "M"
End Sub
