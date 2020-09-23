VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Keeps track of a colleciton of users online/offline activity, including all the computers they are on"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Users"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private pstrDomain As String 'local copy
Private pstrServer As String 'local copy
Private pusrUsers As Users
Private pintUpdateInterval As Integer 'local copy

'RaiseEvent
Public Event UserOnline(usrUser As Object)
Attribute UserOnline.VB_Description = "The specified user is now online"
Public Event UserOffline(usrUser As Object)
Attribute UserOffline.VB_Description = "A user has gone offline."
Public Event CompChange(usrUser As Object)
Attribute CompChange.VB_Description = "The number of computers a user is on has changed."

'API declarations
Private Declare Function NetWkstaGetInfo100 Lib "netapi32" Alias "NetWkstaGetInfo" (ServerName As Byte, ByVal Level As Long, BufPtr As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function NetGetDCName Lib "netapi32.dll" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)

'User types
Private Type WKSTA_INFO_100
    dw_platform_id As Long
    ptr_computername As Long
    ptr_langroup As Long
    dw_ver_major As Long
    dw_ver_minor As Long
End Type

Public Property Let UpdateInterval(ByVal vData As Integer)
Attribute UpdateInterval.VB_Description = "The interval (ms) between updates. Default 10 000ms (10 seconds)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpdateInterval = 5
    pintUpdateInterval = vData
End Property
Public Property Get UpdateInterval() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpdateInterval
    UpdateInterval = pintUpdateInterval
End Property

Public Property Get Users() As Users
    If pusrUsers Is Nothing Then
        Set pusrUsers = New Users
    End If

    Set Users = pusrUsers
End Property

Private Sub Class_Initialize()
    'Set domain to the domain this computer is in, and
    'server to the PDC of that domain.
    Dim lngReturn As Long
    
    'Get domain name
    lngReturn = fnGetDomainName(pstrDomain)
    If lngReturn <> 0 Then
        Err.Raise Err.Number + vbObjectError
    End If

    'Get PDC name
    lngReturn = fnGetPDCName("", pstrDomain, pstrServer)
    If lngReturn <> 0 Then
        Err.Raise Err.Number + vbObjectError
    End If
    
    'Set other defaults
    pintUpdateInterval = 10000

    'Initialise the timer
    SetTimer 0&, &H5000&, pintUpdateInterval, AddressOf TimerTriggered
End Sub

Private Sub Class_Terminate()
    Set pusrUsers = Nothing
    KillTimer 0&, &H5000&
End Sub

Public Property Let Server(ByVal vData As String)
Attribute Server.VB_Description = "The server to get the info from. Defaults to the Primary Domain Controller for the specified domain"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    pstrServer = vData
End Property
Public Property Get Server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    Server = pstrServer
End Property

Public Property Let Domain(ByVal vData As String)
Attribute Domain.VB_Description = "The domain to use for all users. Defaults to the first domain in the list."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Domain = 5
    pstrDomain = vData
End Property
Public Property Get Domain() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Domain
    Domain = pstrDomain
End Property

Public Sub UpdateNow()
    'Update the user list now.
    'Called externally and by the timer sub.
End Sub

Private Sub TimerTriggered(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
    'Ignore the parameters - just a timer
    MsgBox "Timer triggered!"
End Sub

Private Function fnGetDomainName(strDomain) As Long
Dim lngReturn As Long
Dim lngTemp As Long
Dim strTemp As String
Dim bDomain(99) As Byte
Dim bServer() As Byte
Dim lngBuffPtr As Long
Dim typeWorkstation As WKSTA_INFO_100

    fnGetDomainName = 0
    
    bServer = "" + vbNullChar
    
    lngReturn = NetWkstaGetInfo100( _
        bServer(0), _
        100, _
        lngBuffPtr)
        
    If lngReturn <> 0 Then
        fnGetDomainName = lngReturn
        Exit Function
    End If
        
    CopyMem typeWorkstation, _
        ByVal lngBuffPtr, _
        Len(typeWorkstation)
        
    lngTemp = typeWorkstation.ptr_langroup
    
    lngReturn = PtrToStr( _
        bDomain(0), _
        lngTemp)
        
    strTemp = Left( _
        bDomain, _
        StrLen(lngTemp))

    strDomain = strTemp
End Function

Private Function fnGetPDCName(strServer As String, strDomain As String, strPDCName As String) As Long
Dim lngReturn As Long
Dim lngDCNPtr As Long
Dim bDomain() As Byte
Dim bServer() As Byte
Dim bPDCName(100) As Byte

    fnGetPDCName = 0
    
    bServer = strServer & vbNullChar
    bDomain = strDomain & vbNullChar
    lngReturn = NetGetDCName( _
        bServer(0), _
        bDomain(0), _
        lngDCNPtr)
    
    If lngReturn <> 0 Then
        fnGetPDCName = lngReturn
        Exit Function
    End If
    
    lngReturn = PtrToStr(bPDCName(0), lngDCNPtr)
    lngReturn = NetAPIBufferFree(lngDCNPtr)
    strPDCName = bPDCName()
    strPDCName = Mid$(strPDCName, 1, InStr(strPDCName, Chr$(0)) - 1)
End Function

