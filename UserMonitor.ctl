VERSION 5.00
Begin VB.UserControl UserMonitor 
   CanGetFocus     =   0   'False
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   930
   ScaleWidth      =   1245
   Begin VB.Timer tmrUpdate 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "UserMonitor.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "UserMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private pstrDomain As String 'local copy
Private pstrServer As String 'local copy
Private pusrUsers As New Users

'RaiseEvent
Public Event UserOnline(usrUser As User)
Public Event UserOffline(usrUser As User)
Public Event CompChange(usrUser As User)

'API declarations
Private Declare Function NetWkstaGetInfo100 Lib "netapi32" Alias "NetWkstaGetInfo" (ServerName As Byte, ByVal Level As Long, BufPtr As Any) As Long
Private Declare Function NetGetDCName Lib "netapi32.dll" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Private Declare Function NetSessionEnum Lib "netapi32.dll" (ServerName As Byte, UncClientName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PreMaxLen As Long, EntriesRead As Long, TotalEntries As Long, Resume_Handle As Long) As Long

'User types
Private Type WKSTA_INFO_100
    dw_platform_id As Long
    ptr_computername As Long
    ptr_langroup As Long
    dw_ver_major As Long
    dw_ver_minor As Long
End Type

Private Type Session_Info_10
   sesi10_cname                       As Long
   sesi10_username                    As Long
   sesi10_time                        As Long
   sesi10_idle_time                   As Long
End Type

'Public Type USER_INFO_10_API
'  Name As Long
'  Comment As Long
'  UsrComment As Long
'  FullName As Long
'End Type

Public Property Let UpdateInterval(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpdateInterval = 5
    If vData = 0 Then
        tmrUpdate.Enabled = False
    Else
        If Ambient.UserMode Then tmrUpdate.Enabled = True
        tmrUpdate.Interval = vData
    End If
End Property
Public Property Get UpdateInterval() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpdateInterval
    UpdateInterval = tmrUpdate.Interval
End Property

Public Property Get Users() As Users
    Set Users = pusrUsers
End Property

Private Sub tmrUpdate_Timer()
    UpdateNow
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Set domain to the domain this computer is in, and
    'server to the PDC of that domain.
    Dim lngReturn As Long
    
    'If it's runtime
    If Ambient.UserMode Then
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
        pusrUsers.Server = pstrServer
        
        tmrUpdate.Enabled = True
    Else
        tmrUpdate.Enabled = False
    End If
End Sub

Private Sub UserControl_Resize()
    'Set the size
    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height
End Sub

Private Sub usercontrol_Terminate()
    Set pusrUsers = Nothing
End Sub

Public Property Let Server(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    If Ambient.UserMode Then
        pstrServer = vData
        pusrUsers.Server = vData
    Else
        Err.Raise Number:=31013, Description:="Property is read-only at run time."
    End If
End Property
Public Property Get Server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    Server = pstrServer
End Property

Public Property Let Domain(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Domain = 5
    If Ambient.UserMode Then
        pstrDomain = vData
    Else
        Err.Raise Number:=31013, Description:="Property is read-only at run time."
    End If
End Property
Public Property Get Domain() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Domain
    Domain = pstrDomain
End Property

Public Sub UpdateNow()
    'Update the user list now.
    'Called externally and by the timer sub.
    Dim lngReturn As Long
    Dim usrCurrentUser As User
    Dim intNumComputers As Integer
    Dim a As Integer, b As Integer
    
    For a = 1 To pusrUsers.Count
        Set usrCurrentUser = pusrUsers.Item(a)
        'Get the number of computers they are on
        intNumComputers = usrCurrentUser.Computers.Count
        
        'Empty the collection
        For b = usrCurrentUser.Computers.Count To 1 Step -1
            usrCurrentUser.Computers.Remove b
        Next b
        
        'Call SessionEnum to update this users computers
        lngReturn = SessionEnum(pstrServer, "", usrCurrentUser.UserName)
        If lngReturn <> 0 And lngReturn <> 2221 Then
            'Error 2221 means they are not logged in.
            Err.Raise lngReturn
        End If
        
        'If they have not just been added
        If Not usrCurrentUser.JustAdded Then
            'Are they now online/offline/num computers changed?
            If intNumComputers < usrCurrentUser.Computers.Count And intNumComputers = 0 Then
                'They were offline, now online
                RaiseEvent UserOnline(usrCurrentUser)
            ElseIf intNumComputers > usrCurrentUser.Computers.Count And usrCurrentUser.Computers.Count = 0 Then
                'They were online, now offline
                RaiseEvent UserOffline(usrCurrentUser)
            ElseIf intNumComputers <> usrCurrentUser.Computers.Count Then
                'Number of computers has changed
                RaiseEvent CompChange(usrCurrentUser)
            End If
        Else
            usrCurrentUser.JustAdded = False
        End If
    Next a
End Sub

Public Function updateUser(strUsername As String)
    Dim usrUser As User
    Dim intNumComputers As Integer
    Dim a As Integer
    Dim lngReturn As Long
    
    Set usrUser = pusrUsers(strUsername)
    intNumComputers = usrUser.Computers.Count
    
    'Empty the computers collection
    For a = 1 To intNumComputers
        usrUser.Computers.Remove a
    Next a
    
    'Call sessionenum for this computer
    lngReturn = SessionEnum(pstrServer, "", strUsername)
    If lngReturn <> 0 And lngReturn <> 2221 Then
        'Error 2221 means they are not logged in.
        Err.Raise lngReturn
    End If
    
    'If they have not just been added
    If Not usrUser.JustAdded Then
        'Are they now online/offline/num computers changed?
        If intNumComputers < usrUser.Computers.Count And intNumComputers = 0 Then
            'They were offline, now online
            RaiseEvent UserOnline(usrUser)
        ElseIf intNumComputers > usrUser.Computers.Count And usrUser.Computers.Count = 0 Then
            'They were online, now offline
            RaiseEvent UserOffline(usrUser)
        ElseIf intNumComputers <> usrUser.Computers.Count Then
            'Number of computers has changed
            RaiseEvent CompChange(usrUser)
        End If
    Else
        usrUser.JustAdded = False
    End If
End Function

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

Private Function SessionEnum(sServerName As String, sClientName As String, sUserName As String) As Long
   Dim bFirstTime           As Boolean
   Dim lRtn                 As Long
   Dim ServerName()         As Byte
   Dim UncClientName()      As Byte
   Dim UserName()           As Byte
   Dim lptrBuffer           As Long
   Dim lEntriesRead         As Long
   Dim lTotalEntries        As Long
   Dim lResume              As Long
   Dim I                    As Integer
   Dim psComputerName               As String
   Dim psUserName                   As String
   Dim plActiveTime                 As Long
   Dim plIdleTime                   As Long
   Dim typSessionInfo()             As Session_Info_10
    
    lPrefmaxlen = 65535
     
    ServerName = sServerName & vbNullChar
    UncClientName = sClientName & vbNullChar
    UserName = sUserName & vbNullChar
    
Do
   lRtn = NetSessionEnum(ServerName(0), UncClientName(0), UserName(0), 10, lptrBuffer, lPrefmaxlen, lEntriesRead, lTotalEntries, lResume)
     
    If lRtn <> 0 Then
        SessionEnum = lRtn
        Exit Function
    End If

If lTotalEntries <> 0 Then


    ReDim typSessionInfo(0 To lEntriesRead - 1)
    'ReDim msSessionInfo(0 To lEntriesRead - 1)
     
    CopyMem typSessionInfo(0), ByVal lptrBuffer, Len(typSessionInfo(0)) * lEntriesRead
     
    For I = 0 To lEntriesRead - 1
        pusrUsers(sUserName).Computers.Add PointerToStringW(typSessionInfo(I).sesi10_cname), typSessionInfo(I).sesi10_time, typSessionInfo(I).sesi10_idle_time
    Next I
    End If
Loop Until lEntriesRead = lTotalEntries
   
    If lptrBuffer <> 0 Then
        NetAPIBufferFree lptrBuffer
    End If
End Function

Public Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
    
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMem Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function
