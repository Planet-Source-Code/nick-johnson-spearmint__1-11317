VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailslotServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lngMailslotHandle As Long

'All declares containing "Zero as Long" have been modified to
'simplify the calls - just set zero to 0
Private Declare Function APICreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function ReadFileSimple Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal Zero As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'UDT's
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

'General constants
Private Const MAILSLOT_NO_MESSAGE = (-1)
Private Const MAILSLOT_WAIT_FOREVER = (-1)

'Error constants
Private Const INVALID_HANDLE_VALUE = -1

Public Sub CreateMailslot(ByVal strMailSlotName As String)
    Dim saSecurity As SECURITY_ATTRIBUTES
    
    If lngMailslotHandle <= 0 Then 'We have not already opened one
        saSecurity.bInheritHandle = 0
        saSecurity.lpSecurityDescriptor = 0
        saSecurity.nLength = Len(saSecurity)
        
        lngMailslotHandle = APICreateMailslot(strMailSlotName & vbNullChar, 0, 0, saSecurity)
        
        If lngMailslotHandle = INVALID_HANDLE_VALUE Then
            On Error Resume Next
            Err.Raise 1 + vbObjectError, "MailslotServer", "Could not create the Mailslot - Invalid handle value. Possibly the mail slot was not properly closed last time it was used."
            Exit Sub
        End If
    End If
End Sub

Public Function getMessage() As String
    'Gets the next message, if any. If no messages are waiting,
    'getMessage returns an empty string.
    Dim lngResult As Long
    Dim blnResult As Boolean
    Dim lngptrMessage As Long
    Dim lngNextMessageLength As Long
    Dim lngTotalMessages As Long
    Dim strMessage As String
    Dim lngBytesRead As Long
    
    If lngMailslotHandle > 0 Then
        'Check for messages
        lngResult = GetMailslotInfo(lngMailslotHandle, 0, lngNextMessageLength, lngTotalMessages, 0)
        
        If lngResult = 0 Then
            'Something went wrong with GetMailslotInfo!
            Err.Raise 2 + vbObjectError, "MailslotServer", "An error occurred calling GetMailslotInfo: Error number " & GetLastError()
            Exit Function
        End If
        
        If lngTotalMessages = 0 Then
            'No messages waiting
            Exit Function
        End If
        
        'Get the first waiting message (oldest)
        strMessage = String(lngNextMessageLength + 1, " ")
        blnResult = ReadFileSimple(lngMailslotHandle, strMessage, lngNextMessageLength, lngBytesRead, 0)

        If lngResult = 0 Then
            'Another error!
            Err.Raise 3 + vbObjectError, "MailslotServer", "An error occurred calling ReadFile: Error number " & GetLastError()
            Exit Function
        End If
        
        If lngBytesRead <> lngNextMessageLength Then
            'For some reason we got more or less than the number
            'of bytes in the message.
            Err.Raise 4 + vbObjectError, "MailslotServer", "Did not read the correct length of the message. Expected " & Str(lngNextMessageLength) & " bytes, recieved " & lngBytesRead
            Exit Function
        End If
        
        getMessage = strMessage
    End If
End Function

Public Property Get MailslotHandle() As Long
    MailslotHandle = lngMailslotHandle
End Property

Public Property Get MessageCount() As Long
    Dim lngResult As Long
    Dim lngTotalMessages As Long
    
    If lngMailslotHandle > 0 Then
        lngResult = GetMailslotInfo(lngMailslotHandle, 0, 0, lngTotalMessages, 0)
        MessageCount = lngTotalMessages
    End If
End Property

Private Sub Class_Terminate()
    If lngMailslotHandle > 0 Then
        CloseHandle lngMailslotHandle
    End If
End Sub
