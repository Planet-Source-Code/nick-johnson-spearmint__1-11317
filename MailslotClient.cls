VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailslotClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Zero As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function WriteFileSimple Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal Zero As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

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

'General Constants
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_EXISTING = 3

'Error constants
Private Const INVALID_HANDLE_VALUE = -1

Public Sub writeMessage(strMailSlotName As String, strMessage As String)
    Dim saSecurity As SECURITY_ATTRIBUTES
    Dim lngMailslotHandle As Long
    Dim lngResult As Long
    Dim lngMessageLen As Long
    Dim lngBytesSent As Long
    
    saSecurity.bInheritHandle = 0
    saSecurity.lpSecurityDescriptor = 0
    saSecurity.nLength = Len(saSecurity)
    
    'Get a handle to the mailslot
    lngMailslotHandle = CreateFile(strMailSlotName & vbNullChar, GENERIC_WRITE, FILE_SHARE_READ, saSecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    'If an error
    If lngMailslotHandle = INVALID_HANDLE_VALUE Then
        Err.Raise 5 + vbObjectError, "MailslotClient", "Could not open the Mailslot - Invalid handle value. Possibly the mail slot was not properly closed last time it was used."
        Exit Sub
    End If
    
    lngMessageLen = Len(strMessage) + 1
    lngResult = WriteFileSimple(lngMailslotHandle, strMessage + vbNullChar, lngMessageLen, lngBytesSent, 0)
    
    If lngResult = 0 Then
        'A problem sending
        Err.Raise 6 + vbObjectError, "MailslotClient", "Could not send the message - an error occurred calling Writefile: Error Number " & GetLastError()
        CloseHandle lngMailslotHandle
        Exit Sub
    End If
    
    If lngMessageLen <> lngBytesSent Then
        'More or less bytes were sent than expected.
        Err.Raise 7 + vbObjectError, "MailslotClient", "Did not send the correct length of the message. Expected " & Str(lngMessageLen) & " bytes, sent " & lngBytesSent
        CloseHandle lngMailslotHandle
        Exit Sub
    End If
    
    'Close the handle to the mailslot
    CloseHandle lngMailslotHandle
End Sub
