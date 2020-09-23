Attribute VB_Name = "modGlobal"
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Public Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function NetUserGetInfo Lib "netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Type USER_INFO_10_API
  Name As Long
  Comment As Long
  UsrComment As Long
  fullName As Long
End Type

Public Type USERINFO_2_API
  usri2_name As Long
  usri2_password As Long
  usri2_password_age As Long
  usri2_priv As Long
  usri2_home_dir As Long
  usri2_comment As Long
  usri2_flags As Long
  usri2_script_path As Long
  usri2_auth_flags As Long
  usri2_full_name As Long
  usri2_usr_comment As Long
  usri2_parms As Long
  usri2_workstations As Long
  usri2_last_logon As Long
  usri2_last_logoff As Long
  usri2_acct_expires As Long
  usri2_max_storage As Long
  usri2_units_per_week As Long
  usri2_logon_hours As Long
  usri2_bad_pw_count As Long
  usri2_num_logons As Long
  usri2_logon_server As Long
  usri2_country_code As Long
  usri2_code_page As Long
End Type

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

Public Function localUserName() As String
    Dim strUsername As String * 255
    Dim lngLength As Long
    Dim lngResult As Long
    
    lngLength = 255
    lngResult = GetUserName(strUsername, lngLength)
    If lngResult <> 1 Then
        MsgBox "An error occurred with localUserName() - No " & Str(lngResult), vbCritical, "Error in getUserName"
        Exit Function
    End If
    localUserName = Left(strUsername, lngLength - 1)
End Function

Public Function getUserHome() As String
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    Dim userInfo As USERINFO_2_API
    Dim strName As String
    Dim a As Integer
    
    'set variables
    baServerName = frmMain.umUsers.Server & Chr$(0)
    baUserName = localUserName & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 2, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
      getUserHome = "h:\"
      Exit Function
    End If

    'Turn the pointer into a variable
    CopyMem userInfo, ByVal lngptrUserInfo, Len(userInfo)
    
    getUserHome = PointerToStringW(userInfo.usri2_home_dir) & "\"
    NetAPIBufferFree lngptrUserInfo
End Function

Public Function localComputerName() As String
    Dim lngLength As Long
    Dim strCompName As String
    
    strCompName = String(256, " ")
    lngLength = 256
    GetComputerName strCompName, lngLength
    localComputerName = Left(strCompName, lngLength)
End Function

Public Function getRealName(strServer As String, strUsername As String) As String
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    Dim userInfo As USER_INFO_10_API
    Dim strName As String
    Dim a As Integer
    
    'set variables
    baServerName = strServer & Chr$(0)
    baUserName = strUsername & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 10, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
      getRealName = ""
      Exit Function
    End If

    'Turn the pointer into a variable
    CopyMem userInfo, ByVal lngptrUserInfo, Len(userInfo)
    
    strName = PointerToStringW(userInfo.fullName)
    NetAPIBufferFree lngptrUserInfo
    getRealName = strName
End Function
