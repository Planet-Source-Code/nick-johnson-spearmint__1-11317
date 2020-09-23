VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOnlineUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Online Users"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "z.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCollapse 
      Caption         =   "&Collapse All"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3400
      Width           =   2175
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "&Expand All"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3400
      Width           =   2175
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      _Version        =   393217
      Indentation     =   450
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextSend 
         Caption         =   "&Send Message"
      End
      Begin VB.Menu mnuContextSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAdd 
         Caption         =   "&Add to list"
      End
   End
End
Attribute VB_Name = "frmOnlineUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function NetSessionEnum Lib "netapi32.dll" (ServerName As Byte, UncClientName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PreMaxLen As Long, EntriesRead As Long, TotalEntries As Long, Resume_Handle As Long) As Long

Private Type Session_Info_10
   sesi10_cname                       As Long
   sesi10_username                    As Long
   sesi10_time                        As Long
   sesi10_idle_time                   As Long
End Type

Private strComputers() As String
Private strUsers() As String

Private Sub cmdCollapse_Click()
    Dim a As Long
    
    tvwMain.Visible = False
    For a = 1 To tvwMain.Nodes.Count
        tvwMain.Nodes(a).Expanded = False
    Next a
    tvwMain.Visible = True
End Sub

Private Sub cmdExpand_Click()
    Dim a As Long
    
    tvwMain.Visible = False
    For a = 1 To tvwMain.Nodes.Count
        tvwMain.Nodes(a).Expanded = True
    Next a
    tvwMain.Visible = True
End Sub

Private Sub Form_Load()
    Dim a As Long
    Dim intComputerCount As Integer
    Dim intUserCount As Long
    
    Me.Visible = True
    tvwMain.ImageList = frmMain.tvwMain.ImageList
    SessionEnum frmMain.umUsers.Server, "", ""
    On Error Resume Next
    For a = 1 To UBound(strComputers)
        If strUsers(a) <> "" Then
            tvwMain.Nodes.Add , tvwLast, strUsers(a), getRealName(frmMain.umUsers.Server, strUsers(a)) & " (" & strUsers(a) & ")", "Online User"
            tvwMain.Nodes.Add strUsers(a), tvwChild, , strComputers(a), "Computer"
        End If
        DoEvents
    Next a
    
    'Establish the number of computers and users online
    For a = 1 To tvwMain.Nodes.Count
        If tvwMain.Nodes(a).Image = "Online User" Then
            intUserCount = intUserCount + 1
        ElseIf tvwMain.Nodes(a).Image = "Computer" Then
            intComputerCount = intComputerCount + 1
        End If
    Next a
    Me.Caption = "Browse online users (" & Str(intUserCount) & " users," & Str(intComputerCount) & " computers)"
End Sub

Private Function SessionEnum(sServerName As String, sClientName As String, sUserName As String) As Long
   Dim bFirstTime           As Boolean
   Dim lRtn                 As Long
   Dim lPrefmaxlen          As Long
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
    ReDim strComputers(0 To lEntriesRead - 1)
    ReDim strUsers(0 To lEntriesRead - 1)
     
    CopyMem typSessionInfo(0), ByVal lptrBuffer, Len(typSessionInfo(0)) * lEntriesRead
     
    For I = 0 To lEntriesRead - 1
        strComputers(I) = PointerToStringW(typSessionInfo(I).sesi10_cname)
        strUsers(I) = PointerToStringW(typSessionInfo(I).sesi10_username)
        DoEvents
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

Private Sub mnuContextAdd_Click()
    Load frmAdd
    frmAdd.txtUsername = tvwMain.SelectedItem.Key
    frmAdd.Show vbModal
End Sub

Private Sub mnuContextSend_Click()
    Dim msgSend As New frmMessage
    
    msgSend.newMessage tvwMain.SelectedItem.Text, tvwMain.SelectedItem.Parent.Key
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        If tvwMain.HitTest(x, y).Image = "Computer" Then
            If Err.Number = 0 Then
                'Computer selected
                mnuContextSend.Enabled = True
                mnuContextAdd.Enabled = False
                frmOnlineUsers.PopupMenu mnuContext
            Else
                Err.Clear
            End If
        ElseIf tvwMain.HitTest(x, y).Image = "Online User" Then
            If Err.Number = 0 Then
                'Computer selected
                mnuContextSend.Enabled = False
                mnuContextAdd.Enabled = True
                frmOnlineUsers.PopupMenu mnuContext
            Else
                Err.Clear
            End If
        End If
    End If
End Sub
