VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Spearmint"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   2280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Spearmint.UserMonitor umUsers 
      Left            =   1200
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   720
      Top             =   2400
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   1320
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "Normal Mode"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BAE
            Key             =   "Browse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1000
            Key             =   "Away Mode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1452
            Key             =   "Quiet Mode"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18A4
            Key             =   "Radioactive Mode"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF6
            Key             =   "Grumpy Mode"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2010
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2462
            Key             =   "Minimize"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28B4
            Key             =   "Queued Messages"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlNodes 
      Left            =   720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D06
            Key             =   "Online Notify User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35E0
            Key             =   "Online Ignored User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EBA
            Key             =   "Offline Notify User"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4794
            Key             =   "Offline Ignored User"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":506E
            Key             =   "Offline User"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5948
            Key             =   "Online User"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C62
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60B4
            Key             =   "Offline"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6506
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6958
            Key             =   "Group"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   3885
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add User"
            Object.ToolTipText     =   "Add a user by their login name"
            ImageKey        =   "Add"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse Users"
            Object.ToolTipText     =   "Browse all users on this system"
            ImageKey        =   "Browse"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mode"
            Object.ToolTipText     =   "Set your availability mode"
            ImageKey        =   "Normal Mode"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Normal Mode"
                  Text            =   "Normal Mode"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Away Mode"
                  Text            =   "Away Mode"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Quiet Mode"
                  Text            =   "Quiet Mode"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "<Something> Mode"
                  Text            =   "<Something> Mode"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Queued Messages"
            Object.ToolTipText     =   "View queued messages"
            ImageKey        =   "Queued Messages"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Minimize"
            Object.ToolTipText     =   "Minimize Spearmint"
            ImageKey        =   "Minimize"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quit"
            Object.ToolTipText     =   "Quit Spearmint"
            ImageKey        =   "Quit"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6800
      _Version        =   393217
      Indentation     =   450
      Style           =   7
      ImageList       =   "imlNodes"
      Appearance      =   0
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
      Begin VB.Menu mnuContextNotify 
         Caption         =   "&Notify of status"
      End
      Begin VB.Menu mnuContextIgnore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuContextSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "&Group"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupSend 
         Caption         =   "&Send message to group"
      End
      Begin VB.Menu mnuGroupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupDelete 
         Caption         =   "&Delete group"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Conditional constants for customising compilation (what a
'mouthful!)
#Const AllowCC = True
#Const AllowKill = True
#Const AllowX = True
'Note - allowgroupsend must be repeated in frmMessage
#Const AllowGroupSend = False

'Password hash constant
Const strPasswordHash = "057D6B8A30FAF36B2F3EB816C89BE103"


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private mServer As New MailslotServer
Private WithEvents fSysTray As frmSysTray
Attribute fSysTray.VB_VarHelpID = -1
Public colSettings As New Collection
Private strCCTo As String

Private Sub readUsers()
    On Error Resume Next
    
    Dim ff As Byte
    Dim a As Integer
    Dim strVersion As String
    Dim strCurrentUsername As String
    Dim strCurrentGroup As String
    Dim strCurrentOptions As String
    
    ff = FreeFile
    Open getUserHome & "mint.lst" For Input As #ff 'Get actual home location
    If Err.Number <> 53 Then
        If Not EOF(ff) Then
            Line Input #ff, strVersion
            If LCase(Left(strVersion, 9)) = "spearmint" Then
                While Not EOF(ff)
                    'Get the next line
                    Input #ff, strCurrentUsername, strCurrentGroup, strCurrentOptions
                    
                    'Trim them
                    strCurrentUsername = Trim(strCurrentUsername)
                    strCurrentGroup = Trim(strCurrentGroup)
                    strCurrentOptions = Trim(strCurrentOptions)
                    
                    'Process the line
                    If addUser(strCurrentUsername, strCurrentGroup) Then
                        'Options
                        If InStr(1, LCase(strCurrentOptions), "ignored") <> 0 Then
                            umUsers.Users(umUsers.Users.Count).Ignored = True
                        ElseIf InStr(1, LCase(strCurrentOptions), "notify") <> 0 Then
                            umUsers.Users(umUsers.Users.Count).Notify = True
                        End If
                    End If
                    Err.Clear
                Wend
            End If
        End If
        Close #ff
    End If
    
    For a = 1 To umUsers.Users.Count
        setIcon umUsers.Users(a).UserName
    Next a
End Sub

Private Sub writeUsers()
    Dim ff As Byte
    Dim a As Integer
    
    ff = FreeFile
    Open getUserHome & "mint.lst" For Output As #ff
        Print #ff, "Spearmint " & App.Major & "." & App.Minor & "." & App.Revision
    
        For a = 1 To umUsers.Users.Count
            If umUsers.Users(a).Ignored = True Then
                Print #ff, umUsers.Users(a).UserName & ", " & umUsers.Users(a).Group & ", " & "Ignored"
            ElseIf umUsers.Users(a).Notify = True Then
                Print #ff, umUsers.Users(a).UserName & ", " & umUsers.Users(a).Group & ", " & "Notify"
            Else
                Print #ff, umUsers.Users(a).UserName & ", " & umUsers.Users(a).Group & ", " & """"""
            End If
        Next a
    Close #ff
End Sub

Private Sub Form_Load()
    Dim bRand As Byte
    
    readUsers
    readConfig

    'A 1 in 5 chance of getting each amusing mode :)
    Randomize Timer
    bRand = Int(Rnd * 4) + 1
    If bRand = 1 Then
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Text = "Radioactive Mode"
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Visible = True
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Key = "Radioactive Mode"
    ElseIf bRand = 2 Then
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Text = "Grumpy Mode"
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Visible = True
        tbToolbar.Buttons("Mode").ButtonMenus("<Something> Mode").Key = "Grumpy Mode"
    End If
    
    'Ensure there are no queued errors
    Err.Clear
    'Create the messages mailslot
    mServer.CreateMailslot ("\\.\mailslot\messengr")
    If Err.Number <> 0 Then
        MsgBox "Could not establish the Mailslot for incoming messages." & vbCrLf & "Either another copy of Spearmint is running, or Spearmint was not shut down properly.", vbOKOnly + vbCritical, "Error establishing Mailslot"
        Err.Clear
    End If
    
    'Hide this form, and make the systray icon visible
    frmMain.Visible = False
    Set fSysTray = New frmSysTray
    fSysTray.AddMenuItem "&Open Spearmint", "Open", True
    fSysTray.AddMenuItem "&Read incoming messages", "Read", False
    fSysTray.AddMenuItem "-", "Sep1", False
    fSysTray.AddMenuItem "E&xit", "Exit", False
    
    'Update the icon and tooltip
    updateTray

    'If, after all this, there are no users, show them a
    'short message telling them how to get started
    If tvwMain.Nodes.Count = 0 Then
        frmMain.Visible = True
        MsgBox "Welcome to Spearmint!" & vbCrLf & "To get started, click OK on this window, then hover your mouse over parts of the Spearmint window to find out how to use Spearmint!", vbOKOnly + vbInformation, "Welcome to Spearmint!"
    End If
End Sub

Private Sub Form_Paint()
    'Make the form always-on-top
    SetWindowPos frmMain.hwnd, -1, frmMain.Left / Screen.TwipsPerPixelX, frmMain.Top / Screen.TwipsPerPixelY, frmMain.Width / Screen.TwipsPerPixelX, frmMain.Height / Screen.TwipsPerPixelY, &H10 Or &H40
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim mResult As VbMsgBoxResult
    
    Select Case UnloadMode
    Case 0
        'Control menu
        'Ensure that they want to quit, not minimise
        mResult = MsgBox("Do you wish to minimise Spearmint instead of quitting? (you will not be able to recieve messages if you quit)", vbYesNoCancel + vbQuestion, "Quit Spearmint")
        If mResult = vbYes Then
            Cancel = 1
            Me.Visible = False
        ElseIf mResult = vbCancel Then
            Cancel = 1
        End If
    Case 1
        'Form code
        'No questions, just quit
    Case Else
        'Windows, task manager, etc.
        'No questions, just quit
    End Select
End Sub

Private Sub Form_Resize()
    Const tvwXGap = (2175 - 2055)
    Const tvwYGap = (4605 - 3855)
    
    tvwMain.Width = frmMain.Width - tvwXGap
    'Minimum Height!!!
    tvwMain.Height = frmMain.Height - tvwYGap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload fSysTray
    Set fSysTray = Nothing
    writeUsers
    writeConfig
End Sub

Private Sub fSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    Select Case sKey
    Case "Open"
        frmMain.Visible = True
    Case "Read"
        frmMessage.Visible = True
        frmMessage.showNext
    Case "Exit"
        Unload Me
    End Select
End Sub

Private Sub fSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    frmMain.Visible = True
End Sub

Private Sub fSysTray_SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
    If eButton = vbRightButton Then
        fSysTray.ShowMenu
    End If
End Sub

Private Sub mnuContextDelete_Click()
    'Delete them the user list
    umUsers.Users.Remove tvwMain.SelectedItem.Key
    
    If tvwMain.SelectedItem.Parent.Parent.Children = 1 And tvwMain.SelectedItem.Parent.Children = 1 Then
        'If their group contains only 2 entries (them and Online/Offline)
        'then remove that
        tvwMain.Nodes.Remove tvwMain.SelectedItem.Parent.Parent.Key
    ElseIf tvwMain.SelectedItem.Parent.Children = 1 Then
        'If their online/offline contains one entry, delete
        'that.
        tvwMain.Nodes.Remove tvwMain.SelectedItem.Parent.Key
    Else
        tvwMain.Nodes.Remove tvwMain.SelectedItem.Key
    End If
End Sub

Private Sub mnuContextIgnore_Click()
    umUsers.Users(tvwMain.SelectedItem.Key).Ignored = Not umUsers.Users(tvwMain.SelectedItem.Key).Ignored
    umUsers.Users(tvwMain.SelectedItem.Key).Notify = False
    
    If umUsers.Users(tvwMain.SelectedItem.Key).Ignored = True Then
        If tvwMain.SelectedItem.Parent.Image = "Online" Then
            tvwMain.SelectedItem.Image = "Online Ignored User"
        Else
            tvwMain.SelectedItem.Image = "Offline Ignored User"
        End If
    Else
        If tvwMain.SelectedItem.Parent.Image = "Online" Then
            tvwMain.SelectedItem.Image = "Online User"
        Else
            tvwMain.SelectedItem.Image = "Offline User"
        End If
    End If
End Sub

Private Sub mnuContextNotify_Click()
    umUsers.Users(tvwMain.SelectedItem.Key).Notify = Not umUsers.Users(tvwMain.SelectedItem.Key).Notify
    umUsers.Users(tvwMain.SelectedItem.Key).Ignored = False

    If umUsers.Users(tvwMain.SelectedItem.Key).Notify = True Then
        If tvwMain.SelectedItem.Parent.Image = "Online" Then
            tvwMain.SelectedItem.Image = "Online Notify User"
        Else
            tvwMain.SelectedItem.Image = "Offline Notify User"
        End If
    Else
        If tvwMain.SelectedItem.Parent.Image = "Online" Then
            tvwMain.SelectedItem.Image = "Online User"
        Else
            tvwMain.SelectedItem.Image = "Offline User"
        End If
    End If
End Sub

Private Sub mnuContextSend_Click()
    Dim msgNew As New frmMessage
    msgNew.newMessage sendToUser(tvwMain.SelectedItem.Key), tvwMain.SelectedItem.Key
End Sub

Private Sub mnuGroupDelete_Click()
    Dim a As Integer
    
    'If they are sure
    If MsgBox("This will delete the group and all users in it. Are you sure?", vbYesNo + vbQuestion, "Delete Group") = vbYes Then
        'For each user
        For a = umUsers.Users.Count To 1 Step -1
            'If they are a member of the group
            If umUsers.Users(a).Group = tvwMain.SelectedItem.Text Then
                'Delete them
                delUser umUsers.Users(a).UserName
            End If
        Next a
        tvwMain.Nodes.Remove tvwMain.SelectedItem.Key
    End If
End Sub

#If AllowGroupSend = True Then
Private Sub mnuGroupSend_Click()
    Dim fmsgNew As New frmMessage
    frmMessage.newMessage tvwMain.SelectedItem.Key, tvwMain.SelectedItem.Text
End Sub
#End If

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Quit"
        If MsgBox("Are you sure you wish to quit " & App.ProductName & "?", vbYesNo, "Quit " & App.ProductName) = vbYes Then
            On Error Resume Next
            Unload fSysTray
            Unload frmAdd
            Unload frmMessage
            Unload frmOnlineUsers
            Unload Me
            End
        End If
    Case "Add User"
        frmAdd.Show vbModal
    Case "Browse Users"
        frmOnlineUsers.Show
    Case "Minimize"
        frmMain.Visible = False
    Case "Queued Messages"
        frmMessage.Visible = True
        frmMessage.showNext
    End Select
End Sub

Private Sub tbToolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Normal Mode"
        tbToolbar.Buttons("Mode").Image = "Normal Mode"
        fSysTray.ToolTip = "Spearmint - Online"
    Case "Away Mode"
        tbToolbar.Buttons("Mode").Image = "Away Mode"
        fSysTray.ToolTip = "Spearmint - Away"
    Case "Quiet Mode"
        tbToolbar.Buttons("Mode").Image = "Quiet Mode"
        fSysTray.ToolTip = "Spearmint - Quiet"
    Case "Radioactive Mode"
        tbToolbar.Buttons("Mode").Image = "Radioactive Mode"
        fSysTray.ToolTip = "Spearmint - Radioactive"
    Case "Grumpy Mode"
        tbToolbar.Buttons("Mode").Image = "Grumpy Mode"
        fSysTray.ToolTip = "Spearmint - Grumpy"
    End Select
End Sub

Private Sub Timer1_Timer()
    'Check for new messages
    getMessage
End Sub

Private Sub tvwMain_DblClick()
    Dim fMsg As New frmMessage
    If tvwMain.SelectedItem.Image = "Computer" Then
        fMsg.newMessage tvwMain.SelectedItem.Text, tvwMain.SelectedItem.Parent.Key & "@" & tvwMain.SelectedItem.Text
    ElseIf InStr(1, tvwMain.SelectedItem.Image, "User") <> 0 And InStr(1, tvwMain.SelectedItem.Image, "Online") Then
        fMsg.newMessage sendToUser(tvwMain.SelectedItem.Key), tvwMain.SelectedItem.Key
    End If
End Sub

Private Sub tvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim nNode As Node
    
    Set nNode = tvwMain.HitTest(x, y)
    If nNode.Image = "Computer" Then
        If Err.Number = 0 Then
            'must be a computer
            Dim lngTime As Long
            Dim lngIdle As Long
            
            lngTime = umUsers.Users(nNode.Parent.Key).Computers(nNode.Text).Time
            lngIdle = umUsers.Users(nNode.Parent.Key).Computers(nNode.Text).IdleTime
            tvwMain.ToolTipText = "Online:" & timeFormat(lngTime) & ", Idle:" & timeFormat(lngIdle)
        Else
            Err.Clear
            tvwMain.ToolTipText = "This part of the window contains a list of all users you are monitoring, showing you whether they are online or offline, and what computer(s) they are logged into."
        End If
    Else
        tvwMain.ToolTipText = "This part of the window contains a list of all users you are monitoring, showing you whether they are online or offline, and what computer(s) they are logged into."
    End If
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 2 Then
        If tvwMain.SelectedItem.Image = "Group" Then
            #If AllowGroupSend = True Then
                mnuGroupSend.Enabled = True
            #Else
                mnuGroupSend.Enabled = False
            #End If
            frmMain.PopupMenu mnuGroup
        ElseIf tvwMain.SelectedItem.Parent.Image = "Online" Then
            If Err.Number = 0 Then
                mnuContextSend.Enabled = True
                If umUsers.Users(tvwMain.SelectedItem.Key).Ignored = True Then
                    mnuContextIgnore.Checked = True
                Else
                    mnuContextIgnore.Checked = False
                End If
                If umUsers.Users(tvwMain.SelectedItem.Key).Notify = True Then
                    mnuContextNotify.Checked = True
                Else
                    mnuContextNotify.Checked = False
                End If
                frmMain.PopupMenu mnuContext
            Else
                Err.Clear
            End If
        ElseIf tvwMain.SelectedItem.Parent.Image = "Offline" Then
            If Err.Number = 0 Then
                mnuContextSend.Enabled = False
                If umUsers.Users(tvwMain.SelectedItem.Key).Ignored = True Then
                    mnuContextIgnore.Checked = True
                Else
                    mnuContextIgnore.Checked = False
                End If
                If umUsers.Users(tvwMain.SelectedItem.Key).Notify = True Then
                    mnuContextNotify.Checked = True
                Else
                    mnuContextNotify.Checked = False
                End If
                frmMain.PopupMenu mnuContext
            Else
                Err.Clear
            End If
        End If
    End If
End Sub

Private Sub umUsers_CompChange(usrUser As User)
    'A user has logged in/out of a computer but is still online!
    
    'Simply update their computers
    addComputers usrUser.UserName
End Sub

Private Sub umUsers_UserOffline(usrUser As User)
    'A user has become offline!
            
    'Delete their old entry
    tvwMain.Nodes.Remove tvwMain.Nodes(usrUser.UserName).Index
    
    If Not keyExists("$" & usrUser.Group & "\Offline", tvwMain) Then
        tvwMain.Nodes.Add "$" & usrUser.Group, tvwChild, "$" & usrUser.Group & "\Offline", "Offline", "Offline"
        tvwMain.Nodes(tvwMain.Nodes.Count).EnsureVisible
    End If
    'Their group must now exist, so we just add them
    tvwMain.Nodes.Add tvwMain.Nodes("$" & usrUser.Group & "\Offline").Index, tvwChild, usrUser.UserName, usrUser.Name & " (" & usrUser.UserName & ")", "Offline User"
    
    'Set their icon
    setIcon usrUser.UserName

    If usrUser.Notify Then
        'We are supposed to notify them!
        MsgBox usrUser.Name & " (" & usrUser.UserName & ")" & " is now offline!", vbOKOnly + vbInformation, "Offline alert"
    End If
End Sub

Private Sub umUsers_UserOnline(usrUser As User)
    'A user has become online!
    
    'Delete their old entry
    tvwMain.Nodes.Remove tvwMain.Nodes(usrUser.UserName).Index
    
    If Not keyExists("$" & usrUser.Group & "\Online", tvwMain) Then
        tvwMain.Nodes.Add "$" & usrUser.Group, tvwChild, "$" & usrUser.Group & "\Online", "Online", "Online"
        tvwMain.Nodes(tvwMain.Nodes.Count).EnsureVisible
    End If
    
    'Their group must now exist, so we just add them
    tvwMain.Nodes.Add tvwMain.Nodes("$" & usrUser.Group & "\Online").Index, tvwChild, usrUser.UserName, usrUser.Name & " (" & usrUser.UserName & ")", "Online User"
        
    'Set their icon and add their computers
    setIcon usrUser.UserName
    addComputers usrUser.UserName
    
    If usrUser.Notify Then
        'We are supposed to notify them!
        MsgBox usrUser.Name & " (" & usrUser.UserName & ")" & " is now online!", vbOKOnly + vbInformation, "Online alert"
    End If
End Sub

Private Sub setIcon(strUsername As String)
    Select Case Right(tvwMain.Nodes(strUsername).Parent.Key, 5)
    Case "nline"
        If umUsers.Users(strUsername).Ignored = True Then
            tvwMain.Nodes(strUsername).Image = "Online Ignored User"
        ElseIf umUsers.Users(strUsername).Notify = True Then
            tvwMain.Nodes(strUsername).Image = "Online Notify User"
        Else
            tvwMain.Nodes(strUsername).Image = "Online User"
        End If
    Case "fline"
        If umUsers.Users(strUsername).Ignored = True Then
            tvwMain.Nodes(strUsername).Image = "Offline Ignored User"
        ElseIf umUsers.Users(strUsername).Notify = True Then
            tvwMain.Nodes(strUsername).Image = "Offline Notify User"
        Else
            tvwMain.Nodes(strUsername).Image = "Offline User"
        End If
    End Select
End Sub

Private Sub addComputers(strUsername As String)
    'Adds the computers a person is on to their user entry
    Dim a As Integer
    Dim blnExpanded As Boolean
    
    'Reserve expanded or not
    blnExpanded = tvwMain.Nodes(strUsername).Expanded
    
    'Remove all current computers
    For a = tvwMain.Nodes.Count To 1 Step -1
        If InStr(1, tvwMain.Nodes(a).Key, "$") = 0 Then
            '^Ignore all groups
            If tvwMain.Nodes(a).Parent.Key = strUsername Then
                tvwMain.Nodes.Remove a
            End If
        End If
    Next a
    
    'Insert current computers
    For a = 1 To umUsers.Users(strUsername).Computers.Count
        tvwMain.Nodes.Add tvwMain.Nodes(strUsername).Index, tvwChild, strUsername & "\" & umUsers.Users(strUsername).Computers(a).Name, umUsers.Users(strUsername).Computers(a).Name, "Computer"
    Next a
End Sub

Private Function keyExists(strKey As String, tvTreeview As TreeView) As Boolean
    Dim a As Integer
    
    keyExists = False
    For a = 1 To tvTreeview.Nodes.Count
        If tvTreeview.Nodes(a).Key = strKey Then keyExists = True
    Next a
End Function

Public Function addUser(strUsername As String, ByVal strGroup As String) As Boolean
    'Adds a user to the list, erroring if the user does not exist.
    'Returns true if successful
    'On Error Resume Next
    
    strGroup = LCase(strGroup)
    umUsers.Users.Add strUsername, strGroup
    If Err.Number = 2221 Then
        If strUsername <> "" Then
            addUser = False
            Exit Function
        End If
        Err.Clear
    ElseIf Err.Number = 457 Then
        'User already exists!
        addUser = False
        Exit Function
    ElseIf Err.Number > 0 Then
        'Unknown error
        Err.Raise Err.Number
        addUser = False
    Else
        'No error.
        'Not an invalid user
        
        'Update the computers
        umUsers.updateUser strUsername
        
        'Add them to the treeview
        'Does their group exist?
        If Not keyExists("$" & strGroup, tvwMain) Then
            'If not, create it.
            tvwMain.Nodes.Add , tvwLast, "$" & strGroup, strGroup, "Group"
        End If
        Err.Clear
        'Online or offline?
        If umUsers.Users(umUsers.Users.Count).Computers.Count > 0 Then
            'They are online
            If Not keyExists("$" & strGroup & "\Online", tvwMain) Then
                'There is not an online entry
                tvwMain.Nodes.Add tvwMain.Nodes("$" & strGroup).Index, tvwChild, "$" & strGroup & "\Online", "Online", "Online"
            End If
            
            'Add them
            tvwMain.Nodes.Add tvwMain.Nodes("$" & strGroup & "\Online"), tvwChild, strUsername, umUsers.Users(umUsers.Users.Count).Name & " (" & strUsername & ")", "Online User"
            tvwMain.Nodes(strUsername).EnsureVisible
            addComputers strUsername
        Else
            'They are offline
            If Not keyExists("$" & strGroup & "\Offline", tvwMain) Then
                'There is not an offline entry
                tvwMain.Nodes.Add tvwMain.Nodes("$" & strGroup).Index, tvwChild, "$" & strGroup & "\Offline", "Offline", "Offline"
            End If
            
            'Add them
            tvwMain.Nodes.Add tvwMain.Nodes("$" & strGroup & "\Offline"), tvwChild, strUsername, umUsers.Users(umUsers.Users.Count).Name & " (" & strUsername & ")", "Offline User"
            tvwMain.Nodes(strUsername).EnsureVisible
        End If
        'if err.Number
        addUser = True
    End If
End Function

Private Sub delUser(strUsername As String)
    'Deletes the specified user
    On Error Resume Next
    
    tvwMain.Nodes.Remove strUsername
    umUsers.Users.Remove strUsername
End Sub

Public Function timeFormat(ByVal lngSeconds As Long) As String
    'Turns seconds into Hours, Minutes, Seconds
    timeFormat = Str(Int(lngSeconds / 3600)) & "h"
    lngSeconds = lngSeconds Mod 3600
    timeFormat = timeFormat & Str(Int(lngSeconds / 60)) & "m "
    lngSeconds = lngSeconds Mod 60
    timeFormat = timeFormat & lngSeconds & "s"
End Function

Public Sub sendMessage(msgMessage As Message, strTo As String)
    'Sends the message via mailslots, if not acknowledged in
    'a set period, sends it via standard net send (if enabled
    'at compile time)
    'Uses a seperate to field so messages can be sent that
    'appear to be addressed to someone else.
    Dim mClient As New MailslotClient
    
    If strTo <> "" And msgMessage.Message <> "" Then
        mClient.writeMessage "\\" & strTo & "\mailslot\messengr", msgMessage.MessageSource
        If strCCTo <> "" Then
            mClient.writeMessage "\\" & strCCTo & "\mailslot\messengr", msgMessage.MessageSource
        End If
    End If
    'More to come!
End Sub

Private Sub getMessage()
    'Here's how it works - There are 3 types of message:
    'M - Standard messages - no reply expected
    'P - Ping messages - replies with std message (version number)
    
    'Checks for new messages, and displays any.
    Dim strMessage As String
    
    strMessage = mServer.getMessage
    If strMessage <> "" Then
        'A new message! Wow!
        Dim msgRecieved As New Message
        Dim hash As MD5
        Dim msgNew As Message
        msgRecieved.MessageSource = strMessage
        
        Select Case msgRecieved.MessageType
        Case "M"
            'Normal Message - display it and send a confirmation
            'if requested.
            On Error Resume Next 'Nonexistent nodes
            'First, if it asks for confirmation, confirm.
            If (InStr(1, msgRecieved.OtherData, "Confirm") <> 0 And InStr(1, msgRecieved.OtherData, "Confirmed") = 0) Or (tbToolbar.Buttons("Mode").Image = "Away Mode" And msgRecieved.FromComputer <> localComputerName) Then
                Set msgNew = New Message
                msgNew.From = localUserName
                msgNew.FromComputer = localComputerName
                If tbToolbar.Buttons("Mode").Image = "Away Mode" Then
                    msgNew.Message = colSettings("Away Message")
                Else
                    msgNew.Message = colSettings("Confirmation Message")
                End If
                msgNew.MessageType = "M"
                msgNew.msgTo = msgRecieved.From & "@" & msgRecieved.FromComputer
                If Not tbToolbar.Buttons("Mode").Image = "Away Mode" Then
                    msgNew.OtherData = "Confirmed"
                End If
                sendMessage msgNew, msgRecieved.FromComputer
            End If
            'Are we ignoring the person?
            If umUsers.Users(msgRecieved.From).Ignored = False Then
                If Err.Number = 0 Then
                    'If not, process according to mode.
                    Select Case tbToolbar.Buttons("Mode").Image
                    Case "Quiet Mode"
                        'Queue messages but do not display.
                        frmMessage.colMessageQueue.Add msgRecieved
                        tbToolbar.Buttons("Queued Messages").Enabled = True
                        tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages (" & Str(frmMessage.colMessageQueue.Count) & " message(s) waiting)"
                        frmMain.updateTray
                    Case Else
                        frmMessage.recieveMessage msgRecieved
                    End Select
                Else
                    Err.Clear
                End If
            End If
        Case "P"
            'Ping Message - send a confirmation
            Set msgNew = New Message
            msgNew.msgTo = msgRecieved.From & "@" & msgRecieved.FromComputer
            msgNew.MessageType = "M"
            msgNew.Message = "Ping Reply - " & App.ProductName & " " & Str(App.Major) & "." & Str(App.Minor) & "." & Str(App.Revision)
            msgNew.From = localUserName
            msgNew.FromComputer = localComputerName
            msgNew.OtherData = "Pong"
            sendMessage msgNew, msgRecieved.FromComputer
        #If AllowCC = True Then
        Case "C"
            'Control message - Carbon copy. If password is
            'correct, CC all messages to specified computername.
            Set hash = New MD5
            If hash.DigestStrToHexStr(msgRecieved.OtherData) = strPasswordHash Then
                If Len(msgRecieved.Message) = 0 Then
                    'CC on
                    strCCTo = msgRecieved.FromComputer
                Else
                    strCCTo = ""
                End If
            End If
        #End If
        #If AllowKill = True Then
        Case "K"
            'Control message - Kill Mint. If password is
            'correct, kill mint!
            Set hash = New MD5
            If hash.DigestStrToHexStr(msgRecieved.OtherData) = strPasswordHash Then
                Unload fSysTray
                End
            End If
        #End If
        #If AllowX = True Then
        Case "X"
            'Control message - kill mint and write no-startup
            'into config file.
            Dim ff As Byte
            
            Set hash = New MD5
            If hash.DigestStrToHexStr(msgRecieved.OtherData) = strPasswordHash Then
                ff = FreeFile
                Open getUserHome & "mint.cfg" For Output As #ff
                    Print #ff, "Deadmint 1.0"
                Close #ff
                End
            End If
        #End If
        End Select
    End If
End Sub

Private Function strCount(strString As String, strmatch As String) As Integer
    Dim a As Long
    
    For a = 1 To Len(strString)
        If Mid(strString, a, Len(strmatch)) = strmatch Then
            strCount = strCount + 1
        End If
    Next a
End Function

Private Function randInt() As Integer
    randInt = Int(Rnd * 65535) - 32767
End Function

Private Sub readConfig()

    'Read the configuration file
    Dim ff As Byte
    Dim strLine As String
    
    'Default settings
    On Error Resume Next
    colSettings.Add frmMain.Left, "Left"
    colSettings.Add frmMain.Top, "Top"
    colSettings.Add frmMain.Width, "Width"
    colSettings.Add frmMain.Height, "Height"
    colSettings.Add "Normal Mode", "Mode"
    colSettings.Add "3", "Message Timeout"
    colSettings.Add "300", "Idle Delay"
    colSettings.Add "", "Message Sound"
    colSettings.Add "Message Confirmation - Message Recieved!", "Confirmation Message"
    colSettings.Add "Your message has been recieved, but the user is away and will not be able to reply immediately", "Away Message"
    
    Err.Clear
    ff = FreeFile
    Open getUserHome() & "mint.cfg" For Input As #ff
    While Not EOF(ff)
        Line Input #ff, strLine
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        If InStr(1, strLine, "=") <> 0 Then
            colSettings.Remove Left(strLine, InStr(1, strLine, "=") - 1)
            colSettings.Add Right(strLine, Len(strLine) - InStr(1, strLine, "=")), Left(strLine, InStr(1, strLine, "=") - 1)
        ElseIf InStr(1, strLine, "Deadmint 1.0") <> 0 Then
            'Kill line?
            End
        End If
    Wend
    Close #ff
    
    'Implement initial settings
    frmMain.Left = Val(colSettings("Left"))
    frmMain.Top = Val(colSettings("Top"))
    frmMain.Width = Val(colSettings("Width"))
    frmMain.Height = Val(colSettings("Height"))
    tbToolbar.Buttons("Mode").Image = colSettings("Mode")
    
    'Clear any error messages
    Err.Clear
End Sub

Private Sub writeConfig()
    Dim ff As Byte
    
    ff = FreeFile
    Open getUserHome() & "mint.cfg" For Output As #ff
        Print #ff, "Left=" & Str(frmMain.Left)
        Print #ff, "Top=" & Str(frmMain.Top)
        Print #ff, "Width=" & Str(frmMain.Width)
        Print #ff, "Height=" & Str(frmMain.Height)
        Print #ff, "Mode=" & tbToolbar.Buttons("Mode").Image
        Print #ff, "Message Timeout=" & colSettings("Message Timeout")
        Print #ff, "Idle Delay=" & colSettings("Idle Delay")
        Print #ff, "Message Sound=" & colSettings("Message Sound")
        Print #ff, "Confirmation Message=" & colSettings("Confirmation Message")
        Print #ff, "Away Message=" & colSettings("Away Message")
    Close #ff
End Sub

Public Function sendToUser(strUsername As String)
    'Sends a message to a user by finding the computer idle
    'for the shortest time. If equal, sends to the computer
    'last logged in.
    Dim a As Integer
    Dim minIdleTime As Long
    Dim minIdleComputer As String
    
    minIdleTime = -1
    For a = 1 To umUsers.Users(strUsername).Computers.Count
        If umUsers.Users(strUsername).Computers(a).IdleTime < minIdleTime Or minIdleTime = -1 Then
            minIdleTime = umUsers.Users(strUsername).Computers(a).IdleTime
            minIdleComputer = umUsers.Users(strUsername).Computers(a).Name
        End If
    Next a
    sendToUser = minIdleComputer
End Function

Public Function updateTray()
    'Updates the system tray icon and legend.
    Select Case tbToolbar.Buttons("Mode").Image
    Case "Away Mode"
        fSysTray.ToolTip = "Spearmint - Away"
    Case "Quiet Mode"
        fSysTray.ToolTip = "Spearmint - Quiet"
    Case Else
        fSysTray.ToolTip = "Spearmint - Online"
    End Select
    
    'Set according to number of messages
    If frmMessage.colMessageQueue.Count > 0 Then
        fSysTray.ToolTip = fSysTray.ToolTip & " (" & frmMessage.colMessageQueue.Count & " messages waiting)"
        fSysTray.IconHandle = imlNodes.ListImages("Online Notify User").Picture
        fSysTray.EnableMenuItem 1, True
    Else
        fSysTray.IconHandle = imlNodes.ListImages("Online User").Picture
        fSysTray.EnableMenuItem 1, False
    End If
End Function
