VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recieved Message"
   ClientHeight    =   3345
   ClientLeft      =   3900
   ClientTop       =   3030
   ClientWidth     =   5415
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkConfirm 
      Caption         =   "Request Confirmation of Reciept"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   4680
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0442
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0894
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0CE6
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":1138
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":158A
            Key             =   "Reply"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   2745
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Send"
            Object.ToolTipText     =   "Send this message"
            ImageKey        =   "Send"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reply"
            Object.ToolTipText     =   "Reply to this message"
            ImageKey        =   "Reply"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Close this message"
            ImageKey        =   "Cancel"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Read next incoming message"
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete all incoming messages"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtMessage 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   -50
      Width           =   5415
      Begin VB.Label lblTo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "To:"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFromUsername 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFrom 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note - allowgroupsend must be repeated in frmMain
#Const AllowGroupSend = False

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public colMessageQueue As New Collection
Private blnNewMessage As Boolean
Private strReallyTo As String
Private strFromComputer As String

Public Sub newMessage(strTo As String, strShowTo As String)
    'Update the buttons
    tbMain.Buttons("Send").Visible = True
    tbMain.Buttons("Reply").Visible = False
    tbMain.Buttons("Cancel").ToolTipText = "Cancel this message"
    tbMain.Buttons("Next").Visible = False
    tbMain.Buttons("Delete").Visible = False
    
    'Yes, it is a new message
    blnNewMessage = True
    
    'Set the to & from
    lblTo.Caption = strShowTo
    strReallyTo = strTo
    lblFromUsername.Caption = localUserName
    lblFrom.Caption = getRealName(frmMain.umUsers.Server, localUserName)
    chkConfirm.Enabled = True
    
    'Make myself visible
    Me.Caption = "New Message"
    Me.Visible = True
End Sub

Public Sub recieveMessage(msgNew As Message)
    'Add it to the queue
    colMessageQueue.Add msgNew
    Beep
    
    'Keep track of the computer it's from
    strFromComputer = msgNew.FromComputer
    
    'Update the icons
    tbMain.Buttons("Send").Visible = False
    tbMain.Buttons("Reply").Visible = True
    tbMain.Buttons("Cancel").ToolTipText = "Close this message"
    
    'If it is not already displayed
    If frmMessage.Visible = False Then
        'Display the first message
        showNext
    'Else, if there is one displayed, update the buttons
    Else
        tbMain.Buttons("Next").Enabled = True
        tbMain.Buttons("Next").ToolTipText = "Read next message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        Me.Caption = "Recieved Message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        frmMain.tbToolbar.Buttons("Queued Messages").Enabled = True
        frmMain.tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        tbMain.Buttons("Delete").Enabled = True
        frmMain.updateTray
    End If
    
    blnNewMessage = False
    chkConfirm.Enabled = False
End Sub

Public Sub showNext()
    'Shows the next message
    Dim msgRead As Message
    
    'Display the textbox correctly
    txtMessage.Enabled = False
    txtMessage.BackColor = vbWhite
    txtMessage.ForeColor = vbBlack
    
    'Get the oldest message
    Set msgRead = colMessageQueue(1)
    colMessageQueue.Remove 1
    
    'Set the fields
    lblFrom.Caption = msgRead.fromName
    lblFromUsername.Caption = msgRead.From
    lblTo.Caption = msgRead.msgTo
    txtMessage.Text = msgRead.Message
    strFromComputer = msgRead.FromComputer
    
    'Ensure the window is visible
    Me.Visible = True
    
    'Disable the next and delete icons if neccessary, else
    'update them
    If colMessageQueue.Count = 0 Then
        'Update the icons etc - no more messages
        tbMain.Buttons("Next").Enabled = False
        tbMain.Buttons("Next").ToolTipText = "Read the next message"
        Me.Caption = "Recieved Message"
        frmMain.tbToolbar.Buttons("Queued Messages").Enabled = False
        frmMain.tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages"
        tbMain.Buttons("Delete").Enabled = False
        frmMain.updateTray
    Else
        tbMain.Buttons("Next").Enabled = True
        tbMain.Buttons("Next").ToolTipText = "Read next message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        Me.Caption = "Recieved Message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        frmMain.tbToolbar.Buttons("Queued Messages").Enabled = True
        frmMain.tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages (" & Str(colMessageQueue.Count) & " message(s) waiting)"
        tbMain.Buttons("Delete").Enabled = True
        frmMain.updateTray
    End If
End Sub

Private Sub Form_Paint()
    If blnNewMessage = False Then
        'Make the form always-on-top
        SetWindowPos Me.hwnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, &H10 Or &H40
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        If blnNewMessage = False Then
            Unload Me
        Else
            If MsgBox("Are you sure you wish to cancel this message?", vbYesNo + vbQuestion, "Cancel message") = vbYes Then
                Unload Me
            Else
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub lblFrom_Click()
    lblFrom.ToolTipText = lblFrom.Caption
End Sub

Private Sub lblFromUsername_Click()
    lblFromUsername.ToolTipText = lblFromUsername.Caption
End Sub

Private Sub lblTo_Change()
    lblTo.ToolTipText = lblTo.Caption
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim msgNew As Message
    Dim fMsg As frmMessage
    Dim a As Integer
    
    Select Case Button.Key
    Case "Cancel"
        If blnNewMessage = False Then
            Unload Me
        Else
            If MsgBox("Are you sure you wish to cancel this message?", vbYesNo + vbQuestion, "Cancel message") = vbYes Then
                Unload Me
            End If
        End If
    Case "Next"
        showNext
    Case "Delete"
        If MsgBox("Are you sure? This will delete all incoming messages!", vbYesNo + vbExclamation, "Delete all incoming messages") = vbYes Then
            
            For a = colMessageQueue.Count To 1 Step -1
                colMessageQueue.Remove a
            Next a
            If colMessageQueue.Count = 0 Then
                'Update the icons etc - no more messages
                tbMain.Buttons("Next").Enabled = False
                tbMain.Buttons("Next").ToolTipText = "Read the next message"
                Me.Caption = "Recieved Message"
                frmMain.tbToolbar.Buttons("Queued Messages").Enabled = False
                frmMain.tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages"
                tbMain.Buttons("Delete").Enabled = False
            Else
                tbMain.Buttons("Next").Enabled = True
                tbMain.Buttons("Next").ToolTipText = "Read next message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
                Me.Caption = "Recieved Message (" & Str(colMessageQueue.Count) & " message(s) waiting)"
                frmMain.tbToolbar.Buttons("Queued Messages").Enabled = True
                frmMain.tbToolbar.Buttons("Queued Messages").ToolTipText = "View Queued Messages (" & Str(colMessageQueue.Count) & " message(s) waiting)"
                tbMain.Buttons("Delete").Enabled = True
            End If
            Unload Me
        End If
    Case "Send"
        Set msgNew = New Message
        msgNew.msgTo = lblTo.Caption & "@" & strReallyTo
        msgNew.From = localUserName
        msgNew.FromComputer = localComputerName
        msgNew.MessageType = "M"
        msgNew.Message = txtMessage.Text
        If chkConfirm.value Then
            'If they are requesting a confirmation message
            msgNew.OtherData = "Confirm"
        End If
        #If AllowGroupSend = True Then
        If Left(strReallyTo, 1) = "$" Then
            'We are sending to a group
            For a = 1 To frmMain.umUsers.Users.Count
                If "$" & frmMain.umUsers.Users(a).Group = strReallyTo Then
                    'This user is a member of the group
                    frmMain.sendMessage msgNew, frmMain.sendToUser(frmMain.umUsers.Users(a).UserName)
                End If
            Next a
        Else
            'We are sending to a person, send normally
            frmMain.sendMessage msgNew, strReallyTo
        End If
        #Else
        frmMain.sendMessage msgNew, strReallyTo
        #End If
        Unload Me
    Case "Reply"
        Set fMsg = New frmMessage
        fMsg.newMessage strFromComputer, lblFromUsername.Caption
        If colMessageQueue.Count > 0 Then
            showNext
        Else
            Unload Me
        End If
    End Select
End Sub

