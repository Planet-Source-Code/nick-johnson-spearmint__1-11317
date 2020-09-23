VERSION 5.00
Begin VB.Form frmLeaf 
   Caption         =   "Spearmint Leaf"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   Icon            =   "frmLeaf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmLeaf.frx":0442
      Left            =   1320
      List            =   "frmLeaf.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "The type of message to send"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtFromComputer 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "The text to be inserted into the From Computer field of the message"
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go!"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Transmit the message"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "The contents of the additional data field."
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtMessage 
      Height          =   885
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "The text of the message"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "The text to be inserted into the From field of the message"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "The computer, domain, or * to which you wish to send this message to"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "From (Computer):"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblOther 
      Alignment       =   1  'Right Justify
      Caption         =   "Additional Data:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Message:"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "To (Computer):"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Message Type:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbType_Click()
    Select Case Left(cmbType.Text, 1)
    Case "M"
        'Standard message
        txtOther.PasswordChar = ""
        lblOther.Caption = "Additional Data:"
        txtMessage.Enabled = True
        txtOther.Enabled = True
        txtOther.ToolTipText = "The contents of the additional data field."
    Case "P"
        'Ping message
        txtOther.PasswordChar = ""
        lblOther.Caption = "Additional Data:"
        txtMessage.Enabled = False
        txtMessage.Text = ""
        txtOther.Enabled = False
        txtOther.Text = ""
        txtOther.ToolTipText = "The contents of the additional data field."
    Case "C"
        'CC message
        txtOther.PasswordChar = "*"
        lblOther.Caption = "Password:"
        txtMessage.Enabled = True
        txtOther.Enabled = True
        txtOther.ToolTipText = "The password to transmit to the client."
    Case "K"
        'Kill message
        txtOther.PasswordChar = "*"
        lblOther.Caption = "Password:"
        txtMessage.Enabled = False
        txtMessage.Text = ""
        txtOther.Enabled = True
        txtOther.ToolTipText = "The password to transmit to the client."
    Case "X"
        'Kill & do not allow restart message
        txtOther.PasswordChar = "*"
        lblOther.Caption = "Password:"
        txtMessage.Enabled = False
        txtMessage.Text = ""
        txtOther.Enabled = True
        txtOther.ToolTipText = "The password to transmit to the client."
    End Select
End Sub

Private Sub cmdGo_Click()
    Dim msgNew As New Message
    Dim mClient As New MailslotClient
    
    msgNew.msgTo = txtTo.Text
    msgNew.From = txtFrom.Text
    msgNew.FromComputer = txtFromComputer.Text
    msgNew.Message = txtMessage.Text
    msgNew.MessageType = Left(cmbType.Text, 1)
    msgNew.OtherData = txtOther.Text
    
    mClient.writeMessage "\\" & txtTo.Text & "\mailslot\messengr", msgNew.MessageSource
End Sub

Private Sub Form_Load()
    txtFrom.Text = localUserName
    txtFromComputer.Text = localComputerName
    cmbType.ListIndex = 0
End Sub
