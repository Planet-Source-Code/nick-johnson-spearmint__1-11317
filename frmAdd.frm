VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add a user"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmbExistingGroup 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtNewGroup 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1050
      Width           =   1695
   End
   Begin VB.OptionButton optAddType 
      Caption         =   "To a new group called:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.OptionButton optAddType 
      Caption         =   "To an existing group:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmAdd
End Sub

Private Sub cmdOk_Click()
    If txtUsername.Text <> "" And ((optAddType(0).value = True And cmbExistingGroup.Text <> "") Or (optAddType(1).value = True And txtNewGroup.Text <> "")) Then
        If optAddType(0).value = True Then
            'Adding to existing group
            If Not frmMain.addUser(txtUsername.Text, cmbExistingGroup.Text) Then
                'Could not add
                MsgBox "" & txtUsername.Text & "" & " is not a valid username, or is already in your user list. Please enter a valid username in the Username box!", vbOKOnly + vbExclamation, "Add a user"
            Else
                'Added successfully
                Unload Me
                Exit Sub
            End If
        Else
            'New group
            If Not frmMain.addUser(txtUsername.Text, txtNewGroup.Text) Then
                'Could not add
                MsgBox "" & txtUsername.Text & "" & " is not a valid username, or is already in your user list. Please enter a valid username in the Username box!", vbOKOnly + vbExclamation, "Add a user"
            Else
                'Added successfully
                Unload Me
                Exit Sub
            End If
        End If
    Else
        MsgBox "Please enter a valid username and groupname!", vbOKOnly + vbExclamation, "Add a user"
    End If
End Sub

Private Sub Form_Load()
    Dim a As Integer
    
    For a = 1 To frmMain.tvwMain.Nodes.Count
        If frmMain.tvwMain.Nodes(a).Image = "Group" Then
            cmbExistingGroup.AddItem frmMain.tvwMain.Nodes(a).Text
        End If
    Next a
    If cmbExistingGroup.ListCount > 0 Then
        cmbExistingGroup.ListIndex = 0
    End If
End Sub
