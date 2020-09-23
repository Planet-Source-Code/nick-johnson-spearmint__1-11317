VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a single user, their details, and all the computers they are logged into (if any)"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Computers"
Private pstrUserName As String 'local copy
Private pstrName As String 'local copy
Private pstrComment As String 'local copy
Private pstrUserComment As String 'local copy
Private pblnJustAdded As Boolean 'local copy
Private pblnNotify As Boolean
Private pblnIgnored As Boolean
Private pstrGroup As String
Private pcomComputers As Computers


Public Property Get Computers() As Computers
    If pcomComputers Is Nothing Then
        Set pcomComputers = New Computers
    End If


    Set Computers = pcomComputers
End Property

Private Sub Class_Terminate()
    Set pcomComputers = Nothing
End Sub

Friend Property Let Name(ByVal vData As String)
Attribute Name.VB_Description = "The real name of this user"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    pstrName = vData
End Property
Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = pstrName
End Property

Friend Property Let UserName(ByVal vData As String)
Attribute UserName.VB_Description = "The username of this user"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    pstrUserName = vData
End Property
Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    UserName = pstrUserName
End Property

Friend Property Let Comment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    pstrComment = vData
End Property
Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    Comment = pstrComment
End Property

Friend Property Let UserComment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Username = 5
    pstrUserComment = vData
End Property
Public Property Get UserComment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Username
    Comment = pstrUserComment
End Property

Friend Property Let JustAdded(ByVal vData As Boolean)
    pblnJustAdded = vData
End Property
Friend Property Get JustAdded() As Boolean
    JustAdded = pblnJustAdded
End Property

Public Property Let Notify(ByVal vData As Boolean)
    pblnNotify = vData
End Property
Public Property Get Notify() As Boolean
    Notify = pblnNotify
End Property

Public Property Let Ignored(ByVal vData As Boolean)
    pblnIgnored = vData
End Property
Public Property Get Ignored() As Boolean
    Ignored = pblnIgnored
End Property

Public Property Let Group(vData As String)
    pstrGroup = vData
End Property
Public Property Get Group() As String
    Group = pstrGroup
End Property
