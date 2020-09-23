VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

Private mvarName As String 'local copy
Private mvarTime As Long 'local copy
Private mvarIdleTime As Long 'local copy
Friend Property Let IdleTime(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdleTime = 5
    mvarIdleTime = vData
End Property


Property Get IdleTime() As Long
Attribute IdleTime.VB_Description = "The length of time this connection has been idle"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IdleTime
    IdleTime = mvarIdleTime
End Property



Friend Property Let Time(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Time = 5
    mvarTime = vData
End Property


Public Property Get Time() As Long
Attribute Time.VB_Description = "The length of time this user has been logged in"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Time
    Time = mvarTime
End Property



Friend Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
Attribute Name.VB_Description = "The name of this computer"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property
