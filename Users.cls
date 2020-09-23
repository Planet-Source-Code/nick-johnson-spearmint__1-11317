VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contains a User object for each user being monitored"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"User"
Attribute VB_Ext_KEY = "Member0" ,"User"
'local variable to hold collection
Private mCol As Collection

Private pstrServer As String

'API defs
Private Declare Function NetUserGetInfo Lib "netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long) As Long

'User types
Private Type USER_INFO_10_API
  Name As Long
  Comment As Long
  UsrComment As Long
  fullName As Long
End Type

Private Type user_info
    Comment As String
    UserComment As String
    fullName As String
End Type

Friend Property Let Server(ByVal vData As String)
    pstrServer = vData
End Property

Public Function Add(UserName As String, Group As String) As User
    On Error Resume Next 'Duplicate keys will be ignored
    Dim uiInfo As user_info
    
    'create a new object
    Dim objNewMember As User
    Set objNewMember = New User

    'set the properties passed into the method
    objNewMember.UserName = UserName

    'Get their details
    lngReturn = getUserInfo(pstrServer, UserName, uiInfo)
    If lngReturn <> 0 Then
        'Usually a user does not exist error
        On Error Resume Next
        Err.Raise lngReturn
        Exit Function
    End If
    objNewMember.Name = uiInfo.fullName
    objNewMember.Comment = uiInfo.Comment
    objNewMember.UserComment = uiInfo.UserComment
    objNewMember.Group = Group
    objNewMember.JustAdded = True
    
    mCol.Add objNewMember, UserName
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As User
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Private Function getUserInfo(strServer As String, strUsername As String, uinfo As user_info) As Long
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
      getUserInfo = lngReturn
      Exit Function
    End If

    'Turn the pointer into a variable
    CopyMem userInfo, ByVal lngptrUserInfo, Len(userInfo)
    
    uinfo.fullName = PointerToStringW(userInfo.fullName)
    uinfo.Comment = PointerToStringW(userInfo.Comment)
    uinfo.UserComment = PointerToStringW(userInfo.UsrComment)
    NetAPIBufferFree lngptrUserInfo
End Function
