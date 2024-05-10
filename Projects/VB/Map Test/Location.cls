VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public ID As String

'local variable(s) to hold property value(s)
Private mvarStreet As Variant 'local copy
Private mvarCity As Variant 'local copy
Private mvarState As Variant 'local copy
Private mvarZip As Variant 'local copy
Public Property Let Zip(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Zip = 5
    mvarZip = vData
End Property


Public Property Set Zip(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Zip = Form1
    Set mvarZip = vData
End Property


Public Property Get Zip() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Zip
    If IsObject(mvarZip) Then
        Set Zip = mvarZip
    Else
        Zip = mvarZip
    End If
End Property



Public Property Let State(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Set State(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.State = Form1
    Set mvarState = vData
End Property


Public Property Get State() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    If IsObject(mvarState) Then
        Set State = mvarState
    Else
        State = mvarState
    End If
End Property



Public Property Let City(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.City = 5
    mvarCity = vData
End Property


Public Property Set City(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.City = Form1
    Set mvarCity = vData
End Property


Public Property Get City() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.City
    If IsObject(mvarCity) Then
        Set City = mvarCity
    Else
        City = mvarCity
    End If
End Property



Public Property Let Street(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Street = 5
    mvarStreet = vData
End Property


Public Property Set Street(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Street = Form1
    Set mvarStreet = vData
End Property


Public Property Get Street() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Street
    If IsObject(mvarStreet) Then
        Set Street = mvarStreet
    Else
        Street = mvarStreet
    End If
End Property




