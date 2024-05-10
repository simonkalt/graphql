VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Hotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarHotelID As Variant 'local copy
Public Property Let HotelID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HotelID = 5
    mvarHotelID = vData
End Property


Public Property Set HotelID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.HotelID = Form1
    Set mvarHotelID = vData
End Property


Public Property Get HotelID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HotelID
    If IsObject(mvarHotelID) Then
        Set HotelID = mvarHotelID
    Else
        HotelID = mvarHotelID
    End If
End Property



