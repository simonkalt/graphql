VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "getMapURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Location"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarNewEnum As IUnknown 'local copy
Private mvarCount As Long 'local copy
Private mvarItem As Location 'local copy
Public Property Get Item() As Location
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Item
    Set Item = mvarItem
End Property



Public Property Get Count() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Count
    Count = mvarCount
End Property



Public Sub Remove(vntIndexKey As Variant)
End Sub

Public Property Get NewEnum() As IUnknown
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewEnum
    Set NewEnum = mvarNewEnum
End Property





