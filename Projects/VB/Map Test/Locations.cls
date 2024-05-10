VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of Locations"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Location"
Attribute VB_Ext_KEY = "Member0" ,"Location"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarMapHeight As Variant 'local copy
Private mvarMapWidth As Variant 'local copy
Public Property Let MapWidth(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MapWidth = 5
    mvarMapWidth = vData
End Property


Public Property Set MapWidth(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MapWidth = Form1
    Set mvarMapWidth = vData
End Property


Public Property Get MapWidth() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MapWidth
    If IsObject(mvarMapWidth) Then
        Set MapWidth = mvarMapWidth
    Else
        MapWidth = mvarMapWidth
    End If
End Property



Public Property Let MapHeight(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MapHeight = 5
    mvarMapHeight = vData
End Property


Public Property Set MapHeight(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MapHeight = Form1
    Set mvarMapHeight = vData
End Property


Public Property Get MapHeight() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MapHeight
    If IsObject(mvarMapHeight) Then
        Set MapHeight = mvarMapHeight
    Else
        MapHeight = mvarMapHeight
    End If
End Property



Public Property Get Item(vntIndexKey As Variant) As Location
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


Public Function Add(ID As String, Street As Variant, City As Variant, State As Variant, Zip As Variant, Optional sKey As String) As Location
    'create a new object
    Dim objNewMember As Location
    Set objNewMember = New Location


    'set the properties passed into the method
    objNewMember.ID = ID
    If IsObject(Street) Then
        Set objNewMember.Street = Street
    Else
        objNewMember.Street = Street
    End If
    If IsObject(City) Then
        Set objNewMember.City = City
    Else
        objNewMember.City = City
    End If
    If IsObject(State) Then
        Set objNewMember.State = State
    Else
        objNewMember.State = State
    End If
    If IsObject(Zip) Then
        Set objNewMember.Zip = Zip
    Else
        objNewMember.Zip = Zip
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Function getMapUrl() As String
    Dim map As Object, ll, locationfrom As Object
    Dim locationto As Object, rFinder As Object
    Dim RouteID, RoutePlan As Object, ext
    Dim MapImageURL
    
    Set map = CreateObject("RmImsApi.Map")
    Set map = OpenMap(mvarMapWidth, mvarMapHeight)
    
    Set ll = CreateObject("RMIMSAPI.LocationList")
    map.setZoomlevel (6)
    
    For Each Location In Me
        ll.addLocation map.findAddress(Location.Street, Location.City, Location.State, Location.Zip)
    Next
    
    Set rFinder = map.getRouteFinder()
    
    rFinder.setStops ll
    rFinder.setHighwayPreference 60
    rFinder.setWeight 1
    rFinder.setOptimize True

    RouteID = rFinder.findRoute()
    'map.clearRoute
    
    Set RoutePlan = map.getRoutePlan(RouteID)
    Set ext = RoutePlan.getRouteAreaExtent()
    map.setExtent ext
    
    MapImageURL = map.getMapImageURL()

    getMapUrl = MapImageURL
End Function

