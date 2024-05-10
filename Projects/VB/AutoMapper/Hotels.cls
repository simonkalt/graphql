VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Hotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Hotel"
Attribute VB_Ext_KEY = "Member0" ,"Hotel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarLocationID As Variant 'local copy
Public Property Let LocationID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationID = 5
    mvarLocationID = vData
End Property


Public Property Set LocationID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LocationID = Form1
    Set mvarLocationID = vData
End Property


Public Property Get LocationID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationID
    If IsObject(mvarLocationID) Then
        Set LocationID = mvarLocationID
    Else
        LocationID = mvarLocationID
    End If
End Property



Public Function generateMaps() As Boolean
    Dim Map As Object, ll, lZip As String
    Dim LocationFrom As Object, LocationTo As Object, rFinder As Object
    Dim RouteID, RoutePlan As Object, ext
    Dim MapImageURL, i As Integer
    Dim rsCompany As ADODB.Recordset, cn As ADODB.Connection
    Dim rsLocation As ADODB.Recordset
    Dim strFileName As String
    
    Set cn = New ADODB.Connection
    Set rsCompany = New ADODB.Recordset
    Set rsLocation = New ADODB.Recordset
    cn.Open "DSN=InnSightReports", "sa", "sequoia"
    rsLocation.Open "SELECT Street, City, State, Zip FROM tblLocation WHERE LocationID = " & mvarLocationID, cn
    
    Set Map = CreateObject("RmImsApi.Map")
    Set Map = OpenMap()
    Set ll = CreateObject("RMIMSAPI.LocationList")
    Map.setZoomlevel (6)
    
    For Each Hotel In mCol
        rsCompany.Open "SELECT * FROM tblCompany WHERE CompanyID = " & Hotel.HotelID, cn
        
        If IsNull(rsLocation("Zip")) Then
            lZip = ""
        Else
            lZip = rsLocation("Zip")
        End If
        
        Set LocationFrom = Map.findAddress(rsLocation("Street"), rsLocation("City"), rsLocation("State"), lZip)
        Set LocationTo = Map.findAddress(rsCompany("Address1"), rsCompany("City"), rsCompany("State"), rsCompany("PostalCode"))
        ll.Clear
        ll.addLocation LocationFrom
        ll.addLocation LocationTo
        
        strFileName = Hotel.HotelID & "_to_" & mvarLocationID
        
        Set rFinder = Map.getRouteFinder()
        rFinder.setStops ll
        rFinder.setHighwayPreference 60
        rFinder.setWeight 1
        rFinder.setOptimize True
    
        RouteID = rFinder.findRoute()
        
        Map.ClearRoute
        Set RoutePlan = Map.getRoutePlan(RouteID)
        Set ext = RoutePlan.getRouteAreaExtent()
        Map.setExtent ext
        
        x = Map.saveMapImage("c:\inetpub\www.innsightreports.com\Maps\" + strFileName + ".jpg")
        'x = Map.saveMapImage("c:\temp\" + strFileName + ".jpg")
        rsCompany.Close
    Next
    Set rsCompany = Nothing
    generateMaps = True
End Function


Public Function Add(HotelID As Variant, Optional sKey As String) As Hotel
    'create a new object
    Dim objNewMember As Hotel
    Set objNewMember = New Hotel


    'set the properties passed into the method
    If IsObject(HotelID) Then
        Set objNewMember.HotelID = HotelID
    Else
        objNewMember.HotelID = HotelID
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

Public Property Get Item(vntIndexKey As Variant) As Hotel
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

