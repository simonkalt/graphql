VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Hotel"
Attribute VB_Ext_KEY = "Member1" ,"Hotels"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarHotelID As Long 'local copy
Private mvarLocationID As Long 'local copy
Private mvarHotels As Hotels
Public Sub Mapify()
    assignHotelsToLocation
End Sub


Public Property Get Hotels() As Hotels
    If mvarHotels Is Nothing Then
        Set mvarHotels = New Hotels
    End If


    Set Hotels = mvarHotels
End Property


Public Property Set Hotels(vData As Hotels)
    Set mvarHotels = vData
End Property



Public Property Let LocationID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationID = 5
    mvarLocationID = vData
End Property


Public Property Get LocationID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationID
    LocationID = mvarLocationID
End Property


Public Sub assignHotelsToLocation()
    Dim map As Object, ll, locationfrom As Object
    Dim locationto As Object, rFinder As Object
    Dim RouteID, RoutePlan As Object, ext
    Dim MapImageURL, i As Integer
    Dim rsCompany As ADODB.Recordset, cn As ADODB.Connection
    Dim rsLocation As ADODB.Recordset
    Dim strFileName As String
    
    cn.Open "DSN=InnSightReports", "sa", "sequoia"
    rsLocation.Open "SELECT Address1, City, State, Zip FROM tblLocation WHERE LocationID = " & LocationID, cn
    
    Set map = CreateObject("RmImsApi.Map")
    Set map = OpenMap()
    Set ll = CreateObject("RMIMSAPI.LocationList")
    map.setZoomlevel (6)
    
    For Each HotelID In Hotels
        rsCompany.Open "SELECT * FROM tblCompany WHERE CompanyID = " & HotelID, cn
        ll.Clear
        ll.AddLocation map.FindAddress(rsLocation("Address1"), rsLocation("City"), rsLocation("State"), rsLocation("Zip"))
        ll.AddLocation map.FindAddress(rsCompany("Street"), rsCompany("City"), rsCompany("State"), rsCompany("PostalCode"))
        strFileName = HotelID & "_to_" & LocationID
        
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
        
        x = map.saveMapImage("\\inetpub/www.innsightreports.com/Maps/" + strFileName + ".jpg")
        rsCompany.Close
    Next
    Set rsCompany = Nothing
    assignLocationToHotels = True
End Sub






