VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Route"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Sub Mapify()
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

