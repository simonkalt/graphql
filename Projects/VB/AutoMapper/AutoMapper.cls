VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Hotels"
Public Function assignLocationToHotels() As Boolean
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
    
    For Each HotelID In Me.Hotels
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
End Function

Private Function OpenMap()
        Dim mapFactory As Object
        Dim map As Object
        Dim x
        Dim y
        Dim nTop
        Dim nLeft
        Dim nBottom
        Dim nRight
        
        Dim ServerURL
        Dim GroupID
        Dim UserID
        Dim MapID
        
        Set mapFactory = CreateObject("RMImsApi.MapFactory")
        Set map = CreateObject("RmImsApi.Map")

        x = 660
        y = 340
        
        nTop = 0#
        nLeft = 0#
        nBottom = 0#
        nRight = 0#
        
        mapFactory.setImageSize x, y
        mapFactory.setExtentCoord nTop, nLeft, nBottom, nRight
        mapFactory.setLanguage "en"
        mapFactory.setUnits 0
        
        
        ServerURL = "http://www.goconcierge.net/scripts/webgate.dll"
        GroupID = "ROUTEMAP"
        UserID = "demo"
        MapID = "InnSight Reports Beige With FileSave"
        
        Set map = mapFactory.OpenMap(ServerURL, GroupID, UserID, MapID)
        
        If IsNull(map) Then
            MsgBox "Open map failed"
        Else
            Set OpenMap = map
        End If
End Function
