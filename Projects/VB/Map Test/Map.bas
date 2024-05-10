Attribute VB_Name = "Module1"
Function OpenMap(ByVal x As Integer, ByVal y As Integer) As Object
        Dim mapFactory As Object
        Dim map As Object
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
            Set OpenMap = Nothing
        Else
            Set OpenMap = map
        End If
End Function
