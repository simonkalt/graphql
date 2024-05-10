VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLong 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtLat 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "LatLong.frx":0000
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "&Assign Latitudes && Longitudes"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Longitude:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Latitude:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset, cn As ADODB.Connection

Function OpenMap()
    Dim mapFactory As Object
    Dim Map As Object
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
    Set Map = CreateObject("RmImsApi.Map")

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
    
    
    'ServerURL = "http://dev.innsightreports.com/scripts/webgate.dll"
    ServerURL = "http://www.goconcierge.net/scripts/webgate.dll"
    GroupID = "ROUTEMAP"
    UserID = "demo"
    MapID = "InnSight Reports Beige With FileSave"
    
    Set Map = mapFactory.OpenMap(ServerURL, GroupID, UserID, MapID)
    
    If IsNull(Map) Then
        MsgBox "Open map failed"
    Else
        Set OpenMap = Map
    End If
End Function

Private Sub cmdAssign_Click()
    Dim i As Integer, rc As Integer, lZip As Variant
    Dim Map As Object
    Dim Location As Object
    Dim LatLongLocation As Object
    Dim Projection As Object
    
    Set Map = CreateObject("RmImsApi.Map")
    Set Map = OpenMap()
    
    i = 0
    
    Do Until rs.EOF
        If IsNull(rs("Zip")) Then
            lZip = ""
        Else
            lZip = rs("Zip")
        End If
        
        Set Location = Map.findAddress(rs("Street"), rs("City"), rs("State"), lZip)
        Set Projection = Map.getProjection()
        Set LatLongLocation = Projection.unProjectLocation(Location)
        
        Me.txtAddress.Text = rs("Street") & ", " & rs("City") & ", " & rs("State") & "  " & rs("Zip")
        If VarType(LatLongLocation) = 8 Then
            rs("Latitude") = LatLongLocation.y
            rs("Longitude") = LatLongLocation.x
            Me.txtLat = rs("Latitude")
            Me.txtLong = rs("Longitude")
            rs.Update
            Location.Clear
            LatLongLocation.Clear
        Else
            Me.txtLat = "*** Can not Determine ***"
            Me.txtLong = "*** Can not Determine ***"
        End If
        DoEvents
        rs.MoveNext
        i = i + 1
    Loop
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "DSN=InnsightReports", "sa", "sequoia"
    rs.Open "SELECT Street, City, State, Zip, Latitude, Longitude FROM tblLocation WHERE Latitude IS NULL OR Longitude IS NULL", cn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub
