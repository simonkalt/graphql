VERSION 4.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetPak Professional - Bindery"
   ClientHeight    =   5865
   ClientLeft      =   495
   ClientTop       =   3180
   ClientWidth     =   9405
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   6240
   Left            =   450
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9405
   Top             =   2850
   Width           =   9495
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6705
      TabIndex        =   3
      Top             =   5130
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"BINDERY.frx":0000
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   5985
      TabIndex        =   7
      Top             =   1545
      Width           =   2895
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Names"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   3555
      Width           =   2535
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Types"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   5
      Top             =   1035
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Double-click the object type to see object names."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   5970
      TabIndex        =   4
      Top             =   3945
      Width           =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   240
      X2              =   8880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NetWare Functions"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   3435
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
   
   Form_Unload False

End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
    Main.Left = Screen.Width \ 2 - Main.Width \ 2
    Main.Top = Screen.Height \ 2 - Main.Height \ 2
    
    List1.Clear
    List2.Clear

    List1.AddItem "Groups"
    List1.AddItem "Users"
    List1.AddItem "Print Queues"

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

Private Sub List1_DblClick()
    Dim NWBind As BinderyInfo
    
    List2.Clear
    
    For n = 0 To (List1.ListCount - 1)
        If List1.Selected(n) = True Then
            cOName = RTrim(Left(List1.List(n), 48))
        End If
    Next

    Select Case cOName
        Case "Groups"
            ccode% = NW_ScanBindery(-1, &H2, NWBind)
            List2.AddItem NWBind.object_name
    
            Do
                ccode% = NW_ScanBindery(NWBind.object_ID, &H2, NWBind)
                If (ccode% = 0) Then
                    List2.AddItem NWBind.object_name
                End If
            Loop Until ccode%
        
        
        Case "Users"
            ccode% = NW_ScanBindery(-1, &H1, NWBind)
            List2.AddItem NWBind.object_name
    
            Do
                ccode% = NW_ScanBindery(NWBind.object_ID, &H1, NWBind)
                If (ccode% = 0) Then
                    List2.AddItem NWBind.object_name
                End If
            Loop Until ccode%
        Case "Print Queues"
            ccode% = NW_ScanBindery(-1, &H3, NWBind)
            List2.AddItem NWBind.object_name
    
            Do
                ccode% = NW_ScanBindery(NWBind.object_ID, &H3, NWBind)
                If (ccode% = 0) Then
                    List2.AddItem NWBind.object_name
                End If
            Loop Until ccode%
    End Select
End Sub

