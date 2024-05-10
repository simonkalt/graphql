VERSION 4.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetPak Professional Demonstration - USING"
   ClientHeight    =   5985
   ClientLeft      =   1470
   ClientTop       =   1815
   ClientWidth     =   8340
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   6360
   Left            =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8340
   Top             =   1485
   Width           =   8430
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   6420
      TabIndex        =   2
      Top             =   5370
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   300
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search for Connections Using A File"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   2130
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      X1              =   360
      X2              =   7920
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NetWare Functions 16-bit"
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
      Height          =   435
      Left            =   300
      TabIndex        =   12
      Top             =   330
      Width           =   4395
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File in use by:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Path and Filename:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   300
      TabIndex        =   7
      Top             =   1860
      Width           =   1755
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"USING.frx":0000
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
      Height          =   735
      Left            =   300
      TabIndex        =   10
      Top             =   990
      Width           =   7695
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Example: SYS:PTS\PTS_DATA\CONSULT.DBF)"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3720
      TabIndex        =   8
      Top             =   1860
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Connections Using File:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4320
      TabIndex        =   6
      Top             =   3030
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code (0 = Success):"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4320
      TabIndex        =   5
      Top             =   2670
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7380
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7380
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    
    Dim ScanUsing  As ScanCnUsingFile
    Dim nCount As Integer
    Dim cObjName As String * 48
    Dim cLocalObjName As String * 48
    Dim cPValue As String * 128

    List1.Clear
    
    ccode% = NW_ScanConnUsingAFileF2(UCase$(Text1.TEXT), nCount, ScanUsing)

    Label1.Caption = Hex$(ccode%)
    
    If nCount > 1 Then
        Label2.Caption = nCount
    Else
        Label2.Caption = 0
    End If

    For i = 0 To nCount - 1
        cObjName = Space$(48)
        cLocalObjName = Space$(48)
        ccode% = NW_GetConnInfoF2(ScanUsing.connection_number(i), cObjName)
        ccode% = NW_GetConnInfoF2(NW_GetConnNbr(), cLocalObjName)
        
        If cObjName <> cLocalObjName Then
            ccode% = NW_UserFullName(cObjName, cPValue)
            List1.AddItem cPValue
        Else
            ccode% = NW_UserFullName(cObjName, cPValue)
            List1.AddItem "* " + cPValue
        End If
    Next

    Text1.TEXT = UCase$(Text1.TEXT)

End Sub

Private Sub Command2_Click()
  
  Form_Unload False

End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
  Main.Left = Screen.Width \ 2 - Main.Width \ 2
  Main.TOP = Screen.Height \ 2 - Main.Height \ 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

