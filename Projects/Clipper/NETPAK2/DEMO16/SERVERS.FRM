VERSION 4.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetPak Professional - Servers"
   ClientHeight    =   6885
   ClientLeft      =   270
   ClientTop       =   300
   ClientWidth     =   8850
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
   Height          =   7290
   Left            =   210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8850
   Top             =   -45
   Width           =   8970
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   7020
      TabIndex        =   7
      Top             =   6240
      Width           =   1515
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   5760
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   5370
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   4980
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Change Password"
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      Top             =   4260
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Verify Password"
      Height          =   435
      Left            =   6360
      TabIndex        =   6
      Top             =   4260
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Login To File Server"
      Height          =   435
      Left            =   1800
      TabIndex        =   4
      Top             =   4260
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   1800
      TabIndex        =   0
      Top             =   2100
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      X1              =   300
      X2              =   8520
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label12 
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
      TabIndex        =   19
      Top             =   300
      Width           =   4395
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"SERVERS.frx":0000
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
      Height          =   765
      Left            =   1800
      TabIndex        =   18
      Top             =   960
      Width           =   6795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Function Name:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   4380
      Width           =   1395
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   180
      TabIndex        =   16
      Top             =   3930
      Width           =   1575
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6360
      TabIndex        =   15
      Top             =   3900
      Width           =   2175
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   3900
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   3900
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "New Password:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   5790
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Old Password:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   5430
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Name:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   600
      TabIndex        =   10
      Top             =   5010
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File Server Selected:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1800
      TabIndex        =   9
      Top             =   3390
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "No File Server Selected"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3660
      TabIndex        =   8
      Top             =   3390
      Width           =   4815
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    
    Dim nConnID As Integer
    cUserName$ = Text1.Text
    cOldPwd$ = Text2.Text

    Label6.Caption = ""
    
    ccode% = NW_AttachToFS(cOName, nConnID)
    
    ' Added to fix bug Kim reported
    ccode% = NW_SetPrefConnID(nConnID)
    
    ccode% = NW_LoginToFS(cUserName$, cOldPwd$)

    Label6.Caption = Hex$(ccode%)

    Beep
End Sub

Private Sub Command2_Click()
    
    cUserName$ = Text1.Text
    cOldPwd$ = Text2.Text
    
    Label8.Caption = ""

    ccode% = NW_VerifyBObjPwd(cUserName$, cOldPwd$)
    
    Label8.Caption = Hex$(ccode%)

    Beep
End Sub

Private Sub Command3_Click()
    
    cUserName$ = Text1.Text
    cOldPwd$ = Text2.Text
    cNewPwd$ = Text3.Text
    
    Label7.Caption = ""

    ccode% = NW_ChangeBObjPwd(cUserName$, cOldPwd$, cNewPwd$)

    Label7.Caption = Hex$(ccode%)

    Beep
End Sub

Private Sub Command4_Click()
  Form_Unload False
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
    Dim NWBind As BinderyInfo
    
    Main.Left = Screen.Width \ 2 - Main.Width \ 2
    Main.Top = Screen.Height \ 2 - Main.Height \ 2
    
'   Main.WindowState = 2

    List1.Clear
    
    ccode% = NW_ScanBindery(-1, &H4, NWBind)
    List1.AddItem NWBind.object_name
    
    Do
        ccode% = NW_ScanBindery(NWBind.object_ID, &H4, NWBind)
        If (ccode% = 0) Then
            List1.AddItem NWBind.object_name
        End If
    
    Loop Until ccode%

End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub List1_DblClick()
    
    For n = 0 To (List1.ListCount - 1)
        If List1.Selected(n) = True Then
            cOName = RTrim(Left(List1.List(n), 48))
        End If
    Next

    Label1.ForeColor = Label2.ForeColor
    Label1.Caption = cOName

End Sub

