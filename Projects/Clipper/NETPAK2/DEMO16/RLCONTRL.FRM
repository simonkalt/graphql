VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   3645
   ClientLeft      =   1680
   ClientTop       =   2670
   ClientWidth     =   7170
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4020
   Left            =   1635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7170
   Top             =   2340
   Width           =   7260
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Read Login Control Property"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Maximum Concurrent Connections:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter User Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Read The Bindery Object Login Control (16-bit DLL version)"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   6960
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    
    End

End Sub

Private Sub Command2_Click()
    
    Dim RLC As ReadLoginControl
    Dim AT As AllowedTimes

    ccode% = NW_ReadLoginControl(LTrim$(RTrim$(Text1.Text)), RLC, AT)

    Text2.Text = NW_SwapInteger(RLC.max_concurrent_conn)
    Text3.Text = Hex$(ccode%)

End Sub

Private Sub Form_Load()
    
    ccode% = NW_NWDrvrLoaded()
    
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

