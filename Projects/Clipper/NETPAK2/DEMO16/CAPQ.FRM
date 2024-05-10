VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   3330
   ClientLeft      =   2280
   ClientTop       =   2745
   ClientWidth     =   6105
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
   Height          =   3705
   Left            =   2235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6105
   Top             =   2415
   Width           =   6195
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set Capture Queue"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter Queue Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set A Capture Print Queue (16-bit DLL Version)"
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
      TabIndex        =   3
      Top             =   360
      Width           =   6015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   5880
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
    ccode% = NW_EndSPLPTCap(0)
    ccode% = NW_SetCapturePrintQ(0, LTrim$(RTrim$(Text1.TEXT)))
    ccode% = NW_StartSPLPTCap(0)
    Text2.TEXT = Hex$(ccode%)
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

