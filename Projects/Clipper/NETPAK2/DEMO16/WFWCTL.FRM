VERSION 4.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetPak Professional Custom Control Demonstration"
   ClientHeight    =   5625
   ClientLeft      =   2145
   ClientTop       =   1725
   ClientWidth     =   6330
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
   Height          =   6000
   Left            =   2100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6330
   Top             =   1395
   Width           =   6420
   Begin VBX.NPwfw NPwfw1 
      ChangeCachePassword=   -1  'True
      ChangePassword  =   -1  'True
      Left            =   720
      Logoff          =   -1  'True
      Prop10          =   -1  'True
      Prop11          =   -1  'True
      Prop6           =   -1  'True
      Prop7           =   -1  'True
      Prop8           =   -1  'True
      Prop9           =   -1  'True
      NetDeviceMode   =   -1  'True
      Prop12          =   -1  'True
      Prop13          =   -1  'True
      NetServerBrowseDialog=   -1  'True
      NetShareAsDialog=   -1  'True
      NetStopShareDialog=   -1  'True
      Top             =   3180
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   2160
      TabIndex        =   0
      Top             =   4980
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Left            =   720
      Top             =   3180
      Width           =   435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      X1              =   420
      X2              =   5940
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Windows for Workgroups"
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
      Left            =   360
      TabIndex        =   5
      Top             =   330
      Width           =   4455
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1260
      Top             =   3150
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5460
      Top             =   420
      Width           =   480
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WFWCTL.frx":0000
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   540
      TabIndex        =   4
      Top             =   1170
      Width           =   5115
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WFWCTL.frx":0051
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
      Height          =   945
      Left            =   660
      TabIndex        =   3
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WFWCTL.frx":0105
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
      Height          =   465
      Left            =   1740
      TabIndex        =   2
      Top             =   3180
      Width           =   3615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WFWCTL.frx":0153
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
      Height          =   1005
      Left            =   540
      TabIndex        =   1
      Top             =   2010
      Width           =   5115
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

   Form_Unload False

End Sub

Private Sub Form_Load()
    
    Main.Left = Screen.Width \ 2 - Main.Width \ 2
    Main.TOP = Screen.Height \ 2 - Main.Height \ 2
    
    Show Form
    'npwfw1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

