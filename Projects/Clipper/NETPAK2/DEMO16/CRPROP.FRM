VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4425
   ClientLeft      =   1995
   ClientTop       =   1605
   ClientWidth     =   7020
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
   Height          =   4800
   Left            =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7020
   Top             =   1275
   Width           =   7110
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Create Property"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Type:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Property Name To Create:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create A Bindery Object Property 16-bit"
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   6840
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
    ccode% = NW_CreateProperty(Val(Text4.TEXT), LTrim$(RTrim$(Text1.TEXT)), BF_ITEM Or BF_STATIC, BS_LOGGED_READ Or BS_SUPER_WRITE, LTrim$(RTrim$(Text2.TEXT)))
    Text3.TEXT = Hex$(ccode%)
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

