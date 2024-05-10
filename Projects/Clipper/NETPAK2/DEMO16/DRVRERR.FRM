VERSION 4.00
Begin VB.Form NetError 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   Caption         =   "NetWare Driver Error"
   ClientHeight    =   6030
   ClientLeft      =   870
   ClientTop       =   1845
   ClientWidth     =   7890
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
   Height          =   6405
   Left            =   825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7890
   Top             =   1515
   Width           =   7980
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "secondnet.drv=netware.drv"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4800
      Width           =   7215
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "network.drv=wfwnet.drv"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   7215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "[boot]"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   7215
   End
   Begin VB.Label Label5 
      Caption         =   $"DRVRERR.frx":0000
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "SYSTEM.INI for Windows 3.1 (running NetWare only)"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   7215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "network.drv=netware.drv"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   7215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "[boot]"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   $"DRVRERR.frx":0048
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   2295
      Left            =   360
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "NetError"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Beep
End Sub

