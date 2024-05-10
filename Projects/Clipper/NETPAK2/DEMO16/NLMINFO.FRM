VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   5220
   ClientLeft      =   1530
   ClientTop       =   1920
   ClientWidth     =   8310
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
   Height          =   5595
   Left            =   1485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8310
   Top             =   1590
   Width           =   8400
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show NLM"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   7935
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VBX.NLMINFO NLMINFO1 
      AllocAvailableBytes=   0
      AllocFreeCount  =   0
      CompletionCode  =   0
      Copyright       =   ""
      Count           =   29
      Day             =   0
      Description     =   ""
      EndingIdentificationNumber=   30
      Flags           =   0
      IdentificationNumber=   2
      LastGarbageCollection=   0
      LoadedCount     =   33
      MajorVersion    =   0
      MessageLanguage =   0
      MinorVersion    =   0
      Month           =   0
      NLMName         =   ""
      NumOfReferencedPublics=   0
      ParentID        =   0
      Revision        =   0
      StartingIdentificationNumber=   2
      Type            =   ""
      Year            =   0
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NLM Date:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NLM Description:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NLM Type:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NLM Filename:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NLM Identification Number:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"NLMINFO.frx":0000
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
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   7920
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
    
    NLMINFO1.IdentificationNumber = Val(Text1.TEXT)
    
    NLMINFO1.Refresh
    
    Text2.TEXT = LTrim$(RTrim$(NLMINFO1.NLMName))
    Text3.TEXT = LTrim$(RTrim$(NLMINFO1.Type))
    Text4.TEXT = LTrim$(RTrim$(NLMINFO1.Description))
    Text5.TEXT = LTrim$(RTrim$(NLMINFO1.Month))
    Text6.TEXT = LTrim$(RTrim$(NLMINFO1.Day))
    Text7.TEXT = LTrim$(RTrim$(NLMINFO1.Year))

End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
    Text1.TEXT = NLMINFO1.StartingIdentificationNumber
End Sub

