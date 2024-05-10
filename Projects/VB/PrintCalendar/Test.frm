VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "ARViewer.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Print Calendar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13361
      SectionData     =   "Test.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.ARViewer1.ReportSource = ActiveReport1
End Sub
