VERSION 5.00
Object = "{89B71FA8-59F5-4844-8E2E-FBCAAF44893E}#13.0#0"; "ClockOCX.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ClockOCX.AnalogClock AnalogClock1 
      Height          =   3135
      Left            =   735
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnalogClock1_Click()
   AnalogClock1.About
End Sub
