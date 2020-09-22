VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ClockOCX.AnalogClock AnalogClock1 
      Height          =   1950
      Left            =   180
      TabIndex        =   3
      Top             =   225
      Width           =   1905
      _ExtentX        =   3307
      _ExtentY        =   3413
   End
   Begin VB.Timer tmrUnload 
      Interval        =   12000
      Left            =   2505
      Top             =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   2340
      X2              =   2340
      Y1              =   45
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2310
      X2              =   2310
      Y1              =   45
      Y2              =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTERED VERSION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   2565
      TabIndex        =   2
      Top             =   555
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   3765
      TabIndex        =   1
      Top             =   1830
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gradient Analog Clock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2565
      TabIndex        =   0
      Top             =   255
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   330
      Left            =   3765
      Shape           =   4  'Rounded Rectangle
      Top             =   1770
      Width           =   1410
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   3810
      Shape           =   4  'Rounded Rectangle
      Top             =   1815
      Width           =   1410
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim myLS As Long, myTS As Long, myTL As Long, myDL As Long

Private Sub Form_Load()
On Error Resume Next

   Label1(2) = "REGISTERED VERSION" & vbCrLf & "Made by. Batavian" & vbCrLf & "Jakarta - Indonesia" & vbCrLf & "batavian_forever@hotmail.com"
   Me.ScaleMode = vbPixels
   SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H2 Or &H1 Or &H40 Or &H200
   
   AnalogClock1.CircleBorder = vbWhite
   AnalogClock1.ClockBody = vbBlack
   AnalogClock1.HourOutline = vbWhite
   AnalogClock1.HourPointer = vbBlack
   AnalogClock1.MinuteOutline = vbWhite
   AnalogClock1.MinutePointer = vbBlack
   AnalogClock1.ShowMajorP = False
   AnalogClock1.ShowMinorP = False
   
   myLS = Shape1.Left
   myTS = Shape1.Top
   myTL = Label1(1).Top

On Error GoTo 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shape1.BackColor = vbRed Then
      Shape1.BackColor = vbBlack
   End If
End Sub

Private Sub Label1_Click(Index As Integer)
   If Index = 1 Then
      Unload Me
   End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 1 Then
      Shape1.Move Shape2.Left, Shape2.Top
      Label1(1).Move Shape2.Left, Shape1.Top + (60 / Screen.TwipsPerPixelY)
   End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 1 Then
      Shape1.BackColor = vbRed
   End If
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 1 Then
      Shape1.Move myLS, myTS
      Label1(1).Move myLS, myTL
   End If
End Sub
