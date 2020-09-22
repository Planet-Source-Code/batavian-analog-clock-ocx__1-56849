VERSION 5.00
Begin VB.UserControl AnalogClock 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   ClipControls    =   0   'False
   PaletteMode     =   2  'Custom
   PropertyPages   =   "ClockOCX.ctx":0000
   ScaleHeight     =   3240
   ScaleWidth      =   3195
   ToolboxBitmap   =   "ClockOCX.ctx":0014
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2595
      Top             =   2625
   End
End
Attribute VB_Name = "AnalogClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Const ALTERNATE As Long = 1
Private Const Pi As Double = 3.14159265358979
Private Const WINDING As Long = 2

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Default Property Values:
Const m_def_ShowMinorP = True
Const m_def_ShowMajorP = True
Const m_def_MinuteOutline = &H80FF&    'RGB(255, 64, 0)
Const m_def_HourOutline = &H8080FF     'RGB(255, 64, 64)
Const m_def_MajorPoint = &HFFFFFF      'RGB(255, 255, 255)
Const m_def_MinorPoint = &H808080      'RGB(128, 128, 128)
Const m_def_SecondPointer = &HFFFFFF   'vbWhite
Const m_def_MinutePointer = &H80C0FF   'RGB(255, 192, 64)
Const m_def_HourPointer = &HC0C0FF     'RGB(255, 192, 192)
Const m_def_CircleBorder = &H0&        'vbBlack
Const m_def_ClockBody = &H400080       'RGB(128, 0, 64)

'Property Variables:
Dim m_ShowMinorP As Boolean
Dim m_ShowMajorP As Boolean
Dim m_MinuteOutline As OLE_COLOR
Dim m_HourOutline As OLE_COLOR
Dim m_MajorPoint As OLE_COLOR
Dim m_MinorPoint As OLE_COLOR
Dim m_SecondPointer As OLE_COLOR
Dim m_MinutePointer As OLE_COLOR
Dim m_HourPointer As OLE_COLOR
Dim m_CircleBorder As OLE_COLOR
Dim m_ClockBody As OLE_COLOR

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove

Dim XX As Long, YY As Long

Private Function Dec2Rad(ByVal lDec As Long) As Double 'Convert Decimal To Radian
Dim dRad As Double

   dRad = Pi / 180
   Dec2Rad = lDec * dRad
End Function

Private Sub ShowTime()
Dim dH As Long, dH1 As Long, dH2 As Long, iH As Integer 'Hour Variables
Dim dM As Long, dM1 As Long, dM2 As Long, iM As Integer 'Minute Variables
Dim dS As Long 'Second Variable

Dim dHX As Double, dHX1 As Double, dHX2 As Double 'Hour Variables
Dim dHY As Double, dHY1 As Double, dHY2 As Double
Dim dMX As Double, dMX1 As Double, dMX2 As Double 'Minute Variables
Dim dMY As Double, dMY1 As Double, dMY2 As Double
Dim dSX As Double, dSY As Double 'Second Variables

Dim pP(1 To 4) As POINTAPI 'The Polygons Points
Dim hRgn As Long, hBrush As Long 'Fill Color

Dim fW As Long, fH As Long
Dim iCircle As Integer, R As Integer, G As Integer, B As Integer
   

   fW = UserControl.ScaleWidth: fH = UserControl.ScaleHeight
   
   iH = Hour(Time)                                 '| Get the Current Hour
   If iH > 12 Then iH = iH - 12                    '| Make it 12 Hour Format
   If iH = 0 Then iH = 12                          '|
   dH = (iH * 30) + (Int(Minute(Time) / 12) * 6)   '| Hour Original
   dH1 = dH - 40                                   '| Hour Outer
   dH2 = dH + 40                                   '| Hour Inner
   
   iM = Minute(Time)                               '| Get the Current Minute
   If iM = 0 Then iM = 60                          '|
   dM = iM * 6                                     '| Minute Original
   dM1 = dM - 40                                   '| Minute Outer
   dM2 = dM + 40                                   '| Minute Inner
   
   dS = Int(Timer) * 6                             '| Second Code

   dHX = Sin(Dec2Rad(dH))     '| Hour Point.X
   dHY = -Cos(Dec2Rad(dH))    '| Hour Point.Y
   dHX1 = Sin(Dec2Rad(dH1))   '| Hour Left Point.X
   dHY1 = -Cos(Dec2Rad(dH1))  '| Hour Left Point.Y
   dHX2 = Sin(Dec2Rad(dH2))   '| Hour Right Point.X
   dHY2 = -Cos(Dec2Rad(dH2))  '| Hour Right Point.X
                                                            
   dMX = Sin(Dec2Rad(dM))     '| Minute Point.X
   dMY = -Cos(Dec2Rad(dM))    '| Minute Point.Y
   dMX1 = Sin(Dec2Rad(dM1))   '| Minute Left Point.X
   dMY1 = -Cos(Dec2Rad(dM1))  '| Minute Left Point.Y
   dMX2 = Sin(Dec2Rad(dM2))   '| Minute Right Point.X
   dMY2 = -Cos(Dec2Rad(dM2))  '| Minute Right Point.Y
   
   dSX = Sin(Dec2Rad(dS))     '| The Second
   dSY = -Cos(Dec2Rad(dS))    '|
   
   UserControl.Cls            '| Clear the Form
   UserControl.DrawStyle = 5  '| Set to Transparent
   UserControl.FillStyle = 0  '| Set to Solid
   
   hRgn = CreateEllipticRgn(2, 2, XX + XX, YY + YY)   '| Clock's Region >>--------------------------\
   hBrush = CreateSolidBrush(m_ClockBody)             '|                                            |
   FillRgn UserControl.hdc, hRgn, hBrush              '| Fill the Clock                             |
   DeleteObject hBrush                                '|                                            |
                                                      '|                                            |
   R = m_ClockBody And &HFF                           '| Convert the Color to RGB Value             |
   G = m_ClockBody \ &H100 And &HFF                   '|                                            |
   B = m_ClockBody \ &H10000 And &HFF                 '|                                            |
                                                      '|                                            |
   For iCircle = ((XX + YY) / 2) * 2 To 0 Step -1     '| Fill the Clock with Radial Gradation Style |
      UserControl.FillColor = RGB(R, G, B)            '|                                            |
      UserControl.Circle (XX, YY), iCircle              '|                                            |
      R = R + 1: G = G + 1: B = B + 1                 '|                                            |
   Next iCircle                                       '|                                            |
                                                      '|                                            |
   UserControl.DrawStyle = 0                          '| Set Back to Solid                          |
   UserControl.FillStyle = 1                          '| Set Back to Transparent                    |
                                                      '|                                            |
   hBrush = CreateSolidBrush(m_CircleBorder)          '| Set Frame Color                            |
   FrameRgn UserControl.hdc, hRgn, hBrush, 1, 1       '| Draw the Frame of the Region <<------------/
   DeleteObject hBrush
   DeleteObject hRgn
   
   pP(1).x = (dMX * -Round(fW / 19)) + XX:      pP(1).y = (dMY * -Round(fH / 19)) + YY       '|  \
   pP(2).x = (dMX1 * Round(fW / 19)) + pP(1).x: pP(2).y = (dMY1 * Round(fH / 19)) + pP(1).y  '|   > The Minutes's Pointer Points
   pP(3).x = (dMX * Round(fW / 2.2)) + pP(1).x: pP(3).y = (dMY * Round(fH / 2.2)) + pP(1).y  '|  /
   pP(4).x = (dMX2 * Round(fW / 19)) + pP(1).x: pP(4).y = (dMY2 * Round(fH / 19)) + pP(1).y  '| /
   
   hRgn = CreatePolygonRgn(pP(1), 4, WINDING)   '| Create the Minute Region
   OffsetRgn hRgn, 2, 2                         '| Shadow First  <<--------\
   hBrush = CreateSolidBrush(RGB(64, 64, 64))   '| Create a Color Handle   |
   FillRgn UserControl.hdc, hRgn, hBrush        '| Fill the Shadow         |
   OffsetRgn hRgn, -2, -2                       '| Then the Pointer  <<----/
   DeleteObject hBrush                          '| RELEASE THE MEMORY HANDLE, TO AVOID GDI MEMORY LEAK *)
   
   hBrush = CreateSolidBrush(m_MinutePointer)
   FillRgn UserControl.hdc, hRgn, hBrush        '| Fill the Minute Region
   DeleteObject hBrush
   hBrush = CreateSolidBrush(m_MinuteOutline)
   FrameRgn Me.hdc, hRgn, hBrush, 1, 1          '| Draw a Frame On the Minute Region
   DeleteObject hBrush
   DeleteObject hRgn
   
   pP(1).x = (dHX * -Round(fW / 19)) + XX:      pP(1).y = (dHY * -Round(fH / 19)) + YY       '|  \
   pP(2).x = (dHX1 * Round(fW / 19)) + pP(1).x: pP(2).y = (dHY1 * Round(fH / 19)) + pP(1).y  '|   > The Hour's Pointer Polygon Points
   pP(3).x = (dHX * Round(fW / 2.8)) + pP(1).x: pP(3).y = (dHY * Round(fH / 2.8)) + pP(1).y  '|  /
   pP(4).x = (dHX2 * Round(fW / 19)) + pP(1).x: pP(4).y = (dHY2 * Round(fH / 19)) + pP(1).y  '| /

   hRgn = CreatePolygonRgn(pP(1), 4, WINDING)   '| Create the Hour's Region
   OffsetRgn hRgn, 2, 2                         '| Shadow First  <<--------\
   hBrush = CreateSolidBrush(RGB(64, 64, 64))   '| Create a Color Handle   |
   FillRgn UserControl.hdc, hRgn, hBrush        '| Fill the Shadow         |
   OffsetRgn hRgn, -2, -2                       '| Then the Pointer  <<----/
   DeleteObject hBrush
   
   hBrush = CreateSolidBrush(m_HourPointer)
   FillRgn UserControl.hdc, hRgn, hBrush        '| Fill the Hour Region
   DeleteObject hBrush
   hBrush = CreateSolidBrush(m_HourOutline)
   FrameRgn Me.hdc, hRgn, hBrush, 1, 1          '| Draw a Frame On the Hour Region
   DeleteObject hBrush
   DeleteObject hRgn
   
   UserControl.Line ((dSX * -Round(fW / 10)) + XX + 2, (dSY * -Round(fH / 10)) + YY + 2)- _
                    ((dSX * Round(fW / 2.5)) + XX + 2, (dSY * Round(fH / 2.5)) + YY + 2), _
                    RGB(64, 64, 64)             '| Crate a Shadow of the Second Pointer
                    
   UserControl.Line ((dSX * -Round(fW / 10)) + XX, (dSY * -Round(fH / 10)) + YY)- _
                    ((dSX * Round(fW / 2.5)) + XX, (dSY * Round(fH / 2.5)) + YY), _
                    m_SecondPointer             '| Now Create the Simple Second Pointer
   
   UserControl.Circle (XX, YY), 0, vbBlack      '| Draw the Pointers Axis
   UserControl.Circle (XX, YY), 1, vbBlack      '| Draw the Pointers Axis Again
   
   For iCircle = 6 To 360 Step 6                '| Draw the Points
      dHX = Sin(Dec2Rad(iCircle))
      dHY = -(Cos(Dec2Rad(iCircle)))
      dHX1 = Sin(Dec2Rad(iCircle - 1))
      dHY1 = -(Cos(Dec2Rad(iCircle - 1)))
      dHX2 = Sin(Dec2Rad(iCircle + 1))
      dHY2 = -(Cos(Dec2Rad(iCircle + 1)))
      If m_ShowMajorP Then
         If iCircle Mod 30 = 0 Then
            UserControl.Line ((dHX * Round(fW / 2.1)) + XX, (dHY * Round(fH / 2.1)) + YY)- _
                             ((dHX * Round(fW / 2.3)) + XX, (dHY * Round(fH / 2.3)) + YY), m_MajorPoint
            If iCircle Mod 90 = 0 Then
               UserControl.Line ((dHX1 * Round(fW / 2.1)) + XX, (dHY1 * Round(fH / 2.1)) + YY)- _
                                ((dHX2 * Round(fW / 2.3)) + XX, (dHY2 * Round(fH / 2.3)) + YY), m_MajorPoint, BF
               UserControl.Line ((dHX * Round(fW / 2.1)) + XX, (dHY * Round(fH / 2.1)) + YY)- _
                                ((dHX * Round(fW / 2.3)) + XX, (dHY * Round(fH / 2.3)) + YY), m_ClockBody
            End If
         Else
            If m_ShowMinorP Then
               UserControl.Circle ((dHX * Round(fW / 2.2)) + XX, (dHY * Round(fH / 2.2)) + YY), 0, m_MinorPoint
               UserControl.Circle ((dHX * Round(fW / 2.2)) + XX, (dHY * Round(fH / 2.2)) + YY), 1, m_MinorPoint
            End If
         End If
      Else
         If m_ShowMinorP Then
            UserControl.Circle ((dHX * Round(fW / 2.2)) + XX, (dHY * Round(fH / 2.2)) + YY), 0, m_MinorPoint
            UserControl.Circle ((dHX * Round(fW / 2.2)) + XX, (dHY * Round(fH / 2.2)) + YY), 1, m_MinorPoint
         End If
      End If
   Next iCircle
   
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
   UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
   ShowTime
End Sub

Private Sub UserControl_Initialize()
   UserControl.ScaleMode = vbPixels
   UserControl.Refresh
   ShowTime
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   Timer1.Enabled = Ambient.UserMode
   m_ShowMajorP = m_def_ShowMajorP
   m_MinuteOutline = m_def_MinuteOutline
   m_HourOutline = m_def_HourOutline
   m_MajorPoint = m_def_MajorPoint
   m_MinorPoint = m_def_MinorPoint
   m_SecondPointer = m_def_SecondPointer
   m_MinutePointer = m_def_MinutePointer
   m_HourPointer = m_def_HourPointer
   m_CircleBorder = m_def_CircleBorder
   m_ClockBody = m_def_ClockBody
   m_ShowMinorP = m_def_ShowMinorP
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Timer1.Enabled = Ambient.UserMode
   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   m_ShowMajorP = PropBag.ReadProperty("ShowMajorP", m_def_ShowMajorP)
   m_MinuteOutline = PropBag.ReadProperty("MinuteOutline", m_def_MinuteOutline)
   m_HourOutline = PropBag.ReadProperty("HourOutline", m_def_HourOutline)
   m_MajorPoint = PropBag.ReadProperty("MajorPoint", m_def_MajorPoint)
   m_MinorPoint = PropBag.ReadProperty("MinorPoint", m_def_MinorPoint)
   m_SecondPointer = PropBag.ReadProperty("SecondPointer", m_def_SecondPointer)
   m_MinutePointer = PropBag.ReadProperty("MinutePointer", m_def_MinutePointer)
   m_HourPointer = PropBag.ReadProperty("HourPointer", m_def_HourPointer)
   m_CircleBorder = PropBag.ReadProperty("CircleBorder", m_def_CircleBorder)
   m_ClockBody = PropBag.ReadProperty("ClockBody", m_def_ClockBody)
   m_ShowMinorP = PropBag.ReadProperty("ShowMinorP", m_def_ShowMinorP)
End Sub

Private Sub UserControl_Resize()
Dim hRgn As Long

   XX = UserControl.ScaleWidth / 2
   YY = UserControl.ScaleHeight / 2
   
   hRgn = CreateEllipticRgn(2, 2, XX + XX, YY + YY)   'Clock's Face
   SetWindowRgn UserControl.hwnd, hRgn, True          'Set to a Circular Form (Not a Rectangle :)

   ShowTime
End Sub

Private Sub UserControl_Show()
   ShowTime
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
   Call PropBag.WriteProperty("ShowMajorP", m_ShowMajorP, m_def_ShowMajorP)
   Call PropBag.WriteProperty("MinuteOutline", m_MinuteOutline, m_def_MinuteOutline)
   Call PropBag.WriteProperty("HourOutline", m_HourOutline, m_def_HourOutline)
   Call PropBag.WriteProperty("MajorPoint", m_MajorPoint, m_def_MajorPoint)
   Call PropBag.WriteProperty("MinorPoint", m_MinorPoint, m_def_MinorPoint)
   Call PropBag.WriteProperty("SecondPointer", m_SecondPointer, m_def_SecondPointer)
   Call PropBag.WriteProperty("MinutePointer", m_MinutePointer, m_def_MinutePointer)
   Call PropBag.WriteProperty("HourPointer", m_HourPointer, m_def_HourPointer)
   Call PropBag.WriteProperty("CircleBorder", m_CircleBorder, m_def_CircleBorder)
   Call PropBag.WriteProperty("ClockBody", m_ClockBody, m_def_ClockBody)
   Call PropBag.WriteProperty("ShowMinorP", m_ShowMinorP, m_def_ShowMinorP)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hdc = UserControl.hdc
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get ShowMajorP() As Boolean
   ShowMajorP = m_ShowMajorP
End Property

Public Property Let ShowMajorP(ByVal New_ShowMajorP As Boolean)
   m_ShowMajorP = New_ShowMajorP
   PropertyChanged "ShowMajorP"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinuteOutline() As OLE_COLOR
   MinuteOutline = m_MinuteOutline
End Property

Public Property Let MinuteOutline(ByVal New_MinuteOutline As OLE_COLOR)
   m_MinuteOutline = New_MinuteOutline
   PropertyChanged "MinuteOutline"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HourOutline() As OLE_COLOR
   HourOutline = m_HourOutline
End Property

Public Property Let HourOutline(ByVal New_HourOutline As OLE_COLOR)
   m_HourOutline = New_HourOutline
   PropertyChanged "HourOutline"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MajorPoint() As OLE_COLOR
   MajorPoint = m_MajorPoint
End Property

Public Property Let MajorPoint(ByVal New_MajorPoint As OLE_COLOR)
   m_MajorPoint = New_MajorPoint
   PropertyChanged "MajorPoint"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinorPoint() As OLE_COLOR
   MinorPoint = m_MinorPoint
End Property

Public Property Let MinorPoint(ByVal New_MinorPoint As OLE_COLOR)
   m_MinorPoint = New_MinorPoint
   PropertyChanged "MinorPoint"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SecondPointer() As OLE_COLOR
   SecondPointer = m_SecondPointer
End Property

Public Property Let SecondPointer(ByVal New_SecondPointer As OLE_COLOR)
   m_SecondPointer = New_SecondPointer
   PropertyChanged "SecondPointer"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinutePointer() As OLE_COLOR
   MinutePointer = m_MinutePointer
End Property

Public Property Let MinutePointer(ByVal New_MinutePointer As OLE_COLOR)
   m_MinutePointer = New_MinutePointer
   PropertyChanged "MinutePointer"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HourPointer() As OLE_COLOR
   HourPointer = m_HourPointer
End Property

Public Property Let HourPointer(ByVal New_HourPointer As OLE_COLOR)
   m_HourPointer = New_HourPointer
   PropertyChanged "HourPointer"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CircleBorder() As OLE_COLOR
   CircleBorder = m_CircleBorder
End Property

Public Property Let CircleBorder(ByVal New_CircleBorder As OLE_COLOR)
   m_CircleBorder = New_CircleBorder
   PropertyChanged "CircleBorder"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ClockBody() As OLE_COLOR
   ClockBody = m_ClockBody
End Property

Public Property Let ClockBody(ByVal New_ClockBody As OLE_COLOR)
   m_ClockBody = New_ClockBody
   PropertyChanged "ClockBody"
   ShowTime
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowMinorP() As Boolean
   ShowMinorP = m_ShowMinorP
End Property

Public Property Let ShowMinorP(ByVal New_ShowMinorP As Boolean)
   m_ShowMinorP = New_ShowMinorP
   PropertyChanged "ShowMinorP"
   ShowTime
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
Dim frms As Form

   For Each frms In Forms
      If frms.Name = "frmAbout" Then Unload frms
   Next frms
   frmAbout.Show
End Sub
