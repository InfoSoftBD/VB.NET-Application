VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "Mscal.ocx"
Begin VB.Form frmClock 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "My Clock"
   ClientHeight    =   10185
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   3705
   ControlBox      =   0   'False
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   135
      Top             =   -360
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2475
      Left            =   150
      TabIndex        =   0
      Top             =   3285
      Width           =   3315
      _Version        =   524288
      _ExtentX        =   5847
      _ExtentY        =   4366
      _StockProps     =   1
      BackColor       =   -2147483643
      Year            =   2012
      Month           =   5
      Day             =   19
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   6
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   0
      ShowDateSelectors=   0   'False
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   0
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   165
      Left            =   1710
      Shape           =   2  'Oval
      Top             =   1305
      Width           =   165
   End
   Begin VB.Image Image9 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":0442
      Stretch         =   -1  'True
      Top             =   8325
      Width           =   2925
   End
   Begin VB.Image Image8 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":2BAC
      Stretch         =   -1  'True
      Top             =   7215
      Width           =   2895
   End
   Begin VB.Image Image6 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":5DBA
      Stretch         =   -1  'True
      Top             =   6105
      Width           =   2895
   End
   Begin VB.Image Image7 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":9606
      Stretch         =   -1  'True
      Top             =   8325
      Width           =   2925
   End
   Begin VB.Image Image5 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":BA61
      Stretch         =   -1  'True
      Top             =   7215
      Width           =   2895
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   390
      Picture         =   "Clock.frx":E960
      Stretch         =   -1  'True
      Top             =   6105
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1155
      TabIndex        =   4
      Top             =   3315
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "December - 2012"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   660
      TabIndex        =   3
      Top             =   5355
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   1095
      TabIndex        =   2
      Top             =   3765
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rakib's Clock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1350
      TabIndex        =   1
      Top             =   810
      Width           =   945
   End
   Begin VB.Line Linehour 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   1785
      X2              =   1785
      Y1              =   2400
      Y2              =   1380
   End
   Begin VB.Line Linemin 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   1080
      X2              =   1800
      Y1              =   1080
      Y2              =   1380
   End
   Begin VB.Line linesec 
      BorderWidth     =   2
      X1              =   1785
      X2              =   1785
      Y1              =   150
      Y2              =   1380
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   345
      Picture         =   "Clock.frx":11EAA
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2745
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   3075
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   2865
      Left            =   0
      Picture         =   "Clock.frx":15DEE
      Stretch         =   -1  'True
      Top             =   3015
      Width           =   3630
   End
   Begin VB.Image Image3 
      Height          =   3075
      Left            =   360
      Picture         =   "Clock.frx":1F5C8
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   2940
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_DblClick()
Image2.Visible = False
Calendar1.Visible = False
Call SetWindowRgn(Me.hwnd, TransparentForm(App.Path + "\Photo\Clock Back Short.bmp"), True)
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Shape2.Visible = True
Dim Mont As String
 Dim yr As String
 Mont = MonthName(Month(CDate(Today)))
yr = Year(CDate(Today))
Label3.Caption = Mont & " - " & yr
Label2.Caption = Day(CDate(Today))
Label4.Caption = WeekdayName(Weekday(CDate(Today)))
End Sub

Private Sub Form_Load()
frmClock.Top = 700
frmClock.Left = wdt
Call SetWindowRgn(Me.hwnd, TransparentForm(App.Path + "\Photo\Clock Back Short.bmp"), True)

Image3.Visible = True
Image2.Visible = False
Calendar1.Visible = False

Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Shape2.Visible = True
Dim Mont As String
 Dim yr As String
 Mont = MonthName(Month(CDate(Today)))
yr = Year(CDate(Today))
Label3.Caption = Mont & " - " & yr
Label2.Caption = Day(CDate(Today))
Label4.Caption = WeekdayName(Weekday(CDate(Today)))

linesec.X1 = Cos((Second(Now) * 3.14 / 30) - 3.14 / 2) * 1100 + linesec.X2
linesec.Y1 = Sin((Second(Now) * 3.14 / 30) - 3.14 / 2) * 1100 + linesec.Y2
Linemin.X1 = Cos((Minute(Now) * 3.14 / 30) - 3.14 / 2) * 900 + Linemin.X2
Linemin.Y1 = Sin((Minute(Now) * 3.14 / 30) - 3.14 / 2) * 900 + Linemin.Y2
Linehour.X1 = Cos(((Hour(Now) * 30) + Minute(Now) / 2) * 3.14 / 180 - 3.14 / 2) * 700 + Linehour.X2
Linehour.Y1 = Sin(((Hour(Now) * 30) + Minute(Now) / 2) * 3.14 / 180 - 3.14 / 2) * 700 + Linehour.Y2
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image4.Visible = False
Image5.Visible = True
End Sub
Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.Visible = False
Image4.Visible = True
End Sub
Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.Visible = True
Image4.Visible = False
frmProd_Receive.Show 1
End Sub
Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image8.Visible = True
Image5.Visible = False
frmSales.Show 1
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image8.Visible = False
Image5.Visible = True
End Sub
Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image9.Visible = False
Image7.Visible = True
End Sub
Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image9.Visible = True
Image7.Visible = False
If MsgBox("Really want to quit?", vbInformation + vbYesNo, "Exit") = vbYes Then
MsgBox "Thank you n' have a nice day", vbInformation, "Thanks!"
EndPlaySound
End
Else
Exit Sub
End If
End Sub


Private Sub Label1_Click()
Unload Me
End Sub
Private Sub Label2_DblClick()
Image2.Visible = True
Calendar1.Visible = True
Calendar1.Value = CDate(Today)
Call SetWindowRgn(Me.hwnd, TransparentForm(App.Path + "\Photo\Clock Back Long.bmp"), True)
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Shape2.Visible = False
End Sub

Private Sub Timer1_Timer()
linesec.X1 = Cos((Second(Now) * 3.14 / 30) - 3.14 / 2) * 1100 + linesec.X2
linesec.Y1 = Sin((Second(Now) * 3.14 / 30) - 3.14 / 2) * 1100 + linesec.Y2
Linemin.X1 = Cos((Minute(Now) * 3.14 / 30) - 3.14 / 2) * 900 + Linemin.X2
Linemin.Y1 = Sin((Minute(Now) * 3.14 / 30) - 3.14 / 2) * 900 + Linemin.Y2
Linehour.X1 = Cos(((Hour(Now) * 30) + Minute(Now) / 2) * 3.14 / 180 - 3.14 / 2) * 700 + Linehour.X2
Linehour.Y1 = Sin(((Hour(Now) * 30) + Minute(Now) / 2) * 3.14 / 180 - 3.14 / 2) * 700 + Linehour.Y2
'Label5.Caption = Time
End Sub

