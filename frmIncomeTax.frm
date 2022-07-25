VERSION 5.00
Begin VB.Form frmIncomeTax 
   BackColor       =   &H00004000&
   Caption         =   "Income Tax Calculator"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Taxable Income"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   8325
      TabIndex        =   151
      Top             =   135
      Width           =   3435
      Begin VB.TextBox Text121 
         Height          =   285
         Left            =   2610
         TabIndex        =   165
         Text            =   "Text121"
         Top             =   2565
         Width           =   690
      End
      Begin VB.TextBox Text120 
         Height          =   285
         Left            =   2610
         TabIndex        =   164
         Text            =   "Text120"
         Top             =   2205
         Width           =   690
      End
      Begin VB.TextBox Text119 
         Height          =   285
         Left            =   2610
         TabIndex        =   163
         Text            =   "Text119"
         Top             =   1800
         Width           =   690
      End
      Begin VB.TextBox Text118 
         Height          =   285
         Left            =   2610
         TabIndex        =   162
         Text            =   "Text118"
         Top             =   1440
         Width           =   690
      End
      Begin VB.TextBox Text117 
         Height          =   285
         Left            =   2610
         TabIndex        =   161
         Text            =   "Text117"
         Top             =   1080
         Width           =   690
      End
      Begin VB.TextBox Text116 
         Height          =   285
         Left            =   2610
         TabIndex        =   160
         Text            =   "Text116"
         Top             =   675
         Width           =   735
      End
      Begin VB.TextBox Text115 
         Height          =   285
         Left            =   2610
         TabIndex        =   159
         Text            =   "Text115"
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL TAXABLE INCOME"
         Height          =   195
         Left            =   135
         TabIndex        =   158
         Top             =   2655
         Width           =   1965
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f) House Building Loan Interest"
         Height          =   195
         Left            =   135
         TabIndex        =   157
         Top             =   2295
         Width           =   2175
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e) Entertainment"
         Height          =   195
         Left            =   135
         TabIndex        =   156
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d) Provident Fund"
         Height          =   195
         Left            =   90
         TabIndex        =   155
         Top             =   1485
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c) Conveyance Allowance"
         Height          =   195
         Left            =   90
         TabIndex        =   154
         Top             =   1125
         Width           =   1860
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b) Taxable amount of House Rent"
         Height          =   195
         Left            =   90
         TabIndex        =   153
         Top             =   720
         Width           =   2400
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a) Total Basic Salary and Bonus"
         Height          =   195
         Left            =   90
         TabIndex        =   152
         Top             =   360
         Width           =   2265
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2970
      TabIndex        =   141
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pay and Allowances"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5625
      Left            =   120
      TabIndex        =   12
      Top             =   2355
      Width           =   8085
      Begin VB.TextBox Text114 
         Height          =   330
         Left            =   6300
         TabIndex        =   149
         Text            =   "Text114"
         Top             =   5175
         Width           =   1545
      End
      Begin VB.TextBox Text113 
         Height          =   285
         Left            =   7200
         TabIndex        =   148
         Text            =   "Text113"
         Top             =   4815
         Width           =   645
      End
      Begin VB.TextBox Text112 
         Height          =   285
         Left            =   6480
         TabIndex        =   147
         Text            =   "Text112"
         Top             =   4815
         Width           =   645
      End
      Begin VB.TextBox Text111 
         Height          =   285
         Left            =   5400
         TabIndex        =   146
         Text            =   "Text111"
         Top             =   4815
         Width           =   960
      End
      Begin VB.TextBox Text110 
         Height          =   285
         Left            =   4680
         TabIndex        =   145
         Text            =   "Text110"
         Top             =   4815
         Width           =   645
      End
      Begin VB.TextBox Text109 
         Height          =   285
         Left            =   3735
         TabIndex        =   144
         Text            =   "Text109"
         Top             =   4815
         Width           =   870
      End
      Begin VB.TextBox Text108 
         Height          =   285
         Left            =   2745
         TabIndex        =   143
         Text            =   "Text108"
         Top             =   4815
         Width           =   870
      End
      Begin VB.TextBox Text107 
         Height          =   285
         Left            =   1935
         TabIndex        =   142
         Text            =   "Text107"
         Top             =   4815
         Width           =   735
      End
      Begin VB.TextBox Text100 
         Height          =   285
         Left            =   7200
         TabIndex        =   131
         Text            =   "Text100"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text99 
         Height          =   285
         Left            =   7200
         TabIndex        =   130
         Text            =   "Text99"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text98 
         Height          =   285
         Left            =   7200
         TabIndex        =   129
         Text            =   "Text98"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text97 
         Height          =   285
         Left            =   7200
         TabIndex        =   128
         Text            =   "Text97"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text96 
         Height          =   285
         Left            =   7200
         TabIndex        =   127
         Text            =   "Text96"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text95 
         Height          =   285
         Left            =   7200
         TabIndex        =   126
         Text            =   "Text95"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text94 
         Height          =   285
         Left            =   7200
         TabIndex        =   125
         Text            =   "Text94"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text93 
         Height          =   285
         Left            =   7200
         TabIndex        =   124
         Text            =   "Text93"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text92 
         Height          =   285
         Left            =   7200
         TabIndex        =   123
         Text            =   "Text92"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text91 
         Height          =   285
         Left            =   7200
         TabIndex        =   122
         Text            =   "Text91"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text90 
         Height          =   285
         Left            =   7200
         TabIndex        =   121
         Text            =   "Text90"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text89 
         Height          =   285
         Left            =   7200
         TabIndex        =   120
         Text            =   "Text89"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text88 
         Height          =   285
         Left            =   6480
         TabIndex        =   119
         Text            =   "Text88"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text87 
         Height          =   285
         Left            =   6480
         TabIndex        =   118
         Text            =   "Text87"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text86 
         Height          =   285
         Left            =   6480
         TabIndex        =   117
         Text            =   "Text86"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text85 
         Height          =   285
         Left            =   6480
         TabIndex        =   116
         Text            =   "Text85"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text84 
         Height          =   285
         Left            =   6480
         TabIndex        =   115
         Text            =   "Text84"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text83 
         Height          =   285
         Left            =   6480
         TabIndex        =   114
         Text            =   "Text83"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text82 
         Height          =   285
         Left            =   6480
         TabIndex        =   113
         Text            =   "Text82"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text81 
         Height          =   285
         Left            =   6480
         TabIndex        =   112
         Text            =   "Text81"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text80 
         Height          =   285
         Left            =   6480
         TabIndex        =   111
         Text            =   "Text80"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text79 
         Height          =   285
         Left            =   6480
         TabIndex        =   110
         Text            =   "Text79"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text78 
         Height          =   285
         Left            =   6480
         TabIndex        =   109
         Text            =   "Text78"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text77 
         Height          =   285
         Left            =   6480
         TabIndex        =   108
         Text            =   "Text77"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text76 
         Height          =   285
         Left            =   5400
         TabIndex        =   107
         Text            =   "Text76"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text75 
         Height          =   285
         Left            =   5400
         TabIndex        =   106
         Text            =   "Text75"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text74 
         Height          =   285
         Left            =   5400
         TabIndex        =   105
         Text            =   "Text74"
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text73 
         Height          =   285
         Left            =   5400
         TabIndex        =   104
         Text            =   "Text73"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text72 
         Height          =   285
         Left            =   5400
         TabIndex        =   103
         Text            =   "Text72"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text71 
         Height          =   285
         Left            =   5400
         TabIndex        =   102
         Text            =   "Text71"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text70 
         Height          =   285
         Left            =   5400
         TabIndex        =   101
         Text            =   "Text70"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text69 
         Height          =   285
         Left            =   5400
         TabIndex        =   100
         Text            =   "Text69"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text68 
         Height          =   285
         Left            =   5400
         TabIndex        =   99
         Text            =   "Text68"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text67 
         Height          =   285
         Left            =   5400
         TabIndex        =   98
         Text            =   "Text67"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text66 
         Height          =   285
         Left            =   5400
         TabIndex        =   97
         Text            =   "Text66"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text65 
         Height          =   285
         Left            =   5400
         TabIndex        =   96
         Text            =   "Text65"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text64 
         Height          =   285
         Left            =   4680
         TabIndex        =   95
         Text            =   "Text64"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text63 
         Height          =   285
         Left            =   4680
         TabIndex        =   94
         Text            =   "Text63"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text62 
         Height          =   285
         Left            =   4680
         TabIndex        =   93
         Text            =   "Text62"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Left            =   4680
         TabIndex        =   92
         Text            =   "Text61"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text60 
         Height          =   285
         Left            =   4680
         TabIndex        =   91
         Text            =   "Text60"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text59 
         Height          =   285
         Left            =   4680
         TabIndex        =   90
         Text            =   "Text59"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text58 
         Height          =   285
         Left            =   4680
         TabIndex        =   89
         Text            =   "Text58"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text57 
         Height          =   285
         Left            =   4680
         TabIndex        =   88
         Text            =   "Text57"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text56 
         Height          =   285
         Left            =   4680
         TabIndex        =   87
         Text            =   "Text56"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text55 
         Height          =   285
         Left            =   4680
         TabIndex        =   86
         Text            =   "Text55"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text54 
         Height          =   285
         Left            =   4680
         TabIndex        =   85
         Text            =   "Text54"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text53 
         Height          =   285
         Left            =   4680
         TabIndex        =   84
         Text            =   "Text53"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text52 
         Height          =   285
         Left            =   3720
         TabIndex        =   83
         Text            =   "Text52"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Text51 
         Height          =   285
         Left            =   3720
         TabIndex        =   82
         Text            =   "Text51"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text50 
         Height          =   285
         Left            =   3720
         TabIndex        =   81
         Text            =   "Text50"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text49 
         Height          =   285
         Left            =   3720
         TabIndex        =   80
         Text            =   "Text49"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text48 
         Height          =   285
         Left            =   3720
         TabIndex        =   79
         Text            =   "Text48"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text47 
         Height          =   285
         Left            =   3720
         TabIndex        =   78
         Text            =   "Text47"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text46 
         Height          =   285
         Left            =   3720
         TabIndex        =   77
         Text            =   "Text46"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text45 
         Height          =   285
         Left            =   3720
         TabIndex        =   76
         Text            =   "Text45"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   3720
         TabIndex        =   75
         Text            =   "Text44"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   3720
         TabIndex        =   74
         Text            =   "Text43"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text42 
         Height          =   285
         Left            =   3720
         TabIndex        =   73
         Text            =   "Text42"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   3720
         TabIndex        =   72
         Text            =   "Text41"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   2760
         TabIndex        =   71
         Text            =   "Text40"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   2760
         TabIndex        =   70
         Text            =   "Text39"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   2760
         TabIndex        =   69
         Text            =   "Text38"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text37 
         Height          =   285
         Left            =   2760
         TabIndex        =   68
         Text            =   "Text37"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   2760
         TabIndex        =   67
         Text            =   "Text36"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text35 
         Height          =   285
         Left            =   2760
         TabIndex        =   66
         Text            =   "Text35"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text34 
         Height          =   285
         Left            =   2760
         TabIndex        =   65
         Text            =   "Text34"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text33 
         Height          =   285
         Left            =   2760
         TabIndex        =   64
         Text            =   "Text33"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text32 
         Height          =   285
         Left            =   2760
         TabIndex        =   63
         Text            =   "Text32"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   2760
         TabIndex        =   62
         Text            =   "Text31"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   2760
         TabIndex        =   61
         Text            =   "Text30"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   2760
         TabIndex        =   60
         Text            =   "Text29"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   1920
         TabIndex        =   50
         Text            =   "Text28"
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   1920
         TabIndex        =   49
         Text            =   "Text27"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   1920
         TabIndex        =   48
         Text            =   "Text26"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   1920
         TabIndex        =   47
         Text            =   "Text25"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   1920
         TabIndex        =   46
         Text            =   "Text24"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   1920
         TabIndex        =   45
         Text            =   "Text23"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   1920
         TabIndex        =   44
         Text            =   "Text22"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Text            =   "Text21"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1920
         TabIndex        =   42
         Text            =   "Text20"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Text            =   "Text19"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Text            =   "Text18"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Text            =   "Text17"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Text            =   "Text16"
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Text            =   "Text15"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Text            =   "Text14"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1080
         TabIndex        =   34
         Text            =   "Text13"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Text            =   "Text12"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Text            =   "Text11"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Text            =   "Text10"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Text            =   "Text9"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Text            =   "Text8"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Text            =   "Text7"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Text            =   "Text6"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   5130
         TabIndex        =   150
         Top             =   5220
         Width           =   840
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
         Height          =   195
         Left            =   7200
         TabIndex        =   59
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.F."
         Height          =   195
         Left            =   6600
         TabIndex        =   58
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entertainment"
         Height          =   195
         Left            =   5400
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical"
         Height          =   195
         Left            =   4680
         TabIndex        =   56
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance"
         Height          =   195
         Left            =   3720
         TabIndex        =   55
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Rent"
         Height          =   195
         Left            =   2760
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay"
         Height          =   195
         Left            =   1920
         TabIndex        =   53
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   1395
         TabIndex        =   25
         Top             =   4815
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "December"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "November"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "October"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "September"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "August"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "July"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "June"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "May"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "April"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "March"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "February"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "January"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Income Tax Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4455
      TabIndex        =   7
      Top             =   135
      Width           =   3720
      Begin VB.TextBox Text106 
         Height          =   285
         Left            =   2400
         TabIndex        =   137
         Text            =   "Text106"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text105 
         Height          =   285
         Left            =   1170
         TabIndex        =   136
         Text            =   "Text105"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text104 
         Height          =   285
         Left            =   2400
         TabIndex        =   135
         Text            =   "Text104"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text103 
         Height          =   285
         Left            =   1170
         TabIndex        =   134
         Text            =   "Text103"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text102 
         Height          =   285
         Left            =   2400
         TabIndex        =   133
         Text            =   "Text102"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text101 
         Height          =   285
         Left            =   1170
         TabIndex        =   132
         Text            =   "Text101"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1170
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entertainment"
         Height          =   195
         Left            =   135
         TabIndex        =   140
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance"
         Height          =   195
         Left            =   135
         TabIndex        =   139
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Rent"
         Height          =   195
         Left            =   120
         TabIndex        =   138
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Year"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4185
      Begin VB.ComboBox cmbDesignation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1230
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1500
         Width           =   1425
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1230
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   405
         Width           =   945
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1245
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   960
         Width           =   2730
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   1005
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   1560
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmIncomeTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
Text5.Text = Text1.Text
Text6.Text = Text1.Text
Text7.Text = Text1.Text
Text8.Text = Text1.Text
Text9.Text = Text1.Text
Text10.Text = Text1.Text
Text11.Text = Text2.Text
Text12.Text = Text2.Text
Text13.Text = Text2.Text
Text14.Text = Text2.Text
Text15.Text = Text2.Text
Text16.Text = Text2.Text

Text17.Text = Text3.Text
Text18.Text = Text3.Text
Text19.Text = Text3.Text
Text20.Text = Text3.Text
Text21.Text = Text3.Text
Text22.Text = Text3.Text
Text23.Text = Text4.Text
Text24.Text = Text4.Text
Text25.Text = Text4.Text
Text26.Text = Text4.Text
Text27.Text = Text4.Text
Text28.Text = Text4.Text

Text29.Text = Text101.Text
Text30.Text = Text101.Text
Text31.Text = Text101.Text
Text32.Text = Text101.Text
Text33.Text = Text101.Text
Text34.Text = Text101.Text
Text35.Text = Text102.Text
Text36.Text = Text102.Text
Text37.Text = Text102.Text
Text38.Text = Text102.Text
Text39.Text = Text102.Text
Text40.Text = Text102.Text

Text41.Text = Text103.Text
Text42.Text = Text103.Text
Text43.Text = Text103.Text
Text44.Text = Text103.Text
Text45.Text = Text103.Text
Text46.Text = Text103.Text
Text47.Text = Text104.Text
Text48.Text = Text104.Text
Text49.Text = Text104.Text
Text50.Text = Text104.Text
Text51.Text = Text104.Text
Text52.Text = Text104.Text

Text53.Text = Val(Text3.Text) * 0.1
Text54.Text = Val(Text3.Text) * 0.1
Text55.Text = Val(Text3.Text) * 0.1
Text56.Text = Val(Text3.Text) * 0.1
Text57.Text = Val(Text3.Text) * 0.1
Text58.Text = Val(Text3.Text) * 0.1
Text59.Text = Val(Text4.Text) * 0.1
Text60.Text = Val(Text4.Text) * 0.1
Text61.Text = Val(Text4.Text) * 0.1
Text62.Text = Val(Text4.Text) * 0.1
Text63.Text = Val(Text4.Text) * 0.1
Text64.Text = Val(Text4.Text) * 0.1

Text65.Text = Text105.Text
Text66.Text = Text105.Text
Text67.Text = Text105.Text
Text68.Text = Text105.Text
Text69.Text = Text105.Text
Text70.Text = Text105.Text
Text71.Text = Text106.Text
Text72.Text = Text106.Text
Text73.Text = Text106.Text
Text74.Text = Text106.Text
Text75.Text = Text106.Text
Text76.Text = Text106.Text

Text77.Text = Val(Text3.Text) * 0.1
Text78.Text = Val(Text3.Text) * 0.1
Text79.Text = Val(Text3.Text) * 0.1
Text80.Text = Val(Text3.Text) * 0.1
Text81.Text = Val(Text3.Text) * 0.1
Text82.Text = Val(Text3.Text) * 0.1
Text83.Text = Val(Text4.Text) * 0.1
Text84.Text = Val(Text4.Text) * 0.1
Text85.Text = Val(Text4.Text) * 0.1
Text86.Text = Val(Text4.Text) * 0.1
Text87.Text = Val(Text4.Text) * 0.1
Text88.Text = Val(Text4.Text) * 0.1

Text89.Text = 0
Text90.Text = 0
Text91.Text = 0
Text92.Text = Val(Text3.Text) * 1
Text93.Text = 0
Text94.Text = Val(Text3.Text) * 1
Text95.Text = 0
Text96.Text = 0
Text97.Text = Val(Text4.Text) * 3
Text98.Text = 0
Text99.Text = 0
Text100.Text = 0

Text107.Text = Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text) + Val(Text26.Text) + Val(Text27.Text) + Val(Text28.Text)
Text108.Text = Val(Text29.Text) + Val(Text30.Text) + Val(Text31.Text) + Val(Text32.Text) + Val(Text33.Text) + Val(Text34.Text) + Val(Text35.Text) + Val(Text36.Text) + Val(Text37.Text) + Val(Text38.Text) + Val(Text39.Text) + Val(Text40.Text)
Text109.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text) + Val(Text48.Text) + Val(Text49.Text) + Val(Text50.Text) + Val(Text51.Text) + Val(Text52.Text)
Text110.Text = Val(Text53.Text) + Val(Text54.Text) + Val(Text55.Text) + Val(Text56.Text) + Val(Text57.Text) + Val(Text58.Text) + Val(Text59.Text) + Val(Text60.Text) + Val(Text61.Text) + Val(Text62.Text) + Val(Text63.Text) + Val(Text64.Text)
Text111.Text = Val(Text65.Text) + Val(Text66.Text) + Val(Text67.Text) + Val(Text68.Text) + Val(Text69.Text) + Val(Text70.Text) + Val(Text71.Text) + Val(Text72.Text) + Val(Text73.Text) + Val(Text74.Text) + Val(Text75.Text) + Val(Text76.Text)
Text112.Text = Val(Text77.Text) + Val(Text78.Text) + Val(Text79.Text) + Val(Text80.Text) + Val(Text81.Text) + Val(Text82.Text) + Val(Text83.Text) + Val(Text84.Text) + Val(Text85.Text) + Val(Text86.Text) + Val(Text87.Text) + Val(Text88.Text)
Text113.Text = Val(Text89.Text) + Val(Text90.Text) + Val(Text91.Text) + Val(Text92.Text) + Val(Text93.Text) + Val(Text94.Text) + Val(Text95.Text) + Val(Text96.Text) + Val(Text97.Text) + Val(Text98.Text) + Val(Text99.Text) + Val(Text100.Text)

Text114.Text = Val(Text107.Text) + Val(Text108.Text) + Val(Text109.Text) + Val(Text110.Text) + Val(Text111.Text) + Val(Text112.Text) + Val(Text113.Text)

Text115.Text = Val(Text107.Text) + Val(Text113.Text)

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    txtName.SetFocus
End If
End Sub

Private Sub txtID_LostFocus()
If txtID.Text = "" Then
Exit Sub
End If

Dim ID As String
    ID = txtID.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Salary where ID like '" & ID & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtID.Text = rs!ID
        txtName.Text = rs!Name
        cmbDesignation.Text = rs!Designation
        Text3.Text = Format$(rs![Basic Pay], "######0.00")
        Text101.Text = Format$(rs![House Rent Allowance], "######0.00")
        
        Text103.Text = Format$(rs![Conveyance Allowance], "######0.00")
        Text105.Text = Format$(rs![Entertainment Allowance], "######0.00")
        
        rs.Close
        
        
    Else
    If MsgBox("There is no such ID found, Do you want add new employee?", vbCritical + vbYesNo) = vbYes Then
       
           
        rs.Close
    Else
        Exit Sub
    End If
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub
