VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSalary 
   BackColor       =   &H00004000&
   Caption         =   "Entry Form"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DEDUCTION FROM SALARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   180
      TabIndex        =   54
      Top             =   6930
      Width           =   8160
      Begin VB.TextBox txtDeduction 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5940
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   2340
         Width           =   2040
      End
      Begin VB.TextBox txtFlexyloan 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6690
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   1380
         Width           =   1275
      End
      Begin VB.TextBox txtOther_Loan 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6690
         TabIndex        =   61
         Text            =   "0.00"
         Top             =   1845
         Width           =   1275
      End
      Begin VB.TextBox txtPfloan 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6690
         TabIndex        =   60
         Text            =   "0.00"
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtShbl 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6690
         TabIndex        =   59
         Text            =   "0.00"
         Top             =   390
         Width           =   1275
      End
      Begin VB.TextBox txtDeathrisk 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2685
         TabIndex        =   58
         Text            =   "0.00"
         Top             =   1890
         Width           =   1275
      End
      Begin VB.TextBox txtWelfare 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2685
         TabIndex        =   57
         Text            =   "0.00"
         Top             =   1410
         Width           =   1275
      End
      Begin VB.TextBox txtIncometax 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2685
         TabIndex        =   56
         Text            =   "0.00"
         Top             =   930
         Width           =   1275
      End
      Begin VB.TextBox txtPFboth 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2685
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deduction Tk."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4020
         TabIndex        =   72
         Top             =   2430
         Width           =   1845
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flexy Loan Installment Tk."
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
         Left            =   4365
         TabIndex        =   70
         Top             =   1455
         Width           =   2295
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Loan Installment Tk."
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
         Left            =   4320
         TabIndex        =   69
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P/F Loan Installment Tk."
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
         Left            =   4530
         TabIndex        =   68
         Top             =   990
         Width           =   2130
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHBL Loan Installment Tk."
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
         Left            =   4335
         TabIndex        =   67
         Top             =   465
         Width           =   2325
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Death Risk Coverage Tk."
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
         Left            =   510
         TabIndex        =   66
         Top             =   1965
         Width           =   2145
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employees Welfare Fund Tk."
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
         TabIndex        =   65
         Top             =   1485
         Width           =   2520
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax Tk."
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
         Left            =   1320
         TabIndex        =   64
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P/F Both Contribution Tk."
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
         Left            =   450
         TabIndex        =   63
         Top             =   525
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PAY AND ALLOWANCES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Left            =   180
      TabIndex        =   21
      Top             =   1935
      Width           =   8160
      Begin VB.TextBox txtCleaner 
         Alignment       =   1  'Right Justify
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
         Left            =   6660
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   3825
         Width           =   1275
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5925
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   4320
         Width           =   2040
      End
      Begin VB.TextBox txtWater 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6675
         TabIndex        =   53
         Text            =   "0.00"
         Top             =   3300
         Width           =   1275
      End
      Begin VB.TextBox txtNightguard 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6675
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   2775
         Width           =   1275
      End
      Begin VB.TextBox txtTeaboy 
         Alignment       =   1  'Right Justify
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
         Left            =   6675
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   2295
         Width           =   1275
      End
      Begin VB.TextBox txtDriver 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6675
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   1815
         Width           =   1275
      End
      Begin VB.TextBox txtWashing 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6675
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   1350
         Width           =   1275
      End
      Begin VB.TextBox txtGuard 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6675
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtTechnical 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   6675
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   390
         Width           =   1275
      End
      Begin VB.TextBox txtEvening 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2580
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   4275
         Width           =   1275
      End
      Begin VB.TextBox txtCharge 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   3765
         Width           =   1275
      End
      Begin VB.TextBox txtUtilities 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3285
         Width           =   1275
      End
      Begin VB.TextBox txtPFbank 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   2805
         Width           =   1275
      End
      Begin VB.TextBox txtMedical 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   2325
         Width           =   1275
      End
      Begin VB.TextBox txtEntertainment 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   1845
         Width           =   1275
      End
      Begin VB.TextBox txtConveyance 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
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
         Left            =   2595
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1365
         Width           =   1275
      End
      Begin VB.TextBox txtHouserent 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         Left            =   2595
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   885
         Width           =   1275
      End
      Begin VB.TextBox txtBasic 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """#.00"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2595
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cleaner Allowance Tk."
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
         Left            =   4680
         TabIndex        =   76
         Top             =   3870
         Width           =   1935
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment Tk."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4140
         TabIndex        =   74
         Top             =   4395
         Width           =   1725
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Water Suppler Allowance Tk."
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
         Left            =   4095
         TabIndex        =   52
         Top             =   3390
         Width           =   2520
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Night Guard Allowance Tk."
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
         Left            =   4305
         TabIndex        =   50
         Top             =   2910
         Width           =   2310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tea Boy Allowance Tk."
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
         Left            =   4620
         TabIndex        =   48
         Top             =   2385
         Width           =   1995
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Allowance Tk."
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
         Left            =   4860
         TabIndex        =   47
         Top             =   1860
         Width           =   1755
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Washing Allowance Tk."
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
         Left            =   4575
         TabIndex        =   46
         Top             =   1380
         Width           =   2040
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Guard Allowance Tk."
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
         Left            =   4035
         TabIndex        =   45
         Top             =   900
         Width           =   2580
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technical Allowance Tk."
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
         Left            =   4515
         TabIndex        =   44
         Top             =   420
         Width           =   2100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evening Allowance Tk."
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
         Left            =   585
         TabIndex        =   43
         Top             =   4305
         Width           =   1950
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Allowance Tk."
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
         Left            =   645
         TabIndex        =   42
         Top             =   3795
         Width           =   1890
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Residence Utilities Tk."
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
         Left            =   585
         TabIndex        =   41
         Top             =   3315
         Width           =   1950
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P/F Bank's Contribution Tk."
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
         TabIndex        =   40
         Top             =   2835
         Width           =   2385
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Allowance Tk."
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
         Left            =   585
         TabIndex        =   39
         Top             =   2355
         Width           =   1950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entertainment Allowance Tk."
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
         Left            =   45
         TabIndex        =   38
         Top             =   1875
         Width           =   2490
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conveyance Allowance Tk."
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
         Left            =   210
         TabIndex        =   37
         Top             =   1395
         Width           =   2325
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Rent Allowance Tk."
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
         Left            =   240
         TabIndex        =   36
         Top             =   915
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay Tk."
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
         Left            =   1305
         TabIndex        =   35
         Top             =   435
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0FFC0&
      Height          =   9060
      Left            =   8550
      ScaleHeight     =   9000
      ScaleWidth      =   5580
      TabIndex        =   11
      Top             =   810
      Width           =   5640
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmEntry.frx":0000
         Height          =   8115
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   14314
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12648384
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   5
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   5235
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   510
         Left            =   180
         Top             =   8460
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   900
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Payroll"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   9135
      TabIndex        =   6
      Top             =   10080
      Width           =   9195
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3300
         TabIndex        =   20
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   6300
         TabIndex        =   19
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   375
         Left            =   4815
         TabIndex        =   18
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Save"
         Height          =   375
         Left            =   315
         TabIndex        =   16
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   7785
         TabIndex        =   15
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   3285
         TabIndex        =   14
         Top             =   90
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0FFC0&
      Height          =   645
      Left            =   180
      ScaleHeight     =   585
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   10080
      Width           =   4560
      Begin VB.TextBox txtTakehome 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Tk.""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2340
         TabIndex        =   8
         Top             =   90
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Take Home Salary Tk."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   5
         Top             =   150
         Width           =   2100
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   945
      Left            =   180
      ScaleHeight     =   885
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   810
      Width           =   8175
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
         Left            =   2190
         TabIndex        =   77
         Top             =   375
         Width           =   3405
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
         Height          =   375
         Left            =   315
         TabIndex        =   7
         Top             =   375
         Width           =   1410
      End
      Begin VB.TextBox txtDesignation 
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
         Height          =   375
         Left            =   5985
         TabIndex        =   1
         Top             =   375
         Width           =   1950
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employs ID No."
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
         Left            =   360
         TabIndex        =   9
         Top             =   45
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of The Employee"
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
         Left            =   2250
         TabIndex        =   3
         Top             =   45
         Width           =   2025
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
         Left            =   5955
         TabIndex        =   2
         Top             =   45
         Width           =   1065
      End
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALARY AND ALLOWANCES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   195
      TabIndex        =   10
      Top             =   135
      Width           =   14040
   End
End
Attribute VB_Name = "frmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsU As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim str As String

Private Sub ColumnWidth()
        DataGrid1.Columns(0).Width = 300
        DataGrid1.Columns(1).Width = 2400
End Sub
Private Sub Sum()
        SumText1 = Val(txtBasic.Text): SumText2 = Val(txtHouserent.Text): SumText3 = Val(txtMedical.Text)
        SumText4 = Val(txtConveyance.Text): SumText5 = Val(txtEntertainment.Text): SumText6 = Val(txtPFbank.Text)
        SumText7 = Val(txtUtilities.Text): SumText8 = Val(txtCharge.Text): SumText9 = Val(txtEvening.Text)
        SumText10 = Val(txtTechnical.Text): SumText11 = Val(txtGuard.Text): SumText12 = Val(txtWashing.Text)
        SumText13 = Val(txtDriver.Text): SumText14 = Val(txtTeaboy.Text): sumText15 = Val(txtNightguard.Text)
        sumText16 = Val(txtCleaner.Text): sumText17 = Val(txtWater.Text)
        
        Payment = Val(SumText1 + SumText2 + SumText3 + SumText4 + SumText5 + SumText6 + SumText7 + SumText8 + SumText9 + SumText10 + SumText11 + SumText12 + SumText13 + SumText14 + sumText15 + sumText16 + sumText17)
        txtPayment.Text = Format$(Payment, "#0.00")
        
        sumText18 = Val(txtPFboth.Text): sumText19 = Val(txtIncometax.Text): sumText20 = Val(txtWelfare.Text)
        sumText21 = Val(txtShbl.Text): sumText22 = Val(txtDeathrisk.Text): sumText23 = Val(txtPfloan.Text)
        sumText24 = Val(txtFlexyloan.Text): sumText25 = Val(txtOther_Loan.Text)
        
        Deduction = Val(sumText18 + sumText19 + sumText20 + sumText21 + sumText22 + sumText23 + sumText24 + sumText25)
        txtDeduction.Text = Format$(Deduction, "#0.00")
        txtTakehome.Text = Val(Payment - Deduction)

End Sub


Private Sub clearTextboxes()
    txtID.Text = "": txtName.Text = "": txtDesignation.Text = ""
    txtBasic.Text = "0.00": txtHouserent.Text = "0.00": txtMedical.Text = "0.00": txtConveyance.Text = "0.00"
    txtEntertainment.Text = "0.00": txtPFbank.Text = "0.00": txtUtilities.Text = "0.00": txtCharge.Text = "0.00"
    txtEvening.Text = "0.00": txtTechnical.Text = "0.00": txtGuard.Text = "0.00": txtWashing.Text = "0.00"
    txtDriver.Text = "0.00": txtTeaboy.Text = "0.00": txtNightguard.Text = "0.00": txtCleaner.Text = "0.00"
    txtWater.Text = "0.00": txtPayment.Text = "0.00": txtPFboth.Text = "0.00": txtIncometax.Text = "0.00"
    txtWelfare.Text = "0.00": txtShbl.Text = "0.00": txtDeathrisk.Text = "0.00": txtPfloan.Text = "0.00"
    txtFlexyloan.Text = "0.00": txtOther_Loan.Text = "0.00": txtDeduction.Text = "0.00": txtTakehome.Text = "0.00"
End Sub
        
Private Sub cmdAdd_Click()
On Error Resume Next
    Set rsN = New ADODB.Recordset
        str = "select * from Salary where ID='" & txtID.Text & "'"
        rsN.Open str, conn
    If rsN.EOF Then
        rsN.Close
        rsN.Open "Salary", conn, adOpenDynamic, adLockOptimistic, -1
        rsN.AddNew
        rsN!ID = txtID.Text
        rsN!Name = txtName.Text
        rsN!Designation = txtDesignation.Text
        rsN!Basic_Pay = Val(txtBasic.Text)
        rsN!H_Rent = Val(txtHouserent.Text)
        rsN!Medical = Val(txtMedical.Text)
        rsN!Conveyance = Val(txtConveyance.Text)
        rsN!Entertainment = Val(txtEntertainment.Text)
        rsN!PF_Bank = Val(txtPFbank.Text)
        
        SumText1 = Val(txtBasic.Text): SumText2 = Val(txtHouserent.Text): SumText3 = Val(txtMedical.Text)
        SumText4 = Val(txtConveyance.Text): SumText5 = Val(txtEntertainment.Text): SumText6 = Val(txtPFbank.Text)
        Payment = Val(SumText1 + SumText2 + SumText3 + SumText4 + SumText5 + SumText6)
        rsN!Payment = Payment
        
        rsN!Utilities = Val(txtUtilities.Text)
        rsN!Charge = Val(txtCharge.Text)
        rsN!Evening = Val(txtEvening.Text)
        rsN!Technical = Val(txtTechnical.Text)
        rsN!Security = Val(txtGuard.Text)
        rsN!Washing = Val(txtWashing.Text)
        rsN!Driver = Val(txtDriver.Text)
        rsN!Teaboy = Val(txtTeaboy.Text)
        rsN!Nightguard = Val(txtNightguard.Text)
        rsN!Water_Supply = Val(txtWater.Text)
        rsN!Cleaner = Val(txtCleaner.Text)
        
        SumText7 = Val(txtUtilities.Text): SumText8 = Val(txtCharge.Text): SumText9 = Val(txtEvening.Text)
        SumText10 = Val(txtTechnical.Text): SumText11 = Val(txtGuard.Text): SumText12 = Val(txtWashing.Text)
        SumText13 = Val(txtDriver.Text): SumText14 = Val(txtTeaboy.Text): sumText15 = Val(txtNightguard.Text)
        sumText16 = Val(txtCleaner.Text): sumText17 = Val(txtWater.Text)
        Cash_Payment = Val(SumText7 + SumText8 + SumText9 + SumText10 + SumText11 + SumText12 + SumText13 + SumText14 + sumText15 + sumText16 + sumText17)
        rsN!Cash_Payment = Cash_Payment
        
        rsN!PF_Both = Val(txtPFboth.Text)
        rsN!Income_Tax = Val(txtIncometax.Text)
        rsN!Welfare = Val(txtWelfare.Text)
        rsN!SHBL = Val(txtShbl.Text)
        rsN!Death_Risk = Val(txtDeathrisk.Text)
        rsN!PF_Loan = Val(txtPfloan.Text)
        rsN!Flexy_Loan = Val(txtFlexyloan.Text)
        rsN!Other_Loan = Val(txtCcsloan.Text)
        
        sumText18 = Val(txtPFboth.Text): sumText19 = Val(txtIncometax.Text): sumText20 = Val(txtWelfare.Text)
        sumText21 = Val(txtShbl.Text): sumText22 = Val(txtDeathrisk.Text): sumText23 = Val(txtPfloan.Text)
        sumText24 = Val(txtFlexyloan.Text): sumText25 = Val(txtOther_Loan.Text)
        Deduction = Val(sumText18 + sumText19 + sumText20 + sumText21 + sumText22 + sumText23 + sumText24 + sumText25)
        rsN!Deduction = Deduction
        
        rsN!AC_Payment = Payment - Deduction
        rsN!Take_Home = (AC_Payment + Cash_Payment) - Deduction
        rsN.Update
        rsN.Close
    
    Call clearTextboxes

    Else
        MsgBox "There exist a record with this ID no: " & rsN!ID
        rsN.Close
    End If

    Set rs = New ADODB.Recordset
        str = "select * from Salary order by ID"
        rs.Open "Salary", conn, adOpenKeyset, adLockReadOnly, -1
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
    
    Call ColumnWidth
    txtID.SetFocus
    cmdAdd.Enabled = False
Exit Sub
  
   Resume Next
End Sub

Private Sub cmdDelete_Click()
    
    Set rs = New ADODB.Recordset
    str = "select * from Salary where ID like '" & txtID.Text & "'"
    rs.Open str, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    On Error Resume Next
        txtID.Text = rs!ID
        txtName.Text = rs!Name
        txtDesignation.Text = rs!Designation
        txtBasic.Text = rs![Basic Pay]
        txtHouserent.Text = rs![House Rent Allowance]
        txtMedical.Text = rs![Medical Allowance]
        txtConveyance.Text = rs![Conveyance Allowance]
        txtEntertainment.Text = rs![Entertainment Allowance]
        txtPFbank.Text = rs![P/F Bank Contribution]
        txtUtilities.Text = rs![Utilities Residence]
        txtTelephone.Text = rs![Telephone Allowance]
        txtCharge.Text = rs![Charge Allowance]
        txtEvening.Text = rs![Evening Allowance]
        txtTechnical.Text = rs![Technical Allowance]
        txtGuard.Text = rs![Security Guard]
        txtWashing.Text = rs![Washing Allowance]
        txtDriver.Text = rs![Driver Allowance]
        txtPayment.Text = rs![Total paid by the Bank]
        txtPFboth.Text = rs![P/F Both Contribution]
        txtIncometax.Text = rs![Income Tax]
        txtWelfare.Text = rs![Employees Welfare Fund]
        txtShbl.Text = rs![Staff House Building Loan]
        txtDeathrisk.Text = rs![Death Risk Coverage Scheme]
        txtPfloan.Text = rs![P/F loan]
        txtCcsloan.Text = rs![Loan Under CCS]
        txtFlexyloan.Text = rs![Flexy Loan]
        txtDeduction.Text = rs![Total Deduction]
        txtTakehome.Text = rs![Net Payable]
        rs.Close
     
     If MsgBox("Really want to delete?", vbCritical + vbYesNo) = vbYes Then
        
        str = "delete from Salary where ID like '" & txtID.Text & "'"
        rs.Open str, conn, adOpenDynamic, adLockOptimistic
        rs.Close
    Call clearTextboxes
        str = "select * from Salary order by ID"
        rs.Open str, conn, adOpenKeyset, adLockReadOnly
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
    Call ColumnWidth
        End If
    Else
        MsgBox "There is no personal information with this ID: " & txtID.Text
        rs.Close
    End If
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdEdit_Click()
Dim ID As String
    ID = txtID.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Salary where ID like '" & ID & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        txtID.Text = rs!ID
        txtName.Text = rs!Name
        txtDesignation.Text = rs!Designation
        txtBasic.Text = Format$(rs![Basic Pay], "#0.00")
        txtHouserent.Text = Format$(rs![House Rent Allowance], "#0.00")
        txtMedical.Text = Format$(rs![Medical Allowance], "#0.00")
        txtConveyance.Text = Format$(rs![Conveyance Allowance], "#0.00")
        txtEntertainment.Text = Format$(rs![Entertainment Allowance], "#0.00")
        txtPFbank.Text = Format$(rs![P/F Bank Contribution], "#0.00")
        txtUtilities.Text = Format$(rs![Utilities Residence], "#0.00")
        txtTelephone.Text = Format$(rs![Telephone Allowance], "#0.00")
        txtCharge.Text = Format$(rs![Charge Allowance], "#0.00")
        txtEvening.Text = Format$(rs![Evening Allowance], "#0.00")
        txtTechnical.Text = Format$(rs![Technical Allowance], "#0.00")
        txtGuard.Text = Format$(rs![Security Guard], "#0.00")
        txtWashing.Text = Format$(rs![Washing Allowance], "#0.00")
        txtDriver.Text = Format$(rs![Driver Allowance], "#0.00")
        txtPayment.Text = Format$(rs![Total paid by the Bank], "#0.00")
        txtPFboth.Text = Format$(rs![P/F Both Contribution], "#0.00")
        txtIncometax.Text = Format$(rs![Income Tax], "#0.00")
        txtWelfare.Text = Format$(rs![Employees Welfare Fund], "#0.00")
        txtShbl.Text = Format$(rs![Staff House Building Loan], "#0.00")
        txtDeathrisk.Text = Format$(rs![Death Risk Coverage Scheme], "#0.00")
        txtPfloan.Text = Format$(rs![P/F loan], "#0.00")
        txtCcsloan.Text = Format$(rs![Loan Under CCS], "#0.00")
        txtFlexyloan.Text = Format$(rs![Flexy Loan], "#0.00")
        txtDeduction.Text = Format$(rs![Total Deduction], "#0.00")
        txtTakehome.Text = Format$(rs![Net Payable], "#0.00")
        rs.Close
        txtBasic.SetFocus
    End If
        cmdEdit.Visible = False
        cmdUpdate.Visible = True
        
    Exit Sub
    
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub cmdPrint_Click()
Unload Me
Shell App.Path & "\Report1"
End Sub

Private Sub cmdUpdate_Click()
    
    On Error GoTo Last
 
    Set rsU = New ADODB.Recordset
    str = "select * from Salary where ID like '" & txtID.Text & "'"
    rsU.Open str, conn, adOpenDynamic, adLockOptimistic
    If Not rsU.EOF Then
        rsU!ID = txtID.Text
        rsU!Name = txtName.Text
        rsU!Designation = txtDesignation.Text
        rsU![Basic Pay] = Format$(Val(txtBasic.Text), "#0.00")
        rsU![House Rent Allowance] = Val(txtHouserent.Text)
        rsU![Medical Allowance] = Val(txtMedical.Text)
        rsU![Conveyance Allowance] = Val(txtConveyance.Text)
        rsU![Entertainment Allowance] = Val(txtEntertainment.Text)
        rsU![P/F Bank Contribution] = Val(txtPFbank.Text)
        rsU![Utilities Residence] = Val(txtUtilities.Text)
        rsU![Telephone Allowance] = Val(txtTelephone.Text)
        rsU![Charge Allowance] = Val(txtCharge.Text)
        rsU![Evening Allowance] = Val(txtEvening.Text)
        rsU![Technical Allowance] = Val(txtTechnical.Text)
        rsU![Security Guard] = Val(txtGuard.Text)
        rsU![Washing Allowance] = Val(txtWashing.Text)
        rsU![Driver Allowance] = Val(txtDriver.Text)
        rsU![Total paid by the Bank] = Val(txtPayment.Text)
        rsU![P/F Both Contribution] = Val(txtPFboth.Text)
        rsU![Income Tax] = Val(txtIncometax.Text)
        rsU![Employees Welfare Fund] = Val(txtWelfare.Text)
        rsU![Staff House Building Loan] = Val(txtShbl.Text)
        rsU![Death Risk Coverage Scheme] = Val(txtDeathrisk.Text)
        rsU![P/F loan] = Val(txtPfloan.Text)
        rsU![Loan Under CCS] = Val(txtCcsloan.Text)
        rsU![Flexy Loan] = Val(txtFlexyloan.Text)
        rsU![Total Deduction] = Val(txtDeduction.Text)
        rsU![Net Payable] = Val(txtTakehome.Text)
        rsU.Update
        rsU.Close
        cmdUpdate.Visible = False
        cmdEdit.Visible = True
        cmdEdit.Enabled = False
    Else
        MsgBox "There is no such ID No. to update.", 64, "Update Error"
        rsU.Close
        Exit Sub
    End If
    
    
    Set rs = New ADODB.Recordset
        rs.Open "Salary", conn, adOpenKeyset, adLockReadOnly, -1
        Adodc1.RecordSource = "select *  from Salary order by ID"
        Adodc1.Refresh
        DataGrid1.Refresh
    Call ColumnWidth
    Call clearTextboxes
    Exit Sub
Last:
   MsgBox "Error: " + Err.Description
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.SelText
End Sub
Private Sub Form_Load()
    On Error GoTo Last
    
    Set rs = New ADODB.Recordset
        str = "select * from Salary"
        rs.Open str, conn
        
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        rs.Close
        Call clearTextboxes
        Call ColumnWidth
        cmdAdd.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdUpdate.Visible = False
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
If Text1.Text = "" Then
Exit Sub
End If

    Dim s As String
        s = Text1.Text
    
    Set rs = New ADODB.Recordset
        str = "select * from Salary where Name like '" & s & "'"
        rs.Open str, conn
    
    If Not rs.EOF Then
    On Error Resume Next
        Adodc1.ConnectionString = cnStr
        Adodc1.RecordSource = str
        Adodc1.Refresh
        DataGrid1.Refresh
        rs.Close
        Else
    Exit Sub
    End If
    End If
End Sub

Private Sub txtCcsloan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFlexyloan.SetFocus
End If
End Sub

Private Sub txtCleaner_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtPFboth.SetFocus
End If
End Sub

Private Sub txtEvening_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtTechnical.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    txtName.SetFocus
End If
End Sub



Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDesignation.SelStart = 0
    txtDesignation.SelLength = Len(txtDesignation.Text)
    txtDesignation.SetFocus
End If
End Sub
Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBasic.SelStart = 0
    txtBasic.SelLength = Len(txtBasic.Text)
    txtBasic.SetFocus
End If
End Sub
Private Sub txtBasic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'txtBasic.Text = Format$(Val(txtBasic.Text), "###,#0.00")
'txtHouserent.Text = Format$(Val(txtBasic.Text) * (50 / 100), "###,#0.00")
'txtMedical.Text = Format$(Val(txtBasic.Text) * (10 / 100), "###,#0.00")
'txtPFbank.Text = Format$(Val(txtBasic.Text) * (10 / 100), "###,#0.00")
'txtPFboth.Text = Format$(Val(txtBasic.Text) * (10 / 100) * 2, "###,#0.00")
    txtHouserent.SelStart = 0
    txtHouserent.SelLength = Len(txtHouserent.Text)
    txtHouserent.SetFocus
Call Sum
    
End If
End Sub
Private Sub txtHouserent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtConveyance.SelStart = 0
    txtConveyance.SelLength = Len(txtConveyance.Text)
    txtConveyance.SetFocus
    Call Sum
End If
End Sub
Private Sub txtConveyance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEntertainment.SelStart = 0
    txtEntertainment.SelLength = Len(txtEntertainment.Text)
    txtEntertainment.SetFocus
Call Sum
End If
End Sub
Private Sub txtEntertainment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedical.SelStart = 0
    txtMedical.SelLength = Len(txtMedical.Text)
    txtMedical.SetFocus
Call Sum
End If
End Sub
Private Sub txtCharge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtEvening.SetFocus
End If
End Sub
Private Sub txtDeathrisk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtShbl.SetFocus
End If
End Sub
Private Sub txtDriver_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtTeaboy.SetFocus
End If
End Sub
Private Sub txtFlexyloan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtOther_Loan.SetFocus
End If
End Sub
Private Sub txtGuard_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtWashing.SetFocus
End If
End Sub
Private Sub txtIncometax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtWelfare.SetFocus
End If
End Sub
Private Sub txtMedical_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPFbank.SelStart = 0
    txtPFbank.SelLength = Len(txtPFbank.Text)
    txtPFbank.SetFocus
Call Sum
End If
End Sub

Private Sub txtNightguard_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtWater.SetFocus
End If
End Sub

Private Sub txtOther_Loan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtID.SetFocus
End If
End Sub

Private Sub txtPFbank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPFboth.Text = Format$(Val(txtPFbank.Text) * 2, "#0.00")
Call Sum
txtUtilities.SetFocus
End If
End Sub
Private Sub txtPFboth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtIncometax.SetFocus
End If
End Sub
Private Sub txtPfloan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtFlexyloan.SetFocus
End If
End Sub
Private Sub txtShbl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtPfloan.SetFocus
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
        txtDesignation.Text = rs!Designation
        txtBasic.Text = Format$(rs!Basic_Pay, "###,###0.00")
        txtHouserent.Text = Format$(rs!H_Rent, "###,###0.00")
        txtMedical.Text = Format$(rs!Medical, "###,###0.00")
        txtConveyance.Text = Format$(rs!Conveyance, "###,###0.00")
        txtEntertainment.Text = Format$(rs!Entertainment, "###,###0.00")
        txtPFbank.Text = Format$(rs!PF_Bank, "###,###0.00")
        txtUtilities.Text = Format$(rs!Utilities, "###,###0.00")
        txtCharge.Text = Format$(rs!Charge, "###,###0.00")
        txtEvening.Text = Format$(rs!Evening, "###,###0.00")
        txtTechnical.Text = Format$(rs!Technical, "###,###0.00")
        txtGuard.Text = Format$(rs!Security, "###,###0.00")
        txtWashing.Text = Format$(rs!Washing, "###,###0.00")
        txtDriver.Text = Format$(rs!Driver, "###,###0.00")
        txtTeaboy.Text = Format$(rs!Teaboy, "###,###0.00")
        txtNightguard.Text = Format$(rs!Nightguard, "###,###0.00")
        txtWater.Text = Format$(rs!Water_Supply, "###,###0.00")
        txtCleaner.Text = Format$(rs!Cleaner, "###,###0.00")
        
        
        
        Sum1 = Val(rs!Basic_Pay): Sum2 = Val(rs!H_Rent): Sum3 = Val(rs!Medical): Sum4 = Val(rs!Conveyance)
        Sum5 = Val(rs!Entertainment): Sum6 = Val(rs!PF_Bank): Sum7 = Val(rs!Utilities): Sum8 = Val(rs!Charge)
        Sum9 = Val(rs!Evening): Sum10 = Val(rs!Technical): Sum11 = Val(rs!Security): Sum12 = Val(rs!Washing)
        Sum13 = Val(rs!Driver): Sum14 = Val(rs!Teaboy): Sum15 = Val(rs!Nightguard): Sum16 = Val(rs!Water_Supply): Sum17 = Val(rs!Cleaner)
        
        Payment = Sum1 + Sum2 + Sum3 + Sum4 + Sum5 + Sum6 + Sum7 + Sum8 + Sum9 + Sum10 + Sum11 + Sum12 + Sum13 + Sum14 + Sum15 + Sum16 + Sum17
        txtPayment.Text = Format$(Payment, "###,###0.00")
        
        txtPFboth.Text = Format$(rs!PF_Both, "###,###0.00")
        txtIncometax.Text = Format$(rs!Income_Tax, "###,###0.00")
        txtWelfare.Text = Format$(rs!Welfare, "###,###0.00")
        txtShbl.Text = Format$(rs!SHBL, "###,###0.00")
        txtDeathrisk.Text = Format$(rs!Death_Risk, "###,###0.00")
        txtPfloan.Text = Format$(rs!PF_Loan, "###,###0.00")
        txtCcsloan.Text = Format$(rs!Other_Loan, "###,###0.00")
        txtFlexyloan.Text = Format$(rs!Flexy_Loan, "###,###0.00")
       
        Sum18 = Val(rs!PF_Both): Sum19 = Val(rs!Income_Tax): Sum20 = Val(rs!Welfare): Sum21 = Val(rs!SHBL)
        Sum22 = Val(rs!Death_Risk): Sum23 = Val(rs!PF_Loan): Sum24 = Val(rs!Other_Loan): Sum25 = Val(rs!Flexy_Loan)
        
        Deduction = Sum18 + Sum19 + Sum20 + Sum21 + Sum22 + Sum23 + Sum24 + Sum25
        txtDeduction.Text = Format$(Deduction, "###,###0.00")
        
        txtTakehome.Text = Format$((Val(Payment - Deduction)), "###,###0.00")
        
        rs.Close
        cmdUpdate.Visible = False
        cmdEdit.Visible = True
        cmdEdit.Enabled = True
        cmdAdd.Enabled = False
        cmdDelete.Enabled = True
        cmdPrint.Enabled = True
    Else
    If MsgBox("There is no such ID found, Do you want add new employee?", vbCritical + vbYesNo) = vbYes Then
        Call clearTextboxes
        txtID.Text = ID
        cmdAdd.Enabled = True
        rs.Close
    Else
        Call clearTextboxes
    End If
    End If
    Exit Sub
Last:
    MsgBox ("Database Connection error: " + Err.Description)
End Sub

Private Sub txtTeaboy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtNightguard.SetFocus
End If
End Sub

Private Sub txtTechnical_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtGuard.SetFocus
End If
End Sub
Private Sub txtUtilities_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtCharge.SetFocus
End If
End Sub
Private Sub txtWashing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtDriver.SetFocus
End If
End Sub

Private Sub txtWater_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtCleaner.SetFocus
End If
End Sub

Private Sub txtWelfare_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Sum
txtDeathrisk.SetFocus
End If
End Sub
