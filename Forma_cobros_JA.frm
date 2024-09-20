VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form Forma_principal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Payments verification"
   ClientHeight    =   15180
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   28560
   Icon            =   "Forma_cobros_JA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   15180
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   240
      TabIndex        =   183
      Top             =   13080
      Visible         =   0   'False
      Width           =   21135
      _ExtentX        =   37280
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   840
      TabIndex        =   177
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton op_fecha_carga 
         BackColor       =   &H00808080&
         Caption         =   "20 days before"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   182
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton op_fecha_carga 
         BackColor       =   &H00808080&
         Caption         =   "Only month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   179
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton op_fecha_carga 
         BackColor       =   &H00808080&
         Caption         =   "10 days before"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   178
         Top             =   360
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No save"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   176
      Top             =   2160
      Value           =   1  'Checked
      Width           =   700
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   2400
      TabIndex        =   175
      Top             =   10200
      Visible         =   0   'False
      Width           =   4815
      _Version        =   524288
      _ExtentX        =   8493
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483638
      Year            =   2021
      Month           =   5
      Day             =   7
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1680
      TabIndex        =   168
      Top             =   9140
      Width           =   1695
      Begin VB.TextBox txtfecha_cargada 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   174
         Top             =   680
         Width           =   855
      End
      Begin VB.TextBox txtfecha_cargada 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   173
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton op_rango_mes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Partial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   170
         Top             =   60
         Width           =   855
      End
      Begin VB.OptionButton op_rango_mes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   169
         Top             =   60
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   172
         Top             =   680
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   171
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   8040
      ScaleHeight     =   1185
      ScaleWidth      =   8625
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Label lblmsg2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   8400
      End
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading data... "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   8415
      End
   End
   Begin Project1.lvButtons_H btnsort4 
      Height          =   495
      Left            =   17080
      TabIndex        =   142
      ToolTipText     =   "Sort all data"
      Top             =   800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3336E
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   3015
      Left            =   10080
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   4815
      _Version        =   524288
      _ExtentX        =   8493
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   12648447
      Year            =   2021
      Month           =   5
      Day             =   7
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      Height          =   6615
      Left            =   12840
      ScaleHeight     =   6555
      ScaleWidth      =   8475
      TabIndex        =   123
      Top             =   2280
      Visible         =   0   'False
      Width           =   8535
      Begin VB.ListBox List3 
         Height          =   1425
         Left            =   7680
         TabIndex        =   180
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin Project1.lvButtons_H btnsepara_cantidades2 
         Height          =   855
         Left            =   2280
         TabIndex        =   161
         ToolTipText     =   "Separate the currency sign from the amounts"
         Top             =   5280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":33A8D
         ImgSize         =   40
         cBack           =   12632256
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove ""$"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   162
         Top             =   6160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin Project1.lvButtons_H btnsepara_cantidades 
         Height          =   855
         Left            =   120
         TabIndex        =   155
         ToolTipText     =   "Separate amounts  in a PDF file"
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1296
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":34C75
         ImgSize         =   40
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btn_borra_RichTextBox1 
         Height          =   495
         Left            =   3600
         TabIndex        =   154
         Top             =   4080
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   873
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":36124
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnajuste 
         Height          =   1215
         Left            =   6720
         TabIndex        =   141
         ToolTipText     =   "Remove unchecked options"
         Top             =   4320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   2143
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Forma_cobros_JA.frx":36A86
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5100
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   143
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin Project1.lvButtons_H btnctrl_v 
         Height          =   525
         Left            =   3600
         TabIndex        =   153
         Top             =   4680
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         Caption         =   "Paste"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   40
         cBack           =   14737632
      End
      Begin Project1.lvButtons_H btnguardar 
         Height          =   525
         Left            =   4680
         TabIndex        =   149
         Top             =   5640
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   40
         cBack           =   14737632
      End
      Begin Project1.lvButtons_H btnleft 
         Height          =   615
         Left            =   3840
         TabIndex        =   151
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1085
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":379E5
         ImgSize         =   48
         cBack           =   8438015
      End
      Begin VB.OptionButton op_day 
         BackColor       =   &H0080C0FF&
         Caption         =   "4 days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   145
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton op_day 
         BackColor       =   &H0080C0FF&
         Caption         =   "3 days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   144
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton op_day 
         BackColor       =   &H0080C0FF&
         Caption         =   "2 days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   138
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton op_day 
         BackColor       =   &H0080C0FF&
         Caption         =   "1 day"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   137
         Top             =   720
         Value           =   -1  'True
         Width           =   735
      End
      Begin Project1.lvButtons_H btnmerge 
         Height          =   525
         Left            =   5400
         TabIndex        =   150
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         Caption         =   "Merge selected rows"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   40
         Enabled         =   0   'False
         cBack           =   14737632
      End
      Begin Project1.lvButtons_H btnright 
         Height          =   615
         Left            =   4440
         TabIndex        =   152
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1085
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":37E37
         ImgSize         =   48
         cBack           =   8438015
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4335
         Left            =   120
         TabIndex        =   156
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7646
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Forma_cobros_JA.frx":38289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Search for the missing amount(s) and once found, click the ""MERGE"" button to merge the saved amounts with the newly found ones."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Index           =   2
         Left            =   6600
         TabIndex        =   165
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Once you have selected the quantities, Clic on the button below to eliminate the non-selected quantities."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   975
         Index           =   0
         Left            =   6600
         TabIndex        =   164
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "To merge another quantity with select quantities from the list..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   6600
         TabIndex        =   163
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label lblfecha2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm/dd/yyyy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3960
         TabIndex        =   160
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label lblfecha1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm/dd/yyyy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3960
         TabIndex        =   159
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3960
         TabIndex        =   158
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3960
         TabIndex        =   157
         Top             =   840
         Width           =   420
      End
      Begin VB.Image pulgar 
         Height          =   780
         Left            =   7200
         Picture         =   "Forma_cobros_JA.frx":38304
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbltotal_lista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6000
         TabIndex        =   147
         Top             =   5160
         Width           =   90
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   146
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label lbl_total_marcado 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   255
         Left            =   6840
         TabIndex        =   140
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total marked"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   6795
         TabIndex        =   139
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Looking for        Difference"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Left            =   1200
         TabIndex        =   128
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label lbl_diferencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2160
         TabIndex        =   127
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lbltotal_needed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1140
         TabIndex        =   126
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lbltotal_cantidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   45
         TabIndex        =   125
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   124
         Top             =   120
         Width           =   960
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   975
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   6615
         Left            =   6480
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid6 
      Height          =   4215
      Left            =   21480
      TabIndex        =   120
      Top             =   7080
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7435
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox chk_exacto3 
      BackColor       =   &H8000000C&
      Caption         =   "Exact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15480
      TabIndex        =   103
      Top             =   4180
      Value           =   1  'Checked
      Width           =   615
   End
   Begin Project1.lvButtons_H chk_multiple 
      Height          =   615
      Left            =   4200
      TabIndex        =   110
      Top             =   3480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   "Multiple receipts"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   33023
   End
   Begin Project1.lvButtons_H btnsort3 
      Height          =   495
      Left            =   5040
      TabIndex        =   136
      ToolTipText     =   "Sort all data"
      Top             =   3600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":38A3D
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cbosort3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   135
      Top             =   3720
      Width           =   975
   End
   Begin Project1.lvButtons_H btnsort1 
      Height          =   495
      Left            =   12960
      TabIndex        =   134
      ToolTipText     =   "Sort all data"
      Top             =   10080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3915C
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cbosort1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      Style           =   2  'Dropdown List
      TabIndex        =   133
      Top             =   10240
      Width           =   975
   End
   Begin VB.ComboBox cbo_oficina 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   19560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   131
      Top             =   720
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   16560
      Top             =   8040
   End
   Begin VB.PictureBox encapsula 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3240
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   129
      Top             =   2560
      Visible         =   0   'False
      Width           =   855
   End
   Begin Project1.lvButtons_H btnexcel 
      Height          =   615
      Left            =   18600
      TabIndex        =   9
      ToolTipText     =   "Export to Excel"
      Top             =   10680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3987B
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btntransfiere 
      Height          =   615
      Left            =   18600
      TabIndex        =   39
      ToolTipText     =   "Double verification"
      Top             =   11400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3A62E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnexcel3 
      Height          =   615
      Left            =   20640
      TabIndex        =   93
      ToolTipText     =   "Export to Excel"
      Top             =   5280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3B10E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cboprogram 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   18840
      Style           =   2  'Dropdown List
      TabIndex        =   118
      Top             =   280
      Width           =   2295
   End
   Begin VB.TextBox txtcust_id 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17760
      TabIndex        =   115
      Top             =   120
      Width           =   615
   End
   Begin Project1.lvButtons_H btnborra_cant 
      Height          =   375
      Left            =   6000
      TabIndex        =   76
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":3BFB6
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   22080
      Sorted          =   -1  'True
      TabIndex        =   109
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Project1.lvButtons_H op_rem 
      Height          =   495
      Index           =   0
      Left            =   7680
      TabIndex        =   107
      Top             =   8520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":3C918
      ImgSize         =   40
      cBack           =   -2147483645
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      TabIndex        =   98
      Top             =   4320
      Width           =   3375
      Begin Project1.lvButtons_H btnbusca3 
         Height          =   495
         Left            =   2880
         TabIndex        =   99
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":3D630
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.OptionButton op_busca3 
         BackColor       =   &H8000000C&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton op_busca3 
         BackColor       =   &H8000000C&
         Caption         =   "Policy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   101
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcantidad3 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   100
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   1935
      Left            =   21960
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
      _Version        =   393216
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid5 
      Height          =   2415
      Left            =   12960
      TabIndex        =   89
      Top             =   5160
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox chk_exacto2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   92
      Top             =   3000
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox chk_exacto 
      BackColor       =   &H00808080&
      Caption         =   "Exact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   91
      Top             =   9480
      Value           =   1  'Checked
      Width           =   615
   End
   Begin Project1.lvButtons_H btntransfiere_grid4_a_grid3 
      Height          =   975
      Left            =   3160
      TabIndex        =   61
      ToolTipText     =   "Fill up data in row"
      Top             =   2980
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":3E160
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.ComboBox lblfila_grid4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3360
      TabIndex        =   88
      Top             =   2760
      Width           =   560
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   83
      Top             =   3120
      Width           =   2535
      Begin Project1.lvButtons_H btnbusca2 
         Height          =   495
         Left            =   2040
         TabIndex        =   87
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":3EA8B
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtcantidad 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton op_busca 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Policy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   880
         TabIndex        =   85
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton op_busca 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   4800
      TabIndex        =   77
      Top             =   9600
      Width           =   2535
      Begin Project1.lvButtons_H btnsearch2 
         Height          =   495
         Left            =   2040
         TabIndex        =   82
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_cobros_JA.frx":3F5BB
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtcantidad2 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   81
         Top             =   520
         Width           =   1935
      End
      Begin VB.OptionButton op_busca2 
         BackColor       =   &H00808080&
         Caption         =   "VOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1540
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton op_busca2 
         BackColor       =   &H00808080&
         Caption         =   "Policy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   880
         TabIndex        =   79
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton op_busca2 
         BackColor       =   &H00808080&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   30
      Left            =   3120
      TabIndex        =   69
      Top             =   5640
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btneraser_all 
      Height          =   495
      Left            =   2760
      TabIndex        =   68
      ToolTipText     =   "Clean all"
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":400EB
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnlimbiacbo 
      Height          =   375
      Left            =   16560
      TabIndex        =   66
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":40E73
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   3360
      TabIndex        =   63
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   5
      SelectRange     =   -1  'True
      SelStart        =   5
      Value           =   5
   End
   Begin VB.CheckBox chk_exact 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   62
      Top             =   340
      Value           =   1  'Checked
      Width           =   615
   End
   Begin Project1.lvButtons_H btnbusca_registro 
      Height          =   855
      Left            =   3360
      TabIndex        =   46
      ToolTipText     =   "Find payment information"
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":417D5
      ImgSize         =   48
      cBack           =   49344
   End
   Begin VB.CheckBox chkactive 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   17640
      TabIndex        =   60
      Top             =   720
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtfecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   12320
      TabIndex        =   59
      Top             =   80
      Width           =   920
   End
   Begin VB.ComboBox cbocompany 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   14320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtfecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   11160
      TabIndex        =   53
      Top             =   80
      Width           =   920
   End
   Begin VB.TextBox txtpoliza 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7120
      TabIndex        =   51
      Top             =   80
      Width           =   2610
   End
   Begin VB.TextBox txtamount 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5200
      TabIndex        =   49
      Top             =   80
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid4 
      Height          =   1335
      Left            =   4080
      TabIndex        =   45
      Top             =   720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   14737632
      BackColorFixed  =   12632256
      BackColorBkg    =   12632256
      GridColor       =   -2147483648
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btnexcel2 
      Height          =   615
      Left            =   11760
      TabIndex        =   34
      ToolTipText     =   "Export to Excel"
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":4226F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnremove_note 
      Height          =   375
      Left            =   9600
      TabIndex        =   31
      Top             =   7920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "&Remove"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnagregar_nota 
      Height          =   375
      Left            =   9000
      TabIndex        =   30
      Top             =   7920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.TextBox txtnota 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   8280
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   12
      Top             =   8280
      Width           =   3855
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   8
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "September"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnload1 
         Height          =   615
         Left            =   3120
         TabIndex        =   90
         ToolTipText     =   "Load all data from LAE System"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Forma_cobros_JA.frx":43049
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   10
         Left            =   2520
         TabIndex        =   23
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "November"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   22
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "October"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.ComboBox cboyear 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "January"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "February"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "March"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   16
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "April"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "May"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   18
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "June"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "July"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "August"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   11
         Left            =   3120
         TabIndex        =   24
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "December"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin Project1.lvButtons_H btnverifica_polizas 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Check payments"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":43C6B
      ImgSize         =   48
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnsepara_polizas 
      Height          =   495
      Left            =   22920
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "Separa"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   735
      Left            =   20640
      TabIndex        =   4
      Top             =   12360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "&End"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3360
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   10560
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   3625
      _Version        =   393216
      ForeColor       =   0
      Rows            =   1
      FixedRows       =   0
      BackColorFixed  =   32768
      ForeColorFixed  =   16777215
      BackColorBkg    =   14737632
      ScrollTrack     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btncarga_excel 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Load information from Excel"
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":44CC1
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnsql 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Transfer data to LAE System"
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":4536F
      ImgSize         =   48
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid3 
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      BackColorFixed  =   49152
      ForeColorFixed  =   16777215
      BackColorBkg    =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btnrevisado 
      Height          =   615
      Left            =   11760
      TabIndex        =   40
      ToolTipText     =   "Transfer verified data"
      Top             =   4920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":45F7F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnload2 
      Height          =   615
      Left            =   22920
      TabIndex        =   41
      ToolTipText     =   "Import from Excel"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":467A9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnlimpia_grid2 
      Height          =   615
      Left            =   22320
      TabIndex        =   42
      ToolTipText     =   "Clean the grid"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":4728B
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnlimpia_grid1 
      Height          =   615
      Left            =   22320
      TabIndex        =   43
      ToolTipText     =   "Clean the grid"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":47DF7
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtfila 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3440
      MaxLength       =   5
      TabIndex        =   47
      Top             =   3800
      Width           =   435
   End
   Begin Project1.lvButtons_H btnlimpiapoliza 
      Height          =   375
      Left            =   9720
      TabIndex        =   67
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":48954
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtfirst_name 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   17640
      TabIndex        =   71
      Top             =   1320
      Width           =   1095
   End
   Begin Project1.lvButtons_H btnlimpia_cte 
      Height          =   375
      Left            =   20860
      TabIndex        =   72
      Top             =   1320
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":492B6
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtlast_name 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   18840
      TabIndex        =   73
      Top             =   1320
      Width           =   1935
   End
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   615
      Left            =   22320
      TabIndex        =   94
      ToolTipText     =   "Clean the grid"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":49C18
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H op_grid 
      Height          =   615
      Index           =   0
      Left            =   16080
      TabIndex        =   95
      Top             =   10020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Verified"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   2
      Value           =   0   'False
      cBack           =   49152
   End
   Begin Project1.lvButtons_H op_grid 
      Height          =   495
      Index           =   1
      Left            =   9360
      TabIndex        =   96
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "To be checked"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   2
      Value           =   -1  'True
      cBack           =   49152
   End
   Begin Project1.lvButtons_H op_grid 
      Height          =   495
      Index           =   2
      Left            =   17880
      TabIndex        =   97
      Top             =   4720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "To be checked"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   2
      Value           =   0   'False
      cBack           =   49152
   End
   Begin Project1.lvButtons_H op_rem 
      Height          =   495
      Index           =   1
      Left            =   7680
      TabIndex        =   108
      Top             =   9000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":4A7FD
      ImgSize         =   40
      cBack           =   -2147483645
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000C&
      Height          =   975
      Left            =   13200
      TabIndex        =   111
      Top             =   7320
      Width           =   2775
      Begin VB.OptionButton op_mes 
         BackColor       =   &H8000000C&
         Caption         =   "Show last days and full month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   113
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton op_mes 
         BackColor       =   &H8000000C&
         Caption         =   "show only the month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   1935
      End
   End
   Begin Project1.lvButtons_H btnborra_cust_id 
      Height          =   375
      Left            =   18360
      TabIndex        =   116
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":4B517
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnerase_program 
      Height          =   375
      Left            =   21200
      TabIndex        =   119
      Top             =   240
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":4BE79
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btndepura 
      Height          =   615
      Left            =   11760
      TabIndex        =   122
      ToolTipText     =   "Transfer verified data"
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "depura"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnborra_oficinas 
      Height          =   375
      Left            =   21200
      TabIndex        =   132
      Top             =   720
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_cobros_JA.frx":4C7DB
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H BTNELIMINA_FILA 
      Height          =   495
      Left            =   17080
      TabIndex        =   166
      ToolTipText     =   "Delete a row"
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":4D13D
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnborra_fila 
      Height          =   615
      Left            =   11760
      TabIndex        =   167
      ToolTipText     =   "remove rows"
      Top             =   5640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Forma_cobros_JA.frx":4D7E8
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Label lblarchivo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   960
      TabIndex        =   181
      Top             =   120
      Width           =   90
   End
   Begin VB.Image Image6 
      Height          =   2655
      Left            =   120
      Picture         =   "Forma_cobros_JA.frx":4E173
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   4455
   End
   Begin VB.Label lblamount_total 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   148
      Top             =   2140
      Width           =   150
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Office:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   18840
      TabIndex        =   130
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April, 2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   18600
      TabIndex        =   121
      Top             =   12480
      Width           =   675
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   -2400
      Top             =   13560
      Width           =   26775
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   18840
      TabIndex        =   117
      Top             =   0
      Width           =   810
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   80
      Left            =   2640
      Top             =   0
      Width           =   18975
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   17040
      Top             =   620
      Width           =   4575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cust.ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   17040
      TabIndex        =   114
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To double check, simply click on the VERIFIED column of the payments you want to verify and then click on the double check button"
      ForeColor       =   &H00808080&
      Height          =   1815
      Left            =   19560
      TabIndex        =   106
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label lbltotal5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   18960
      TabIndex        =   105
      Top             =   7660
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   18480
      TabIndex        =   104
      Top             =   7680
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   3
      Left            =   20280
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image Image5 
      Height          =   660
      Left            =   19440
      Picture         =   "Forma_cobros_JA.frx":4ED68
      Stretch         =   -1  'True
      Top             =   4600
      Width           =   660
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   2
      Left            =   18120
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2295
      Index           =   1
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   940
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   18960
      TabIndex        =   75
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   17760
      TabIndex        =   74
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   17640
      TabIndex        =   70
      Top             =   1035
      Width           =   975
   End
   Begin VB.Label lbltotal4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      TabIndex        =   65
      Top             =   2120
      Width           =   165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:              Total Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   5520
      TabIndex        =   64
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4440
      TabIndex        =   48
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   12100
      TabIndex        =   58
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   57
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   13440
      TabIndex        =   55
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10200
      TabIndex        =   52
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v3.42"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   18645
      TabIndex        =   44
      Top             =   12195
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   0
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   17520
      Picture         =   "Forma_cobros_JA.frx":4F449
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   10920
      Picture         =   "Forma_cobros_JA.frx":4FADF
      Stretch         =   -1  'True
      Top             =   3520
      Width           =   615
   End
   Begin VB.Label lbltotal2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   885
      TabIndex        =   38
      Top             =   12640
      Width           =   195
   End
   Begin VB.Label lbltotal1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   885
      TabIndex        =   37
      Top             =   7600
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   36
      Top             =   7680
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   35
      Top             =   12720
      Width           =   510
   End
   Begin VB.Label lblrow 
      BackStyle       =   0  'Transparent
      Caption         =   "......"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   33
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Row#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   32
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   28
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Hector Navarro"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   18720
      TabIndex        =   27
      Top             =   13200
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   10320
      Picture         =   "Forma_cobros_JA.frx":50198
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   19320
      Picture         =   "Forma_cobros_JA.frx":50CC7
      Stretch         =   -1  'True
      Top             =   10680
      Width           =   2265
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   18975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   17060
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4560
   End
   Begin VB.Shape Shape11 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   0
      Top             =   8760
      Width           =   22575
   End
   Begin VB.Shape Shape12 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   0
      Top             =   2520
      Width           =   12855
   End
   Begin VB.Shape Shape13 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   12840
      Top             =   3840
      Width           =   9735
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Menu mnufile 
      Caption         =   "Just &Auto Insurance"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuTrayMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuTraySize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuTrayMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Forma_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim mes_actual As Integer, ano_actual As Integer, calen1 As Integer, calen2 As Integer, guarda$, recibos$(5000), concepto$

Dim opcion_find As Integer, fila_total As Integer, fila_hecha As Integer, seg As Integer, valor_barra As Integer
' Dim fila_actual_de_busqueda As Integer


Dim SHORTNAME$, limite_inferior, limite_superior, grabado As Integer, ID_COMPANY1(5) As Integer, rango_dia As Integer, rango As Integer


Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&



Public Sub Checa_status()
On Error Resume Next




End Sub
Public Sub SetTrayMenuItems(window_state As Integer)
    Select Case window_state
        Case vbMinimized
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = False
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        Case vbMaximized
            mnuTrayMaximize.Enabled = False
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        Case vbNormal
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = True
            mnuTrayRestore.Enabled = False
            mnuTraySize.Enabled = True
    End Select
    
End Sub

    
Private Sub btn_borra_RichTextBox1_Click()
On Error Resume Next
RichTextBox1.Text = ""
RichTextBox1.Font.Name = "tahoma"
RichTextBox1.Font.Size = 7

RichTextBox1.SetFocus
End Sub

Private Sub btnagregar_nota_Click()
On Error Resume Next
If lblrow.Caption = "" Then
  MsgBox "There is not a row selected. Choose a row where you want to add the comment.", 16, "Attention"
  Exit Sub
End If


If op_rem(1).Value = True Then

Grid1.Row = lblrow.Caption
Grid1.Col = 15
Grid1.Text = txtnota.Text
Grid1.ColAlignment = 1

Grid1.Col = 22
Grid1.Text = "0"

ElseIf op_rem(0).Value = True Then

Grid3.Row = lblrow.Caption
Grid3.Col = 15
Grid3.Text = txtnota.Text
Grid3.ColAlignment = 1

Grid3.Col = 22
Grid3.Text = "0"

Else



End If


txtnota.Text = ""
lblrow.Caption = ""
grabado = 0
End Sub




Private Sub btnajuste_Click()
On Error Resume Next

msg.Visible = True
lblmsg.Caption = "Processing all the information..."
msg.Refresh



 
List3.Clear
 
again:
For t = 0 To List2.ListCount - 1
   'If List2.Selected(t) = False Then
   '  List2.RemoveItem t
   '  GoTo again
   'End If
   
   If List2.Selected(t) = True Then
      List3.AddItem List2.List(t)
   End If
   
Next t


List2.Clear


For t = 0 To List3.ListCount - 1
   List2.AddItem List3.List(t)
Next t

For t = 0 To List2.ListCount - 1
   List2.Selected(t) = True
Next t


ajusta_tabla



 'calcula_total_multiple
    
  
    lbltotal_lista.Caption = List2.ListCount
    
    

 btnguardar_Click
btnmerge.Enabled = True

msg.Visible = False


End Sub

Private Sub btnborra_cant_Click()
On Error Resume Next
txtamount.Text = ""
End Sub

Private Sub btnborra_cust_id_Click()
On Error Resume Next
txtcust_id.Text = ""
txtcust_id.SetFocus
End Sub

Private Sub btnborra_fila_Click()
On Error Resume Next

Grid3.Col = 0

r$ = InputBox("Type the row number you want to remove", "Remove row", Grid3.Text)
If r$ = "" Then
   Exit Sub
End If

Grid3.RemoveItem (Val(r$))  '+ 1)

conta = 0
For t = 1 To Grid3.Rows - 1
   conta = conta + 1
   Grid3.Row = t
   Grid3.Col = 0
   Grid3.Text = conta
Next t

lbltotal1.Caption = Grid3.Rows - 1

End Sub

Private Sub btnborra_oficinas_Click()
On Error Resume Next
cbo_oficina.ListIndex = -1

End Sub

Private Sub btnbusca_registro_Click()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset
    
    
    
busca_x_nombre = False
If txtfirst_name.Text <> "" Then
   busca_x_nombre = True
End If


busca_x_apellido = False
If txtlast_name.Text <> "" Then
   busca_x_apellido = True
End If
        
        
        
        
    
If txtfecha(0).Text = "" And txtfecha(1).Text = "" Then
   MsgBox "You have not set a date range yet", 16, "Attention"
   Exit Sub
End If
    
    
If txtamount.Text = "" And txtpoliza.Text = "" And cbocompany.ListIndex = -1 And busca_x_apellido = False And busca_x_nombre = False Then
   MsgBox "You have not established the amount, the policy or the company yet", 16, "Attention"
   Exit Sub
End If




    
    
grid4.Clear
grid4.Rows = 2

poliza$ = UCase(txtpoliza.Text)

idcompany$ = Right(cbocompany.List(cbocompany.ListIndex), 4)
compania$ = Left(cbocompany.List(cbocompany.ListIndex), Len(cbocompany.List(cbocompany.ListIndex)) - 4)



If idcompany$ = "" Then
  r$ = ""

    r$ = UCase(concepto$)
    SHORTNAME$ = ""
    
    X1$ = revisa_compania(r$)
    
    
    
    
   
   
continua_aqui:

 
 
  
    Set Rs = New ADODB.Recordset
            'Checa_status
   
            idcompany$ = ""
            sSelect = "select idcompany from insurancecatalog where shortname='" + SHORTNAME$ + "'"
    
           ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
            Rs.Open sSelect, base, adOpenUnspecified
    
            idcompany$ = Rs(0)
    
            Rs.Close
    
    
    
End If
  
  

Set Rs = New ADODB.Recordset
    
    
activo = chkactive.Value
    
    

    
    
rango1 = Val(txtamount.Text) * limite_inferior
rango2 = Val(txtamount.Text) * limite_superior
    
    
poliza2$ = "xxx"
If Val(idcompany$) = 10 Then
   If Mid$(poliza$, 4, 1) = "-" Then
         poliza2$ = Left(poliza$, 3) + Mid$(poliza$, 5, 5) + Mid$(poliza$, 11, 4) + Right$(poliza$, 3)
   Else
         poliza2$ = Left(poliza$, 3) + "-" + Mid$(poliza$, 4, 5) + "-" + Mid$(poliza$, 9, 4) + "-" + Right(poliza$, 2)
   End If
End If
    
    


encontrado = False
    
    
If chk_exact.Value = 1 Then
   
    
  If txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
  "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
  "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount='" + Format(txtamount.Text, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True

  ElseIf txtamount.Text = "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then


  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice  from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"

   encontrado = True

  ElseIf txtamount.Text <> "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then


  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice  from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount='" + Format(txtamount.Text, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
  ElseIf txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount='" + Format(txtamount.Text, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and iicat.IsPremium=1"
  
  encontrado = True

  ElseIf txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text = "" And cbocompany.ListIndex >= 0 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and recdtl.amount='" + Format(txtamount.Text, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
  ElseIf txtamount.Text = "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and iicat.IsPremium=1"
   
   encontrado = True
   
  ElseIf txtamount.Text <> "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount='" + Format(txtamount.Text, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and iicat.IsPremium=1"
    
  encontrado = True

  ElseIf txtamount.Text = "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then
    
    
    If idcompany$ = "1095" Then
        idcompany$ = "1095, 1094"
    End If
    
 sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and polhdr.idcompany in (" + idcompany$ + ") and iicat.IsPremium=1"
  

  encontrado = True
  
 End If

    
    
    
    
    
Else
'  *********************  cantidades aproximadas   *********************************




 If txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount>='" + Format(rango1, "#######0.00") + "' and recdtl.amount<='" + Format(rango2, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
 ElseIf txtamount.Text = "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then


  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"

   encontrado = True
   
 ElseIf txtamount.Text <> "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then


  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount>='" + Format(rango1, "#######0.00") + "' and recdtl.amount<='" + Format(rango2, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
 ElseIf txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount>='" + Format(rango1, "#######0.00") + "' and recdtl.amount<='" + Format(rango2, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and iicat.IsPremium=1"
  
  
  encontrado = True
  
 ElseIf txtamount.Text <> "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text = "" And cbocompany.ListIndex >= 0 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer , cust.firstname, cust.lastname1, rechdr.idoffice  from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and recdtl.amount>='" + Format(rango1, "#######0.00") + "' and recdtl.amount<='" + Format(rango2, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
 ElseIf txtamount.Text = "" And txtpoliza.Text <> "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and (polhdr.PolicyNumber='" + txtpoliza.Text + "' or polhdr.PolicyNumber='" + poliza2$ + "') and iicat.IsPremium=1"
   
   
  encontrado = True
  
 ElseIf txtamount.Text <> "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex = -1 Then
    
    
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1 , rechdr.idoffice  from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + "'  and recdtl.amount>='" + Format(rango1, "#######0.00") + "' and recdtl.amount<='" + Format(rango2, "#######0.00") + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and iicat.IsPremium=1"
    
  encontrado = True
  
 ElseIf txtamount.Text = "" And txtpoliza.Text = "" And txtfecha(0).Text <> "" And txtfecha(1).Text <> "" And cbocompany.ListIndex >= 0 Then
    
    
 sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and polhdr.idcompany='" + idcompany$ + "' and iicat.IsPremium=1"
  
  encontrado = True
  
 End If



End If
    
    
    
If encontrado = False Then
 
  sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
  "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
  "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.date>='" + txtfecha(0).Text + "' and rechdr.date<='" + txtfecha(1).Text + _
  "' and rechdr.Active='" + Format(activo, "0") + "' and iicat.IsPremium=1"

End If

    
    
' agrega en la busqueda el nombre o el apellido
    
If busca_x_nombre = True And busca_x_apellido = False Then
    
      sSelect = sSelect + " and cust.firstname='" + txtfirst_name.Text + "'"
 
ElseIf busca_x_apellido = True And busca_x_nombre = False Then
     sSelect = sSelect + " and cust.lastname1='" + txtlast_name.Text + "'"

ElseIf busca_x_apellido = True And busca_x_nombre = True Then
     sSelect = sSelect + " and cust.lastname1='" + txtlast_name.Text + "' and cust.firstname='" + txtfirst_name.Text + "'"
     
        
End If
    
    
    
' agrega el cust ID a la busqueda
    
If txtcust_id.Text <> "" Then
   sSelect = sSelect + " and rechdr.IdCustomer='" + LTrim(RTrim(txtcust_id.Text)) + "'"
End If
    
    
' agrega el programa a la busqueda
    
If cboprogram.ListIndex >= 0 Then
   sSelect = sSelect + " and catal.idprogram='" + LTrim(RTrim(Right(cboprogram.List(cboprogram.ListIndex), 10))) + "'"
End If
    
    
' agrega la oficina a la busqueda
If cbo_oficina.ListIndex >= 0 Then
   sSelect = sSelect + " and rechdr.idoffice='" + LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 10))) + "'"
End If
    
    
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid4.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid4.DataSource = Rs
    ' user1$ = Rs(0)
          
                        
                         
    Rs.Close
    
    
       lblfila_grid4.Clear
       grid4.Col = 0
    For t = 1 To grid4.Rows - 1
       grid4.Row = t
       grid4.Text = t
       lblfila_grid4.AddItem t
    Next t
    
    
    
    enca_grid4
    
    
    calcula_total_multiple

    
    lineas = grid4.Rows - 1
    If lineas < 0 Then lineas = 0
    lbltotal4.Caption = Str(lineas)
    
   
    
    
    
    
    
End Sub

Private Sub btnbusca2_Click()
On Error Resume Next
If txtcantidad.Text = "" Then Exit Sub
If op_busca(0).Value = True Then

  If chk_exacto2.Value = True Or chk_exacto2.Value = 1 Then
     cant = Val(txtcantidad.Text)
  Else
     cant = Int(Val(txtcantidad.Text))
  End If
  
  
  
  existe = 0
  For t = 1 To Grid3.Rows - 1
     Grid3.Row = t
     Grid3.Col = 3
     debito = Val(Grid3.Text)
     Grid3.Col = 4
     credito = Val(Grid3.Text)
     
     cantidad = debito + credito
     
     If chk_exacto2.Value = False Then
       debito = Int(debito)
       credito = Int(credito)
     End If
     
     
     
     If cant = debito Or cant = credito Then
       existe = 1
       With Grid3
            .Row = t
            .RowSel = .Row
            .Col = 3
            .ColSel = 9
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the amount of" + Str(cantidad) + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
     End If
     
  Next t
  
  
  
  
Else


  p$ = txtcantidad.Text
  existe = 0
  For t = 1 To Grid3.Rows - 1
     Grid3.Row = t
     Grid3.Col = 8
     poliza$ = UCase(Grid3.Text)
     
     
     If chk_exacto2.Value = True Then
     
      
      If UCase(p$) = UCase(poliza$) Then
       existe = 1
       With Grid3
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
      
       
       
      End If
     Else
     
      If UCase(p$) = Left(UCase(poliza$), Len(p$)) Then
       existe = 1
       With Grid3
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
       
      End If
     
     
     End If
     
     
     
     
       
      
      
  Next t
  
  
  

End If


MsgBox "End of search", 64, "Attention"

End Sub

Private Sub btnbusca3_Click()
On Error Resume Next
If txtcantidad3.Text = "" Then Exit Sub
If op_busca3(0).Value = True Then



  If chk_exacto3.Value = True Or chk_exacto3.Value = 1 Then
     cant = Val(txtcantidad3.Text)
  Else
     cant = Int(Val(txtcantidad3.Text))
  End If
  
  
  
  existe = 0
  For t = 1 To grid5.Rows - 1
     grid5.Row = t
     grid5.Col = 6
     cantidad = Val(grid5.Text)
     
     If chk_exacto3.Value = False Then
       cantidad = Int(cantidad)
       
     End If
     
     
     
     If cant = cantidad Then
       existe = 1
       With grid5
            .Row = t
            .RowSel = .Row
            .Col = 6
            .ColSel = 9
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the amount of " + grid5.Text + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
     End If
     
  Next t
  
  
  
  
Else


  p$ = txtcantidad3.Text
  existe = 0
  For t = 1 To grid5.Rows - 1
     grid5.Row = t
     grid5.Col = 11
     poliza$ = UCase(grid5.Text)
     
     
     If chk_exacto3.Value = True Then
     
      
      If UCase(p$) = UCase(poliza$) Then
       existe = 1
       With grid5
            .Row = t
            .RowSel = .Row
            .Col = 11
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
      
       
       
      End If
     Else
     
      If UCase(p$) = Left(UCase(poliza$), Len(p$)) Then
       existe = 1
       With grid5
            .Row = t
            .RowSel = .Row
            .Col = 11
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
       
      End If
     
     
     End If
     
     
     
     
       
      
      
  Next t
  
  
  

End If


MsgBox "End of search", 64, "Attention"
End Sub


Private Sub btncarga_excel_Click()
On Error Resume Next
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long

barra.Visible = True
barra.Min = 1

barra.Min = 1
barra.Max = Grid1.Rows
barra.Value = 0

' establece el color de la barra en verde
    Color_Progreso barra.hwnd, &HC0FFC0
    

barra.Refresh


Grid1.Visible = False
'barra.Max = 5000
lblmsg.Caption = "Loading the information from Excel..."
lblmsg2.Caption = ""
Timer1.Enabled = True
openforms = DoEvents
n$ = ""
cd1.FileName = ""
cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  r$ = ""
  For t = Len(n$) To 1 Step -1
     If Mid(n$, t, 1) <> "\" Then
       r$ = Mid(n$, t, 1) + r$
     Else
       Exit For
     End If
  Next t
  
  lblarchivo.Caption = r$
  
  If n$ = "" Then
    Grid1.Visible = True
    barra.Visible = False
    Exit Sub
  End If
  
msg.Visible = True
msg.Refresh

'btn_NB.Visible = False

contador = 0
Erase recibos$
Grid3.Clear

inicio:

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

grabado = 0
Grid1.Clear

'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing


'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True



'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(FileName:=n$, ReadOnly:=True)



' Get the first worksheet.
 
  Set xlHoja = xlApp.Worksheets(1)
 

 ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.Row
    ultimacolumnax = ActiveCell.Column
    

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

If ultimafilax = 0 Then
  lngUltimaFila = 5000
Else
  lngUltimaFila = ultimafilax
End If
   

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_NB
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_NB = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), xlHoja.Cells(lngUltimaFila, 22))   ' cambie 10 por 19

Grid1.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid1.Rows = lngUltimaFila + 1
Grid1.cols = 23

cont = 0
linea_vacia = 0
barra.Max = Grid1.Rows

For t = 1 To Grid1.Rows ' - 2
  
  barra.Value = t
  openforms = DoEvents
  
  
  Grid1.Row = t
  If t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
   
    If t > 0 Then
     cont = cont + 1
     Grid1.Col = 0
     Grid1.Text = cont
    End If
    
  End If
  
  Grid1.Row = t - 1
  veces = 0
  For Y = 1 To 21
   Grid1.Col = Y
   Grid1.Text = varMatriz(t, Y)
   If Grid1.Text = "" Then
     veces = veces + 1
     If veces >= 18 Then
        linea_vacia = linea_vacia + 1
     End If
   End If
  Next Y
  If linea_vacia >= 5 Then
     Exit For
  End If
Next t
   




Grid1.Rows = cont + 1 '+ 2


'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing


' arreglamos el orden de las columnas
' ****************************************************************************************

grid2.Clear

For Y = 1 To Grid1.cols - 1
  Grid1.Row = 0
  Grid1.Col = Y
  titulo$ = LTrim(RTrim(UCase(Grid1.Text)))
  
  
  
  
Next Y
 
 
 
grid2.Rows = Grid1.Rows
grid2.cols = 9


 For t = 1 To Grid1.Rows - 1
 
  grid2.Col = 0
  grid2.Row = t
  grid2.Text = Str(t)
  
  For Y = 1 To 8
  
    Grid1.Row = 0
    Grid1.Col = Y
    titulo$ = LTrim(RTrim(UCase(Grid1.Text)))
    
    If titulo$ = "ACCOUNT NUMBER" Or titulo$ = "ACCOUNT" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 1
       grid2.Text = Grid1.Text
       
    End If
      
      
      
   
    If titulo$ = "POST DATE" Or titulo$ = "DATE" Or titulo$ = "DATE CREATED" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 6
       grid2.Text = Grid1.Text
      
    End If
    
    
    If titulo$ = "CHECK" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 2
       grid2.Text = Grid1.Text
      
    End If
    
    
    
    If titulo$ = "DESCRIPTION" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 7
       grid2.Text = Grid1.Text
      
    End If
    
    
     If titulo$ = "DEBIT" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 3
       grid2.Text = Grid1.Text
      
    End If
    
    
     If titulo$ = "CREDIT" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 4
       grid2.Text = Grid1.Text
      
    End If
    
    
    If titulo$ = "STATUS" Then
      
    End If
    
    
     If titulo$ = "BALANCE" Then
       Grid1.Row = t
       Grid1.Col = Y
       
       grid2.Row = t
       grid2.Col = 5
       grid2.Text = Grid1.Text
      
    End If
    
    
    
   
  Next Y
  
 Next t




 Grid1.Clear
 
 
 For t = 1 To grid2.Rows - 1
    grid2.Row = t
    Grid1.Row = t
    Grid1.Col = 0
    Grid1.Text = Str(t)
    
    For Y = 1 To grid2.cols - 1
       grid2.Col = Y
       Grid1.Col = Y
       Grid1.Text = grid2.Text
    Next Y
  Next t


 

' ****************************************************************************************



enca_grid1
enca_grid3




final:
barra.Visible = False
msg.Visible = False


' ajusta el mes
ajusta_mes


Grid3.Rows = 2
Grid3.cols = 23








separa_efectivos





btnsepara_polizas_Click





Grid3.FixedRows = 1
Grid3.FixedCols = 1

Grid1.Visible = True
   

Grid1.Rows = Grid1.Rows - 1
btncarga_excel.Enabled = False


Grid1.Row = Grid1.Rows
vacio = True
For Y = 1 To Grid1.cols - 1
   Grid1.Col = Y
   r$ = Grid1.Text
   If r$ <> "" Then
      vacio = False
      Exit For
   End If
Next Y

If vacio = True Then
   Grid1.Rows = Grid1.Rows - 1
End If

lbltotal2.Caption = Format(Grid1.Rows - 1, "###,##0")


'Grid1.Row = Grid1.Rows
' Grid1.Col = 0
 
' Grid1.Text = "" ' Grid1.Rows - 1
 Timer1.Enabled = False
 
 btnverifica_polizas.Enabled = True
 


End Sub


Public Sub separa_polizas()
On Error Resume Next

Exit Sub

 Dim sSelect As String
    
 Dim Rs As ADODB.Recordset
 

barra.Visible = True
barra.Min = 1
barra.Max = Grid1.Rows
barra.Value = 0

' establece el color de la barra en verde
    Color_Progreso barra.hwnd, &HC0FFC0
    

barra.Refresh

   ' fila_actual_de_busqueda = 0
   
   
'valor_inicio = 1
'If Grid1.Rows <= 100 Then
'  valor_final = Grid1.Rows - 1
'Else
'  valor_final = 100
'End If

 
rutina:
   
'For w = 1 To (Grid1.Rows - 1) Step 100
   
'   valor_inicio = w
'   valor_final = valor_inicio + 99
   
'For t = valor_inicio To valor_final


 For t = 1 To Grid1.Rows - 1
   
   'fila_actual_de_busqueda = fila_actual_de_busqueda + 1
   barra.Value = t
   Grid1.Row = t
   Grid1.Col = 7
   descrip$ = ""
   descrip$ = Grid1.Text
   r$ = ""
   pos = InStr(1, descrip$, "MIL")
   
   Grid1.Col = 3
   cantidad = Val(Format(Grid1.Text, "0000.00"))
   
   Grid1.Col = 6
   fecha_banco$ = Format(Grid1.Text, "mm/dd/yyyy")
   
   Select Case Val(Left(fecha_banco$, 2))
   Case 1
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "12/24/" + Format(Val(Right(fecha_banco$, 4)) - 1, "0000")
      Else
        fecha_busqueda_inicial$ = "01/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
      
   Case 2
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "01/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "02/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
      
   Case 3
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "02/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "03/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 4
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "03/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "04/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 5
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "04/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "05/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 6
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "05/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "06/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 7
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "06/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "07/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 8
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "07/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "08/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 9
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "08/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "09/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 10
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "09/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "10/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 11
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "10/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "11/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   Case 12
      If Val(Mid(fecha_banco$, 4, 2)) < 10 Then
        fecha_busqueda_inicial$ = "11/24/" + Format(Val(Right(fecha_banco$, 4)), "0000")
      Else
        fecha_busqueda_inicial$ = "12/" + Format(Val(Mid(fecha_banco$, 4, 2)) - 9, "00") + Right(fecha_banco$, 5)
      End If
   End Select
                 
   
   
   
   
   
   
   
   'If cantidad = 147.88 Then Stop
   
   If pos = 0 Then
      GoTo anchor
   End If
   
   
   lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid1.Rows - 1, "###0")
   lblmsg2.Refresh
   openforms = DoEvents
   
   r$ = Mid$(descrip$, pos, Len(descrip$) - pos + 1)
   
   p$ = ""
   ' le quita el extra a la poliza en ALLIANCE  MILxxxxxxx
   If r$ <> "" Then
     pos = InStr(1, r$, "-")
     p$ = Left(r$, pos - 1)
     
     
     
     resul$ = p$
     r$ = verifica_existencia_de_poliza(resul$)
     
     If resul$ = "0" Then
       
       Grid1.Row = t
       Grid1.Col = 8
       Grid1.Text = Left(p$, 10)
     End If
     
   End If
   
   Grid1.ColWidth(8) = 1900
   
   GoTo brincado
   
   
   
anchor:
   
   If UCase(Left(descrip$, 6)) = "ANCHOR" Then
       p$ = ""
       pos = InStr(1, descrip$, "PREM")
   
       If pos = 0 Then
          GoTo OCEAN_HARBOR
       End If
       
       p$ = Mid$(descrip$, pos, Len(descrip$) - pos + 1)
       a$ = ""
       a$ = Mid(p$, 6, Len(p$) - 5)
       p$ = RTrim(LTrim(Left(a$, 7)))
       
     resul$ = p$
     r$ = verifica_existencia_de_poliza(resul$)
     
     If resul$ = "0" Then
        Grid1.Row = t
        Grid1.Col = 8
        Grid1.Text = p$
     
     End If
     
       
   
   End If
   
   
OCEAN_HARBOR:

   If UCase(Left(descrip$, 12)) = UCase("Ocean Harbor") Then
              
       p$ = ""
       pos = InStr(1, descrip$, "OHA")
           
       p$ = Mid$(descrip$, pos, 10)
              
       
     resul$ = p$
     r$ = verifica_existencia_de_poliza(resul$)
     
     If resul$ = "0" Then
        Grid1.Row = t
        Grid1.Col = 8
        Grid1.Text = p$
     
     End If
       
            
   
   End If
   
   
   
Alliance_KEMPER:

   If UCase(Left(descrip$, 14)) = UCase("AllianceUnited") Then
              
       p$ = ""
       pos = InStr(1, descrip$, "MNS")
           
       p$ = Mid$(descrip$, pos, 10)
              
     resul$ = p$
     r$ = verifica_existencia_de_poliza(resul$)
     
     If resul$ = "0" Then
        Grid1.Row = t
        Grid1.Col = 8
        Grid1.Text = p$
     
     End If
       
       
       
   
   End If
   
   
   
   
   
infinity:


   If UCase(Left(descrip$, 15)) = UCase("Upload Infinity") Then
       p$ = ""
       p$ = Right(descrip$, 15)
       
     resul$ = p$
     r$ = verifica_existencia_de_poliza(resul$)
     
     If resul$ = "0" Then
        Grid1.Row = t
        Grid1.Col = 8
        Grid1.Text = p$
     
     End If
       
       
   
   End If
   
   
   
national_general:


   If UCase(Left(descrip$, 16)) = UCase("National General") Then
       
       SHORTNAME$ = "NATIONAL GENERAL"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
       
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
    
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         If UCase(Left(p$, 3)) <> "MIL" Then
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         End If
       End If
       
       
       
       Next z
                         
       Rs.Close
  


   End If
   
   
   
   
SAFEWAY:


   If UCase(Left(descrip$, 7)) = UCase("Safeway") Then
       
       SHORTNAME$ = "SAFEWAY"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  



   End If
   
   
   
ASPIRE:


   If UCase(Left(descrip$, 6)) = UCase("Aspire") Then
       
       SHORTNAME$ = "ASPIRE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
   
   
   
   
BRIDGER:


   If UCase(Left(descrip$, 7)) = UCase("Bridger") Then
       
       SHORTNAME$ = "BRIDGER"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
   
   
   
   
BRISTOL:


   If UCase(Left(descrip$, 11)) = UCase("Bristolwest") Then
       
       SHORTNAME$ = "BRISTOL"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
           
       If grid2.Rows > 1 Then
           
        For z = 1 To grid2.Rows - 1
       
         grid2.Row = z
         grid2.Col = 1
         p$ = grid2.Text
       
       
              
         If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
         End If
       
       
       
        Next z
        
       Else
       
       
       
        p$ = ""
        pos = InStr(1, descrip$, "BRISTOLWEST")
   
        If pos = 0 Then
          GoTo CARNEGIE
        End If
       
        p$ = Mid$(descrip$, pos, Len(descrip$) - pos + 1)
        a$ = ""
        a$ = Mid(p$, 13, 10)
        p$ = RTrim(LTrim(a$))
        p$ = Left(p$, 3) + " " + Mid$(p$, 4, 7) + " " + "00"
       
       
        resul$ = p$
        r$ = verifica_existencia_de_poliza(resul$)
     
        If resul$ = "0" Then
          Grid1.Row = t
          Grid1.Col = 8
          Grid1.Text = p$
     
        End If
     
     
     
                         
       End If
                         
       Rs.Close
  
  


   End If
   
   
   
   
CARNEGIE:


   If UCase(Left(descrip$, 8)) = UCase("Carnegie") Then
       
       SHORTNAME$ = "CARNEGIE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
     
     
     
MULTI_State:


   If UCase(Left(descrip$, 16)) = UCase("Century-national") Or UCase(Left(descrip$, 10)) = UCase("Multistate") Then
       
       SHORTNAME$ = "MULTI-STATE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
   
   
   
   
   
COAST_NATIONAL:


   If UCase(Left(descrip$, 14)) = UCase("COAST NATIONAL") Then
       
       SHORTNAME$ = "BRISTOL"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
   
   
   
   
   
commerce_west:


   If UCase(Left(descrip$, 13)) = UCase("COMMERCE WEST") Then
       
       SHORTNAME$ = "MAPFRE"""
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  


   End If
   
   
   
   
SUN_COAST:


   If UCase(Left(descrip$, 15)) = UCase("DEBIT SUN COAST") Then
       
       SHORTNAME$ = "SUNCOAST"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If
   
   
   
   
   
KEMPER:


   If UCase(Left(descrip$, 6)) = UCase("KEMPER") Then
       p$ = ""
       grid2.Clear
       
       SHORTNAME$ = "KEMPER"
       
       ' separa el nombre o apellido
       nombre$ = ""
       For Y = Len(descrip$) To 1 Step -1
          r$ = Mid$(descrip$, Y, 1)
          If r$ = "0" Or r$ = "1" Or r$ = "2" Or r$ = "3" Or r$ = "4" Or r$ = "5" Or r$ = "6" Or r$ = "7" Or r$ = "8" Or r$ = "9" Then
                Exit For
          Else
               nombre$ = r$ + nombre$
          End If
       Next Y
       
            
       Set Rs = New ADODB.Recordset
   
            
       sSelect = "select polhdr.PolicyNumber, cust.firstname, cust.lastname1 from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
      
             
       
                         Rs.Open sSelect, base, adOpenUnspecified
        
                       ' Permitir redimensionar las columnas
                        grid2.AllowUserResizing = flexResizeColumns

                        ' Asignar el recordset al FlexGrid
                        Set grid2.DataSource = Rs
                      
                        Rs.Close
         
                  
                               ' ************************
        If grid2.Rows > 1 Then
              existe = 0
              For Y = 1 To grid2.Rows
                  grid2.Row = Y
                  grid2.Col = 2
                  n1$ = grid2.Text
                  grid2.Col = 3
                  n2$ = grid2.Text
                  
                  If UCase(Left(n1$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
                  
                  If UCase(Left(n2$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
               Next Y
               
                  
        End If
       
       
       If p$ <> "" Then
         resul$ = p$
         r$ = verifica_existencia_de_poliza(resul$)
     
         If resul$ = "0" Then
           Grid1.Row = t
           Grid1.Col = 8
           Grid1.Text = p$
     
         End If
     
       End If
                         
       


   End If
   
   
   
   
   
MCGRAW:


   If UCase(Left(descrip$, 6)) = UCase("MCGRAW") Then
       
       SHORTNAME$ = "PACIFIC SPECIALTY"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  


   End If
   
   
   
   
NATIONS:


   If UCase(Left(descrip$, 7)) = UCase("NATIONS") Then
       
       
       p$ = ""
       pos = InStr(1, descrip$, "NMC")
           
       If pos > 0 Then
          p$ = Mid$(descrip$, pos, 10)
          GoTo salta_siguiente
       End If
       
       
       p$ = ""
       pos = InStr(1, descrip$, "NIC")
           
       If pos > 0 Then
          p$ = Mid$(descrip$, pos, 10)
          GoTo salta_siguiente
       End If
       
       
salta_siguiente:
         resul$ = p$
         r$ = verifica_existencia_de_poliza(resul$)
     
         If resul$ = "0" Then
           Grid1.Row = t
           Grid1.Col = 8
           Grid1.Text = p$
     
         End If
       
       
       

   End If
   
   
   
   
PAYMENT_NATIONS:


   If UCase(Left(descrip$, 15)) = UCase("PAYMENT NATIONS") Then
       p$ = ""
       grid2.Clear
       
       SHORTNAME$ = "NATIONS INSURANCE"
       p$ = ""
       ' separa el nombre o apellido
       nombre$ = Right(descrip$, Len(descrip$) - 25)
       
       
            
       Set Rs = New ADODB.Recordset
   
            
       sSelect = "select polhdr.PolicyNumber, cust.firstname, cust.lastname1 from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
      
             
       
                         Rs.Open sSelect, base, adOpenUnspecified
        
                       ' Permitir redimensionar las columnas
                        grid2.AllowUserResizing = flexResizeColumns

                        ' Asignar el recordset al FlexGrid
                        Set grid2.DataSource = Rs
                      
                        Rs.Close
         
                  
                               ' ************************
        If grid2.Rows > 1 Then
              existe = 0
              For Y = 1 To grid2.Rows
                  grid2.Row = Y
                  grid2.Col = 2
                  n1$ = grid2.Text
                  grid2.Col = 3
                  n2$ = grid2.Text
                  
                  If UCase(Left(n1$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
                  
                  If UCase(Left(n2$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
               Next Y
               
                  
        End If
       
       
       If p$ <> "" Then
         resul$ = p$
         r$ = verifica_existencia_de_poliza(resul$)
     
         If resul$ = "0" Then
           Grid1.Row = t
           Grid1.Col = 8
           Grid1.Text = p$
     
         End If
     
       End If
                         
       


   End If
   
   
   
   
Perman_gen:


   If UCase(Left(descrip$, 10)) = UCase("PERMAN GEN") Then
       
       SHORTNAME$ = "ANCHOR"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
        ' separa el nombre o apellido
       nombre$ = ""
       For Y = Len(descrip$) To 1 Step -1
          r$ = Mid$(descrip$, Y, 1)
          If r$ = "0" Or r$ = "1" Or r$ = "2" Or r$ = "3" Or r$ = "4" Or r$ = "5" Or r$ = "6" Or r$ = "7" Or r$ = "8" Or r$ = "9" Then
                Exit For
          Else
               nombre$ = r$ + nombre$
          End If
       Next Y
       
            
       Set Rs = New ADODB.Recordset
   
            
       sSelect = "select polhdr.PolicyNumber, cust.firstname, cust.lastname1 from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
      
             
       
                         Rs.Open sSelect, base, adOpenUnspecified
        
                       ' Permitir redimensionar las columnas
                        grid2.AllowUserResizing = flexResizeColumns

                        ' Asignar el recordset al FlexGrid
                        Set grid2.DataSource = Rs
                      
                        Rs.Close
         
                  
                               ' ************************
        If grid2.Rows > 1 Then
              existe = 0
              For Y = 1 To grid2.Rows
                  grid2.Row = Y
                  grid2.Col = 2
                  n1$ = grid2.Text
                  grid2.Col = 3
                  n2$ = grid2.Text
                  
                  If UCase(Left(n1$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
                  
                  If UCase(Left(n2$, 4)) = UCase(Left(nombre$, 4)) Then
                      existe = 1
                      grid2.Col = 1
                      p$ = grid2.Text
                      Exit For
                  End If
               Next Y
               
                  
        End If
       
       
       If p$ <> "" Then
         resul$ = p$
         r$ = verifica_existencia_de_poliza(resul$)
     
         If resul$ = "0" Then
           Grid1.Row = t
           Grid1.Col = 8
           Grid1.Text = p$
     
         End If
     
       End If
                         
  


   End If
   
   


pronto:


   If UCase(Left(descrip$, 6)) = UCase("PRONTO") Then
       
       SHORTNAME$ = "PRONTO INSURANCE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If



RIC:


   If UCase(Left(descrip$, 15)) = UCase("Reliant General") Then
       
       SHORTNAME$ = "RIC"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If



stonewood:


   If UCase(Left(descrip$, 9)) = UCase("STONEWOOD") Then
       
       SHORTNAME$ = "BLUEFIRE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  


   End If



aegis:


   If UCase(Left(descrip$, 5)) = UCase("AEGIS") Or UCase(Left(descrip$, 11)) = UCase("DB PREM INS") Then
       
       SHORTNAME$ = "AEGIS"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
ALLSTAR:


   If UCase(Left(descrip$, 16)) = UCase("ALL STAR GENERAL") Then
       
       SHORTNAME$ = "BLUEFIRE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
TAPCO:


   If UCase(Left(descrip$, 5)) = UCase("TAPCO") Then
       
       SHORTNAME$ = "TAPCO INSURANCE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
   
WORKMENS:


   If UCase(Left(descrip$, 8)) = UCase("WORKMENS") Then
       
       SHORTNAME$ = "WORKMENS"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
cypress:


   If UCase(Left(descrip$, 7)) = UCase("Cypress") Or UCase(Left(descrip$, 8)) = UCase("Scottish") Then
       
       SHORTNAME$ = "SCOTTISH AMERICAN"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If




ARROWHEAD:


   If UCase(Left(descrip$, 9)) = UCase("ARROWHEAD") Then
       
       SHORTNAME$ = "ARROWHEAD"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
ACH:


   If UCase(Left(descrip$, 11)) = UCase("ACH PROFILE") Then
       
       SHORTNAME$ = "BLUEFIRE"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
   
INS_PREM:


   If UCase(Left(descrip$, 24)) = UCase("INS PREM WESTERN GENERAL") Then
       
       SHORTNAME$ = "WESTERN"
       p$ = ""
       Set Rs = New ADODB.Recordset
   
       sSelect = "select polhdr.PolicyNumber from ReceiptsHDR rechdr " & _
       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
       "where rechdr.date>='" + fecha_busqueda_inicial$ + "' and rechdr.date<='" + fecha_banco$ + "'  and recdtl.amount='" + Format(cantidad, "#######0.00") + _
       "' and rechdr.Active=1 and shortname='" + SHORTNAME$ + "' and iicat.IsPremium=1"
       
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
           Rs.Open sSelect, base, adOpenUnspecified
   
    
           ' Permitir redimensionar las columnas
           grid2.AllowUserResizing = flexResizeColumns

           ' Asignar el recordset al FlexGrid
           Set grid2.DataSource = Rs
                      
           
       For z = 1 To grid2.Rows - 1
       
       grid2.Row = z
       grid2.Col = 1
       p$ = grid2.Text
       
       
              
       If p$ <> "" Then
         
           resul$ = p$
           r$ = verifica_existencia_de_poliza(resul$)
     
           If resul$ = "0" Then
             Grid1.Row = t
             Grid1.Col = 8
             Grid1.Text = p$
             Exit For
           
           End If
     
         
       End If
       
       
       
       Next z
                         
       Rs.Close
  
  

   End If
   
   
brincado:

Next t



' Next w



barra.Visible = False

End Sub









Private Sub btnctrl_v_Click()
RichTextBox1.Text = Clipboard.GetText
End Sub


Private Sub btndepura_Click()
On Error Resume Next
 If Grid3.Rows <= 2 Then
     Exit Sub
 End If
 
  
'r$ = MsgBox("Do you want to save all the information in the database?", 4, "Attention")
'If r$ = "7" Then Exit Sub
 
 
'Exit Sub



Dim sSelect As String
    
Dim Rs As ADODB.Recordset

 barra.Visible = True
  barra.Min = 1
  barra.Max = Grid3.Rows - 1
  
  msg.Visible = True
  
  lblmsg.Caption = "Debugging all the information in the database..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  
  contador = 0
  Erase recibos$
  
' -------------------------------------------------------------------------------------------------------

  For t = 1 To Grid3.Rows - 1
    
    Grid3.Row = t
    
    Grid3.Col = 21
    Idconciliation$ = Grid3.Text
    
    
       
    IDReceiptHDR$ = ""
    
    Grid3.Col = 1 ' account
    account$ = "243162505"
    
    Grid3.Col = 2 ' chkref
    chkref$ = Grid3.Text
    
    Grid3.Col = 3  ' debito
    debito = Val(Grid3.Text)
        
    Grid3.Col = 4  ' credit
    credito = Val(Grid3.Text)
    
    Grid3.Col = 5 ' balance
    balance = Val(Grid3.Text)
        
    Grid3.Col = 6   '  date
    fecha_pagada$ = Format(Grid3.Text, "yyyy-mm-dd")
    
    
    Grid3.Col = 7 ' descripcion
    Description$ = Grid3.Text
           
    Grid3.Col = 8  ' poliza
    poliza$ = UCase(Grid3.Text)
    
    Grid3.Col = 9  ' receipt HDR
    IDReceiptHDR$ = Grid3.Text
    
    Grid3.Col = 10 ' amount
    amount = Val(Format(Grid3.Text, "0000000.00"))
    
    Grid3.Col = 11 ' verificado
    verificado$ = Grid3.Text
    
    Grid3.Col = 12  ' date created
    fecha_creacion$ = Grid3.Text
        
    Grid3.Col = 13   ' Id cust
    IdCustomer$ = Grid3.Text
    
    Grid3.Col = 14   ' company
    compania$ = Grid3.Text
    
    Grid3.Col = 15 'Comment
    nota$ = Grid3.Text
    
    Grid3.Col = 16 ' program name Company
    programname$ = Grid3.Text
    
    Grid3.Col = 17 ' idprogram
    idprogram$ = Grid3.Text
    
    Set Rs = New ADODB.Recordset
      
    idprogram2$ = ""
    
    sSelect = "select idprogram from ProgramsCatalog where programname='" + idprogram$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    idprogram2$ = Rs(0)
    Rs.Close
    
    
    If idprogram2$ <> "" Then
       idprogram$ = idprogram2$
    End If
    
    
    
    
    
    Grid3.Col = 18 ' idreceiptDTL
    idreceiptdtl$ = Grid3.Text
    
    Grid3.Col = 19 ' idpolizaHDR
    idpolizahdr$ = Grid3.Text
    
    Grid3.Col = 20 '
    idcompany$ = Grid3.Text
    
    Grid3.Col = 21
    Idconciliation$ = Grid3.Text
    
    
    Idconciliation1$ = ""
    sSelect = "select idconciliation from ConciliationBankRec where MonthConciliation='" + Format(mes_actual, "#0") + "' and YearConciliation='" & _
    Format(cboyear.List(cboyear.ListIndex), "0000") + "' and debit='" + Format(debito, "###0.00") + "' and policyNo='" + poliza$ + "' and description='" & _
    Description$ + "' and idreceiptHDR='0'"
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation1$ = Rs(0)  'correcto
                         
    Rs.Close
    
    
    
    
    
    Idconciliation2$ = ""
    If Val(IDReceiptHDR$) > 0 Then
      sSelect = "select idconciliation from ConciliationBankRec where MonthConciliation='" + Format(mes_actual, "#0") + "' and YearConciliation='" & _
      Format(cboyear.List(cboyear.ListIndex), "0000") + "' and debit='" + Format(debito, "###0.00") + "' and policyNo='" + poliza$ + "' and description='" & _
      Description$ + "' and idreceiptHDR='" + IDReceiptHDR$ + "'"
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
      Idconciliation2$ = Rs(0)  'correcto
                         
      Rs.Close
    End If
    
    
    If Idconciliation1$ <> "" And Idconciliation2$ <> "" Then
        sSelect = "delete from ConciliationBankRec where Idconciliation='" + Idconciliation1$ + "'"
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
        Rs.Open sSelect, base, adOpenUnspecified
        Idconciliation3$ = Rs(0)  'correcto
                        
        Rs.Close
    
    End If
    
    
    
    
    
    ' If amount = 0 Then GoTo al_final
    
    
    contador = contador + 1
    barra.Value = contador
    lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid3.Rows - 1, "###0")
    lblmsg2.Refresh
    openforms = DoEvents
    
    
   
    
    
   
no_grabes:

    
al_final:
          
  Next t




msg.Visible = False
barra.Visible = False
End Sub

Private Sub BTNELIMINA_FILA_Click()
On Error Resume Next

If grid4.Rows <= 2 Or lblfila_grid4.Text = "" Then
  Exit Sub
End If

r$ = MsgBox("Do you wish to delete the row# " + lblfila_grid4.Text + "?", 4, "Attention")
If r$ = "7" Then Exit Sub

'For t = Val(lblfila_grid4.Text) To grid4.Rows - 1
   ' grid4.Row = t
   grid4.RemoveItem Val(lblfila_grid4.Text)
'Next t



lblfila_grid4.Clear

For t = 1 To grid4.Rows - 1
  grid4.Row = t
  grid4.Col = 0
  grid4.Text = t
  lblfila_grid4.AddItem t

Next t

lbltotal4.Caption = Str(grid4.Rows - 1)


Total = 0
For t = 1 To grid4.Rows - 1
   grid4.Row = t
   grid4.Col = 5
   Total = Total + Val(grid4.Text)
Next t

lblamount_total.Caption = Format(Total, "$###,##0.00")





'calcula_total_multiple

End Sub

Private Sub btnerase_program_Click()
On Error Resume Next
cboprogram.ListIndex = -1
End Sub

Private Sub btneraser_all_Click()
On Error Resume Next
cbocompany.ListIndex = -1
cboprogram.ListIndex = -1
txtpoliza.Text = ""
txtamount.Text = ""
txtfecha(0).Text = ""
txtfecha(1).Text = ""
grid4.Clear
grid4.Rows = 2
lbltotal4.Caption = "0"
lblfila_grid4.Text = ""
txtfila.Text = ""
txtlast_name.Text = ""
txtfirst_name.Text = ""
txtcust_id.Text = ""
lblamount_total.Caption = "0"

txtfecha(0).Text = Format(fecha_rango1$, "mm/dd/yyyy")
txtfecha(1).Text = Format(fecha_rango2$, "mm/dd/yyyy")
End Sub

Private Sub btnexcel_Click()
On Error Resume Next

'If lista_agentes.ListCount = 0 Then Exit Sub
If Grid1.Rows = 1 Then Exit Sub


r$ = MsgBox("Do you wish to transfer all the data from this grid to one Excel datasheet?", 4, "Confirm action")
If r$ = "7" Then Exit Sub



Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
    
msg.Visible = True
lblmsg.Caption = "Please, wait a moment... transferring data to Excel"
lblmsg2.Caption = ""
msg.Refresh
openforms = DoEvents

cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowSave
  n$ = cd1.FileName

    
    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)
'oSheet.range("A1").Value = "Last Name"
'oSheet.range("B1").Value = "First Name"
'oSheet.range("A1:B1").Font.Bold = True
'oSheet.range("A2").Value = "Doe"
'oSheet.range("B2").Value = "John"
'oSheet.range("A3").Value = "Vazquez"
'oSheet.range("B3").Value = "Maria"


'Create an array with 20 columns and 100 rows
Dim DataArray(1 To 5000, 1 To 21) As Variant
'Dim r As Integer



grandtotal = 0
num = 0
For t = 1 To Grid1.Rows - 1
      
    num = t
    Grid1.Row = t
    
        
    Grid1.Col = 1
    DataArray(num, 1) = Grid1.Text
    
    Grid1.Col = 2
    DataArray(num, 2) = Grid1.Text
    
    Grid1.Col = 3
    DataArray(num, 3) = Format(Grid1.Text, "#####0.00")
    
    Grid1.Col = 4
    DataArray(num, 4) = Grid1.Text
    
    Grid1.Col = 5
    DataArray(num, 5) = Format(Grid1.Text, "#####0.00")
    
    Grid1.Col = 6
    DataArray(num, 6) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.Col = 7
    DataArray(num, 7) = Grid1.Text
        
    Grid1.Col = 8
    DataArray(num, 8) = Grid1.Text
    
    Grid1.Col = 9
    DataArray(num, 9) = Grid1.Text
    
    Grid1.Col = 10
    DataArray(num, 10) = Format(Grid1.Text, "#####0.00")
    
    Grid1.Col = 11
    DataArray(num, 11) = Grid1.Text
    
    Grid1.Col = 12
    DataArray(num, 12) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.Col = 13
    DataArray(num, 13) = Grid1.Text
        
    Grid1.Col = 14
    DataArray(num, 14) = Grid1.Text
    
    Grid1.Col = 15
    DataArray(num, 15) = Grid1.Text
    
    Grid1.Col = 16
    DataArray(num, 16) = Grid1.Text
    
    Grid1.Col = 17
    DataArray(num, 17) = Grid1.Text
  
    Grid1.Col = 18
    DataArray(num, 18) = Grid1.Text
   
    Grid1.Col = 19
    DataArray(num, 19) = Grid1.Text
   
   Grid1.Col = 20
    DataArray(num, 20) = Grid1.Text
   
   Grid1.Col = 21
    DataArray(num, 21) = Grid1.Text
   


Next t


'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
' A1:N1 es la cantidad de columnas (rango)
oSheet.Range("A1:U1").Value = Array("Account", "CHK", "Debit", "Credit", "Balance", "Date", "Description", "Policy", "Receipt", "Amount", "Verified", "Date Created", "IDCustomer", "Company", "Comment", "Program Name", "IDprogram", "IDReceiptDTL", "IDPolicyHDR", "IDCompany", "IDConciliation")

'Transfer the array to the worksheet starting at cell A2
oSheet.Range("A2").Resize(5000, 21).Value = DataArray


'lblgrandtotal.Caption = Format(grandtotal, "#####0.00")



'Save the Workbook and Quit Excel
oBook.SaveAs n$
oExcel.Quit

msg.Visible = False

End Sub


Private Sub btnexcel2_Click()
On Error Resume Next

'If lista_agentes.ListCount = 0 Then Exit Sub
If Grid3.Rows = 1 Then Exit Sub

r$ = MsgBox("Do you wish to transfer all the data from this grid to one Excel datasheet?", 4, "Confirm action")
If r$ = "7" Then Exit Sub


Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
    
    
    
cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowSave
  n$ = cd1.FileName
  
  
If n$ = "" Then Exit Sub
  
    
msg.Visible = True
lblmsg.Caption = "Please, wait a moment... transferring data to Excel"
lblmsg2.Caption = ""
msg.Refresh
openforms = DoEvents


    
    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)
'oSheet.range("A1").Value = "Last Name"
'oSheet.range("B1").Value = "First Name"
'oSheet.range("A1:B1").Font.Bold = True
'oSheet.range("A2").Value = "Doe"
'oSheet.range("B2").Value = "John"
'oSheet.range("A3").Value = "Vazquez"
'oSheet.range("B3").Value = "Maria"


'Create an array with 20 columns and 100 rows
Dim DataArray(1 To 5000, 1 To 21) As Variant
'Dim r As Integer



grandtotal = 0
num = 0
For t = 1 To Grid3.Rows - 1
      
    num = t
    Grid3.Row = t
    
        
    Grid3.Col = 1
    DataArray(num, 1) = Grid3.Text
    
    Grid3.Col = 2
    DataArray(num, 2) = Grid3.Text
    
    Grid3.Col = 3
    DataArray(num, 3) = Format(Grid3.Text, "#####0.00")
    
    Grid3.Col = 4
    DataArray(num, 4) = Grid3.Text
    
    Grid3.Col = 5
    DataArray(num, 5) = Format(Grid3.Text, "#####0.00")
    
    Grid3.Col = 6
    DataArray(num, 6) = Format(Grid3.Text, "mm/dd/yyyy")
    
    Grid3.Col = 7
    DataArray(num, 7) = Grid3.Text
        
    Grid3.Col = 8
    DataArray(num, 8) = Grid3.Text
    
    Grid3.Col = 9
    DataArray(num, 9) = Grid3.Text
    
    Grid3.Col = 10
    DataArray(num, 10) = Format(Grid3.Text, "#####0.00")
    
    Grid3.Col = 11
    DataArray(num, 11) = Grid3.Text
    
    Grid3.Col = 12
    DataArray(num, 12) = Format(Grid3.Text, "mm/dd/yyyy")
    
    Grid3.Col = 13
    DataArray(num, 13) = Grid3.Text
        
    Grid3.Col = 14
    DataArray(num, 14) = Grid3.Text
    
    Grid3.Col = 15
    DataArray(num, 15) = Grid3.Text
    
    Grid3.Col = 16
    DataArray(num, 16) = Grid3.Text
    
    Grid3.Col = 17
    DataArray(num, 17) = Grid3.Text
  
    Grid3.Col = 18
    DataArray(num, 18) = Grid3.Text
   
  Grid3.Col = 19
    DataArray(num, 19) = Grid3.Text
   
   Grid3.Col = 20
    DataArray(num, 20) = Grid3.Text
   
   Grid3.Col = 21
    DataArray(num, 21) = Grid3.Text
   


Next t


'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
' A1:N1 es la cantidad de columnas (rango)
oSheet.Range("A1:U1").Value = Array("Account", "CHK", "Debit", "Credit", "Balance", "Date", "Description", "Policy", "Receipt", "Amount", "Verified", "Date Created", "IDCustomer", "Company", "Comment", "Program Name", "IDprogram", "IDReceiptDTL", "IDPolicyHDR", "IDCompany", "IDConciliation")

'Transfer the array to the worksheet starting at cell A2
oSheet.Range("A2").Resize(5000, 21).Value = DataArray


'lblgrandtotal.Caption = Format(grandtotal, "#####0.00")



'Save the Workbook and Quit Excel
oBook.SaveAs n$
oExcel.Quit

msg.Visible = False

End Sub

Private Sub btnexcel3_Click()
On Error Resume Next

'If lista_agentes.ListCount = 0 Then Exit Sub
If grid5.Rows = 1 Then Exit Sub

r$ = MsgBox("Do you wish to transfer all the data from this grid to one Excel datasheet?", 4, "Confirm action")
If r$ = "7" Then Exit Sub


Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
    
    
    
cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowSave
  n$ = cd1.FileName
  
  
If n$ = "" Then Exit Sub
  
    
msg.Visible = True
lblmsg.Caption = "Please, wait a moment... transferring data to Excel"
lblmsg2.Caption = ""
msg.Refresh
openforms = DoEvents


    
    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)
'oSheet.range("A1").Value = "Last Name"
'oSheet.range("B1").Value = "First Name"
'oSheet.range("A1:B1").Font.Bold = True
'oSheet.range("A2").Value = "Doe"
'oSheet.range("B2").Value = "John"
'oSheet.range("A3").Value = "Vazquez"
'oSheet.range("B3").Value = "Maria"


'Create an array with 20 columns and 100 rows
Dim DataArray(1 To 5000, 1 To 15) As Variant
'Dim r As Integer



grandtotal = 0
num = 0
For t = 1 To grid5.Rows - 1
      
    num = t
    grid5.Row = t
    
        
    grid5.Col = 1
    DataArray(num, 1) = grid5.Text
    
    grid5.Col = 2
    DataArray(num, 2) = grid5.Text
    
    grid5.Col = 3
    DataArray(num, 3) = grid5.Text
    
    grid5.Col = 4
    DataArray(num, 4) = grid5.Text
    
    grid5.Col = 5
    DataArray(num, 5) = Format(grid5.Text, "mm/dd/yyyy")
    
    grid5.Col = 6
    DataArray(num, 6) = Format(grid5.Text, "#####0.00")
    
    grid5.Col = 7
    DataArray(num, 7) = grid5.Text
        
    grid5.Col = 8
    DataArray(num, 8) = grid5.Text
    
    grid5.Col = 9
    DataArray(num, 9) = grid5.Text
    
    grid5.Col = 10
    DataArray(num, 10) = grid5.Text
    
    grid5.Col = 11
    DataArray(num, 11) = grid5.Text
    
    grid5.Col = 12
    DataArray(num, 12) = grid5.Text
    
    grid5.Col = 13
    DataArray(num, 13) = grid5.Text
        
    grid5.Col = 14
    DataArray(num, 14) = grid5.Text
    
    
Next t


'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
' A1:N1 es la cantidad de columnas (rango)
oSheet.Range("A1:N1").Value = Array("ReceiptHDR", "ReceiptDTL", "ReceiptDTL", "IDPoliciesHDR", "Date", "Amount", "IdCompany", "Idprogram", "Programname", "Companyname", "policynumber", "IDcustomer", "First name", "Last name")

'Transfer the array to the worksheet starting at cell A2
oSheet.Range("A2").Resize(5000, 15).Value = DataArray


'lblgrandtotal.Caption = Format(grandtotal, "#####0.00")



'Save the Workbook and Quit Excel
oBook.SaveAs n$
oExcel.Quit

msg.Visible = False

End Sub

Private Sub btnguardar_Click()
On Error Resume Next

grid6.Clear
grid6.Rows = grid4.Rows
grid6.cols = grid4.cols

For t = 0 To grid4.Rows - 1
   grid4.Row = t
   grid6.Row = t
   For Y = 1 To grid4.cols - 1
      grid4.Col = Y
      grid6.Col = Y
      grid6.Text = grid4.Text
   Next Y
Next t
      
'MsgBox "It was saved", 64, "Attention"


'btnguardar.Caption = "Saved"
'btnguardar.Enabled = False

End Sub

Private Sub btnleft_Click()
On Error Resume Next

valor_barra = valor_barra - 1
If valor_barra < 1 Then valor_barra = 1



Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dias_actual = 31
Case 4, 6, 9, 11
  dias_actual = 30
Case 2
   cant = (ano_actual / 4)
   residuo = cant - Int(cant)
   If residuo = 0 Then
      dias_actual = 29
   Else
      dias_actual = 28
   End If
End Select


If valor_barra >= dias_actual Or (valor_barra + rango_dia) >= dias_actual Then
   op_day(0).Value = True
   valor_barra = dias_actual
   rango_dia = 0
   
End If


  f1$ = Format(mes_actual, "00") + "/" + Format(valor_barra, "00") + "/" + Format(ano_actual, "0000")
  f2$ = Format(mes_actual, "00") + "/" + Format(valor_barra + rango_dia, "00") + "/" + Format(ano_actual, "0000")



txtfecha(0).Text = f1$
txtfecha(1).Text = f2$

lblfecha1.Caption = txtfecha(0).Text
lblfecha2.Caption = txtfecha(1).Text


btnbusca_registro_Click

carga_lista

calcula_total_multiple
End Sub

Private Sub btnlimbiacbo_Click()
On Error Resume Next
cbocompany.ListIndex = -1
End Sub

Private Sub btnlimpia_cte_Click()
txtlast_name.Text = ""
txtfirst_name.Text = ""
End Sub

Private Sub btnlimpia_grid1_Click()
On Error Resume Next
If Grid1.Rows = 1 Then Exit Sub

r$ = MsgBox("Do you want to clean the lower grid?", 4, "Attention")
If r$ = "7" Then Exit Sub

Grid1.Clear
Grid1.Rows = 1
enca_grid1
End Sub

Private Sub btnlimpia_grid2_Click()
On Error Resume Next
If Grid3.Rows = 1 Then Exit Sub

r$ = MsgBox("Do you want to clean the upper grid?", 4, "Attention")
If r$ = "7" Then Exit Sub

Grid3.Clear
Grid3.Rows = 1
enca_grid3

End Sub

Private Sub btnlimpiapoliza_Click()
On Error Resume Next
txtpoliza.Text = ""
End Sub

Private Sub btnload1_Click()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset


Dim rsVar As Variant
Dim i As Integer
  
Conecta_SQL
btncarga_excel.Enabled = False
  
  If Grid1.Rows > 1 Or Grid3.Rows > 1 Then
    graba_sql1
    graba_sql2
  End If
  
  msg.Visible = True
  
  lblmsg.Caption = "Loading all the information..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  
 
  
  Grid1.Visible = False
  Grid3.Visible = False
  grid5.Visible = False
  Grid1.Clear
  grid2.Clear
  Grid3.Clear
  grid5.Clear
  '
  
 Set Rs = New ADODB.Recordset
    
    
   If mes_actual = 0 Then
    For Y = 0 To 11
      If btnmes(Y).Value = True Then
         mes_actual = Y + 1
         Exit For
      End If
    Next Y
    
   End If
    
   ano_actual = cboyear.List(cboyear.ListIndex)
    
   'sSelect = "select * from ConciliationBankRec"
   
  If op_rango_mes(0).Value = True Then  ' TODO EL MES
    
   sSelect = "select cbank.account, cbank.chkref, cbank.debit, cbank.credit, cbank.balance, cbank.date, cbank.description, " & _
   "cbank.policyno, cbank.idreceipthdr, cbank.amount, cbank.clear, cbank.receiptdate, cbank.idcustomer, ins.CompanyName, " & _
   "cbank.notes , progcat.ProgramName, cbank.idprogram, cbank.idreceiptdtl, cbank.idpolicieshdr, cbank.idcompany, cbank.idconciliation " & _
   "from ConciliationBankRec cbank " & _
   "left join InsuranceCatalog ins on ins.IdCompany=cbank.IdCompany " & _
   "left join ProgramsCatalog progcat on progcat.IdProgram=cbank.IdProgram " & _
   "Where cbank.MonthConciliation ='" + Format(mes_actual, "#0") + "' And cbank.YearConciliation ='" + Format(ano_actual, "###0") + "'"  ' and clear=1"

  Else
  
   sSelect = "select cbank.account, cbank.chkref, cbank.debit, cbank.credit, cbank.balance, cbank.date, cbank.description, " & _
   "cbank.policyno, cbank.idreceipthdr, cbank.amount, cbank.clear, cbank.receiptdate, cbank.idcustomer, ins.CompanyName, " & _
   "cbank.notes , progcat.ProgramName, cbank.idprogram, cbank.idreceiptdtl, cbank.idpolicieshdr, cbank.idcompany, cbank.idconciliation " & _
   "from ConciliationBankRec cbank " & _
   "left join InsuranceCatalog ins on ins.IdCompany=cbank.IdCompany " & _
   "left join ProgramsCatalog progcat on progcat.IdProgram=cbank.IdProgram " & _
   "Where cbank.MonthConciliation ='" + Format(mes_actual, "#0") + "' And cbank.YearConciliation ='" + Format(ano_actual, "###0") + "' " & _
   "and cbank.date>='" + txtfecha_cargada(0).Text + "' and cbank.date<='" + txtfecha_cargada(1).Text + "'"
   
  
  
  End If
  
  
  
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   grid2.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   grid2.cols = Rs.Fields.Count
    
    
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      grid2.TextMatrix(0, i) = Rs.Fields(i).Name
   Next

   grid2.Row = 1
   grid2.Col = 0

   ' Set range of cells in the grid
   grid2.RowSel = grid2.Rows - 1
   grid2.ColSel = grid2.cols - 1
   grid2.clip = rsVar
   

   ' Reset the grid's selected range of cells
   grid2.RowSel = grid2.Row
   grid2.ColSel = grid2.Col

   Rs.Close

   Set Rs = Nothing
   
    
     ' Permitir redimensionar las columnas
   grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
   'Set grid2.DataSource = Rs
    ' user1$ = Rs(0)
   'Rs.Close

  

   Grid1.cols = 23 'grid2.cols + 1
   Grid3.cols = 23 'grid2.cols + 1
   
   Grid1.Rows = grid2.Rows '- 1
   Grid3.Rows = Grid1.Rows '- 1
   
   
   'Exit Sub


   cont1 = 0
   cont3 = 0
   fila1 = 0
   fila3 = 0
   
   For t = 1 To grid2.Rows - 1
      grid2.Row = t
      grid2.Col = 10
      
      If grid2.Text = "-1" Then
          fila1 = fila1 + 1
          Grid1.Row = fila1
          cont1 = cont1 + 1
          Grid1.Col = 0
          Grid1.Text = Str(cont1)
                
          For Y = 0 To grid2.cols - 1
        
              grid2.Col = Y
              Grid1.Col = Y + 1
          
              If Y <> 10 Then
                  Grid1.Text = grid2.Text
              Else
                  If grid2.Text = "-1" Then
                       Grid1.Text = "Ok"
                  Else
                       Grid1.Text = ""
                  End If
              End If
        
          Next Y
          
         
          
      Else
          fila3 = fila3 + 1
          Grid3.Row = fila3
          cont3 = cont3 + 1
          Grid3.Col = 0
          Grid3.Text = Str(cont3)
                
          For Y = 0 To grid2.cols - 1
        
              grid2.Col = Y
              Grid3.Col = Y + 1
          
              If Y <> 10 Then
                  Grid3.Text = grid2.Text
              Else
                  If grid2.Text = "-1" Then
                       Grid3.Text = "Ok"
                  Else
                       Grid3.Text = ""
                  End If
              End If
        
          Next Y
          
          
      
      End If
sigue_aqui:
      a = 1
   Next t


   Grid1.Rows = fila1 + 1
   Grid3.Rows = fila3 + 1
   
   enca_grid1
   enca_grid3
  


   Grid1.Visible = True
   Grid3.Visible = True
   grid5.Visible = True
   
  


   ajusta_mes

   ASIGNA_UNOS


   '
   carga_grid5

   enca_grid5
   
  
   lbltotal1.Caption = Format(fila3, "###,##0")
   lbltotal2.Caption = Format(fila1, "###,##0")
   'lbltotal5.Caption = Format(grid5.Rows - 2, "###,##0")
   

barra.Visible = False
msg.Visible = False












' ----------------------------------------------------------------------------------------------------------------------------------

Exit Sub
' rutina para importar desde excel

Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long

n$ = ""
r$ = MsgBox("Do you wish to load all the data from Excel datasheet to this grid?", 4, "Confirm action")
If r$ = "7" Then Exit Sub

barra.Visible = True
barra.Min = 1
Grid1.Visible = False
barra.Max = 5000
lblmsg.Caption = "Loading the information from Excel..."
lblmsg2.Caption = ""
openforms = DoEvents

cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  If n$ = "" Then
    barra.Visible = False
    Exit Sub
  End If
  
msg.Visible = True

'btn_NB.Visible = False

contador = 0
Erase recibos$
Grid1.Clear

inicio:

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing


'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(FileName:=n$, ReadOnly:=True)



' Get the first worksheet.
 
  Set xlHoja = xlApp.Worksheets(1)
 

 ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.Row
    ultimacolumnax = ActiveCell.Column
    

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

If ultimafilax = 0 Then
  lngUltimaFila = 5000
Else
  lngUltimaFila = ultimafilax
End If
   

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_NB
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_NB = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), xlHoja.Cells(lngUltimaFila, 22))   ' cambie 10 por 19

Grid1.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid1.Rows = lngUltimaFila + 2
Grid1.cols = 23

cont = 0
linea_vacia = 0
For t = 1 To Grid1.Rows - 2
  
  barra.Value = t
  openforms = DoEvents
  
  Grid1.Row = t
  If t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
   
    If t > 0 Then
     cont = cont + 1
     Grid1.Col = 0
     Grid1.Text = cont
    End If
    
  End If
  
  Grid1.Row = t - 1
  veces = 0
  For Y = 1 To 18
   Grid1.Col = Y
   Grid1.Text = varMatriz(t, Y)
   If Grid1.Text = "" Then
     veces = veces + 1
     If veces >= 18 Then
        linea_vacia = linea_vacia + 1
     End If
   End If
  Next Y
  If linea_vacia >= 5 Then
     Exit For
  End If
Next t
   




Grid1.Rows = cont  '+ 2


'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit

'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing






enca_grid1

Grid1.Visible = True
' separa_polizas

'verifica_polizas

final:

barra.Visible = False
msg.Visible = False
End Sub

Private Sub btnload2_Click()
On Error Resume Next
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long

n$ = ""
r$ = MsgBox("Do you wish to load all the data from Excel datasheet to this grid?", 4, "Confirm action")
If r$ = "7" Then Exit Sub



barra.Visible = True
barra.Min = 1
Grid1.Visible = False
barra.Max = 5000
lblmsg.Caption = "Loading the information from Excel..."
lblmsg2.Caption = ""
openforms = DoEvents

cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\pagos"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  If n$ = "" Then
     barra.Visible = False
     Exit Sub
  End If
  
msg.Visible = True

'btn_NB.Visible = False

contador = 0
Erase recibos$
Grid1.Clear

inicio:

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing


'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(FileName:=n$, ReadOnly:=True)



' Get the first worksheet.
 
  Set xlHoja = xlApp.Worksheets(1)
 

 ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.Row
    ultimacolumnax = ActiveCell.Column
    

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

If ultimafilax = 0 Then
  lngUltimaFila = 5000
Else
  lngUltimaFila = ultimafilax
End If
   

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_NB
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_NB = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), xlHoja.Cells(lngUltimaFila, 22))   ' cambie 10 por 19

Grid1.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid1.Rows = lngUltimaFila + 2
Grid1.cols = 23

cont = 0
linea_vacia = 0
For t = 1 To Grid1.Rows - 2
  
  barra.Value = t
  openforms = DoEvents
  
  Grid1.Row = t
  If t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
   
    If t > 0 Then
     cont = cont + 1
     Grid1.Col = 0
     Grid1.Text = cont
    End If
    
  End If
  
  Grid1.Row = t - 1
  veces = 0
  For Y = 1 To 18
   Grid1.Col = Y
   Grid1.Text = varMatriz(t, Y)
   If Grid1.Text = "" Then
     veces = veces + 1
     If veces >= 18 Then
        linea_vacia = linea_vacia + 1
     End If
   End If
  Next Y
  If linea_vacia >= 5 Then
     Exit For
  End If
Next t
   




Grid1.Rows = cont  '+ 2


'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit

'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing






enca_grid1

Grid1.Visible = True
' separa_polizas

'verifica_polizas

final:
barra.Visible = False
msg.Visible = False
End Sub

Private Sub btnmerge_Click()
On Error Resume Next


ultima_fila = grid6.Rows - 1
'cambia_filas
Dim tt As Single
grid6.Rows = grid6.Rows + grid4.Rows - 1
'grid6.Rows = 100

For t = 1 To grid4.Rows - 1
   grid4.Row = t
   tt = (ultima_fila) + t
   grid6.Row = tt
   For Y = 1 To grid4.cols - 1
      grid4.Col = Y
      grid6.Col = Y
      grid6.Text = grid4.Text
   Next Y
Next t

' crea el grid4 con las filas igual que el grid6
grid4.Clear

grid4.Rows = 0
For i = 0 To grid6.Rows - 1
  mydata = vbTab
  grid4.AddItem mydata
Next i

enca_grid4



For t = 0 To grid6.Rows - 1
   grid4.Row = t
   grid6.Row = t
   For Y = 1 To grid6.cols - 1
      grid4.Col = Y
      grid6.Col = Y
      grid4.Text = grid6.Text
   Next Y
Next t





suma = 0
For t = 1 To grid4.Rows - 1
   
   grid4.Row = t
   grid4.Col = 0
   grid4.Text = t
   
   grid4.Col = 5
   cant = Val(grid4.Text)
   
   suma = suma + cant
Next t

   
lbltotal4.Caption = Format(grid4.Rows - 1, "###,###0")
lblamount_total.Caption = Format(suma, "$###,##0.00")

carga_lista


'btnmerge.Enabled = False

For t = 0 To List2.ListCount - 1
  List2.Selected(t) = True
Next t


' grid6.Clear

MsgBox "The grids were merged", 64, "Attention"

btnmerge.Enabled = False

End Sub

Private Sub btnmes_Click(Index As Integer)
On Error Resume Next
mes_actual = Index + 1


Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dias_actual = 31
Case 4, 6, 9, 11
  dias_actual = 30
Case 2
   cant = (ano_actual / 4)
   residuo = cant - Int(cant)
   If residuo = 0 Then
      dias_actual = 29
   Else
      dias_actual = 28
   End If
End Select


If mes_actual > 1 Then

  fecha_rango1$ = Format(ano_actual, "00") + "-" + Format(mes_actual - 1, "00") + "-24"
  'fecha_rango1$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-01"
  fecha_rango2$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-" + Format(dias_actual, "00")
  
Else

  fecha_rango1$ = Format(ano_actual - 1, "00") + "-" + "12" + "-24"
  'fecha_rango1$ = Format(ano_actual, "00") + "-" + "01" + "-01"
  fecha_rango2$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-" + Format(dias_actual, "00")


End If


End Sub



Private Sub btnok_Click()
On Error Resume Next




'If Check2.Value = 1 Then
  
  If grabado = 0 Then
     r$ = MsgBox("You have not saved the last changes made. Do you want to save your work?", 4, "Attention")
     If r$ = "6" Then
       Check2.Value = 0
       graba_sql1
       graba_sql2
     End If
  End If

'End If


base.Close
    
End
End Sub



Private Sub btnrecalcula_Click()


End Sub

Private Sub btnremove_note_Click()
On Error Resume Next
If lblrow.Caption = "" Then
  MsgBox "There is not a row selected. Choose a row where you want to remove the comment.", 16, "Attention"
  Exit Sub
End If


If op_rem(1).Value = True Then

Grid1.Row = lblrow.Caption
Grid1.Col = 15
Grid1.Text = ""

Else

Grid3.Row = lblrow.Caption
Grid3.Col = 15
Grid3.Text = ""

End If

txtnota.Text = ""
lblrow.Caption = ""
grabado = 0
End Sub

Private Sub btnrevisado_Click()
On Error Resume Next
If Grid3.Rows <= 1 Then Exit Sub

r$ = MsgBox("Do you wish to transfer the policies marked to the lower grid?", 4, "Confirm action")
If r$ = "7" Then Exit Sub

Grid1.Visible = False
Grid3.Visible = False

    

grid2.Clear
grid2.Rows = Grid3.Rows
grid2.cols = Grid3.cols

cont = 0
For t = 1 To Grid3.Rows - 1
   Grid3.Row = t
   Grid3.Col = 11
   p$ = Grid3.Text
   

   
   If p$ = "Ok" Then
      Grid1.Rows = Grid1.Rows + 1
      Grid1.Row = Grid1.Rows - 1
      For Y = 1 To Grid1.cols
         Grid3.Col = Y
         Grid1.Col = Y
         Grid1.Text = Grid3.Text
      Next Y
   Else
     cont = cont + 1
     grid2.Row = cont
      For Y = 1 To Grid3.cols
         Grid3.Col = Y
         grid2.Col = Y
         grid2.Text = Grid3.Text
      Next Y
     
   End If
   
   
   
Next t
        
'Grid3.FixedRows = 1
'Grid3.FixedCols = 1

        
' borra las filas con "---"

grid2.Rows = cont + 1
Grid3.Clear
Grid3.Rows = cont + 1


For t = 1 To grid2.Rows - 1
    grid2.Row = t
    Grid3.Row = t
    
    Grid3.Col = 0
    Grid3.Text = Format(t, "####0")

    For Y = 1 To grid2.cols - 1
       grid2.Col = Y
      Grid3.Col = Y
      Grid3.Text = grid2.Text
    Next Y
 Next t

enca_grid1
enca_grid3





For t = 1 To Grid3.Rows - 1
  Grid3.Row = t
  Grid3.Col = 0
  Grid3.Text = Format(t, "####0")
Next t


For t = 1 To Grid1.Rows - 1
  Grid1.Row = t
  Grid1.Col = 0
  Grid1.Text = Format(t, "####0")
Next t


Grid1.Visible = True
Grid3.Visible = True


lbltotal2.Caption = Format(Grid1.Rows - 1, "###,##0")
lbltotal1.Caption = Format(Grid3.Rows - 1, "###,##0")
lbltotal5.Caption = Format(grid5.Rows - 1, "###,##0")

End Sub



Private Sub btnright_Click()
On Error Resume Next


valor_barra = valor_barra + 1
If valor_barra > 31 Then valor_barra = 31



Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dias_actual = 31
Case 4, 6, 9, 11
  dias_actual = 30
Case 2
   cant = (ano_actual / 4)
   residuo = cant - Int(cant)
   If residuo = 0 Then
      dias_actual = 29
   Else
      dias_actual = 28
   End If
End Select


If valor_barra >= dias_actual Or (valor_barra + rango_dia) >= dias_actual Then
   op_day(0).Value = True
   valor_barra = dias_actual
   rango_dia = 0
   
End If


  f1$ = Format(mes_actual, "00") + "/" + Format(valor_barra, "00") + "/" + Format(ano_actual, "0000")
  f2$ = Format(mes_actual, "00") + "/" + Format(valor_barra + rango_dia, "00") + "/" + Format(ano_actual, "0000")



txtfecha(0).Text = f1$
txtfecha(1).Text = f2$

lblfecha1.Caption = txtfecha(0).Text
lblfecha2.Caption = txtfecha(1).Text


btnbusca_registro_Click

carga_lista

calcula_total_multiple
End Sub

Private Sub btnsearch2_Click()
On Error Resume Next
If op_busca2(0).Value = True Then
  If txtcantidad2.Text = "" Then Exit Sub
  If chk_exacto.Value = True Or chk_exacto.Value = 1 Then
     cant = Val(txtcantidad2.Text)
  Else
     cant = Int(Val(txtcantidad2.Text))
  End If
  
  
  
  existe = 0
  For t = 1 To Grid1.Rows - 1
     Grid1.Row = t
     Grid1.Col = 10
     cantidad_correcta = Val(Format(Grid1.Text, "000000.00"))
     
     
     'Grid1.Col = 3
     'debito = Val(Grid1.Text)
     'Grid1.Col = 4
     'credito = Val(Grid1.Text)
     
     'cantidad = debito + credito
     
     If chk_exacto.Value = False Then
       'debito = Int(debito)
       'credito = Int(credito)
       cantidad_correcta = Int(cantidad_correcta)
     End If
     
     
     
     If cant = cantidad_correcta Then
       existe = 1
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 10
            .ColSel = 9
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the amount of " + Grid1.Text + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
     End If
     
  Next t
  
  
  
  
ElseIf op_busca2(1).Value = True Then

  If txtcantidad2.Text = "" Then Exit Sub
  
  p$ = txtcantidad2.Text
  existe = 0
  For t = 1 To Grid1.Rows - 1
     Grid1.Row = t
     Grid1.Col = 8
     poliza$ = UCase(Grid1.Text)
     
     
     If chk_exacto.Value = True Then
     
      
      If UCase(p$) = UCase(poliza$) Then
       existe = 1
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
      
       
       
      End If
     Else
     
      If UCase(p$) = Left(UCase(poliza$), Len(p$)) Then
       existe = 1
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
       
      End If
     
     
     End If
     
     
     
     
       
      
      
  Next t
  
  
  
ElseIf op_busca2(2).Value = True Then
  
  
   p$ = "V O I D"
  existe = 0
  For t = 1 To Grid1.Rows - 1
     Grid1.Row = t
     Grid1.Col = 15
     Comment$ = Grid1.Text
     
     
     If chk_exacto.Value = 1 Then
     
      
      If UCase(p$) = UCase(Comment$) Then
       existe = 1
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
      
       
       
      End If
      
     Else
     
      If UCase(p$) = Left(UCase(Comment$), Len(p$)) Then
       existe = 1
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 8
            .ColSel = 10
            '.CellBackColor = &HC0FFFF
            .TopRow = .Row
       End With
       
       r$ = MsgBox("Is this the policy #" + p$ + "?", 4, "Attention")
       
       If r$ = "6" Then
             Exit For
       End If
       
       
      End If
     
     
     End If
     
     
     
     
       
      
      
  Next t
  

End If




MsgBox "End of search", 64, "Attention"

End Sub

Private Sub btnsepara_cantidades_Click()
On Error Resume Next
RichTextBox1.Font.Name = "arial"
RichTextBox1.Font.Size = 10



Dim arr() As String
    Dim i As Integer
    arr = Split(RichTextBox1.Text, vbCrLf)
    cont = 0
    
    For i = 0 To UBound(arr)
       a = 0
       For k = Len(arr(i)) To 1 Step -1
          a = a + 1
          If Mid$(arr(i), k, 1) = Space(1) Then
              
              cont = cont + 1
              
          End If
          
          
          If cont >= 3 Then
            arr(i) = LTrim(RTrim(Right(arr(i), a)))
            pos = InStr(1, arr(i), " ")
            arr(i) = Format(Left(arr(i), pos - 1), "###0.00") + vbCrLf
            a = 0
            cont = 0
            Exit For
          End If
          
       Next k
       
       
    Next i

    
    
    RichTextBox1.Text = ""
    
    
    
    For i = 0 To UBound(arr)
        
        existe = 0
        For Y = 0 To List2.ListCount - 1
           If Val(Left(List2.List(Y), 10)) = Val(arr(i)) Then
               List2.Selected(Y) = True
               existe = 1
               Exit For
           End If
        Next Y
        
        If existe = 0 Then
           RichTextBox1.Text = RichTextBox1.Text + arr(i)
        Else
           RichTextBox1.Text = RichTextBox1.Text + " OK- " + arr(i)
        End If
    
    Next i
    
    
    lbltotal_lista.Caption = List2.ListCount

    
    

   
    
   
    
    
    calcula_total_multiple
    
  
    lbltotal_lista.Caption = List2.ListCount

End Sub

Private Sub btnsepara_cantidades2_Click()
On Error Resume Next
RichTextBox1.Font.Name = "arial"
RichTextBox1.Font.Size = 10



Dim arr() As String
    Dim i As Integer
    
    
If Picture1.Visible = True Then
  btnbusca_registro_Click
  carga_lista
  calcula_total_multiple
End If
    
    
    
    arr = Split(RichTextBox1.Text, vbCrLf)
    cont = 0
    
 If Check1.Value = True Then
    
    For i = 0 To UBound(arr)
       a = 0
       For k = Len(arr(i)) To 1 Step -1
          a = a + 1
          If Mid$(arr(i), k, 1) = "$" Then
              
              cont = cont + 1
              
          End If
          
          
          If cont >= 1 Then
            arr(i) = LTrim(RTrim(Right(arr(i), a)))
            pos = InStr(1, arr(i), "$")
            arr(i) = Format(Right(arr(i), Len(arr(i)) - 1), "###0.00") + vbCrLf
            a = 0
            cont = 0
            Exit For
          End If
          
       Next k
       
       
    Next i

 Else
 
    For i = 0 To UBound(arr)
            arr(i) = LTrim(RTrim(arr(i)))
            arr(i) = Format(arr(i), "###0.00") + vbCrLf
       
    Next i

 End If
    
    
    
    
    
    
    RichTextBox1.Text = ""
    
    
    
    For i = 0 To UBound(arr)
        
        existe = 0
        For Y = 0 To List2.ListCount - 1
           If Val(Left(List2.List(Y), 10)) = Val(arr(i)) Then
               List2.Selected(Y) = True
               existe = 1
               Exit For
           End If
        Next Y
        
        If existe = 0 Then
           RichTextBox1.Text = RichTextBox1.Text + arr(i)
        Else
           RichTextBox1.Text = RichTextBox1.Text + " OK- " + arr(i)
        End If
    
    Next i
    
    
    lbltotal_lista.Caption = List2.ListCount

    
    

   
    
   
    
    
    calcula_total_multiple
    
  
    lbltotal_lista.Caption = List2.ListCount

End Sub


Private Sub btnsepara_polizas_Click()
On Error Resume Next
msg.Visible = True
msg.Refresh

lblmsg.Caption = "Separating the policies from the field..."

lblmsg.Refresh
Grid1.Visible = False
Grid1.Refresh
separa_polizas
Grid1.Visible = True
final:
msg.Visible = False

End Sub

Private Sub btnsort1_Click()
On Error Resume Next



msg.Visible = True
  
  lblmsg.Caption = "Sorting all the information..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  posicion = Val(Right(cbosort1.List(cbosort1.ListIndex), 5))
  
  If posicion = 6 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 6
       Grid1.Text = Format(Grid1.Text, "yyyymmdd")
    Next t
    
  End If
  
  
  If posicion = 12 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 12
       Grid1.Text = Format(Grid1.Text, "yyyymmdd")
    Next t
    
  End If
  
  
  
  If posicion = 3 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 3
       Grid1.Text = Format(Grid1.Text, " 00000.00")
       
    Next t
    
  End If
  
  
  
  If posicion = 10 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 10
       Grid1.Text = Format(Grid1.Text, " 00000.00")
    Next t
    
  End If
  
  
  
  
     Grid1.Col = posicion
     Grid1.Sort = flexSortGenericAscending
     
  ' asigna numeracion nueva por fila
  For t = 1 To Grid1.Rows - 1
    Grid1.Row = t
    Grid1.Col = 0
    Grid1.Text = t
  Next t
  
  
  
  If posicion = 6 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 6
       r$ = Grid1.Text
       f$ = Mid$(r$, 5, 2) + "/" + Right(r$, 2) + "/" + Left(r$, 4)
       Grid1.Text = f$
       
       
    Next t
    
  End If
  
  
  
  If posicion = 12 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 12
       r$ = Grid1.Text
       f$ = Mid$(r$, 5, 2) + "/" + Right(r$, 2) + "/" + Left(r$, 4)
       Grid1.Text = f$
       
       
    Next t
    
  End If
  
  
  
  
  If posicion = 3 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 3
       Grid1.Text = Format(Grid1.Text, " ####0.00")
       
    Next t
    
  End If
  
  
  
  If posicion = 10 Then
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       Grid1.Col = 10
       Grid1.Text = Format(Grid1.Text, " ####0.00")
    Next t
    
  End If
  
  
  
  
  msg.Visible = False
  
End Sub

Private Sub btnsort3_Click()
On Error Resume Next



msg.Visible = True
  
  lblmsg.Caption = "Sorting all the information..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  posicion = Val(Right(cbosort3.List(cbosort3.ListIndex), 5))
  
  If posicion = 6 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 6
       Grid3.Text = Format(Grid3.Text, "yyyymmdd")
    Next t
    
  End If
  
  
  If posicion = 12 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 12
       Grid3.Text = Format(Grid3.Text, "yyyymmdd")
    Next t
    
  End If
  
  
  
  
  If posicion = 3 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 3
       Grid3.Text = Format(Grid3.Text, " 000000.00")
    Next t
    
  End If
  
  
  
  If posicion = 10 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 10
       Grid3.Text = Format(Grid3.Text, " 000000.00")
    Next t
    
  End If
  
  
  
  
  
  
     Grid3.Col = posicion
     Grid3.Sort = flexSortGenericAscending
  
  For t = 1 To Grid3.Rows - 1
    Grid3.Row = t
    Grid3.Col = 0
    Grid3.Text = t
  Next t
  
  
  
  If posicion = 6 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 6
       r$ = Grid3.Text
       f$ = Mid$(r$, 5, 2) + "/" + Right(r$, 2) + "/" + Left(r$, 4)
       Grid3.Text = f$
       
       
    Next t
    
  End If
  
  
  
  If posicion = 12 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 12
       r$ = Grid3.Text
       f$ = Mid$(r$, 5, 2) + "/" + Right(r$, 2) + "/" + Left(r$, 4)
       Grid3.Text = f$
       
       
    Next t
    
  End If
  
  
  
  
   
  If posicion = 3 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 3
       Grid3.Text = Format(Grid3.Text, " ####0.00")
       
    Next t
    
  End If
  
  
  
  If posicion = 10 Then
    For t = 1 To Grid3.Rows - 1
       Grid3.Row = t
       Grid3.Col = 10
       Grid3.Text = Format(Grid3.Text, " ####0.00")
    Next t
    
  End If
  
  
  
   
  
  
  msg.Visible = False
  

  

End Sub

Private Sub btnsort4_Click()
On Error Resume Next
grid4.Col = 5
     grid4.Sort = flexSortGenericAscending
  
  For t = 1 To grid4.Rows - 1
    grid4.Row = t
    grid4.Col = 0
    grid4.Text = t
  Next t
  
  
End Sub


Private Sub btnsql_Click()
On Error Resume Next
 


r$ = MsgBox("Do you want to save all data?", 4, "Attention")
  If r$ = "7" Then
    Exit Sub
  End If

 
 
 
graba_sql1

graba_sql2


carga_grid5
grabado = 1

End Sub

Private Sub btntransfiere_Click()
On Error Resume Next

If Grid1.Rows <= 1 Then Exit Sub

r$ = MsgBox("Do you wish to transfer the policies marked as doubtful to the upper grid?", 4, "Confirm action")
If r$ = "7" Then Exit Sub

Grid1.Visible = False
Grid3.Visible = False

msg.Visible = True
msg.Refresh

lblmsg.Caption = "Transfering the policies to the upper grid..."
lblmsg.Refresh



grid2.Clear
grid2.Rows = Grid1.Rows
grid2.cols = Grid1.cols

cont = 0
For t = 1 To Grid1.Rows - 1
   Grid1.Row = t
   Grid1.Col = 11
   p$ = Grid1.Text
   

   
   If p$ = "---" Then
      Grid3.Rows = Grid3.Rows + 1
      Grid3.Row = Grid3.Rows - 1
      For Y = 1 To Grid1.cols - 1
         Grid1.Col = Y
         Grid3.Col = Y
         Grid3.Text = Grid1.Text
      Next Y
   Else
     cont = cont + 1
     grid2.Row = cont
      For Y = 1 To Grid1.cols - 1
         Grid1.Col = Y
         grid2.Col = Y
         grid2.Text = Grid1.Text
      Next Y
     
   End If
   
   
   
Next t
        
'Grid3.FixedRows = 1
'Grid3.FixedCols = 1

        
' borra las filas con "---"

grid2.Rows = cont + 1
Grid1.Clear
Grid1.Rows = cont + 1


For t = 1 To grid2.Rows - 1
    grid2.Row = t
    Grid1.Row = t
    
    Grid1.Col = 0
    Grid1.Text = Format(t, "####0")

    For Y = 1 To grid2.cols - 1
      grid2.Col = Y
      Grid1.Col = Y
      Grid1.Text = grid2.Text
    Next Y
 Next t

enca_grid1
enca_grid3




For t = 1 To Grid3.Rows - 1
  Grid3.Row = t
  Grid3.Col = 0
  Grid3.Text = Format(t, "####0")
Next t


Grid1.Visible = True
Grid3.Visible = True

msg.Visible = False

lbltotal2.Caption = Format(Grid1.Rows - 1, "###,##0")
lbltotal1.Caption = Format(Grid3.Rows - 1, "###,##0")
lbltotal5.Caption = Format(grid5.Rows - 1, "###,##0")

End Sub

Private Sub btntransfiere_grid4_a_grid3_Click()
On Error Resume Next

ya_existen_mas = 0

If Val(Format(lbl_diferencia.Caption, "00000.00")) > 0 And chk_multiple.Value = True Then
  r$ = MsgBox("It has a difference greater than zero..." + Chr$(13) + "Do you want to continue?", 4, "Attention")
  If r$ = "7" Then
    Exit Sub
  End If
End If




If lblfila_grid4.Text = "" And chk_multiple.Value = False Then
   MsgBox "You have not selected the row number", 64, "Attention"
   chk_multiple.Value = False
   Picture1.Visible = False
   encapsula.Visible = False
  ' Picture2.Visible = False

   Exit Sub
End If


If Grid3.Rows <= 1 Or grid4.Rows <= 1 Then
  chk_multiple.Value = False
  Picture1.Visible = False
   encapsula.Visible = False
  ' Picture2.Visible = False
  Exit Sub
End If

'If txtfila.Text = "" Then Exit Sub

fila_grid3 = Val(txtfila.Text)
If fila_grid3 = 0 Then
   MsgBox "You have not selected the row in grid 2", 64, "Attention"
   chk_multiple.Value = False
   Picture1.Visible = False
   encapsula.Visible = False
  ' Picture2.Visible = False
   Exit Sub
End If

Grid3.Col = 11
If Grid3.Text = "Ok" Then
   MsgBox "You already have information in that row of grid 2", 16, "Attention"
   chk_multiple.Value = False
   Picture1.Visible = False
   encapsula.Visible = False
 ' Picture2.Visible = False
   Exit Sub
End If




grid4.Row = 2
grid4.Col = 0
If Val(grid4.Text) = 0 Then
   chk_multiple.Value = False
   Picture1.Visible = False
   encapsula.Visible = False
  ' Picture2.Visible = False
   Exit Sub
End If









If chk_multiple.Value = False Then

    grid4.Row = lblfila_grid4.List(lblfila_grid4.ListIndex)

' VERIFICA EN LA CUADRICULA1 QUE NO EXISTA

    Dim sSelect As String
    Dim Rs As ADODB.Recordset

    grid4.Col = 10
    poliza$ = UCase(grid4.Text)

    grid4.Col = 5
    cantidad$ = grid4.Text

    grid4.Col = 1
    recibo$ = grid4.Text

    grid4.Col = 4
    fecha_recibo$ = Format(grid4.Text, "mm/dd/yyyy")

    msg.Visible = True
    lblmsg.Caption = "Verifying that it is not duplicated in grid # 1 and grid# 2"
    lblmsg2.Caption = ""
    msg.Refresh
openforms = DoEvents
' verifica en grid1

    existe = 0

    For t = 1 To Grid1.Rows - 1
          Grid1.Row = t
          Grid1.Col = 8
          poliza_x$ = Grid1.Text
  
          Grid1.Col = 10
          cantidad_x$ = Format(Grid1.Text, "######0.00")

          Grid1.Col = 9
          recibo_x$ = Grid1.Text
  
          Grid1.Col = 12
          fecha_recibo_x$ = Format(Grid1.Text, "mm/dd/yyyy")
  
          Grid1.Col = 2
          multiple_recibo$ = Grid1.Text
  
  
          If Format(cantidad_x$, "####0.00") = Format(cantidad$, "####0.00") And recibo_x$ = recibo$ And poliza_x$ = poliza$ And fecha_recibo$ = fecha_recibo_x$ Then
               existe = 1
               Exit For
          End If
  
    Next t
    
    

    If existe = 1 And multiple_recibo$ = "" Then
        MsgBox "That amount has already been marked like <found> in grid#1", 64, "Attention"
        msg.Visible = False
        Exit Sub
    End If

    existe = 0

    For t = 1 To Grid3.Rows - 1
        Grid3.Row = t
        Grid3.Col = 8
        poliza_x$ = Grid3.Text
  
        Grid3.Col = 10
        cantidad_x$ = Format(Grid3.Text, "######0.00")

        Grid3.Col = 9
        recibo_x$ = Grid3.Text
  
        Grid3.Col = 12
        fecha_recibo_x$ = Format(Grid3.Text, "mm/dd/yyyy")
  
        Grid3.Col = 2
        multiple_recibo$ = Grid3.Text
  
        Grid3.Col = 11
        estatus$ = Grid3.Text
    
        If Format(cantidad_x$, "####0.00") = Format(cantidad$, "####0.00") And recibo_x$ = recibo$ And poliza_x$ = poliza$ And fecha_recibo$ = fecha_recibo_x$ Then
           existe = 1
           Exit For
        End If
  

    Next t


    If existe = 1 And multiple_recibo$ = "" And estatus$ <> "---" Then
          MsgBox "That amount has already been marked like <found> in grid#2", 64, "Attention"
          msg.Visible = False

          Exit Sub
    End If


    fila = Val(txtfila.Text)
    Grid3.Row = fila

    grid4.Row = lblfila_grid4.Text

    Grid3.Col = 8 ' poliza
    grid4.Col = 10
    Grid3.Text = grid4.Text
    poliza$ = UCase(grid4.Text)

    Grid3.Col = 9 ' recibo
    grid4.Col = 1
    Grid3.Text = grid4.Text
    recibo$ = grid4.Text

    If chk_multiple.Value = 1 Or chk_multiple.Value = True Then
       Grid3.Col = 2
       Grid3.Text = recibo$
    End If

    Grid3.Col = 10 ' amount
    grid4.Col = 5
    cantidad = Val(grid4.Text)
    Grid3.Text = Format(grid4.Text, "#####0.00")


    Grid3.Col = 11 ' verified
    Grid3.Text = "Ok"

    Grid3.Col = 12 ' date created
    grid4.Col = 4
    Grid3.Text = Format(grid4.Text, "mm/dd/yyyy")

    Grid3.Col = 13 ' idcustomer
    grid4.Col = 11
    Grid3.Text = grid4.Text

    Grid3.Col = 14 ' company
    grid4.Col = 9
    Grid3.Text = grid4.Text

    Grid3.Col = 15 ' Comment
    Grid3.Text = "Added and reviewed by Joselyn"


    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idpolizahdr$ = ""
    sSelect = "select idpolicieshdr from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idpolizahdr$ = Rs(0)   ' CORRECTO
                         
    Rs.Close
          
    If idpolizahdr$ = "" Then GoTo salta
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    IDReceiptHDR$ = ""
               
    sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + txtfecha(0).Text + "' and date<'" + txtfecha(1).Text + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IDReceiptHDR$ = Rs(0)  'correcto
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idreceiptdtl$ = ""
               
    sSelect = "select IdReceiptDTL from Receiptsdtl where idpolicieshdr='" + idpolizahdr$ + "' and datecreated>='" + txtfecha(0).Text + "' and datecreated<'" + txtfecha(1).Text + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idreceiptdtl$ = Rs(0)
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
        
    sSelect = "select ins.CompanyName " & _
            "FROM   ReceiptsHDR  rechdr " & _
            "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
            "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
            "Where rechdr.IdReceiptHDR = '" + recibo$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    compania$ = Rs(0)
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idcompany$ = ""
    sSelect = "select idcompany from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idcompany$ = Rs(0)
                         
    Rs.Close



    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idprogram$ = ""
    sSelect = "select idprogram from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idprogram$ = Rs(0)
                         
    Rs.Close



    Set Rs = New ADODB.Recordset
    'Checa_status
   
    programname$ = ""
    sSelect = "select programname from programscatalog where idcompany='" + idcompany$ + "' and idprogram='" + idprogram$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    programname$ = Rs(0)
                         
    Rs.Close



salta:

    Grid3.Col = 16 ' program name
    Grid3.Text = programname$

    Grid3.Col = 17 ' idprogram
    Grid3.Text = idprogram$

    Grid3.Col = 18 ' idreceiptDTL
    Grid3.Text = idreceiptdtl$

    Grid3.Col = 19 ' idpolizaHDR
    Grid3.Text = idpolizahdr$

    Grid3.Col = 20 ' idcompany
    Grid3.Text = idcompany$

    Grid3.Col = 22
    Grid3.Text = "0"


    chk_multiple.Value = False
    Picture1.Visible = False
    encapsula.Visible = False
  '  Picture2.Visible = False
    msg.Visible = False
    grabado = 0
    
    
    
Else

    ' *******************************  MULTIPLES  ***********************************************************************
 
  Dim line1$(7)
  fila_actual = Grid3.Row
  ultima_fila_grid3 = Grid3.Rows
  Grid3.Rows = Grid3.Rows + (grid4.Rows - 2)
    
  filax = ultima_fila_grid3 - 1
  
  Grid3.Col = 1
  line1$(1) = Grid3.Text

      Grid3.Col = 2
      line1$(2) = Grid3.Text

      Grid3.Col = 3
      line1$(3) = Grid3.Text

      Grid3.Col = 4
      line1$(4) = Grid3.Text

      Grid3.Col = 5
      line1$(5) = Grid3.Text

      Grid3.Col = 6
      line1$(6) = Grid3.Text

      Grid3.Col = 7
      line1$(7) = Grid3.Text
      
      
      
      
    
  For Y = ultima_fila_grid3 To Grid3.Rows
      
      filax = filax + 1
      Grid3.Row = filax
      
      
      Grid3.Col = 0
      Grid3.Text = Y
      
      Grid3.Col = 1
      Grid3.Text = line1$(1)

      Grid3.Col = 2
      Grid3.Text = line1$(2)

      Grid3.Col = 3
      Grid3.Text = line1$(3)

      Grid3.Col = 4
      Grid3.Text = line1$(4)

      Grid3.Col = 5
      Grid3.Text = line1$(5)

      Grid3.Col = 6
      Grid3.Text = line1$(6)

      Grid3.Col = 7
      Grid3.Text = line1$(7)
      
      
  Next Y
  
  
    
    
    
    
    
  For z = 1 To grid4.Rows - 1
    grid4.Row = z
    
    
    
      
    

' VERIFICA EN LA CUADRICULA1 QUE NO EXISTA

    'Dim sSelect As String
    'Dim Rs As ADODB.Recordset

    grid4.Col = 10
    poliza$ = UCase(grid4.Text)

    grid4.Col = 5
    cantidad$ = grid4.Text

    grid4.Col = 1
    recibo$ = grid4.Text

    grid4.Col = 4
    fecha_recibo$ = Format(grid4.Text, "mm/dd/yyyy")

    msg.Visible = True
    lblmsg.Caption = "Verifying that it is not duplicated in grid # 1 and grid# 2"
    lblmsg2.Caption = ""
    msg.Refresh
openforms = DoEvents
' verifica en grid1

    existe = 0

    For t = 1 To Grid1.Rows - 1
          Grid1.Row = t
          Grid1.Col = 8
          poliza_x$ = Grid1.Text
  
          Grid1.Col = 10
          cantidad_x$ = Format(Grid1.Text, "######0.00")

          Grid1.Col = 9
          recibo_x$ = Grid1.Text
  
          Grid1.Col = 12
          fecha_recibo_x$ = Format(Grid1.Text, "mm/dd/yyyy")
  
          multiple_recibo$ = ""
          Grid1.Col = 2
          multiple_recibo$ = Grid1.Text
  
  
          If Format(cantidad_x$, "####0.00") = Format(cantidad$, "####0.00") And recibo_x$ = recibo$ And poliza_x$ = poliza$ And fecha_recibo$ = fecha_recibo_x$ Then
               existe = 1
               ya_existen_mas = 1
               Exit For
          End If
  
    Next t
    
    

    If existe = 1 And multiple_recibo$ = "" Then
        MsgBox "That amount has already been marked like <found> in grid#1" + Chr$(13) + Chr$(13) + "Amount: " + cantidad$ + Chr$(13) + "Receipt: " + recibo$ + Chr$(13) + "Policy# " + poliza$ + Chr$(13) + "Date: " + fecha_recibo$, 64, "Attention"
        msg.Visible = False
        Grid3.Rows = ultima_fila_grid3 - 1
        Exit Sub
    End If
    
    
    If existe = 1 And multiple_recibo$ <> "" Then
       'Grid3.Rows = Grid3.Rows - 1
       ya_existen_mas = 1
       Grid3.RemoveItem (Grid3.Rows - 1)
       GoTo final
    End If
    
    
     
    
    
    

    existe = 0

    For t = 1 To Grid3.Rows - 1
        Grid3.Row = t
        Grid3.Col = 8
        poliza_x$ = Grid3.Text
  
        Grid3.Col = 10
        cantidad_x$ = Format(Grid3.Text, "######0.00")

        Grid3.Col = 9
        recibo_x$ = Grid3.Text
  
        Grid3.Col = 12
        fecha_recibo_x$ = Format(Grid3.Text, "mm/dd/yyyy")
  
        Grid3.Col = 2
        multiple_recibo$ = Grid3.Text
  
        Grid3.Col = 11
        estatus$ = Grid3.Text
    
        If Format(cantidad_x$, "####0.00") = Format(cantidad$, "####0.00") And recibo_x$ = recibo$ And poliza_x$ = poliza$ And fecha_recibo$ = fecha_recibo_x$ Then
           existe = 1
           Exit For
        End If
  

    Next t


    If existe = 1 And estatus$ <> "---" Then  ' And multiple_recibo$ = "" Then
         ' MsgBox "That amount has already been marked like <found> in grid#2", 64, "Attention"
         ' msg.Visible = False
         ' Grid3.Rows = ultima_fila_grid3 - 1
         ' Exit Sub
    End If


    
    If z = 1 Then
        
        grid4.Row = z  'lblfila_grid4.Text
        fila = Val(txtfila.Text)
        Grid3.Row = fila
        
    
    Else
    
        If ya_existen_mas = 1 Then
           
           Grid3.Row = fila_actual  ' ultima_fila_grid3
           ultima_fila_grid3 = fila_actual  'ultima_fila_grid3 + 1
           grid4.Row = z
        
        Else
        
           Grid3.Row = ultima_fila_grid3
           ultima_fila_grid3 = ultima_fila_grid3 + 1
           grid4.Row = z
        
        
        End If
        
    End If
        

    Grid3.Col = 8 ' poliza
    grid4.Col = 10
    Grid3.Text = grid4.Text
    poliza$ = UCase(grid4.Text)

    Grid3.Col = 9 ' recibo
    grid4.Col = 1
    Grid3.Text = grid4.Text
    recibo$ = grid4.Text

    If chk_multiple.Value = 1 Or chk_multiple.Value = True Then
       Grid3.Col = 2
       Grid3.Text = recibo$
    End If

    Grid3.Col = 10 ' amount
    grid4.Col = 5
    cantidad = Val(grid4.Text)
    Grid3.Text = Format(grid4.Text, "#####0.00")


    Grid3.Col = 11 ' verified
    Grid3.Text = "Ok"

    Grid3.Col = 12 ' date created
    grid4.Col = 4
    Grid3.Text = Format(grid4.Text, "mm/dd/yyyy")

    Grid3.Col = 13 ' idcustomer
    grid4.Col = 11
    Grid3.Text = grid4.Text

    Grid3.Col = 14 ' company
    grid4.Col = 9
    Grid3.Text = grid4.Text

    Grid3.Col = 15 ' Comment
    Grid3.Text = "Added and reviewed by Joselyn"


    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idpolizahdr$ = ""
    sSelect = "select idpolicieshdr from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idpolizahdr$ = Rs(0)   ' CORRECTO
                         
    Rs.Close
          
    If idpolizahdr$ = "" Then GoTo salta
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    IDReceiptHDR$ = ""
               
    sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and datecreated>='" + txtfecha(0).Text + "' and datecreated<'" + txtfecha(1).Text + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IDReceiptHDR$ = Rs(0)  'correcto
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idreceiptdtl$ = ""
               
    sSelect = "select IdReceiptDTL from Receiptsdtl where idpolicieshdr='" + idpolizahdr$ + "' and datecreated>='" + txtfecha(0).Text + "' and datecreated<'" + txtfecha(1).Text + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idreceiptdtl$ = Rs(0)
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
        
    sSelect = "select ins.CompanyName " & _
            "FROM   ReceiptsHDR  rechdr " & _
            "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
            "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
            "Where rechdr.IdReceiptHDR = '" + recibo$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    compania$ = Rs(0)
                         
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idcompany$ = ""
    sSelect = "select idcompany from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idcompany$ = Rs(0)
                         
    Rs.Close



    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idprogram$ = ""
    sSelect = "select idprogram from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idprogram$ = Rs(0)
                         
    Rs.Close



    Set Rs = New ADODB.Recordset
    'Checa_status
   
    programname$ = ""
    sSelect = "select programname from programscatalog where idcompany='" + idcompany$ + "' and idprogram='" + idprogram$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    programname$ = Rs(0)
                         
    Rs.Close



salta2:

   If ya_existen_mas = 0 Then
    Grid3.Col = 2
    Grid3.Text = Format(z, "#0") + " of " + lbltotal4.Caption
   Else
    Grid3.Col = 2
    Grid3.Text = Format("1", "#0") + " of " + lbltotal4.Caption
   End If
   
   


    Grid3.Col = 16 ' program name
    Grid3.Text = programname$

    Grid3.Col = 17 ' idprogram
    Grid3.Text = idprogram$

    Grid3.Col = 18 ' idreceiptDTL
    Grid3.Text = idreceiptdtl$

    Grid3.Col = 19 ' idpolizaHDR
    Grid3.Text = idpolizahdr$

    Grid3.Col = 20 ' idcompany
    Grid3.Text = idcompany$

    Grid3.Col = 22
    Grid3.Text = "0"


    chk_multiple.Value = False
    Picture1.Visible = False
    encapsula.Visible = False
  '  Picture2.Visible = False
    msg.Visible = False
    grabado = 0
    
final:
   
 Next z
    
    
End If


 chk_multiple.Value = False
    Picture1.Visible = False
    encapsula.Visible = False
  '  Picture2.Visible = False
  
    btneraser_all_Click
    
    msg.Visible = False

    
End Sub

Private Sub btnverifica_polizas_Click()
On Error Resume Next

If Grid3.Rows > 1 Then
  r$ = MsgBox("You have data in the upper grid. Everything will be erased. Do you want to save the data?", 4, "Attention")
  If r$ = "6" Then
    btnexcel2_Click
  End If

End If

msg.Visible = True
msg.Refresh

lblmsg.Caption = "Verifying paid policies..."
lblmsg.Refresh
Grid1.Visible = False
Grid1.Refresh
Grid3.Visible = False
'

GoTo verifica1





fila_total = Grid1.Rows

grid6.Rows = Grid1.Rows
grid6.cols = Grid1.cols
grid6.Clear

For t = 1 To Grid1.Rows - 1
        Grid1.Row = t
        grid6.Row = t
   For Y = 1 To Grid1.cols - 1
       Grid1.Col = Y
       grid6.Col = Y
       grid6.Text = Grid1.Text
   Next Y
Next t



' comienza proceso
fila_inicial = 1
fila_final = 20
linea = 1

Do Until fila_final >= fila_total
   'Grid1.Clear
   For t = fila_inicial To fila_final
       Grid1.Row = t
       grid6.Row = linea
       linea = linea + 1
       For Y = 1 To grid6.cols - 1
           Grid1.Col = Y
           grid6.Col = Y
           Grid1.Text = grid6.Text
       Next Y
   Next t
   
   fila_inicial = fila_inicial + 20
   fila_final = fila_final + 20
   
   If fila_final > (grid6.Rows + 1) Then
       fila_final = (grid6.Rows + 1)
   End If
   
   verifica_polizas
   
   
   For z = 1 To Grid1.Rows - 1
     Grid1.Row = z
     For t = fila_inicial To fila_final
        grid6.Row = t
        For Y = 1 To Grid1.cols - 1
           Grid1.Col = Y
           grid6.Col = Y
           grid6.Text = Grid1.Text
        Next Y
     Next t
   Next z
      
   
   Grid1.Clear
   
Loop
   
       
       
Grid1.Clear
For t = 1 To grid6.Rows - 1
   grid6.Row = t
   Grid1.Row = t
   For Y = 1 To grid6.cols - 1
      grid6.Col = Y
      Grid1.Col = Y
      Grid1.Text = grid6.Text
   Next Y
Next t
   




verifica1:

verifica_polizas




 ASIGNA_UNOS



   carga_grid5

   enca_grid5


graba_sql1
graba_sql2


Grid1.Visible = True
Grid3.Visible = True
final:
msg.Visible = False


Grid1.Row = 1
Grid1.Col = 1

Grid3.Row = 1
Grid3.Col = 1

grid4.Row = 1
grid4.Col = 1

grabado = 0
txtfecha(0).Text = Format(fecha_rango1$, "mm/dd/yyyy")
txtfecha(1).Text = Format(fecha_rango2$, "mm/dd/yyyy")

lbltotal2.Caption = Format(Grid1.Rows - 1, "###,##0")
lbltotal1.Caption = Format(Grid3.Rows - 1, "###,##0")
lbltotal5.Caption = Format(grid5.Rows - 1, "###,##0")

btnverifica_polizas.Enabled = False

End Sub








Private Sub Calendar1_Click()
On Error Resume Next
txtfecha_cargada(calen1).Text = Calendar1.Value

If calen1 = 0 Then
  lblfecha1.Caption = txtfecha_cargada(calen1).Text
Else
  lblfecha2.Caption = txtfecha_cargada(calen1).Text
End If

If Picture1.Visible = True Then
  btnbusca_registro_Click
  carga_lista
  calcula_total_multiple
End If



Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
On Error Resume Next
txtfecha(calen2).Text = Calendar2.Value

If calen2 = 0 Then
  lblfecha1.Caption = txtfecha(calen2).Text
Else
  lblfecha2.Caption = txtfecha(calen2).Text
End If

If Picture1.Visible = True Then
  btnbusca_registro_Click
  carga_lista
  calcula_total_multiple
End If



Calendar2.Visible = False
End Sub

Private Sub cboyear_Click()
On Error Resume Next
ano_actual = cboyear.List(cboyear.ListIndex)

End Sub




Private Sub Check2_Click()
On Error Resume Next

If Check2.Value = 0 Then
  btnsql.Enabled = True
Else
  btnsql.Enabled = False
End If


End Sub

Private Sub chk_multiple_Click()
On Error Resume Next
If grid4.Rows <= 1 Or grid4.cols <= 5 Then
   chk_multiple.Value = False
   Exit Sub
End If

If Picture1.Visible = True Then
  btnbusca_registro_Click
  carga_lista
  'calcula_total_multiple
End If


Resetea_multiple
List2.Visible = False

Picture1.Visible = chk_multiple.Value
Picture2.Visible = chk_multiple.Value

'btnguardar.Caption = "Save"
'btnguardar.Enabled = True
'btnmerge.Enabled = True



pulgar.Visible = False
RichTextBox1.Text = ""
List2.Visible = chk_multiple.Value
encapsula.Visible = chk_multiple.Value
' Picture2.Visible = chk_multiple.Value
txtamount.Text = ""
txtcust_id.Text = ""
txtfirst_name.Text = ""
txtlast_name.Text = ""
txtpoliza.Text = ""
Grid3.Row = Val(txtfila.Text)
Grid3.Col = 6
fecha$ = Format(Grid3.Text, "mm/dd/yyyy")

dia$ = Mid$(fecha$, 4, 2)
mes$ = Left(fecha$, 2)
ano$ = Right(fecha$, 4)

mes_actual = Val(mes$)
ano_actual = Val(ano$)


d = Val(dia$) - 5
If d <= 0 Then d = 1
'fecha1$ = mes$ + "/" + Format(d, "00") + "/" + ano$
'fecha2$ = mes$ + "/" + Format(d + 1, "00") + "/" + ano$


'txtfecha(0).Text = fecha1$
'txtfecha(1).Text = fecha2$

valor_barra = d


If mes_actual = 2 Then
   dia_actual = 28
Else
   dia_actual = 30
End If


op_day(0).Value = True
op_day(3).Value = True
f1$ = Format(mes_actual, "00") + "/" + Format(1, "00") + "/" + Format(ano_actual, "0000")
f2$ = Format(mes_actual, "00") + "/" + Format(dia_actual, "00") + "/" + Format(ano_actual, "0000")

txtfecha(0).Text = f1$
txtfecha(1).Text = f2$

lblfecha1.Caption = txtfecha(0).Text
lblfecha2.Caption = txtfecha(1).Text

lbltotal_needed.Caption = Format(txtcantidad.Text, "$###,##0.00")
calcula_total_multiple
btnbusca_registro_Click
carga_lista


End Sub

Private Sub Form_Load()
On Error Resume Next

Top = 0
Left = (Screen.Width - Width) / 2


 ' Establece el Backcolor
    Color_Fondo barra.hwnd, &H80000008
    ' establece el color de la barra en rojo
    Color_Progreso barra.hwnd, &H2B36FF    '  &HFF00&
    
   
    
    
    

actualiza = 0
  nf = FreeFile
  Open "\\192.168.84.215\pagos\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\pagos\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     r$ = Shell("\\192.168.84.215\pagos\actualizador.exe", vbNormalFocus)
     
     Hide
    '
     End
     Exit Sub
  End If
  
  
rango_dia = 1


cbosort1.Clear
cbosort1.AddItem "Amount" + Space(20) + "10"
cbosort1.AddItem "Debit" + Space(20) + "3"
cbosort1.AddItem "Description" + Space(20) + "7"
cbosort1.AddItem "Policy" + Space(20) + "8"
cbosort1.AddItem "Receipt" + Space(20) + "9"
cbosort1.AddItem "Date" + Space(20) + "12"
cbosort1.AddItem "Id Customer" + Space(20) + "13"
cbosort1.AddItem "Company" + Space(20) + "14"
cbosort1.AddItem "Bank Date" + Space(20) + "6"



cbosort3.Clear
cbosort3.AddItem "Amount" + Space(20) + "10"
cbosort3.AddItem "Debit" + Space(20) + "3"
cbosort3.AddItem "Description" + Space(20) + "7"
cbosort3.AddItem "Policy" + Space(20) + "8"
cbosort3.AddItem "Receipt" + Space(20) + "9"
cbosort3.AddItem "Date" + Space(20) + "12"
cbosort3.AddItem "Id Customer" + Space(20) + "13"
cbosort3.AddItem "Company" + Space(20) + "14"
cbosort3.AddItem "Bank Date" + Space(20) + "6"





limite_inferior = 0.95
limite_superior = 1.05
grabado = 1

lblfila_grid4.Clear
For t = 1 To 10
  lblfila_grid4.AddItem t
Next t

If (App.PrevInstance = True) Then
  'base.Close
  End
End If

cboyear.Clear
ano_actual = Format(Now, "yyyy")

cboyear.AddItem ano_actual - 2
cboyear.AddItem ano_actual - 1
cboyear.AddItem ano_actual
cboyear.AddItem ano_actual + 1

' asigna el año actual
For t = 0 To cboyear.ListCount - 1
  If ano_actual = cboyear.List(t) Then
     cboyear.ListIndex = t
     Exit For
  End If
Next t

mes_actual2 = Format(Now, "mm")
btnmes(mes_actual2 - 1).Value = True
btnmes_Click (mes_actual2 - 1)




op_searchx = 1
Checa_status

 Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 940 '1024
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
        'ScaleFactorX = (Xpixels / DesignX)
        'ScaleFactorY = (Ypixels / DesignY)
      
        ScaleFactorX = 1360 / DesignX
        ScaleFactorY = 1024 / DesignY
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        forma_main.Height = 9000 'Me.Height ' Remember the current size
        forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0


Conecta_SQL


carga_aseguranzas

carga_programas


carga_combo_oficinas




Exit Sub


If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    AddToTray Me, mnuTray

    SetTrayTip "VB Helper tray icon program"

End Sub


Private Sub Form_Resize()
 On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1
End Sub


Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
 

 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE" '"admin"
 user_ini$ = "payroll"  '"sa"
 bd_ini$ = "laesystemja"  '"CallCenter"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com" ' "167.114.199.93"

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub
Private Sub Form_Terminate()
On Error Resume Next
 base.Close
End Sub





Private Sub Grid1_EnterCell()
On Error Resume Next

If op_rem(1).Value = True Then

  guarda$ = Grid1.Text

  If Grid1.Col = 15 Or Grid1.Col = 1 Then
     lblrow.Caption = Grid1.Row
     Grid1.Col = 15
     txtnota.Text = Grid1.Text
  Else
     lblrow.Caption = ""
  End If



End If




If Grid1.Col = 11 Then

    If Grid1.Text = "Ok" Then
       Grid1.Text = "---"
    ElseIf Grid1.Text = "---" Then
       Grid1.Text = "Ok"
    End If
    
    Grid1.Col = 12
End If

  

   

If op_grid(0).Value = True Then
If Grid1.Col = 1 Then
    Grid1.Col = 0
    txtfila.Text = Grid1.Text
    
    Grid1.Col = 3
    debito$ = Grid1.Text
    
    Grid1.Col = 4
    credito$ = Grid1.Text
    
    
    If debito <> "" Then
      txtamount.Text = debito$
    Else
      txtamount.Text = credito$
    End If
    
    
    
    Grid1.Col = 7
    concepto$ = Grid1.Text
    
    Grid1.Col = 14
    SHORTNAME$ = Grid1.Text
    
    cbocompany.ListIndex = -1
    For Y = 0 To cbocompany.ListCount - 1
       If UCase(Left(cbocompany.List(Y), Len(SHORTNAME$))) = UCase(SHORTNAME$) Then
            cbocompany.ListIndex = Y
            Exit For
       End If
    Next Y
    
        
    Grid1.Col = 8
    txtpoliza.Text = Grid1.Text
    
    Grid1.Col = 13
  r$ = Grid1.Text
  
  If Val(r$) = 0 Then
     txtcust_id.Text = ""
  Else
     txtcust_id.Text = r$
  End If
  

     cboprogram.ListIndex = -1
  Grid1.Col = 16
  programa$ = UCase(Grid1.Text)
  
  existe = 0
  For Y = 0 To cboprogram.ListCount - 1
      p$ = UCase(LTrim(RTrim(Left(cboprogram.List(Y), Len(cboprogram.List(Y)) - 15))))
      If p$ = programa$ Then
         cboprogram.ListIndex = Y
         existe = 1
         Exit For
      End If
  Next Y
    
    txtfirst_name.Text = ""
    txtlast_name.Text = ""
    
    
    
  End If



  If op_busca2(0).Value = True Then
    Grid1.Col = 10
    txtcantidad2.Text = Grid1.Text
  ElseIf op_busca2(1).Value = True Then
    Grid1.Col = 8
    txtcantidad2.Text = Grid1.Text
  End If


  grid4.Clear
  grid4.Rows = 2
  enca_grid4

  lineas = grid4.Rows - 1
  If lineas < 0 Then lineas = 0
  grid4.Row = 1
  grid4.Col = 0
     
     If Val(grid4.Text) < 1 Then
         lbltotal4.Caption = "0"
     Else
         lbltotal4.Caption = Str(lineas)
     End If


End If


grabado = 0

End Sub






Private Sub Grid3_Click()

On Error Resume Next
chk_multiple.Value = False
Picture1.Visible = False
encapsula.Visible = False
End Sub

Private Sub Grid3_EnterCell()
On Error Resume Next


'Conecta_SQL

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
           
cbo_oficina.ListIndex = -1

If op_rem(0).Value = True Then

  guarda$ = Grid3.Text

  If Grid3.Col = 15 Or Grid3.Col = 1 Then
     lblrow.Caption = Grid3.Row
      Grid3.Col = 15
     txtnota.Text = Grid3.Text
  Else
     lblrow.Caption = ""
  End If
  
End If




 
  
  
If Grid3.Col = 11 Then

    If Grid3.Text = "---" Then
      Grid3.Col = 11
      Grid3.Text = ""
      
      Grid3.Col = 22
      Grid3.Text = "1"
      
      GoTo vente_aqui
    End If
    
    
    If Grid3.Text = "" Then
      Grid3.Col = 11
      Grid3.Text = "---"
      
      Grid3.Col = 22
      Grid3.Text = "1"
      
      GoTo vente_aqui
    End If
    




    If Grid3.Text = "Ok" Or Grid3.Text = "True" Then
       For Y = 9 To 20
          Grid3.Col = Y
          If Y <> 14 Then
             Grid3.Text = ""
          End If
       Next Y
       
    End If
    
    
    
    
vente_aqui:

    Grid3.Col = 2
    Grid3.Text = ""
    
    Grid3.Col = 12
   
End If
  
  

  
  
If op_grid(1).Value = True Then


  If Grid3.Col = 1 Then
    Grid3.Col = 0
    txtfila.Text = Grid3.Text
    
    Grid3.Col = 3
    debito$ = Grid3.Text
    
    Grid3.Col = 4
    credito$ = Grid3.Text
    
    
    If debito <> "" Then
      txtamount.Text = debito$
    Else
      txtamount.Text = credito$
    End If
    
    
    Grid3.Col = 7
    concepto$ = Grid3.Text
    
      
    
    
    
    Grid3.Col = 14
    SHORTNAME$ = Grid3.Text
    
    c = revisa_compania(concepto$)
    
     Grid3.Col = 13
  r$ = Grid3.Text
  
  If Val(r$) = 0 Then
     txtcust_id.Text = ""
  Else
     txtcust_id.Text = r$
  End If
  
  
  cboprogram.ListIndex = -1
  Grid3.Col = 16
  programa$ = UCase(Grid3.Text)
  
  existe = 0
  For Y = 0 To cboprogram.ListCount - 1
      p$ = UCase(LTrim(RTrim(Left(cboprogram.List(Y), Len(cboprogram.List(Y)) - 15))))
      If p$ = programa$ Then
         cboprogram.ListIndex = Y
         existe = 1
         Exit For
      End If
  Next Y


    
    txtfirst_name.Text = ""
    txtlast_name.Text = ""
    
    If UCase(SHORTNAME$) = "INFINITY" Then
       Name1$ = Mid(concepto$, 25, Len(concepto$) - 23)
       pos = InStr(1, Name1$, ".")
       apell$ = Left(Name1$, pos - 1)
       txtlast_name.Text = apell$
       
    End If
    
    idoficina = 0
    If UCase(SHORTNAME$) = "NATIONAL GENERAL" Then
       bankcode = Val(Right(concepto$, 6))
       Select Case bankcode
       Case 31349  ' florence
         idoficina = 4
       Case 31350  ' haven
         idoficina = 5
       Case 57237  ' citrus
         idoficina = 2
       Case 57238  ' compton
         idoficina = 1
       Case 57240  ' San Bernardino
         idoficina = 7
       Case 57246   ' santa ana
         idoficina = 8
       Case 57248   ' whittier
         idoficina = 10
       Case 73138   ' Vannuys
         idoficina = 9
       Case 31348   ' PHS
         idoficina = 13
       End Select
       
       
       sSelect = "select office from OfficesCatalog where idoffice='" + Format(idoficina, "#0") + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       oficina$ = UCase(Rs(0))
       Rs.Close
       
       For Y = 0 To cbo_oficina.ListCount - 1
          ofic$ = UCase(LTrim(RTrim(Left(cbo_oficina.List(Y), Len(cbo_oficina.List(Y)) - 6))))
          If oficina$ = ofic$ Then
           ' cbo_oficina.ListIndex = Y
            Exit For
          End If
       Next Y
    
    End If
    
    
    
    
    
    If SHORTNAME$ <> "" Then
      cbocompany.ListIndex = -1
      For Y = 0 To cbocompany.ListCount - 1
       If UCase(Left(cbocompany.List(Y), Len(SHORTNAME$))) = UCase(SHORTNAME$) Then
            cbocompany.ListIndex = Y
            Exit For
       End If
      Next Y
    Else
       cbocompany.ListIndex = -1
    End If
    
        
    Grid3.Col = 8
    txtpoliza.Text = Grid3.Text
    
    
    
  End If


  If op_busca(0).Value = True Then
    Grid3.Col = 3
    txtcantidad.Text = Grid3.Text
    txtcantidad3.Text = txtcantidad.Text
  ElseIf op_busca(1).Value = True Then
    Grid3.Col = 8
    txtcantidad.Text = Grid3.Text
    txtcantidad3.Text = txtcantidad.Text
  End If
   
  



  grid4.Clear
  grid4.Rows = 2
  enca_grid4

  lineas = grid4.Rows - 1
  If lineas < 0 Then lineas = 0
  grid4.Row = 1
  grid4.Col = 0
  If Val(grid4.Text) < 1 Then
    lbltotal4.Caption = "0"
  Else
    lbltotal4.Caption = Str(lineas)
  End If

  
End If

grabado = 0



End Sub


Private Sub grid4_EnterCell()
On Error Resume Next
lblfila_grid4.Text = grid4.Row

End Sub





Private Sub grid5_EnterCell()
On Error Resume Next

If op_grid(2).Value = True Then

  guarda$ = grid5.Text

  If grid5.Col = 15 Or grid5.Col = 1 Then
     lblrow.Caption = grid5.Row
  Else
     lblrow.Caption = ""
  End If
  
  
  
  If grid5.Col = 11 Then

    If grid5.Text = "Ok" Or grid5.Text = "---" Or grid5.Text = "True" Then
       For Y = 9 To 20
          grid5.Col = Y
          If Y <> 14 Then
             grid5.Text = ""
          End If
       Next Y
       
    End If
    
    grid5.Col = 12
  End If
  
  
   
   
   

  If grid5.Col = 1 Then
    grid5.Col = 0
    txtfila.Text = "NEW"
    
    
    grid5.Col = 6
    txtamount.Text = grid5.Text
    
    
    grid5.Col = 7
    concepto$ = grid5.Text
    
    grid5.Col = 10
    SHORTNAME$ = grid5.Text
    
    If SHORTNAME$ <> "" Then
    cbocompany.ListIndex = -1
    For Y = 0 To cbocompany.ListCount - 1
       If UCase(Left(cbocompany.List(Y), Len(SHORTNAME$))) = UCase(SHORTNAME$) Then
            cbocompany.ListIndex = Y
            Exit For
       End If
    Next Y
    Else
       cbocompany.ListIndex = -1
    End If
    
        
    grid5.Col = 11
    txtpoliza.Text = grid5.Text
    
    grid5.Col = 13
    txtfirst_name.Text = grid5.Text
    
    grid5.Col = 14
    txtlast_name.Text = grid5.Text
    
  End If



   If op_busca3(0).Value = True Then
    grid5.Col = 6
    txtcantidad3.Text = grid5.Text
   ElseIf op_busca3(1).Value = True Then
    grid5.Col = 11
    txtcantidad3.Text = grid5.Text
   
   End If


grid4.Clear
grid4.Rows = 2
enca_grid4

lineas = grid4.Rows - 1
If lineas < 0 Then lineas = 0
grid4.Row = 1
grid4.Col = 0
If Val(grid4.Text) < 1 Then
   lbltotal4.Caption = "0"
Else
   lbltotal4.Caption = Str(lineas)
End If


End If


grabado = 0

End Sub



Private Sub HScroll1_Change()

End Sub

Private Sub List2_Click()
On Error Resume Next
suma = 0
For t = 0 To List2.ListCount - 1
  If List2.Selected(t) = True Then
     suma = suma + Val(Left(List2.List(t), Len(List2.List(t)) - 10))
  End If
Next t

lbl_total_marcado.Caption = Format(suma, "$###,##0.00")


End Sub




Private Sub op_day_Click(Index As Integer)
On Error Resume Next

  rango_dia = Index + 1

End Sub


Private Sub op_fecha_carga_Click(Index As Integer)
rango = Index
End Sub

Private Sub op_mes_Click(Index As Integer)
carga_grid5
End Sub



Private Sub op_rango_mes_Click(Index As Integer)
If Index = 1 Then
  txtfecha_cargada(0).Enabled = True
  txtfecha_cargada(1).Enabled = True
Else
  txtfecha_cargada(0).Enabled = False
  txtfecha_cargada(1).Enabled = False
End If
  
  
  
  
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next
num = Slider1.Value
Select Case num
Case 0
  limite_inferior = 1
  limite_superior = 1

Case 1
  limite_inferior = 0.99
  limite_superior = 1.01

Case 2
  limite_inferior = 0.98
  limite_superior = 1.02

Case 3
  limite_inferior = 0.97
  limite_superior = 1.03

Case 4
  limite_inferior = 0.96
  limite_superior = 1.04

Case 5
  limite_inferior = 0.95
  limite_superior = 1.05


End Select




End Sub
























Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
If seg >= 3 Then
  msg.Refresh
  lblmsg.Refresh
  lblmsg2.Refresh
 
End If


If seg = 10 Then
   Refresh
   seg = 0
End If



End Sub






Private Sub txtfecha_cargada_Click(Index As Integer)
On Error Resume Next
calen1 = Index

If Index = 0 Then

  If txtfecha_cargada(0).Text = "" Then
     Calendar1.Value = Format(mes_actual, "00") + "-01-" + cboyear.List(cboyear.ListIndex)
     'Calendar1.Today
  Else
     Calendar1.Value = txtfecha_cargada(0).Text
  End If
Else
  If txtfecha_cargada(1).Text = "" Then
     Calendar1.Value = Format(mes_actual, "00") + "-01-" + cboyear.List(cboyear.ListIndex)
     'Calendar1.Today
  Else
     Calendar1.Value = txtfecha_cargada(1).Text
  End If


End If

Calendar1.Visible = True


End Sub


Private Sub txtfecha_cargada_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


Private Sub txtfecha_cargada_LostFocus(Index As Integer)
Calendar1.Visible = False

End Sub

Private Sub txtfecha_Click(Index As Integer)
On Error Resume Next
calen2 = Index

If Index = 0 Then

  If txtfecha(0).Text = "" Then
     Calendar2.Today
  Else
     Calendar2.Value = txtfecha(0).Text
  End If
Else
  If txtfecha(1).Text = "" Then
     Calendar2.Today
  Else
     Calendar2.Value = txtfecha(1).Text
  End If


End If

Calendar2.Visible = True

End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub




Private Sub txtfecha_LostFocus(Index As Integer)
Calendar2.Visible = False
End Sub


Private Sub txtfila_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
   Exit Sub
End If

If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
Else
  KeyAscii = 0
End If


End Sub






Public Sub separa_efectivos()
On Error Resume Next

msg.Visible = True
msg.Refresh

lblmsg.Caption = "Removing the cash deposits from the excel sheet..."

lblmsg.Refresh






Grid3.Visible = False
Grid3.Clear
Grid3.Rows = Grid1.Rows
cont = 0
For t = 1 To Grid1.Rows - 1
   Grid1.Row = t
   
   Grid1.Col = 4
   cant = Val(Grid1.Text)
   
   Grid1.Col = 7
   a$ = UCase(Grid1.Text)
   
   
   pos1 = InStr(1, a$, "MERCHANT")
   pos2 = InStr(1, a$, "CASH VAULT")
   pos3 = InStr(1, a$, "AUTHNET")
   pos4 = InStr(1, a$, "BRINKS")
   pos5 = InStr(1, a$, "VENMO")
   pos6 = InStr(1, a$, "PAYPAL")
   pos7 = InStr(1, a$, "E TRANSFER")
   pos8 = InStr(1, a$, "DEPOSIT")
   
   
   
   
   
   
   If Left(a$, 8) = "MERCHANT" Or Left(a$, 10) = "CASH VAULT" Or Left(a$, 7) = "AUTHNET" Or Left(a$, 6) = "BRINKS" Or Left(a$, 5) = "VENMO" Or Left(a$, 6) = "PAYPAL" Or Left(a$, 10) = "E TRANSFER" Or Left(a$, 7) = "DEPOSIT" Then
   
   ElseIf pos1 > 0 Or pos2 > 0 Or pos3 > 0 Or pos4 > 0 Or pos5 > 0 Or pos6 > 0 Or pos8 > 0 Or cant > 0 Then
     
   Else
      cont = cont + 1
      Grid3.Row = cont
      For Y = 1 To 18
         Grid1.Row = t
         Grid1.Col = Y
         Grid3.Col = Y
         Grid3.Text = Grid1.Text
      Next Y
   End If
  
Next t


filas = cont + 1

cont = 0
Grid1.Clear
Grid1.Rows = filas ' - 1

For t = 1 To filas '- 2
  Grid3.Row = t
  Grid1.Row = t
  cont = cont + 1
  Grid1.Col = 0
  Grid1.Text = Str(cont)
  
  For Y = 1 To 18
    Grid3.Col = Y
    Grid1.Col = Y
    Grid1.Text = Grid3.Text
  Next Y
Next t



enca_grid1
 
Grid3.Clear
Grid3.Rows = 1
enca_grid3

Grid3.Visible = True

msg.Visible = False
End Sub

Public Sub sql()
On Error Resume Next
 
 If Grid1.Rows <= 2 Then
     Exit Sub
 End If
 
  
r$ = MsgBox("Do you want to save all the data in SQL?", 4, "Attention")
If r$ = "7" Then Exit Sub
 

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

 barra.Visible = True
  barra.Min = 1
  barra.Max = Grid1.Rows - 1
  
  msg.Visible = True
  
  lblmsg.Caption = "Saving all the information in the database..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  contador = 0
  Erase recibos$
  
' -------------------------------------------------------------------------------------------------------

  For t = 1 To 24 'Grid1.Rows - 1
    
    Grid1.Row = t
    
    IDReceiptHDR$ = ""
    
    Grid1.Col = 1 ' account
    account$ = Grid1.Text
    account$ = "243162505"
    
     X1 = IsNumeric(account$)
    If X1 = False Then
       If Right$(account$, 4) = "2505" Then
          account$ = "243162505"
       Else
          account$ = InputBox("Type the account number: (Only numeric)", "Attention")
       End If
    End If
    
    Grid1.Col = 2 ' chkref
    chkref$ = Grid1.Text
    
    Grid1.Col = 3  ' debito
    debito = Val(Grid1.Text)
        
    Grid1.Col = 4  ' credit
    credito = Val(Grid1.Text)
    
    Grid1.Col = 5 ' balance
    balance = Val(Grid1.Text)
        
    Grid1.Col = 6   '  date
    fecha_pagada$ = Grid1.Text
    
    Grid1.Col = 7 ' descripcion
    Description$ = Grid1.Text
           
    Grid1.Col = 8  ' poliza
    poliza$ = UCase(Grid1.Text)
    
    Grid1.Col = 9  ' receipt HDR
    IDReceiptHDR$ = Grid1.Text
    
    Grid1.Col = 10 ' amount
    amount = Val(Format(Grid1.Text, "0000000.00"))
    
    Grid1.Col = 11 ' verificado
    verificado$ = Grid1.Text
    
    Grid1.Col = 12  ' date created
    fecha_creacion$ = Grid1.Text
        
    Grid1.Col = 13   ' Id cust
    IdCustomer$ = Grid1.Text
    
    Grid1.Col = 14   ' company
    compania$ = Grid1.Text
    
    Grid1.Col = 15
    nota$ = Grid1.Text
    
    ' If amount = 0 Then GoTo al_final
    
    
    contador = contador + 1
    barra.Value = contador
    lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid1.Rows - 1, "###0")
    lblmsg2.Refresh
    openforms = DoEvents
    
    If IDReceiptHDR$ <> "" Then
        GoTo salta
    Else
        GoTo al_final
    End If
    
    
    If poliza$ = "" Then GoTo al_final
    
    
    
   
    
    

    
    
   
    
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    IDReceiptHDR$ = ""
    
    sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + fecha_rango1$ + "' and date<'" + fecha_rango2$ + "' and amount='" + Format(amount, "#######0.00") + "' and active='1'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IDReceiptHDR$ = Rs(0)  'correcto
                         
    Rs.Close
    
    
    
    ' busca en la lista de recibos por si ya se encuentra
    encontrado = 0
    For Y = 0 To 5000
       If IDReceiptHDR$ = recibos$(Y) Then
           encontrado = 1
           Exit For
       End If
       
       If recibos$(Y) = "" Then
         Exit For
       End If
       
    Next Y
    
    If encontrado = 0 Then
       recibos$(Y) = IDReceiptHDR$
    Else
         Set Rs = New ADODB.Recordset
         'Checa_status
   
         IDReceiptHDR$ = ""
               
         sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + fecha_rango1$ + "' and date<'" + fecha_rango2$ + "' and amount='" + Format(amount, "#######0.00") + "' and IdReceiptHDR<>'" + recibos$(Y) + "' and active='1'"             ' where idpolicieshdr='" + idpolizahdr$ + "'"
    
         ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
         Rs.Open sSelect, base, adOpenUnspecified
    
         IDReceiptHDR$ = Rs(0)  'correcto
                         
         Rs.Close
    
    End If
    
    
    
    
        
    
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    fechacreada$ = ""
    sSelect = "select date from Receiptsdtl where IdReceiptHDR='" + IDReceiptHDR$ + "' and idpolicieshdr='" + idpolizahdr$ + "' and date>='" + fecha_rango1$ + "' and date<'" + fecha_rango2$ + "' and amount='" + Format(amount, "#######0.00") + "' and active='1'"
    
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    fechacreada$ = Rs(0)
    
    Rs.Close
       
    
    
    
   ' Set Rs = New ADODB.Recordset
    'Checa_status
        
   ' sSelect = "select ins.CompanyName " & _
    '        "FROM   ReceiptsHDR  rechdr " & _
    '        "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    '        "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    '        "Where rechdr.IdReceiptHDR = '" + IdReceiptHDR$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    'Rs.Open sSelect, base, adOpenUnspecified
    
    'compania$ = Rs(0)
                         
    'Rs.Close
    
    
    
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    IdCustomer$ = ""
     sSelect = "select idcustomer from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IdCustomer$ = Rs(0)
                         
    Rs.Close
    
    
salta:


        

    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idpolizahdr$ = ""
    sSelect = "select idpolicieshdr from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idpolizahdr$ = Rs(0)
                         
    Rs.Close
          
   ' If idpolizahdr$ = "" Then GoTo al_final
    
    


   
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    idcompany$ = ""
     sSelect = "select idcompany from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idcompany$ = Rs(0)
                         
    Rs.Close
    
    
    
    
    
    

    
    
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    idreceiptdtl$ = ""
               
    sSelect = "select IdReceiptsDTL from Receiptsdtl where idreceipthdr='" + IDReceiptHDR$ + "' and idpolicieshdr='" + idpolizahdr$ + "' and date>='" + fecha_rango1$ + "' and date<'" + fecha_rango2$ + "' and amount='" + Format(amount, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idreceiptdtl$ = Rs(0)
                         
    Rs.Close
    
    
    'If t = 16 Then Stop
    
    
    Set Rs = New ADODB.Recordset


' verifica si ya existe el registro en SQL
    a$ = ""
    sSelect = "select idconciliation from Conciliationbankrec  where idcustomer='" + IdCustomer$ + "' and policyNo='" + poliza$ + "' and amount='" + Format(amount, "#######0.00") + "' and IdReceiptHDR='" + IDReceiptHDR$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    a$ = Rs(0)
    Rs.Close
    
           
           
    If UCase(verificado$) = "OK" Then
       valor_verificado$ = "1"
    Else
       valor_verificado$ = "0"
    End If
    
    
     
    If a$ = "" Then
     
    sSelect = "INSERT INTO ConciliationBankRec (Account,chkref,debit,credit,balance,date,description,idcompany,policyno,idpolicieshdr,idcustomer,idreceipthdr,idreceiptdtl,amount,receiptdate,clear,notes)  VALUES ('" + _
    account$ + "', '" + chkref$ + "', convert(money,'" + Format(debito, "#####0.00") + "'), convert(money,'" + Format(credito, "#####0.00") + "'), convert(money,'" + Format(balance, "#######0.00") + "')," + _
    "convert(datetime, '" + fecha_pagada$ + "'), '" + Description$ + "', '" + idcompany$ + "', '" + poliza$ + "', '" + idpolizahdr$ + "', '" + IdCustomer$ + "', '" + IDReceiptHDR$ + "', '" + _
    idreceiptdtl$ + "', convert(money,'" + Format(amount, "#####0.00") + "'), convert(datetime,'" + fecha_creacion$ + "'), '" + valor_verificado$ + "', '" + nota$ + "')"
    
       
     
    
    Else
    
    
     sSelect = "update ConciliationBankRec set Account='" + account$ + "', chkref='" + chkref$ + "', debit= convert(money,'" + Format(debito, "#####0.00") + "'), credit=" + _
     "convert(money,'" + Format(credito, "#####0.00") + "'), balance= convert(money,'" + Format(balance, "#######0.00") + "'), date=convert(datetime, '" + fecha_pagada$ + "')," + _
     "description='" + Description$ + "', idcompany='" + idcompany$ + "', policyno='" + poliza$ + "', idpolicieshdr='" + idpolizahdr$ + "', idcustomer='" + IdCustomer$ + "', idreceipthdr='" + _
     IDReceiptHDR$ + "', idreceiptdtl='" + idreceiptdtl$ + "', amount= convert(money,'" + Format(amount, "#####0.00") + "'), receiptdate= convert(datetime,'" + fecha_creacion$ + _
     "'), clear='" + valor_verificado$ + "', notes='" + nota$ + "' where idconciliation='" + a$ + "'"
     

    
    End If
        
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    'Rs.Close
    
al_final:
          
  Next t




msg.Visible = False
barra.Visible = False
End Sub

Public Sub graba_sql1()
On Error Resume Next
 If Grid1.Rows <= 1 Then
     Exit Sub
 End If
 
  
'r$ = MsgBox("Do you want to save all the information in the database?", 4, "Attention")
'If r$ = "7" Then Exit Sub




If Check2.Value = 1 Then
   Exit Sub
End If

Exit Sub


Dim sSelect As String
    
Dim Rs As ADODB.Recordset



  barra.Visible = True
  barra.Min = 1
  barra.Max = Grid1.Rows - 1
  
  msg.Visible = True
  
  lblmsg.Caption = "Saving all the information in the database..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  contador = 0
  Erase recibos$
  
' -------------------------------------------------------------------------------------------------------



  For t = 1 To Grid1.Rows - 1
    
    Grid1.Row = t
    
    Grid1.Col = 21
    Idconciliation$ = Grid1.Text
    
    
    Grid1.Col = 22
    If Grid1.Text = "1" And Idconciliation$ <> "" Then
      If Check2.Value = 1 Then
         GoTo no_grabes
      End If
    End If

    
    
    IDReceiptHDR$ = ""
    
    Grid1.Col = 1 ' account
    account$ = "243162505"
    
    Grid1.Col = 2 ' chkref
    chkref$ = Grid1.Text
    
    Grid1.Col = 3  ' debito
    debito = Val(Grid1.Text)
        
    Grid1.Col = 4  ' credit
    credito = Val(Grid1.Text)
    
    Grid1.Col = 5 ' balance
    balance = Val(Grid1.Text)
        
    Grid1.Col = 6   '  date
    fecha_pagada$ = Format(Grid1.Text, "yyyy-mm-dd")
    
    Grid1.Col = 7 ' descripcion
    Description$ = Grid1.Text
           
    Grid1.Col = 8  ' poliza
    poliza$ = UCase(Grid1.Text)
    
    Grid1.Col = 9  ' receipt HDR
    IDReceiptHDR$ = Grid1.Text
    
    Grid1.Col = 10 ' amount
    amount = Val(Format(Grid1.Text, "0000000.00"))
    
    Grid1.Col = 11 ' verificado
    verificado$ = Grid1.Text
    
    Grid1.Col = 12  ' date created
    fecha_creacion$ = Grid1.Text
        
    Grid1.Col = 13   ' Id cust
    IdCustomer$ = Grid1.Text
    
    Grid1.Col = 14   ' company
    compania$ = Grid1.Text
    
    Grid1.Col = 15 'Comment
    nota$ = Grid1.Text
    
    Grid1.Col = 16 ' program name Company
    programname$ = Grid1.Text
    
    Grid1.Col = 17 ' idprogram
    idprogram$ = Grid1.Text
    
    
    Set Rs = New ADODB.Recordset
      
    idprogram2$ = ""
    
    sSelect = "select idprogram from ProgramsCatalog where programname='" + idprogram$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    idprogram2$ = Rs(0)
    Rs.Close
    
    If idprogram2$ <> "" Then
       idprogram$ = idprogram2$
    End If
    
    
    Grid1.Col = 18 ' idreceiptDTL
    idreceiptdtl$ = Grid1.Text
    
    Grid1.Col = 19 ' idpolizaHDR
    idpolizahdr$ = Grid1.Text
    
    Grid1.Col = 20 '
    idcompany$ = Grid1.Text
    
    Grid1.Col = 21
    Idconciliation$ = Grid1.Text
    
    
    contador = contador + 1
    barra.Value = contador
    lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid1.Rows - 1, "###0")
    lblmsg2.Refresh
    openforms = DoEvents
    
    If IDReceiptHDR$ <> "" Then
        GoTo salta
    Else
        GoTo al_final
    End If
    
    
salta:
   
' verifica si ya existe el registro en SQL
       
           
           
      Set Rs = New ADODB.Recordset
    'Checa_status
   
  
    Idconciliation2$ = ""
    
    sSelect = "select Idconciliation from ConciliationBankRec where date='" + fecha_pagada$ + "' and amount='" + Format(amount, "#######0.00") + "' and description='" + Description$ + "' and policyno='" + poliza$ + "' and idreceipthdr='" + IDReceiptHDR$ + "'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation2$ = Rs(0)  'correcto
                         
    Rs.Close
 
   
    If Idconciliation2$ = Idconciliation$ And Idconciliation2$ <> "" Then
    
    ElseIf Idconciliation2$ = "" And Idconciliation$ <> "" Then
    
    Else
       Idconciliation$ = Idconciliation2$
    End If
            
           
           
           
    If UCase(verificado$) = "OK" Then
       valor_verificado$ = "1"
       logx$ = "Found by the PAGOS program"
    Else
       If Left(UCase(nota$), 1) = "V" Then
          logx$ = "voided receipt"
       End If
       
       valor_verificado$ = "0"
       
    End If
    
    
    
    ano_selecto = cboyear.List(cboyear.ListIndex)
        
     
    If Idconciliation$ = "" Then
     
       sSelect = "INSERT INTO ConciliationBankRec (Account,chkref,debit,credit,balance,date,description,idcompany,idprogram, policyno,idpolicieshdr,idcustomer,idreceipthdr,idreceiptdtl,amount,receiptdate,clear,notes, logs, monthconciliation, yearconciliation, uploaddate)  VALUES ('" + _
       account$ + "', '" + chkref$ + "', convert(money,'" + Format(debito, "#####0.00") + "'), convert(money,'" + Format(credito, "#####0.00") + "'), convert(money,'" + Format(balance, "#######0.00") + "')," + _
       "convert(datetime, '" + fecha_pagada$ + "'), '" + Description$ + "', '" + idcompany$ + "', '" + idprogram$ + "', '" + poliza$ + "', '" + idpolizahdr$ + "', '" + IdCustomer$ + "', '" + IDReceiptHDR$ + "', '" + _
       idreceiptdtl$ + "', convert(money,'" + Format(amount, "#####0.00") + "'), convert(datetime,'" + fecha_creacion$ + "'), '" + valor_verificado$ + "', '" + nota$ + "', '" + logx$ + "', '" + Format(mes_actual, "00") + "', '" + Format(ano_selecto, "0000") + "', convert(datetime, '" + Format(Now, "mm-dd-yyyy") + "'))"
    
    
       Rs.Open sSelect, base, adOpenUnspecified
       Rs.Close
            
       Set Rs = New ADODB.Recordset
       Checa_status
   
       Idconciliation$ = ""
       sSelect = "select Idconciliation from ConciliationBankRec where date='" + fecha_pagada$ + "' and debit='" + Format(debito, "#######0.00") + "' and credit='" + Format(credito, "#######0.00") + "' and description='" + Description$ + "' and policyno='" + poliza$ + "'"
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
    
       Idconciliation$ = Rs(0)  'correcto
                
       Rs.Close
    
    
       Grid1.Col = 21
       Grid1.Text = Idconciliation$
       
       Grid1.Col = 22
       Grid1.Text = "1"
    
    
    
    Else
    
    
       sSelect = "update ConciliationBankRec set Account='" + account$ + "', chkref='" + chkref$ + "', debit= convert(money,'" + Format(debito, "#####0.00") + "'), credit=" + _
       "convert(money,'" + Format(credito, "#####0.00") + "'), balance= convert(money,'" + Format(balance, "#######0.00") + "'), date=convert(datetime, '" + fecha_pagada$ + "')," + _
       "description='" + Description$ + "', idcompany='" + idcompany$ + "', idprogram='" + idprogram$ + "', policyno='" + poliza$ + "', idpolicieshdr='" + idpolizahdr$ + "', idcustomer='" + IdCustomer$ + "', idreceipthdr='" + _
       IDReceiptHDR$ + "', idreceiptdtl='" + idreceiptdtl$ + "', amount= convert(money,'" + Format(amount, "#####0.00") + "'), receiptdate= convert(datetime,'" + fecha_creacion$ + _
       "'), clear='" + valor_verificado$ + "', notes='" + nota$ + "', logs='" + logx$ + "', monthconciliation='" + Format(mes_actual, "00") + "', yearconciliation='" + Format(ano_selecto, "0000") + "', uploaddate=convert(datetime, '" + Format(Now, "mm-dd-yyyy") + "') where idconciliation='" + Idconciliation$ + "'"
     

       Rs.Open sSelect, base, adOpenUnspecified
          
       Rs.Close

       Grid1.Col = 21
       Grid1.Text = Idconciliation$
  
       Grid1.Col = 22
       Grid1.Text = "1"
    
    
    End If
        
                      
   
    
   
no_grabes:

    
al_final:
          
  Next t




msg.Visible = False
barra.Visible = False
End Sub

Public Sub graba_sql2()
On Error Resume Next
 If Grid3.Rows <= 1 Then
     Exit Sub
 End If
 
  
'r$ = MsgBox("Do you want to save all the information in the database?", 4, "Attention")
'If r$ = "7" Then Exit Sub
 
 
 
If Check2.Value = 1 Then
   Exit Sub
End If

 
Exit Sub



Dim sSelect As String
    
Dim Rs As ADODB.Recordset

 barra.Visible = True
  barra.Min = 1
  barra.Max = Grid3.Rows - 1
  
  msg.Visible = True
  
  lblmsg.Caption = "Saving all the information in the database..."
  lblmsg.Refresh
  lblmsg2.Caption = ""
  openforms = DoEvents
  
  contador = 0
  Erase recibos$
  
' -------------------------------------------------------------------------------------------------------

  For t = 1 To Grid3.Rows - 1
    
    Grid3.Row = t
    
    Grid3.Col = 21
    Idconciliation$ = Grid3.Text
    
    
    Grid3.Col = 22
    If Grid3.Text = "1" And Idconciliation$ <> "" Then
      If Check2.Value = 1 Then
         GoTo no_grabes
      End If
    End If

    
    IDReceiptHDR$ = ""
    
    Grid3.Col = 1 ' account
    account$ = "243162505"
    
     'X1 = IsNumeric(account$)
    'If X1 = False Then
    '   If Right$(account$, 4) = "2505" Then
    '      account$ = "243162505"
    '   Else
    '      account$ = InputBox("Type the account number: (Only numeric)", "Attention")
    '   End If
    'End If
    
    
    
    Grid3.Col = 2 ' chkref
    chkref$ = Grid3.Text
    
    Grid3.Col = 3  ' debito
    debito = Val(Grid3.Text)
        
    Grid3.Col = 4  ' credit
    credito = Val(Grid3.Text)
    
    Grid3.Col = 5 ' balance
    balance = Val(Grid3.Text)
        
    Grid3.Col = 6   '  date
    fecha_pagada$ = Format(Grid3.Text, "yyyy-mm-dd")
    
    
    Grid3.Col = 7 ' descripcion
    Description$ = Grid3.Text
           
    Grid3.Col = 8  ' poliza
    poliza$ = UCase(Grid3.Text)
    
    Grid3.Col = 9  ' receipt HDR
    IDReceiptHDR$ = Grid3.Text
    
    Grid3.Col = 10 ' amount
    amount = Val(Format(Grid3.Text, "0000000.00"))
    
    Grid3.Col = 11 ' verificado
    verificado$ = Grid3.Text
    
    Grid3.Col = 12  ' date created
    fecha_creacion$ = Grid3.Text
        
    Grid3.Col = 13   ' Id cust
    IdCustomer$ = Grid3.Text
    
    Grid3.Col = 14   ' company
    compania$ = Grid3.Text
    
    Grid3.Col = 15 'Comment
    nota$ = Grid3.Text
    
    Grid3.Col = 16 ' program name Company
    programname$ = Grid3.Text
    
    Grid3.Col = 17 ' idprogram
    idprogram$ = Grid3.Text
    
    Set Rs = New ADODB.Recordset
      
    idprogram2$ = ""
    
    sSelect = "select idprogram from ProgramsCatalog where programname='" + idprogram$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    idprogram2$ = Rs(0)
    Rs.Close
    
    
    If idprogram2$ <> "" Then
       idprogram$ = idprogram2$
    End If
    
    
    
    
    
    Grid3.Col = 18 ' idreceiptDTL
    idreceiptdtl$ = Grid3.Text
    
    Grid3.Col = 19 ' idpolizaHDR
    idpolizahdr$ = Grid3.Text
    
    Grid3.Col = 20 '
    idcompany$ = Grid3.Text
    
    Grid3.Col = 21
    Idconciliation$ = Grid3.Text
    
    
    
    
    
    ' If amount = 0 Then GoTo al_final
    
    
    contador = contador + 1
    barra.Value = contador
    lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid3.Rows - 1, "###0")
    lblmsg2.Refresh
    openforms = DoEvents
    
    
   
    
    
salta:
    
    
    
    Set Rs = New ADODB.Recordset


' verifica si ya existe el registro en SQL
     Set Rs = New ADODB.Recordset
    'Checa_status
    
   
 
   
    Idconciliation2$ = ""
    
    sSelect = "select Idconciliation from ConciliationBankRec where date='" + fecha_pagada$ + "' and debit='" + Format(debito, "#######0.00") + "' and credit='" + Format(credito, "#######0.00") + "' and description='" + Description$ + "'"  ' and policyno='" + poliza$ + "'"
    
  
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation2$ = Rs(0)  'correcto
                         
    Rs.Close
    
  
    
           
           
    If Idconciliation2$ = Idconciliation$ And Idconciliation2$ <> "" Then
           
    ElseIf Idconciliation2$ = "" And Idconciliation$ <> "" Then
    
    Else
       Idconciliation$ = Idconciliation2$
    End If
    
     
           
     If UCase(verificado$) = "OK" Then
       valor_verificado$ = "1"
       logx$ = "Found and cleared by Joselin"
     Else
       If Left(UCase(nota$), 1) = "V" Then
          logx$ = "voided receipt"
       End If
       
       valor_verificado$ = "0"
       
     End If
    
    
     ano_selecto = cboyear.List(cboyear.ListIndex)
    
        
     
    If Idconciliation$ = "" Then
     
   
     sSelect = "INSERT INTO ConciliationBankRec (Account,chkref,debit,credit,balance,date,description,idcompany,idprogram, policyno,idpolicieshdr,idcustomer,idreceipthdr,idreceiptdtl,amount,receiptdate,clear,notes, logs, monthconciliation, yearconciliation, uploaddate)  VALUES ('" + _
    account$ + "', '" + chkref$ + "', convert(money,'" + Format(debito, "#####0.00") + "'), convert(money,'" + Format(credito, "#####0.00") + "'), convert(money,'" + Format(balance, "#######0.00") + "')," + _
    "convert(datetime, '" + fecha_pagada$ + "'), '" + Description$ + "', '" + idcompany$ + "', '" + idprogram$ + "', '" + poliza$ + "', '" + idpolizahdr$ + "', '" + IdCustomer$ + "', '" + IDReceiptHDR$ + "', '" + _
    idreceiptdtl$ + "', convert(money,'" + Format(amount, "#####0.00") + "'), convert(datetime,'" + fecha_creacion$ + "'), '" + valor_verificado$ + "', '" + nota$ + "', '" + logx$ + "', '" + Format(mes_actual, "00") + "', '" + Format(ano_selecto, "00") + "', convert(datetime, '" + Format(Now, "mm-dd-yyyy") + "'))"
    
     Rs.Open sSelect, base, adOpenUnspecified
    
    
    
    Rs.Close
    
       
         Set Rs = New ADODB.Recordset
    'Checa_status
   
    Idconciliation$ = ""
    
    sSelect = "select Idconciliation from ConciliationBankRec where account='" + account$ + "' and date='" + fecha_pagada$ + "' and debit='" + Format(debito, "#######0.00") + "' and credit='" + Format(credito, "#######0.00") + "' and description='" + Description$ + "' and policyno='" + poliza$ + "'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation$ = Rs(0)  'correcto
                         
    Rs.Close
 
 
    
       Grid3.Col = 21
       Grid3.Text = Idconciliation$
       
       Grid3.Col = 22
       Grid3.Text = "1"
    
       
    
    Else
    
         
     sSelect = "update ConciliationBankRec set Account='" + account$ + "', chkref='" + chkref$ + "', debit= convert(money,'" + Format(debito, "#####0.00") + "'), credit=" + _
     "convert(money,'" + Format(credito, "#####0.00") + "'), balance= convert(money,'" + Format(balance, "#######0.00") + "'), date=convert(datetime, '" + fecha_pagada$ + "')," + _
     "description='" + Description$ + "', idcompany='" + idcompany$ + "', idprogram='" + idprogram$ + "', policyno='" + poliza$ + "', idpolicieshdr='" + idpolizahdr$ + "', idcustomer='" + IdCustomer$ + "', idreceipthdr='" + _
     IDReceiptHDR$ + "', idreceiptdtl='" + idreceiptdtl$ + "', amount= convert(money,'" + Format(amount, "#####0.00") + "'), receiptdate= convert(datetime,'" + fecha_creacion$ + _
     "'), clear='" + valor_verificado$ + "', notes='" + nota$ + "', logs='" + logx$ + "', monthconciliation='" + Format(mes_actual, "00") + "', yearconciliation='" + Format(ano_selecto, "0000") + "', uploaddate=convert(datetime, '" + Format(Now, "mm-dd-yyyy") + "') where idconciliation='" + Idconciliation$ + "'"
     
     
     Rs.Open sSelect, base, adOpenUnspecified
    
    
    
    Rs.Close
    
    
       Grid3.Col = 21
       Grid3.Text = Idconciliation$
       
       Grid3.Col = 22
       Grid3.Text = "1"
    
       
    
    End If
        
                      
    
   
no_grabes:

    
al_final:
          
  Next t




msg.Visible = False
barra.Visible = False
End Sub

Public Sub enca_grid1()
On Error Resume Next


Grid1.ColWidth(0) = 800
Grid1.ColWidth(1) = 1200 'account
Grid1.ColAlignment(1) = flexAlignRightCenter

Grid1.ColWidth(2) = 800   ' chkref
Grid1.ColAlignment(2) = flexAlignLeftCenter

Grid1.ColWidth(3) = 1000   ' debit
Grid1.ColAlignment(3) = flexAlignRightCenter

Grid1.ColWidth(4) = 1000   ' credit
Grid1.ColAlignment(4) = flexAlignRightCenter

Grid1.ColWidth(5) = 1200   'balance
Grid1.ColAlignment(5) = flexAlignRightCenter

Grid1.ColWidth(6) = 1100   ' date
Grid1.ColAlignment(6) = flexAlignLeftCenter

Grid1.ColWidth(7) = 4400   ' descrip
Grid1.ColAlignment(7) = flexAlignLeftCenter

Grid1.ColWidth(8) = 2200   ' policy
Grid1.ColAlignment(8) = flexAlignLeftCenter

Grid1.ColWidth(9) = 1200   ' receipt
Grid1.ColAlignment(9) = flexAlignCenterCenter

Grid1.ColWidth(10) = 1000   ' amount
Grid1.ColAlignment(10) = flexAlignRightCenter

Grid1.ColWidth(11) = 800   ' verified
Grid1.ColAlignment(11) = flexAlignCenterCenter

Grid1.ColWidth(12) = 1200   ' date created
Grid1.ColAlignment(12) = flexAlignLeftCenter

Grid1.ColWidth(13) = 1100   ' idcustomer
Grid1.ColAlignment(13) = flexAlignCenterCenter

Grid1.ColWidth(14) = 2000   ' company
Grid1.ColAlignment(14) = flexAlignLeftCenter

Grid1.ColWidth(15) = 4000  ' comment
Grid1.ColAlignment(15) = flexAlignLeftCenter

Grid1.ColWidth(16) = 2200  ' program name
Grid1.ColAlignment(16) = flexAlignLeftCenter

Grid1.ColWidth(17) = 1200  ' idprogram
Grid1.ColAlignment(17) = flexAlignCenterCenter

Grid1.ColWidth(18) = 1200  ' idreceiptDTL
Grid1.ColAlignment(18) = flexAlignRightCenter

Grid1.ColWidth(19) = 1200  ' idpolizaHDR
Grid1.ColAlignment(19) = flexAlignRightCenter

Grid1.ColWidth(20) = 1200  ' idcompany
Grid1.ColAlignment(20) = flexAlignCenterCenter

Grid1.ColWidth(21) = 900  ' idconciliation
Grid1.ColAlignment(21) = flexAlignCenterCenter

Grid1.ColWidth(22) = 400  '
Grid1.ColAlignment(22) = flexAlignCenterCenter




Grid1.Row = 0

Grid1.Col = 1
Grid1.Text = "Account"

Grid1.Col = 2
Grid1.Text = "ChkRef"

Grid1.Col = 3
Grid1.Text = "Debit"

Grid1.Col = 4
Grid1.Text = "Credit"

Grid1.Col = 5
Grid1.Text = "Balance"

Grid1.Col = 6
Grid1.Text = "Bank Date"

Grid1.Col = 7
Grid1.Text = "Description"


Grid1.Col = 8
Grid1.Text = "Policy"


Grid1.Col = 9
Grid1.Text = "Receipt"


Grid1.Col = 10
Grid1.Text = "Amount"

Grid1.Col = 11
Grid1.Text = "Verified"

Grid1.Col = 12
Grid1.Text = "Date"

Grid1.Col = 13
Grid1.Text = "IdCustomer"

Grid1.Col = 14
Grid1.Text = "Company"

Grid1.Col = 15
Grid1.Text = "Comment"


Grid1.Col = 16
Grid1.Text = "Program Name"
          
Grid1.Col = 17
Grid1.Text = "IDprogram"
          
Grid1.Col = 18
Grid1.Text = "IDreceiptDTL"
          
Grid1.Col = 19
Grid1.Text = "IDpoliciesHDR"
          
Grid1.Col = 20
Grid1.Text = "IDCompany"

Grid1.Col = 21
Grid1.Text = "IDconciliation"

Grid1.FixedRows = 1
Grid1.FixedCols = 1

lbltotal2.Caption = Format(Grid1.Rows - 1, "###,##0")

End Sub

Public Sub enca_grid3()
On Error Resume Next



Grid3.ColWidth(0) = 800
Grid3.ColWidth(1) = 1200 'account
Grid3.ColAlignment(1) = flexAlignRightCenter

Grid3.ColWidth(2) = 900   ' chkref
Grid3.ColAlignment(2) = flexAlignLeftCenter

Grid3.ColWidth(3) = 1000   ' debit
Grid3.ColAlignment(3) = flexAlignRightCenter

Grid3.ColWidth(4) = 1000   ' credit
Grid3.ColAlignment(4) = flexAlignRightCenter

Grid3.ColWidth(5) = 1200   'balance
Grid3.ColAlignment(5) = flexAlignRightCenter

Grid3.ColWidth(6) = 1100   ' date
Grid3.ColAlignment(6) = flexAlignLeftCenter

Grid3.ColWidth(7) = 4400   ' descrip
Grid3.ColAlignment(7) = flexAlignLeftCenter

Grid3.ColWidth(8) = 2200   ' policy
Grid3.ColAlignment(8) = flexAlignLeftCenter

Grid3.ColWidth(9) = 1200   ' receipt
Grid3.ColAlignment(9) = flexAlignCenterCenter

Grid3.ColWidth(10) = 1000   ' amount
Grid3.ColAlignment(10) = flexAlignRightCenter

Grid3.ColWidth(11) = 800   ' verified
Grid3.ColAlignment(11) = flexAlignCenterCenter

Grid3.ColWidth(12) = 1200   ' date created
Grid3.ColAlignment(12) = flexAlignLeftCenter

Grid3.ColWidth(13) = 1100   ' idcustomer
Grid3.ColAlignment(13) = flexAlignCenterCenter

Grid3.ColWidth(14) = 2000   ' company
Grid3.ColAlignment(14) = flexAlignLeftCenter

Grid3.ColWidth(15) = 4000  ' comment
Grid3.ColAlignment(15) = flexAlignLeftCenter

Grid3.ColWidth(16) = 2200  ' program name
Grid3.ColAlignment(16) = flexAlignLeftCenter

Grid3.ColWidth(17) = 1200  ' idprogram
Grid3.ColAlignment(17) = flexAlignCenterCenter

Grid3.ColWidth(18) = 1200  ' idreceiptDTL
Grid3.ColAlignment(18) = flexAlignRightCenter

Grid3.ColWidth(19) = 1200  ' idpolizaHDR
Grid3.ColAlignment(19) = flexAlignRightCenter

Grid3.ColWidth(20) = 1200  ' idcompany
Grid3.ColAlignment(20) = flexAlignCenterCenter

Grid3.ColWidth(21) = 900  ' idconciliation
Grid3.ColAlignment(21) = flexAlignCenterCenter

Grid3.ColWidth(22) = 400  '
Grid3.ColAlignment(22) = flexAlignCenterCenter





Grid3.Row = 0

Grid3.Col = 1
Grid3.Text = "Account"

Grid3.Col = 2
Grid3.Text = "ChkRef"

Grid3.Col = 3
Grid3.Text = "Debit"

Grid3.Col = 4
Grid3.Text = "Credit"

Grid3.Col = 5
Grid3.Text = "Balance"

Grid3.Col = 6
Grid3.Text = "Bank date"

Grid3.Col = 7
Grid3.Text = "Description"


Grid3.Col = 8
Grid3.Text = "Policy"


Grid3.Col = 9
Grid3.Text = "Receipt"


Grid3.Col = 10
Grid3.Text = "Amount"

Grid3.Col = 11
Grid3.Text = "Verified"

Grid3.Col = 12
Grid3.Text = "Date"

Grid3.Col = 13
Grid3.Text = "IdCustomer"

Grid3.Col = 14
Grid3.Text = "Company"

Grid3.Col = 15
Grid3.Text = "Comment"


Grid3.Col = 16
Grid3.Text = "Program Name"
          
Grid3.Col = 17
Grid3.Text = "IDprogram"
          
Grid3.Col = 18
Grid3.Text = "IDreceiptDTL"
          
Grid3.Col = 19
Grid3.Text = "IDpoliciesHDR"
          
Grid3.Col = 20
Grid3.Text = "IDCompany"

Grid3.Col = 21
Grid3.Text = "IDConciliation"


Grid3.FixedRows = 1
Grid3.FixedCols = 1

lbltotal1.Caption = Format(Grid3.Rows - 1, "###,##0")
End Sub

Public Sub carga_aseguranzas()
On Error Resume Next
cbocompany.Clear

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
  grid2.Clear
    
  sSelect = "select idcompany, shortname from insurancecatalog"
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid2.DataSource = Rs
                         
    Rs.Close
    
    
For t = 1 To grid2.Rows - 1
   grid2.Row = t
   grid2.Col = 1
   idcompany$ = grid2.Text
    
   grid2.Col = 2
   SHORTNAME$ = grid2.Text
   
   a$ = SHORTNAME$ + Space(20) + idcompany$
   cbocompany.AddItem a$
Next t

   
   
   
   
   
End Sub

Public Sub enca_grid4()
On Error Resume Next




grid4.ColWidth(0) = 800
grid4.ColAlignment(0) = flexAlignLeftCenter

grid4.ColWidth(1) = 1200 ' rec hdr
grid4.ColAlignment(1) = flexAlignRightCenter

grid4.ColWidth(2) = 1200  ' rec dtl
grid4.ColAlignment(2) = flexAlignRightCenter

grid4.ColWidth(3) = 1200   ' pol hdr
grid4.ColAlignment(3) = flexAlignRightCenter

grid4.ColWidth(4) = 1800 ' date
grid4.ColAlignment(4) = flexAlignLeftCenter

grid4.ColWidth(5) = 1200   ' amount
grid4.ColAlignment(5) = flexAlignRightCenter

grid4.ColWidth(6) = 1200   ' idcompany
grid4.ColAlignment(6) = flexAlignCenterCenter

grid4.ColWidth(7) = 1200   ' id program
grid4.ColAlignment(7) = flexAlignCenterCenter

grid4.ColWidth(8) = 1800   ' prog
grid4.ColAlignment(8) = flexAlignLeftCenter

grid4.ColWidth(9) = 2400 ' compania
grid4.ColAlignment(9) = flexAlignLeftCenter

grid4.ColWidth(10) = 1800  ' poliza
grid4.ColAlignment(10) = flexAlignLeftCenter

grid4.ColWidth(11) = 1200  ' id cust
grid4.ColAlignment(11) = flexAlignCenterCenter

grid4.ColWidth(12) = 1500  ' first name cust
grid4.ColAlignment(12) = flexAlignLeftCenter

grid4.ColWidth(13) = 1500  ' last name cust
grid4.ColAlignment(13) = flexAlignLeftCenter


grid4.Row = 0

grid4.Col = 1
grid4.Text = "RecHDR"

grid4.Col = 2
grid4.Text = "RecDTL"

grid4.Col = 3
grid4.Text = "PolHDR"

grid4.Col = 4
grid4.Text = "Date"

grid4.Col = 5
grid4.Text = "Amount"

grid4.Col = 6
grid4.Text = "Id Company"

grid4.Col = 7
grid4.Text = "Id Program"


grid4.Col = 8
grid4.Text = "Program"


grid4.Col = 9
grid4.Text = "Company Name"


grid4.Col = 10
grid4.Text = "Pol.Number"

grid4.Col = 11
grid4.Text = "Id Cust"

grid4.Col = 12
grid4.Text = "First name"

grid4.Col = 13
grid4.Text = "Last name"



grid4.FixedRows = 1
grid4.FixedCols = 1
End Sub


 Public Function revisa_compania(X As String)
On Error Resume Next

r$ = UCase(X)
Erase ID_COMPANY1


pos = InStr(1, r$, "ALLIANCE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 1   ' alliance
              ID_COMPANY1(1) = 11  ' kemper
              ID_COMPANY1(2) = 10  ' infinity
              
              SHORTNAME$ = "ALLIANCE"
              
              Exit Function
        
    End If
    
    
    pos = InStr(1, r$, "ANCHOR")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 3
              ID_COMPANY1(1) = 48
        
              SHORTNAME$ = "ANCHOR"
              Exit Function
        
    End If
    
    
     pos = InStr(1, r$, "ARROW")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 47
              
        
              SHORTNAME$ = "ARROWHEAD"
              Exit Function
        
    End If
    
    
    pos = InStr(1, r$, "ASPIRE")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 4
              
              SHORTNAME$ = "ASPIRE"
              Exit Function
              
    End If
    
    
    
     pos = InStr(1, r$, "ALLSTAR")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 2
              ID_COMPANY1(1) = 31
        
              SHORTNAME$ = "ALLSTAR"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "AGIS")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 31
              
        
              SHORTNAME$ = "BLUEFIRE"
              Exit Function
    End If
       
       
    pos = InStr(1, r$, "AEGIS")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 49
              
        
              SHORTNAME$ = "AEGIS"
              Exit Function
              
    End If
    
    
    pos = InStr(1, r$, "AMERICAN")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 73
              
        
              SHORTNAME$ = "AMERICAN COLLECTORS"
              Exit Function
              
    End If
    
    
    
    pos = InStr(1, r$, "ACCEPTANCE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 1095
              
        
              SHORTNAME$ = "ACCEPTANCE INSURANCE"
              Exit Function
              
    End If
    
    
    
    
    
          
    
    pos = InStr(1, r$, "BRIDGER")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 34
              
        
              SHORTNAME$ = "BRIDGER"
              Exit Function
              
    End If
    
    
    pos = InStr(1, r$, "BRISTOL")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 5
                      
              SHORTNAME$ = "BRISTOL"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "BLUEFIRE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 31
                      
              SHORTNAME$ = "BLUEFIRE"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "CASUAL")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 29
              
          
              SHORTNAME$ = "SCOTTISH AMERICAN"
              Exit Function
    End If
    
     pos = InStr(1, r$, "CYPRESS")  ' SCOTTISH AMERICAN
    If pos > 0 Then
        
              ID_COMPANY1(0) = 29
              ID_COMPANY1(1) = 27
        
              SHORTNAME$ = "SCOTTISH AMERICAN"
              Exit Function
    End If
    
    pos = InStr(1, r$, "CARNEGIE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 7
                      
              SHORTNAME$ = "CARNEGIE"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "COAST NATIONAL")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 5
                      
              SHORTNAME$ = "BRISTOL"
              Exit Function
    End If
    
        
        
         
        
        
     pos = InStr(1, r$, "COMMERCE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 8
              ID_COMPANY1(1) = 43
        
              SHORTNAME$ = "COMMERCE WEST"
              Exit Function
              
    End If
    
    
    
    
    
    
    pos = InStr(1, r$, "CENTURY")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 13
                      
              SHORTNAME$ = "MULTI-STATE"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "DAIRYLAND")
    If pos > 0 Then
        
        
              ID_COMPANY1(0) = 9
                      
              SHORTNAME$ = "DAIRYLAND"
              Exit Function
    End If
    
    
    
    
    pos = InStr(1, r$, "DB PREM")
    If pos > 0 Then
              ID_COMPANY1(0) = 49
              
        
              SHORTNAME$ = "AEGIS"
              Exit Function
    End If
    
    
    
    
    pos = InStr(1, r$, "DEBIT INCLINE")
    If pos > 0 Then
              ID_COMPANY1(0) = 34
              
        
              SHORTNAME$ = "BRIDGER"
              Exit Function
    End If
    
    
    
    
    pos = InStr(1, r$, "EVANSTON")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 30
              
        
              SHORTNAME$ = "ATM-AMERICAN TEAM MANAGERS"
              Exit Function
    End If
    
    
    
     pos = InStr(1, r$, "EVEREST NATIONAL")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 47
              
        
              SHORTNAME$ = "ARROWHEAD"
              Exit Function
    End If
    
    
      pos = InStr(1, r$, "EQUITY")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 2
              
        
              SHORTNAME$ = "ALLSTAR"
              Exit Function
    End If
    
    
    
     pos = InStr(1, r$, "FOREMOST")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 1074
              
        
              SHORTNAME$ = "FOREMOST INSURANCE COMPANY"
              Exit Function
    End If
    
    
    
    pos = InStr(1, r$, "HIPPO")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 1091
              
        
              SHORTNAME$ = "HIPPO"
              Exit Function
    End If
    
    
    
    
      pos = InStr(1, r$, "INFINITY")
    If pos > 0 Then
        
              'ID_COMPANY1(0) = 10
              
              ID_COMPANY1(0) = 1   ' alliance
              ID_COMPANY1(1) = 11  ' kemper
              ID_COMPANY1(2) = 10  ' infinity
              
        
              SHORTNAME$ = "INFINITY"
              Exit Function
    End If
    
  
        
    
     pos = InStr(1, r$, "KEMPER")
    If pos > 0 Then
        
              'ID_COMPANY1(0) = 11
              'ID_COMPANY1(1) = 1
              
              ID_COMPANY1(0) = 1   ' alliance
              ID_COMPANY1(1) = 11  ' kemper
              ID_COMPANY1(2) = 10  ' infinity
        
              SHORTNAME$ = "KEMPER"
              Exit Function
    End If
    
        
    pos = InStr(1, r$, "KNIGHT")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 34
                            
              SHORTNAME$ = "BRIDGER"
              Exit Function
              
    End If
    
    
    pos = InStr(1, r$, "MAPFRE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 8
              ID_COMPANY1(1) = 43
              
              SHORTNAME$ = "MAPFRE"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "MULTISTATE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 13
                            
               SHORTNAME$ = "MULTI-STATE"
               Exit Function
    End If
    
    
     pos = InStr(1, r$, "MACAFEE")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 55
              
              
              SHORTNAME$ = "MACAFEE"
              Exit Function
    End If
          
    
     pos = InStr(1, r$, "MILLENIUM")
    If pos > 0 Then
    
              'ID_COMPANY1(0) = 11
              'ID_COMPANY1(1) = 1
              
              ID_COMPANY1(0) = 1   ' alliance
              ID_COMPANY1(1) = 11  ' kemper
              ID_COMPANY1(2) = 10  ' infinity
        
              SHORTNAME$ = "KEMPER"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "MOTORCLUB")
    If pos > 0 Then
          
              ID_COMPANY1(0) = 15
                            
              SHORTNAME$ = "NATIONS INSURANCE"
              Exit Function
    End If
    
           
       
     pos = InStr(1, r$, "MCGRAW")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 12
              ID_COMPANY1(1) = 53
              
              SHORTNAME$ = "MCGRAW"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "MERIPLAN")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 47
              
              SHORTNAME$ = "ARROWHEAD"
              Exit Function
    End If
    
       
   
    pos = InStr(1, r$, "NATIONAL")
    If pos > 0 Then
               
              ID_COMPANY1(0) = 14
                      
              SHORTNAME$ = "NATIONAL GENERAL"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "NATIONS")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 15
              ID_COMPANY1(1) = 13
              ID_COMPANY1(2) = 31
              
              SHORTNAME$ = "NATIONS INSURANCE"
              Exit Function
    End If
   
   
       
    pos = InStr(1, r$, "OCEAN")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 31
                            
              SHORTNAME$ = "BLUEFIRE"
              Exit Function
    End If
       
       
    pos = InStr(1, r$, "PACIFIC")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 12
              ID_COMPANY1(1) = 53
              
              SHORTNAME$ = "ANCHOR"
              Exit Function
    End If
    
   
   
    pos = InStr(1, r$, "PERMANENT")  ' ANCHOR
    If pos > 0 Then
        
              SHORTNAME$ = "ANCHOR"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "PERMAN")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 48
              ID_COMPANY1(1) = 3
              
              SHORTNAME$ = "PAC STAR GENERAL"
              Exit Function
    End If
    
  
  
    pos = InStr(1, r$, "PRIME")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 31
                            
              SHORTNAME$ = "BLUEFIRE"
              Exit Function
    End If
       
     
    pos = InStr(1, r$, "PRONTO")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 24
                      
              SHORTNAME$ = "PRONTO INSURANCE"
              Exit Function
    End If
      
    
     pos = InStr(1, r$, "PAC STAR")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 48
        
              SHORTNAME$ = "PAC STAR GENERAL"
              Exit Function
    End If
    
           
    
     pos = InStr(1, r$, "PACIFIC")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 53
              ID_COMPANY1(1) = 12
              
              SHORTNAME$ = "MCGRAW"
              Exit Function
    End If
     
    
    
     pos = InStr(1, r$, "QBE")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 47
              
              SHORTNAME$ = "ARROWHEAD"
              Exit Function
    End If
    
    
    
    pos = InStr(1, r$, "QUIC")
    If pos > 0 Then
        
        
        ID_COMPANY1(0) = 1097
        SHORTNAME$ = "QUALITAS INSURANCE COMPANY"
              '
        Exit Function
    
    End If
         
         
    
         
         
     pos = InStr(1, r$, "RELIANT")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 57
              
              SHORTNAME$ = "RIC"
              Exit Function
    End If
    
    
    
    pos = InStr(1, r$, "ROBERT MORENO")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 16
        
              SHORTNAME$ = "RMIS"
              Exit Function
    End If
    
    pos = InStr(1, r$, "SAFE AUTO")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 34
              
              SHORTNAME$ = "BRIDGER"
              Exit Function
    End If
    
        
    pos = InStr(1, r$, "SUN COAST")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 18
        
              SHORTNAME$ = "SUNCOAST"
              Exit Function
    End If
    
   
    
     pos = InStr(1, r$, "SAFEWAY")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 17
              
              SHORTNAME$ = "SAFEWAY"
              Exit Function
       
    End If
    
    
     pos = InStr(1, r$, "SCOTT")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 27
              ID_COMPANY1(1) = 29
        
              SHORTNAME$ = UCase("Scottish American")
              Exit Function
    End If
     
   
   pos = InStr(1, r$, "STARR")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 47
              
              SHORTNAME$ = "ARROWHEAD"
              Exit Function
    End If
    
    
   pos = InStr(1, r$, "STONEWOOD")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 31
              
              SHORTNAME$ = "BLUEFIRE"
              Exit Function
    End If
    
    
     pos = InStr(1, r$, "SAFEBUILT")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 1077
              
              SHORTNAME$ = "SIS INSURE"
              Exit Function
    End If
    
    
    pos = InStr(1, r$, "TAPCO")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 72
              
              SHORTNAME$ = "TAPCO INSURANCE"
              Exit Function
    End If
   
    
    
     pos = InStr(1, r$, "WESTERN")
    If pos > 0 Then
        
              ID_COMPANY1(0) = 19
              
              SHORTNAME$ = "WESTERN"
              Exit Function
    End If
    
    
    
    
     pos = InStr(1, r$, "WORKMEN")
    If pos > 0 Then
    
              ID_COMPANY1(0) = 20
    
              SHORTNAME$ = "WORKMENS"
              Exit Function
    End If
    
   
    

End Function


Public Sub asigna_IDconciliation()
On Error Resume Next


End Sub

Public Sub ajusta_mes()
On Error Resume Next

' revisa en los primeros 50 para ver la fecha mayor

If Grid1.Rows < 50 Then
  limite_max = Grid1.Rows
Else
  limite_max = 50
End If

mes_mayor = 0
ano_mayor = 0
For t = 1 To limite_max
  Grid1.Row = t
  Grid1.Col = 6
  f$ = Format(Grid1.Text, "mm/dd/yyyy")
  
  If Val(Right(f$, 4)) > ano_mayor Then
    ano_mayor = Val(Right(f$, 4))
    
  End If
  
  
  If Val(Left(f$, 2)) > mes_mayor And (Val(Right(f$, 4)) >= ano_mayor) Then
     mes_mayor = Val(Left(f$, 2))
  ElseIf Val(Left(f$, 2)) < mes_mayor And (Val(Right(f$, 4)) >= ano_mayor) Then
     mes_mayor = Val(Left(f$, 2))
     
  End If
  
Next t

mes_actual = mes_mayor
ano_actual = ano_mayor




pos = InStr(1, f$, "/")

mes_act = mes_actual


mm$ = Left$(f$, pos - 1)
btnmes(Val(mm$) - 1).Value = True
btnmes_Click (Val(mm$) - 1)


aa = Val(Right(f$, 4))

For t = 0 To cboyear.ListCount - 1
  If cboyear.List(t) = aa Then
     cboyear.ListIndex = t
     Exit For
  End If
Next t


If mes_actual = 0 Then
   mes_actual = mes_act
End If




mes_actual = mes_mayor
ano_actual = ano_mayor


Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dias_actual = 31
Case 4, 6, 9, 11
  dias_actual = 30
Case 2
   cant = (ano_actual / 4)
   residuo = cant - Int(cant)
   If residuo = 0 Then
      dias_actual = 29
   Else
      dias_actual = 28
   End If
End Select


If mes_actual > 1 Then

  fecha_rango1$ = Format(ano_actual, "00") + "-" + Format(mes_actual - 1, "00") + "-24"
  'fecha_rango1$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-01"
  fecha_rango2$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-" + Format(dias_actual, "00")
  
Else

  fecha_rango1$ = Format(ano_actual - 1, "00") + "-" + "12" + "-24"
  'fecha_rango1$ = Format(ano_actual, "00") + "-" + "01" + "-01"
  fecha_rango2$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-" + Format(dias_actual, "00")


End If



txtfecha(0).Text = Format(fecha_rango1$, "mm/dd/yyyy")
txtfecha(1).Text = Format(fecha_rango2$, "mm/dd/yyyy")

End Sub

Public Sub ASIGNA_UNOS()
On Error Resume Next

For t = 1 To Grid1.Rows - 1
  Grid1.Row = t
  Grid1.Col = 22
  Grid1.Text = "1"
Next t


For t = 1 To Grid3.Rows - 1
  Grid3.Row = t
  Grid3.Col = 22
  Grid3.Text = "1"
Next t

  
End Sub

Public Sub carga_grid5()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset


Dim rsVar As Variant
Dim i As Integer

' carga el grid5
grid2.Clear
grid5.Clear









If op_mes(1).Value = True Then
  fecha_rango1x$ = fecha_rango1$
  fecha_rango2x$ = fecha_rango2$
Else
   If mes_actual > 1 Then
       fecha_rango1x$ = Format(ano_actual, "00") + "-" + Format(mes_actual, "00") + "-01"
   Else
       fecha_rango1x$ = Format(ano_actual, "00") + "-" + "01" + "-01"
   End If

   fecha_rango2x$ = fecha_rango2$
End If








Set Rs = New ADODB.Recordset
    
    
   'sSelect = "select * from ConciliationBankRec"
    
   sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL, conci.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, " & _
   "recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber , " & _
   "rechdr.IdCustomer, cust.firstname, cust.lastname1, conci.Clear from ReceiptsHDR rechdr " & _
   "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
   "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
   "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
   "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
   "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
   "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
   "left join ConciliationBankRec conci on conci.IdReceiptDTL=recdtl.IdReceiptDTL " & _
   "where rechdr.date>='" + fecha_rango1x$ + "' and rechdr.date<='" + fecha_rango2x$ + "' " & _
   "and iicat.IsPremium=1 and conci.IdReceiptDTL is null and rechdr.void=0"
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   grid2.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   grid2.cols = Rs.Fields.Count
    
    
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      grid2.TextMatrix(0, i) = Rs.Fields(i).Name
   Next

   grid2.Row = 1
   grid2.Col = 0

   ' Set range of cells in the grid
   grid2.RowSel = grid2.Rows - 1
   grid2.ColSel = grid2.cols - 1
   grid2.clip = rsVar

   ' Reset the grid's selected range of cells
   grid2.RowSel = grid2.Row
   grid2.ColSel = grid2.Col

   Rs.Close

   Set Rs = Nothing

   grid5.Visible = False

   grid5.Rows = grid2.Rows
   grid5.cols = grid2.cols
   
   
      grid2.Row = 0
      grid5.Row = 0
      For Y = 1 To grid2.cols - 1
        grid2.Col = Y - 1
        grid5.Col = Y
        grid5.Text = grid2.Text
      Next Y
   
   
   
   
   For t = 1 To grid2.Rows - 1
      grid2.Row = t
      grid5.Row = t
      
      grid5.Col = 0
      grid5.Text = Format(t, "###0")
      
      For Y = 1 To grid2.cols - 1
         grid2.Col = Y - 1
         grid5.Col = Y
         grid5.Text = grid2.Text
      Next Y
   Next t


   grid5.Visible = True
   

 lbltotal5.Caption = Format(grid5.Rows - 1, "###,##0")
   







End Sub

Public Sub enca_grid5()
'On Error Resume Next




grid5.ColWidth(0) = 800
grid5.ColAlignment(0) = flexAlignLeftCenter

grid5.ColWidth(1) = 1300 ' rec hdr
grid5.ColAlignment(1) = flexAlignCenterCenter

grid5.ColWidth(2) = 1300  ' rec dtl
grid5.ColAlignment(2) = flexAlignCenterCenter

grid5.ColWidth(3) = 1300   ' rec dtl
grid5.ColAlignment(3) = flexAlignCenterCenter

grid5.ColWidth(4) = 1300 ' id policies hdr
grid5.ColAlignment(4) = flexAlignCenterCenter

grid5.ColWidth(5) = 2400   ' date
grid5.ColAlignment(5) = flexAlignLeftCenter

grid5.ColWidth(6) = 1200   ' amount
grid5.ColAlignment(6) = flexAlignRightCenter

grid5.ColWidth(7) = 1200   ' id company
grid5.ColAlignment(7) = flexAlignCenterCenter

grid5.ColWidth(8) = 1600   ' id prog
grid5.ColAlignment(8) = flexAlignCenterCenter

grid5.ColWidth(9) = 2800 ' program name
grid5.ColAlignment(9) = flexAlignLeftCenter

grid5.ColWidth(10) = 2400  ' company name
grid5.ColAlignment(10) = flexAlignLeftCenter

grid5.ColWidth(11) = 2200  ' policy number
grid5.ColAlignment(11) = flexAlignLeftCenter

grid5.ColWidth(12) = 1100  ' id cust
grid5.ColAlignment(12) = flexAlignCenterCenter

grid5.ColWidth(13) = 1600  ' first name
grid5.ColAlignment(13) = flexAlignLeftCenter

grid5.ColWidth(14) = 1600  ' last name
grid5.ColAlignment(13) = flexAlignLeftCenter


grid5.Row = 0

grid5.Col = 1
grid5.Text = "RecHDR"

grid5.Col = 2
grid5.Text = "RecDTL"

grid5.Col = 3
grid5.Text = "RecDTL"

grid5.Col = 4
grid5.Text = "IDpoliciesHDR"

grid5.Col = 5
grid5.Text = "Date"

grid5.Col = 6
grid5.Text = "Amount"

grid5.Col = 7
grid5.Text = "Id Company"


grid5.Col = 8
grid5.Text = "ID Program"


grid5.Col = 9
grid5.Text = "Program Name"


grid5.Col = 10
grid5.Text = "Company name"

grid5.Col = 11
grid5.Text = "Policy number"

grid5.Col = 12
grid5.Text = "ID Customer"

grid5.Col = 13
grid5.Text = "First name"

grid5.Col = 14
grid5.Text = "Last name"


grid5.FixedRows = 1
grid5.FixedCols = 1
End Sub

Public Function verifica_existencia_de_poliza(poliza_a_verificar As String)
On Error Resume Next

Grid1.Visible = False
existe = 0
For Y = 1 To Grid1.Rows - 1
  Grid1.Row = Y
  Grid1.Col = 8
  X$ = Grid1.Text
  
  If UCase(X$) = UCase(poliza_a_verificar) Then
    existe = 1
    Exit For
  End If
  
Next Y

poliza_a_verificar = Format(existe, "0")

End Function

Public Sub carga_programas()
On Error Resume Next
cboprogram.Clear

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
  grid2.Clear
    
  sSelect = "select programname, idprogram from ProgramsCatalog"
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid2.DataSource = Rs
                         
    Rs.Close
    
    
For t = 1 To grid2.Rows - 1
   grid2.Row = t
   grid2.Col = 1
   programa$ = grid2.Text
    
   grid2.Col = 2
   id_programa$ = grid2.Text
   
   a$ = programa$ + Space(30) + id_programa$
   cboprogram.AddItem a$
Next t

   
   
End Sub

Public Sub calcula_total_multiple()
On Error Resume Next

Total = 0
For t = 1 To grid4.Rows - 1
   grid4.Row = t
   grid4.Col = 5
   Total = Total + Val(grid4.Text)
Next t

lbltotal_cantidad.Caption = Format(Total, "$###,##0.00")

lbl_diferencia.Caption = Format(Total - Val(txtcantidad.Text), "$###,##0.00")

lbltotal_lista.Caption = Format(List2.ListCount, "###0")


If (Val(Total) - Val(txtcantidad.Text)) = 0 And Val(Format(lbltotal_cantidad.Caption, "0000.00")) > 0 Then
  pulgar.Visible = True
Else
  pulgar.Visible = False
End If

calcula_total_grid4


End Sub



Public Sub carga_combo_oficinas()
On Error Resume Next
cbo_oficina.Clear

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
  grid2.Clear
    
  sSelect = "select office, idoffice from OfficesCatalog"
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid2.DataSource = Rs
                         
    Rs.Close
    
    
For t = 1 To grid2.Rows - 1
   grid2.Row = t
   grid2.Col = 1
   oficina$ = grid2.Text
    
   grid2.Col = 2
   id_oficina$ = grid2.Text
   
   a$ = oficina$ + Space(30) + id_oficina$
   cbo_oficina.AddItem a$
Next t





End Sub

Public Sub Resetea_multiple()
On Error Resume Next
lbltotal_cantidad.Caption = "$"
lbltotal_needed.Caption = "$"
lbl_diferencia.Caption = "$"
lbl_total_marcado.Caption = "$"

List2.Clear
op_day(0).Value = True


End Sub

Public Sub carga_lista()
On Error Resume Next
List2.Clear
For t = 1 To grid4.Rows - 1
  grid4.Row = t
  grid4.Col = 5
  cant = Val(grid4.Text)
  
  List2.AddItem Format(grid4.Text, "0000.00") + Space(20) + Str(t)
  
Next t

lbltotal_lista.Caption = List2.ListCount


End Sub

Public Sub calcula_total_grid4()
 ' calcula el total
    On Error Resume Next
    
    amount = 0
    For t = 0 To grid4.Rows - 1
       grid4.Row = t
       grid4.Col = 5
       c = Val(grid4.Text)
       
       amount = amount + c
    Next t
    lblamount_total.Caption = Format(amount, "$###,##0.00")
    
    
End Sub

Public Sub ajusta_tabla()
On Error Resume Next

 grid6.Clear
 grid6.Rows = List2.ListCount + 1
 

grid6.Row = 0
grid4.Row = 0
For t = 0 To grid4.cols - 1
   grid4.Col = t
   grid6.Col = t
   grid6.Text = grid4.Text
Next t



 c = 0
    For t = 0 To List2.ListCount - 1
   
     'If List2.Selected(t) = False Then
       X1 = Val(Left(List2.List(t), 10))
       c = c + 1
       For Y = 1 To grid4.Rows - 1
          grid4.Row = Y
          grid4.Col = 5
          X2 = Val(grid4.Text)
          If X1 = X2 Then
                         
              grid6.Row = c
              For z = 0 To grid4.cols - 1
                 grid6.Col = z
                 grid4.Col = z
                 grid6.Text = grid4.Text
              Next z
                         
              
             Exit For
          End If
       Next Y
     'End If
     

    Next t
          
          
          
  grid4.Clear
  grid4.Rows = grid6.Rows
  
  
  For t = 0 To grid6.Rows - 1
     grid6.Row = t
     grid4.Row = t
     
     For Y = 0 To grid6.cols - 1
        grid6.Col = Y
        grid4.Col = Y
        grid4.Text = grid6.Text
     Next Y
  Next t
          
          
 'c = 0
 '   For t = 0 To List2.ListCount - 1
   
     'If List2.Selected(t) = False Then
 '      X1 = Val(Left(List2.List(t), 10))
 '      c = c + 1
 '      For Y = 1 To grid4.Rows - 1
 '         grid4.Row = Y
 '         grid4.Col = 5
 '         X2 = Val(grid4.Text)
 '         If X1 = X2 Then
 '            grid4.RemoveItem Y
             
 '            Exit For
 '         End If
 '      Next Y
     'End If
     

  '  Next t
          
          
          
    
    
    
    For t = 1 To grid4.Rows - 1
       grid4.Row = t
       grid4.Col = 0
       grid4.Text = t
    Next t
    
    lbltotal4.Caption = grid4.Rows - 1
    
     calcula_total_multiple
    
  
    lbltotal_lista.Caption = List2.ListCount
    
End Sub

Public Sub cambia_filas()
On Error Resume Next
grid4.Rows = grid4.Rows + grid6.Rows - 1

End Sub

Public Sub verifica_polizas()
On Error Resume Next

' ************************************************************
  

  
  
 Dim sSelect As String
    
 Dim Rs As ADODB.Recordset
    
   Image1.Visible = False
  contador = 0
  barra.Visible = True
  barra.Min = 1
  
  ' establece el color de la barra en rojo
  Color_Progreso barra.hwnd, &H2B36FF
    
  
  
  Grid3.Rows = Grid1.Rows
  Grid3.Clear
  
  lineas = 0
  Erase recibos$
  
Grid1.Row = 0
Grid3.cols = Grid1.cols


Grid1.Visible = False
Grid3.Visible = False
grid2.Visible = False
Grid1.Refresh


'Grid1.Visible = True
'Grid3.Visible = True

'



' asigna encabezados
For w = 1 To Grid1.cols - 1
   Grid3.Row = 0
   Grid1.Col = w
   Grid3.Col = w
   Grid3.Text = Grid1.Text
Next w
   
   

grid2.Clear

Timer1.Enabled = True
seg = 0
crea_registro = 0

barra.Max = Grid1.Rows

For t = 1 To Grid1.Rows - 1

    checada = 0
    
    contador = contador + 1
    
    barra.Value = contador
    
    crea_registro = 0
    estatus_poliza = 0
    
    Grid1.Row = t
    
    poliza$ = ""
    GUIONES = 0
    
    Grid1.Row = t
    Grid1.Col = 8
    poliza$ = UCase(Grid1.Text)
    
    poliza_banco$ = UCase(poliza$)
    
   ' If poliza$ = "MIL3919618" Then
   '    Stop
   ' End If
        
    Grid1.Col = 3
    debito = Val(Grid1.Text)
    
    Grid1.Col = 4
    credito = Val(Grid1.Text)
    
    Grid1.Col = 6
    fecha_pagada$ = Grid1.Text
    
    
    
    If debito <> 0 Then
       cantidad = debito
    ElseIf credito <> 0 Then
       
       GoTo viene_de_credito
    End If
    
    
     'If cantidad = 20000 Then Stop
    
     'If cantidad = 84 Then Stop
    
     'If cantidad = 468.95 Then Stop
     
    ' If cantidad = 85.99 Then Stop
     
    ' If cantidad = 741.63 Then Stop
     
     'If cantidad = 315.78 Then Stop
     
     'If cantidad = 197.24 Then Stop
    
    
    
   
    
    ' genera fecha inicial de busqueda
    ' ******************************************************************
    
            'mes_actual = Val(Left(Format(Now, "mm/dd/yyyy"), 2))
            
            
            Select Case mes_actual
            Case 1, 3, 5, 7, 8, 10, 12
               dias_actual = 31
            Case 4, 6, 9, 11
               dias_actual = 30
            Case 2
               cant = (ano_actual / 4)
               residuo = cant - Int(cant)
               If residuo = 0 Then
                   dias_actual = 29
               Else
                   dias_actual = 28
               End If
            End Select
            
            
            
            Select Case Left(Format(fecha_pagada$, "mm/dd/yyyy"), 2)
                 Case 1, 3, 5, 7, 8, 10, 12
                   dia = 31
                 Case 4, 6, 9, 11
                   dia = 30
                 Case 2
                   cant = (ano_actual / 4)
                   residuo = cant - Int(cant)
                   If residuo = 0 Then
                      dia = 29
                   Else
                      dia = 28
                   End If
                 
            End Select
                 
                 
            fecha_pagada$ = Format(fecha_pagada$, "mm/dd/yyyy")
                           
                           
            If Val(Mid(fecha_pagada$, 4, 2)) < dia Then
            
               fecha_pagada2$ = Left(fecha_pagada$, 2) + "/" + Format(Val(Mid(fecha_pagada$, 4, 2)) + 1, "00") + "/" + Right(fecha_pagada$, 4)
            
            Else
               fecha_pagada2$ = fecha_pagada$
            End If
                      
            
                           
            



            If mes_actual >= 1 And mes_actual <= 12 Then
        
                  dia_x1 = Format(fecha_pagada$, "y")
                  
                  ' dia_x1 = Val(dia_x2) - 10
                     
                 
                  m1 = 31      ' Enero
                 
                  ano_actual = Val(Right(fecha_pagada$, 4))
                  cant = (ano_actual / 4)
                  residuo = cant - Int(cant)
                  If residuo = 0 Then
                     dias_actualx = 29
                  Else
                     dias_actualx = 28
                  End If
                   
                  m2 = m1 + dias_actualx   ' Febrero
                  m3 = m2 + 31    ' marzo
                  m4 = m3 + 30    ' abril
                  m5 = m4 + 31    ' mayo
                  m6 = m5 + 30    ' junio
                  m7 = m6 + 31    ' julio
                  m8 = m7 + 31    ' agosto
                  m9 = m8 + 30    ' septiembre
                  m10 = m9 + 31   ' octubre
                  m11 = m10 + 30  ' noviembre
                  m12 = m11 + 31  ' diciembre
                 
                  dia = dias_actual
                 
                 
                  Select Case Val(dia_x1)
                  Case Is <= m1
                        If dia_x1 > 8 Then
                             'dia = dia_x1 - 8
                             'Fecha_inicio_busqueda$ = "01" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                      
                            ' If rango = 0 Then
                                    fecha_21_mes = "12/21/" + Format(ano_actual - 1, "0000")
                            ' ElseIf rango = 2 Then
                                    fecha_11_mes = "12/11/" + Format(ano_actual - 1, "0000")
                            ' Else
                                    fecha_01_mes = "01/01/" + Format(ano_actual, "0000")
                            ' End If
                        End If
                   
                  Case Is <= m2
                        'dia = dia_x1 - m1
                        'Fecha_inicio_busqueda$ = "02" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                   
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "01/21/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "01/11/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "02/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                  Case Is <= m3
                        'dia = dia_x1 - m2
                        'Fecha_inicio_busqueda$ = "03" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                   
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "02/18/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "02/08/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "03/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                  Case Is <= m4
                        'dia = dia_x1 - m3
                        'Fecha_inicio_busqueda$ = "04" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "03/21/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "03/11/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "04/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                  Case Is <= m5
                        'dia = dia_x1 - m4
                        'Fecha_inicio_busqueda$ = "05" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "04/20/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "04/11/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "05/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                  Case Is <= m6
                        'dia = dia_x1 - m5
                        'Fecha_inicio_busqueda$ = "06" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "05/21/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "05/11/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "06/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                  Case Is <= m7
                        'dia = dia_x1 - m6
                        'Fecha_inicio_busqueda$ = "07" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                        'If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "06/20/" + Format(ano_actual, "0000")
                        'ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "06/11/" + Format(ano_actual, "0000")
                        'Else
                                    fecha_01_mes = "07/01/" + Format(ano_actual, "0000")
                        'End If
                                      
                  Case Is <= m8
                        'dia = dia_x1 - m7
                        'Fecha_inicio_busqueda$ = "08" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                        'If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "07/21/" + Format(ano_actual, "0000")
                        'ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "07/11/" + Format(ano_actual, "0000")
                        'Else
                                    fecha_01_mes = "08/01/" + Format(ano_actual, "0000")
                        'End If
                   
                  Case Is <= m9
                        'dia = dia_x1 - m8
                        'Fecha_inicio_busqueda$ = "09" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                        'If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "08/20/" + Format(ano_actual, "0000")
                        'ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "08/11/" + Format(ano_actual, "0000")
                        'Else
                                    fecha_01_mes = "09/01/" + Format(ano_actual, "0000")
                        'End If
                   
                  Case Is <= m10
                        'dia = dia_x1 - m9
                         '           Fecha_inicio_busqueda$ = "10" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        'If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "09/20/" + Format(ano_actual, "0000")
                        'ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "09/11/" + Format(ano_actual, "0000")
                        'Else
                                    fecha_01_mes = "10/01/" + Format(ano_actual, "0000")
                        'End If
                   
                  Case Is <= m11
                        'dia = dia_x1 - m10
                        'Fecha_inicio_busqueda$ = "11" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                        'If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "10/21/" + Format(ano_actual, "0000")
                        'ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "10/11/" + Format(ano_actual, "0000")
                        'Else
                                    fecha_01_mes = "11/01/" + Format(ano_actual, "0000")
                        'End If
                   
                 Case Is <= m12
                        'dia = dia_x1 - m11
                        'Fecha_inicio_busqueda$ = "12" + "/" + Format(dia, "00") + "/" + Format(ano_actual, "0000")
                        
                       ' If op_fecha_carga(0).Value = True Then
                                    fecha_21_mes = "11/20/" + Format(ano_actual, "0000")
                       ' ElseIf op_fecha_carga(2).Value = True Then
                                    fecha_11_mes = "11/11/" + Format(ano_actual, "0000")
                       ' Else
                                    fecha_01_mes = "12/01/" + Format(ano_actual, "0000")
                       ' End If
                   
                 End Select
                 
                                  
                 
                
                
                
   
            'ElseIf mes_actual = 1 Then
                 
             '    ano_anterior = cboyear.List(cboyear.ListIndex) - 1 'Val(Format(Now, "yyyy")) - 1
                 
              '   If op_fecha_carga(0).Value = True Then
                      
               '       fecha_21_mes = "12/20/" + Format(ano_anterior, "0000")
                      
               '  ElseIf op_fecha_carga(2).Value = True Then
                '      fecha_11_mes = "12/11/" + Format(ano_anterior, "0000")
                      
               '  Else
                '       fecha_01_mes = "01/01/" + Format(ano_anterior + 1, "0000")
               '  End If
                 
                 
                 
            'ElseIf mes_actual = 12 Then
           
             '    ano_anterior = cboyear.List(cboyear.ListIndex)
                 
              '   If op_fecha_carga(0).Value = True Then
               '       fecha_21_mes = "11/20/" + Format(ano_anterior, "0000")
              '   ElseIf op_fecha_carga(2).Value = True Then
               '       Fecha_inicio_busqueda$ = "11/11/" + Format(ano_anterior, "0000")
               '  Else
                '      Fecha_inicio_busqueda$ = "12/01/" + Format(ano_anterior, "0000")
                ' End If
                            

            End If
    
   
    ' ******************************************************************
    
    
    Fecha_inicio_busqueda$ = fecha_01_mes
    
revisa_otravez:
    
    
    
    Grid1.Col = 7
    concepto$ = Grid1.Text
    
    Grid1.Col = 11
    Grid1.Text = ""
    
    
    
    
    lblmsg2.Caption = "Processing " + Format(t, "###0") + " of " + Format(Grid1.Rows - 1, "###0")
    lblmsg2.Refresh
    openforms = DoEvents
    
    'If poliza$ = "" Then GoTo salta
    
    
   
    
    
   
    r$ = ""

    r$ = UCase(concepto$)
    SHORTNAME$ = ""
    
    
     p$ = ""
    
   
    
    Grid1.Row = t
    Grid1.Col = 7
    r$ = UCase(Grid1.Text)
    
    
    
    
    VERIFICA_COMPANIA = 0
    
    X1$ = revisa_compania(r$)
    
    ' shortname$ tiene el nombre de la compañia
    
    
    If SHORTNAME$ = "ACCEPTANCE INSURANCE" Then
       'Stop
    End If
    
    
    
   
continua_aqui:
   
   
   
     Set Rs = New ADODB.Recordset
    'Checa_status
    
   
    If poliza$ <> "" Then
     idcompany$ = ""
     sSelect = "select idcompany from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
     Rs.Open sSelect, base, adOpenUnspecified
    
     idcompany$ = Rs(0)
                         
     Rs.Close
     
    End If
   
   
   
   
   
  poliza2$ = "..."
  If Val(idcompany$) = 10 Then
     poliza2$ = Left(poliza$, 3) + "-" + Mid$(poliza$, 4, 5) + "-" + Mid$(poliza$, 9, 4) + "-" + Right(poliza$, 2)
  End If
   
   void = 0
   
   Set Rs = New ADODB.Recordset
    
    
    
  ' asigna companias con id
   g$ = "polhdr.idcompany="
  compania$ = ""
  
  If ID_COMPANY1(0) > 0 Then
    c$ = g$ + "'" + Format(ID_COMPANY1(0), "###0") + "'"
    compania$ = compania$ + c$
  End If
  
  If ID_COMPANY1(1) > 0 Then
    c$ = g$ + "'" + Format(ID_COMPANY1(1), "###0") + "'"
    compania$ = compania$ + " or " + c$
  End If
    
  If ID_COMPANY1(2) > 0 Then
    c$ = g$ + "'" + Format(ID_COMPANY1(2), "###0") + "'"
    compania$ = compania$ + " or " + c$
  End If
    
  If ID_COMPANY1(3) > 0 Then
    c$ = g$ + "'" + Format(ID_COMPANY1(3), "###0") + "'"
    compania$ = compania$ + " or " + c$
  End If
    
    
    
    
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    
sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName,  polhdr.PolicyNumber, rechdr.IdCustomer,INS.shortname from ReceiptsHDR rechdr " & _
"inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
"inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
"inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
"inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
"inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem "

' verifica condiciones



condicion$ = ""
c1 = 0
If Fecha_inicio_busqueda$ <> "" Then
   condicion$ = "rechdr.date>='" + Fecha_inicio_busqueda$ + "'"
  
   If fecha_pagada2$ <> "" Then
      cond1$ = "rechdr.date<='" + fecha_pagada2$ + "'"
      condicion$ = condicion$ + " and " + cond1$
      
   End If

   c1 = 1
End If


' fecha_24_mes_anterior,   tambien puede ir en la condicion

If c1 = 1 Then
   condicion$ = condicion$ + " and "
End If


If condicion$ <> "" And Right(condicion$, 4) <> "and " Then
  condicion$ = condicion$ + " and "
End If




cantidad1 = cantidad * 0.999
cantidad2 = cantidad * 1.001
cantidad3 = Fix(cantidad)


If cantidad > 0 Then

  'If poliza$ = "" Then
     'cond1$ = "(recdtl.amount='" + Format(cantidad, "#######0.00") + "' or (recdtl.amount>='" + Format(cantidad1, "#######0.00") + "' and recdtl.amount<='" + Format(cantidad2, "#######0.00") + "') or (recdtl.amount='" + Format(cantidad3, "#######0.00") + "'))"
      
      cond1$ = "recdtl.amount='" + Format(cantidad, "#######0.00") + "'"
      condicion$ = condicion$ + cond1$
      If c1 = 1 Then
         condicion$ = condicion$ + " and "
      End If
   
  'Else
  
   '   cond1$ = "(recdtl.amount='" + Format(cantidad, "#######0.00") + "'"
   '   condicion$ = condicion$ + cond1$
   '   If c1 = 1 Then
   '      condicion$ = condicion$ + " or "
   '   End If
   
  
  'End If

End If


If condicion$ <> "" And Right(condicion$, 4) <> "and " Then
  condicion$ = condicion$ + " and "
End If




grid2.Col = 10
encontrado = 0
poliza_con_guiones$ = ""
poliza_sin_guiones$ = poliza$
poliza_con_guiones$ = Left(poliza$, 3) + "-" + Mid$(poliza$, 4, 5) + "-" + Mid$(poliza$, 9, 4) + "-" + Right(poliza$, 3)

If poliza$ <> "" Then
   cond1$ = "(polhdr.PolicyNumber='" + poliza$ + "' or polhdr.PolicyNumber='" + poliza_con_guiones$ + "')"
   condicion$ = condicion$ + cond1$
   If c1 = 1 Then
     condicion$ = condicion$ + " and "
   End If

End If


If condicion$ <> "" And Right(condicion$, 4) <> "and " Then
  condicion$ = condicion$ + " and "
End If


If compania$ <> "" Then
   cond1$ = "(" + compania$ + ")"
   condicion$ = condicion$ + cond1$
   If c1 = 1 Then
     condicion$ = condicion$ + " and "
   End If

End If

If condicion$ <> "" And Right(condicion$, 4) <> "and " Then
  condicion$ = condicion$ + " and "
End If

condicion$ = condicion$ + "rechdr.Active=1 and iicat.IsPremium=1"


sSelect = sSelect + Chr$(13) + "Where " + condicion$


' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


estatus_poliza = 1
    
Rs.Open sSelect, base, adOpenUnspecified
grid2.AllowUserResizing = flexResizeColumns
Set grid2.DataSource = Rs
Rs.Close


valor = 0
If poliza$ <> poliza_banco$ Or UCase(poliza$) = "POLICYNUMBER" Then
       valor = 1
       grid2.Clear
End If
                                                                                                                              
                                            
If grid2.Rows > 1 And valor = 0 Then GoTo saltada_por_encontrada
                                            
                    
                    ' checa con los otros valores
                    checada = checada + 1
                    
                    If checada = 1 Then
                       Fecha_inicio_busqueda$ = fecha_21_mes
                       GoTo revisa_otravez
                    ElseIf checada = 2 Then
                       Fecha_inicio_busqueda$ = fecha_11_mes
                       GoTo revisa_otravez
                    End If
                    
                    
                    
                    
                                                             ' =========================================================================================================
                    
                                                                   If SHORTNAME$ = "ACCEPTANCE INSURANCE" Then
                                                                      comp$ = "1095, 1094"
                                                                   ElseIf SHORTNAME$ = "COMMERCE WEST" Then
                                                                      comp$ = "8"
                                                                   ElseIf SHORTNAME$ = "NATIONAL GENERAL" Then
                                                                      comp$ = "14"
                                                                   ElseIf SHORTNAME$ = "SAFEWAY" Then
                                                                      comp$ = "17"
                                                                   
                                                                   
                                                                   End If
                                                                      
                                                                   
                                                                         fecha_pagada$ = Format(fecha_pagada2$, "mm/dd/yyyy")
                                                                   
                                                                         dia_pagado$ = Mid$(Format(fecha_pagada2$, "mm/dd/yyyy"), 4, 2)
                                                                         mes_pagado$ = Left$(fecha_pagada2$, 2)
                                                                         ano_pagado$ = Right(fecha_pagada2$, 4)
                                                                    
                                                                   
                                                                    
                                                                         For w = 1 To 30
                                                                            
                                                                                                                                                    
                                                                           
                                                                           f$ = ano_pagado$ + "-" + mes_pagado$ + "-" + Format(w, "00")
                                                                           f2$ = ano_pagado$ + "-" + mes_pagado$ + "-" + Format(w + 1, "00")
                                                                            
                                                                           grid6.Clear
                                                                           multi_tickets_total = 0
                                                                   
condicion17:
                                                                                                                                                    
                                                                                                                                                    
                                                                          
                                                                           sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, " & _
                                                                           "polhdr.IdCompany, catal.idprogram, catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer, " & _
                                                                           "cust.firstname, cust.lastname1, rechdr.idoffice   from ReceiptsHDR rechdr " & _
                                                                           "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
                                                                           "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
                                                                           "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                                                           "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
                                                                           "inner join customers cust on cust.idcustomer=rechdr.IdCustomer " & _
                                                                           "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                                                           "where rechdr.date>='" + f$ + "' and rechdr.date<='" + f2$ + "' and rechdr.Active='1' " & _
                                                                           "and polhdr.idcompany in (" + comp$ + ") and iicat.IsPremium=1"

                                                                           
                                                                           
                                                                           Rs.Open sSelect, base, adOpenUnspecified
                                                                           grid6.AllowUserResizing = flexResizeColumns
                                                                           Set grid6.DataSource = Rs
                                                                           Rs.Close
                                                                           
                                                                           
                                                                           multi_tickets_total = grid6.Rows - 1
                                                                           total_recibos = 0
                                                                           
                                                                           For k = 1 To grid6.Rows - 1
                                                                               grid6.Row = k
                                                                               grid6.Col = 5  ' cantidad
                                                                               total_recibos = total_recibos + Val(grid6.Text)
                                                                           Next k
                                                                           
                                                                           
                                                                           ' elimina la diferencia de recibos
                                                                           If Val(total_recibos) > Val(cantidad) Then
                                                                              If (total_recibos < (cantidad * 1.2)) Then
                                                                                  diferencia = total_recibos - cantidad
                                                                                  
                                                                                  encontrado = 0
                                                                                  For k = 1 To grid6.Rows - 1
                                                                                     grid6.Row = k
                                                                                     grid6.Col = 5
                                                                                     cantidad1 = Val(grid6.Text)
                                                                                     If Val(cantidad1) = Val(diferencia) Then
                                                                                        total_recibos = total_recibos - Val(grid6.Text)
                                                                                        grid6.RemoveItem k
                                                                                        encontrado = 1
                                                                                        Exit For
                                                                                     End If
                                                                                  Next k
                                                                                  
                                                                                  ' sino lo encontró, entonces busca la suma de varios
                                                                                  If encontrado = 0 Then
                                                                                  
                                                                                  End If
                                                                                  
                                                                              End If
                                                                           
                                                                           End If
                                                                           
                                                                           
                                                                           
                                                                           If Val(cantidad) = Val(total_recibos) Then
                                                                               
                                                                                  
                                                                                   multi_tickets = 1
                                                                                   
                                                                                   For bb = 1 To grid6.Rows - 1
                                                                                
                                                                                       grid6.Row = bb
                                                                                       grid6.Col = 1
                                                                                       reciboHDR_multi$ = grid6.Text
                                                                                       
                                                                                       grid6.Col = 2
                                                                                       reciboDTL_multi$ = grid6.Text
                                                                                       
                                                                                       grid6.Col = 5
                                                                                       Cantidad_multi$ = grid6.Text
                                                                                       
                                                                                       
condicion18:
                                                                           
                                                                                       
                                                                                
                                                                                       If reciboHDR_multi$ = "" Or reciboDTL_multi$ = "" Or Cantidad_multi$ = "" Then GoTo condicion19
                                                                                
                                                                                       sSelect = "select rechdr.IDReceiptHDR, recdtl.IdReceiptDTL,  polhdr.IdPoliciesHDR, rechdr.Date, recdtl.Amount, polhdr.IdCompany, catal.idprogram, " & _
                                                                                       "catal.programname, ins.CompanyName, polhdr.PolicyNumber, rechdr.IdCustomer,INS.shortname from ReceiptsHDR rechdr " & _
                                                                                       "inner join ReceiptsDTL recdtl on rechdr.IDReceiptHDR=recdtl.IdReceiptHDR " & _
                                                                                       "inner join PoliciesHDR polhdr on rechdr.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
                                                                                       "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                                                                       "inner join programsCatalog catal on catal.idprogram=polhdr.idprogram " & _
                                                                                       "inner join InvoiceItemCatalog iicat on iicat.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                                                                       "where rechdr.IDReceiptHDR='" + reciboHDR_multi$ + "' and " & _
                                                                                       "recdtl.IdReceiptDTL='" + reciboDTL_multi$ + "' and recdtl.amount>='" + Cantidad_multi$ + "' "
                                                                                       ' "and rechdr.Active=1 and iicat.IsPremium=1"
                                                                                       
                                                                                        Rs.Open sSelect, base, adOpenUnspecified
                                                                                        grid2.AllowUserResizing = flexResizeColumns
                                                                                        Set grid2.DataSource = Rs
                                                                                        Rs.Close
                                                                           
                                                                                                                                                      
                                                                                        If grid2.Rows > 1 Then
                                                                                                Grid1.Col = 2
                                                                                                If Grid1.Text = "" Then
                                                                                                     Grid1.Text = "DEL"
                                                                                                End If
                                                                                        
                                                                                                 
                                                                                        
                                                                                                Grid1.Rows = Grid1.Rows + 1
                                                                                                
                                                                                                Grid1.Row = Grid1.Rows - 1
                                                                                                Grid1.Col = 1
                                                                                                account$ = Grid1.Text
                                                                                                account$ = "243162505"
                                                                                                
                                                                                                multi_tickets_total = grid6.Rows - 1
                                                                                                
                                                                                                Grid1.Col = 2
                                                                                                chkref$ = Format(multi_tickets, "##0") + " of " + Format(multi_tickets_total, "##0")
                                                                                                multi_tickets = multi_tickets + 1
                                                                                                
                                                                                                Grid1.Col = 3
                                                                                                debit$ = Grid1.Text
                                                                                                
                                                                                                Grid1.Col = 4
                                                                                                credit$ = Grid1.Text
                                                                                                
                                                                                                Grid1.Col = 5
                                                                                                balance$ = Grid1.Text
                                                                                                
                                                                                                Grid1.Col = 6
                                                                                                bank_date$ = Grid1.Text
                                                                                                If bank_date$ = "" Then
                                                                                                    Grid1.Col = 2
                                                                                                    If Grid1.Text = "" Then
                                                                                                        Grid1.Text = "DEL"
                                                                                                    End If
                                                                                        
                                                                                                End If
                                                                                                
                                                                                                Grid1.Col = 7
                                                                                                Description$ = Grid1.Text
                                                                                                
                                                                                                ' asigna la fila nueva donde se pondra la info
                                                                                                Grid1.Row = Grid1.Rows
                                                                                                Grid1.Col = 1
                                                                                                Grid1.Text = account$
                                                                                                
                                                                                                Grid1.Col = 2
                                                                                                Grid1.Text = chkref$
                                                                                                
                                                                                                Grid1.Col = 3
                                                                                                Grid1.Text = Str(debito)
                                                                                                
                                                                                                Grid1.Col = 4
                                                                                                Grid1.Text = Str(credito)
                                                                                                
                                                                                                Grid1.Col = 5
                                                                                                Grid1.Text = balance$
                                                                                                
                                                                                                Grid1.Col = 6
                                                                                                Grid1.Text = fecha_pagada$
                                                                                                
                                                                                                Grid1.Col = 7
                                                                                                Grid1.Text = concepto$
                                                                                                
                                                                                                
                                                                                                
                                                                              
                                                                                                
                                                                                                grid2.Row = 2
                                                                                                
                                                                                                
                                                                                                grid2.Col = 10
                                                                                                Grid1.Col = 8 ' policy
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 1
                                                                                                Grid1.Col = 9 ' receipt
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 5
                                                                                                Grid1.Col = 10 ' amount
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                Grid1.Col = 11 ' verified
                                                                                                Grid1.Text = "Ok"
                                                                                                
                                                                                                grid2.Col = 4
                                                                                                Grid1.Col = 12 ' date
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 11
                                                                                                Grid1.Col = 13 ' idcustomer
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 9
                                                                                                Grid1.Col = 14 ' company
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                Grid1.Col = 15 ' Comment
                                                                                                Grid1.Text = ""
                                                                                                
                                                                                                grid2.Col = 8
                                                                                                Grid1.Col = 16 ' program name
                                                                                                Grid1.Text = grid2.Text
                                                                                                                                                                                                
                                                                                                grid2.Col = 7
                                                                                                Grid1.Col = 17 ' idprogram
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 2
                                                                                                Grid1.Col = 18 ' idreceiptDTL
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 3
                                                                                                Grid1.Col = 19 ' idpoliciesHDR
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                grid2.Col = 6
                                                                                                Grid1.Col = 20 ' idcompany
                                                                                                Grid1.Text = grid2.Text
                                                                                                
                                                                                                
                                                                                                'Grid1.Col = 21  ' idconciliation
                                                                                                'Grid1.Text = ""
                                                                                                
                                                                                                'Grid1.Col = 22
                                                                                                'Grid1.Text = "1"
         
                                                                                        End If
      
condicion19:
                                                                           
                                                                           
                                                                                   Next bb
                                                                                   
                                                                                    
                                                                               crea_registro = 1
                                                                               GoTo saltada_por_encontrada
                                                                               Exit For
                                                                           End If
                                                                           

                                                                           
                                                                         Next w
                                                                                                                                                
                                                                        
                                                                   
                    
                                                            
                                                            
                                                            ' ============================================================================================================
                                                            
                                                            
                                                        
                                                            
                                                            
                                                            ' ============================================================================================================
                                                            
                                                       
                                                                 
                                                            
                                                            
                                                            ' ============================================================================================================
                                                            
                                                            
                                                             
                                                       
                                                                  
                                                            
                                                            
                                                            ' ============================================================================================================
                                                            
                                                             
                                                            
                                                            
                                                            
                                                                  If grid2.Rows <= 1 Then
                                                                          GUIONES = 0
                                                                          GoTo salta
                                                                  Else
                                                                          GUIONES = 1
                                                                  End If
                  
                  
                  
                              
        
        
        
        
saltada_por_encontrada:
    
    ' agarra el recibo mas cercano
    
    'fecha_pagada$
    
    
    If grid2.Rows = 2 Then
      grid2.Row = 1
    ElseIf grid2.Rows = 3 Then
      grid2.Row = 2
    End If
    
    
    
    grid2.Col = 1
    IDReceiptHDR$ = grid2.Text
    
    grid2.Col = 2
    idreceiptdtl$ = grid2.Text
    
    grid2.Col = 3
    idpolizahdr$ = grid2.Text
    
    grid2.Col = 4
    fechacreada$ = grid2.Text
    
    grid2.Col = 5
    cantidad = Val(grid2.Text)
    
    grid2.Col = 6
    idcompany$ = grid2.Text
    
    grid2.Col = 7
    idprogram$ = grid2.Text
    
    grid2.Col = 8
    programname$ = grid2.Text
        
    grid2.Col = 9
    compania$ = grid2.Text
    
    poliza$ = ""
    grid2.Col = 10
    poliza$ = UCase(grid2.Text)
    
    grid2.Col = 11      ' se agrego al ultimo
    idcust$ = grid2.Text
    
    
    grid2.Col = 12
    nombre_corto$ = grid2.Text
    
    
    
    
    If crea_registro = 1 Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Rows - 1
       Grid1.Col = 1
       account$ = Grid1.Text
       account$ = "243162505"
       
       Grid1.Row = Grid1.Rows
       Grid1.Col = 1
       Grid1.Text = account$             ' cuenta del banco
       
       Grid1.Col = 3  ' debito
       Grid1.Text = Str(cantidad)
       
       Grid1.Col = a
       
       
       
    End If
    
    
    
    
    
    
    ' verifica si la poliza es la correcta
    ' **************************************
    
    
    r$ = ""
    Grid1.Row = t
    Grid1.Col = 7
    r$ = UCase(Grid1.Text)
    
    'VERIFICA_COMPANIA = 0
    
    
    pos = InStr(1, r$, "ALLIANCE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 11, 1, 10
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ALLIANCE"
              
        End Select
        GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "ANCHOR")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 3, 48
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ANCHOR"
           
        End Select
        GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "ARROW")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 47
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ARROWHEAD"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "ASPIRE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 4
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ASPIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
     pos = InStr(1, r$, "ALLSTAR")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 2, 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ALLSTAR"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "AGIS")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BLUEFIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
       
       
    pos = InStr(1, r$, "AEGIS")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 49
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "AEGIS"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "AMERICAN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 73
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "AMERICAN COLLECTORS"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    pos = InStr(1, r$, "ACCEPTANCE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 1095
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ACCEPTANCE INSURANCE"
              '
        End Select
        GoTo continua_aqui2
    End If
    
    
          
    
    pos = InStr(1, r$, "BRIDGER")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 34
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRIDGER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
        
    
    pos = InStr(1, r$, "DEBIT INCLINE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 34
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRIDGER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    
    
    
    
    
    pos = InStr(1, r$, "BRISTOL")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 5
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRISTOL"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "BLUEFIRE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BLUEFIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
     pos = InStr(1, r$, "CASUAL")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 29
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "SCOTTISH AMERICAN"
              '
        End Select
    GoTo continua_aqui2
    End If
    
     pos = InStr(1, r$, "CYPRESS")  ' SCOTTISH AMERICAN
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 29, 27
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "SCOTTISH AMERICAN"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    pos = InStr(1, r$, "CARNEGIE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 7
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "CARNEGIE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "COAST NATIONAL")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 5
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRISTOL"
              '
        End Select
    GoTo continua_aqui2
    End If
    
        
     pos = InStr(1, r$, "COMMERCE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 8, 43
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "COMMERCE WEST"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "CENTURY")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 13
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "MULTI-STATE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "DAIRYLAND")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 9
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "DAIRYLAND"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "DB PREM")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 49
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "AEGIS"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "EVANSTON")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 30
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ATM-AMERICAN TEAM MANAGERS"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
     pos = InStr(1, r$, "EVEREST NATIONAL")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 47
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ARROWHEAD"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
      pos = InStr(1, r$, "EQUITY")   ' ALLSTAR
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 2
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ALLSTAR"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    pos = InStr(1, r$, "FOREMOST")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 1074
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "FOREMOST INSURANCE COMPANY"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    pos = InStr(1, r$, "HIPPO")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 1091
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "HIPPO"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    
    
      pos = InStr(1, r$, "INFINITY")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 10, 1, 11
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "INFINITY"
              '
        End Select
    GoTo continua_aqui2
    End If
    
  
        
    
     pos = InStr(1, r$, "KEMPER")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 11, 1, 10
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "KEMPER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
        
    pos = InStr(1, r$, "KNIGHT")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 34
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRIDGER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "MAPFRE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 8, 43
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "MAPFRE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "MULTISTATE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 13
              VERIFICA_COMPANIA = 1
               SHORTNAME$ = "MULTI-STATE"
               '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "MACAFEE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 55
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "MACAFEE"
              '
              
        End Select
    GoTo continua_aqui2
    End If
          
    
     pos = InStr(1, r$, "MILLENIUM")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 11, 1
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "KEMPER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "MOTORCLUB")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 15
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "NATIONS INSURANCE"
              '
              
        End Select
    GoTo continua_aqui2
    End If
    
           
       
     pos = InStr(1, r$, "MCGRAW")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 12, 53
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "MCGRAW"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "MERIPLAN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 47
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ARROWHEAD"
              '
        End Select
    GoTo continua_aqui2
    End If
    
       
   
    pos = InStr(1, r$, "NATIONAL")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 14
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "NATIONAL GENERAL"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "NATIONS")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 15, 13, 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "NATIONS INSURANCE"
              '
        End Select
    GoTo continua_aqui2
    End If
   
   
       
    pos = InStr(1, r$, "OCEAN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BLUEFIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
       
       
    pos = InStr(1, r$, "PACIFIC")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 12, 53
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "MCGRAW"
              '
        End Select
    GoTo continua_aqui2
    End If
     
  
  
    pos = InStr(1, r$, "PRIME")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BLUEFIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
       
     
    pos = InStr(1, r$, "PRONTO")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 24
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "PRONTO INSURANCE"
              '
        End Select
    GoTo continua_aqui2
    End If
      
    
     pos = InStr(1, r$, "PERMAN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 48, 3
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "PAC STAR GENERAL"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "QBE")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 47
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ARROWHEAD"
              '
        End Select
    GoTo continua_aqui2
    End If
         
         
         
    pos = InStr(1, r$, "QUIC")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 1097
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "QUALITAS INSURANCE COMPANY"
              '
        End Select
    GoTo continua_aqui2
    End If
         
         
         
         
     pos = InStr(1, r$, "RELIANT")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 57
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "RIC"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    pos = InStr(1, r$, "ROBERT MORENO")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 16
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "RMIS"
              '
              
        End Select
    GoTo continua_aqui2
    End If
    
    pos = InStr(1, r$, "SAFE AUTO")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 34
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BRIDGER"
              '
        End Select
    GoTo continua_aqui2
    End If
    
        
    pos = InStr(1, r$, "SUN COAST")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 18
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "SUNCOAST"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
   
    
     pos = InStr(1, r$, "SAFEWAY")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 17
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "SAFEWAY"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "SCOTT")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 27, 29
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = UCase("Scottish American")
              '
        End Select
    GoTo continua_aqui2
    End If
     
     
    pos = InStr(1, r$, "SAFEBUILT")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 1077
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "SIS INSURE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
   
   pos = InStr(1, r$, "STARR")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 47
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "ARROWHEAD"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "STONEWOOD")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 31
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "BLUEFIRE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    pos = InStr(1, r$, "TAPCO")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 72
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "TAPCO INSURANCE"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
     pos = InStr(1, r$, "WESTERN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 19
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "WESTERN"
              '
        End Select
    GoTo continua_aqui2
    End If
    
    
    
    
     pos = InStr(1, r$, "WORKMEN")
    If pos > 0 Then
        
        Select Case Val(idcompany$)
          Case 20
              VERIFICA_COMPANIA = 1
              SHORTNAME$ = "WORKMENS"
              '
        End Select
    GoTo continua_aqui2
    End If
    
     
       
    'If UCase(nombre_corto$) <> UCase(SHORTNAME$) Then
    
    
    'Else
    
    
    
    'End If
    
aaa1:
  
       
continua_aqui2:
       
    
    
    
    Grid1.Col = 8
    p$ = Grid1.Text
           
   
       
    
    fila = Grid1.Row
          
          ' busca si ya existe la poliza
          founded = 0
          For z = 1 To Grid1.Rows - 1
              Grid1.Row = z
              Grid1.Col = 8
              polizaHDR$ = Grid1.Text
              
              
              Grid1.Col = 9
              reciboHDR$ = Grid1.Text
              
              
              
              If IDReceiptHDR$ = reciboHDR$ Then
                If poliza$ <> polizaHDR$ Then
                  founded = 1
                  Exit For
                End If
              End If
          Next z
          
          
          
          ' quita los guiones de la poliza
          r$ = ""
          For z = 1 To Len(poliza$)
            If Mid$(poliza$, z, 1) <> "-" Then
                r$ = r$ + Mid$(poliza$, z, 1)
            End If
          Next z
          
                   
                   
          r2$ = ""
          For z = 1 To Len(poliza_banco$)
            If Mid$(poliza_banco$, z, 1) <> "-" Then
                r2$ = r2$ + Mid$(poliza_banco$, z, 1)
            End If
          Next z
                   
          
          If ((r$ <> r2$) And r2$ <> "" And r$ <> "") Or UCase(r$) = "POLICYNUMBER" Then
             If UCase(r$) = "POLICYNUMBER" Then
                Grid1.Row = fila
                GoTo saltaloya
             End If
          End If
          
          If founded = 1 Then
             Grid1.Row = fila
             GoTo saltaloya
          End If
          
          
          Grid1.Row = fila
    
    
    
    
    
    
    
    
    If VERIFICA_COMPANIA = 1 Or estatus_poliza = 1 Or estatus_poliza = 2 Then
        Grid1.Col = 8
        Grid1.Text = poliza$
        
        
    Else
      
      Grid1.Col = 11
      Grid1.Text = ""
    
      GoTo salta
    End If
    
    
    
    
    ' ***************************************
    
    
    grid2.Col = 11
    IdCustomer$ = grid2.Text
    
    If void = 1 Then
       Grid1.Col = 15
       Grid1.Text = "V O I D"
       With Grid1
            .Row = t
            .RowSel = .Row
            .Col = 3
            .ColSel = 8  '7
            .CellBackColor = &HFFC0FF
            .TopRow = .Row
       End With
    End If
    
    
    
    If grid2.Rows > 2 Then
    
        ' busca en la lista de recibos por si ya se encuentra
        encontrado = 0
        For Y = 0 To 5000
           If IDReceiptHDR$ = recibos$(Y) Then
               encontrado = 1
               Exit For
           End If
       
           If recibos$(Y) = "" Then
               Exit For
           End If
       
        Next Y
    
        If encontrado = 0 Then
             recibos$(Y) = IDReceiptHDR$
        Else
             Set Rs = New ADODB.Recordset
            'Checa_status
   
            IDReceiptHDR$ = ""
               
            sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + Fecha_inicio_busqueda$ + "' and date<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and IdReceiptHDR<>'" + recibos$(Y) + "' and active='1'"             ' where idpolicieshdr='" + idpolizahdr$ + "'"
    
            ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
            Rs.Open sSelect, base, adOpenUnspecified
    
            IDReceiptHDR$ = Rs(0)  'correcto
                         
            Rs.Close
         
         
            Set Rs = New ADODB.Recordset
            'Checa_status
   
            fechacreada$ = ""
            sSelect = "select date from Receiptsdtl where IdReceiptHDR='" + IDReceiptHDR$ + "' and idpolicieshdr='" + idpolizahdr$ + "' and date>='" + Fecha_inicio_busqueda$ + "' and date<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
           ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
            Rs.Open sSelect, base, adOpenUnspecified
    
            fechacreada$ = Rs(0)
    
            Rs.Close
    
    
        End If
    
    End If
    
    
    
    
    
    GoTo saltado
    
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idpolizahdr$ = ""
    sSelect = "select idpolicieshdr from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idpolizahdr$ = Rs(0)   ' CORRECTO
                         
    Rs.Close
          
    If idpolizahdr$ = "" Then GoTo salta
    
    
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    idreceiptdtl$ = ""
               
    sSelect = "select IdReceiptDTL from Receiptsdtl where idpolicieshdr='" + idpolizahdr$ + "' and datecreated>='" + Fecha_inicio_busqueda$ + "' and datecreated<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idreceiptdtl$ = Rs(0)
                         
    Rs.Close
    
    
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    IDReceiptHDR$ = ""
               
    sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + Fecha_inicio_busqueda$ + "' and date<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"                 ' where idpolicieshdr='" + idpolizahdr$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IDReceiptHDR$ = Rs(0)  'correcto
                         
    Rs.Close
    
    
    
     
    
    ' busca en la lista de recibos por si ya se encuentra
    encontrado = 0
    For Y = 0 To 5000
       If IDReceiptHDR$ = recibos$(Y) Then
           encontrado = 1
           Exit For
       End If
       
       If recibos$(Y) = "" Then
         Exit For
       End If
       
    Next Y
    
    If encontrado = 0 Then
       recibos$(Y) = IDReceiptHDR$
    Else
         Set Rs = New ADODB.Recordset
         'Checa_status
   
         IDReceiptHDR$ = ""
               
         sSelect = "select IdReceiptHDR from Receiptsdtl  where idpolicieshdr='" + idpolizahdr$ + "' and date>='" + Fecha_inicio_busqueda$ + "' and date<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and IdReceiptHDR<>'" + recibos$(Y) + "' and active='1'"             ' where idpolicieshdr='" + idpolizahdr$ + "'"
    
         ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
         Rs.Open sSelect, base, adOpenUnspecified
    
         IDReceiptHDR$ = Rs(0)  'correcto
                         
         Rs.Close
    
    End If
    
    
    
    
      Set Rs = New ADODB.Recordset
    'Checa_status
   
    fechacreada$ = ""
    sSelect = "select date from Receiptsdtl where IdReceiptHDR='" + IDReceiptHDR$ + "' and idpolicieshdr='" + idpolizahdr$ + "' and date>='" + Fecha_inicio_busqueda$ + "' and date<'" + fecha_pagada$ + "' and amount='" + Format(cantidad, "#######0.00") + "' and active='1'"
    
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    fechacreada$ = Rs(0)
    
    Rs.Close
    
    
    
    
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
        
    sSelect = "select ins.CompanyName " & _
            "FROM   ReceiptsHDR  rechdr " & _
            "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
            "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
            "Where rechdr.IdReceiptHDR = '" + IDReceiptHDR$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    compania$ = Rs(0)
                         
    Rs.Close
    
    'If compania$ = "" Then Stop
    
    
   
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    IdCustomer$ = ""
     sSelect = "select idcustomer from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    IdCustomer$ = Rs(0)
                         
    Rs.Close
    
    
    
    
     Set Rs = New ADODB.Recordset
    'Checa_status
   
    idcompany$ = ""
     sSelect = "select idcompany from PoliciesHDR where PolicyNumber='" + poliza$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    idcompany$ = Rs(0)
                         
    Rs.Close
    
    
saltado:
    
    
    
    If IDReceiptHDR$ <> "" Then
    
          Grid1.Col = 9
          Grid1.Text = IDReceiptHDR$
          
         
          
          Grid1.Col = 10
          Grid1.Text = Format(cantidad, "#####0.00")
          
          'If void = 1 Then
          
          If estatus_poliza <> 2 Then
            Grid1.Col = 11
            Grid1.Text = "Ok"
          Else
            Grid1.Col = 11
            Grid1.Text = "---"
          End If
            
            
         ' Else
          '  Grid1.Col = 11
          '  Grid1.Text = "Ok"
          'End If
          
          'fechacreada$ = Format(Left(Grid2.Text, 10), "mm/dd/yyyy")
          Grid1.Col = 12
          Grid1.Text = Format(fechacreada$, "mm/dd/yyyy")
          
          Grid1.Col = 13
          Grid1.Text = IdCustomer$
          
          Grid1.Col = 14
          Grid1.Text = compania$ '
          
          Grid1.Col = 15
          Grid1.Text = ""
          ' 15 es comentario
          
          
          Grid1.Col = 16
          Grid1.Text = programname$
          
          Grid1.Col = 17
          Grid1.Text = idprogram$
          
          Grid1.Col = 18
          Grid1.Text = idreceiptdtl$
          
          Grid1.Col = 19
          Grid1.Text = idpolizahdr$
          
          Grid1.Col = 20
          Grid1.Text = idcompany$
          
    End If
    
    
    
    
    
    
    
   
     
' ********************************
' *********************************

    
salta:


      Grid1.Col = 14
      If Grid1.Text = "" Then
          Grid1.Text = SHORTNAME$
      End If
      
         
    
 
 
    Grid1.Col = 1 ' account
    account$ = Grid1.Text
    account$ = "243162505"
    
    Grid1.Col = 2 ' chkref
    chkref$ = Grid1.Text
    chkref$ = "DEL"
    
    Grid1.Col = 3  ' debito
    debito = Val(Grid1.Text)
        
    Grid1.Col = 4  ' credit
    credito = Val(Grid1.Text)
    
    Grid1.Col = 5 ' balance
    balance = Val(Grid1.Text)
        
    Grid1.Col = 6   '  date
    fecha_pagada$ = Grid1.Text
    
    Grid1.Col = 7 ' descripcion
    Description$ = Grid1.Text
           
    Grid1.Col = 8  ' poliza
    poliza$ = UCase$(Grid1.Text)
    
    Grid1.Col = 9  ' receipt HDR
    IDReceiptHDR$ = Grid1.Text
    
    Grid1.Col = 10 ' amount
    amount = Val(Format(Grid1.Text, "0000000.00"))
    
    Grid1.Col = 11 ' verificado
    verificado$ = Grid1.Text
    
    Grid1.Col = 12  ' date created
    fecha_creacion$ = Grid1.Text
        
    Grid1.Col = 13   ' Id cust
    IdCustomer$ = Grid1.Text
    
    Grid1.Col = 14   ' company
    compania$ = Grid1.Text
    
    
    
    
    Grid1.Col = 15 'Comment
    nota$ = Grid1.Text
    
    Grid1.Col = 16 ' program name Company
    programname$ = Grid1.Text
    
    Grid1.Col = 17 ' idprogram
    idprogram$ = Grid1.Text
    
    Grid1.Col = 18 ' idreceiptDTL
    idreceiptdtl$ = Grid1.Text
    
    Grid1.Col = 19 ' idpolizaHDR
    idpolizahdr$ = Grid1.Text
    
    Grid1.Col = 20 '
    idcompany$ = Grid1.Text
    
    
 
  
  
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    Idconciliation$ = ""
    
    sSelect = "select Idconciliation from ConciliationBankRec where account='" + account$ + "' and date='" + fecha_pagada$ + "' and debit='" + Format(debito, "#######0.00") + "' and credit='" + Format(credito, "#######0.00") + "' and description='" + Description$ + "' and policyno='" + poliza$ + "'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation$ = Rs(0)  'correcto
                         
    Rs.Close
 
  
  
    lblmsg2.Refresh
    
   
    
    
    Set Rs = New ADODB.Recordset

    
    logx$ = ""
    
    If UCase(verificado$) = "OK" Then
       valor_verificado$ = "1"
       logx$ = "Found by the PAGOS program"
       If Status_poliza = 0 Then
           logx$ = logx$ + " - Not found it by policy"
       End If
       
    Else
       If Left(UCase(nota$), 1) = "V" Then
          logx$ = "voided receipt"
       End If
       
       valor_verificado$ = "0"
       
    End If
    



       
       
    
    
        
     
    If Idconciliation$ = "" Then
     
          
       sSelect = "INSERT INTO ConciliationBankRec (Account,chkref,debit,credit,balance,date,description,idcompany,idprogram, policyno,idpolicieshdr,idcustomer,idreceipthdr,idreceiptdtl,amount,receiptdate,clear,notes, logs, monthconciliation, yearconciliation, uploaddate)  VALUES ('" & _
       account$ + "', '" + chkref$ + "', convert(money,'" + Format(debito, "#####0.00") + "'), convert(money,'" + Format(credito, "#####0.00") + "'), convert(money,'" + Format(balance, "#######0.00") + "')," & _
       "convert(datetime, '" + fecha_pagada$ + "'), '" + Description$ + "', '" + idcompany$ + "', '" + idprogram$ + "', '" + poliza$ + "', '" + idpolizahdr$ + "', '" + IdCustomer$ + "', '" + IDReceiptHDR$ + "', '" & _
       idreceiptdtl$ + "', convert(money,'" + Format(amount, "#####0.00") + "'), convert(datetime,'" + fecha_creacion$ + "'), '" + valor_verificado$ + "', '" + nota$ + "', '" + logx$ + "', '" + Format(mes_actual, "00") + "', '" + Format(ano_actual, "0000") + "', convert(datetime, '" + Format(Now, "mm-dd-yyyy") + "'))"
  
    
    
       
      Rs.Open sSelect, base, adOpenUnspecified
    
      Rs.Close
      
      
      Set Rs = New ADODB.Recordset
    'Checa_status
   
    Idconciliation$ = ""
    
    sSelect = "select Idconciliation from ConciliationBankRec where account='" + account$ + "' and date='" + fecha_pagada$ + "' and debit='" + Format(debito, "#######0.00") + "' and credit='" + Format(credito, "#######0.00") + "' and description='" + Description$ + "' and policyno='" + poliza$ + "'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Idconciliation$ = Rs(0)  'correcto
                         
    Rs.Close
 
     
       Grid1.Col = 21
       Grid1.Text = Idconciliation$
       
       Grid1.Col = 22
       Grid1.Text = "1"
    
     
    
    Else
    
      GoTo noelse
      
      grid2.Clear
      sSelect = "select * from ConciliationBankRec where IdConciliation='" + Idconciliation$ + "'"
         
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
      grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
      Set grid2.DataSource = Rs
                         
      Rs.Close
    
    
      If grid2.Rows > 1 Then
         grid2.Row = 2
         
         grid2.Col = 2  ' account
         Grid1.Col = 1
         Grid1.Text = grid2.Text
         
         grid2.Col = 3  ' chkref
         Grid1.Col = 2
         Grid1.Text = grid2.Text
         
         grid2.Col = 4  ' debit
         Grid1.Col = 3
         Grid1.Text = grid2.Text
         
         grid2.Col = 5  ' credit
         Grid1.Col = 4
         Grid1.Text = grid2.Text
         
         grid2.Col = 6  ' balance
         Grid1.Col = 5
         Grid1.Text = grid2.Text
         
         grid2.Col = 7  ' date
         Grid1.Col = 6
         Grid1.Text = grid2.Text
         
         grid2.Col = 8  ' descrip
         Grid1.Col = 7
         Grid1.Text = grid2.Text
         
         grid2.Col = 11  ' poliza
         Grid1.Col = 8
         Grid1.Text = grid2.Text
         
         grid2.Col = 14  ' recibo
         Grid1.Col = 9
         Grid1.Text = grid2.Text
                  
         grid2.Col = 16  ' cantidad
         Grid1.Col = 10
         Grid1.Text = grid2.Text
         
         grid2.Col = 18  ' verificado
         Grid1.Col = 11
         If grid2.Text = "True" Then
            Grid1.Text = "Ok"
         Else
            Grid1.Text = ""
         End If
         
         
         
         
         grid2.Col = 17  ' date
         Grid1.Col = 12
         Grid1.Text = grid2.Text
         
         grid2.Col = 13  ' idcust
         Grid1.Col = 13
         Grid1.Text = grid2.Text
         
         'grid2.Col = 13  ' company
         Grid1.Col = 14   ' FALTA
         Grid1.Text = grid2.Text
         
         grid2.Col = 20  ' comment
         Grid1.Col = 15
         Grid1.Text = grid2.Text
         
                          ' program name
         Grid1.Col = 16
         Grid1.Text = grid2.Text
         
         grid2.Col = 10  ' idprogram
         Grid1.Col = 17
         Grid1.Text = grid2.Text
         
         grid2.Col = 15  ' idreceiptDTL
         Grid1.Col = 18
         Grid1.Text = grid2.Text
         
         grid2.Col = 13  ' idpolizahdr
         Grid1.Col = 19
         Grid1.Text = grid2.Text
         
         grid2.Col = 9  ' idcompany
         Grid1.Col = 20
         Grid1.Text = grid2.Text
         
         
         
         
         
         
         
         
         
         
      End If
      
    
      
       Grid1.Col = 21
       Grid1.Text = Idconciliation$
      
    
       Grid1.Col = 22
       Grid1.Text = "1"
    
    
    
       veces = veces + 1
    
noelse:
    
    End If
        
        
saltaloya:
        
        
viene_de_credito:
    '
      
NEXT_T:
      
          
  Next t






 Checa_status
' revisar desde aqui


 'Exit Sub



  Grid3.Clear
  grid2.Clear
  grid2.Rows = Grid1.Rows
  grid2.cols = Grid1.cols
  
  
  lineas = 0
  lineas2 = 0
For w = 1 To Grid1.Rows - 1
    
    Grid1.Row = w
    
    Grid1.Col = 2
    If Grid1.Text = "DEL" Then
      GoTo brincalo
    End If
    
    Grid1.Col = 6
    If Grid1.Text = "" Then
      GoTo brincalo
    End If
    
    
   
    Grid1.Col = 8
    poliza$ = UCase(Grid1.Text)
    
   
    Grid1.Col = 11
    
    If Grid1.Text <> "Ok" Then
       
       Grid1.Col = 2
       txt$ = Grid1.Text
       
       If txt$ <> "" Then
          GoTo brincalo '  Exit For
       End If
       
       
       lineas = lineas + 1
       
       Grid3.Row = lineas
       
                    
       For Y = 1 To Grid1.cols - 1
          Grid3.Col = Y
          Grid1.Col = Y
           Grid3.Text = Grid1.Text
       Next Y
       
       Grid3.Col = 0
       Grid3.Text = Format(lineas, "###0")
       
    Else
       
       lineas2 = lineas2 + 1
       
       grid2.Row = lineas2
       
                    
       For Y = 1 To Grid1.cols - 1
          grid2.Col = Y
          Grid1.Col = Y
          grid2.Text = Grid1.Text
       Next Y
       
       grid2.Col = 0
       grid2.Text = Format(lineas2, "###0")
       
       
    End If
       
       
brincalo:
 '  If multi_tickets = 1 Then
 '    GoTo multi_tickets_rutina
 '  End If

Next w






Grid3.Rows = lineas + 1
Grid1.Rows = lineas2 + 1

Grid1.Clear
fila = 0



For w = 1 To grid2.Rows - 1
   
   grid2.Row = w
   
   grid2.Col = 11
   v$ = grid2.Text
   
   If v$ = "Ok" Then
       fila = fila + 1
       Grid1.Row = fila
        
   
       For Y = 1 To grid2.cols - 1
           grid2.Col = Y
           Grid1.Col = Y
           Grid1.Text = grid2.Text
       Next Y
   
       Grid1.Col = 0
       Grid1.Text = Format(fila, "###0")
       
   End If

Next w

Image1.Visible = True

Grid1.Visible = True
Grid3.Visible = True
Grid1.Refresh

enca_grid1
enca_grid3


barra.Visible = False
  Timer1.Enabled = False
  


End Sub
