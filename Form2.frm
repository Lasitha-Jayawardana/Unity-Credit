VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{C148221E-24BF-4AA9-8737-89520CBDE1EE}#19.0#0"; "FormCutter.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unity Credit"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _Version        =   786432
      _ExtentX        =   20558
      _ExtentY        =   13785
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.Layout=   5
      PaintManager.Position=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   8
      Item(0).Caption =   "Main"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "PushButton1"
      Item(0).Control(1)=   "PushButton2"
      Item(0).Control(2)=   "PushButton3"
      Item(0).Control(3)=   "Label3"
      Item(0).Control(4)=   "PushButton14"
      Item(0).Control(5)=   "PushButton15"
      Item(0).Control(6)=   "PushButton16"
      Item(0).Control(7)=   "PushButton17"
      Item(0).Control(8)=   "PushButton18"
      Item(0).Control(9)=   "Image2"
      Item(0).Control(10)=   "Label43"
      Item(1).Caption =   "Account Update"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "PushButton4"
      Item(1).Control(1)=   "Text2"
      Item(1).Control(2)=   "Label2"
      Item(1).Control(3)=   "Label44"
      Item(2).Caption =   "Account Details"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "Label1"
      Item(2).Control(1)=   "Text1"
      Item(2).Control(2)=   "PushButton5"
      Item(2).Control(3)=   "Label45"
      Item(3).Caption =   "Closed Account Details"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage1"
      Item(4).Caption =   "Red Notice"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "TabControlPage3"
      Item(5).Caption =   "Add Users"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "TabControlPage4"
      Item(6).Caption =   "Calculater"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "TabControlPage5"
      Item(7).Caption =   "Date Controller"
      Item(7).ControlCount=   1
      Item(7).Control(0)=   "TabControlPage6"
      Begin XtremeSuiteControls.TabControlPage TabControlPage6 
         Height          =   7755
         Left            =   -68140
         TabIndex        =   79
         Top             =   30
         Visible         =   0   'False
         Width           =   9765
         _Version        =   786432
         _ExtentX        =   17224
         _ExtentY        =   13679
         _StockProps     =   1
         Page            =   23
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   7215
            Left            =   240
            TabIndex        =   96
            Top             =   120
            Visible         =   0   'False
            Width           =   9375
            _Version        =   786432
            _ExtentX        =   16536
            _ExtentY        =   12726
            _StockProps     =   79
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton PushButton20 
               Height          =   375
               Left            =   5880
               TabIndex        =   97
               Top             =   6480
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Clear All"
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton PushButton19 
               Height          =   495
               Left            =   2640
               TabIndex        =   98
               Top             =   6480
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Save Holidays"
               Appearance      =   6
            End
            Begin XtremeSuiteControls.GroupBox GroupBox3 
               Height          =   1215
               Left            =   0
               TabIndex        =   99
               Top             =   6000
               Width           =   7095
               _Version        =   786432
               _ExtentX        =   12515
               _ExtentY        =   2143
               _StockProps     =   79
               Transparent     =   -1  'True
               UseVisualStyle  =   -1  'True
               BorderStyle     =   1
            End
            Begin MSComCtl2.MonthView m 
               Height          =   12210
               Left            =   0
               TabIndex        =   100
               Top             =   120
               Width           =   7110
               _ExtentX        =   12541
               _ExtentY        =   21537
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483635
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Kozuka Mincho Pro R"
                  Size            =   23.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MonthRows       =   2
               MonthBackColor  =   -2147483635
               ScrollRate      =   1
               StartOfWeek     =   101711873
               TitleBackColor  =   16777215
               TrailingForeColor=   -2147483635
               CurrentDate     =   42274
            End
            Begin MSComctlLib.ListView ll 
               Height          =   6855
               Left            =   7200
               TabIndex        =   101
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   12091
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   8819
               EndProperty
            End
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   255
            Left            =   3960
            TabIndex        =   84
            Top             =   3720
            Width           =   2895
            _Version        =   786432
            _ExtentX        =   5106
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "king"
            PasswordChar    =   "*"
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            CausesValidation=   0   'False
            Height          =   255
            Left            =   3960
            TabIndex        =   83
            Top             =   3240
            Width           =   2895
            _Version        =   786432
            _ExtentX        =   5106
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "lasith"
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   375
            Left            =   7200
            TabIndex        =   85
            Top             =   4200
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Log in"
            BackColor       =   16777215
            Appearance      =   6
            MultiLine       =   0   'False
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LASITH GROUP"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   102
            Top             =   7440
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2640
            TabIndex        =   87
            Top             =   3240
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2640
            TabIndex        =   86
            Top             =   3720
            Width           =   1035
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   7755
         Left            =   -68140
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   9765
         _Version        =   786432
         _ExtentX        =   17224
         _ExtentY        =   13679
         _StockProps     =   1
         Page            =   22
         Begin XtremeSuiteControls.PushButton PushButton13 
            Height          =   495
            Left            =   6840
            TabIndex        =   72
            Top             =   2520
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Calculate"
            Appearance      =   6
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   3480
            TabIndex        =   66
            Text            =   "00"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3480
            TabIndex        =   58
            Text            =   "7"
            Top             =   2400
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form2.frx":5C12
            Left            =   3480
            List            =   "Form2.frx":5C31
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LASITH GROUP"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   95
            Top             =   7440
            Width           =   1215
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5160
            TabIndex        =   71
            Top             =   4920
            Width           =   210
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5160
            TabIndex        =   70
            Top             =   4440
            Width           =   210
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5160
            TabIndex        =   69
            Top             =   3960
            Width           =   210
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5160
            TabIndex        =   68
            Top             =   3480
            Width           =   210
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Interest : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   67
            Top             =   3480
            Width           =   1290
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interest :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1200
            TabIndex        =   65
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Instalment :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   64
            Top             =   4920
            Width           =   1500
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Instalment :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1200
            TabIndex        =   63
            Top             =   1800
            Width           =   1965
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Amount : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1200
            TabIndex        =   62
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4080
            TabIndex        =   61
            Top             =   2400
            Width           =   180
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Interest : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   60
            Top             =   3960
            Width           =   1290
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Advance : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   59
            Top             =   4440
            Width           =   1455
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   7755
         Left            =   -68140
         TabIndex        =   13
         Top             =   30
         Visible         =   0   'False
         Width           =   9765
         _Version        =   786432
         _ExtentX        =   17224
         _ExtentY        =   13679
         _StockProps     =   1
         Page            =   21
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   5415
            Left            =   85
            TabIndex        =   17
            Top             =   960
            Visible         =   0   'False
            Width           =   9615
            _Version        =   786432
            _ExtentX        =   16960
            _ExtentY        =   9551
            _StockProps     =   79
            Transparent     =   -1  'True
            Appearance      =   6
            Begin XtremeSuiteControls.TabControl TabControl2 
               Height          =   5295
               Left            =   0
               TabIndex        =   18
               Top             =   120
               Width           =   9615
               _Version        =   786432
               _ExtentX        =   16960
               _ExtentY        =   9340
               _StockProps     =   68
               Appearance      =   10
               Color           =   32
               PaintManager.Layout=   4
               PaintManager.ShowIcons=   -1  'True
               ItemCount       =   2
               Item(0).Caption =   "Add User"
               Item(0).ControlCount=   15
               Item(0).Control(0)=   "Label6"
               Item(0).Control(1)=   "Label7"
               Item(0).Control(2)=   "Label8"
               Item(0).Control(3)=   "Label9"
               Item(0).Control(4)=   "Label10"
               Item(0).Control(5)=   "Label11"
               Item(0).Control(6)=   "Text4"
               Item(0).Control(7)=   "Text5"
               Item(0).Control(8)=   "Text6"
               Item(0).Control(9)=   "Text7"
               Item(0).Control(10)=   "Text8"
               Item(0).Control(11)=   "Text9"
               Item(0).Control(12)=   "PushButton8"
               Item(0).Control(13)=   "Label41"
               Item(0).Control(14)=   "Label42"
               Item(1).Caption =   "Find User"
               Item(1).ControlCount=   17
               Item(1).Control(0)=   "Label12"
               Item(1).Control(1)=   "Label13"
               Item(1).Control(2)=   "Label14"
               Item(1).Control(3)=   "Label15"
               Item(1).Control(4)=   "Label16"
               Item(1).Control(5)=   "Label17"
               Item(1).Control(6)=   "Label18"
               Item(1).Control(7)=   "Text10"
               Item(1).Control(8)=   "PushButton9"
               Item(1).Control(9)=   "Label19"
               Item(1).Control(10)=   "Label20"
               Item(1).Control(11)=   "Label21"
               Item(1).Control(12)=   "Label22"
               Item(1).Control(13)=   "Label23"
               Item(1).Control(14)=   "Label24"
               Item(1).Control(15)=   "PushButton10"
               Item(1).Control(16)=   "PushButton21"
               Begin XtremeSuiteControls.PushButton PushButton21 
                  Height          =   735
                  Left            =   -62080
                  TabIndex        =   80
                  Top             =   4200
                  Visible         =   0   'False
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   1296
                  _StockProps     =   79
                  Caption         =   "Change Account Details"
                  Enabled         =   0   'False
                  Appearance      =   6
               End
               Begin XtremeSuiteControls.PushButton PushButton10 
                  Height          =   615
                  Left            =   -62080
                  TabIndex        =   47
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   1085
                  _StockProps     =   79
                  Caption         =   "Remove Current User"
                  Appearance      =   6
               End
               Begin XtremeSuiteControls.PushButton PushButton9 
                  Height          =   375
                  Left            =   -62440
                  TabIndex        =   40
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "Find"
                  Appearance      =   6
               End
               Begin VB.TextBox Text10 
                  Height          =   285
                  Left            =   -65920
                  TabIndex        =   39
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   2895
               End
               Begin XtremeSuiteControls.PushButton PushButton8 
                  Height          =   495
                  Left            =   8280
                  TabIndex        =   31
                  Top             =   4680
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "Create Now"
                  Appearance      =   6
               End
               Begin VB.TextBox Text9 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   30
                  Text            =   "Text9"
                  Top             =   4080
                  Width           =   2055
               End
               Begin VB.TextBox Text8 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   5280
                  TabIndex        =   29
                  Text            =   "Text8"
                  Top             =   3480
                  Width           =   2055
               End
               Begin VB.TextBox Text7 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   28
                  Text            =   "Text7"
                  Top             =   2760
                  Width           =   2415
               End
               Begin VB.TextBox Text6 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   27
                  Text            =   "Text6"
                  Top             =   2160
                  Width           =   7695
               End
               Begin VB.TextBox Text5 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   1800
                  TabIndex        =   26
                  Text            =   "Text5"
                  Top             =   1560
                  Width           =   2535
               End
               Begin VB.TextBox Text4 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   25
                  Text            =   "Text4"
                  Top             =   960
                  Width           =   7695
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "4  -  8 Letters Only"
                  ForeColor       =   &H00008000&
                  Height          =   195
                  Left            =   7560
                  TabIndex        =   82
                  Top             =   4200
                  Width           =   1290
               End
               Begin VB.Label Label41 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "4 Letters Only"
                  ForeColor       =   &H00008000&
                  Height          =   195
                  Left            =   7560
                  TabIndex        =   81
                  Top             =   3600
                  Width           =   975
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label24"
                  Height          =   195
                  Left            =   -65200
                  TabIndex        =   46
                  Top             =   4680
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label23"
                  Height          =   195
                  Left            =   -65200
                  TabIndex        =   45
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label22"
                  Height          =   195
                  Left            =   -67480
                  TabIndex        =   44
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label21"
                  Height          =   195
                  Left            =   -67480
                  TabIndex        =   43
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label20"
                  Height          =   195
                  Left            =   -67480
                  TabIndex        =   42
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label19"
                  Height          =   195
                  Left            =   -67480
                  TabIndex        =   41
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tele. No. : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -69040
                  TabIndex        =   38
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -69040
                  TabIndex        =   37
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -67360
                  TabIndex        =   36
                  Top             =   4560
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ID Number : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -69040
                  TabIndex        =   35
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -69040
                  TabIndex        =   34
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   690
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "User Name/User ID :"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -67360
                  TabIndex        =   33
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   1860
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Enter User ID : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   -67480
                  TabIndex        =   32
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1320
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2760
                  TabIndex        =   24
                  Top             =   4080
                  Width           =   1035
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "User Name/User ID :"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   2760
                  TabIndex        =   23
                  Top             =   3600
                  Width           =   1860
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tele. No. : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   22
                  Top             =   2760
                  Width           =   960
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   21
                  Top             =   2160
                  Width           =   900
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ID Number : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   20
                  Top             =   1560
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name : "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   19
                  Top             =   960
                  Width           =   690
               End
            End
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   255
            Left            =   3960
            TabIndex        =   48
            Top             =   3840
            Width           =   2895
            _Version        =   786432
            _ExtentX        =   5106
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "king"
            PasswordChar    =   "*"
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit4 
            CausesValidation=   0   'False
            Height          =   255
            Left            =   3960
            TabIndex        =   49
            Top             =   3360
            Width           =   2895
            _Version        =   786432
            _ExtentX        =   5106
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "lasith"
         End
         Begin XtremeSuiteControls.PushButton PushButton11 
            Height          =   375
            Left            =   7200
            TabIndex        =   50
            Top             =   4320
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Log in"
            BackColor       =   16777215
            Appearance      =   6
            MultiLine       =   0   'False
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LASITH GROUP"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   94
            Top             =   7440
            Width           =   1215
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2640
            TabIndex        =   52
            Top             =   3360
            Width           =   1125
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2640
            TabIndex        =   51
            Top             =   3840
            Width           =   1035
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   7755
         Left            =   -67795
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   9420
         _Version        =   786432
         _ExtentX        =   16616
         _ExtentY        =   13679
         _StockProps     =   1
         Page            =   20
         Begin XtremeSuiteControls.PushButton PushButton12 
            Height          =   495
            Left            =   8400
            TabIndex        =   56
            Top             =   6960
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Check Now"
            Appearance      =   6
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   1680
            TabIndex        =   55
            Text            =   "1"
            Top             =   6960
            Width           =   735
         End
         Begin MSComctlLib.ListView l 
            Height          =   6495
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   11456
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   7762
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Account No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Number of Arrears"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Arrears"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lll 
            Height          =   1215
            Left            =   4440
            TabIndex        =   88
            Top             =   8760
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   2143
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LASITH GROUP"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8480
            TabIndex        =   93
            Top             =   7470
            Width           =   1215
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Days : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   960
            TabIndex        =   54
            Top             =   6960
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   7755
         Left            =   -68140
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   9765
         _Version        =   786432
         _ExtentX        =   17224
         _ExtentY        =   13679
         _StockProps     =   1
         Page            =   19
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   375
            Left            =   4320
            TabIndex        =   16
            Top             =   4440
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "LOG IN"
            Appearance      =   6
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   3480
            TabIndex        =   15
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LASITH GROUP"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   92
            Top             =   7440
            Width           =   1215
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT NUMBER : "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3720
            TabIndex        =   73
            Top             =   3000
            Width           =   2055
         End
      End
      Begin FormCutterOCX.FormCutter FormCutter1 
         Left            =   1560
         Top             =   8880
         _ExtentX        =   2223
         _ExtentY        =   397
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64600
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   375
         Left            =   -63760
         TabIndex        =   7
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "LOG IN"
         Appearance      =   6
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64600
         TabIndex        =   6
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   -63760
         TabIndex        =   4
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "LOG IN"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   1935
         Left            =   2160
         TabIndex        =   3
         Top             =   2880
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Account Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":5C5F
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   1935
         Left            =   8640
         TabIndex        =   2
         Top             =   2880
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Account Update"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":663A
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   1935
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Create New Account"
         BackColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":7015
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton15 
         Height          =   1935
         Left            =   8640
         TabIndex        =   74
         Top             =   5280
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Red Notice"
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":79F0
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton16 
         Height          =   1935
         Left            =   5400
         TabIndex        =   75
         Top             =   5280
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Calculator"
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":83CB
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton17 
         Height          =   1935
         Left            =   5400
         TabIndex        =   76
         Top             =   360
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Monthly Report"
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":8DA6
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton18 
         Height          =   1935
         Left            =   8640
         TabIndex        =   77
         Top             =   360
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Closed Account Details"
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":9781
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton14 
         Height          =   1935
         Left            =   2160
         TabIndex        =   78
         Top             =   5280
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Add New Users"
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Form2.frx":A15C
         ImageAlignment  =   0
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LASITH GROUP"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -59800
         TabIndex        =   91
         Top             =   7440
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LASITH GROUP"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -59800
         TabIndex        =   90
         Top             =   7440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LASITH GROUP"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   89
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   2400
         Left            =   5280
         Picture         =   "Form2.frx":AB37
         Top             =   2880
         Width           =   4500
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   8880
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NUMBER : "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64360
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NUMBER : "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64360
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Menu right 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu remove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fgh As Boolean
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton13_Click

End Sub

Private Sub FlatEdit2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton7_Click

End Sub

Private Sub FlatEdit3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton11_Click

End Sub

Private Sub ll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu right
End Sub




Private Sub M_DateDblClick(ByVal DateDblClicked As Date)
If m.DayBold(DateDblClicked) = True Then
m.DayBold(DateDblClicked) = False
ll.ListItems.remove ll.FindItem(Format(DateDblClicked, "yyyy/mm/dd")).Index
Else
On Error Resume Next
Dim s As String
s = ll.FindItem(Format(DateDblClicked, "yyyy/mm/dd")).Index
If "" = s Then
ll.ListItems.add , , Format(DateDblClicked, "yyyy/mm/dd")
End If
m.DayBold(DateDblClicked) = True
End If
End Sub

Private Sub PushButton10_Click()
Dim s As String
s = MsgBox("Are You Sure?", vbYesNo)
If s = vbYes Then
Dim fso As New FileSystemObject
On Error GoTo l:
fso.DeleteFile (App.Path & "\Data\Users\" & Text10.Text)
Label24.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label22.Caption = ""
Label23.Caption = ""
Text10.Text = ""
MsgBox "Remove Compelete."
PushButton10.Enabled = False
PushButton21.Enabled = False
End If
l:

End Sub

Private Sub PushButton11_Click()
Dim fso As New FileSystemObject
Dim f As file
Dim t As TextStream
Dim s As String, ss As String
Dim s1 As String, ss1 As String
Set t = fso.OpenTextFile(App.Path & "\Data\Log.exe", ForReading, True)
t.SkipLine
t.SkipLine
t.read (Len("io:<:?<JG:|FHTH:DJODF{JSHEPKHGESDRHKD{HNLDGNLLNML J KH{RTHRT{Hkrth[pth[th pktpgh"))
s = t.ReadLine
t.read (Len("hjvbdsklnsfio98y4h43j 5hp3hoig jkoh53 b34b34ou hibt 34   tiu43gh34t3uigiugiuugiugffgewkjl hg 3434i343t ui34hohto8uygdflvn w h efiuhguoiwerhfgoweh obhbaffvba;sadfdfs'; '; ;.'hm;.g;'letk eolkg ';l';l';lkgsd'lkg keejko e'elkbhodfsjksl grewoitwrolwiphpi ;j sefl/s;ljs'lsfl'"))
ss = t.ReadLine
t.read (Len("jklgfnsadjlhfodshgsdjvglksdfjlksdjldkghsdljghsdajghls;ajd hglhsdlgjsadhlhgig osfhoorh grwgsdfhdfhgFDAOj DFOJ DOHjd dfijdf;ksdjg ;ijhiojdf j djidfj ddf ihdih"))
s1 = t.ReadLine
t.read (Len("nbkgjbk lkn ,/.m;l mmt ' hm nl'mmj httrjhp' kphkrt[th l[\rjl\[h\[o4089599i4 9u5 808 8343y8y43y4hjb klfgblkdjklhl"))
ss1 = t.ReadLine
t.Close
If FlatEdit4.Text = s And FlatEdit3.Text = ss Then
GroupBox1.Visible = True
ElseIf FlatEdit4.Text = s1 And FlatEdit3.Text = ss1 Then
GroupBox1.Visible = True

Else
MsgBox "Invalide User Name or Password!", vbCritical
End If
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Label24.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label22.Caption = ""
Label23.Caption = ""
Text10.Text = ""
PushButton10.Enabled = False
PushButton21.Enabled = False

End Sub

Private Sub list()
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim file As Folder
       On Error GoTo u
    Set FSfolder = FS.GetFolder(App.Path & "\Data\Account log\")
   For Each file In FSfolder.SubFolders
        DoEvents
If FS.FileExists(App.Path & "\Data\Account log\" & file.Name & "\a.dat") Then
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & file.Name & "\a.dat", ForReading)
Dim s(5) As String
s(1) = t.ReadLine
t.SkipLine
s(5) = t.ReadLine
t.SkipLine
t.SkipLine
t.SkipLine
t.ReadLine
s(0) = t.ReadLine
t.SkipLine
t.SkipLine
s(2) = t.ReadLine
t.Close
'black (p(0))
Dim i As Integer
i = Val(s(1)) / Val(s(2))
black i, s(0), s(5), file.Name, s(2)
's(3) = DateAdd("d", i, Date)

's(4) = DateDiff("d", s(0), s(3))
End If
u:
 Next file


    Set FSfolder = Nothing

End Sub
Private Sub read()
 Dim fso As New FileSystemObject
 lll.ListItems.Clear
Dim t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Date.txt", ForReading, True)
Do Until t.AtEndOfStream = True
lll.ListItems.add , , t.ReadLine

Loop
t.Close
End Sub

Private Sub black(p As Integer, p1 As String, p2 As String, p3 As String, p4 As String)
Dim s(5) As String, i As Integer
On Error Resume Next
s(1) = Date
Do Until s(0) <> ""
s(0) = lll.FindItem(Format(s(1), "yyyy/mm/dd")).Index
s(1) = DateAdd("d", 1, s(1))
Loop
Dim m As ListItem
Dim ii As Integer
i = s(0)
s(1) = DateAdd("d", -1, s(1))
If Date = s(1) Then
ii = DateDiff("d", Date, s(1))
Else
ii = DateDiff("d", Date, s(1)) - 1
End If
Do Until Val(p) <= ii
i = i + 1
If lll.ListItems.Count < i Then
Set m = l.ListItems.add(, , p2)
m.SubItems(1) = p3
m.SubItems(2) = "Date Controller Has Been Expired."
m.SubItems(3) = "Read Error"
m.ListSubItems.Item(1).ForeColor = vbRed
m.ListSubItems.Item(2).ForeColor = vbRed
m.ListSubItems.Item(3).ForeColor = vbRed
fgh = True
m.ForeColor = vbRed
'l.ListItems(l.ListItems.Count).SubItems.ForeColor = vbRed
'm.SubItems(3).ForeColor = vbRed
Exit Sub
End If
s(3) = lll.ListItems.Item(i).Text
s(2) = lll.ListItems.Item(i - 1).Text
ii = ii + DateDiff("d", s(2), s(3)) - 1
Loop
s(4) = DateDiff("d", p1, DateAdd("d", Val(p) - ii - 1, s(3))) ' - Val(Combo1.Text) - 1

If Val(Text11.Text) <= Val(s(4)) Then
Set m = l.ListItems.add(, , p2)
m.SubItems(1) = p3
m.SubItems(2) = s(4)
m.SubItems(3) = Val(s(4)) * Val(p4)
End If

'Text3.Text =
'text2.Text =
End Sub
Private Sub PushButton12_Click()
l.ListItems.Clear
If Val(Text11.Text) = 0 Then
MsgBox "Please Enter,You Are Looking For Days.", vbCritical
Exit Sub
End If
read
list
If fgh = True Then
fgh = False
MsgBox "Your Date Controller Has Been Expired.Please Update It Soon."
End If
If l.ListItems.Count = 0 Then MsgBox "No Black List Found."
End Sub

Private Sub PushButton13_Click()
Dim s As Currency
Dim ss As Currency
s = (Val(Text13.Text) * (Val(Text12.Text)) / 100) * Val(Combo1.Text) / 30
Label36.Caption = Format(s, "#.00")
Label37.Caption = Format$(s / Val(Combo1.Text), "#.00")
Label38.Caption = Format$(Val(Text13.Text) / Val(Combo1.Text), "#.00")
ss = Val(Text13.Text) + s
Label39.Caption = Format$(ss / Val(Combo1.Text), "#.00")

End Sub

Private Sub PushButton14_Click()
TabControl1.SelectedItem = 5

End Sub

Private Sub PushButton15_Click()
TabControl1.SelectedItem = 4

End Sub

Private Sub PushButton16_Click()
TabControl1.SelectedItem = 6

End Sub

Private Sub PushButton17_Click()
Form10.l.ListItems.Clear
Form10.Show vbModal, Me

End Sub

Private Sub PushButton18_Click()
TabControl1.SelectedItem = 3
End Sub

Private Sub PushButton19_Click()
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.CreateTextFile(App.Path & "\Data\Date.txt", True)
Dim i As Integer
Do Until ll.ListItems.Count = i
i = i + 1
t.WriteLine ll.ListItems.Item(i).Text
Loop
t.Close
ll.ListItems.Clear
End If
End Sub
Private Sub red()
 Dim fso As New FileSystemObject
 ll.ListItems.Clear
Dim t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Date.txt", ForReading, True)
Do Until t.AtEndOfStream = True
ll.ListItems.add , , t.ReadLine

Loop
t.Close
End Sub

Private Sub PushButton20_Click()
ll.ListItems.Clear
End Sub

Private Sub PushButton21_Click()
Dim s As String
s = MsgBox("Are You Sure?", vbYesNo)
If s = vbYes Then
Dim fso As New FileSystemObject
On Error GoTo l:
fso.DeleteFile (App.Path & "\Data\Users\" & Text10.Text)
Label24.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Text8.Text = Label23.Caption

Label22.Caption = ""
Label23.Caption = ""
Text10.Text = ""
PushButton10.Enabled = False
PushButton21.Enabled = False

TabControl2.SelectedItem = 0
l:
Else


End If
End Sub

Private Sub PushButton8_Click()
Dim fso As New FileSystemObject
Dim t As TextStream
If Text4.Text = "" Or Text5.Text = "" Or Len(Text8.Text) <> 4 Or Len(Text9.Text) <= 4 Or Len(Text9.Text) >= 8 Then
MsgBox "Application is Incompelete!", vbCritical

Else

If fso.FileExists(App.Path & "\Data\Users\" & Text8.Text) Then
MsgBox "User ID is Already Exits!", vbCritical
Else
Set t = fso.CreateTextFile(App.Path & "\Data\Users\" & Text8.Text, True)
t.WriteLine Text9.Text
t.WriteLine Text4.Text
t.WriteLine Text5.Text
t.WriteLine Text6.Text
t.WriteLine Text7.Text
t.Close
Dim OL As String
OL = "User ID : " & Text8.Text & vbCrLf
OL = OL + "Password : " + Text9.Text & vbCrLf
OL = OL + "Create Compelete."
MsgBox OL
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

End If

End If
End Sub

Private Sub PushButton9_Click()
Dim fso As New FileSystemObject
Dim t As TextStream
If Text10.Text <> "" Then
If fso.FileExists(App.Path & "\Data\Users\" & Text10.Text) Then
Set t = fso.OpenTextFile(App.Path & "\Data\Users\" & Text10.Text, ForReading)
On Error Resume Next
Label24.Caption = t.ReadLine
Label19.Caption = t.ReadLine
Label20.Caption = t.ReadLine
Label21.Caption = t.ReadLine
Label22.Caption = t.ReadLine
Label23.Caption = Text10.Text
PushButton10.Enabled = True
PushButton21.Enabled = True

Else
MsgBox "Invalid User ID!", vbCritical
PushButton10.Enabled = False
PushButton21.Enabled = False

End If
Else
PushButton10.Enabled = False
PushButton21.Enabled = False

MsgBox "Insert User ID", vbCritical
End If

End Sub

Private Sub Form_Load()
FormCutter1.SetGlobelWindow_Z_order Me, zTop
Combo1.Text = 30
End Sub

Private Sub PushButton1_Click()


Form3.Show vbModal, Me

End Sub

Private Sub PushButton2_Click()
TabControl1.SelectedItem = 1
End Sub

Private Sub PushButton3_Click()
TabControl1.SelectedItem = 2

End Sub

Private Sub PushButton4_Click()
If Not Text2.Text = "" Then
Dim fso As New FileSystemObject
If fso.FolderExists(App.Path & "\Data\Account log\" & Text2.Text) Then
If fso.FileExists(App.Path & "\Data\Account log\" & Text2.Text & "\c.dat") Then
Form5.Show vbModal, Me
Else
MsgBox "Account Is Closed", vbOKOnly
End If
Else
MsgBox "Sorry! Invalide id number .", vbOKOnly
End If
End If
End Sub

Private Sub PushButton5_Click()
If Not Text1.Text = "" Then

Dim fso As New FileSystemObject
If fso.FolderExists(App.Path & "\Data\Account log\" & Text1.Text) Then
If Not fso.FileExists(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat") Then MsgBox "Current Account Is Closed.But Customer Details Are Here.", vbOKOnly

Form4.Show vbModal, Me
Else
MsgBox "Sorry! Invalide id number .", vbOKOnly
End If
End If
End Sub

Private Sub PushButton6_Click()
Dim fso As New FileSystemObject
If fso.FolderExists(App.Path & "\Data\Closed Account Log\" & Text3.Text) And Text3.Text <> "" Then

 

Form8.Show vbModal, Me
Else
MsgBox "Sorry! Invalide Account number .", vbOKOnly
End If


End Sub

Private Sub PushButton7_Click()
Dim fso As New FileSystemObject
Dim f As file
Dim t As TextStream
Dim s As String, ss As String
Dim s1 As String, ss1 As String
Set t = fso.OpenTextFile(App.Path & "\Data\Log.exe", ForReading, True)
t.SkipLine
t.SkipLine
t.read (Len("io:<:?<JG:|FHTH:DJODF{JSHEPKHGESDRHKD{HNLDGNLLNML J KH{RTHRT{Hkrth[pth[th pktpgh"))
s = t.ReadLine
t.read (Len("hjvbdsklnsfio98y4h43j 5hp3hoig jkoh53 b34b34ou hibt 34   tiu43gh34t3uigiugiuugiugffgewkjl hg 3434i343t ui34hohto8uygdflvn w h efiuhguoiwerhfgoweh obhbaffvba;sadfdfs'; '; ;.'hm;.g;'letk eolkg ';l';l';lkgsd'lkg keejko e'elkbhodfsjksl grewoitwrolwiphpi ;j sefl/s;ljs'lsfl'"))
ss = t.ReadLine
t.read (Len("jklgfnsadjlhfodshgsdjvglksdfjlksdjldkghsdljghsdajghls;ajd hglhsdlgjsadhlhgig osfhoorh grwgsdfhdfhgFDAOj DFOJ DOHjd dfijdf;ksdjg ;ijhiojdf j djidfj ddf ihdih"))
s1 = t.ReadLine
t.read (Len("nbkgjbk lkn ,/.m;l mmt ' hm nl'mmj httrjhp' kphkrt[th l[\rjl\[h\[o4089599i4 9u5 808 8343y8y43y4hjb klfgblkdjklhl"))
ss1 = t.ReadLine
t.Close
If FlatEdit1.Text = s And FlatEdit2.Text = ss Then
GroupBox2.Visible = True
ElseIf FlatEdit1.Text = s1 And FlatEdit2.Text = ss1 Then
GroupBox2.Visible = True
red
Else
MsgBox "Invalide User Name or Password!", vbCritical
GroupBox2.Visible = False
red
End If

End Sub

Private Sub remove_Click()
ll.ListItems.remove ll.SelectedItem.Index
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next
If TabControl1.SelectedItem = 1 Then Text2.SetFocus
If TabControl1.SelectedItem = 2 Then Text1.SetFocus
If TabControl1.SelectedItem = 3 Then Text3.SetFocus
If TabControl1.SelectedItem = 5 Then FlatEdit4.SetFocus
'If TabControl1.SelectedItem = 5 Then PushButton10.SetFocus
If TabControl1.SelectedItem = 6 Then Text13.SetFocus
'If TabControl1.SelectedItem = 7 Then FlatEdit4.SetFocus
'If TabControl1.SelectedItem = 7 Then red

GroupBox2.Visible = False

GroupBox1.Visible = False
End Sub



Private Sub TabControl2_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If TabControl2.SelectedItem = 1 Then Text10.SetFocus
If TabControl2.SelectedItem = 0 Then Text4.SetFocus

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton5_Click
End Sub


Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton9_Click

End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton12_Click
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton13_Click
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton13_Click

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton4_Click

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then PushButton6_Click

End Sub

