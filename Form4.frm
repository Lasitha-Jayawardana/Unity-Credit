VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Account Details"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12630
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      _Version        =   786432
      _ExtentX        =   22251
      _ExtentY        =   16536
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Customer Details"
      Item(0).ControlCount=   50
      Item(0).Control(0)=   "Label3"
      Item(0).Control(1)=   "Label4"
      Item(0).Control(2)=   "Label5"
      Item(0).Control(3)=   "Label6"
      Item(0).Control(4)=   "Label7"
      Item(0).Control(5)=   "Label8"
      Item(0).Control(6)=   "Label9"
      Item(0).Control(7)=   "Text1"
      Item(0).Control(8)=   "Text2"
      Item(0).Control(9)=   "Text3"
      Item(0).Control(10)=   "Text4"
      Item(0).Control(11)=   "Text5"
      Item(0).Control(12)=   "Text6"
      Item(0).Control(13)=   "Text7"
      Item(0).Control(14)=   "Text8"
      Item(0).Control(15)=   "Label10"
      Item(0).Control(16)=   "Label11"
      Item(0).Control(17)=   "Label12"
      Item(0).Control(18)=   "Label13"
      Item(0).Control(19)=   "Label14"
      Item(0).Control(20)=   "Label15"
      Item(0).Control(21)=   "Label16"
      Item(0).Control(22)=   "Label17"
      Item(0).Control(23)=   "Text9"
      Item(0).Control(24)=   "Text10"
      Item(0).Control(25)=   "Text11"
      Item(0).Control(26)=   "Text12"
      Item(0).Control(27)=   "Text13"
      Item(0).Control(28)=   "Text14"
      Item(0).Control(29)=   "Text15"
      Item(0).Control(30)=   "Text16"
      Item(0).Control(31)=   "PushButton1"
      Item(0).Control(32)=   "Label18"
      Item(0).Control(33)=   "Label39"
      Item(0).Control(34)=   "Text34"
      Item(0).Control(35)=   "Text36"
      Item(0).Control(36)=   "Text37"
      Item(0).Control(37)=   "Label40"
      Item(0).Control(38)=   "Label41"
      Item(0).Control(39)=   "Label42"
      Item(0).Control(40)=   "Label52"
      Item(0).Control(41)=   "Text46"
      Item(0).Control(42)=   "Shape1"
      Item(0).Control(43)=   "Shape2"
      Item(0).Control(44)=   "Shape3"
      Item(0).Control(45)=   "Shape4"
      Item(0).Control(46)=   "Shape5"
      Item(0).Control(47)=   "PushButton4"
      Item(0).Control(48)=   "Label53"
      Item(0).Control(49)=   "Combo1"
      Item(1).Caption =   "Garenter 1 Details"
      Item(1).ControlCount=   36
      Item(1).Control(0)=   "Text17"
      Item(1).Control(1)=   "Label19"
      Item(1).Control(2)=   "Label20"
      Item(1).Control(3)=   "Label21"
      Item(1).Control(4)=   "Label22"
      Item(1).Control(5)=   "Label23"
      Item(1).Control(6)=   "Label24"
      Item(1).Control(7)=   "Label25"
      Item(1).Control(8)=   "Label26"
      Item(1).Control(9)=   "Text18"
      Item(1).Control(10)=   "Text19"
      Item(1).Control(11)=   "Text20"
      Item(1).Control(12)=   "Text21"
      Item(1).Control(13)=   "Text22"
      Item(1).Control(14)=   "Text23"
      Item(1).Control(15)=   "Text24"
      Item(1).Control(16)=   "Label27"
      Item(1).Control(17)=   "Text25"
      Item(1).Control(18)=   "Label28"
      Item(1).Control(19)=   "PushButton2"
      Item(1).Control(20)=   "Label43"
      Item(1).Control(21)=   "Text38"
      Item(1).Control(22)=   "Text39"
      Item(1).Control(23)=   "Text40"
      Item(1).Control(24)=   "Text41"
      Item(1).Control(25)=   "Label44"
      Item(1).Control(26)=   "Label45"
      Item(1).Control(27)=   "Label46"
      Item(1).Control(28)=   "Label47"
      Item(1).Control(29)=   "Shape6"
      Item(1).Control(30)=   "Shape7"
      Item(1).Control(31)=   "Shape8"
      Item(1).Control(32)=   "Shape9"
      Item(1).Control(33)=   "Shape10"
      Item(1).Control(34)=   "PushButton5"
      Item(1).Control(35)=   "Label54"
      Item(2).Caption =   "Garenter 2 Details"
      Item(2).ControlCount=   37
      Item(2).Control(0)=   "Text26"
      Item(2).Control(1)=   "Label29"
      Item(2).Control(2)=   "Label30"
      Item(2).Control(3)=   "Label31"
      Item(2).Control(4)=   "Text27"
      Item(2).Control(5)=   "Text28"
      Item(2).Control(6)=   "Text29"
      Item(2).Control(7)=   "Text30"
      Item(2).Control(8)=   "Text31"
      Item(2).Control(9)=   "Text32"
      Item(2).Control(10)=   "Label37"
      Item(2).Control(11)=   "Text33"
      Item(2).Control(12)=   "Label38"
      Item(2).Control(13)=   "PushButton3"
      Item(2).Control(14)=   "r"
      Item(2).Control(15)=   "Text35"
      Item(2).Control(16)=   "Label32"
      Item(2).Control(17)=   "Label33"
      Item(2).Control(18)=   "Label34"
      Item(2).Control(19)=   "Label35"
      Item(2).Control(20)=   "Label36"
      Item(2).Control(21)=   "Label48"
      Item(2).Control(22)=   "Label49"
      Item(2).Control(23)=   "Label50"
      Item(2).Control(24)=   "Label51"
      Item(2).Control(25)=   "Text42"
      Item(2).Control(26)=   "Text43"
      Item(2).Control(27)=   "Text44"
      Item(2).Control(28)=   "Text45"
      Item(2).Control(29)=   "Label1"
      Item(2).Control(30)=   "Shape11"
      Item(2).Control(31)=   "Shape12"
      Item(2).Control(32)=   "Shape13"
      Item(2).Control(33)=   "Shape14"
      Item(2).Control(34)=   "Shape15"
      Item(2).Control(35)=   "PushButton6"
      Item(2).Control(36)=   "Label55"
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form4.frx":5C12
         Left            =   480
         List            =   "Form4.frx":5C1F
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   960
         Width           =   1815
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   495
         Left            =   10080
         TabIndex        =   102
         Top             =   8520
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin VB.TextBox Text46 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   100
         Top             =   8880
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Text            =   "jgxgaXKJbzkjxkjaK"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   46
         Text            =   "Form4.frx":5C5B
         Top             =   1440
         Width           =   8775
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "Form4.frx":5C61
         Top             =   2160
         Width           =   8775
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   2880
         Width           =   8775
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   3360
         Width           =   8775
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   42
         Text            =   "Text6"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   41
         Text            =   "Text7"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         TabIndex        =   40
         Text            =   "Text8"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   4320
         Width           =   8775
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   4800
         Width           =   8775
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   5280
         Width           =   2175
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   6240
         Width           =   8895
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   7680
         Width           =   2415
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   33
         Text            =   "Text4"
         Top             =   7680
         Width           =   2415
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   8400
         Width           =   2415
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -67600
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "Form4.frx":5C67
         Top             =   840
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67600
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67600
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   2040
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67600
         TabIndex        =   27
         Text            =   "Text4"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67600
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   3000
         Visible         =   0   'False
         Width           =   7695
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66880
         TabIndex        =   25
         Text            =   "Text6"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -63880
         TabIndex        =   24
         Text            =   "Text7"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60760
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67600
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   5880
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -67720
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "Form4.frx":5C6D
         Top             =   840
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.TextBox Text27 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67720
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67720
         TabIndex        =   19
         Text            =   "Text3"
         Top             =   2040
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.TextBox Text29 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67720
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text30 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67000
         TabIndex        =   17
         Text            =   "Text6"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -64000
         TabIndex        =   16
         Text            =   "Text7"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text32 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60880
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text33 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67720
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   5760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text34 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         TabIndex        =   12
         Text            =   "Text34"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text35 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67720
         TabIndex        =   11
         Text            =   "Text35"
         Top             =   3000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox Text36 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7680
         TabIndex        =   10
         Text            =   "Text36"
         Top             =   7680
         Width           =   2295
      End
      Begin VB.TextBox Text37 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10200
         TabIndex        =   9
         Text            =   "Text37"
         Top             =   7680
         Width           =   2175
      End
      Begin VB.TextBox Text38 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -68080
         TabIndex        =   8
         Text            =   "Text38"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text39 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65320
         TabIndex        =   7
         Text            =   "Text39"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text40 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -62800
         TabIndex        =   6
         Text            =   "Text40"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text41 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60280
         TabIndex        =   5
         Text            =   "Text41"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text42 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -67960
         TabIndex        =   4
         Text            =   "Text42"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65320
         TabIndex        =   3
         Text            =   "Text43"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text44 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -62680
         TabIndex        =   2
         Text            =   "Text44"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text45 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -60160
         TabIndex        =   1
         Text            =   "Text45"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin RichTextLib.RichTextBox r 
         Height          =   135
         Left            =   -52120
         TabIndex        =   13
         Top             =   9360
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   238
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Form4.frx":5C73
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   11280
         TabIndex        =   31
         Top             =   8520
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Next Page"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   495
         Left            =   -58720
         TabIndex        =   48
         Top             =   8520
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Next Page"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   495
         Left            =   -58720
         TabIndex        =   49
         Top             =   8520
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Update"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   495
         Left            =   -59920
         TabIndex        =   103
         Top             =   8520
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   495
         Left            =   -59920
         TabIndex        =   104
         Top             =   8520
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Garenter 2  Details"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -65440
         TabIndex        =   107
         Top             =   360
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Garenter 1 Details"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -65560
         TabIndex        =   106
         Top             =   360
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Details"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4560
         TabIndex        =   105
         Top             =   360
         Width           =   2700
      End
      Begin VB.Shape Shape15 
         Height          =   735
         Left            =   -68080
         Top             =   4920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape Shape14 
         Height          =   375
         Left            =   -68080
         Top             =   4560
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Shape Shape13 
         Height          =   1095
         Left            =   -62800
         Top             =   4560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Shape Shape12 
         Height          =   1095
         Left            =   -60280
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Shape Shape11 
         Height          =   1095
         Left            =   -65440
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape Shape10 
         Height          =   375
         Left            =   -68200
         Top             =   4560
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Shape Shape9 
         Height          =   1095
         Left            =   -62920
         Top             =   4560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Shape Shape8 
         Height          =   1095
         Left            =   -60400
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Shape Shape7 
         Height          =   1095
         Left            =   -65560
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape Shape6 
         Height          =   1095
         Left            =   -68200
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape Shape5 
         Height          =   375
         Left            =   2280
         Top             =   7200
         Width           =   10215
      End
      Begin VB.Shape Shape4 
         Height          =   1095
         Left            =   10080
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         Height          =   1095
         Left            =   4920
         Top             =   7200
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   2280
         Top             =   7200
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   2280
         Top             =   7200
         Width           =   10215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69160
         TabIndex        =   101
         Top             =   4080
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         Height          =   375
         Left            =   480
         TabIndex        =   99
         Top             =   8880
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name with Initials :     "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   98
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Names Denoted by Init. : "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   97
         Top             =   2160
         Width           =   2040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address :      "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   96
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Address :        "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   95
         Top             =   3360
         Width           =   1770
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tele. num :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   94
         Top             =   3840
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home -                   "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2640
         TabIndex        =   93
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile -                 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5520
         TabIndex        =   92
         Top             =   3840
         Width           =   840
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   91
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Business Name :          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   90
         Top             =   4320
         Width           =   1845
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Business Address :       "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   89
         Top             =   4800
         Width           =   1860
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Individual/Join        "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   88
         Top             =   5280
         Width           =   1530
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   87
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reasone for Loan :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   86
         Top             =   6240
         Width           =   2190
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Bank         "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   85
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   84
         Top             =   6840
         Width           =   2805
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   8400
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile -                 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -64720
         TabIndex        =   82
         Top             =   3480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home -                   "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -67600
         TabIndex        =   81
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tele. num :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   80
         Top             =   3480
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employment :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   79
         Top             =   3000
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   78
         Top             =   2520
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Address :                  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   77
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address :                "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   76
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full name           "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   75
         Top             =   840
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -61720
         TabIndex        =   74
         Top             =   3480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   255
         Left            =   -69640
         TabIndex        =   73
         Top             =   5880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile -                 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -64840
         TabIndex        =   72
         Top             =   3480
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home -                   "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -67720
         TabIndex        =   71
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -61840
         TabIndex        =   70
         Top             =   3480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   255
         Left            =   -69520
         TabIndex        =   69
         Top             =   5760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File NO           "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9240
         TabIndex        =   68
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   5040
         TabIndex        =   67
         Top             =   7320
         Width           =   510
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         Height          =   195
         Left            =   7680
         TabIndex        =   66
         Top             =   7320
         Width           =   1005
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   10440
         TabIndex        =   65
         Top             =   7320
         Width           =   900
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69280
         TabIndex        =   64
         Top             =   4080
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Bank         "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -67960
         TabIndex        =   63
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   -65320
         TabIndex        =   62
         Top             =   4680
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         Height          =   195
         Left            =   -62680
         TabIndex        =   61
         Top             =   4680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   -60160
         TabIndex        =   60
         Top             =   4680
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tele. num :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   59
         Top             =   3480
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full name           "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69640
         TabIndex        =   58
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address :                "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   57
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Address :                  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   56
         Top             =   2040
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number :              "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   55
         Top             =   2520
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employment :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69640
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Bank         "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -67960
         TabIndex        =   53
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   -65200
         TabIndex        =   52
         Top             =   4680
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         Height          =   195
         Left            =   -62560
         TabIndex        =   51
         Top             =   4680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   -60040
         TabIndex        =   50
         Top             =   4680
         Visible         =   0   'False
         Width           =   900
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Boolean

Private Sub Form_Load()
Combo1.ListIndex = 0
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
Text37.Text = ""
Text38.Text = ""
Text39.Text = ""
Text40.Text = ""
Text41.Text = ""
Text42.Text = ""
Text43.Text = ""
Text44.Text = ""
Text45.Text = ""

Dim fso As New FileSystemObject
Dim t As TextStream
If fso.FileExists(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\closed 1.dat") Then b = True
If b = True Then
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\closed 1.dat", ForReading)

Else
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\c.dat", ForReading)

End If
Text1.Text = t.ReadLine
 Text2.Text = t.ReadLine
 Text3.Text = t.ReadLine
Text4.Text = t.ReadLine
Text5.Text = t.ReadLine
Text6.Text = t.ReadLine
Text7.Text = t.ReadLine
Text8.Text = t.ReadLine
 Text9.Text = t.ReadLine
 Text10.Text = t.ReadLine
 Text11.Text = t.ReadLine
 Text12.Text = t.ReadLine
 Text13.Text = t.ReadLine
 Text14.Text = t.ReadLine
 Text15.Text = t.ReadLine
  Text36.Text = t.ReadLine
 Text37.Text = t.ReadLine
 Text16.Text = t.ReadLine
Text34.Text = t.ReadLine
On Error Resume Next

Text46.Text = t.ReadLine

t.Close

Set fso = Nothing
If b = True Then
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\closed 2.dat", ForReading)
Else
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\p1.dat", ForReading)
End If
 Text17.Text = t.ReadLine
 Text18.Text = t.ReadLine
 Text19.Text = t.ReadLine
 Text20.Text = t.ReadLine
 Text21.Text = t.ReadLine
 Text22.Text = t.ReadLine
 Text23.Text = t.ReadLine
 Text24.Text = t.ReadLine
  Text38.Text = t.ReadLine
 Text39.Text = t.ReadLine
 Text40.Text = t.ReadLine
 Text41.Text = t.ReadLine
 Text25.Text = t.ReadLine

 
 t.Close
Set fso = Nothing
If b = True Then
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\closed 3.dat", ForReading)
Else
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text1.Text & "\p2.dat", ForReading)
End If
 Text26.Text = t.ReadLine
 Text27.Text = t.ReadLine
 Text28.Text = t.ReadLine
 Text29.Text = t.ReadLine
 Text35.Text = t.ReadLine
 Text30.Text = t.ReadLine
Text31.Text = t.ReadLine
Text32.Text = t.ReadLine
 Text42.Text = t.ReadLine
 Text43.Text = t.ReadLine
Text44.Text = t.ReadLine
Text45.Text = t.ReadLine

 Text33.Text = t.ReadLine

 t.Close
Set fso = Nothing

End Sub

Private Sub PushButton1_Click()
TabControl1.SelectedItem = 1
End Sub

Private Sub PushButton2_Click()
TabControl1.SelectedItem = 2
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
MkDir App.Path & "\Data\Account log\" & Text1.Text
r.Text = ""
r.Text = Text1.Text & vbCrLf
r.Text = r.Text + Text2.Text & vbCrLf
r.Text = r.Text + Text3.Text & vbCrLf
r.Text = r.Text + Text4.Text & vbCrLf
r.Text = r.Text + Text5.Text & vbCrLf
r.Text = r.Text + Text6.Text & vbCrLf
r.Text = r.Text + Text7.Text & vbCrLf
r.Text = r.Text + Text8.Text & vbCrLf
r.Text = r.Text + Text9.Text & vbCrLf
r.Text = r.Text + Text10.Text & vbCrLf
r.Text = r.Text + Text11.Text & vbCrLf
r.Text = r.Text + Trim(Str(Val(Text12.Text))) & vbCrLf
r.Text = r.Text + Text13.Text & vbCrLf
r.Text = r.Text + Text14.Text & vbCrLf
r.Text = r.Text + Text15.Text & vbCrLf
r.Text = r.Text + Text36.Text & vbCrLf
r.Text = r.Text + Text37.Text & vbCrLf
r.Text = r.Text + Text16.Text & vbCrLf
r.Text = r.Text + Text34.Text & vbCrLf
r.Text = r.Text + Text46.Text
Dim fso As New FileSystemObject
Dim t As TextStream
If b = True Then
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 1.dat", True)
Else
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat", True)
End If
t.Write r.Text
t.Close
Set fso = Nothing
r.Text = ""

r.Text = Text17.Text & vbCrLf
r.Text = r.Text + Text18.Text & vbCrLf
r.Text = r.Text + Text19.Text & vbCrLf
r.Text = r.Text + Text20.Text & vbCrLf
r.Text = r.Text + Text21.Text & vbCrLf
r.Text = r.Text + Text22.Text & vbCrLf
r.Text = r.Text + Text23.Text & vbCrLf
r.Text = r.Text + Text24.Text & vbCrLf
r.Text = r.Text + Text38.Text & vbCrLf
r.Text = r.Text + Text39.Text & vbCrLf
r.Text = r.Text + Text40.Text & vbCrLf
r.Text = r.Text + Text41.Text & vbCrLf
r.Text = r.Text + Text25.Text
If b = True Then
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 2.dat", True)
Else
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\p1.dat", True)
End If
t.Write r.Text
t.Close
Set fso = Nothing
r.Text = ""

r.Text = Text26.Text & vbCrLf
r.Text = r.Text + Text27.Text & vbCrLf
r.Text = r.Text + Text28.Text & vbCrLf
r.Text = r.Text + Text29.Text & vbCrLf
r.Text = r.Text + Text35.Text & vbCrLf
r.Text = r.Text + Text30.Text & vbCrLf
r.Text = r.Text + Text31.Text & vbCrLf
r.Text = r.Text + Text32.Text & vbCrLf
r.Text = r.Text + Text42.Text & vbCrLf
r.Text = r.Text + Text43.Text & vbCrLf
r.Text = r.Text + Text44.Text & vbCrLf
r.Text = r.Text + Text45.Text & vbCrLf
r.Text = r.Text + Text33.Text
If b = True Then
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 3.dat", True)
Else
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\p2.dat", True)
End If
t.Write r.Text
t.Close
Set fso = Nothing
r.Text = ""
MsgBox "Update Compelete.", vbOKOnly
End Sub

Private Sub PushButton5_Click()
On Error Resume Next
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Me.PrintForm
End If
End Sub

Private Sub PushButton6_Click()
On Error Resume Next
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Me.PrintForm
End If
End Sub

Private Sub PushButton4_Click()
On Error Resume Next
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Me.PrintForm
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 10 Then
Combo1.ListIndex = 0
ElseIf Len(Text1.Text) = 8 Then
Combo1.ListIndex = 1
Else
Combo1.ListIndex = 2
End If
End Sub
