VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Application"
   ClientHeight    =   8895
   ClientLeft      =   2685
   ClientTop       =   2385
   ClientWidth     =   12930
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8895
      Left            =   0
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Width           =   12975
      _Version        =   786432
      _ExtentX        =   22886
      _ExtentY        =   15690
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Customer Application"
      Item(0).ControlCount=   52
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
      Item(0).Control(25)=   "Text12"
      Item(0).Control(26)=   "Text14"
      Item(0).Control(27)=   "Text15"
      Item(0).Control(28)=   "Text16"
      Item(0).Control(29)=   "PushButton1"
      Item(0).Control(30)=   "Label18"
      Item(0).Control(31)=   "Label39"
      Item(0).Control(32)=   "Text34"
      Item(0).Control(33)=   "Text36"
      Item(0).Control(34)=   "Text37"
      Item(0).Control(35)=   "Label40"
      Item(0).Control(36)=   "Label41"
      Item(0).Control(37)=   "Label42"
      Item(0).Control(38)=   "Label52"
      Item(0).Control(39)=   "Text46"
      Item(0).Control(40)=   "Shape1"
      Item(0).Control(41)=   "Shape12"
      Item(0).Control(42)=   "Shape13"
      Item(0).Control(43)=   "Shape14"
      Item(0).Control(44)=   "Shape15"
      Item(0).Control(45)=   "Combo1"
      Item(0).Control(46)=   "Combo2"
      Item(0).Control(47)=   "Combo3"
      Item(0).Control(48)=   "PushButton4"
      Item(0).Control(49)=   "Shape16"
      Item(0).Control(50)=   "Shape17"
      Item(0).Control(51)=   "Shape18"
      Item(1).Caption =   "Garenter 1 Application"
      Item(1).ControlCount=   38
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
      Item(1).Control(29)=   "Shape2"
      Item(1).Control(30)=   "Shape8"
      Item(1).Control(31)=   "Shape9"
      Item(1).Control(32)=   "Shape10"
      Item(1).Control(33)=   "Shape11"
      Item(1).Control(34)=   "PushButton5"
      Item(1).Control(35)=   "Shape19"
      Item(1).Control(36)=   "Shape20"
      Item(1).Control(37)=   "Shape21"
      Item(2).Caption =   "Garenter 2 Application"
      Item(2).ControlCount=   39
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
      Item(2).Control(30)=   "Shape3"
      Item(2).Control(31)=   "Shape4"
      Item(2).Control(32)=   "Shape5"
      Item(2).Control(33)=   "Shape6"
      Item(2).Control(34)=   "Shape7"
      Item(2).Control(35)=   "PushButton6"
      Item(2).Control(36)=   "Shape22"
      Item(2).Control(37)=   "Shape23"
      Item(2).Control(38)=   "Shape24"
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   255
         Left            =   11880
         TabIndex        =   103
         Top             =   2820
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Same"
         Appearance      =   6
         Checked         =   -1  'True
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Form3.frx":5C12
         Left            =   2640
         List            =   "Form3.frx":5C1C
         TabIndex        =   11
         Top             =   4920
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Form3.frx":5C33
         Left            =   2640
         List            =   "Form3.frx":5C40
         TabIndex        =   13
         Text            =   "For the Business Activities."
         Top             =   5880
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Form3.frx":5C88
         Left            =   360
         List            =   "Form3.frx":5C95
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   8895
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   8895
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   8895
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   3000
         Width           =   8895
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6480
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9600
         TabIndex        =   8
         Text            =   "Text8"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3960
         Width           =   8895
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   4440
         Width           =   8895
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   5400
         Width           =   2535
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   7320
         Width           =   2415
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4800
         TabIndex        =   15
         Text            =   "Text4"
         Top             =   7320
         Width           =   2415
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   7920
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "Form3.frx":5CD1
         Top             =   960
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "Form3.frx":5CD7
         Top             =   1680
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "Form3.frx":5CDD
         Top             =   2520
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67480
         TabIndex        =   23
         Text            =   "Text4"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67480
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   4080
         Visible         =   0   'False
         Width           =   7935
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66760
         TabIndex        =   25
         Text            =   "Text6"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -63880
         TabIndex        =   26
         Text            =   "Text7"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -61000
         TabIndex        =   27
         Text            =   "Text8"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67600
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   7320
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "Form3.frx":5CE3
         Top             =   960
         Visible         =   0   'False
         Width           =   9015
      End
      Begin VB.TextBox Text27 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -67480
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "Form3.frx":5CE9
         Top             =   2520
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.TextBox Text29 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67480
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   3480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox Text30 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66760
         TabIndex        =   39
         Text            =   "Text6"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -63880
         TabIndex        =   40
         Text            =   "Text7"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text32 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -61120
         TabIndex        =   41
         Text            =   "Text8"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text33 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67480
         TabIndex        =   46
         Text            =   "Text5"
         Top             =   6960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text34 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10200
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text35 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67480
         TabIndex        =   38
         Text            =   "Text35"
         Top             =   4080
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.TextBox Text36 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         TabIndex        =   16
         Text            =   "Text36"
         Top             =   7320
         Width           =   2175
      End
      Begin VB.TextBox Text37 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9840
         TabIndex        =   17
         Text            =   "Text37"
         Top             =   7320
         Width           =   2175
      End
      Begin VB.TextBox Text38 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67840
         TabIndex        =   28
         Text            =   "Text38"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text39 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -65200
         TabIndex        =   29
         Text            =   "Text39"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text40 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -62800
         TabIndex        =   30
         Text            =   "Text40"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text41 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -60400
         TabIndex        =   31
         Text            =   "Text41"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text42 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67960
         TabIndex        =   42
         Text            =   "Text42"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -65320
         TabIndex        =   43
         Text            =   "Text43"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text44 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -62800
         TabIndex        =   44
         Text            =   "Text44"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text45 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -60520
         TabIndex        =   45
         Text            =   "Text45"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text46 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         TabIndex        =   49
         Top             =   8400
         Width           =   1815
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   11520
         TabIndex        =   19
         Top             =   8040
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
         Left            =   -58480
         TabIndex        =   33
         Top             =   8040
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
         Left            =   -58480
         TabIndex        =   47
         Top             =   8040
         Visible         =   0   'False
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Create Account"
         Appearance      =   6
      End
      Begin RichTextLib.RichTextBox r 
         Height          =   735
         Left            =   -51400
         TabIndex        =   50
         Top             =   11280
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Form3.frx":5CEF
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   255
         Left            =   -58120
         TabIndex        =   104
         Top             =   2370
         Visible         =   0   'False
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Same"
         Appearance      =   6
         Checked         =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   255
         Left            =   -58240
         TabIndex        =   105
         Top             =   2350
         Visible         =   0   'False
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Same"
         Appearance      =   6
         Checked         =   -1  'True
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   4
         Height          =   15
         Left            =   -58480
         Top             =   2160
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   4
         Height          =   15
         Left            =   -58480
         Top             =   2760
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   4
         Height          =   615
         Left            =   -58360
         Top             =   2160
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   4
         Height          =   15
         Left            =   -58360
         Top             =   2160
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape20 
         BorderWidth     =   4
         Height          =   15
         Left            =   -58360
         Top             =   2760
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape19 
         BorderWidth     =   4
         Height          =   615
         Left            =   -58240
         Top             =   2160
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   4
         Height          =   15
         Left            =   11640
         Top             =   2640
         Width           =   135
      End
      Begin VB.Shape Shape17 
         BorderWidth     =   4
         Height          =   15
         Left            =   11640
         Top             =   3240
         Width           =   135
      End
      Begin VB.Shape Shape16 
         BorderWidth     =   4
         Height          =   615
         Left            =   11760
         Top             =   2640
         Width           =   15
      End
      Begin VB.Shape Shape15 
         Height          =   375
         Left            =   1920
         Top             =   6720
         Width           =   10335
      End
      Begin VB.Shape Shape14 
         Height          =   1095
         Left            =   7320
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Shape Shape13 
         Height          =   1095
         Left            =   4680
         Top             =   6720
         Width           =   2655
      End
      Begin VB.Shape Shape12 
         Height          =   1095
         Left            =   1920
         Top             =   6720
         Width           =   2775
      End
      Begin VB.Shape Shape11 
         Height          =   375
         Left            =   -67960
         Top             =   5880
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.Shape Shape10 
         Height          =   1095
         Left            =   -62920
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Shape Shape9 
         Height          =   1095
         Left            =   -65320
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Shape Shape8 
         Height          =   1095
         Left            =   -67960
         Top             =   5880
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape Shape7 
         Height          =   1095
         Left            =   -62920
         Top             =   5640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Shape Shape6 
         Height          =   1095
         Left            =   -65440
         Top             =   5640
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Shape Shape5 
         Height          =   1095
         Left            =   -68200
         Top             =   5640
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   -68200
         Top             =   5640
         Visible         =   0   'False
         Width           =   9975
      End
      Begin VB.Shape Shape3 
         Height          =   1095
         Left            =   -68200
         Top             =   5640
         Visible         =   0   'False
         Width           =   9975
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   -67960
         Top             =   5880
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   1920
         Top             =   6720
         Width           =   10335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69160
         TabIndex        =   101
         Top             =   5280
         Visible         =   0   'False
         Width           =   2505
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
         Left            =   360
         TabIndex        =   100
         Top             =   1320
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
         Left            =   360
         TabIndex        =   99
         Top             =   1800
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
         Left            =   360
         TabIndex        =   98
         Top             =   2520
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
         Left            =   360
         TabIndex        =   97
         Top             =   3000
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
         Left            =   360
         TabIndex        =   96
         Top             =   3480
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
         TabIndex        =   95
         Top             =   3480
         Width           =   735
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
         Left            =   5640
         TabIndex        =   94
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         Height          =   255
         Left            =   8760
         TabIndex        =   93
         Top             =   3480
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
         Left            =   360
         TabIndex        =   92
         Top             =   3960
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
         Left            =   360
         TabIndex        =   91
         Top             =   4440
         Width           =   1860
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Individual/Joint         "
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
         Left            =   360
         TabIndex        =   90
         Top             =   4920
         Width           =   1620
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
         Left            =   360
         TabIndex        =   89
         Top             =   5400
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
         Left            =   360
         TabIndex        =   88
         Top             =   5880
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
         Left            =   2160
         TabIndex        =   87
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   86
         Top             =   6360
         Width           =   2625
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   195
         Left            =   480
         TabIndex        =   85
         Top             =   7920
         Width           =   1395
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
         Left            =   -64600
         TabIndex        =   84
         Top             =   4920
         Visible         =   0   'False
         Width           =   600
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
         TabIndex        =   83
         Top             =   4920
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
         Left            =   -69520
         TabIndex        =   82
         Top             =   4800
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
         Left            =   -69520
         TabIndex        =   81
         Top             =   4200
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
         Left            =   -69520
         TabIndex        =   80
         Top             =   3360
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   -69520
         TabIndex        =   79
         Top             =   2640
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
         Left            =   -69520
         TabIndex        =   78
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   -69520
         TabIndex        =   77
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         Height          =   255
         Left            =   -61840
         TabIndex        =   76
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   255
         Left            =   -69520
         TabIndex        =   75
         Top             =   7440
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
         Left            =   -64720
         TabIndex        =   74
         Top             =   4680
         Visible         =   0   'False
         Width           =   720
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
         Left            =   -67600
         TabIndex        =   73
         Top             =   4680
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Business -               "
         Height          =   255
         Left            =   -61960
         TabIndex        =   72
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Date sing                "
         Height          =   255
         Left            =   -69520
         TabIndex        =   71
         Top             =   7080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File NO        "
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
         Left            =   9480
         TabIndex        =   70
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   4800
         TabIndex        =   69
         Top             =   6840
         Width           =   510
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type :"
         Height          =   195
         Left            =   7440
         TabIndex        =   68
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   9840
         TabIndex        =   67
         Top             =   6840
         Width           =   900
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account Details          "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69160
         TabIndex        =   66
         Top             =   5400
         Visible         =   0   'False
         Width           =   2625
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
         Left            =   -67840
         TabIndex        =   65
         Top             =   6000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   -65080
         TabIndex        =   64
         Top             =   6000
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         Height          =   195
         Left            =   -62680
         TabIndex        =   63
         Top             =   6000
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   -60400
         TabIndex        =   62
         Top             =   6000
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
         Left            =   -69520
         TabIndex        =   61
         Top             =   4680
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
         Height          =   225
         Left            =   -69520
         TabIndex        =   60
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
         Left            =   -69520
         TabIndex        =   59
         Top             =   1680
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
         Left            =   -69520
         TabIndex        =   58
         Top             =   2640
         Visible         =   0   'False
         Width           =   2220
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
         Left            =   -69520
         TabIndex        =   57
         Top             =   3480
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
         Left            =   -69520
         TabIndex        =   56
         Top             =   4080
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
         TabIndex        =   55
         Top             =   5760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   -65200
         TabIndex        =   54
         Top             =   5760
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
         Height          =   195
         Left            =   -62680
         TabIndex        =   53
         Top             =   5760
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Left            =   -60280
         TabIndex        =   52
         Top             =   5760
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         Height          =   195
         Left            =   480
         TabIndex        =   51
         Top             =   8400
         Width           =   630
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add()
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 1.dat", ForReading)
t.SkipLine
Text2.Text = t.ReadLine
 Text3.Text = t.ReadLine
Text4.Text = t.ReadLine
Text5.Text = t.ReadLine
Text6.Text = t.ReadLine
Text7.Text = t.ReadLine
Text8.Text = t.ReadLine
 Text9.Text = t.ReadLine
 Text10.Text = t.ReadLine
 Combo3.Text = t.ReadLine
 Text12.Text = t.ReadLine
 Combo2.Text = t.ReadLine
 Text14.Text = t.ReadLine
 Text15.Text = t.ReadLine
  Text36.Text = t.ReadLine
 Text37.Text = t.ReadLine
 t.ReadLine
Text34.Text = t.ReadLine
On Error Resume Next

Text46.Text = t.ReadLine

t.Close

Set fso = Nothing
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 2.dat", ForReading)
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
 t.ReadLine

 
 t.Close
Set fso = Nothing
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\closed 3.dat", ForReading)
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

 t.ReadLine

 t.Close
Set fso = Nothing

End Sub

Private Sub ComboBox1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo3.ListIndex = 0

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
'Text11.Text = ""
Text12.Text = ""
'Combo2.Text = ""
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
Text46.Text = Form2.Label3.Caption
Text16.Text = Date
Text25.Text = Date
Text33.Text = Date

End Sub

Private Sub Label2_Click()
End Sub

Private Sub PushButton1_Click()
TabControl1.SelectedItem = 1
End Sub

Private Sub PushButton2_Click()
TabControl1.SelectedItem = 2
End Sub

Private Sub PushButton3_Click()
If Text2.Text = "" Then
MsgBox "Please enter Customer Name.", vbCritical
TabControl1.SelectedItem = 0
Text2.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "Please enter Customer Address.", vbCritical
TabControl1.SelectedItem = 0
Text4.SetFocus
Exit Sub
End If



If Len(Text1.Text) = 10 Or Len(Text1.Text) = 8 Or Len(Text1.Text) = 6 Then
Else
MsgBox "Please enter Customer N.I.C. or Licence or Passport Number", vbCritical
TabControl1.SelectedItem = 0
Text1.SetFocus
Exit Sub
End If
If 0 = Val(Text12.Text) Then
MsgBox "Please enter Customer Loan Amount", vbCritical
TabControl1.SelectedItem = 0
Text12.SetFocus
Exit Sub
End If

Dim sd As String
Dim fso1 As New FileSystemObject
Dim t1 As TextStream
Set t1 = fso1.OpenTextFile(App.Path & "\Data\Count.txt", ForReading)
Dim s As String
s = t1.ReadLine

Text34.Text = Val(s) + 1
sd = MsgBox("Your Application No Is : " & Text34.Text & vbCrLf & "Your Account Number Is : " & Text1.Text & vbCrLf & "Create Now ?", vbYesNo)
t1.Close
If sd = vbYes Then
Set t1 = fso1.CreateTextFile(App.Path & "\Data\Count.txt", True)
t1.Write Text34.Text
t1.Close

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
r.Text = r.Text + Combo3.Text & vbCrLf
r.Text = r.Text + Trim(Str(Val(Text12.Text))) & vbCrLf
r.Text = r.Text + Combo2.Text & vbCrLf
r.Text = r.Text + Text14.Text & vbCrLf
r.Text = r.Text + Text15.Text & vbCrLf
r.Text = r.Text + Text36.Text & vbCrLf
r.Text = r.Text + Text37.Text & vbCrLf
r.Text = r.Text + Text16.Text & vbCrLf
r.Text = r.Text + Text34.Text & vbCrLf
r.Text = r.Text + Text46.Text

Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat", True)
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
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\p1.dat", True)
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
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Text1.Text & "\p2.dat", True)
t.Write r.Text
t.Close
Set fso = Nothing
r.Text = ""
Form6.Label9.Caption = ""
Form6.Label10.Caption = ""
Form6.Label11.Caption = ""
Form6.Label12.Caption = "00 0"
Form6.Text1.Text = ""
Form6.Text2.Text = "0"
Form6.Text3.Text = ""
Form6.Text4.Text = "7"
'Form6.Combo1.Text = "30"
Form6.Text1.Text = Date
Form6.Text5.Text = Form2.Label3.Caption
Form6.Label9.Caption = Text2.Text
Form6.Label10.Caption = Text9.Text
Form6.Label11.Caption = Text12.Text
PushButton3.Enabled = False

On Error Resume Next
fso.DeleteFile (App.Path & "\Data\Account log\" & Text1.Text & "\closed 1.dat")
fso.DeleteFile (App.Path & "\Data\Account log\" & Text1.Text & "\closed 2.dat")
fso.DeleteFile (App.Path & "\Data\Account log\" & Text1.Text & "\closed 3.dat")
Form6.Show vbModal, Me
Else
Exit Sub
End If
Unload Form3
End Sub


Private Sub PushButton4_Click()
If PushButton4.Checked = True Then
PushButton4.Checked = False
Else
PushButton4.Checked = True
Text5.Text = Text4.Text
End If
End Sub

Private Sub PushButton5_Click()
If PushButton5.Checked = True Then
PushButton5.Checked = False
Else
PushButton5.Checked = True
Text19.Text = Text18.Text
End If

End Sub

Private Sub PushButton6_Click()
If PushButton6.Checked = True Then
PushButton6.Checked = False
Else
PushButton6.Checked = True
Text28.Text = Text27.Text
End If

End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If TabControl1.SelectedItem <> 0 Then

If Len(Text1.Text) = 10 Or Len(Text1.Text) = 8 Or Len(Text1.Text) = 6 Then
Else
MsgBox "Please enter Customer N.I.C. or Licence or Passport Number", vbCritical
TabControl1.SelectedItem = 0
Text1.SetFocus

End If
End If
If TabControl1.SelectedItem = 1 Then Text17.SetFocus
If TabControl1.SelectedItem = 2 Then Text26.SetFocus

End Sub

Private Sub Text1_Change()
Dim fso As New FileSystemObject
If Len(Text1.Text) = 10 Then
 Combo1.ListIndex = 0
If fso.FolderExists(App.Path & "\Data\Account log\" & Text1.Text) Then
If fso.FileExists(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat") Then
MsgBox "Sorry! Account is Already Exists.", vbOKOnly
Text1.Text = ""
Else
add
End If
End If
ElseIf Len(Text1.Text) = 8 Then
 Combo1.ListIndex = 1
If fso.FolderExists(App.Path & "\Data\Account log\" & Text1.Text) Then
If fso.FileExists(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat") Then
MsgBox "Sorry! Account is Already Exists.", vbOKOnly
Text1.Text = ""
Else
add
End If
End If
ElseIf Len(Text1.Text) = 6 Then
Combo1.ListIndex = 2
If fso.FolderExists(App.Path & "\Data\Account log\" & Text1.Text) Then
If fso.FileExists(App.Path & "\Data\Account log\" & Text1.Text & "\c.dat") Then
MsgBox "Sorry! Account is Already Exists.", vbOKOnly
Text1.Text = ""
Else
add
End If
End If

End If
End Sub

Private Sub Text18_Change()
If PushButton5.Checked = True Then
Text19.Text = Text18.Text
End If
End Sub

Private Sub Text27_Change()
If PushButton6.Checked = True Then
Text28.Text = Text27.Text
End If
End Sub

Private Sub Text4_Change()
If PushButton4.Checked = True Then
Text5.Text = Text4.Text
End If
End Sub
