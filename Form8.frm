VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form8"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10440
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _Version        =   786432
      _ExtentX        =   18441
      _ExtentY        =   13361
      _StockProps     =   68
      DrawFocusRect   =   0   'False
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
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
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text2 
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
         Left            =   2520
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text3 
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
         Left            =   2520
         TabIndex        =   6
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text4 
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
         Left            =   2520
         TabIndex        =   5
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox Text7 
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
         Left            =   2520
         TabIndex        =   4
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         Left            =   9360
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   495
         Left            =   9000
         TabIndex        =   1
         Top             =   6960
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Print Account Log"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   615
         Left            =   9360
         TabIndex        =   2
         Top             =   3240
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin MSComctlLib.ListView L 
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5741
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "warikaya"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "sign"
            Object.Width           =   3572
         EndProperty
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Left            =   5400
         TabIndex        =   26
         Top             =   1080
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   5880
         TabIndex        =   30
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/YYYY"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/YYYY"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Account Number"
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
         Left            =   5400
         TabIndex        =   27
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name : "
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
         Height          =   225
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name : "
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
         Height          =   225
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount : "
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
         Height          =   225
         Left            =   480
         TabIndex        =   23
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Date : "
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
         Height          =   225
         Left            =   480
         TabIndex        =   22
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays : "
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
         TabIndex        =   21
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Instalments :"
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
         TabIndex        =   20
         Top             =   2280
         Width           =   1950
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Instalment :"
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
         TabIndex        =   19
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Date"
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
         TabIndex        =   18
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   17
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   16
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   15
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   14
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest :"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
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
         Left            =   8160
         TabIndex        =   12
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Log"
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
         Left            =   3840
         TabIndex        =   11
         Top             =   0
         Width           =   1965
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   2640
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub list()
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim file As file
       On Error GoTo u
    Set FSfolder = FS.GetFolder(App.Path & "\Data\Closed Account Log\" & Form2.Text3.Text)
   For Each file In FSfolder.Files
        DoEvents
 ComboBox1.AddItem file.Name
 
 Next file
 ComboBox1.Text = ComboBox1.list(0)
u:

    Set FSfolder = Nothing

End Sub


Private Sub ComboBox1_Click()
On Error Resume Next
Dim fso As New FileSystemObject
Dim m As ListItem, t As TextStream

Set t = fso.OpenTextFile(App.Path & "\Data\Closed Account Log\" & Form2.Text3.Text & "\" & ComboBox1.Text, ForReading)
t.SkipLine
t.SkipLine

Label9.Caption = t.ReadLine
Label10.Caption = t.ReadLine
Label11.Caption = t.ReadLine
Text1.Text = t.ReadLine
Text2.Text = t.ReadLine
Text3.Text = t.ReadLine
Text4.Text = t.ReadLine
Text7.Text = t.ReadLine
Label12.Caption = t.ReadLine
Text8.Text = t.ReadLine
On Error GoTo j:
Do Until t.AtEndOfStream = True
Set m = l.ListItems.add(, , t.ReadLine)
m.SubItems(1) = t.ReadLine
m.SubItems(2) = t.ReadLine
m.SubItems(3) = t.ReadLine
m.SubItems(4) = t.ReadLine

Loop
t.Close
j:


End Sub


Private Sub Form_Load()
On Error Resume Next
Dim fso As New FileSystemObject
Dim m As ListItem, t As TextStream
list
Set t = fso.OpenTextFile(App.Path & "\Data\Closed Account Log\" & Form2.Text3.Text & "\" & ComboBox1.Text, ForReading)
t.SkipLine
t.SkipLine
Label9.Caption = t.ReadLine
Label10.Caption = t.ReadLine
Label11.Caption = t.ReadLine
Text1.Text = t.ReadLine
Text2.Text = t.ReadLine
Text3.Text = t.ReadLine
Text4.Text = t.ReadLine
Text7.Text = t.ReadLine
Label12.Caption = t.ReadLine
Text8.Text = t.ReadLine
On Error GoTo j:
Do Until t.AtEndOfStream = True
Set m = l.ListItems.add(, , t.ReadLine)
m.SubItems(1) = t.ReadLine
m.SubItems(2) = t.ReadLine
m.SubItems(3) = t.ReadLine
m.SubItems(4) = t.ReadLine

Loop
t.Close
j:
End Sub

Private Sub Form_Resize()
'On Error Resume Next
'TabControl1.Height = Me.Height - 500
'l.Height = Me.Height - (L.Top + 700)
'Me.Width = 10680

End Sub

Private Sub Label23_Click()

End Sub

Private Sub PushButton2_Click()
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Me.PrintForm
End If
End Sub

Private Sub r_Change()

End Sub

Private Sub PushButton4_Click()
Form9.Show vbModal, Me
End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text6_Change()

End Sub

