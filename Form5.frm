VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Account Update"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10440
   DrawMode        =   14  'Copy Pen
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
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
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   495
         Left            =   9000
         TabIndex        =   25
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
         Left            =   9240
         TabIndex        =   24
         Top             =   1440
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   855
         Left            =   9240
         TabIndex        =   22
         Top             =   2880
         Visible         =   0   'False
         Width           =   1005
         _Version        =   786432
         _ExtentX        =   1773
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Close Current Account"
         Appearance      =   6
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
         TabIndex        =   21
         Top             =   600
         Width           =   975
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
         TabIndex        =   19
         Top             =   2160
         Width           =   615
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
         TabIndex        =   13
         Top             =   2520
         Width           =   495
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
         TabIndex        =   12
         Top             =   2880
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
         TabIndex        =   11
         Top             =   1800
         Width           =   615
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
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin MSComctlLib.ListView L 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5530
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
            Text            =   "Instalment"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "User ID"
            Object.Width           =   3572
         EndProperty
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1935
         Left            =   5280
         TabIndex        =   30
         Top             =   1680
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Add New"
         Transparent     =   -1  'True
         Appearance      =   6
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1200
            TabIndex        =   31
            Top             =   960
            Width           =   1695
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   1920
            TabIndex        =   33
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Update"
            Appearance      =   6
         End
         Begin VB.Label Label13 
            Caption         =   "Date : "
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Paid Money : "
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   1095
         End
      End
      Begin RichTextLib.RichTextBox R 
         Height          =   1095
         Left            =   6120
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"Form5.frx":5C12
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2520
         TabIndex        =   39
         Top             =   3600
         Width           =   45
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2520
         TabIndex        =   38
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Advance : "
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
         TabIndex        =   37
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Interest : "
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
         TabIndex        =   36
         Top             =   3600
         Width           =   1185
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/YYYY"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label Label20 
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
         TabIndex        =   27
         Top             =   2520
         Width           =   375
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
         TabIndex        =   26
         Top             =   0
         Width           =   1965
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
         TabIndex        =   20
         Top             =   600
         Width           =   705
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
         TabIndex        =   18
         Top             =   2520
         Width           =   690
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
         TabIndex        =   17
         Top             =   3960
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
         TabIndex        =   16
         Top             =   1080
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
         TabIndex        =   15
         Top             =   720
         Width           =   45
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
         TabIndex        =   14
         Top             =   360
         Width           =   45
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
         TabIndex        =   9
         Top             =   2880
         Width           =   1050
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
         TabIndex        =   8
         Top             =   3960
         Width           =   1395
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
         TabIndex        =   7
         Top             =   2160
         Width           =   1950
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
         TabIndex        =   5
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Date : "
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount : "
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name : "
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
         TabIndex        =   2
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name : "
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
         TabIndex        =   1
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   495
      Left            =   5040
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Text6.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As New FileSystemObject
Text5.Text = Date
Dim m As ListItem, t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat", ForReading)
t.ReadLine
t.ReadLine
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
Label24.Caption = Format(Label11.Caption / Text7.Text, "#.00")
Label25.Caption = Format(Label12.Caption - Label24.Caption, "#.00")

If l.ListItems(l.ListItems.Count).SubItems(3) <= 0 Then
PushButton2.Visible = True
Else
PushButton2.Visible = False
End If
j:

End Sub

Private Sub Form_Resize()
'On Error Resume Next
'TabControl1.Height = Me.Height - 500
'l.Height = Me.Height - (L.Top + 700)
'Me.Width = 10680

End Sub

Private Sub PushButton1_Click()
Dim fso As New FileSystemObject
Dim m As ListItem, t As TextStream
If Val(Text6.Text) >= 100 Then
If l.ListItems.Count >= 1 Then
If l.ListItems(l.ListItems.Count).SubItems(3) - Val(Text6.Text) < 0 Then
MsgBox "Your Account Closing Balance Is -  " & l.ListItems(l.ListItems.Count).SubItems(3)
Exit Sub
End If
End If
Set m = l.ListItems.add(, , Val(l.ListItems.Count + 1))

m.SubItems(1) = Text5.Text
m.SubItems(2) = Val(Text6.Text)
If l.ListItems.Count = 1 Then

m.SubItems(3) = Val(Val(Label11.Caption) + Val(Label11.Caption) * (Val(Text4.Text) / 100) * Val(Text7.Text) / 30) - Val(Text6.Text)
Else
m.SubItems(3) = Val(l.ListItems.Item(l.ListItems.Count - 1).SubItems(3)) - Val(Text6.Text)

End If
m.SubItems(4) = Form2.Label3.Caption
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat", ForAppending)
t.WriteLine l.ListItems.Count
t.WriteLine Text5.Text
t.WriteLine Val(Text6.Text)
t.WriteLine l.ListItems.Item(l.ListItems.Count).SubItems(3)
t.WriteLine Form2.Label3.Caption
t.Close
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat", ForReading)
Dim nm As String
Dim mn As String
nm = t.ReadLine
mn = t.ReadLine
r.Text = r.Text + l.ListItems.Item(l.ListItems.Count).SubItems(3) + vbCrLf
r.Text = r.Text & mn & vbCrLf
r.Text = r.Text + t.ReadAll
t.Close
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat", ForWriting, True)
t.Write r.Text
t.Close
r.Text = ""
If l.ListItems(l.ListItems.Count).SubItems(3) <= 0 Then
PushButton2.Visible = True
Else
PushButton2.Visible = False
End If
End If
Text6.Text = ""
End Sub

Private Sub PushButton2_Click()
Dim s As String
Dim fso As New FileSystemObject
Dim f As file
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
Dim t As TextStream
Set t = fso.OpenTextFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\c.dat", ForReading, True)
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
t.SkipLine
Dim aa As String
aa = t.ReadLine
On Error Resume Next
MkDir (App.Path & "\Data\Closed Account Log\" & Form2.Text2.Text)
fso.CopyFile App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat", App.Path & "\Data\Closed Account Log\" & Form2.Text2.Text & "\" & aa & ".dat"
t.Close
fso.DeleteFile (App.Path & "\Data\Account log\" & Form2.Text2.Text & "\a.dat")
Set fso = Nothing
Set f = fso.GetFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\c.dat")
f.Name = "closed 1.dat"
Set f = fso.GetFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\p1.dat")
f.Name = "closed 2.dat"
Set f = fso.GetFile(App.Path & "\Data\Account log\" & Form2.Text2.Text & "\p2.dat")
f.Name = "closed 3.dat"
MsgBox "Closed Compelete", vbOKOnly
Unload Me
End If
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
'r.SelPrint
End Sub

Private Sub PushButton4_Click()
Form7.Show vbModal, Me
End Sub

Private Sub Text6_Change()
If Val(Text6.Text) >= 0 Then
PushButton1.Visible = True
Else
PushButton1.Visible = False
End If
End Sub

