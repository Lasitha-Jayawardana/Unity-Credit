VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Account Log"
   ClientHeight    =   6600
   ClientLeft      =   7275
   ClientTop       =   4035
   ClientWidth     =   10320
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _Version        =   786432
      _ExtentX        =   18230
      _ExtentY        =   11668
      _StockProps     =   68
      DrawFocusRect   =   0   'False
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin MSComctlLib.ListView ll 
         Height          =   1215
         Left            =   1200
         TabIndex        =   29
         Top             =   6840
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form6.frx":5C12
         Left            =   2280
         List            =   "Form6.frx":5C31
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9000
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   600
         Width           =   975
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   8520
         TabIndex        =   18
         Top             =   5640
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Create"
         Appearance      =   6
      End
      Begin VB.TextBox Text1 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   4800
         Width           =   1215
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
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   8280
         TabIndex        =   31
         Top             =   2280
         Width           =   45
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Interest : "
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
         Left            =   6000
         TabIndex        =   30
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/YYYY"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   4800
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/YYYY"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   8280
         TabIndex        =   25
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   8280
         TabIndex        =   24
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label17 
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
         Left            =   6000
         TabIndex        =   23
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label Label16 
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
         Left            =   6000
         TabIndex        =   22
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label14 
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
         Left            =   2880
         TabIndex        =   21
         Top             =   4200
         Width           =   180
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   5760
         Picture         =   "Form6.frx":5C5F
         Top             =   4440
         Width           =   1920
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         Height          =   255
         Left            =   7920
         TabIndex        =   19
         Top             =   600
         Width           =   1095
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
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1485
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
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   1470
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
         Left            =   360
         TabIndex        =   15
         Top             =   1800
         Width           =   1230
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
         Left            =   360
         TabIndex        =   14
         Top             =   2400
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
         Left            =   360
         TabIndex        =   13
         Top             =   3000
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Instalment :"
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
         TabIndex        =   12
         Top             =   3600
         Width           =   1860
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
         Left            =   6000
         TabIndex        =   11
         Top             =   3600
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
         Left            =   360
         TabIndex        =   10
         Top             =   4800
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
         Left            =   2400
         TabIndex        =   9
         Top             =   600
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
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
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
         Left            =   2400
         TabIndex        =   7
         Top             =   1800
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
         Left            =   8280
         TabIndex        =   6
         Top             =   3600
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
         Left            =   360
         TabIndex        =   5
         Top             =   4200
         Width           =   690
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Dim s(5) As String, i As Integer
On Error Resume Next
's(0) = ll.FindItem(Format(Text1.Text, "yyyy/mm/dd")).Index
'ss =
'sss = DateDiff("d", Text1.Text, )
If DateDiff("d", Text1.Text, ll.ListItems.Item(1)) > 0 Then
MsgBox "Invalid Opening Date.Please Change It.", vbCritical

Text1.SetFocus
Exit Sub
End If
s(1) = DateValue(Text1.Text)

Do Until s(0) <> ""
s(0) = ll.FindItem(Format(s(1), "yyyy/mm/dd")).Index
s(1) = DateAdd("d", 1, s(1))
Loop
Dim ii As Integer
i = s(0)
s(1) = DateAdd("d", -1, s(1))
If Text1.Text = s(1) Then
ii = DateDiff("d", Text1.Text, s(1))
Else
ii = DateDiff("d", Text1.Text, s(1)) - 1
End If
Do Until Val(Combo1.Text) <= ii
i = i + 1
If ll.ListItems.Count < i Then
MsgBox "Your Date Controller Has Been Expired.Please Update It Soon.Now Enter Holidays Manually."
Text2.Enabled = True
Text3.Text = ""
Text2.Text = 0
Exit Sub
End If

s(3) = ll.ListItems.Item(i).Text
's(3) = ll.FindItem(Format(s(2), "yyyy/mm/dd")).Index
s(2) = ll.ListItems.Item(i - 1).Text
ii = ii + DateDiff("d", s(2), s(3)) - 1
Loop
'If ii = Val(Combo1.Text) + 1 Then
's(4) = DateAdd("d", i - s(0), Text1.Text)

'Else
'Text2.Text = s(2)
's(4) = DateAdd("d", i - s(0), Text1.Text)
'End If

Text3.Text = DateAdd("d", Val(Combo1.Text) - ii - 1, s(3))

Text2.Text = DateDiff("d", Text1.Text, Text3.Text) - Val(Combo1.Text)

End Sub

Private Sub Form_Activate()
Combo1.Text = 30

End Sub

Private Sub Form_Load()
red
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

Private Sub Form_Unload(Cancel As Integer)
If PushButton1.Enabled = True Then
Cancel = 1


End If

End Sub

Private Sub PushButton1_Click()
Dim s As Currency
Dim ss As Currency
If Text3.Text = "" Then Text3.Text = DateAdd("d", Val(Combo1.Text) + Val(Text2.Text), Text1.Text)
s = (Val(Label11.Caption) * (Val(Text4.Text)) / 100) * Val(Combo1.Text) / 30
Label23.Caption = s
Label18.Caption = Format$(s / Val(Combo1.Text), "#.00")
Label19.Caption = Format$(Val(Label11.Caption) / Val(Combo1.Text), "#.00")
ss = Val(Label11.Caption) + s
Label12.Caption = Format$(ss / Val(Combo1.Text), "#.00")
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.CreateTextFile(App.Path & "\Data\Account log\" & Form3.Text1.Text & "\a.dat", True)
t.WriteLine ss
t.WriteLine ss
t.WriteLine Label9.Caption
t.WriteLine Label10.Caption
t.WriteLine Label11.Caption
t.WriteLine Text1.Text
t.WriteLine Text2.Text
t.WriteLine Text3.Text
t.WriteLine Text4.Text
t.WriteLine Combo1.Text
t.WriteLine Label12.Caption
t.WriteLine Text5.Text
t.Close
Set fso = Nothing
PushButton1.Enabled = False
'Unload Form3
MsgBox "Create Compelete ."
End Sub

