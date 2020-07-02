VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monthly Report"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11970
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Calculate"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      _Version        =   786432
      _ExtentX        =   21167
      _ExtentY        =   13996
      _StockProps     =   68
      Appearance      =   10
      Color           =   16
      PaintManager.ShowTabs=   0   'False
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6600
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6600
         Width           =   2535
      End
      Begin MSComctlLib.ListView l 
         Height          =   6015
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Daily Total Interest"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Daily Total Advance"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Daily Total Amount"
            Object.Width           =   4410
         EndProperty
      End
      Begin RichTextLib.RichTextBox r 
         Height          =   735
         Left            =   10680
         TabIndex        =   1
         Top             =   9240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         _Version        =   393217
         TextRTF         =   $"Form10.frx":5C12
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
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
         Left            =   4920
         TabIndex        =   12
         Top             =   6600
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   9600
         TabIndex        =   7
         Top             =   7320
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   6360
         TabIndex        =   6
         Top             =   7320
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   7320
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount : "
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
         Left            =   7800
         TabIndex        =   4
         Top             =   7320
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Advance : "
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
         Left            =   4320
         TabIndex        =   3
         Top             =   7320
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   2
         Top             =   7320
         Width           =   1290
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qq As Currency
Dim qqq As Currency
Dim qqqq As Currency
Dim gh As Boolean
Dim u As Integer
Dim uu As Integer
Private Sub ap(s As String)
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.OpenTextFile(s, ForReading, True)
Dim sss As String
Dim ss As String
ss = t.ReadLine
sss = t.ReadLine
r.Text = ""
r.Text = ss + vbCrLf
r.Text = r.Text + ss + vbCrLf
r.Text = r.Text + t.ReadAll
t.Close
Set t = fso.CreateTextFile(s, True)
t.Write r.Text
t.Close

End Sub
Private Sub ch(s As String)
On Error Resume Next
Dim fso As New FileSystemObject
Dim t As TextStream
Set t = fso.OpenTextFile(s, ForReading, True)
Dim x As String
Dim xx As String
Dim xxx As String
Dim u(20) As String
u(6) = t.ReadLine
u(7) = t.ReadLine
u(1) = Val(u(7)) - Val(u(6))
If Val(u(1)) = 0 Then Exit Sub
t.SkipLine
t.SkipLine
u(2) = t.ReadLine
t.SkipLine
t.SkipLine
t.SkipLine
u(3) = t.ReadLine
u(4) = t.ReadLine
u(5) = t.ReadLine
t.Close
u(8) = Left(Format$(Val(u(1)) / Val(u(5)), "000#.00"), 4)
u(9) = Val(u(8)) * Val(u(5))
u(10) = Val(u(1)) - Val(u(9))
u(11) = Val(u(2)) * Val(u(3)) / 100
u(12) = Format(Val(u(11)) / Val(u(4)), "#.00")

If Val(u(10)) > Val(u(12)) Then
u(13) = Val(u(10)) - Val(u(12))
qq = qq + Val(u(12)) * (Val(u(8)) + 1)
qqq = qqq + (Val(u(9)) - Val(u(12)) * (Val(u(8)) + 1)) + Val(u(10))
qqqq = qqqq + Val(u(1))
Else
qq = qq + Val(u(12)) * Val(u(8))
qqq = qqq + (Val(u(9)) - Val(u(12)) * Val(u(8))) + Val(u(10))
qqqq = qqqq + Val(u(1))

End If
'qqqq = 0

End Sub
Private Sub cl()
Dim fso As New FileSystemObject
Dim FS As Folder
Dim ff As Folder
Dim fff As Folder
Dim f As file
       On Error GoTo u
    Set FS = fso.GetFolder(App.Path & "\Data\Closed Account Log\")
   For Each ff In FS.SubFolders
        DoEvents
    Set fff = fso.GetFolder(ff.Path)
   For Each f In fff.Files
        DoEvents

ch (f.Path)
ap (f.Path)
 Next f


 Next ff
u:

    Set FS = Nothing

End Sub

Private Sub list()
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim file As Folder
       On Error GoTo u
    Set FSfolder = FS.GetFolder(App.Path & "\Data\Account log\")
   For Each file In FSfolder.SubFolders
        DoEvents
cal file.Path
u:

 Next file


    Set FSfolder = Nothing

End Sub

Private Sub cal(s As String)
Dim fso As New FileSystemObject
Dim t As TextStream
On Error Resume Next
Set t = fso.OpenTextFile(s & "\a.dat", ForReading)
Dim x As String
Dim xx As String
Dim xxx As String
Dim u(20) As String
u(6) = t.ReadLine
u(7) = t.ReadLine
u(1) = Val(u(7)) - Val(u(6))
If Val(u(1)) = 0 Then Exit Sub
t.SkipLine
t.SkipLine
u(2) = t.ReadLine
t.SkipLine
t.SkipLine
t.SkipLine
u(3) = t.ReadLine
u(4) = t.ReadLine
u(5) = t.ReadLine
t.Close
u(8) = Left(Format$(Val(u(1)) / Val(u(5)), "000#.00"), 4)
u(9) = Val(u(8)) * Val(u(5))
u(10) = Val(u(1)) - Val(u(9))
u(11) = Val(u(2)) * Val(u(3)) / 100
u(12) = Format(Val(u(11)) / Val(u(4)), "#.00")

If Val(u(10)) > Val(u(12)) Then
u(13) = Val(u(10)) - Val(u(12))
qq = qq + Val(u(12)) * (Val(u(8)) + 1)
qqq = qqq + (Val(u(9)) - Val(u(12)) * (Val(u(8)) + 1)) + Val(u(10))
qqqq = qqqq + Val(u(1))
Else
qq = qq + Val(u(12)) * Val(u(8))
qqq = qqq + (Val(u(9)) - Val(u(12)) * Val(u(8))) + Val(u(10))
qqqq = qqqq + Val(u(1))

End If
'qqqq = 0
End Sub

Private Sub Form_Activate()
Dim fso As New FileSystemObject
Dim t As TextStream
On Error GoTo j:
Set t = fso.OpenTextFile(App.Path & "\Data\Report.txt", ForReading, True)
Dim m As ListItem

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
list
apply
cl
Dim fso As New FileSystemObject
Dim t As TextStream
If qqqq = 0 Then Exit Sub

Set t = fso.OpenTextFile(App.Path & "\Data\Report.txt", ForAppending, True)
t.WriteLine Date
t.WriteLine Time
t.WriteLine qq
t.WriteLine qqq
t.WriteLine qqqq
t.Close
qq = 0
qqq = 0
qqqq = 0
End Sub
Private Sub apply()
list1
End Sub
Private Sub list1()
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim file As Folder
       On Error GoTo u
    Set FSfolder = FS.GetFolder(App.Path & "\Data\Account log\")
   For Each file In FSfolder.SubFolders
        DoEvents
cal1 file.Path

u:
 Next file


    Set FSfolder = Nothing

End Sub

Private Sub cal1(s As String)
Dim fso As New FileSystemObject
Dim t As TextStream
On Error GoTo l:
Set t = fso.OpenTextFile(s & "\a.dat", ForReading)
Dim sss As String
Dim ss As String
ss = t.ReadLine
sss = t.ReadLine
r.Text = ""
r.Text = ss + vbCrLf
r.Text = r.Text + ss + vbCrLf
r.Text = r.Text + t.ReadAll
t.Close
Set t = fso.CreateTextFile(s & "\a.dat", True)
t.Write r.Text
t.Close
l:
End Sub

Private Sub l_ItemClick(ByVal Item As MSComctlLib.ListItem)
If gh = False Then
Text1.Text = l.SelectedItem.Text & "   " & l.SelectedItem.SubItems(1)
u = l.SelectedItem.Index
Else
Text2.Text = l.SelectedItem.Text & "   " & l.SelectedItem.SubItems(1)
uu = l.SelectedItem.Index
End If
End Sub

Private Sub PushButton1_Click()
Dim kk As Currency
Dim kkk As Currency
Dim kkkk As Currency
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please Select Date ", vbCritical
Exit Sub
End If
On Error GoTo b:
Do Until u > uu

kk = kk + Val(l.ListItems.Item(u).SubItems(2))
kkk = kkk + Val(l.ListItems.Item(u).SubItems(3))
kkkk = kkkk + Val(l.ListItems.Item(u).SubItems(4))
u = u + 1

Loop

b:
If Err.Number = 13 Then MsgBox "Invalid Date", vbCritical
Label4.Caption = kk
Label5.Caption = kkk
Label6.Caption = kkkk
u = 0
uu = 0
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Text1_Click()
gh = False
End Sub

Private Sub Text2_Click()
gh = True
End Sub
