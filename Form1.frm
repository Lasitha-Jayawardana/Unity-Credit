VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   7365
   ClientTop       =   2310
   ClientWidth     =   6015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox r 
      Height          =   855
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":5C12
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2535
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   4471
      _StockProps     =   68
      Appearance      =   10
      Color           =   64
      PaintManager.ShowTabs=   0   'False
      Begin XtremeSuiteControls.PushButton PushButton2 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   2040
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cancel"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   1560
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "king"
         PasswordChar    =   "*"
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         Top             =   1080
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "lasith"
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   2040
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Log in"
         Appearance      =   6
         MultiLine       =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   1920
         Left            =   4440
         Picture         =   "Form1.frx":5C94
         Top             =   480
         Width           =   1920
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form1.frx":793F
         Top             =   0
         Width           =   6000
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1035
         _Version        =   786432
         _ExtentX        =   1826
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Password : "
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "User Name : "
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qq As Currency
Dim qqq As Currency
Dim qqqq As Currency


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
Private Sub ap(s As String)
Dim fso As New FileSystemObject
Dim t As TextStream
On Error GoTo L:
Set t = fso.OpenTextFile(s, ForReading, True)
Dim sss As String
Dim ss As String
ss = t.ReadLine
sss = t.ReadLine
R.Text = ""
R.Text = ss + vbCrLf
R.Text = R.Text + ss + vbCrLf
R.Text = R.Text + t.ReadAll
t.Close
Set t = fso.CreateTextFile(s, True)
t.Write R.Text
t.Close
L:
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

Private Sub cal1(s As String)
Dim fso As New FileSystemObject
Dim t As TextStream
On Error GoTo L:
Set t = fso.OpenTextFile(s & "\a.dat", ForReading)
Dim sss As String
Dim ss As String
ss = t.ReadLine
sss = t.ReadLine
R.Text = ""
R.Text = ss + vbCrLf
R.Text = R.Text + ss + vbCrLf
R.Text = R.Text + t.ReadAll
t.Close
Set t = fso.CreateTextFile(s & "\a.dat", True)
t.Write R.Text
t.Close
L:
End Sub




Private Sub Form_Activate()
list
list1
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

End Sub


Private Sub Form_Load()
If App.PrevInstance Then End
'MsgBox "Please Check Your Computer Date &  Time." & vbCrLf & vbCrLf & Date & vbCrLf & Time, vbSystemModal
End Sub

Private Sub PushButton1_Click()
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
Form2.Label3.Caption = "MDU"
Form2.Show

Unload Me
Exit Sub
ElseIf FlatEdit1.Text = s1 And FlatEdit2.Text = ss1 Then
Form2.Label3.Caption = "ND"
Form2.Show
Unload Me
Exit Sub
End If

If fso.FileExists(App.Path & "\Data\Users\" & FlatEdit1.Text) Then
Set t = fso.OpenTextFile(App.Path & "\Data\Users\" & FlatEdit1.Text)
If FlatEdit2.Text = t.ReadLine Then
Form2.Label3.Caption = FlatEdit1.Text
t.Close
Form2.Show
Unload Me
Else
GoTo f:
End If
Else
f:
MsgBox "Invalide User Name or Password!", vbCritical
End If

End Sub

Private Sub PushButton2_Click()
Dim s As String
s = MsgBox("Are You Sure Want Exit?", vbYesNo, "Unity Micro Credit")
If s = vbYes Then
Unload Me
End If
End Sub

