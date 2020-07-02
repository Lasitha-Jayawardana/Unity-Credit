VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form9 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form9"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   10470
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _Version        =   786432
      _ExtentX        =   18441
      _ExtentY        =   10610
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   9240
         TabIndex        =   4
         Top             =   5520
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Print Page"
         Appearance      =   6
      End
      Begin RichTextLib.RichTextBox r 
         Height          =   4815
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Form9.frx":5C12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "   No                     Date                 Instalment                 Balance                   User ID"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   9870
      End
      Begin VB.Label Label2 
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
         TabIndex        =   1
         Top             =   0
         Width           =   1965
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
r.Text = ""
Dim i As Integer
With Form8.l.ListItems
'r.SelAlignment = 2
'FormatNumber
'FormatPercent
'Format .Item(1).Text, "          "
'Format$
'r.Text= "   "+"No"+
Do Until .Count = i
i = i + 1
r.Text = r.Text + "   " + Format(.Item(i).Text, "#000")
r.Text = r.Text + "                                    " + .Item(i).SubItems(1)
r.Text = r.Text + "                                    " + .Item(i).SubItems(2)
r.Text = r.Text + "                                    " + .Item(i).SubItems(3)
r.Text = r.Text + "                                    " + .Item(i).SubItems(4) + vbCrLf

Loop
End With
End Sub

Private Sub Form_Resize()
r.Height = Me.Height - 1800
Me.Width = 10710
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
Dim s As String
s = MsgBox("Are You Sure ?", vbYesNo)
If s = vbYes Then
PushButton1.Visible = False
Me.PrintForm
End If

End Sub


