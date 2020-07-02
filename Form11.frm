VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form Form11 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _Version        =   786432
      _ExtentX        =   9975
      _ExtentY        =   4895
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.ShowTabs=   0   'False
      ItemCount       =   1
      Item(0).Caption =   ""
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   2715
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   5595
         _Version        =   786432
         _ExtentX        =   9869
         _ExtentY        =   4789
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   4080
            TabIndex        =   5
            Top             =   2040
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Apply"
            ForeColor       =   0
            Appearance      =   6
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            Text            =   "2012"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MM/DD/YYYY"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   6
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please Check Your System Date  ::-"
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   1560
            TabIndex        =   4
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "System Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   2
            Top             =   720
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const MAX_PATH = 260
Private Const CSIDL_FONTS = &H14
Private Function GetSpecialFolderPath(ByVal folder_number As Long) As String
Dim Path As String

    Path = Space$(MAX_PATH)
    If SHGetSpecialFolderPath(hwnd, Path, _
        folder_number, False) _
    Then
    Dim fso As New FileSystemObject

        GetSpecialFolderPath = Left$(Path, InStr(Path, Chr$(0)))
        Dim C As String
        C = Left(GetSpecialFolderPath, Len(GetSpecialFolderPath) - 1)
        On Error Resume Next
  fso.CopyFile App.Path & "\Data\Font\*.*", C, False

    End If
End Function

Private Sub Form_Activate()
  GetSpecialFolderPath CSIDL_FONTS
End Sub

Private Sub Form_Load()
Text1.Text = Date
End Sub

Private Sub PushButton1_Click()
On Error Resume Next

Date = DateValue(Text1.Text)
If Err.Number > 0 Then
Err.Clear
MsgBox "Invaild Date,Please Check Again.", vbCritical
Else
Form1.Show
Unload Form11

End If
End Sub

