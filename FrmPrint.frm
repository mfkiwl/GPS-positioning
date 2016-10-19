VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form FrmPrint 
   Caption         =   "水准成果表打印"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5130
   StartUpPosition =   3  '窗口缺省
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cboSelType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmPrint.frx":0000
      Left            =   2640
      List            =   "FrmPrint.frx":0007
      TabIndex        =   0
      Top             =   450
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "请选择打印内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSelect As String
Dim idx As Integer

Private Sub cboSelType_Click()
     StrSelect = cboSelType.Text
     idx = cboSelType.ListIndex
     If idx < 0 Then idx = 0
End Sub

Private Sub CmdExit_Click()
     Unload FrmPrint
End Sub

Private Sub CmdPrint_Click()
    Dim i As Integer
    Dim strtmp As String
    Dim ReportPath As String
    Dim arecord As Recordset
    If idx = 0 Then
       ReportPath = App.path + "\三角形闭合差表.rpt"
    End If
    CrystalReport1.ReportFileName = ReportPath
    CrystalReport1.DataFiles(0) = g_ProjectFile
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
'     Unload FrmPrint
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    cboSelType.Text = cboSelType.List(0)
    StrSelect = cboSelType.List(0)
End Sub
