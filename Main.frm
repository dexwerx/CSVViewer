VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   Caption         =   "CSV Viewer"
   ClientHeight    =   4425
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   30
      Cols            =   10
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      GridLinesFixed  =   1
      GridLinesUnpopulated=   2
      AllowUserResizing=   1
      BorderStyle     =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileDiv1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Data As New ADODB.Recordset

Function Min(ByVal A As Long, ByVal B As Long) As Long
    Min = IIf(A < B, A, B)
End Function

Function Max(ByVal A As Long, ByVal B As Long) As Long
    Max = IIf(A > B, A, B)
End Function

Private Sub ResizeColumns()
    Dim i As Long
    Dim MinWidth As Long
    MinWidth = TextWidth("WWWWW")

    For i = 1 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.ColWidth(i) = Max(MinWidth, TextWidth(MSHFlexGrid1.TextMatrix(0, i)))
    Next

    MSHFlexGrid1.ColWidth(0) = MSHFlexGrid1.RowHeight(0) * 1.2
End Sub
Private Sub Form_Load()
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\;" & _
        "Extended Properties=""text;HDR=Yes;FMT=delimited;"""
    Data.Open "SELECT top 1000 * FROM [test.csv]", con, adOpenDynamic, adLockReadOnly
    Set MSHFlexGrid1.DataSource = Data
    ResizeColumns
End Sub

Private Sub Form_Resize()
    MSHFlexGrid1.Width = ScaleWidth
    MSHFlexGrid1.Height = ScaleHeight
End Sub

Private Sub mnuFileExit_Click()
    Dim l As Long
    Dim mmi As MINMAXINFO
    
    l = WM_GETMINMAXINFO
    End
End Sub

