VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFM 
   Caption         =   "DGSS属性数据可视化编辑"
   ClientHeight    =   8850
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   15615
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1041
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   8595
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15120
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   7935
      Left            =   13080
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1140
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   2655
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11880
      _Version        =   393217
      Indentation     =   176
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid DG 
      Height          =   7935
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   10095
      _cx             =   17806
      _cy             =   13996
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   100
      RowHeightMax    =   3000
      ColWidthMin     =   450
      ColWidthMax     =   13500
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   2
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "回退"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "前进"
      Enabled         =   0   'False
      Height          =   315
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存到数据库"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   90
      Width           =   1455
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Begin VB.Menu SaveDB 
         Caption         =   "保存（数据库）"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "另存为"
         Begin VB.Menu SaveAsExcel 
            Caption         =   "Excel文件"
         End
      End
      Begin VB.Menu LoadFrom 
         Caption         =   "导入"
         Begin VB.Menu LoadFromExcel 
            Caption         =   "从Excel文件(数据库不可用）"
         End
      End
   End
   Begin VB.Menu Setpath 
      Caption         =   "设置"
      Begin VB.Menu SetDGSSDBPath 
         Caption         =   "设置DGSS路径"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "工具"
      Begin VB.Menu FindAndReplce 
         Caption         =   "DGSS属性数据库查找替换"
      End
   End
   Begin VB.Menu VSFGMenu 
      Caption         =   "VSFG菜单"
      Visible         =   0   'False
      Begin VB.Menu Copy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu Fangda 
         Caption         =   "放大"
      End
      Begin VB.Menu Suoxiao 
         Caption         =   "缩小"
      End
   End
End
Attribute VB_Name = "MainFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const InitialRecordCount& = 5000, UndoDepth& = 10

Private DBName$, Cnn As cConnection, WithEvents CurRs As cRecordset
Attribute CurRs.VB_VarHelpID = -1
Private HistoryList As cHistoryList, TopRow As Long, LeftCol As Long
Private UnComprSize As Long, ComprSize As Long

Private Sub Copy_Click()
Call ExportExcelclicp(DG)
End Sub

Private Sub DG_Click()
With DG
Text1.Text = DG.TextMatrix(.Row, .Col)
End With
End Sub


Private Sub DG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu VSFGMenu
End Sub

Private Sub Fangda_Click()
DG.Height = DG.Height * 5 / 4
DG.FontSize = DG.FontSize + 2
DG.Width = DG.Width * 5 / 4
Form_Resize
End Sub

Private Sub FindAndReplce_Click()
Form2.Show
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Dim tempStr, a As String
    Open App.Path & "\path.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, a
        tempStr = tempStr & a
    Loop
    Close #1
    If tempStr <> "" Then
    Call fileconnection(tempStr, TreeView1, False)
    Else
    MsgBox ("先设置数据库目录!")
    End If
    'Call SetRowColor(Me.DG)
End Sub

Private Sub InsertDemoData()
Dim i&, J&, FldArr(1 To 49) As String
Dim Cmd As cCommand, DblDate As Double, RowTxtPart As String

  'Create a CommandObject from an Insert-Statement for faster Inserts
  For J = 1 To UBound(FldArr): FldArr(J) = "?": Next J
  Set Cmd = Cnn.CreateCommand("Insert Into Test Values(" & Join(FldArr, ",") & ")")
  
  'now the inserts (to stress the State-CompressionFeature a bit, we ensure
  'that each of the inserted FieldValues is unique in this initial TableContent
  DblDate = Now
  Cnn.BeginTrans
    Cmd.SetNull 1 'for our AutoID-Field
    For i = 1 To InitialRecordCount 'create a bunch of records
      RowTxtPart = ", Row-Index " & i
      For J = 1 To 16
        Cmd.SetInt32 3 * J - 1, J * 100000 + i
        Cmd.SetText 3 * J, "Text " & J & RowTxtPart
        Cmd.SetDate 3 * J + 1, DblDate + J - 1 + (i / 86400)
      Next J
      Cmd.Execute
    Next i
  Cnn.CommitTrans
End Sub

'Get a fresh Recordset from the DB and reset our HistoryList,
'followed by our first initial State-Saving (respecting the last Scrollposition)
Private Sub UpdateViewFromDB()
  'On Error Resume Next
  Set CurRs = Cnn.OpenRecordset("Select * From " & List1.Text)
  Set DG.DataSource = CurRs.DataSource
  DG.Refresh
  'DG.Columns(0).Visible = False
  HistoryList.Clear
  StoreState TopRow, StateReason.DBRead
  'Call SetRowColor(Me.DG)
End Sub

'Implementation of the Change-Reactions (coming from the Recordset-Events)
Private Sub CurRs_FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
  StoreState RowIdxZeroBased, StateReason.Update
End Sub
Private Sub CurRs_AddNew(ByVal NewRowIdxZeroBased As Long)
  StoreState NewRowIdxZeroBased, StateReason.AddNew
End Sub
Private Sub CurRs_Delete(ByVal NewRowIdxZeroBased As Long)
  StoreState NewRowIdxZeroBased, StateReason.Delete
End Sub

'Implementation of the State-Handling (save/restore from HistoryList)
Private Sub StoreState(ByVal CurRow As Long, Reason As StateReason)
Dim NewState As cState, B() As Byte
  New_c.Timing True
  
    'LeftCol = DG.LeftCol - 1 'reflect the HScroll-Value of the DataGrid
    'If DG.Row < 0 Then
    '  TopRow = CurRow
    'Else 'somewhat "weird construction", since the DataGrid has no real VScroll-Reflection
    '  TopRow = CurRow - CLng((DG.RowTop(DG.Row) - 13.984) / (DG.RowHeight + 1))
    'End If
    
    Set NewState = New cState
    B = CurRs.Content
    UnComprSize = UBound(B) + 1
    ComprSize = NewState.SaveContent(B, TopRow, LeftCol, Reason)
    HistoryList.SaveState NewState
    UpdateUndoRedoButtons
  
  'only a Timing-output
  'Select Case Reason
   ' Case StateReason.DBRead: lblTiming = New_c.Timing & " Initial-State set after DBRead"
   ' Case StateReason.Update: lblTiming = New_c.Timing & " State saved after Fld-Update"
   ' Case StateReason.AddNew: lblTiming = New_c.Timing & " State saved after AddNew"
   ' Case StateReason.Delete: lblTiming = New_c.Timing & " State saved after Delete"
  'End Select
  'lblTiming = lblTiming & " (compressed Size=" & CLng(ComprSize / 1024) & "kB, uncompressed=" & CLng(UnComprSize / 1024) & "kB)"
End Sub
Private Function ReStoreState(State As cState)
Dim B() As Byte, Reason As StateReason
  New_c.Timing True
  
    State.GetContent B, TopRow, LeftCol, Reason
    Set DG.DataSource = Nothing
    Set CurRs = New_c.Recordset
    CurRs.Content = B
    Set CurRs.ActiveConnection = Cnn
    Set DG.DataSource = CurRs.DataSource
    'DG.Columns(0).Visible = False
    'DG.Scroll LeftCol, TopRow 'restore the last Scroll-Position
    UpdateUndoRedoButtons
    
  'only a Timing-output
  'Select Case Reason
    'Case StateReason.DBRead: lblTiming = New_c.Timing & " Initial-State restored (DBRead)"
    'Case StateReason.Update: lblTiming = New_c.Timing & " State restored (Fld-Update)"
    'Case StateReason.AddNew: lblTiming = New_c.Timing & " State restored (AddNew)"
    'Case StateReason.Delete: lblTiming = New_c.Timing & " State restored (Delete)"
  'End Select
End Function

'Implementation of the Undo/Redo-Button-Events
Private Sub cmdRedo_Click()
Dim NextState As cState
  Set NextState = HistoryList.NextState
  If Not NextState Is Nothing Then ReStoreState NextState
End Sub
Private Sub cmdUndo_Click()
Dim PreviousState As cState
  Set PreviousState = HistoryList.PreviousState
  If Not PreviousState Is Nothing Then ReStoreState PreviousState
End Sub
Private Sub UpdateUndoRedoButtons()
  cmdRedo.Enabled = HistoryList.RedoEnabled
  cmdUndo.Enabled = HistoryList.UndoEnabled
  cmdSave.Enabled = cmdUndo.Enabled 'if no Undo possible, then we don't need to save to the DB
  SaveDB.Enabled = cmdSave.Enabled
End Sub
'and the appropriate DB-Save-Button
Private Sub cmdSave_Click()
  CurRs.UpdateBatch
  UpdateViewFromDB
End Sub

Private Sub Form_Resize()
'lblTiming.Width = ScaleWidth
On Error Resume Next
DG.Width = Me.ScaleWidth - 367
DG.Height = Me.ScaleHeight - 64
TreeView1.Height = Me.ScaleHeight - 144
List1.Top = Me.ScaleHeight - 97
Text1.Left = Me.ScaleWidth - 168
Text1.Height = Me.ScaleHeight - 64
End Sub


Private Sub List1_DblClick()
    On Error Resume Next
    Set HistoryList = New cHistoryList
    HistoryList.UndoDepth = UndoDepth
    Set Cnn = New_c.connection
    DBName = TreeView1.SelectedItem.Key
    If New_c.fso.FileExists(DBName) Then
    Cnn.OpenDB DBName
    End If
    UpdateViewFromDB
End Sub


Private Sub LoadFromExcel_Click()
CommonDialog1.Filter = "Excel文件|*.xls"
CommonDialog1.ShowOpen
If CommonDialog1.fileName <> "" Then DG.LoadGrid CommonDialog1.fileName, flexFileExcel
End Sub

Private Sub Paste_Click()
DG.Paste
End Sub

Private Sub SaveAsExcel_Click()
CommonDialog1.Filter = "Excel文件|*.xls"
CommonDialog1.ShowSave
If CommonDialog1.fileName <> "" Then DG.SaveGrid CommonDialog1.fileName, flexFileExcel, flexXLSaveFixedCells
End Sub

Private Sub SaveDB_Click()
    CurRs.UpdateBatch
    UpdateViewFromDB
End Sub

Private Sub SetDGSSDBPath_Click()
    Dim Pathstr As String
    Pathstr = GetFolder(Me.hWnd, "选择DGSS数据库目录")
    If Pathstr <> "" Then
        Open App.Path & "\path.txt" For Output As #1
        Print #1, Pathstr
        Close #1
        Call fileconnection(Pathstr, TreeView1, False)
    Else
        MsgBox ("请选择正确目录!")
    End If
End Sub
Private Sub Suoxiao_Click()
DG.Height = DG.Height * 4 / 5
DG.FontSize = DG.FontSize - 2
DG.Width = DG.Width * 4 / 5
Form_Resize
End Sub

Private Sub Text1_Change()
With DG
DG.TextMatrix(.Row, .Col) = Text1.Text
End With
End Sub

Private Sub TreeView1_DblClick()
   Dim Tbl As cTable, B() As Byte, CnnSrc As cConnection, RsSrc As cRecordset, T!
  'open a filebased Src-Database (the one we want to dump into memory)
  DBName = TreeView1.SelectedItem.Key
  Set CnnSrc = New_c.connection
    If New_c.fso.FileExists(DBName) Then
        New_c.Timing True
        List1.Clear
        CnnSrc.OpenDB DBName
        For Each Tbl In CnnSrc.DataBases(1).Tables
            List1.AddItem Tbl.Name
        Next Tbl
    End If
End Sub
