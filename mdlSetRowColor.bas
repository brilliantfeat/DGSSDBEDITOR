Attribute VB_Name = "mdlvsFlexGrid"

Public Sub SetRowColor(ByRef MSHFlexGrid As Object)
    Dim J, I, objName
    objName = TypeName(MSHFlexGrid)

    If StrConv(Trim(objName), vbUpperCase) <> "VSFLEXGRID" Then
        Exit Sub
    End If

    MSHFlexGrid.FillStyle = 1

    For I = 1 To MSHFlexGrid.Rows - 1
        MSHFlexGrid.Row = I

        If I Mod 2 = 0 Then
            MSHFlexGrid.Col = 0
            MSHFlexGrid.ColSel = MSHFlexGrid.Cols - 1
            MSHFlexGrid.CellBackColor = &H80000018
        End If

    Next I
    
    For I = 1 To MSHFlexGrid.Rows - 1
        MSHFlexGrid.Row = I
        MSHFlexGrid.Col = 0
        MSHFlexGrid.CellBackColor = &H8000000F

    Next I

    MSHFlexGrid.FillStyle = 0
    MSHFlexGrid.Row = 0
    MSHFlexGrid.Col = 0
End Sub
Function ExportExcelclicp(FLex As VSFlexGrid)
    '------------------------------------------------
    '功能:将MSHFlexGrid表中内容复制至剪粘板
    '变量:
    '调用方法: call ExportExcelclicp(MSHFlexGrid1)
    '    [Scols]................MSHFlexGrid表格的启始列
    '    [Srows]............... MSHFlexGrid表格的启始行
    '    [Ecols]................MSHFlexGrid表格的结束列
    '    [Erows]............... MSHFlexGrid表格的结束行
    '------------------------------------------------
    Screen.MousePointer = 13
    '
    Dim Scols, Srows, Ecols, Erows As Integer

    With FLex
        Scols = .Col
        Srows = .Row
        Ecols = .ColSel
        Erows = .RowSel

        If .ColSel > .Col And .RowSel > .Row Then
            Scols = .Col
            Srows = .Row
            Ecols = .ColSel
            Erows = .RowSel
        ElseIf .ColSel < .Col And .RowSel < .Row Then
            Scols = .ColSel
            Srows = .RowSel
            Ecols = .Col
            Erows = .Row
        ElseIf .ColSel > .Col And .RowSel < .Row Then
            Scols = .Col
            Srows = .RowSel
            Ecols = .ColSel
            Erows = .Row
        ElseIf .ColSel < .Col And .RowSel > .Row Then
            Scols = .ColSel
            Srows = .Row
            Ecols = .Col
            Erows = .RowSel
        End If

    End With

    Dim I, J As Integer
    Dim str       As String
    Dim Fileopens As Boolean
    On Error GoTo err

    str = ""

    If Srows > 0 Then

        For I = Scols To Ecols '复制题头

            If I = Scols Then
                str = str & FLex.TextMatrix(0, I)
            Else
                str = str & Chr(9) & FLex.TextMatrix(0, I)
            End If

        Next

    End If

    For J = Srows To Erows
        str = str & vbCrLf

        For I = Scols To Ecols

            If I = Scols Then
                str = str & FLex.TextMatrix(J, I)
            Else
                str = str & Chr(9) & FLex.TextMatrix(J, I)
            End If

        Next
    Next

    Clipboard.Clear   ' 清除剪贴板。
    Clipboard.SetText str  ' 将正文放置在剪贴板上。
    Screen.MousePointer = 0

err:

    Select Case err.Number

        Case 0

        Case Else
            Screen.MousePointer = 0
            MsgBox err.Description, vbInformation, "出错"
            Exit Function
    End Select
  
End Function
