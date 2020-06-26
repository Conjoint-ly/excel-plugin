
Option Explicit
Sub CONJOINTLY_Centre_Across_Cells_eventhandler(control As IRibbonControl)
    CONJOINTLY_Centre_Across_Cells
End Sub
Sub CONJOINTLY_Centre_Across_Cells()
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .MergeCells = False
    End With
End Sub
Sub CONJOINTLY_Find_Ref_eventhandler(control As IRibbonControl)
    CONJOINTLY_Find_Ref
End Sub
Sub CONJOINTLY_Find_Ref()
    Dim wbSheet As Worksheet
    Dim wbFound As Range
    Dim wbaFound As Range
    Dim wkcount As Integer
    Set wbSheet = ActiveSheet

    With Application
       .DisplayAlerts = False
       .ScreenUpdating = False
    End With
    
    Set wbFound = Cells.Find(What:="#ref", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)
    If wbFound Is Nothing Then
    
        For wkcount = 1 To Worksheets.Count
            Worksheets(wkcount).Activate
            Set wbaFound = Cells.Find(What:="#ref", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)
            If wbaFound Is Nothing Then
            Else
                wbaFound.Activate
                GoTo EXITHerefind_ref
            End If
        Next
        wbSheet.Activate
        MsgBox "No #REF was found in values"
    Else
        wbFound.Activate
    End If
    
EXITHerefind_ref:
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub


Private Sub TurnOffGridLines()
    Dim view As WorksheetView
    For Each view In ActiveSheet.Parent.Windows(1).SheetViews
        If view.Sheet.Name = ActiveSheet.Name Then
            view.DisplayGridlines = False
            Exit Sub
        End If
    Next
End Sub

Sub CONJOINTLY_Make_Solid_Table_eventhandler(control As IRibbonControl)
    CONJOINTLY_Make_Solid_Table
End Sub
Sub CONJOINTLY_Make_Solid_Table()

    Dim FirstRowNumber As Integer
    
    TurnOffGridLines
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    
    FirstRowNumber = Selection.Columns(1).Column
    Range(Cells(Selection.Rows(1).Row, FirstRowNumber), Cells(Selection.Rows(1).Row, Selection.Columns.Count + FirstRowNumber - 1)).Select
    Selection.Font.Bold = True
End Sub


Sub CONJOINTLY_Kill_Custom_Styles_eventhandler(control As IRibbonControl)
    CONJOINTLY_Kill_Custom_Styles
End Sub
Sub CONJOINTLY_Kill_Custom_Styles()
     Dim styT As Style
     Dim intRet As Integer
     On Error Resume Next
     For Each styT In ActiveWorkbook.Styles
         If Not styT.BuiltIn Then
             If styT.Name <> "1" And styT.Name <> "Assumption" Then styT.Delete
         End If
     Next styT
 End Sub
 
Sub CONJOINTLY_Wrap_Formula_In_If_Error_eventhandler(control As IRibbonControl)
    CONJOINTLY_Wrap_Formula_In_If_Error
End Sub
Sub CONJOINTLY_Wrap_Formula_In_If_Error()
    Dim mycell As Range
    Dim display_if_true As Integer
    Dim error_condition As String
    Dim error_value As String
    Dim oldvalue As String
    Dim oldvalue1 As String
    Dim error_condition1 As String
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    oldvalue1 = Right(Selection.Cells(1, 1).Formula, Len(Selection.Cells(1, 1).Formula) - 1)
    If Len(oldvalue1) > 10 Then
        oldvalue1 = Left(oldvalue1, 10) & "..."
    End If
    error_condition = InputBox(Prompt:="What is the condition?" & vbNewLine & "=IF(" & oldvalue1 & " ??? , ... , ...)", Title:="Wrap formula in IF value", Default:="=0")
    error_condition1 = error_condition
    If Len(error_condition1) > 10 Then
        error_condition1 = Left(error_condition1, 10) & "..."
    End If
    display_if_true = MsgBox(Prompt:="Should the alternative value be displayed when the condition is satisfied?", Buttons:=vbYesNoCancel, Title:="Wrap formula in IF value")
    If display_if_true = 6 Then
        'true
        error_value = InputBox(Prompt:="What is the alternative value?" & vbNewLine & "=IF(" & oldvalue1 & error_condition1 & ", ???, " & oldvalue1 & " )", Title:="Wrap formula in IF value", Default:="""""")

        For Each mycell In Selection.Cells
            If mycell.HasFormula And Not mycell.HasArray Then
                oldvalue = Right(mycell.Formula, Len(mycell.Formula) - 1)
                mycell.Formula = "=IF(" & oldvalue & error_condition & "," & error_value & "," & oldvalue & ")"
            End If
        Next
    ElseIf display_if_true = 7 Then
        'false
        error_value = InputBox(Prompt:="What is the alternative value?" & vbNewLine & "=IF(" & oldvalue1 & error_condition1 & ", " & oldvalue1 & ", ??? )", Title:="Wrap formula in IF value", Default:="""""")

        For Each mycell In Selection.Cells
            If mycell.HasFormula And Not mycell.HasArray Then
                oldvalue = Right(mycell.Formula, Len(mycell.Formula) - 1)
                mycell.Formula = "=IF(" & oldvalue & error_condition & "," & oldvalue & "," & error_value & ")"
            End If
        Next
    
    
    End If
            
With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
End With
    
End Sub
Sub CONJOINTLY_Trace_Dependents_Outside_Range_eventhandler(control As IRibbonControl)
    CONJOINTLY_Trace_Dependents_Outside_Range
End Sub
Sub CONJOINTLY_Trace_Dependents_Outside_Range()
    On Error Resume Next
    Dim dpns As Range
    Dim fullset As Range
    Dim mycell As Range

    For Each mycell In Selection
        Set dpns = mycell.DirectDependents
        If dpns Then
            Set fullset = CONJOINTLY_Union(fullset, CONJOINTLY_SubstractRanges(dpns, Selection))
        End If
    Next
    If fullset Is Nothing Then
        MsgBox "No dependents for this range were found on this sheet."
    Else
        MsgBox "This range has the following dependents on this sheet: " & fullset.Address
        fullset.Select
    End If
End Sub
Sub CONJOINTLY_Trace_Precedents_Outside_Range_eventhandler(control As IRibbonControl)
    CONJOINTLY_Trace_Precedents_Outside_Range
End Sub
Sub CONJOINTLY_Trace_Precedents_Outside_Range()
    On Error Resume Next
    Dim dpns As Range
    Dim fullset As Range
    Dim mycell As Range

    For Each mycell In Selection
        Set dpns = mycell.DirectPrecedents
        If dpns Then
            Set fullset = CONJOINTLY_Union(fullset, CONJOINTLY_SubstractRanges(dpns, Selection))
        End If
    Next
    If fullset Is Nothing Then
        MsgBox "No precedents for this range were found on this sheet."
    Else
        MsgBox "This range has the following precedents on this sheet: " & fullset.Address
        fullset.Select
    End If
End Sub

Function CONJOINTLY_Union(Rng1 As Range, Rng2 As Range) As Range
    If Rng1 Is Nothing Then
        Set CONJOINTLY_Union = Rng2
    ElseIf Rng2 Is Nothing Then
        Set CONJOINTLY_Union = Rng1
    Else
        Set CONJOINTLY_Union = Application.Union(Rng1, Rng2)
    End If
End Function

Function CONJOINTLY_SubstractRanges(BigRangeToSubstractFrom As Range, SmallRangeToSubstract As Range) As Range
    Dim SubstractedRanges As Range
    Dim c As Range
    For Each c In BigRangeToSubstractFrom
        If Intersect(c, SmallRangeToSubstract) Is Nothing Then
            If SubstractedRanges Is Nothing Then
                Set SubstractedRanges = c
            Else
                Set SubstractedRanges = CONJOINTLY_Union(SubstractedRanges, c)
            End If
        End If
    Next c
    Set CONJOINTLY_SubstractRanges = SubstractedRanges
End Function
Sub CONJOINTLY_Red_All_Errors_eventhandler(control As IRibbonControl)
    CONJOINTLY_Red_All_Errors
End Sub
Sub CONJOINTLY_Red_All_Errors()
Cells.Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISERROR(A1)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
Sub CONJOINTLY_Format_As_Error_Check_eventhandler(control As IRibbonControl)
    CONJOINTLY_Format_As_Error_Check
End Sub
Sub CONJOINTLY_Format_As_Error_Check()

    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""N/A"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""NA"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & ActiveCell.Offset(0, 0).Address(False, False) & "=FALSE,TYPE(" & ActiveCell.Offset(0, 0).Address(False, False) & ")=4)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    With Selection
        .HorizontalAlignment = xlCenter
    End With
End Sub


Sub CONJOINTLY_CellColorsToChart_eventhandler(control As IRibbonControl)
    CONJOINTLY_CellColorsToChart
End Sub
Sub CONJOINTLY_CellColorsToChart()
    Dim xChart As Chart
    Dim I As Long, J As Long
    Dim xRowsOrCols As Long, xSCount As Long
    Dim xRg As Range, xCell As Range
    On Error Resume Next
    
    Set xChart = ActiveChart
    If xChart Is Nothing Then Exit Sub
    Application.StatusBar = "Starting for the chart called: " & xChart.Name
    
    xSCount = xChart.SeriesCollection.Count
    For I = 1 To xSCount
        J = 1
        With xChart.SeriesCollection(I)
            Set xRg = ActiveSheet.Range(Split(Split(.Formula, ",")(2), "!")(1))
            If xSCount > 4 Then
                xRowsOrCols = xRg.Columns.Count
            Else
                xRowsOrCols = xRg.Rows.Count
            End If
            
            .Format.Line.Visible = msoTrue
            .Format.Line.Weight = 1
            .Format.Line.ForeColor.RGB = RGB(100, 100, 100)
            
            For Each xCell In xRg
                Application.StatusBar = "Updating row " & I & " in column " & J
                .Points(J).Format.Fill.ForeColor.RGB = xCell.Interior.Color
                .Points(J).Format.Line.Visible = msoTrue
                .Points(J).Format.Line.Weight = 1
                .Points(J).Format.Line.ForeColor.RGB = RGB(100, 100, 100)
                J = J + 1
            Next
            
        End With
    Next
    Application.StatusBar = False
End Sub


Sub CONJOINTLY_Open_Experiments_eventhandler(control As IRibbonControl)
    CONJOINTLY_Open_Experiments
End Sub
Sub CONJOINTLY_Open_Experiments()
    ThisWorkbook.FollowHyperlink ("https://run.conjoint.ly/experiments")
End Sub


Sub CONJOINTLY_CreateLinksToAllSheets_eventhandler(control As IRibbonControl)
    CONJOINTLY_CreateLinksToAllSheets
End Sub
Sub CONJOINTLY_CreateLinksToAllSheets()
    Dim sh As Worksheet
    Dim cell As Range
    Dim view As WorksheetView
    
    If MsgBox("These changes cannot be undone. Proceed?", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    
    ActiveCell.Value = "Index"
    ActiveCell.Font.Bold = True
    ActiveCell.Font.Underline = True
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = ""
    ActiveCell.Offset(1, 0).Select
    
    For Each sh In ActiveWorkbook.Worksheets
        If ActiveSheet.Name <> sh.Name And sh.Visible = xlSheetVisible Then
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
            "'" & sh.Name & "'" & "!A1", TextToDisplay:=sh.Name
            ActiveCell.Offset(1, 0).Select
        End If
       
        For Each view In sh.Parent.Windows(1).SheetViews
            If view.Sheet.Name = sh.Name Then
                view.DisplayGridlines = False
            End If
        Next
    Next sh
    
    For Each sh In ActiveWorkbook.Worksheets
        With sh.Cells.Font
            .Name = "Helvetica Neue"
        End With
    Next sh
    
End Sub



