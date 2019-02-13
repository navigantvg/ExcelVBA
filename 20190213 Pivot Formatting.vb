''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''Table of Contents'''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'SegmentFormatting
''Used for Themes Coverage Chart formatting

'ThemeExpandFormatting
''Used for Rule Applicability Chart formatting

'SegExpandFormatting
''Used for Product Groups by Segment, Themes By Segment formatting

'HeatMapCopyPaste
''Used for Themes Coverage Chart copy paste

'RuleChartCopyPaste
''Used for Rule Applicability Chart copy paste

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SegmentFormatting()

'
' Formatting Macro
' Refresh and Format
'

Dim Filter As Integer
Dim ColOffSet As Integer
Dim SegRow As Integer
Dim NumSegments As Integer
Dim FilterNum As Integer
Dim ThemeCol As Integer
Dim RuleCol As Integer
Dim FinalRow As Integer
Dim CoverageCells As Range
Dim Cond1 As FormatCondition
Dim Cond2 As FormatCondition
Dim Cond3 As FormatCondition

    ColOffSet = 1
    Filter = 6
    SegRow = 0
    FilterNum = Filter + SegRow
    RuleCol = Cells(FilterNum + 3, ColOffSet + 1).End(xlToRight).End(xlToRight).End(xlToLeft).Column
    NumSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, RuleCol)).Cells.SpecialCells(xlCellTypeConstants).Count
    With Range(Cells(FilterNum + 3, ColOffSet + 1), Cells(FilterNum + 3, RuleCol))
        'ScotiaRed
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 192
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .VerticalAlignment = xlBottom
        ''Original Blue
        '.Interior.Pattern = xlSolid
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.ThemeColor = xlThemeColorAccent5
        '.Interior.TintAndShade = -0.249977111117893
        '.Interior.PatternTintAndShade = 0
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 1
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .Orientation = 90
        .ColumnWidth = 3.75
        .RowHeight = 150
    End With
    
    Dim RunningSegments As Integer
    Dim PastSegments As Integer
    Dim NoneSegments As Integer
    
    For j = ColOffSet + 1 To RuleCol
        If Cells(FilterNum + 2, j).Value2 = "None" Then
            NoneSegments = j
            Exit For
        End If
    Next j
        
    If NoneSegments = 0 Then
        NoneSegments = RuleCol
    End If
    
    For i = ColOffSet + 1 To NoneSegments
        If i = ColOffSet + 1 Then
            RunningSegments = 1
            PastSegments = 0
        Else
            RunningSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, i)).Cells.SpecialCells(xlCellTypeConstants).Count
            PastSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, i - 1)).Cells.SpecialCells(xlCellTypeConstants).Count
        End If
        
        
        If RunningSegments Mod 2 = 1 Then
            GoTo NextLoop
        ElseIf RunningSegments = PastSegments Then
            Cells(FilterNum + 3, i - 1).Copy
            Cells(FilterNum + 3, i).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        ' Else
            ' With Cells(FilterNum + 3, i)
                ' .Interior.Pattern = xlSolid
                ' .Interior.PatternColorIndex = xlAutomatic
                ' .Interior.ThemeColor = xlThemeColorAccent2
                ' .Interior.TintAndShade = 0
                ' .Interior.PatternTintAndShade = 0
                ' .Borders.LineStyle = xlContinuous
                ' .Borders.ThemeColor = 1
                ' .Borders.TintAndShade = 0
                ' .Borders.Weight = xlThin
            ' End With
        End If
NextLoop:
    Next i
    
    If NoneSegments = RuleCol Then
        GoTo SkipColor
    End If
    
    With Range(Cells(FilterNum + 3, NoneSegments), Cells(FilterNum + 3, RuleCol))
        '.Interior.Pattern = xlSolid
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.ThemeColor = xlThemeColorDark1
        '.Interior.TintAndShade = -0.149998474074526
        '.Interior.PatternTintAndShade = 0
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent5
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
        .Font.ColorIndex = xlAutomatic
        .Font.TintAndShade = 0
    End With

SkipColor:
    FinalRow = Cells(FilterNum + 4, ColOffSet).End(xlDown).Row
    'Range(Cells(FilterNum + 4, ColOffSet+ 1), Cells(FinalRow, RuleCol)).Select
    With Range(Cells(FilterNum + 4, ColOffSet), Cells(FilterNum + 4, ColOffSet).End(xlDown))
        .HorizontalAlignment = xlLeft
    End With
    Set CoverageCells = Range(Cells(FilterNum + 4, ColOffSet + 1), Cells(FinalRow, RuleCol))
    Cells.FormatConditions.Delete
    
    'Setting up Formatting Conditions
    Set Cond1 = CoverageCells.FormatConditions.Add(xlCellValue, xlEqual, "=1")
    Set Cond2 = CoverageCells.FormatConditions.Add(xlCellValue, xlEqual, "=3")
    Set Cond3 = CoverageCells.FormatConditions.Add(xlCellValue, xlBetween, "=1", "=3")

    With CoverageCells.Borders
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlHairline
    End With

    'Full Coverage
    With Cond1
        .Font.ThemeColor = xlThemeColorAccent6
        .Font.TintAndShade = 0
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .StopIfTrue = True
    End With
    
    'No Coverage
    With Cond2
        '.Font.Strikethrough = False
        '.Font.Color = -16777024
        '.Font.TintAndShade = 0
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.Color = 192
        '.Interior.TintAndShade = 0
        .Font.ThemeColor = xlThemeColorAccent4
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 49407
        .StopIfTrue = True
    End With
    
    'Partial Coverage
    With Cond3
        '.Font.ThemeColor = xlThemeColorAccent4
        '.Font.TintAndShade = 0
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.Color = 49407
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.599993896298105
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorAccent6
        .Font.TintAndShade = 0.599993896298105
        .StopIfTrue = True
    End With
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ThemeExpandFormatting()

'
' Formatting Macro
' Refresh and Format
'

Dim Filter As Integer
Dim ColOffSet As Integer
Dim SegRow As Integer
Dim NumSegments As Integer
Dim FilterNum As Integer
Dim ThemeCol As Integer
Dim RuleCol As Integer
Dim FinalRow As Integer
Dim CoverageCells As Range
Dim Cond1 As IconSetCondition
Dim Cond2 As IconSetCondition
Dim Cond3 As IconSetCondition

    Filter = 6
    ColOffSet = 2
    SegRow = 0
    FilterNum = Filter + SegRow
    RuleCol = Cells(FilterNum + 3, ColOffSet + 1).End(xlToRight).End(xlToRight).End(xlToLeft).Column
    NumSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, RuleCol)).Cells.SpecialCells(xlCellTypeConstants).Count
    With Range(Cells(FilterNum + 3, ColOffSet + 1), Cells(FilterNum + 3, RuleCol))
        'ScotiaRed
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 192
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .VerticalAlignment = xlBottom
        ''Original Blue
        '.Interior.Pattern = xlSolid
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.ThemeColor = xlThemeColorAccent5
        '.Interior.TintAndShade = -0.249977111117893
        '.Interior.PatternTintAndShade = 0
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 1
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .Orientation = 90
        .ColumnWidth = 3.75
        .RowHeight = 150
    End With
    
    Dim RunningSegments As Integer
    Dim PastSegments As Integer
    
        
    
    For i = 2 To NoneSegments
        If i = 2 Then
            RunningSegments = 1
            PastSegments = 0
        Else
            RunningSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, ColOffSet + i - 1)).Cells.SpecialCells(xlCellTypeConstants).Count
            PastSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, ColOffSet + i - 2)).Cells.SpecialCells(xlCellTypeConstants).Count
        End If
        
        
        If RunningSegments Mod 2 = 1 Then
            GoTo NextLoop
        ElseIf RunningSegments = PastSegments Then
            Cells(FilterNum + 3, ColOffSet + i - 2).Copy
            Cells(FilterNum + 3, ColOffSet + i - 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Else
            With Cells(FilterNum + 3, ColOffSet + i - 1)
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorAccent2
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                .Borders.LineStyle = xlContinuous
                .Borders.ThemeColor = 1
                .Borders.TintAndShade = 0
                .Borders.Weight = xlThin
            End With
        End If
NextLoop:
    Next i
    
    FinalRow = Cells(FilterNum + 4, ColOffSet).End(xlDown).Row
    'Range(Cells(FilterNum + 4, ColOffSet + 1), Cells(FinalRow, RuleCol)).Select
        With Range(Cells(FilterNum + 4, ColOffSet), Cells(FilterNum + 4, ColOffSet).End(xlDown))
        .HorizontalAlignment = xlLeft
    End With
    Set CoverageCells = Range(Cells(FilterNum + 4, ColOffSet + 1), Cells(FinalRow, RuleCol))
    Cells.FormatConditions.Delete
    
    
    Set Cond1 = CoverageCells.FormatConditions.AddIconSetCondition
    
    'Setting up Formatting Conditions
    With Cond1
        .ReverseOrder = True
        .ShowIconOnly = True
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
        
        With .IconCriteria(1)
            .Icon = xlIconBlackCircleWithBorder
            '.Type = xlConditionValueFormula
            '.Value = 0
            '.Operator = 7
        End With
        
        With .IconCriteria(2)
            .Icon = xlIconBlackCircleWithBorder
            .Type = xlConditionValueNumber
            .Value = 1
            .Operator = xlGreaterEqual
        End With
        
        With .IconCriteria(3)
            .Icon = xlIconNoCellIcon
            .Type = xlConditionValueNumber
            .Value = 3
            .Operator = xlGreaterEqual
        End With

    End With
    

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SegExpandFormatting()

'
' Formatting Macro
' Refresh and Format
'

Dim Filter As Integer
Dim ColOffSet As Integer
Dim SegRow As Integer
Dim NumSegments As Integer
Dim FilterNum As Integer
Dim ThemeCol As Integer
Dim RuleCol As Integer
Dim FinalRow As Integer
Dim CoverageCells As Range
Dim Cond1 As IconSetCondition
Dim Cond2 As IconSetCondition
Dim Cond3 As IconSetCondition

    Filter = 1
    ColOffSet = 5
    SegRow = 1
    FilterNum = Filter + SegRow
    RuleCol = Cells(FilterNum + 3, ColOffSet + 1).End(xlToRight).End(xlToRight).End(xlToLeft).Column
    NumSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, RuleCol)).Cells.SpecialCells(xlCellTypeConstants).Count
    With Range(Cells(FilterNum + 3, ColOffSet + 1), Cells(FilterNum + 3, RuleCol))
        'ScotiaRed
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 192
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .VerticalAlignment = xlBottom
        ''Original Blue
        '.Interior.Pattern = xlSolid
        '.Interior.PatternColorIndex = xlAutomatic
        '.Interior.ThemeColor = xlThemeColorAccent5
        '.Interior.TintAndShade = -0.249977111117893
        '.Interior.PatternTintAndShade = 0
        .Borders.LineStyle = xlContinuous
        .Borders.ThemeColor = 1
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .Orientation = 90
        .ColumnWidth = 9
        .RowHeight = 150
    End With
    
    Dim RunningSegments As Integer
    Dim PastSegments As Integer
    
        
    
    For i = 2 To NoneSegments
        If i = 2 Then
            RunningSegments = 1
            PastSegments = 0
        Else
            RunningSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, ColOffSet + i - 1)).Cells.SpecialCells(xlCellTypeConstants).Count
            PastSegments = Range(Cells(FilterNum + 2, ColOffSet + 1), Cells(FilterNum + 2, ColOffSet + i - 2)).Cells.SpecialCells(xlCellTypeConstants).Count
        End If
        
        
        If RunningSegments Mod 2 = 1 Then
            GoTo NextLoop
        ElseIf RunningSegments = PastSegments Then
            Cells(FilterNum + 3, ColOffSet + i - 2).Copy
            Cells(FilterNum + 3, ColOffSet + i - 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Else
            With Cells(FilterNum + 3, ColOffSet + i - 1)
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorAccent2
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                .Borders.LineStyle = xlContinuous
                .Borders.ThemeColor = 1
                .Borders.TintAndShade = 0
                .Borders.Weight = xlThin
            End With
        End If
NextLoop:
    Next i
    
    FinalRow = Cells(FilterNum + 4, ColOffSet).End(xlDown).Row
    'Range(Cells(FilterNum + 4, ColOffSet + 1), Cells(FinalRow, RuleCol)).Select
    With Range(Cells(FilterNum + 4, ColOffSet), Cells(FilterNum + 4, ColOffSet).End(xlDown))
        .HorizontalAlignment = xlLeft
    End With
    Set CoverageCells = Range(Cells(FilterNum + 4, ColOffSet + 1), Cells(FinalRow, RuleCol))
    Cells.FormatConditions.Delete
    
    
    Set Cond1 = CoverageCells.FormatConditions.AddIconSetCondition
    
    'Setting up Formatting Conditions
    With Cond1
        .ReverseOrder = True
        .ShowIconOnly = True
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
        
        With .IconCriteria(1)
            .Icon = xlIconBlackCircleWithBorder
            '.Type = xlConditionValueFormula
            '.Value = 0
            '.Operator = 7
        End With
        
        With .IconCriteria(2)
            .Icon = xlIconBlackCircleWithBorder
            .Type = xlConditionValueNumber
            .Value = 1
            .Operator = xlGreaterEqual
        End With
        
        With .IconCriteria(3)
            .Icon = xlIconNoCellIcon
            .Type = xlConditionValueNumber
            .Value = 3
            .Operator = xlGreaterEqual
        End With

    End With
    
    With Range(Cells(FilterNum + 4, 1), Cells(FinalRow, RuleCol))
        .Borders.LineStyle = xlContinuous
        .Borders.ColorIndex = xlAutomatic
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        .Borders(xlInsideHorizontal).TintAndShade = 0
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).ColorIndex = xlAutomatic
        .Borders(xlInsideVertical).TintAndShade = 0
        .Borders(xlInsideVertical).Weight = xlHairline
        .Font.ColorIndex = xlAutomatic
        .Font.TintAndShade = 0
    End With
    
End Sub

Sub CopyToWord()

    Dim objWord As New Word.Application
    'Copy the range Which you want to paste in a New Word Document
    Range("A1:C18").Copy

    With objWord
        .Documents.Add
        .Selection.Paste
        .Visible = True
    End With

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub HeatMapCopyPaste()
'
' HeatMapCopyPaste Macro
'

Dim ThemeHeatMap As PivotTable
Dim ConduciveField As PivotField
Dim ApplicableField As PivotField
Dim DuplicateField As PivotField
Dim PriorityField As PivotField
Dim SegmentField As PivotField
Dim CoverageField As PivotField

Dim Conducive As String
Dim Applicable As String
Dim Duplicate As String
Dim Priority As String
Dim Segment As String
Dim Coverage As String
Dim Deployed(1 To 2) As Variant
Dim Suggested(1 To 2) As String
Dim SegArray(1 To 6) As Variant
Dim ChartCode As String
Dim PriorityRules As String


Set ThemeHeatMap = ActiveSheet.PivotTables("PivotTable10")

'Filter Fields
Set ConduciveField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsConduciveToAutomatedMonitoring].[IsConduciveToAutomatedMonitoring]")
Set ApplicableField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsApplicableToBank].[IsApplicableToBank]")
Set DuplicateField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsDuplicate].[IsDuplicate]")
Set PriorityField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Priority].[Priority]")
Set SegmentField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Product Segment].[Product Segment]")
Set CoverageField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Oracle Rule Indicator].[Oracle Rule Indicator]")


'Conducive.EnableMultiplePageItems = True
'Conducive.PivotItems("N").Visible = False

Conducive = "[Pivot Theme Coverage].[IsConduciveToAutomatedMonitoring]."
Applicable = "[Pivot Theme Coverage].[IsApplicableToBank]."
Duplicate = "[Pivot Theme Coverage].[IsDuplicate]."
Priority = "[Pivot Theme Coverage].[Priority]."
Segment = "[Pivot Theme Coverage].[Product Segment]."
Coverage = "[Pivot Theme Coverage].[Oracle Rule Indicator]."

'Activate Multiple Priorities
ThemeHeatMap.CubeFields(46).EnableMultiplePageItems = True

Suggested(1) = Priority & "&[2]"
Suggested(2) = Priority & "&[3]"

SegArray(1) = "&[CB - Commercial]"
SegArray(2) = "&[CB - Retail]"
SegArray(3) = "&[CB - Small Business]"
SegArray(4) = "&[Financial Institutions]"
SegArray(5) = "&[GBM]"
SegArray(6) = "&[Wealth]"


Deployed(1) = Priority & "&[1]"
Deployed(2) = Suggested

'PriorityField.VisibleItemsList = Deployed(2)

ThemeHeatMap.ClearAllFilters
ConduciveField.CurrentPageName = Conducive & "&[Y]"
ApplicableField.CurrentPageName = Applicable & "&[Y]"
DuplicateField.CurrentPageName = Duplicate & "&[N]"

'NumSegments = UBound(SegArray) - LBound(SegArray) + 1

'For SegNum = 0 To NumSegments
For SegSwitch = 0 To 1
    SegmentField.ClearAllFilters
    If SegSwitch = 0 Then
        For RuleSet = 1 To 2
            PriorityField.ClearAllFilters
            PriorityField.VisibleItemsList = Array(Deployed(RuleSet))
            If RuleSet = 1 Then
                PriorityRules = "&[Required]"
            ElseIf RuleSet = 2 Then
                PriorityRules = "&[Suggested]"
            End If
            ChartCode = "[Theme Chart]&[All Segments]" & PriorityRules
            MsgBox (ChartCode)
            'Call SegmentFormatting
            'Range("A1").End(xlDown).End(xlDown).CurrentRegion.Copy
        Next RuleSet
    ElseIf SegSwitch = 1 Then
        For Each Seg In SegArray
            SegmentField.CurrentPageName = Segment & Seg
            For RuleSet = 1 To 2
                PriorityField.ClearAllFilters
                PriorityField.VisibleItemsList = Array(Deployed(RuleSet))
                If RuleSet = 1 Then
                    PriorityRules = "&[Required]"
                ElseIf RuleSet = 2 Then
                    PriorityRules = "&[Suggested]"
                End If
                ChartCode = "[Theme Chart]" & Seg & PriorityRules
                MsgBox (ChartCode)
                'Call SegmentFormatting
                'Range("A1").End(xlDown).End(xlDown).CurrentRegion.Copy
            Next RuleSet
        Next Seg
    End If
Next SegSwitch

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RuleChartCopyPaste()
'
' HeatMapCopyPaste Macro
'

Dim ThemeHeatMap As PivotTable
Dim ConduciveField As PivotField
Dim ApplicableField As PivotField
Dim DuplicateField As PivotField
Dim PriorityField As PivotField
Dim SegmentField As PivotField
Dim CoverageField As PivotField

Dim Conducive As String
Dim Applicable As String
Dim Duplicate As String
Dim Priority As String
Dim Segment As String
Dim Coverage As String
Dim Deployed(1 To 2) As Variant
Dim Suggested(1 To 2) As String
Dim SegArray(1 To 6) As Variant
Dim CovArray(1 To 3) As Variant
Dim ChartCode As String
Dim PriorityRules As String

Set ThemeHeatMap = ActiveSheet.PivotTables("PivotTable10")

'Filter Fields
Set ConduciveField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsConduciveToAutomatedMonitoring].[IsConduciveToAutomatedMonitoring]")
Set ApplicableField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsApplicableToBank].[IsApplicableToBank]")
Set DuplicateField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[IsDuplicate].[IsDuplicate]")
Set PriorityField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Priority].[Priority]")
Set SegmentField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Product Segment].[Product Segment]")
Set CoverageField = ThemeHeatMap.PivotFields("[Pivot Theme Coverage].[Oracle Rule Indicator].[Oracle Rule Indicator]")


'Conducive.EnableMultiplePageItems = True
'Conducive.PivotItems("N").Visible = False

Conducive = "[Pivot Theme Coverage].[IsConduciveToAutomatedMonitoring]."
Applicable = "[Pivot Theme Coverage].[IsApplicableToBank]."
Duplicate = "[Pivot Theme Coverage].[IsDuplicate]."
Priority = "[Pivot Theme Coverage].[Priority]."
Segment = "[Pivot Theme Coverage].[Product Segment]."
Coverage = "[Pivot Theme Coverage].[Oracle Rule Indicator]."

'Activate Multiple Priorities
ThemeHeatMap.CubeFields(46).EnableMultiplePageItems = True

Suggested(1) = Priority & "&[2]"
Suggested(2) = Priority & "&[3]"

SegArray(1) = "&[CB - Commercial]"
SegArray(2) = "&[CB - Retail]"
SegArray(3) = "&[CB - Small Business]"
SegArray(4) = "&[Financial Institutions]"
SegArray(5) = "&[GBM]"
SegArray(6) = "&[Wealth]"

CovArray(1) = "&[Full]"
CovArray(2) = "&[Partial]"
CovArray(3) = "&[None]"


Deployed(1) = Priority & "&[1]"
Deployed(2) = Suggested

'PriorityField.VisibleItemsList = Deployed(2)

ThemeHeatMap.ClearAllFilters
ConduciveField.CurrentPageName = Conducive & "&[Y]"
ApplicableField.CurrentPageName = Applicable & "&[Y]"
DuplicateField.CurrentPageName = Duplicate & "&[N]"

'NumSegments = UBound(SegArray) - LBound(SegArray) + 1

'For SegNum = 0 To NumSegments
For SegSwitch = 0 To 1
    SegmentField.ClearAllFilters
    If SegSwitch = 0 Then
        For RuleSet = 1 To 2
            PriorityField.ClearAllFilters
            PriorityField.VisibleItemsList = Array(Deployed(RuleSet))
            For Each CovSet In CovArray
                CoverageField.ClearAllFilters
                CoverageField.CurrentPageName = Coverage & CovSet
                If RuleSet = 1 Then
                    PriorityRules = "&[Required]"
                ElseIf RuleSet = 2 Then
                    PriorityRules = "&[Suggested]"
                End If
                ChartCode = "&[All Segments]" & PriorityRules & CovSet
                MsgBox (ChartCode)
                'Call SegmentFormatting
                'Range("A1").End(xlDown).End(xlDown).CurrentRegion.Copy
            Next CovSet
        Next RuleSet
    ElseIf SegSwitch = 1 Then
        For Each Seg In SegArray
            SegmentField.CurrentPageName = Segment & Seg
            For RuleSet = 1 To 2
                PriorityField.ClearAllFilters
                PriorityField.VisibleItemsList = Array(Deployed(RuleSet))
                For Each CovSet In CovArray
                    CoverageField.ClearAllFilters
                    CoverageField.CurrentPageName = Coverage & CovSet
                    If RuleSet = 1 Then
                        PriorityRules = "&[Required]"
                    ElseIf RuleSet = 2 Then
                        PriorityRules = "&[Suggested]"
                    End If
                    ChartCode = Seg & PriorityRules & CovSet
                    MsgBox (ChartCode)
                    'Call SegmentFormatting
                    'Range("A1").End(xlDown).End(xlDown).CurrentRegion.Copy
                Next CovSet
            Next RuleSet
        Next Seg
    End If
Next SegSwitch

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use to Split Long Column to Many and Formula for indiv lists
'=TEXTJOIN(",", TRUE, H2:Q2)

Sub SplitColumn()
    'Updateby20141106
    Dim rng As Range
    Dim InputRng As Range
    Dim OutRng As Range
    Dim xRow As Integer
    Dim xCol As Integer
    Dim xArr As Variant
    xTitleId = "KutoolsforExcel"
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", xTitleId, InputRng.Address, Type:=8)
    xRow = Application.InputBox("Rows :", xTitleId)
    Set OutRng = Application.InputBox("Out put to (single cell):", xTitleId, Type:=8)
    Set InputRng = InputRng.Columns(1)
    xCol = InputRng.Cells.Count / xRow
    ReDim xArr(1 To xRow, 1 To xCol + 1)
    For i = 0 To InputRng.Cells.Count - 1
        xValue = InputRng.Cells(i + 1)
        iRow = i Mod xRow
        iCol = VBA.Int(i / xRow)
        xArr(iRow + 1, iCol + 1) = xValue
    Next
    OutRng.Resize(UBound(xArr, 1), UBound(xArr, 2)).Value = xArr
End Sub
