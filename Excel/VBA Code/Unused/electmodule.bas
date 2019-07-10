Sub Elect_Engr_BOM()




    'Elect_Engr_BOM Macro

    'Turn off screen updating while executing macro
    Application.ScreenUpdating = False
    'Sheets("Master").Visible = True
    'Sheets("Data Validation").Visible = True
    

    'Copy Columns and Ranges
    Sheets("Master").Select
    Range("A1:X33").Select
    Selection.Copy
    
    'Paste Columns and Ranges into new sheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Application.PrintCommunication = False
    
    
    
   'New Page Setup
   With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PrintArea = ""
        
        'New Page Margins
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.1)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        
    End With
    
   'Hide Notes_Comp Weight_Total Weight Columns
    Columns("M:O").Select
    Selection.EntireColumn.Hidden = True
    
    'Hide Estimating and Mech Titles
    Rows("8:10").Select
    
    Selection.EntireRow.Hidden = True
    
    'Set Title Row Height
    Rows("11:11").RowHeight = 12.5
    Rows("12:12").RowHeight = 12.5
    
    'Set Body Row Height
    Rows("13:33").RowHeight = 40
    
    
    'Set Disclaimer Block
    Range("A7:F7").Select
    With Selection
        '.Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .RowHeight = 27
    End With
        
    
    'Merge and Position BOM TITLE
    Range("A4:D4").Select
    With Selection
        '.Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'Merge and Position SM PROJECT DESCRIPTION
    Range("A6:D6").Select
     With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With

    
   
    
    
    'Set Thick Borders around new sheet
    Range("A1:X33").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    'Restore functional settings and restore to Normal View
    ActiveWindow.View = xlNormalView
    Range("A13").Select
    Application.PrintCommunication = True
    Application.ScreenUpdating = True


End Sub
