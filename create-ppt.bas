Attribute VB_Name = "Módulo1"
Sub FillPPT()
    
    Dim pptPath As String
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim Pos1(1) As Double
    Dim Pos2(1) As Double
    Dim Pos3(1) As Double
    Dim Pos4(1) As Double
    Dim Pos5(1) As Double
    Dim wbPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheet As String
    Dim slides As Variant
    Dim chart1 As Variant
    Dim chart1Var As Variant
    Dim table1 As Variant
    Dim chart2 As Variant
    Dim i As Integer
    Dim pptSlide As Object
    Dim chart As Object
    Dim table As Range
    Dim pptShape As Object
    
    ' Opens a PowerPoint presentation from a file
    pptPath = ThisWorkbook.Path & "\presentation\sources\templatePPT.pptx" ' Registers the path of the presentation template (template.pptx)
    Set pptApp = CreateObject("PowerPoint.Application") ' Initializes PowerPoint
    Set pptPresentation = pptApp.Presentations.Open(pptPath) ' Loads the presentation template
    
    ' Registers the coordinates of the 1st graph (points)
    Pos1(0) = 30.3
    Pos1(1) = 116.9
    
    ' Registers the coordinates of the variation of the 1st graph (points)
    Pos2(0) = 204.1
    Pos2(1) = 114.6
    
    ' Registers the coordinates of the table (points)
    Pos3(0) = 380.2
    Pos3(1) = 157.1
    
    ' Registers the coordinates of the 2st graph (points)
    Pos4(0) = 33.5
    Pos4(1) = 367.9
    
    ' Opens an Excel workbook from a file
    wbPath = ThisWorkbook.Path & "\presentation\sources\data.xlsx" ' Registers the path of the workbook with the data (data.xlsx)
    Set wb = Workbooks.Open(wbPath) ' Loads the workbook
    
    ' Loads the data from the worksheet into the presentation template
    sheet = "data01" ' Registers the name of the worksheet
    Set ws = wb.Sheets(sheet) ' Selects the worksheet
    
    slides = Array(2, 4, 6, 8) ' Registers the slide numbers
    chart1 = Array("Chart01", "Chart03", "Chart05", "Chart07") ' Registers the names of the 1st graphs
    chart1Var = Array("BA10:BA10", "BA20:BA20", "BA30:BA30", "BA40:BA40") ' Registers the ranges of the variations of the 1st graphs
    table1 = Array("AO10:AO15", "AO20:AO25", "AO30:AO35", "AO40:AO45") ' Registers the ranges of the tables
    chart2 = Array("Chart02", "Chart04", "Chart06", "Chart08") ' Registers the names of the 2st graphs
    
    For i = 0 To UBound(slides)
        
        Set pptSlide = pptPresentation.slides(slides(i))
        
        Set chart = ws.ChartObjects(chart1(i)) ' Selects the 1st graph in the worksheet
        chart.Copy ' Copy the 1st graph
        Set pptShape = pptSlide.Shapes.Paste ' Pastes the 1st graph in the presentation
        pptShape.Left = Pos1(0) ' Positions the 1st graph on the left
        pptShape.Top = Pos1(1) ' Positions the 1st graph on the top
        
        Set table = ws.Range(chart1Var(i)) ' Selects the variation of the 1st graph in the worksheet
        table.Copy ' Copy the variation of the 1st graph
        pptSlide.Shapes.PasteSpecial DataType:=2 ' Pastes the variation of the 1st graph in the presentation as an image
        With pptSlide.Shapes(pptSlide.Shapes.Count)
            .Left = Pos2(0) ' Positions the variation of the 1st graph on the left
            .Top = Pos2(1) ' Positions the variation of the 1st graph on the top
        End With
        
        Set table = ws.Range(table1(i)) ' Selects the table in the worksheet
        table.Copy ' Copy the table
        pptSlide.Shapes.PasteSpecial DataType:=2 ' Pastes the table in the presentation as an image
        With pptSlide.Shapes(pptSlide.Shapes.Count)
            .Left = Pos3(0) ' Positions the table on the left
            .Top = Pos3(1) ' Positions the table on the top
        End With
        
        Set chart = ws.ChartObjects(chart2(i)) ' Selects the 2st graph in the worksheet
        chart.Copy ' Copy the 2st graph
        Set pptShape = pptSlide.Shapes.PasteSpecial(DataType:=2) ' Pastes the 2st graph in the presentation as an image
        pptShape.Left = Pos4(0) ' Positions the 2st graph on the left
        pptShape.Top = Pos4(1) ' Positions the 2st graph on the top
        
        Application.Wait (Now + TimeValue("0:00:05"))
        
        Next i
    
    ' Close and save the updated presentation
    pptPresentation.SaveAs ThisWorkbook.Path & "\presentation\aaaamm - presentation update.pptx"
    pptPresentation.Close
    pptApp.Quit
    
End Sub
