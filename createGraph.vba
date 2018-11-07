Sub graphResults()
    
    Dim dataSht, graphSht As Worksheet
    Set dataSht = Worksheets("Sample")
    Set graphSht = Worksheets("Graph")
    
    
    Dim rng1, rng2 As Range
    Dim cht As Object

    'Your data range for the chart
    Set rng1 = Range(dataSht.Range("A2"), dataSht.Range("A2").End(xlDown))
    Set rng2 = Range(dataSht.Range("B2"), dataSht.Range("B2").End(xlDown))
    

    'Create a chart
    Set cht = graphSht.Shapes.AddChart2
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.SeriesCollection(1).XValues = rng1
    cht.Chart.SeriesCollection(1).Values = rng2
    'Give chart some data
    'cht.Chart.SetSourceData Source:=rng

    'Determine the chart type
    cht.Chart.ChartType = xlXYScatter
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Characters.Text = "Test Chart Title"
End Sub