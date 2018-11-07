Sub createPivot()

    Dim pivSht As Worksheet, dataSht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String, pivName As String
    
    Set dataSht = Worksheets("Data")
    Set pivSht = Worksheets("Pivot")
    pivName = "PivotTable1"

    'Determine the data range you want to pivot
      SrcData = dataSht.Name & "!" & Range(Range("A1:H1"), Range("A1:H1").End(xlDown)).Address(ReferenceStyle:=xlR1C1)

    'Where do you want Pivot Table to start?
      StartPvt = pivSht.Name & "!" & pivSht.Range("A3").Address(ReferenceStyle:=xlR1C1)
    
    'Create Pivot Cache from Source Data
      Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=SrcData)
    
    'Create Pivot table from Pivot Cache
      Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=StartPvt, _
        TableName:=pivName)
    Set pvt = pivSht.PivotTables(pivName)
    'Add item to the Report Filter
    'pvt.PivotFields("feeamount").Orientation = xlPageField
  
    'Add item to the Column Labels
    pvt.PivotFields("groupname").Orientation = xlColumnField
    
    'Add item to the Row Labels
    pvt.PivotFields("date_created").Orientation = xlRowField
    
    Dim pf_Name As String
    pf_Name = "Sum of fees"
    
    ' Add item to Values
    pvt.AddDataField pvt.PivotFields("feeamount"), pf_Name, xlSum

    'Turn Off Grand Totals for Rows and Columns
    pvt.ColumnGrand = False
    pvt.RowGrand = False
    
    

End Sub
