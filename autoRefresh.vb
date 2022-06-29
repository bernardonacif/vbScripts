
' scrip autialização de toda a planilha inclusive dados externos e planilhas dinamicas

Sub RefreshPivotTables()
  'Objeto de tabela dinâmica
  Dim pivotTable As pivotTable
 
  'Loop por todos os objetos da planilha
  For Each plan In ActiveWorkbook.Sheets
    For Each pivotTable In plan.PivotTables
        pivotTable.RefreshTable
    Next
  Next
End Sub