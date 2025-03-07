Option Infer On

Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace SpreadsheetApiDataBinding
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim workbook As New Workbook()
			workbook.CreateNewDocument()

			' Bind a first worksheet to a BindingList
			BindWeatherReportToRange(MyWeatherReportSource.DataAsBindingList, workbook.Worksheets(0))

			' Bind a second worksheet to a List
			workbook.Worksheets.Add()
			BindWeatherReportToTable(MyWeatherReportSource.Data, workbook.Worksheets(1))

			' Bind a data source to the fixed table on the third worksheet
			workbook.Worksheets.Add()
			BindWeatherReportToFixedTable(MyWeatherReportSource.Data, workbook.Worksheets(2))
			AddHandler workbook.Worksheets(0).DataBindings.Error, AddressOf DataBindings_Error
			workbook.SaveDocument("DataBindings.xlsx")
			Process.Start(New ProcessStartInfo("DataBindings.xlsx") With {.UseShellExecute = True})
		End Sub

		Private Shared Sub DataBindings_Error(ByVal sender As Object, ByVal e As DataBindingErrorEventArgs)
			Console.WriteLine(String.Format("Error at worksheet.Rows[{0}]." & vbLf & " The error is : {1}", e.RowIndex, e.ErrorType.ToString()), "Binding Error")
		End Sub

		Private Shared Sub BindWeatherReportToRange(ByVal weatherDatasource As Object, ByVal worksheet As Worksheet)
			' Check for range conflicts.
			Dim bindingRange As CellRange = worksheet.Range("A1:C5")
			Dim dataBindingConflicts = worksheet.DataBindings.Where(Function(binding) (binding.Range.RightColumnIndex >= bindingRange.LeftColumnIndex) OrElse (binding.Range.BottomRowIndex >= bindingRange.TopRowIndex))
			If dataBindingConflicts.Count() > 0 Then
				Console.WriteLine("Cannot bind the range to data." & vbCrLf & "The worksheet contains other binding ranges which may conflict.", "Range Conflict")
				Return
			End If

			' Specify the binding options.
			Dim dsOptions As New ExternalDataSourceOptions()
			dsOptions.ImportHeaders = True
			dsOptions.CellValueConverter = New MyWeatherConverter()
			dsOptions.SkipHiddenRows = True

			' Bind the data source to the worksheet range.
			Dim sheetDataBinding As WorksheetDataBinding = worksheet.DataBindings.BindToDataSource(weatherDatasource, bindingRange, dsOptions)

			' Adjust the column width.
			sheetDataBinding.Range.AutoFitColumns()
		End Sub

		Private Shared Sub BindWeatherReportToTable(ByVal weatherDatasource As Object, ByVal worksheet As Worksheet)
			Dim bindingRange As CellRange = worksheet("A1:C5")
			' Remove all data bindings bound to the specified data source.
			worksheet.DataBindings.Remove(weatherDatasource)

			' Specify the binding options.
			Dim dsOptions As New ExternalDataSourceOptions()
			dsOptions.ImportHeaders = True
			dsOptions.CellValueConverter = New MyWeatherConverter()
			dsOptions.SkipHiddenRows = True

			' Create a table and bind the data source to the table.
			Try
				Dim sheetDataBinding As WorksheetTableDataBinding = worksheet.DataBindings.BindTableToDataSource(weatherDatasource, bindingRange, dsOptions)
				sheetDataBinding.Table.Style = worksheet.Workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium14)

				' Adjust the column width.
				sheetDataBinding.Range.AutoFitColumns()
			Catch e As Exception
				Console.WriteLine(e.Message, "Binding Exception")
			End Try
		End Sub

		Private Shared Sub BindWeatherReportToFixedTable(ByVal weatherDatasource As Object, ByVal worksheet As Worksheet)
			' Remove all data bindings bound to the specified data source.
			worksheet.DataBindings.Remove(weatherDatasource)

			Dim bindingRange As CellRange = worksheet("A1:C5")

			' Specify the binding options.
			Dim dsOptions As New ExternalDataSourceOptions()
			dsOptions.ImportHeaders = True
			dsOptions.CellValueConverter = New MyWeatherConverter()
			dsOptions.SkipHiddenRows = True

			' Create a table and bind the data source to the table.
			Try
				Dim boundTable As Table = worksheet.Tables.Add(weatherDatasource, bindingRange, dsOptions)
				boundTable.Style = worksheet.Workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium15)

				' Adjust the column width.
				boundTable.Range.AutoFitColumns()
			Catch e As Exception
				Console.WriteLine(e.Message, "Binding Exception")
			End Try
		End Sub
	End Class
End Namespace
