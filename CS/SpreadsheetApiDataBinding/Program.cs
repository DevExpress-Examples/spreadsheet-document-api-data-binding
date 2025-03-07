using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace SpreadsheetApiDataBinding
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();
            workbook.CreateNewDocument();

            // Bind a first worksheet to a BindingList
            BindWeatherReportToRange(MyWeatherReportSource.DataAsBindingList, workbook.Worksheets[0]);
            
            // Bind a second worksheet to a List
            workbook.Worksheets.Add();
            BindWeatherReportToTable(MyWeatherReportSource.Data, workbook.Worksheets[1]);
            
            // Bind a data source to the fixed table on the third worksheet
            workbook.Worksheets.Add();
            BindWeatherReportToFixedTable(MyWeatherReportSource.Data, workbook.Worksheets[2]);
            workbook.Worksheets[0].DataBindings.Error += DataBindings_Error;
            workbook.SaveDocument("DataBindings.xlsx");
            Process.Start(new ProcessStartInfo("DataBindings.xlsx") { UseShellExecute = true });
        }

        private static void DataBindings_Error(object sender, DataBindingErrorEventArgs e)
        {
            Console.WriteLine(String.Format("Error at worksheet.Rows[{0}].\n The error is : {1}", e.RowIndex, e.ErrorType.ToString()), "Binding Error");
        }

        private static void BindWeatherReportToRange(object weatherDatasource, Worksheet worksheet)
        {
            // Check for range conflicts.
            CellRange bindingRange = worksheet.Range["A1:C5"];
            var dataBindingConflicts = worksheet.DataBindings.
                Where(binding => (binding.Range.RightColumnIndex >= bindingRange.LeftColumnIndex)
                || (binding.Range.BottomRowIndex >= bindingRange.TopRowIndex));
            if (dataBindingConflicts.Count() > 0)
            {
                Console.WriteLine("Cannot bind the range to data.\r\nThe worksheet contains other binding ranges which may conflict.", "Range Conflict");
                return;
            }

            // Specify the binding options.
            ExternalDataSourceOptions dsOptions = new ExternalDataSourceOptions();
            dsOptions.ImportHeaders = true;
            dsOptions.CellValueConverter = new MyWeatherConverter();
            dsOptions.SkipHiddenRows = true;

            // Bind the data source to the worksheet range.
            WorksheetDataBinding sheetDataBinding = worksheet.DataBindings.BindToDataSource(weatherDatasource, bindingRange, dsOptions);

            // Adjust the column width.
            sheetDataBinding.Range.AutoFitColumns();
        }

        private static void BindWeatherReportToTable(object weatherDatasource, Worksheet worksheet)
        {
            CellRange bindingRange = worksheet["A1:C5"];
            // Remove all data bindings bound to the specified data source.
            worksheet.DataBindings.Remove(weatherDatasource);

            // Specify the binding options.
            ExternalDataSourceOptions dsOptions = new ExternalDataSourceOptions();
            dsOptions.ImportHeaders = true;
            dsOptions.CellValueConverter = new MyWeatherConverter();
            dsOptions.SkipHiddenRows = true;

            // Create a table and bind the data source to the table.
            try
            {
                WorksheetTableDataBinding sheetDataBinding = worksheet.DataBindings.BindTableToDataSource(weatherDatasource, bindingRange, dsOptions);
                sheetDataBinding.Table.Style = worksheet.Workbook.TableStyles[BuiltInTableStyleId.TableStyleMedium14];

                // Adjust the column width.
                sheetDataBinding.Range.AutoFitColumns();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Binding Exception");
            }
        }

        private static void BindWeatherReportToFixedTable(object weatherDatasource, Worksheet worksheet)
        {
            // Remove all data bindings bound to the specified data source.
            worksheet.DataBindings.Remove(weatherDatasource);

            CellRange bindingRange = worksheet["A1:C5"];

            // Specify the binding options.
            ExternalDataSourceOptions dsOptions = new ExternalDataSourceOptions();
            dsOptions.ImportHeaders = true;
            dsOptions.CellValueConverter = new MyWeatherConverter();
            dsOptions.SkipHiddenRows = true;

            // Create a table and bind the data source to the table.
            try
            {
                Table boundTable = worksheet.Tables.Add(weatherDatasource, bindingRange, dsOptions);
                boundTable.Style = worksheet.Workbook.TableStyles[BuiltInTableStyleId.TableStyleMedium15];

                // Adjust the column width.
                boundTable.Range.AutoFitColumns();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Binding Exception");
            }
        }
    }
}
