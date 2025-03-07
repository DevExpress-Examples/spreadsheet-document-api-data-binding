<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/944428246/24.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1281203)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
# Spreadsheet Document API: Bind a Worksheet to a Generic List or a BindingList Data Source

This example demonstrates the use of a [List\<T\>](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=net-7.0) and [BindingList\<T\>](https://learn.microsoft.com/en-us/dotnet/api/system.componentmodel.bindinglist-1?view=net-7.0) objects as data sources to bind data to the worksheet range.

## Implementation Details

Use the [WorksheetDataBindingCollection.BindToDataSource](https://docs.devexpress.com/OfficeFileAPI/devexpress.spreadsheet.worksheetdatabindingcollection.bindtodatasource.overloads) method to bind data to the range, and the [WorksheetDataBindingCollection.BindTableToDataSource](https://docs.devexpress.com/OfficeFileAPI/devexpress.spreadsheet.worksheetdatabindingcollection.bindtabletodatasource.overloads) method to bind data to the worksheet table.

The [ExternalDataSourceOptions](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.ExternalDataSourceOptions) object specifies various data binding options. A custom converter with the [IBindingRangeValueConverter](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IBindingRangeValueConverter) interface converts weather data between the data source and a worksheet.

If the data source does not allow modification, the binding worksheet range also prevents modification.

Data binding error results in the [WorksheetDataBinding.Error](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.WorksheetDataBindingCollection.Error) event and cancels data update. The event handler in this example displays a message containing the error type.

## Files to Review

* [Program.cs](./CS/SpreadsheetApiDataBinding/Program.cs) (VB: [Program.vb](./VB/SpreadsheetApiDataBinding/Program.vb))
* [MyConverter.cs](./CS/SpreadsheetApiDataBinding/MyConverter.cs) (VB: [MyConverter.vb](./VB/SpreadsheetApiDataBinding/MyConverter.vb))
* [WeatherReport.cs](./CS/SpreadsheetApiDataBinding/WeatherReport.cs) (VB: [WeatherReport.vb](./VB/SpreadsheetApiDataBinding/WeatherReport.vb))

## Documentation

* [Spreadsheet Data Binding](https://docs.devexpress.com/OfficeFileAPI/118785/spreadsheet-document-api/data-binding)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-data-binding&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-data-binding&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
