Option Infer On

Imports System.ComponentModel

Namespace SpreadsheetApiDataBinding
	Public Class WeatherReport
		<DisplayName("Date")>
		Public Property [Date]() As Date
		<DisplayName("Weather Condition")>
		Public Property Weather() As Weather
		<DisplayName("Max and Min Temperature")>
		Public Property HourlyReport() As List(Of HourlyReport)
	End Class

	Public Class HourlyReport
		Public Property Hour() As Integer
		Public Property Temperature() As Integer
	End Class

	Public Enum Weather
		Sunny
		Cloudy
		Windy
		Gloomy
		Foggy
		Misty
		Rainy
		Undefined
	End Enum
	Public Module MyWeatherReportSource
		Private rand As Random = New System.Random()
'INSTANT VB NOTE: The field data was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private data_Renamed As List(Of WeatherReport)
		Private dataBindingList As BindingList(Of WeatherReport)

		Public ReadOnly Property Data() As List(Of WeatherReport)
			Get
				If data_Renamed Is Nothing Then
					data_Renamed = GetReport()
				End If
				Return data_Renamed
			End Get
		End Property
		Public ReadOnly Property DataAsBindingList() As BindingList(Of WeatherReport)
			Get
				If dataBindingList Is Nothing Then
					dataBindingList = New BindingList(Of WeatherReport)(Data)
				End If
				Return dataBindingList
			End Get
		End Property
		Public Function GetReport() As List(Of WeatherReport)
			Dim report = New List(Of WeatherReport)()

			report.Add(New WeatherReport() With {
				.Date = Date.Today,
				.Weather = Weather.Rainy,
				.HourlyReport = GenerateRandomHourlyReport()
			})

			report.Add(New WeatherReport() With {
				.Date = Date.Today.AddDays(-1),
				.Weather = Weather.Cloudy,
				.HourlyReport = GenerateRandomHourlyReport()
			})

			report.Add(New WeatherReport() With {
				.Date = Date.Today.AddDays(-2),
				.Weather = Weather.Sunny,
				.HourlyReport = GenerateRandomHourlyReport()
			})

			report.Add(New WeatherReport() With {
				.Date = Date.Today.AddDays(-3),
				.Weather = Weather.Gloomy,
				.HourlyReport = GenerateRandomHourlyReport()
			})
			Return report
		End Function
		Public Function GenerateRandomHourlyReport() As List(Of HourlyReport)
			Dim report = New List(Of HourlyReport)()

			For i As Integer = 0 To 23
				Dim hourlyReport = New HourlyReport()
				hourlyReport.Hour = i
				hourlyReport.Temperature = rand.Next(30)
				report.Add(hourlyReport)
			Next i
			Return report
		End Function
		Public Sub Reload()
			data_Renamed = GetReport()
		End Sub
	End Module
End Namespace
