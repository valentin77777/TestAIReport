
&AtClient
Procedure ShowReport(Command)
	
	SpreadsheetDoc = GenerateReportAtServer();
	
	// Show the template in a separate window
	SpreadsheetDoc.ShowGrid = False;
	SpreadsheetDoc.ShowHeaders = False;
	SpreadsheetDoc.ReadOnly = True;
	SpreadsheetDoc.Show("Report: clients and contracts");
	
EndProcedure

&AtServer
Function GenerateReportAtServer()
	
	ReportObject = FormAttributeToValue("Object");
	Template = ReportObject.GetTemplate("Template");
	
	SpreadsheetDoc = New SpreadsheetDocument;
	
	AreaTitle  = Template.GetArea("Title");
	AreaHeader = Template.GetArea("Header");
	AreaRow    = Template.GetArea("Row");
	
	AreaTitle.Parameters.ReportTitle = "Clients and contracts report";
	SpreadsheetDoc.Put(AreaTitle);
	SpreadsheetDoc.Put(AreaHeader);
	
	ReportData = GetReportData();
	LineNumber = 0;
	For Each DataRow In ReportData Do
		LineNumber = LineNumber + 1;
		AreaRow.Parameters.LineNumber = LineNumber;
		AreaRow.Parameters.Client     = DataRow.Client;
		AreaRow.Parameters.Contract   = DataRow.Contract;
		AreaRow.Parameters.ContractDate = DataRow.ContractDate;
		SpreadsheetDoc.Put(AreaRow);
	EndDo;
	
	Return SpreadsheetDoc;
	
EndFunction

&AtServer
Function GetReportData()
	
	// Sample data so the report runs in any infobase.
	// For real data, replace with a query to your configuration, e.g.:
	//	Query = New Query;
	//	Query.Text =
	//	"SELECT
	//	|	Contracts.Owner AS Client,
	//	|	Contracts.Ref AS Contract,
	//	|	Contracts.Date AS ContractDate
	//	|FROM
	//	|	Catalog.ДоговорыКонтрагентов AS Contracts";
	//	Return Query.Execute().Unload();
	
	ReportData = New ValueTable;
	ReportData.Columns.Add("Client");
	ReportData.Columns.Add("Contract");
	ReportData.Columns.Add("ContractDate", New TypeDescription("Date"));
	
	AddDataRow(ReportData, "Alpha LLC",  "Contract No. 1", Date(2026, 1, 1));
	AddDataRow(ReportData, "Beta LLC",   "Contract No. 2", Date(2026, 2, 15));
	AddDataRow(ReportData, "John Smith", "Contract No. 3", Date(2026, 3, 10));
	
	Return ReportData;
	
EndFunction

&AtServer
Procedure AddDataRow(ReportData, Client, Contract, ContractDate)
	NewRow = ReportData.Add();
	NewRow.Client   = Client;
	NewRow.Contract = Contract;
	NewRow.ContractDate = ContractDate;
EndProcedure

#Region FormEvents

&AtServer
Procedure OnCreateAtServer(Cancel, StandardProcessing)
	
	Items.WeatherLabel.Title = GetWeatherForecast();
	
EndProcedure

&AtClient
Procedure OnOpen(Cancel)
	
	UpdateClock();
	AttachIdleHandler("UpdateClock", 1, False);
	
EndProcedure

&AtClient
Procedure OnClose(Exit)
	
	DetachIdleHandler("UpdateClock");
	
EndProcedure

#EndRegion

#Region Clock

&AtClient
Procedure UpdateClock() Export
	
	Items.ClockLabel.Title = Format(CurrentDate(), "DF=dd.MM.yyyy HH:mm:ss");
	
EndProcedure

#EndRegion

#Region Weather

// Current weather and today's forecast for Chisinau from Open-Meteo (no API key needed).
&AtServer
Function GetWeatherForecast()
	
	Try
		
		Connection = New HTTPConnection("api.open-meteo.com", 443, , , , 10,
			New OpenSSLSecureConnection);
		Request = New HTTPRequest("/v1/forecast?latitude=47.0105&longitude=28.8638"
			+ "&current=temperature_2m,weather_code,wind_speed_10m"
			+ "&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max"
			+ "&timezone=Europe%2FChisinau&forecast_days=1");
		Response = Connection.Get(Request);
		
		If Response.StatusCode <> 200 Then
			Return "Weather in Chisinau: service unavailable (HTTP " + Response.StatusCode + ")";
		EndIf;
		
		Reader = New JSONReader;
		Reader.SetString(Response.GetBodyAsString("UTF-8"));
		Data = ReadJSON(Reader);
		Reader.Close();
		
		Current = Data.current;
		Daily = Data.daily;
		
		Return StrTemplate("Weather in Chisinau: %1, %2 °C, wind %3 km/h. Today: %4…%5 °C, precipitation %6%%",
			WeatherDescription(Current.weather_code),
			Round(Current.temperature_2m),
			Round(Current.wind_speed_10m),
			Round(Daily.temperature_2m_min[0]),
			Round(Daily.temperature_2m_max[0]),
			Daily.precipitation_probability_max[0]);
		
	Except
		Return "Weather in Chisinau: failed to load (" + BriefErrorDescription(ErrorInfo()) + ")";
	EndTry;
	
EndFunction

// WMO weather code -> text.
&AtServer
Function WeatherDescription(Code)
	
	If Code = 0 Then
		Return "clear sky";
	ElsIf Code <= 2 Then
		Return "partly cloudy";
	ElsIf Code = 3 Then
		Return "overcast";
	ElsIf Code <= 48 Then
		Return "fog";
	ElsIf Code <= 57 Then
		Return "drizzle";
	ElsIf Code <= 67 Then
		Return "rain";
	ElsIf Code <= 77 Then
		Return "snow";
	ElsIf Code <= 82 Then
		Return "rain showers";
	ElsIf Code <= 86 Then
		Return "snow showers";
	Else
		Return "thunderstorm";
	EndIf;
	
EndFunction

#EndRegion
