
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
	
	WidgetHTML = BuildWidgetHTML();
	
EndProcedure

&AtClient
Procedure OnOpen(Cancel)
	
	// The clock ticks inside the widget itself (JavaScript, once per second).
	// The idle handler refreshes the weather data every 10 minutes.
	AttachIdleHandler("RefreshWeather", 600, False);
	
EndProcedure

&AtClient
Procedure OnClose(Exit)
	
	DetachIdleHandler("RefreshWeather");
	
EndProcedure

#EndRegion

#Region Widget

&AtClient
Procedure RefreshWeather() Export
	
	WidgetHTML = BuildWidgetHTMLAtServer();
	
EndProcedure

&AtClient
Procedure WidgetOnClick(Item, EventData, StandardProcessing)
	
	// Links in the widget look like "#act=<action>&value=<url-encoded text>".
	// We catch the click, stop the browser from navigating and run a 1C action instead.
	Href = EventData.Href;
	If Not ValueIsFilled(Href) Then
		Return;
	EndIf;
	
	Position = StrFind(Href, "#act=");
	If Position = 0 Then
		Return; // an ordinary link - let the HTML field handle it
	EndIf;
	
	StandardProcessing = False;
	
	Action = Mid(Href, Position + 5);
	Value = "";
	ValuePosition = StrFind(Action, "&value=");
	If ValuePosition > 0 Then
		Value = DecodeString(Mid(Action, ValuePosition + 7), StringEncodingMethod.URLEncoding);
		Action = Left(Action, ValuePosition - 1);
	EndIf;
	
	If Action = "report" Then
		ShowReport(Undefined);
		
	ElsIf Action = "notify" Then
		ShowUserNotification("Weather in Chisinau", , Value, PictureLib.Information);
		
	ElsIf Action = "day" Then
		ShowMessageBox(, Value, , "Forecast for Chisinau");
		
	ElsIf Action = "url" Then
		GotoURL(Value);
		
	ElsIf Action = "open" Then
		OpenAnyLink(Value);
		
	EndIf;
	
EndProcedure

// Opens a 1C navigation link (e1cib/...), a web address or a file path.
&AtClient
Procedure OpenAnyLink(Link)
	
	If Not ValueIsFilled(Link) Then
		ShowMessageBox(, "Paste a link first: a 1C navigation link (Shift+F11 on any document or list, e.g. e1cib/list/Catalog.Products) or a web address.");
		Return;
	EndIf;
	
	If StrStartsWith(Lower(Link), "www.") Then
		Link = "https://" + Link;
	EndIf;
	
	Try
		GotoURL(Link);
	Except
		ShowMessageBox(, "Cannot open the link:" + Chars.LF + Link + Chars.LF + BriefErrorDescription(ErrorInfo()));
	EndTry;
	
EndProcedure

&AtServer
Function BuildWidgetHTMLAtServer()
	
	Return BuildWidgetHTML();
	
EndFunction

// Builds the clock + weather widget from the WeatherWidget template.
&AtServer
Function BuildWidgetHTML()
	
	ReportObject = FormAttributeToValue("Object");
	HTMLText = ReportObject.GetTemplate("WeatherWidget").GetText();
	
	HTMLText = StrReplace(HTMLText, "/*WEATHER_JSON*/null", GetWeatherJSON());
	HTMLText = StrReplace(HTMLText, "/*UPDATED*/", Format(CurrentSessionDate(), "DF=HH:mm"));
	
	Return HTMLText;
	
EndFunction

#EndRegion

#Region Weather

// Current weather and a 5-day forecast for Chisinau from Open-Meteo (no API key needed).
// Returns the JSON text for the widget, or {"error": "..."} if the service is unavailable.
&AtServer
Function GetWeatherJSON()
	
	Try
		
		Connection = New HTTPConnection("api.open-meteo.com", 443, , , , 10,
			New OpenSSLSecureConnection);
		Request = New HTTPRequest("/v1/forecast?latitude=47.0105&longitude=28.8638"
			+ "&current=temperature_2m,apparent_temperature,relative_humidity_2m,weather_code,wind_speed_10m,is_day"
			+ "&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max"
			+ "&timezone=Europe%2FChisinau&forecast_days=5");
		Response = Connection.Get(Request);
		
		If Response.StatusCode <> 200 Then
			Return ErrorJSON("service returned HTTP " + Response.StatusCode);
		EndIf;
		
		Return Response.GetBodyAsString("UTF-8");
		
	Except
		Return ErrorJSON(BriefErrorDescription(ErrorInfo()));
	EndTry;
	
EndFunction

&AtServer
Function ErrorJSON(Text)
	
	Writer = New JSONWriter;
	Writer.SetString();
	Data = New Structure("error", Text);
	WriteJSON(Writer, Data);
	Return Writer.Close();
	
EndFunction

#EndRegion
