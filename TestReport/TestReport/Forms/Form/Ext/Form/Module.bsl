
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
	//	|	Contracts.Ref AS Contract
	//	|FROM
	//	|	Catalog.ДоговорыКонтрагентов AS Contracts";
	//	Return Query.Execute().Unload();
	
	ReportData = New ValueTable;
	ReportData.Columns.Add("Client");
	ReportData.Columns.Add("Contract");
	
	AddDataRow(ReportData, "Alpha LLC",    "Contract No. 1 dated 01.01.2026");
	AddDataRow(ReportData, "Beta LLC",     "Contract No. 2 dated 15.02.2026");
	AddDataRow(ReportData, "John Smith",   "Contract No. 3 dated 10.03.2026");
	
	Return ReportData;
	
EndFunction

&AtServer
Procedure AddDataRow(ReportData, Client, Contract)
	NewRow = ReportData.Add();
	NewRow.Client   = Client;
	NewRow.Contract = Contract;
EndProcedure
