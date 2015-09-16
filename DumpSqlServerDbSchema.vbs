Option Explicit
Const REFDATA_NONE = 0
Const REFDATA_REGEXP = 1
Const REFDATA_LIST = 2
Dim goDb, gbWriteToc, gnRefData, gsRefDataRegexp
Dim gaRefDataList()
Dim gsTitle
Dim goFormatter

''' Class to produce properly-indented output
Class Formatter
	Private nIndentLevel
	Private bUseTabs
	Private nSpacesPerTab
	Private oOutputStream
	
	Public Property Get IndentLevel()
		IndentLevel = nIndentLevel
	End Property
	
	Public Property Let IndentLevel(ByVal value)
		If Not IsNumeric(value) Or value < 0 Then
			Err.Raise 5
		End If
		nIndentLevel = CLng(value)
	End Property
	
	Public Property Get UseTabs()
		UseTabs = bUseTabs
	End Property
	
	Public Property Let UseTabs(ByVal value)
		bUseTabs = CBool(value)
	End Property
	
	Public Property Get SpacesPerTab()
		SpacesPerTab = nSpacesPerTab
	End Property
	
	Public Property Let SpacesPerTab(ByVal value)
		If Not IsNumeric(value) Or value < 1 Then
			Err.Raise 5
		End If
		nSpacesPerTab = CLng(value)
	End Property
	
	Public Property Get OutputStream()
		Set OutputStream = oOutputStream
	End Property
	
	Public Property Set OutputStream(value)
		Set oOutputStream = value
	End Property
	
	Public Sub Class_Initialize()
		IndentLevel = 0
		UseTabs = False
		SpacesPerTab = 3
		Set OutputStream = WScript.StdOut
	End Sub
	
	Public Sub Indent()
		IndentLevel = IndentLevel + 1
	End Sub
	
	Public Sub Outdent()
		If IndentLevel > 0 Then
			IndentLevel = IndentLevel - 1
		End If
	End Sub
	
	Private Function GetIndent()
		If UseTabs Then
			GetIndent = String(IndentLevel, Chr(9))
		Else
			GetIndent = Space(IndentLevel * SpacesPerTab)
		End If
	End Function
	
	Public Sub Write(ByVal s)
		OutputStream.Write(GetIndent() & s)
	End Sub
	
	Public Sub WriteLine(ByVal s)
		OutputStream.WriteLine(GetIndent() & s)
	End Sub
	
End Class


''' Renders all the tables in the database
Sub WriteTables()
	Dim i, oRs, sSql
	Dim bFirst, sLastSchemaName, sSchemaName, sTableName, sFullName, sLinkId, sDescription, sTableType
	
	sSql = "SELECT t.TABLE_SCHEMA, t.TABLE_NAME, " & _
		"CASE WHEN t.TABLE_TYPE = 'BASE TABLE' THEN 'table' ELSE 'view' END AS TABLE_TYPE , " & _
		"ISNULL(p.value, '') AS [DESCRIPTION] " & _
		"FROM INFORMATION_SCHEMA.TABLES t " & _
		"OUTER APPLY fn_listextendedproperty(N'MS_Description', N'Schema', t.TABLE_SCHEMA, N'TABLE', t.TABLE_NAME, DEFAULT, DEFAULT) p " & _
		"ORDER BY t.TABLE_SCHEMA, t.TABLE_TYPE, t.TABLE_NAME"
	Set oRs = goDb.Execute(sSql)
	With oRs
		' 2 passes: One to generate table of contents, one to write the table data
		For i = 1 To 2
			If i = 1 And gbWriteToc Then
				With goFormatter
					.WriteLine "<ul class=""toc"">"
					.Indent
				End With
			End If
			sLastSchemaName = ""
			bFirst = True
			Do Until .EOF
				sSchemaName = .Fields("TABLE_SCHEMA").Value
				sTableName = .Fields("TABLE_NAME").Value
				sTableType = .Fields("TABLE_TYPE").Value
				sDescription = .Fields("DESCRIPTION").Value
				sFullName = Interpolate("%1%.%2%", Array(sSchemaName, sTableName))
				sLinkId = "tbl_" & LCase(Replace(sFullName, ".", "_"))
				If i = 1 Then
					If gbWriteToc Then
						If sLastSchemaName <> sSchemaName Then
							If Not bFirst Then
								With goFormatter
									.Outdent
									.WriteLine "</ul>"
									.Outdent
									.WriteLine "</li>"
								End With
							End If
							With goFormatter
								.WriteLine Interpolate("<li>%1%", Array(HtmlEscape(sSchemaName)))
								.Indent
								.WriteLine "<ul>"
								.Indent
							End With
						End If
						goFormatter.WriteLine Interpolate( _
							"<li><a href=""#%1%"">%2%</a></li>", _
							Array(HtmlEscape(sLinkId), HtmlEscape(sTableName)) _
						)
					End If
				Else
					If bFirst Then
						With goFormatter
							.WriteLine "<div class=""main"">"
							.Indent
						End With
					End If
					With goFormatter
						.WriteLine "<div class=""table-container"">"
						.Indent
					End With
					goFormatter.WriteLine Interpolate( _
						"<h1 id=""%1%"" data-table-type=""%3%"">%4%: %2%</h1>", _
						Array( _
							HtmlEscape(sLinkId), _ 
							HtmlEscape(sFullName), _ 
							HtmlEscape(sTableType), _
							HtmlEscape(UCase(Left(sTableType, 1)) & Mid(sTableType, 2)) _ 
						) _
					)
					If Len(sDescription) > 0 Then
						goFormatter.WriteLine Interpolate("<p class=""description"">%1%</p>", Array(HtmlEscape(sDescription)))
					End If
					goFormatter.WriteLine "<h2>Columns</h2>"
					WriteTable sSchemaName, sTableName, sTableType
					If sTableType = "table" Then
						WriteFks sSchemaName, sTableName
					End If
					If IsRefDataTable(sSchemaName, sTableName) Then
						WriteRefData sSchemaName, sTableName
					End If
					With goFormatter
						.Outdent
						.WriteLine "</div>"
					End With
				End If
				.MoveNext
				sLastSchemaName = sSchemaName
				bFirst = False
			Loop
			If i = 1 Then
				If gbWriteToc Then
					With goFormatter
						.Outdent
						.WriteLine "</ul>"
						.Outdent
						.WriteLine "</li>"
						.Outdent
						.WriteLine "</ul>"
					End With
				End If
			Else
				With goFormatter
					.Outdent
					.WriteLine "</div>"
				End With
			End If
			.MoveFirst
		Next
		.Close
	End With
	Set oRs = Nothing
End Sub


''' Renders a database table's metadata (column name, data type, etc.) as an HTML table
Sub WriteTable(sSchemaName, sTableName, sTableType)
	Dim oRs, sSql, sColId
	sSql = Interpolate( _
		"SELECT c.COLUMN_NAME, " & _
		"UPPER(CASE " & _
		"  WHEN c.DATA_TYPE IN ('char', 'nchar', 'varchar', 'nvarchar') THEN " & _
		"    CASE " & _
		"      WHEN c.CHARACTER_MAXIMUM_LENGTH >= 0 THEN " & _
		"        c.DATA_TYPE + '(' + LTRIM(RTRIM(STR(c.CHARACTER_MAXIMUM_LENGTH))) + ')'  " & _
		"      ELSE " & _
		"        c.DATA_TYPE + '(max)' " & _
		"      END " & _
		"  ELSE c.DATA_TYPE " & _
		"END) AS DATA_TYPE, " & _
		"c.IS_NULLABLE, ISNULL(c.COLUMN_DEFAULT, '') AS COLUMN_DEFAULT, " & _
		"ISNULL(p.value, '') AS [DESCRIPTION] " & _
		"FROM INFORMATION_SCHEMA.COLUMNS c " & _
		"OUTER APPLY fn_listextendedproperty(N'MS_Description', N'Schema', c.TABLE_SCHEMA, N'TABLE', c.TABLE_NAME, N'COLUMN', c.COLUMN_NAME) p " & _
		"WHERE c.TABLE_SCHEMA = N'%1%' " & _
		"AND c.TABLE_NAME = N'%2%' " & _
		"ORDER BY ORDINAL_POSITION", _
		Array(SqlEscape(sSchemaName), SqlEscape(sTableName)) _
	)
	Set oRs = goDb.Execute(sSql)
	With oRs
		With goFormatter
			.WriteLine "<table>"
			.Indent
			.WriteLine "<thead>"
			.Indent
			.WriteLine "<tr>"
			.Indent
			.WriteLine "<th>Name</th><th>Type</th><th>Nullable?</th><th>Default</th><th>Description</th>"
			.Outdent
			.WriteLine "</tr>"
			.Outdent
			.WriteLine "</thead>"
			.WriteLine "<tbody>"
			.Indent
		End With
		Do Until .EOF
			sColId = LCase(Interpolate("col_%1%_%2%_%3%", Array(sSchemaName, sTableName, .Fields("COLUMN_NAME").Value)))
			goFormatter.WriteLine Interpolate("<tr id=""%1%"">", Array(sColId))
			goFormatter.Indent
			goFormatter.WriteLine Interpolate( _
				"<td><code>%1%</code></td><td><code>%2%</code></td><td>%3%</td><td>%4%&nbsp;</td><td>%5%&nbsp;</td>", _
				Array( _
					HtmlEscape(.Fields("COLUMN_NAME").Value), _
					HtmlEscape(.Fields("DATA_TYPE").Value), _
					HtmlEscape(.Fields("IS_NULLABLE").Value), _
					HtmlEscape(.Fields("COLUMN_DEFAULT").Value), _
					HtmlEscape(.Fields("DESCRIPTION").Value) _
				) _
			)
			With goFormatter
				.Outdent
				.WriteLine "</tr>"
			End With
			.MoveNext
		Loop
		With goFormatter
			.Outdent
			.WriteLine "</tbody>"
			.Outdent
			.WriteLine "</table>"
		End With
		.Close
	End With
	Set oRs = Nothing
End Sub


''' Renders the header portion of the output HTML document
Sub WriteHeader()
	With goFormatter
		.WriteLine "<!DOCTYPE html>"
		.WriteLine "<html>"
		.Indent
		.WriteLine "<head>"
		.Indent
		.WriteLine Interpolate("<title>%1%</title>", Array(HtmlEscape(gsTitle)))
		.WriteLine "<style>"
		.Indent		
		If gbWriteToc Then
			.WriteLine ".toc {"
			.Indent
			.WriteLine "float:left;"
			.WriteLine "width: 300px;"
			.Outdent
			.WriteLine "}"
			.WriteLine ".main {"
			.Indent
			.WriteLine "margin-left: 320px;"
			.Outdent
			.WriteLine "}"
		End If
		.WriteLine "h1 {"
		.Indent
		.WriteLine "margin-top: 72px;"
		.Outdent
		.WriteLine "}"
		.WriteLine "table {"
		.Indent
		.WriteLine "border: 1px solid silver;"
		.WriteLine "border-collapse: collapse;"
		.Outdent
		.WriteLine "}"
		.WriteLine "td, th {"
		.Indent
		.WriteLine "border: 1px solid silver;"
		.WriteLine "padding: 4px;"	
		.Outdent
		.WriteLine "}"
		.WriteLine "td.numeric {"
		.Indent
		.WriteLine "text-align: right;"
		.Outdent
		.WriteLine "}"		
		.WriteLine "ul.popup-menu {"
		.Indent
		.WriteLine "background-color: #eeeeee;"
		.WriteLine "border: 1px solid #cccccc;"
		.WriteLine "margin: 0;"
		.WriteLine "padding: 0;"
		.Outdent
		.WriteLine "}"
		.WriteLine ""
		.WriteLine "ul.popup-menu li {"
		.Indent
		.WriteLine "list-style-type: none;"
		.Outdent
		.WriteLine "}"
		.WriteLine ""
		.WriteLine "ul.popup-menu li a:visited {"
		.Indent
		.WriteLine "color: inherit;"
		.Outdent
		.WriteLine "}"
		.WriteLine ""
		.WriteLine "ul.popup-menu li a:hover {"
		.Indent
		.WriteLine "background-color: navy;"
		.WriteLine "color: white;"
		.Outdent
		.WriteLine "}"
		.WriteLine ""
		.WriteLine "ul.popup-menu li a {"
		.Indent
		.WriteLine "display: block;"
		.WriteLine "padding-left: 4pt;"
		.WriteLine "padding-right: 4pt;"
		.WriteLine "text-decoration: none;"
		.Outdent
		.WriteLine "}"
		.Outdent
		.WriteLine "</style>"
		.Outdent
		.WriteLine "</head>"
		.WriteLine "<body>"
		.Indent
	End With
End Sub


''' Renders the header portion of the output HTML document
Sub WriteFooter()
	With goFormatter
		.WriteLine "<script src=""http://code.jquery.com/jquery-1.11.3.min.js""></script>"
		' The un-minified version of this code is in PopupMenu.js
		.WriteLine "<script>"
		.Indent
		.WriteLine "var timer=-1;"
		.WriteLine "function hidePopup(){$("".popup-menu"").hide();timer=-1}"
		.WriteLine "function showPopup(c,a,h){var g=$("".popup-menu"");g.empty();for(var e=0;e<c.length;++e){var f=c[e],d=$(""<li></li>""),b=$(""<a></a>"").attr(""href"",""#""+f.id).text(f.text);d.append(b);g.append(d)}if(timer>=0){window.clearTimeout(timer)}g.css(""position"",""absolute"").css(""left"",a).css(""top"",h).show()}"
		.WriteLine "$(document).ready(function(){$(""<ul></ul>"").attr(""class"",""popup-menu"").css(""display"",""none"").on(""mouseleave"",function(){if(timer==-1){timer=window.setTimeout(hidePopup,1000)}}).on(""mouseenter"",function(){if(timer>=0){window.clearTimeout(timer);timer=-1}}).insertAfter("".main"");$(""H1"").each(function(){var b=$(this),d=b.attr(""id""),a=[];if(d){var c=$(""a[href='#""+d+""']"");c.each(function(){var g=$(this),h=g.closest("".table-container""),e=h.find(""h1"").filter("":first""),f=e.attr(""id"");if(e.length>0&&!a.some(function(l,k,j){return l&&l.id&&l.id===f})){a.push({id:f,text:e.text()})}})}if(a.length>0){b.data(""refs"",a).css(""cursor"",""pointer"").on(""click"",function(g){var f=$(this).data(""refs"");showPopup(f,g.target.offsetLeft+g.offsetX,g.target.offsetTop+g.offsetY)})}})});"
		.Outdent
		.WriteLine "</script>"
		.Outdent
		.WriteLine "</body>"
		.Outdent
		.WriteLine "</html>"
	End With
End Sub


''' Renders a table's foreign keys as a list of links
Sub WriteFks(sSchemaName, sTableName)
	Dim oRs, sSql, bFirst, sConstraintName, sLinkId, sRefSchemaName, sRefTableName, sFullName
	sSql = Interpolate("SELECT c1.CONSTRAINT_NAME, c2.TABLE_SCHEMA, c2.TABLE_NAME " & _
		"FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS c1 " & _
		"INNER JOIN INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS rc " & _
		"ON c1.CONSTRAINT_SCHEMA = rc.CONSTRAINT_SCHEMA " & _
		"AND c1.CONSTRAINT_NAME = rc.CONSTRAINT_NAME " & _
		"INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS c2 " & _
		"ON rc.UNIQUE_CONSTRAINT_SCHEMA = c2.CONSTRAINT_SCHEMA " & _
		"AND rc.UNIQUE_CONSTRAINT_NAME = c2.CONSTRAINT_NAME " & _
		"WHERE c1.TABLE_SCHEMA = N'%1%' AND c1.TABLE_NAME = N'%2%' " & _
		"AND c1.CONSTRAINT_TYPE = 'FOREIGN KEY' " & _
		"ORDER BY c1.CONSTRAINT_NAME", _
		Array(SqlEscape(sSchemaName), SqlEscape(sTableName)) _
	)
	Set oRs = goDb.Execute(sSql)
	With oRs
		bFirst = True
		Do Until .EOF
			sConstraintName = .Fields("CONSTRAINT_NAME").Value
			sRefSchemaName = .Fields("TABLE_SCHEMA").Value
			sRefTableName = .Fields("TABLE_NAME").Value
			sFullName = Interpolate("%1%.%2%", Array(sRefSchemaName, sRefTableName))
			sLinkId = "tbl_" & LCase(Replace(sFullName, ".", "_"))
			If bFirst Then
				With goFormatter
					.WriteLine "<h2>Foreign Keys</h2>"
					.WriteLine "<ul>"
					.Indent
				End With
			End If
			goFormatter.WriteLine Interpolate( _
				"<li><a href=""#%1%"">%2%</a></li>", _
				Array(HtmlEscape(sLinkId), HtmlEscape(sConstraintName)) _
			)
			.MoveNext
			bFirst = False
		Loop
		If Not bFirst Then
			With goFormatter
				.Outdent
				.WriteLine "</ul>"
			End With
		End If
		.Close
	End With
	Set oRs = Nothing
End Sub


Function IsRefDataTable(sSchemaName, sTableName)
	Dim oRegExp
	Dim sFullName : sFullName = Interpolate("%1%.%2%", Array(sSchemaName, sTableName))

	Select Case gnRefData
	Case REFDATA_REGEXP
		Set oRegExp = New RegExp
		With oRegExp
			.Global = True
			.IgnoreCase = True
			.Pattern = gsRefDataRegexp
			IsRefDataTable = oRegExp.Test(sFullName)
		End With
		Set oRegExp = Nothing
	Case REFDATA_LIST
		IsRefDataTable = IsAmong(UCase(sFullName), gaRefDataList)
	Case Else
		IsRefDataTable = False
	End Select
End Function


''' Renders the contents of a database table as an HTML table
Sub WriteRefData(sSchemaName, sTableName)
	Dim oRs, sSql, bFirst, i
	
	sSql = Interpolate("SELECT * " & _
		"FROM %1%.%2% " & _
		"ORDER BY 1", _
		Array(SqlEscape(sSchemaName), SqlEscape(sTableName)) _ 
	)
	Set oRs = goDb.Execute(sSql)
	With oRs
		If Not (.BOF And .EOF) Then
			With goFormatter
				.WriteLine "<h2>Reference Data</h2>"
				.WriteLine "<table>"
				.Indent
			End With
			bFirst = True
			Do Until .EOF
				If bFirst Then
					With goFormatter
						.WriteLine "<thead>"
						.Indent
						.WriteLine "<tr>"
						.Indent
					End With
					For i = 0 To .Fields.Count - 1
						goFormatter.WriteLine Interpolate("<th>%1%</th>", Array(HtmlEscape(.Fields(i).Name)))
					Next
					With goFormatter
						.Outdent
						.WriteLine "</tr>"
						.Outdent
						.WriteLine "</thead>"
						.WriteLine "<tbody>"
						.Indent
					End With
				End If
				With goFormatter
					.WriteLine "<tr>"
					.Indent
				End With
				For i = 0 To .Fields.Count - 1
					If IsNull(.Fields(i).Value) Then
						goFormatter.WriteLine "<td><em>NULL</em></td>"
					ElseIf IsNumeric(.Fields(i).Value) Or IsDate(.Fields(i).Value) Then
						goFormatter.WriteLine Interpolate("<td class=""numeric"">%1%</td>", Array(HtmlEscape(.Fields(i).Value)))
					Else
						goFormatter.WriteLine Interpolate("<td>%1%</td>", Array(HtmlEscape(.Fields(i).Value)))
					End If
				Next
				With goFormatter
					.Outdent
					.WriteLine "</tr>"
				End With
				bFirst = False
				.MoveNext
			Loop
		End If
		.Close
		With goFormatter
			.Outdent
			.WriteLine "</tbody>"
			.Outdent
			.WriteLine "</table>"
		End With
	End With
	Set oRs = Nothing
End Sub


''' Returns True if a value v is in the array a, False otherwise
Function IsAmong(v, a)
	Dim i, lo, hi
	
	On Error Resume Next
	lo = LBound(a)
	If Err.Number <> 0 Then
		IsAmong = False
		Exit Function
	End If
	hi = UBound(a)
	If Err.Number <> 0 Then
		IsAmong = False
		Exit Function
	End If
	On Error GoTo 0
	
	For i = lo To hi
		If a(i) = v Then
			IsAmong = True
			Exit Function
		End If
	Next
	IsAmong = False
End Function


''' Given a string with %-delimited numbered placeholders and an array,
''' return a string replacing each placeholder with the corresponding 
''' array element
Function Interpolate(s, a)
	Dim p : p = Split(s, "%")
	Dim i
	Dim r : r = ""
	For i = LBound(p) To UBound(p)
		If i Mod 2 = 0 Then
			r = r & p(i)
		Else
			If p(i) = "" Then
				r = r & "%"
			ElseIf IsNumeric(p(i)) Then
				Dim v : v = CInt(p(i)) - 1
				If v >= LBound(a) And v <= UBound(a) Then
					r = r & CStr(a(v))
				Else
					r = r & "%" & p(i) & "%"
				End If
			Else
				r = r & "%" & p(i) & "%"
			End If
		End If
	Next
	Interpolate = r
End Function


''' Replaces characters significant in HTML with their entity references
Function HtmlEscape(s)
	On Error Resume Next
	HtmlEscape = Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
	On Error GoTo 0
End Function


''' Escapes quote characters
Function SqlEscape(s)
	SqlEscape = Replace(s, "'", "''")
End Function


''' Loads the list of reference data table names into an array
Sub LoadList(sFileName)
	Dim oFs : Set oFs = WScript.CreateObject("Scripting.FileSystemObject")
	Dim oFile, sLine
	Dim hi
	If Not oFs.FileExists(sFileName) Then
		WScript.Echo "Error: File """ & sFileName & """ not found."
		WScript.Quit 1
	End If
	Set oFile = oFs.OpenTextFile(sFileName)
	With oFile
		Do Until .AtEndOfStream
			sLine = Trim(.ReadLine)
			If Len(sLine) > 0 And Not IsAmong(UCase(sLine), gaRefDataList) Then
				hi = 0
				On Error Resume Next
				hi = UBound(gaRefDataList)
				On Error GoTo 0
				ReDim Preserve gaRefDataList(hi + 1)
				gaRefDataList(UBound(gaRefDataList)) = UCase(sLine)
			End If
		Loop
		.Close
	End With
	Set oFile = Nothing
	Set oFs = Nothing
End Sub


''' Displays usage information and quits with the specified error code
Sub Usage(nErrorCode)
	WScript.Echo "DumpSqlServerDbSchema - Generates a HTML containing all tables in a SQL Server database"
	WScript.Echo ""
	WScript.Echo "Usage:"
	WScript.Echo "DumpSqlServerDbSchema /c connection-string [/t] [/i title] [/rp regexp | /rf file-name ]"
	WScript.Echo ""
	WScript.Echo "connection-string: ADO connection string to connect to database (required)"
	WScript.Echo "/t: include table of contents (optional)"
	WScript.Echo "/t: set the title of the HTML page (optional)"
	WScript.Echo "regexp: regular expression that matches the names of tables containing reference data (optional)"
	WScript.Echo "file-name: file name containing a list of names of tables containing reference data (optional)"
	WScript.Echo
	WScript.Quit nErrorCode
End Sub

''' MAIN
Dim sConnectionString, sArg, sSwitch, ii
Dim nParseState : nParseState = 0

' Set defaults
gbWriteToc = False
gnRefData = REFDATA_NONE
gsTitle = "Database Schema"
Set goFormatter = New Formatter

' Parse command-line options
If WScript.Arguments.Count > 0 Then
	If WScript.Arguments.Count = 1 And WScript.Arguments(0) = "/?" Then
		Usage 0
	End If
	For ii = 0 To WScript.Arguments.Count - 1
		sArg = WScript.Arguments(ii)
		If nParseState = 0 Then
			If Left(sArg, 1) <> "/" Then
				WScript.Echo "Error: Unexpected value """ & sArg & """, expected switch"
				Usage 1
			ElseIf UCase(sArg) = "/T" Then
				gbWriteToc = True
			ElseIf IsAmong(UCase(sArg), Array("/C", "/RP", "/RF", "/I")) Then
				sSwitch = UCase(sArg)
				nParseState = 1
			Else
				WScript.Echo "Error: Unknown switch: """ & sArg & """"
				Usage 1
			End If
		Else
			If sSwitch = "/C" Then
				sConnectionString = sArg
			ElseIf sSwitch = "/RP" Then
				If gnRefData = REFDATA_LIST Then
					WScript.Echo "Warning: /RF and /RP switches are mutually exclusive"
				End If
				gnRefData = REFDATA_REGEXP
				gsRefDataRegexp = sArg
			ElseIf sSwitch = "/RF" Then
				If gnRefData = REFDATA_REGEXP Then
					WScript.Echo "Warning: /RF and /RP switches are mutually exclusive"
				End If
				gnRefData = REFDATA_LIST
				LoadList sArg
			ElseIf sSwitch = "/I" Then
				gsTitle = sArg
			End If
			nParseState = 0
		End If
	Next
Else
	Usage 1
End If

If sConnectionString = "" Then
	WScript.Echo "Error: A connection string must be specified"
	Usage 1
End If

Set goDb = WScript.CreateObject("ADODB.Connection")
goDb.ConnectionString = sConnectionString
goDb.Open
WriteHeader
WriteTables
WriteFooter
goDb.Close
Set goDb = Nothing
