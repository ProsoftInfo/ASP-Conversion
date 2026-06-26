<%
	'Program Name				:	PopulateMenu.asp
	'Module Name				:	Menu
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 18, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>

<%	' Function for populating the Menu for respective Applications
	'1.	Application Code

	Function PopulateMenu(iApplication)
		dim dcrs,dcrs1

		dim sEmployeeNo

		sEmployeeNo = Session("employeenumber")

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT PROCESSCODE,PROCESSNAME,ORDERNUMBER FROM VWUSERACTIVITY WHERE APPLICATIONCODE = " & iApplication & " AND INTERNALUSERID = " & sEmployeeNo & " ORDER BY ORDERNUMBER"
			.ActiveConnection = con
			.Open
		end with
		'Response.write dcrs.Source
		if not dcrs.EOF then
			do while not dcrs.EOF
		%>
			<b><%=trim(dcrs(1))%></b><br>
		<%
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT ACTIVITYCODE,ACTIVITYNAME,PROGRAMPATH,ActivityTemplateNo,ActivityTemplateName FROM VWUSERACTIVITY WHERE APPLICATIONCODE = " & iApplication & " AND INTERNALUSERID = " & sEmployeeNo & " AND PROCESSCODE = " & trim(dcrs(0)) & ""
					.ActiveConnection = con
					.Open
				end with

				if not dcrs1.EOF then
					do while not dcrs1.EOF
		%>
			<li><a href="../include/ShowActivity.asp?iApplication=<%=iApplication%>&iProcess=<%=trim(dcrs(0))%>&iActivity=<%=trim(dcrs1(0))%>&sPath=<%=trim(dcrs1(2))%>&iActTempNo=<%=trim(dcrs1(3)) %>" target="bodyFrame"><%=trim(dcrs1(4))%></a><br>

		<%
					dcrs1.MoveNext
					loop
				end if
				dcrs1.Close
		%>
			<hr>

		<%
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	End Function
%>