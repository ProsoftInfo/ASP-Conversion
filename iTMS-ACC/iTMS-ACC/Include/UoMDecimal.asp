<%
	'Program Name				:	UoMDecimal.asp
	'Module Name				:	
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 27, 2003
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
<%	
	' Function for returning whether for UoM decimal allowed or not

	'1.	UoM 
	
	Function UoMDecimal(sUoMCode)
		if sUoMCode = "" then
			UoMDecimal = "Err : UoM Code not passed"
			exit function
		end if
		
		dim dcrs,sAllowed
		
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
			
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DECIMALALLOWED FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = '" & trim(sUoMCode) & "'"
			.ActiveConnection = con
			.Open
		end with
		if not dcrs.EOF then
			sAllowed = trim(dcrs(0))
			UoMDecimal = sAllowed
		else
			UoMDecimal = "Err : No UoM Code created"
		end if
		dcrs.Close
		
	End Function
%>
