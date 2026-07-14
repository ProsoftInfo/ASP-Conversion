<%
	'Program Name				:	GetSerialDetail.asp
	'Module Name				:	Inventory (Serial Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 19, 2003
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

<%	' Function for returning the pack number for the Serial Number
	Function GetSerialDetail(iSerialNo)
		dim dcrs,sPackNo

		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PACKINGNUMBER FROM INV_T_LocationLot WHERE SERIALNUMBER = " & iSerialNo & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sPackNo = trim(dcrs(0))
		else
			sPackNo = "N/A"
		end if
		dcrs.close
		
		GetSerialDetail = sPackNo
		
	End Function
%>
