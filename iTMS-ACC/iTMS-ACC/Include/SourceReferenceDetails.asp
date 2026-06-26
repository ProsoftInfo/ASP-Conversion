<%
	'Program Name				:	SourceReferenceDetails.asp
	'Module Name				:	Inventory 
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 04, 2004
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
	'Description				:	Used to display the Header information from various Applications
%>

<%	
	' Function for displaying the Work Order details part

	'1.	Source Reference Number
	'2. From which Application 
	
	Function DisplayWODetail(sRefNumber,sWho)
		dim dcrs,sWCName,sMCName,iCount
		
		sWCName = "-"
		iCount = 0
		
		set dcrs = Server.CreateObject("ADODB.RECORDSET")
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			' Inventory Application
			if sWho = "INV" then
				.Source = "SELECT DISTINCT WORKCENTERNAME FROM VWMRWODETAILS WHERE MRSNUMBER = " & sRefNumber & ""
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		iCount = dcrs.RecordCount
		
		if not dcrs.EOF then
			sWCName = trim(dcrs(0))
		end if
		dcrs.Close
		
		if cint(iCount) > 0 then
%>
		<tr>
			<td class="FieldCell">Work Center</td>
			<td class="FieldCellSub">
				<span class="DataOnly"><%=sWCName%>&nbsp;</span>
			</td>
		</tr>

<%		
		end if
		iCount = 0

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			' Inventory Application
			if sWho = "INV" then
				.Source = "SELECT DISTINCT MACHINECENTERNAME FROM VWMRWODETAILS WHERE MRSNUMBER = " & sRefNumber & ""
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		iCount = dcrs.RecordCount
		
		if not dcrs.EOF then
			do while not dcrs.EOF
				sMCName = sMCName & "," & trim(dcrs(0))
			dcrs.MoveNext
			loop
			sMCName = mid(sMCName,2)
		else
			sMCName = "-"
		end if
		dcrs.Close
		if cint(iCount) > 0 then
%>
		<tr >
			<td class="FieldCell">Machine Center(s)</td>
			<td class="FieldCellSub" colspan=4>
				<span class="DataOnly"><%=sMCName%>&nbsp;</span>
			</td>
		</tr>
<%
		end if
		iCount = 0
		
	End Function 
%>

<%	
	' Function for displaying the MRP details part

	'1.	Source Reference Number
	'2. From which Application 
	
	Function DisplayMRPDetail(sRefNumber,sWho)
		dim dcrs,sMRPNo,sPRDONo,sMixCode,iCount,sTemp
		
		sMRPNo = "-"
		iCount = 0
		
		set dcrs = Server.CreateObject("ADODB.RECORDSET")
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			' Inventory Application
			if sWho = "INV" then
				.Source = "SELECT DISTINCT MRPNO FROM VWMRMRPDETAILS WHERE MRSNUMBER = " & sRefNumber & ""
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		iCount = dcrs.RecordCount
		
		if not dcrs.EOF then
			sMRPNo = trim(dcrs(0))
		end if
		dcrs.Close
		
		if cint(iCount) > 0 then
%>
		<tr>
			<td class="FieldCell">MRP Number</td>
			<td class="FieldCellSub">
				<span class="DataOnly"><%=sMRPNo%>&nbsp;</span>
			</td>
		</tr>

<%		
		end if
		iCount = 0

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			' Inventory Application
			if sWho = "INV" then
				.Source = "SELECT DISTINCT PRODUCTIONORDERNO,MIXCODE FROM VWMRMRPDETAILS WHERE MRSNUMBER = " & sRefNumber & ""
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		
		iCount = dcrs.RecordCount
		
		if not dcrs.EOF then
			do while not dcrs.EOF
				sPRDONo = trim(dcrs(0))
				sMixCode = trim(dcrs(1))
				sTemp = sTemp & "," & sPRDONo & " - " & sMixCode
			dcrs.MoveNext
			loop
			sTemp = mid(sTemp,2)
		else
			sTemp = "-"
		end if
		dcrs.Close
		if cint(iCount) > 0 then
%>
		<tr >
			<td class="FieldCell">Production Order - MixCode</td>
			<td class="FieldCellSub" colspan=4>
				<span class="DataOnly"><%=sTemp%>&nbsp;</span>
			</td>
		</tr>
<%
		end if
		iCount = 0
		
	End Function 
%>