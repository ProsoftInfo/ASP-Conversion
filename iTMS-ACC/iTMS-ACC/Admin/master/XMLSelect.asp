<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLSelect.asp
	'Module Name				:	Admin (Activity Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	Decembe 02, 2003
	'Modified By				:	TAJUDEEN S
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

<!--#include virtual="/include/DatabaseConnection.asp"-->

<%
	dim dcrs,dcrs1,dcrs2,OutData,sProcess,Root,newElem,newElem1,sWho,sPractice
	dim sUser,sRole,sOrgId,sItemTypeId,Node,sProcessId, sSql,iApplication
	dim iProcess,iActivity,sActions,sTemp,iCtr,sCallFrom,nProcessCode,sRoleName
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	
	sWho = Request("sWho")
	'sWho = "UANM"
	' Practice
	if sWho = "PR" then
		sProcess = Request("sProcess")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PROCESSCODE,PROCESSNAME,ORDERNUMBER FROM MS_APPLICATIONPROCESS WHERE APPLICATIONCODE = " & sProcess & " ORDER BY ORDERNUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		Set Root = OutData.createElement("PROCESS")												
		OutData.appendChild Root

		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("ITEMS")
				newElem.setAttribute "PRCode", trim(dcrs(0))
				newElem.setAttribute "PRName", trim(dcrs(1))
				newElem.setAttribute "ORDERNO", trim(dcrs(2)) 'added by Tajudeen
				Root.appendChild newElem
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	' User Role Allocation
	elseif sWho = "UR" then
		sUser = Request("sUser")

		Set Root = OutData.createElement("ROLES")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES WHERE ROLEID NOT IN (SELECT ROLEID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & ") ORDER BY ROLEDESCRIPTION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("TOMAP")
				newElem.setAttribute "RLCode", trim(dcrs(0))
				newElem.setAttribute "RLName", trim(dcrs(1))
				Root.appendChild newElem
			dcrs.MoveNext
			loop
		else
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES ORDER BY ROLEDESCRIPTION"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				do while not dcrs1.EOF
					Set newElem = OutData.createElement("TOMAP")
					newElem.setAttribute "RLCode", trim(dcrs1(0))
					newElem.setAttribute "RLName", trim(dcrs1(1))
					Root.appendChild newElem
				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close
		end if
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES WHERE ROLEID IN (SELECT ROLEID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & ") ORDER BY ROLEDESCRIPTION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("MAPPED")
				newElem.setAttribute "RLCode", trim(dcrs(0))
				newElem.setAttribute "RLName", trim(dcrs(1))
				Root.appendChild newElem
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	' Role Activity Allocation
	elseif sWho = "RA" then
		sRole = Request("sRole")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM VWROLEACTIVITY WHERE ROLEID NOT IN (SELECT ROLEID FROM MS_ROLEACTIVITY WHERE ROLEID = " & sRole & ") ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if 1 = 2 then
		'if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("TOMAP")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY ORDER BY 1,2,3"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				do while not dcrs1.EOF
					Set newElem = OutData.createElement("TOMAP")
					newElem.setAttribute "APPCode", trim(dcrs1(0))
					newElem.setAttribute "PRCode", trim(dcrs1(1))
					newElem.setAttribute "ACCode", trim(dcrs1(2))
					newElem.setAttribute "APPName", trim(dcrs1(3))
					newElem.setAttribute "PAName", trim(dcrs1(4))
					newElem.setAttribute "ACName", trim(dcrs1(5))
					Root.appendChild newElem
				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close
		'end if
		'dcrs.Close

	' User Activity Allocation
	elseif sWho = "UA" then
		sUser = Request("sUser")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM VWROLEACTIVITY WHERE ROLEID IN (SELECT ROLEID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & ") ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("TOMAP")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	' User Activity Amendment added by Tajudeen
	elseif sWho = "UAA" then
	
		sProcess = Request("sProcess")
		sPractice = Request("sPractice")
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,ACTIVITYNAME,PROGAMPATH FROM MS_APPLICATIONACTIVITY WHERE APPLICATIONCODE=" & sProcess & " AND PROCESSCODE=" & sPractice
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		while not dcrs.EOF 
			Set newElem = OutData.createElement("DETAILS")
			newElem.setAttribute "APPCode", trim(dcrs(0))
			newElem.setAttribute "PRCode", trim(dcrs(1))
			newElem.setAttribute "ACCCode", trim(dcrs(2))
			newElem.setAttribute "ACCName", trim(dcrs(3))
			newElem.setAttribute "ProgramPath", trim(dcrs(4))
			Root.appendChild newElem
			dcrs.MoveNext
		wend
		dcrs.Close 
		
	' User Activity DeAllocation added by Tajudeen
	'Assigned Activities
	elseif sWho = "UAD" then
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		sItemTypeId = replace(Request("sItemTypeId"),",","','")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE IN ('" & sOrgId & "') AND ITEMTYPEID IN ('" & sItemTypeId & "') ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	' User Activity DeAllocation added by Tajudeen
	'UnAssigned Activities
	elseif sWho = "UADU" then 'User Activities Deallocation Unassigned
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		sItemTypeId = replace(Request("sItemTypeId"),",","','")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM VWROLEACTIVITY a WHERE NOT EXISTS (SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIVITY b WHERE a.APPLICATIONCODE = b.APPLICATIONCODE and a.PROCESSCODE = b.PROCESSCODE and a.ACTIVITYCODE = b.ACTIVITYCODE AND INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE IN ('" & sOrgId & "') AND ITEMTYPEID IN ('" & sItemTypeId & "')) AND ROLEID IN (SELECT ROLEID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & ") ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	'User Activity View added by Tajudeen
	elseif sWho = "UAV" then
		sUser = Request("sUser")
		sOrgId = Request("sOrgId")
		sProcessId = Request("sProcessId")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root
		
		if trim(sProcessId) = "ALL" then
			sSql= "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE = '" & sOrgId & "' ORDER BY 1,2,3"
		else
			sSql= "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE = '" & sOrgId & "' AND APPLICATIONCODE = " & sProcessId & " ORDER BY 1,2,3"
		end if
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					
					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT DISTINCT ORGANISATIONCODE FROM VWUSERACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & " AND ORGANISATIONCODE <> '" & sOrgId & "'"
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing

					do while not dcrs2.EOF 
						sTemp = sTemp & "," & trim(dcrs2(0))
						dcrs2.MoveNext 
					loop
					dcrs2.Close

					newElem.setAttribute "ForUnit", DisplayOrganization(trim(mid(sTemp,2)))
					Root.appendChild newElem
					sTemp = ""
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	'Role Activity View added by Tajudeen
	elseif sWho = "RAV" then
		sRole = Request("sRole")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_ROLEACTIVITY WHERE ROLEID = " & sRole  & " ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close

	'User Role View added by Tajudeen
	elseif sWho = "URV" then
		sUser = Request("sUser")

		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root
	
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES WHERE ROLEID IN (SELECT ROLEID FROM MS_USERROLES WHERE INTERNALUSERID = " & sUser & ") ORDER BY ROLEDESCRIPTION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		if not dcrs1.EOF then
			do while not dcrs1.EOF
				Set newElem = OutData.createElement("ROLES")
				newElem.setAttribute "RLCODE", trim(dcrs1(0))
				newElem.setAttribute "RLNAME", trim(dcrs1(1))
				Root.appendChild newElem
			dcrs1.MoveNext
			loop
		end if
		dcrs1.Close
	
	'User Actions by Tajudeen
	elseif sWho = "USERACTION" then
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		iApplication = Request("iApplication")
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE IN ('" & sOrgId & "') AND APPLICATIONCODE = " & iApplication & " ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then

					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	
	'User Allocated Activity added by Tajudeen
	elseif sWho = "USERACTIVITYALLOCATED" then
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		iApplication = Request("iApplication")
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIONS WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE IN ('" & sOrgId & "') AND APPLICATIONCODE = " & iApplication & " ORDER BY 1,2,3"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					
					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "APPCode", trim(dcrs(0))
					newElem.setAttribute "PRCode", trim(dcrs(1))
					newElem.setAttribute "ACCode", trim(dcrs(2))
					newElem.setAttribute "APPName", trim(dcrs1(0))
					newElem.setAttribute "PAName", trim(dcrs1(1))
					newElem.setAttribute "ACName", trim(dcrs1(2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	
	'User Allocated Actions added by Tajudeen
	elseif sWho = "USERACTIONALLOCATED" then
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		iApplication = Request("iApplication")
		iProcess = Request("iProcess")
		iActivity = Request("iActivity")
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root
		set Node = OutData.createElement("MAPPED")
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ACTIONS FROM MS_USERACTIONS WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & iApplication & " AND PROCESSCODE = " & iProcess & " AND ACTIVITYCODE = " & iActivity & " AND ORGANISATIONCODE = '" & sOrgId & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.EOF
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ACTIONNAME FROM MS_APPLICATIONACTION WHERE ACTION = '" & trim(dcrs(0)) & "'"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				Set newElem = OutData.createElement("ACTIONS")
				newElem.setAttribute "CODE", trim(dcrs(0))
				newElem.setAttribute "NAME", trim(dcrs1(0))
				Node.appendChild newElem
			end if
			dcrs1.Close 
			
			dcrs.MoveNext
		loop
		dcrs.Close
		Root.appendChild Node
		
		set Node = OutData.createElement("TOBEMAPPED")
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ACTION,ACTIONNAME FROM MS_APPLICATIONACTION WHERE ACTION NOT IN (SELECT ACTIONS FROM MS_USERACTIONS WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & iApplication & " AND PROCESSCODE = " & iProcess & " AND ACTIVITYCODE = " & iActivity & " AND ORGANISATIONCODE = '" & sOrgId & "')"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.EOF
			Set newElem = OutData.createElement("ACTIONS")
			newElem.setAttribute "CODE", trim(dcrs(0))
			newElem.setAttribute "NAME", trim(dcrs(1))
			Node.appendChild newElem
			dcrs.MoveNext
		loop
		dcrs.Close
		Root.appendChild Node
	
	'User Action View added by Tajudeen
	elseif sWho = "USERACTIONVIEW" then
		sUser = Request("sUser")
		sOrgId =replace(Request("sOrgId"),",","','")
		iApplication = Request("iApplication")
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT PROCESSCODE,ACTIVITYCODE FROM MS_USERACTIONS WHERE INTERNALUSERID = " & sUser & " AND ORGANISATIONCODE IN ('" & sOrgId & "') AND APPLICATIONCODE = " & iApplication & " ORDER BY 1,2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & iApplication & " AND PROCESSCODE = " & trim(dcrs(0)) & " AND ACTIVITYCODE = " & trim(dcrs(1)) 
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					
					Set newElem = OutData.createElement("DETAILS")
					newElem.setAttribute "PRACTICECODE", trim(dcrs(0))
					newElem.setAttribute "ACTIVITYCODE", trim(dcrs(1))
					newElem.setAttribute "PRACTICENAME", trim(dcrs1(0))
					newElem.setAttribute "ACTIVITYNAME", trim(dcrs1(1))

					sActions = ""
					with dcrs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ACTIONNAME FROM MS_APPLICATIONACTION WHERE ACTION IN (SELECT ACTIONS FROM MS_USERACTIONS WHERE INTERNALUSERID = " & sUser & " AND APPLICATIONCODE = " & iApplication & " AND PROCESSCODE = " & trim(dcrs(0)) & " AND ACTIVITYCODE = " & trim(dcrs(1)) & " AND ORGANISATIONCODE = '" & sOrgId & "')"
						.ActiveConnection = con
						.Open
					end with
					set dcrs2.ActiveConnection = nothing

					do while not dcrs2.EOF
						sActions = sActions & ", " & trim(dcrs2(0))
						dcrs2.MoveNext 
					loop
					dcrs2.Close 
					
					newElem.setAttribute "ACTIONNAME", UCase(mid(sActions,2))
					Root.appendChild newElem
				end if
				dcrs1.Close
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	'
	elseif sWho = "USERACTIONS" then
		iApplication = Request("iApplication")
		
		Set Root = OutData.createElement("ACTIONS")												
		OutData.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ACTIONNAME, ACTION FROM MS_APPLICATIONACTION WHERE APPLICATIONCODE = " & iApplication & " ORDER BY 1,2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("DETAILS")
				newElem.setAttribute "ACTIONNAME", trim(dcrs(0))
				newElem.setAttribute "ACTION", trim(dcrs(1))

				Root.appendChild newElem
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	
	Elseif sWho = "PPA" then	'Added By UmaMaheswari S on 09 December 2010, For Role Activity Mapping
		Dim nAppCode,nRoleID
		
		sTemp = Request("sProcess")
		
		nProcessCode = Split(sTemp,":")(0)
		nAppCode     = Split(sTemp,":")(1)
		nRoleID      = Split(sTemp,":")(2)
		
		Set Root = OutData.createElement("ACTIVITY")
		OutData.appendChild Root
		
		sSql = " SELECT DISTINCT PROCESSCODE,ACTIVITYCODE,ACTIVITYNAME FROM Ms_ApplicationActivity "&_
			   " WHERE APPLICATIONCODE = " & nAppCode & "  AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE NOT in "&_
			   " (SELECT DISTINCT ACTIVITYCODE FROM MS_ROLEACTIVITY WHERE PROCESSCODE = " & nProcessCode & " AND APPLICATIONCODE = " & nAppCode & " AND ROLEID="& nRoleID &") "&_
			   " ORDER BY PROCESSCODE "
			   
		sSql = " SELECT DISTINCT A.PROCESSCODE,A.ACTIVITYCODE,A.ACTIVITYNAME FROM Ms_ApplicationActivity A,Ms_ApplicationActivityTemplates T "&_
		       " WHERE A.ApplicationCode = T.ApplicationCode and A.ProcessCode = T.ProcessCode and A.ActivityCode = T.ActivityCode "&_
		       " and A.APPLICATIONCODE = "& nAppCode &" AND A.PROCESSCODE = "& nProcessCode &" and Cast(T.ActivityCode as varchar) +':'+ Cast(T.ActivityTemplateNo as varchar)"&_
		       " Not in (Select distinct Cast(ActivityCode as varchar)+':'+ Cast(ActivityTemplateNo as Varchar) "&_
		       " FROM MS_ROLEACTIVITY WHERE PROCESSCODE = "& nProcessCode &" AND APPLICATIONCODE = "& nAppCode &" AND ROLEID="& nRoleID &") "&_
		       " Group By A.ProcessCode,A.ActivityCode,ActivityName ORDER BY A.PROCESSCODE "
		
		'Response.write sSql	  
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		
		set dcrs1.ActiveConnection = nothing
		'Root.setAttribute "QUERY", dcrs1.Source
		'Root.setAttribute "DATA", sProcess  
		If not dcrs1.EOF then
			Do while Not dcrs1.EOF 
				
				Set newElem = OutData.createElement("TOMAP")
				
				newElem.setAttribute "PRCode", trim(dcrs1(0))
				newElem.setAttribute "ACCode", trim(dcrs1(1))
				newElem.setAttribute "ACName", trim(dcrs1(2))
				Root.appendChild newElem
				
				dcrs1.MoveNext 
			Loop
		End if
		dcrs1.Close
	'OutData.save server.MapPath("../Temp/NonMAP.xml")	
		
	Elseif sWho = "PPAONLOAD" Then	''Added By UmaMaheswari S on 09 December 2010, For Role Activity Mapping
		Dim nActCode,sActName
		
		'sRole = Request("RoleID")
		sTemp  = Request("sPassData")
		sCallFrom = Split(sTemp,":")(0)
		sRole = Split(sTemp,":")(1)
		nProcessCode = Split(sTemp,":")(2)
		
		If trim(sCallFrom) = "FROMPRACTICE" Then
			nPracticeCode = Split(sTemp,":")(3)
		End IF
		
		'sCallFrom = "FROMPROCESS"	'FROMPRACTICE
		'sRole = 1
		'nProcessCode = 0
		
		Set Root = OutData.createElement("ROOT")
		OutData.appendChild Root
		
		with dcrs 
			.CursorLocation = 3
			.CursorType = 3
			'If sCallFrom = "FROMPRACTICE" Then
			'	.Source = "select Distinct R.APPLICATIONCODE,A.APPLICATIONNAME from MS_ROLEACTIVITY R,Ms_Applications A where R.APPLICATIONCODE  = A.APPLICATIONCODE  and R.ROLEID = "& sRole &" "
			'Else
				.Source = "select Distinct R.APPLICATIONCODE,A.APPLICATIONNAME from MS_ROLEACTIVITY R,Ms_Applications A where R.APPLICATIONCODE  = A.APPLICATIONCODE  and R.ROLEID = "& sRole &" and R.APPLICATIONCODE = "& nProcessCode &" "
			'End IF
			.ActiveConnection = con
			.Open
		end with
		
		set dcrs.ActiveConnection = nothing
		'Response.Write "<p>Query1="&dcrs.source & "<BR><BR>"
		'Root.setAttribute "sTemp",sTemp
		'Root.setAttribute "Query1",dcrs.source
		If not dcrs.EOF then
			
			with dcrs1 
				.CursorLocation = 3
				.CursorType = 3
				If sCallFrom = "FROMPRACTICE" Then
					'.Source = "select Distinct PROCESSCODE,PROCESSNAME,ORDERNUMBER from Ms_ApplicationProcess where APPLICATIONCODE = "& dcrs(0)&" and PROCESSCODE = "& nProcessCode &" Order by ORDERNUMBER"
					'.Source = "select Distinct PROCESSCODE,PROCESSNAME,ORDERNUMBER from Ms_ApplicationProcess where APPLICATIONCODE = "& dcrs(0)&" and PROCESSCODE = "& nPracticeCode &" Order by ORDERNUMBER"
					'.Source = "select Distinct PROCESSCODE,PROCESSNAME from VwROLEACTIVITY where APPLICATIONCODE = "& dcrs(0)&" and PROCESSCODE = "& nPracticeCode &" and RoleID = "& sRole &" "
					.Source = "select Distinct A.PROCESSCODE,A.PROCESSNAME from Ms_ApplicationProcess A, MS_ROLEACTIVITY R  where R.APPLICATIONCODE = "& dcrs(0)&" and R.PROCESSCODE = "& nPracticeCode &" AND R.RoleID = "& sRole &"  AND A.APPLICATIONCODE = R.APPLICATIONCODE AND A.PROCESSCODE = R.PROCESSCODE"
				Else
					'.Source = "select Distinct PROCESSCODE,PROCESSNAME from Ms_ApplicationProcess where APPLICATIONCODE = "& dcrs(0)&" "
					'.Source = "select Distinct PROCESSCODE,PROCESSNAME from VwROLEACTIVITY where APPLICATIONCODE = "& dcrs(0)&"  and RoleID = "& sRole &" "
					.Source = "select Distinct A.PROCESSCODE,A.PROCESSNAME from Ms_ApplicationProcess A, MS_ROLEACTIVITY R  where R.APPLICATIONCODE = "& dcrs(0)&" AND R.RoleID = "& sRole &"  AND A.APPLICATIONCODE = R.APPLICATIONCODE AND A.PROCESSCODE = R.PROCESSCODE"
				End if
				.ActiveConnection = con
				.Open
			end with
			'Response.Write "<p>Query1="&dcrs1.source & "<BR><BR>"
			'Root.setAttribute "Query2",dcrs1.Source
			set dcrs1.ActiveConnection = nothing
			Do while Not dcrs1.EOF
				
				iCtr = iCtr + 1
				
				Set newElem = OutData.createElement("ACTIVITYMAPPING")	
				newElem.setAttribute "CTR",iCtr
				newElem.setAttribute "APPCODE",dcrs(0)
				newElem.setAttribute "APPNAME",dcrs(1)
				newElem.setAttribute "PROCESSCODE",dcrs1(0)
				newElem.setAttribute "PROCESSNAME",dcrs1(1)
				Root.appendChild newElem
				
				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "select DISTINCT ACTIVITYCODE,ACTIVITYNAME from Ms_ApplicationActivity WHERE STATUS = 'A' AND APPLICATIONCODE = "& dcrs(0) &" and PROCESSCODE = "& dcrs1(0) &" AND ACTIVITYCODE in (SELECT distinct ACTIVITYCODE FROM MS_ROLEACTIVITY WHERE PROCESSCODE = "& dcrs1(0) &" AND APPLICATIONCODE = "& dcrs(0) &" AND ROLEID = "& sRole &") "
					.Source = "Select T.ActivityCode,ActivityName,ActivityTemplateNo,ActivityTemplateName from Ms_ApplicationActivity A,Ms_ApplicationActivityTemplates T where ActivityTemplateNo in ( "&_
					          " SELECT distinct IsNull(ActivityTemplateNo,1) FROM MS_ROLEACTIVITY WHERE PROCESSCODE = "& dcrs1(0) &" AND APPLICATIONCODE = "& dcrs(0) &" AND ROLEID = "& sRole &") "&_
					          " and A.ApplicationCode = T.ApplicationCode and A.ProcessCode = T.ProcessCode and A.ActivityCode = T.ActivityCode and A.Status = 'A' and T.Status='A' and T.ApplicationCode = "& dcrs(0) &_
					          " and T.ProcessCode="& dcrs1(0) &" and cast(T.ActivityCode as varchar)+':'+cast(T.ActivityTemplateNo as varchar) in ( SELECT distinct cast(ActivityCode as varchar)+':'+cast(ActivityTemplateNo as varchar) FROM MS_ROLEACTIVITY WHERE PROCESSCODE = "& dcrs1(0) &" AND APPLICATIONCODE = "& dcrs(0) &" AND ROLEID = "& sRole &") "
					.ActiveConnection = con
					.Open
				end with
				'Response.Write "<p>Query3="&dcrs2.source & "<BR><BR>"
				set dcrs2.ActiveConnection = nothing
				
				nActCode = ""
				sActName = ""
				Do while Not dcrs2.EOF
					Set newElem1 = OutData.createElement("ACTIVITY")	
					newElem1.setAttribute "ACTCTR",iCtr
					newElem1.setAttribute "CODE",dcrs2(0)
					newElem1.setAttribute "NAME",dcrs2(1)
					newElem1.setAttribute "TCODE",dcrs2(2)
					newElem1.setAttribute "TNAME",dcrs2(3)
					newElem.appendChild newElem1
					
					'nActCode = nActCode & "," & dcrs2(0)
					'sActName = sActName & "," & dcrs2(1)
					dcrs2.MoveNext 
				Loop
				dcrs2.Close 
				
				'If nActCode <> "" Then nActCode = mid(nActCode,2)
				'If sActName <> "" Then sActName = mid(sActName,2)
				
				'newElem.setAttribute "ACTCODE",nActCode
				'newElem.setAttribute "ACTNAME",sActName
				'Root.appendChild newElem
				
				dcrs1.MoveNext 
			Loop
			dcrs1.Close 
			
		End IF
		dcrs.close 
	OutData.save server.MapPath("../Temp/AppRole.xml")	
	Elseif sWho = "ACTONLOAD" Then	''Added By UmaMaheswari S on 15 December 2010, For Application Activity
		Dim nPracticeCode
		
		sTemp  = Request("sPassData")
		sCallFrom = Split(sTemp,":")(0)
		nProcessCode = Split(sTemp,":")(1)
		nPracticeCode = Split(sTemp,":")(2)
		
		Set Root = OutData.createElement("ROOT")
		OutData.appendChild Root
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'If sCallFrom = "PROCESS" Then
				.Source = "Select DISTINCT ProcessCode,ProcessName,OrderNumber From Ms_ApplicationProcess where ApplicationCode = "& nProcessCode &" order by OrderNumber"
			'Else
			'	.Source = "Select ProcessCode,ProcessName,OrderNumber From Ms_ApplicationProcess where ApplicationCode = "& nProcessCode &" and ProcessCode="& nPracticeCode &" Order by OrderNumber"
			'End IF
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			Do While Not dcrs.EOF
				iCtr = iCtr + 1
				
				Set newElem = OutData.createElement("DETAILS")	
				newElem.setAttribute "CTR",iCtr
				newElem.setAttribute "PROCESSCODE",dcrs(0)
				newElem.setAttribute "PROCESSNAME",dcrs(1)
				newElem.setAttribute "ORDERNUMBER",dcrs(2)
				Root.appendChild newElem
				
				dcrs.movenext
			Loop
		End IF
		dcrs.close
		
	Elseif sWho = "UAMAP" Then 'Mapped Activity selection 
		
		Dim nPrevRoleID 
		
	 	'sUser = Request("sUser")
	 	sTemp  = Request("sPassData")
	 	sUser  = Split(sTemp,":")(0)
	 	sRole  = Split(sTemp,":")(1)
	 	
	 	'If sRole = "ALL" Then sRole = 1
		
	 	Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root
		
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			If sRole = "ALL" Then
				.Source = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM Ms_Roles "
			Else
				.Source = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM Ms_Roles WHERE ROLEID ="& sRole &" Order by ROLEID"
			End IF
			.ActiveConnection = con
			.Open
		end with
		
		set dcrs2.ActiveConnection = nothing
		
		nPrevRoleID = cdbl("0")
		
		If Not dcrs2.EOF Then
		Do while Not dcrs2.EOF 
			
			sRoleName = dcrs2(1)
		
		'Root.setAttribute "ROLENAME",sRoleName
		Root.setAttribute "DATA",dcrs2.Source 
		
		
		'sSql = " select DISTINCT V.APPLICATIONCODE,V.PROCESSCODE,V.ACTIVITYCODE,V.PROCESSNAME,V.ACTIVITYNAME "&_
		'	   " from VwUSERACTIVITY V, Ms_RoleActivity R WHERE V.INTERNALUSERID ="& sUser &" and R.RoleID = "& dcrs2(0) &"  "&_
		'	   " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE "
		
		sSql = " select DISTINCT V.APPLICATIONCODE,V.PROCESSCODE,V.ACTIVITYCODE,V.PROCESSNAME,V.ACTIVITYNAME,A.APPLICATIONNAME "&_
			   " from VwUSERACTIVITY V, Ms_RoleActivity R,Ms_Applications A  WHERE V.INTERNALUSERID ="& sUser &" and R.RoleID = "& dcrs2(0) &"  "&_
			   " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE "&_
			   " and R.APPLICATIONCODE = A.APPLICATIONCODE ORDER BY 1,2,3"
		
		sSql = " select DISTINCT V.APPLICATIONCODE,V.PROCESSCODE,V.ACTIVITYCODE,V.PROCESSNAME,V.ACTIVITYNAME,A.APPLICATIONNAME "&_
			   " from VwUSERACTIVITY V, Ms_UserActivity R,Ms_Applications A ,MS_USERROLES UR  WHERE V.INTERNALUSERID ="& sUser &" and UR.RoleID = "& dcrs2(0) &"  "&_
			   " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE "&_
			   " and R.APPLICATIONCODE = A.APPLICATIONCODE AND V.INTERNALUSERID = UR.INTERNALUSERID ORDER BY 1,2,3"
		
		'sSql = " select DISTINCT V.APPLICATIONCODE,V.PROCESSCODE,V.ACTIVITYCODE,V.PROCESSNAME,V.ACTIVITYNAME,A.APPLICATIONNAME "&_
		'	   " from VwUSERACTIVITY V, Ms_UserActivity R,Ms_Applications A ,MS_USERROLES UR,Ms_RoleActivity RA   WHERE V.INTERNALUSERID ="& sUser &" and UR.RoleID = "& dcrs2(0) &"  "&_
		'	   " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE "&_
		'	   " and R.APPLICATIONCODE = A.APPLICATIONCODE AND V.INTERNALUSERID = UR.INTERNALUSERID "&_
		'	   "AND RA.ROLEID = UR.ROLEID AND RA.APPLICATIONCODE = R.APPLICATIONCODE  and RA.PROCESSCODE = R.PROCESSCODE and RA.ACTIVITYCODE = R.ACTIVITYCODE  ORDER BY 1,2,3"
		
		sSql = " SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_ROLEACTIVITY "&_
			   " WHERE RoleID = "& dcrs2(0) &" AND Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) IN "&_
               " (select DISTINCT (Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) ) "&_
		       " FROM MS_USERACTIVITY WHERE INTERNALUSERID = "& sUser &" ) "
			   
		'Root.setattribute "Query",sSql
		
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE,PROCESSNAME,ACTIVITYNAME FROM VwUSERACTIVITY WHERE INTERNALUSERID = "& sUser &" "
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		
		set dcrs1.ActiveConnection = nothing
		
		if not dcrs1.EOF then
			do while not dcrs1.EOF 
			
				with dcrs 
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT DISTINCT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs1(0)) & " AND PROCESSCODE = " & trim(dcrs1(1)) & " AND ACTIVITYCODE = " & trim(dcrs1(2)) & ""
					.source = " SELECT DISTINCT V.APPLICATIONNAME,V.PROCESSNAME,V.ACTIVITYNAME FROM VWACTIVITY V,MS_ROLEACTIVITY R WHERE V.APPLICATIONCODE = " & trim(dcrs1(0)) & " and V.PROCESSCODE = " & trim(dcrs1(1)) & " and V.ACTIVITYCODE = " & trim(dcrs1(2)) & " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE =  R.ACTIVITYCODE AND R.ROLEID = "& dcrs2(0) &" "
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				
				if not dcrs.EOF then
				
					If cdbl(nPrevRoleID) = cdbl("0") or cdbl(nPrevRoleID) <> cdbl(dcrs2(0)) Then
						nPrevRoleID = dcrs2(0)
						Set newElem = OutData.createElement("ROLE")
						newElem.setAttribute "ROLEID", dcrs2(0)
						newElem.setAttribute "ROLENAME", sRoleName
						Root.appendchild newElem 
					End IF
						
					Set newElem1 = OutData.createElement("TOMAP")
					newElem1.setAttribute "APPCode", trim(dcrs1(0))
					newElem1.setAttribute "PRCode", trim(dcrs1(1))
					newElem1.setAttribute "ACCode", trim(dcrs1(2))
					newElem1.setAttribute "APPName", trim(dcrs(0))'trim(dcrs1(5))
					newElem1.setAttribute "PAName", trim(dcrs(1)) 'trim(dcrs1(3))
					newElem1.setAttribute "ACName", trim(dcrs(2)) 'trim(dcrs1(4))
					newElem.appendChild newElem1
				End IF
				dcrs.Close 
				
			dcrs1.MoveNext 
			Loop
		end if
		dcrs1.Close
				
				dcrs2.MoveNext 
			Loop
		End IF
		dcrs2.Close 
		
		'OutData.save server.MapPath("../Temp/MappedActvities.xml")
	
	Elseif sWho = "UANM" then	'User Un Mapped Activity Selection
	
		
	 	sTemp  = Request("sPassData")
	 	sUser  = Split(sTemp,":")(0)
	 	sRole  = Split(sTemp,":")(1)
	 	nPracticeCode = Split(sTemp,":")(2)
	 	nProcessCode  = Split(sTemp,":")(3)
	 	
	 	'If sRole="ALL" Then sRole = 1
	 	
	 	Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root
		
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			If sRole = "ALL" Then
				.Source = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM Ms_Roles "
			Else
				.Source = "SELECT DISTINCT ROLEID,ROLEDESCRIPTION FROM Ms_Roles WHERE ROLEID ="& sRole &" Order by ROLEID"
			End IF
			.ActiveConnection = con
			.Open
		end with
		
		set dcrs2.ActiveConnection = nothing
		
		nPrevRoleID = cdbl("0")
		
		If Not dcrs2.EOF Then
			Do while Not dcrs2.EOF 
			
				sRoleName = dcrs2(1)
				'Root.setAttribute "DATA",dcrs2.Source 
		
				sSql = " SELECT DISTINCT APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE FROM MS_ROLEACTIVITY "&_
					   " WHERE RoleID = "& dcrs2(0) &" AND Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) NOT IN "&_
				       " (select DISTINCT (Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) ) "&_
				       " FROM MS_USERACTIVITY WHERE INTERNALUSERID = "& sUser &" ) "
				
				If nProcessCode <> "S" Then
					sSql = sSql & " AND APPLICATIONCODE = "& nProcessCode &" "
				End IF
				
				If nPracticeCode <> "S" Then
					sSql = sSql & " AND PROCESSCODE IN ("& nPracticeCode &") "
				End IF
						
				'Root.setattribute "Query",sSql
		
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
		
				set dcrs1.ActiveConnection = nothing
		
				if not dcrs1.EOF then
					do while not dcrs1.EOF 
						
					with dcrs 
						.CursorLocation = 3
						.CursorType = 3
						.source = " SELECT DISTINCT V.APPLICATIONNAME,V.PROCESSNAME,V.ACTIVITYNAME FROM VWACTIVITY V,MS_ROLEACTIVITY R WHERE V.APPLICATIONCODE = " & trim(dcrs1(0)) & " and V.PROCESSCODE = " & trim(dcrs1(1)) & " and V.ACTIVITYCODE = " & trim(dcrs1(2)) & " and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE =  R.ACTIVITYCODE AND R.ROLEID = "& dcrs2(0) &" "
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
				
					if not dcrs.EOF then
				
						If cdbl(nPrevRoleID) = cdbl("0") or cdbl(nPrevRoleID) <> cdbl(dcrs2(0)) Then
							nPrevRoleID = dcrs2(0)
							Set newElem = OutData.createElement("ROLE")
							newElem.setAttribute "ROLEID", dcrs2(0)
							newElem.setAttribute "ROLENAME", sRoleName
							Root.appendchild newElem 
						End IF
							
						Set newElem1 = OutData.createElement("TOMAP")
						newElem1.setAttribute "APPCode", trim(dcrs1(0))
						newElem1.setAttribute "PRCode", trim(dcrs1(1))
						newElem1.setAttribute "ACCode", trim(dcrs1(2))
						newElem1.setAttribute "APPName", trim(dcrs(0))'trim(dcrs1(5))
						newElem1.setAttribute "PAName", trim(dcrs(1)) 'trim(dcrs1(3))
						newElem1.setAttribute "ACName", trim(dcrs(2)) 'trim(dcrs1(4))
						newElem.appendChild newElem1
					End IF
					dcrs.Close 
				
				dcrs1.MoveNext 
				Loop
			end if
			dcrs1.Close
					
				dcrs2.MoveNext 
			Loop
		End IF
		dcrs2.Close 
		
		'OutData.save server.MapPath("../Temp/UnMappedActvities.xml")
	
	Elseif sWho = "UANM_OLD" then	'User Un Mapped Activity Selection
	
		sTemp = Request("sPassData")
		sUser = Split(sTemp,":")(0)
		sRole = Split(sTemp,":")(1)
		
		Set Root = OutData.createElement("ACTIVITY")												
		OutData.appendChild Root

		'sSql = " Select Distinct R.RoleId,A.APPLICATIONCODE,A.PROCESSCODE,A.ACTIVITYCODE "&_
		'	   " From MS_USERROLES R,Ms_UserActivity A Where R.INTERNALUSERID = A.INTERNALUSERID and R.INTERNALUSERID Not in "&_
		'	   " (SELECT Distinct INTERNALUSERID From VwUSERACTIVITY WHERE INTERNALUSERID = "&sUser&" ) and R.RoleId = "& sRole &" order by 2,3,4 "
		
		'sSql = " Select Distinct R.RoleId,A.APPLICATIONCODE,A.PROCESSCODE,A.ACTIVITYCODE "&_
		'	   " From MS_USERROLES R,Ms_UserActivity A Where R.INTERNALUSERID = A.INTERNALUSERID and A.ACTIVITYCODE Not in "&_
		'	   " (SELECT Distinct ACTIVITYCODE From VwUSERACTIVITY WHERE INTERNALUSERID = "&sUser&" ) order by 2,3,4 "
		
		sSql = " Select Distinct R.RoleId,A.APPLICATIONCODE,A.PROCESSCODE,A.ACTIVITYCODE "&_
			   " From MS_USERROLES R,Ms_UserActivity A ,Ms_ApplicationActivity AA "&_
			   " Where R.INTERNALUSERID = A.INTERNALUSERID and A.APPLICATIONCODE = AA.APPLICATIONCODE and AA.STATUS = 'A'"&_
			   " and A.PROCESSCODE = AA.PROCESSCODE and A.ACTIVITYCODE = AA.ACTIVITYCODE and AA.ACTIVITYNAME Not in "&_
			   " (SELECT Distinct ACTIVITYNAME From VwUSERACTIVITY WHERE INTERNALUSERID = "&sUser&" ) order by 2,3,4 "
		
		
		
		sSql = "Select Distinct R.RoleId,A.APPLICATIONCODE,A.PROCESSCODE,A.ACTIVITYCODE "&_
				" From MS_USERROLES R,Ms_UserActivity A ,Ms_ApplicationActivity AA "&_
				" Where R.INTERNALUSERID = A.INTERNALUSERID and A.APPLICATIONCODE = AA.APPLICATIONCODE and AA.STATUS = 'A'"&_
				" and A.PROCESSCODE = AA.PROCESSCODE and A.ACTIVITYCODE = AA.ACTIVITYCODE and R.RoleID = "& sRole &" "&_
				" and R.INTERNALUSERID ="&sUser&"  "&_
				" and (Cast(A.APPLICATIONCODE as Varchar)+ ':' + cast(A.PROCESSCODE as Varchar) + ':' + cast(A.ACTIVITYCODE as Varchar) ) Not in "&_
				"( select DISTINCT (Cast(V.APPLICATIONCODE as Varchar)+ ':' + cast(V.PROCESSCODE as Varchar) + ':' + cast(V.ACTIVITYCODE as Varchar) ) "&_
				" from VwUSERACTIVITY V, Ms_RoleActivity R WHERE V.INTERNALUSERID ="&sUser&" and R.RoleID = "& sRole &" "&_
				" and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE)"
				
		sSql = "Select Distinct R.RoleId,A.APPLICATIONCODE,A.PROCESSCODE,A.ACTIVITYCODE "&_
				" From MS_USERROLES R,Ms_UserActivity A ,Ms_ApplicationActivity AA "&_
				" Where R.INTERNALUSERID = A.INTERNALUSERID and A.APPLICATIONCODE = AA.APPLICATIONCODE and AA.STATUS = 'A'"&_
				" and A.PROCESSCODE = AA.PROCESSCODE and A.ACTIVITYCODE = AA.ACTIVITYCODE and R.RoleID = "& sRole &" "&_
				" and R.INTERNALUSERID ="&sUser&"  "&_
				" and (Cast(A.APPLICATIONCODE as Varchar)+ ':' + cast(A.PROCESSCODE as Varchar) + ':' + cast(A.ACTIVITYCODE as Varchar) ) Not in "&_
				"( select DISTINCT (Cast(V.APPLICATIONCODE as Varchar)+ ':' + cast(V.PROCESSCODE as Varchar) + ':' + cast(V.ACTIVITYCODE as Varchar) ) "&_
				" from VwUSERACTIVITY V, MS_USERACTIVITY R ,MS_USERROLES UR WHERE V.INTERNALUSERID ="&sUser&" and UR.RoleID = "& sRole &" "&_
				" and V.APPLICATIONCODE = R.APPLICATIONCODE and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE AND V.INTERNALUSERID = UR.INTERNALUSERID)"
				
				
		sSql = " SELECT DISTINCT APPLICATIONCODE, PROCESSCODE,ACTIVITYCODE,APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY V where  "&_
				" Cast(V.APPLICATIONCODE as Varchar)+ ':' + cast(V.PROCESSCODE as Varchar) + ':' + cast(V.ACTIVITYCODE as Varchar) NOT IN  "&_
				" (select DISTINCT (Cast(V.APPLICATIONCODE as Varchar)+ ':' + cast(V.PROCESSCODE as Varchar) + ':' + cast(V.ACTIVITYCODE as Varchar) ) "&_
				"  from VwUSERACTIVITY V, MS_USERACTIVITY R ,MS_USERROLES UR "&_
				" WHERE V.INTERNALUSERID = UR.INTERNALUSERID and V.INTERNALUSERID ="&sUser&"  and UR.RoleID = "& sRole &" and V.APPLICATIONCODE = R.APPLICATIONCODE  "&_
				" and V.PROCESSCODE = R.PROCESSCODE and V.ACTIVITYCODE = R.ACTIVITYCODE )ORDER BY 1,2,3 "
				
		sSql = " SELECT DISTINCT APPLICATIONCODE, PROCESSCODE,ACTIVITYCODE FROM MS_ROLEACTIVITY  "&_
			   " WHERE RoleID = "& sRole &" AND Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) NOT IN  "&_
			   " (select DISTINCT (Cast(APPLICATIONCODE as Varchar)+ ':' + cast(PROCESSCODE as Varchar) + ':' + cast(ACTIVITYCODE as Varchar) ) "&_
			   " FROM MS_USERACTIVITY WHERE INTERNALUSERID = "&sUser&" ) "


		'Root.setattribute "Query",sSql
		
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql 
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				
				iCtr = iCtr + 1
				
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
				'	'.Source = "SELECT DISTINCT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(1)) & " AND PROCESSCODE = " & trim(dcrs(2)) & " AND ACTIVITYCODE = " & trim(dcrs(3)) & ""
					.Source = "SELECT DISTINCT APPLICATIONNAME,PROCESSNAME,ACTIVITYNAME FROM VWACTIVITY WHERE APPLICATIONCODE = " & trim(dcrs(0)) & " AND PROCESSCODE = " & trim(dcrs(1)) & " AND ACTIVITYCODE = " & trim(dcrs(2)) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				
				if not dcrs1.EOF then

					Set newElem = OutData.createElement("TOMAP")
					newElem.setAttribute "CTR",iCtr
					newElem.setAttribute "APPCode", trim(dcrs(0))'trim(dcrs(1))
					newElem.setAttribute "PRCode", trim(dcrs(1)) 'trim(dcrs(2))
					newElem.setAttribute "ACCode", trim(dcrs(2)) 'trim(dcrs(3))
					newElem.setAttribute "APPName", trim(dcrs1(0)) 'trim(dcrs(3))
					newElem.setAttribute "PAName", trim(dcrs1(1))  'trim(dcrs(4))
					newElem.setAttribute "ACName", trim(dcrs1(2))  'trim(dcrs(5))
					Root.appendChild newElem
				end if
				dcrs1.Close
				
				dcrs.MoveNext
			loop
		end if
		dcrs.Close
		'OutData.save server.MapPath("../Temp/APpUser.xml")
	end if
	
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>

<%
	Function DisplayOrganization(sOrgID)
		Dim dcrs,str

		if sOrgID = "" then
			DisplayOrganization = "-"
			exit function
		end if
		
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (" & sOrgID & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.EOF 
			str = str & ", " & trim(dcrs(0)) 
			dcrs.moveNext
		loop
		dcrs.Close

		DisplayOrganization = mid(str,2)
	End Function
%>
