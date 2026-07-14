<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetMappedContraDetails.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 11,2010
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
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
Dim Objrs,Objrs1,objDOM,objRs2
Dim ndRoot,ndAcc
Dim sQuery,sOrgid,sAccName,sToAccName
Dim iCnt,iFromHead,iToAccCode,iRecordsCount,iBookCode,iBookNumber

sOrgid=trim(Request.QueryString("OrgCode"))
iFromHead=trim(Request.QueryString("FromHead"))

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
Set Objrs = Server.CreateObject("ADODB.RecordSet")
Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

set ndRoot = objDOM.createElement("Root")
objDOM.appendChild ndRoot


sQuery="select b.AccountDescription,a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
				"where a.OUDefinitionID='"&sorgID&"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead="& iFromHead 
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
	
		set objRs.ActiveConnection = nothing
	
		set sAccName=objRs(0)
		set iToAccCode=objRs(1)
	
		if not objRs.EOF then
			do while not objRs.EOF
				iRecordsCount = 0
				sQuery="select AccountDescription from Acc_M_GLAccountHead "&_
						"where AccountHead="&iToAccCode
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing
				if not objRs1.EOF then
					sToAccName=objRs1(0)
				end if
				objRs1.Close
				
				sQuery = "Select BookCode,BookNumber from Acc_R_ApplicableAccountHeads where BookAccountHead = "& iToAccCode
				'Response.Write sQuery
				objRs1.Open sQuery,con
				if not objRs1.EOF then
					iBookCode = objRs1(0)
					iBookNumber = objrs1(1)
					sQuery = "Select Count(CreatedTransNo) from ACC_T_CreatedVoucherheader where BookCode = "& iBookCode  &" and BookNumber = "& iBookNumber 
					'Response.Write sQuery
					objRs2.Open sQuery,con
					if not objRs2.EOF then
						iRecordsCount = objRs2(0)
					end if
					objRs2.Close 
				end if
				objRs1.Close 
				
				
				
				set ndAcc = objDOM.createElement("Acc")
					ndAcc.setAttribute "No",iToAccCode
					ndAcc.setAttribute "Name",sToAccName
					if iRecordsCount=0 then
						ndAcc.setAttribute "Records","N"
					else
						ndAcc.setAttribute "Records","Y"
					end if
				ndRoot.appendChild ndAcc
			objRs.MoveNext
			loop				
		end if	

Response.ContentType = "text/xml"
Response.Write objDOM.xml
%>
