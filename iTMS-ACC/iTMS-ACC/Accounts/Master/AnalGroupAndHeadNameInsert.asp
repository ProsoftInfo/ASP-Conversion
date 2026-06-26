<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AnalGroupAndHeadNameInsert.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Created On					:	January 19,2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	March 20, 2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "ANALYTICALHEADS.ASP"	//"ANALGroupCreationMain.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
dim dcrs,sQuery,Temp,sCallFrom ,sUnitID
dim sClassName,sGcode,sParentCode,arrpGroupCode,sGroupValue
Dim iCounter,arrUnit,objRs,iANALCode,sANALName,sANALShortName

Set objRs = Server.CreateObject("ADODB.RecordSet")

sCallFrom = Trim(Request.Form("hCallFrom"))
	
If sCallFrom = "GP" Then

	
	sClassName = UCase(replace (trim(Request.Form("txtClassName")),"'","''"))
	sGroupValue = trim(Request.Form("hpGroup"))

	if sGroupValue<>"0" then
		sParentCode = sGroupValue
	else
		sParentCode="GRP"	
	end if

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sQuery = "SELECT AHGroupName FROM Acc_M_AnalyticalGroup WHERE LOWER(AHGroupName) = '" & lcase(Trim(sClassName)) & "'"
    with dcrs
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
    If Not dcrs.EOF Then
%>
	<BODY onLoad = "msgbox('Analytical Group Name Already Exists','N')">
<%
	
	Response.End
    End If
    dcrs.Close

    If Trim(sParentCode) = "GRP" Then
			sQuery = "Select Max(AHGroupCode) + 1 from Acc_M_AnalyticalGroup where AHParentGroup=AHGroupCode "
			
			dcrs.Open sQuery,con
			
			If dcrs.EOF or isnull(dcrs(0)) then
				sGCode = "01"
			ElseIF dcrs(0)<10 THEN
				sGCode = "0"&dcrs(0)
			ELSE
				sGCode = dcrs(0)
			End if
			dcrs.Close
		sQuery = "Insert into Acc_M_AnalyticalGroup (AHGroupCode,AHGroupName,AHParentGroup,ChildCount) values('"&sGCode&"','"&Trim(sClassName)&"','"&sGCode&"',0)"	
        con.Execute sQuery
%>
	<BODY  onLoad = "msgbox('Group has been Created Successfully','Y')">
<%
    Else
		'*********
		sQuery = "Select childcount + 1 from Acc_M_AnalyticalGroup where AHGroupCode = '"&sParentCode&"'"
		
		dcrs.Open sQuery,con
		
		If Not dcrs.EOF then
		  If dcrs(0) < 10 then			
			 sGcode = "0"&trim(dcrs(0))
		  Else
			sGcode = dcrs(0)
		  End if	
		End if
			
		sGcode = trim(sParentCode)&trim(sGcode)
			
		sQuery = "Insert into Acc_M_AnalyticalGroup (AHGroupCode,AHGroupName,AHParentGroup,ChildCount) values('"&sGcode&"','"&Trim(sClassName)&"','"&sParentCode&"',0)"	
		con.Execute sQuery
			
		sQuery = "Update Acc_M_AnalyticalGroup SET Childcount = childcount + 1 where AHGroupCode = '"&sParentCode&"' "
		con.execute sQuery
        '*********
%>
	<BODY  onLoad = "msgbox('Group has been Created Successfully','Y')">
<%
    End If

Else
	
	sANALName = UCase(replace (trim(Request.Form("txtClassName")),"'","''"))
	sANALShortName=UCase(replace (trim(Request.Form("txtShortName")),"'","''"))
	sParentCode = trim(Request.Form("hParentCode"))

	Temp=Split(Request.Form("selAnalHead"),"?")
	iANALCode=trim(Temp(0))

	sUnitID =trim(Request.Form("hUnitID"))
	
if cint(iANALCode)=0 then
	sQuery = "SELECT AnalyticalShortName FROM Acc_M_AnalyticalHeads WHERE LOWER(AnalyticalShortName) = '" & lcase(Trim(sANALShortName)) & "'"
	with objRs
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	If Not objRs.EOF Then
%>
		<BODY onLoad = "msgbox(' Analytical Head Already Exists','N')">
<%
	
		Response.End
	End If
	objRs.Close

	sQuery="select isnull(max(AnalyticalCode)+1,1)  from Acc_M_AnalyticalHeads"
	with objRs	
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	iANALCode=objRs(0)

	set objRs=nothing

	sQuery="INSERT INTO Acc_M_AnalyticalHeads(AnalyticalCode, AnalyticalName, AnalyticalShortName) "&_
			"VALUES("&iANALCode&",'"&sANALName&"','"&sANALShortName&"')"

	con.Execute(sQuery)
end if

'for iCounter=0 to UBound(arrUnit)
	sQuery="INSERT INTO Acc_R_OrgAnalyticalHeads(OUDefinitionID,AnalyticalCode,AHGroupCode,CreatedBy, CreatedOn, ApprovedBy, ApprovedOn) "&_
			"VALUES('"&trim(sUnitID)&"',"&iANALCode&",'"&sParentCode&"',"&getUserid&",getdate(),"&getUserid&",getdate())"
	con.Execute(sQuery)
'next		
%>
	<BODY  onLoad = "msgbox('Analytical Head has been Created Successfully','Y')">
<%
	
End IF	'If sCallFrom = "GP" Then
con.close
set con = nothing

%>
</HTML>