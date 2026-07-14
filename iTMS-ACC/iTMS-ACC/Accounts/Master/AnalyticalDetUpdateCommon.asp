<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AnalyticalDetUpdateCommon.asp	
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Created On					:	July 21, 2002
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
<!--#include virtual="/include/populate.asp"-->
<SCRIPT>
<!--
	function msgbox(strr,flag) {
		
		if (flag == "Y") {
			alert(strr);
			window.location.href = "ANALYTICALHEADS.ASP"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
Dim sQuery,sCallFrom,sClassName,sGcode,sUnitID
Dim Temp,iCounter,arrUnit,objRs,sCallTy,sRemUnits,sAddUnits,sDisUnits
Dim iANALCode,sANALName,sANALShortName,iHisno,sAnalGpCode,sTemp,iParCode

Set objRs = Server.CreateObject("ADODB.RecordSet")

sCallFrom = trim(Request.Form("hCallFrom"))
sUnitID = Trim(Session("organizationcode"))

If sCallFrom = "GP" Then

	sClassName = replace (trim(Request.Form("txtClassName")),"'","''")
	sGcode = trim(Request.Form("hGroupCode"))
	
	sQuery = "update Acc_M_AnalyticalGroup set AHGroupName='"&sClassName&"' where AHGroupCode='"&sGcode&"'"
	con.Execute sQuery

%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Group has been Updated Successfully','Y')">
	<form name="formname"></form>
<%
Else

	sANALName = trim(Request.Form("txtClassName"))
	sANALShortName=trim(Request.Form("txtShortName"))
	iANALCode = trim(Request.Form("hHeadValue"))
	iParCode = trim(Request.Form("hParentCode"))
	
	sCallTy = "K"

	Con.BeginTrans
			
	sQuery = "Select isNull(Max(OAHAmendmentNo),0) + 1 From Acc_R_HistoryOrgAH "
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set objRs.ActiveConnection = Nothing
	IF Not objRs.EOF Then
		iHisno = objRs(0)
	End IF
	objRs.Close
				
				
	sQuery = "SELECT AHGroupCode, CreatedBy, Convert(char,CreatedOn,103), ApprovedBy, Convert(char,ApprovedOn,103) "&_
			 "FROM Acc_R_OrgAnalyticalHeads Where OUDefinitionID = '"&Trim(sUnitID)&"' and AnalyticalCode = "&iANALCode&" "
						 
				
	'Response.Write sQuery &"<br><br><br><br><br>"
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set objRs.ActiveConnection = Nothing
	IF Not objRs.EOF Then
		sQuery = "INSERT INTO Acc_R_HistoryOrgAH (OAHAmendmentNo, OUDefinitionID, AnalyticalCode, AHGroupCode, "&_
				 "CreatedBy, CreatedOn, ApprovedBy, ApprovedOn, HistoryTypeOAH, HistoryBy, HistoryOn) "&_
				 "VALUES ("&iHisno&", '"&Trim(sUnitID)&"', "&iANALCode&", '"&objRs(0)&"', "&objRs(1)&", Convert(datetime,'"&objRs(2)&"',103), "&objRs(3)&", Convert(datetime,'"&objRs(4)&"',103), 'A', "&getUserid&", Convert(datetime,getDate(),103)) "
							 
		'Response.Write sQuery &"<br>"
		Con.Execute sQuery
	End IF
	objRs.Close
	
	sANALName = Replace(sANALName,"'"," ")
	sQuery = "UPDATE Acc_M_AnalyticalHeads SET AnalyticalName = '"&sANALName&"', AnalyticalShortName = '"&sANALShortName&"' "&_
			 "WHERE AnalyticalCode = "&iANALCode&" "
			 
	'Response.Write sQuery
	Con.Execute sQuery

	Con.CommitTrans

%>
<BODY onLoad = "msgbox('Head Updated Successfully','Y')">
</BODY>
<%	

End IF	'If sCallFrom = "GP" Then

con.close
set con = nothing
%>
</HTML>