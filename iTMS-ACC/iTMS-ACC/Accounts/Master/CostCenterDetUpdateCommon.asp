<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CostCenterDetUpdateCommon.asp	
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	March 19, 2011
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
	function msgboxGP(strr,flag) {
		var sTarget
		sTarget=document.formname.hTargetPage.value
		if (flag == "Y") {
			alert(strr);
			window.location.href = sTarget
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
	
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "CostCenters.asp" 
			//"CCEditMain.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
dim sQuery,objRs,iUserid,sReason,sClassName,sGcode,sGroupFlag,sTargetName,sCallFrom
dim iHisno,iCCCode,sCCName,sCCShortName,iCCGroupCode,sAppBy,sAppon,sUnitId

Set objRs = Server.CreateObject("ADODB.RecordSet")

sCallFrom = trim(Request.Form("hCallFrom"))
iUserid = getUserID()
sUnitId = Session("organizationcode")

If sCallFrom = "GP" Then

	sClassName = replace (trim(Request.Form("txtClassName")),"'","''")
	sGcode = trim(Request.Form("hGroupCode"))
	sGroupFlag = trim(Request.Form("hGroupFlag"))

	select case  sGroupFlag
	case "G"
		sGcode=MID (sGcode,3)
		sTargetName="AccGroupCreationMain.asp"
		sQuery = "update Acc_M_AccountGroups set AccountsGroupName='"&sClassName&"' where AccountsGroupCode='"&sGcode&"'"
	case "A"
		sTargetName="ANALGroupCreationMain.asp"
		sQuery = "update Acc_M_AnalyticalGroup set AHGroupName='"&sClassName&"' where AHGroupCode='"&sGcode&"'"
	case "C"
		'sTargetName="CCGroupCreationMain.asp"
		sTargetName = "CostCenters.asp"
		sQuery = "update Acc_M_CostCenterGroup set CCGroupName='"&sClassName&"' where CCGroupCode='"&sGcode&"'"
	end select 
	con.Execute sQuery
%>
	<script src="/Scripts/itms-modern-compat.js"></script>
<BODY BGCOLOR="#336699" onLoad = "msgboxGP('Group has been Updated Successfully','Y')">
	<form name="formname">
		<input type="hidden" name="hTargetPage" value="<%=sTargetName%>">
	</form>

<%

Else
	sCCName = trim(Request.Form("txtClassName"))
	sCCShortName=trim(Request.Form("txtShortName"))
	iCCCode = trim(Request.Form("hHeadValue"))
	sReason = Request.Form("txtReason")

	sQuery = "Select CCGroupCode,ApprovedBy,Convert(Char,ApprovedOn,103) From Acc_R_OrgCostCenter Where CostCenterHead = "&iCCCode&" "
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
		
	Set objRs.ActiveConnection = Nothing

	IF Not objRs.EOF Then 
		iCCGroupCode = objRs(0)
		sAppBy = objRs(1)
		sAppon = objRs(2)
	End IF
	objRs.Close


	Con.BeginTrans

	sQuery = "Select isNull(Max(OCCAmendmentNo),0) + 1 From Acc_R_HistoryOrgCC "
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
	
	sQuery = "SELECT ApprovedBy, Convert(Char,CreatedOn,103), CreatedBy, Convert(Char,ApprovedOn,103)  FROM Acc_R_OrgCostCenter "&_
			 "WHERE OUDefinitionID = '"&sUnitId&"' AND CostCenterHead = "&iCCCode&" "
					 
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	Set objRs.ActiveConnection = Nothing
	IF Not objRs.EOF Then
		sQuery = "INSERT INTO Acc_R_HistoryOrgCC (OCCAmendmentNo, OUDefinitionID, CostCenterHead, CCGroupCode, "&_
				 "CreatedBy, CreatedOn, ApprovedBy, ApprovedOn, HistoryTypeOCC, HistoryBy, HistoryOn) "&_
				 "VALUES ("&iHisno&", '"&Trim(sUnitId)&"', "&iCCCode&", '"&iCCGroupCode&"', "&objRs(0)&", Convert(datetime,'"&Trim(objRs(1))&"',103), "&objRs(2)&", Convert(datetime,'"&Trim(objRs(3))&"',103), 'A', "&getUserid&", getDate()) "
						 
		'Response.Write sQuery &"<br>"
		Con.Execute sQuery
	End IF
	objRs.Close
	

	sQuery = "UPDATE Acc_M_CCAccountHead SET CCHeadCode = '"&sCCShortName&"', CCAccountDescription = '"&sCCName&"' "&_
			 "WHERE CostCenterHead = "&iCCCode&" "
		 
	'Response.Write sQuery 
	Con.Execute sQuery
			
	Con.CommitTrans		

%>
<HTML>
	<script src="/Scripts/itms-modern-compat.js"></script>
<BODY BGCOLOR="#336699" onLoad = "msgbox('Head Updated Successfully','Y')"></BODY>
<%End If 'If sCallFrom = "GP" Then
con.close
set con = nothing
%>
</HTML>