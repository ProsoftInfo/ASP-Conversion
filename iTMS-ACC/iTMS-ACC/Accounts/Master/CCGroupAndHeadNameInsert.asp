<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CCGroupAndHeadNameInsert.asp	
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<SCRIPT>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "COSTCENTERS.ASP"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
dim dcrs,objRs,sQuery,Temp,sUnitID
dim sClassName,sGcode,sParentCode,arrpGroupCode,sGroupValue,sCallFrom,iCounter,arrUnit
Dim iCCCode,sCCName,sCCShortName,sUserID

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set dcrs = Server.CreateObject("ADODB.RecordSet")

sCallFrom = Trim(Request.form("hCallFrom"))
sUserID = Session("userid")

If sCallFrom = "GP" Then
	sClassName = UCase(replace (trim(Request.Form("txtClassName")),"'","''"))
	sGroupValue = trim(Request.Form("hpGroup"))
	sParentCode = sGroupValue

	sQuery = "SELECT CCGroupName FROM Acc_M_CostCenterGroup WHERE LOWER(CCGroupName) = '" & lcase(Trim(sClassName)) & "'"
    with dcrs
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
    If Not dcrs.EOF Then
%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox(' Cost Center Group Name Already Exists','N')">
<%
	
	Response.End
    End If
    dcrs.Close

    If Trim(sParentCode) = "0" Then
			sQuery = "Select Max(CCGroupCode) + 1 from Acc_M_CostCenterGroup where CCParentGroup=CCGroupCode "
			dcrs.Open sQuery,con
			If dcrs.EOF or isnull(dcrs(0)) then
				sGCode = "01"
			ElseIF dcrs(0)<10 THEN
				sGCode = "0"&dcrs(0)
			ELSE
				sGCode = dcrs(0)
			End if
			dcrs.Close
		sQuery = "Insert into Acc_M_CostCenterGroup (CCGroupCode,CCGroupName,CCParentGroup,ChildCount) values('"&sGCode&"','"&Trim(sClassName)&"','"&sGCode&"',0)"	
        con.Execute sQuery
%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Group has been Created Successfully','Y')">
<%
    Else
		'*********
		sQuery = "Select childcount + 1 from Acc_M_CostCenterGroup where CCGroupCode = '"&sParentCode&"'"
		
		dcrs.Open sQuery,con
		
		If Not dcrs.EOF then
		  If dcrs(0) < 10 then			
			 sGcode = "0"&trim(dcrs(0))
		  Else
			sGcode = dcrs(0)
		  End if	
		End if
			
		sGcode = trim(sParentCode)&trim(sGcode)
			
		sQuery = "Insert into Acc_M_CostCenterGroup (CCGroupCode,CCGroupName,CCParentGroup,ChildCount) values('"&sGcode&"','"&Trim(sClassName)&"','"&sParentCode&"',0)"	
		con.Execute sQuery
			
		sQuery = "Update Acc_M_CostCenterGroup SET Childcount = childcount + 1 where CCGroupCode = '"&sParentCode&"' "
		con.execute sQuery
        '*********
%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Group has been Created Successfully','Y')">
<%
    End If

Elseif sCallFrom = "HD" Then

	sCCName = UCase(replace (trim(Request.Form("txtClassName")),"'","''"))
	sCCShortName=UCase(replace (trim(Request.Form("txtShortName")),"'","''"))
	sParentCode = trim(Request.Form("hParentCode"))

	sUnitID=trim(Request.Form("hUnitID"))
	
	sQuery="select isnull(max(CostCenterHead)+1,1)  from Acc_M_CCAccountHead"

	with objRs	
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	iCCCode=objRs(0)

	set objRs=nothing

	sQuery="INSERT INTO Acc_M_CCAccountHead(CostCenterHead, CCAccountDescription, CCHeadCode) "&_
			"VALUES("&iCCCode&",'"&sCCName&"','"&sCCShortName&"')"

	con.Execute(sQuery)

	'FOR iCounter=0 to UBound(arrUnit)
		sQuery="INSERT INTO Acc_R_OrgCostCenter(OUDefinitionID,CostCenterHead,CCGroupCode,CreatedBy, CreatedOn, ApprovedBy, ApprovedOn) "&_
				"VALUES('"& sUnitID &"',"&iCCCode&",'"&sParentCode&"',"&sUserID&",getdate(),"&sUserID&",getdate())"
			'Response.Write "<p>sql="&sQuery
		con.Execute(sQuery)
	'next		
	%>
	<HTML>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Cost Center has been Created Successfully','Y')">

	<%
	
End IF 'If sCallFrom = "GP" Then
'con.rollbackTrans
'Response.End

'Con.CommitTrans

con.close
set con = nothing
%>

</HTML>
