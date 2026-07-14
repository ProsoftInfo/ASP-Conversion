<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgSalesControlInsert.asp	
	'Module Name				:	Inventory (Organization Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 03, 2002
    'Modified By                :   Ragavendran R
	'Modified On				:   July 21,2011
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

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,org) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "OrgManufacturingControlEntry.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
dim dcrs,dcrs1,sSql
Dim sTaxCatDesc,sTaxCatShDesc,iTaxCatCode,sTaxDesc,sTaxShDesc,iTaxCode
dim sOrgCode,iOrgIndex,sOrgName,iCtr,sChecked
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

sOrgCode = trim(Request.Form("hOrgCode"))
sOrgName = trim(Request.Form("hOrgName"))

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT OUDEFINITIONID FROM INV_CONTROL_ORGSALES WHERE OUDEFINITIONID = " & Pack(sOrgCode) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if dcrs.EOF then
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT AC.TAXCATEGORYCODE,TAXCATEGORYNAME,TAXCATEGORYSHORTNAME,TAXCODE,TAXNAME,TAXSHORTNAME FROM APP_M_CHARGESHEADER AC,APP_M_CHARGESDETAILS AD WHERE AC.TAXCATEGORYCODE = AD.TAXCATEGORYCODE ORDER BY AC.TAXCATEGORYCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		set iTaxCatCode = dcrs1(0)
		set sTaxCatDesc = dcrs1(1)
		set sTaxCatShDesc = dcrs1(2)
		set iTaxCode = dcrs1(3)
		set sTaxDesc = dcrs1(4)
		set sTaxShDesc = dcrs1(5)
		
		Do While Not dcrs1.EOF
			sChecked = trim(Request.Form("chk"&iTaxCatCode&":"&iTaxCode))
			if sChecked = "1" then
				sSql = "INSERT INTO INV_CONTROL_ORGSALES (OUDEFINITIONID,TAXCATEGORYCODE,TAXCODE,TAXCATEGORYNAME," &_
					" TAXCATEGORYELIGIBILITY) VALUES" &_
					"(" & Pack(sOrgCode) & "," & iTaxCatCode & "," & iTaxCode & "," & Pack(sTaxCatDesc) & "," &_
					"1)"
				'Response.Write sSql & "<BR>"
				con.Execute sSql
			end if
		dcrs1.MoveNext
		loop
		dcrs1.Close
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Sales \nhas been defined Successfully','Y','1')">
<%
else
%>
	<BODY BGCOLOR="#FFFFFF" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> Control for Sales \nhas been already defined','N','1')">
<%end if
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
	
	con.CommitTrans
end if

con.close
set con = nothing
%>
