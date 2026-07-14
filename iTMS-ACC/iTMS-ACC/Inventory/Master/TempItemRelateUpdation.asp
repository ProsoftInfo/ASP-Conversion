<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TempItemRelateUpdation.asp
	'Module Name				:	Inventory (Temporary Item Relation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	October 06, 2011
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

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to Relate one more Item"))
				window.location.href = "TEMPORARYITEMS.asp"
			else
				window.location.href = "../welcome_Inventory.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
dim dcrs,iTICode,iItemCode,iClassCode,sOrgCode,sSql

iTICode		= trim(Request.Form("hTempItemCode"))
iItemCode	= trim(Request.Form("hItemCode"))
iClassCode	= trim(Request.Form("hClassCode"))
sOrgCode	= trim(Request.Form("hOrgId"))

con.beginTrans
	
Set dcrs = Server.CreateObject("ADODB.RecordSet")

sSql = "INSERT INTO MS_TEMPFINALITEMDETAIL (TEMPITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
	"ITEMCODE) VALUES " &_
	"(" & iTICode & "," & Pack(sOrgCode) & "," & iClassCode & "," &_
	"" & iItemCode & ")"
'Response.Write sSql & "<BR>"
con.Execute sSql

sSql = "UPDATE MS_TEMPORARYITEMMASTER SET FINALSTATUS = 'Y' WHERE TEMPITEMCODE = " & trim(iTICode) & ""
'Response.Write sSql & "<BR>"

con.Execute sSql

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing

%>
<BODY onLoad = "msgbox('Item has been related Successfully','Y')">
