<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParPrefPopupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 19,2010
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%
Dim rsObj,sQuery
Dim sPayTerms,sBop,sDespatch,sPayMode,sTransport,sCurrency,iPartyCode

set rsObj = Server.CreateObject("ADODB.Recordset")
	con.begintrans

		iPartyCode = Request.Form("hPartyCode")
		sPayTerms = trim(Request.Form("selPayTerms"))
		sBop= trim(Request.Form("selBop"))
		sDespatch= trim(Request.Form("selDespatch"))
		sPayMode = trim(Request.Form("selPayMode"))
		sTransport= trim(Request.Form("selTransport"))
		sCurrency = trim(Request.Form("selCurrency"))

		sQuery =  "Update APP_R_OrgParty set PrefTransporterCode="& sTransport &","&_
				  "PrefDespatchMode="& sDespatch &",PrefCurrencyCode="& sCurrency &","&_
				  "PrefPaymentMode="& sPayMode &",PrefBasisOfPricing="& sBop &","&_
				  "PrefPaymentTerms="& sPayTerms &" where PartyCode = "& iPartyCode
		Response.Write sQuery
		con.execute sQuery

'	con.rollbacktrans
'	Response.End 

	Response.Clear 
	con.committrans
%>

<HTML>
<HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<base target="_self">
<script language="javascript">
window.__itmsPopupCompat = {
	type: "autoClose",
	message: "Perference Updated Successfully",
	returnValue: "Done"
};
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY onload="init()">
</BODY>
</HTML>


