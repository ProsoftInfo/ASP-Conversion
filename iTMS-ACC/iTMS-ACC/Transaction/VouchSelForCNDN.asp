<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	VouchSelForCNDN.asp
	'Module Name				:	ACCOUNTS (Transcation Voucher Selection Only for Sal and Purchase)
	'Author Name				:	SENTHIL E
	'Created On					:	December 02,2003
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
<%
dim sOrgId,objRs,sQuery,objRs2
dim sBookCode,iBookNo,sTransType,bFlag,sCallTy,sAmount
Dim iPurTy,iParTy,iParCode,sVouchTy,iParSubTy,sTemp
Dim iPartyCode,iPartyType,sUptoDate,sFromDate

sTemp = Session("FinPeriod")
sTemp = Split(sTemp,":")
sUptoDate = Trim(sTemp(1))
IF Cstr(sUptoDate) <> "" Then
	sUptoDate = "31/03/"&sUptoDate
	sFromDate = "01/04/"&sTemp(0)
Else
	sUptoDate = "31/03/"&Year(Date)
	sFromDate = "01/04/"&Year(Date)+1
End IF

sUptoDate = Trim(sUptoDate)
sFromDate = Trim(sFromDate)

sOrgId=Request("orgid")
sBookCode=Request("BookCode")
iBookNo=Request("BookNo")
sTransType=Request("TransType")
bFlag=Request("flag")
iPurTy = Request("sPurTy")
iParTy = Request("sParTy")
iParCode = Request("iParCode")
sVouchTy = Request("VouchTy")



IF CStr(iParCode) <> "" Then
	sTemp = Split(iParCode,"?")
	iPartyCode = sTemp(3)
	iPartyType = sTemp(0)
	iParSubTy = sTemp(1)
Else
	iPartyCode = 0
	iPartyType = ""
	iParSubTy = 0
End IF

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Voucher No Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
var iTransNo = "0";
var dialogCompleted = false;

function completeDialog(value) {
	iTransNo = value == null ? "0" : String(value);
	dialogCompleted = true;
	window.returnValue = iTransNo;
	window.returnvalue = iTransNo;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(iTransNo);
	}
	window.close();
}

function finaldone() {
	var selectedValue;
	var selectedText;
	var valueParts;
	var textParts;
	if (String(document.formname.hCheckTy.value) === "0") {
		finalcancel();
		return;
	}
	selectedValue = document.formname.selAccountHead.value;
	selectedText = document.formname.selAccountHead.options[document.formname.selAccountHead.selectedIndex].text;
	valueParts = selectedValue.split("--");
	textParts = selectedText.split(" -- ");
	if (String(document.formname.hTransTy.value) === "SJR") {
		completeDialog((valueParts[0] || "") + "~" + (textParts[0] || "") + "~" + (valueParts[1] || ""));
		return;
	}
	completeDialog((valueParts[0] || "") + "~" + (textParts[0] || "") + "~0~" + (valueParts[1] || ""));
}

function finalcancel() {
	completeDialog(iTransNo);
}
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onunload()
{
	if (!dialogCompleted) {
		window.returnValue = iTransNo;
		window.returnvalue = iTransNo;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(iTransNo);
		}
	}
}
function  selectTheItem(obj,srcCombo){

		objSel = document.forms[0].elements[srcCombo];
		for(i=0; i < objSel.options.length; i++){
				objSel.options[i].selected = false;
		}

		for(i=0; i < objSel.options.length; i++){
			if ( obj.value != "" && objSel.options[i].text.toUpperCase().indexOf(obj.value.toUpperCase()) >=0 ){
				objSel.options[i].selected = true;
				return;
			}
		}
}
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function document_onkeypress(evt)
{
	evt = evt || window.event;
	if (evt && evt.keyCode==27)
	{
		finalcancel();
	}
}
document.addEventListener("keydown", document_onkeypress);
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=document EVENT=onkeypress>
	 document_onkeypress()
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" LANGUAGE=javascript onunload="return window_onunload()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hTransTy" value="<%=sTransType%>">
<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popuptable">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
                                    &nbsp;
                                    <p>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                    </p>
								</td>
								<td valign="top">
                                    <div align="center" height="20">
                                      <center>
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell">
 <input type="text" name="txtSearch" size="35"  ONKEYUP="selectTheItem(this,'selAccountHead')"  class="FormElem">
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell">

<%
dim sParCode,sParName,iRecCount,iPayCount
if bFlag="C" then
sQuery=" select CreatedTransNo,CreatedVoucherNo,convert(char,VoucherDate,103),str(VoucherAmount,10,2),PayToRecdFrom,isNull(BankInstrumentType,'N')  from Acc_T_CreatedVoucherHeader Where"&_
				" OUDefinitionID='"&sorgID&"' and BookCode='"&sBookCode&"' and "&_
				"BookNumber="&iBookNo&" and TransactionType='"&sTransType&"' "&_
				"and CreatedVouchStatus in('010101','010102','010103','010105') and "&_
				"convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sFromDate&"',103) "&_
				"and  convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sUptoDate&"',103) "&_
				" and FromApplication is Null "

else
'sQuery=" select CreatedTransNo,CreatedVoucherNo,convert(char,VoucherDate,103),str(VoucherAmount,10,2),PayToRecdFrom from Acc_T_VoucherHeader Where"&_
'				" OUDefinitionID='"&sorgID&"' and BookCode='"&sBookCode&"' and "&_
'				"BookNumber="&iBookNo&" and TransactionType='"&sTransType&"' "&_
'				"and VoucherStatus in('010104','010106','010107','010108') and "&_
'				"convert(datetime,VoucherDate,103) > Convert(datetime,'31/03/2004',103) "&_
'				" and BRSTransactionNo is Null "

sQuery = " select CreatedTransNo,CreatedVoucherNo,convert(char,VoucherDate,103),str(VoucherAmount,10,2),PayToRecdFrom,isNull(BankInstrumentType,'N')  from Acc_T_CreatedVoucherHeader Where"&_
		 " OUDefinitionID='"&sorgID&"' and BookCode='"&sBookCode&"' and "&_
		 "BookNumber="&iBookNo&" and TransactionType='"&sTransType&"' "&_
		 "and CreatedVouchStatus in('010104','010106','010107','010108') and "&_
		 "convert(datetime,VoucherDate,103) >= Convert(datetime,'"&sFromDate&"',103) "&_
		 "and  convert(datetime,VoucherDate,103) <= Convert(datetime,'"&sUptoDate&"',103) "&_
		 " and FromApplication is Null "

end if

IF CStr(sBookCode) = "07" or CStr(sBookCode) = "06" Then

	IF CStr(sVouchTy) <> "A" Then
		sQuery = sQuery &" and isNull(BankInstrumentType,'N') = '"&sVouchTy&"'"
	End IF

	IF CStr(iPartyType) <> "" Then
		sQuery = sQuery &" and PartyType = '"&iPartyType&"' "
	End IF

	IF CStr(iParSubTy) <> "0" Then
		sQuery = sQuery &" and PartySubType = "&iParSubTy&" "
	End IF

	IF CStr(iPartyCode) <> "0" Then
		sQuery = sQuery &" and PartyCode = "&iPartyCode&" "
	End IF

End IF

'Response.Write sQuery

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing


If not objRs.EOF then
%>
<Input type="hidden" name="hCheckTy" value="1">
<%

	if objRs.RecordCount>20 then
%>
 <select size="20" name="selAccountHead"  onDblclick="finaldone()" class="FormElem">
<%	else%>
 <select size="<%=objRs.RecordCount%>" name="selAccountHead"  onDblclick="finaldone()" class="FormElem">
<%
	end if

	While Not objRs.EOF
		sParCode = objRs(0)
		sAmount = objRs(3)
		IF CStr(sTransType) = "SJR" Then
			Response.Write("<OPTION VALUE="&objRs(0)&"--"&objRs(4)&"--"&Objrs(5)&">"&_
				""&trim(objRs(1))&" -- "&objRs(2)&" -- "&sAmount&"</OPTION>")
		Else
			Response.Write("<OPTION VALUE="&objRs(0)&"--"&Objrs(5)&">"&_
				""&trim(objRs(1))&" -- "&objRs(2)&" -- "&sAmount&"</OPTION>")
		End IF
		objRs.MoveNext
	wend
Response.Write "</select>"
Else
%>
<Input type="hidden" name="hCheckTy" value="0">
<%
end if
objRs.Close
%>

</select>

                            </td>
                                </tr>
                                    </table>
                                      </center>
                                    </div>


								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Done" name="B7" onclick="finaldone()" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="B8" onClick="finalcancel()" class="ActionButton">
                                                                 <input type="reset" value="Reset" name="B9" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                                </tr>
						</table>
		</td>
	</tr>
</table>
  </center>
</div>
</form>
</BODY>
</HTML>
