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
	'Program Name				:	PaymtReqTypeSelection.asp
	'Module Name				:	Accounts (Transaction)
	'Author Name				:	UmaMaheswari S
	'Created On					:	05 April 2010
	'Modified By				:	
	'Modified On				:	
	'Modified By				:   
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher - Instrument Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<xml id="VoucherData" >
<Root></Root>
</xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=Vbscript>
Function CheckSubmit()
	Dim sReqType,sReqName
	
	sReqType = document.formname.selReqType.value 
	
	If sReqType = "S" or sReqType = "" Then
		alert("Select Requisition Type")
		document.formname.selReqType.focus 
		Exit Function
	End IF
	'sReqName = document.formname.selReqType(document.formname.selReqType.selectedIndex).text 
	document.formname.hReqType.value = sReqType
	window.close 
End Function

Function window_onunload()
	window.returnValue = document.formname.hReqType.value 
	window.close 
End Function
</SCRIPT>
<script language="javascript">
(function () {
	"use strict";
	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}
	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}
	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		if (id && window.opener && window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
		}
	}
	window.CheckSubmit = function () {
		var form = document.formname;
		var reqType = form && form.selReqType ? form.selReqType.value : "";
		if (reqType === "S" || trim(reqType) === "") {
			alert("Select Requisition Type");
			if (form && form.selReqType && form.selReqType.focus) {
				form.selReqType.focus();
			}
			return;
		}
		form.hReqType.value = reqType;
		returnValue(reqType);
		window.close();
	};
	window.onunload = function () {
		returnValue(document.formname && document.formname.hReqType ? document.formname.hReqType.value : "");
	};
}());
</script>
<%
Dim sUnitId,sBookId,sUnitName,sBookName,sAccNo,sAccType,sChkN,sChkH,sChkC
Dim sQry,objrs,objrs1,oDOM,iEntNo,sDrawnOn,sPayAt,iStartNo,iEndNo,dtIssueDate,sStatus

Set objrs = Server.CreateObject("ADODB.Recordset")


%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
	<Input type="Hidden" Name="hReqType" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Requisition Type
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">

													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                                <table cellpadding="0" cellspacing="0">
                                            <tr>
                                        <td>
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                    <tr>
										<td class="FieldCellsub" valign="Top">Requisition Type</td>
										<td class="FieldcellSub"> 
											<select Name="selReqType" class="FormElem" size="6">
												<!--<option value="S">Select a Request Type</option>-->
									    		<option value="H">Hire Purchase</option>
												<option value="L">Loan</option>
												<option value="B">Blank Cheque</option> 
												<option value="A">Regular Payment Cheque</option>
												<option value="O">Regular Payment Chash</option>
												<option value="V">Advance</option>
											</Select>
										</td>
									</tr>

                                    <tr>
                                <td class="MiddlePack" colspan="4"><p align="left"></td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                                </table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
