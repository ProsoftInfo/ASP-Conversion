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
	'Program Name				:	GlSummVouTy.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/accpopulate.asp"-->
<%
dim sOrgId,objRs,sQuery,sCallTy,Temparr,sSelVal,sTitle
dim sBookCode,iBookNo,sTransType,bFlag,sUnit,iAccHead
Dim iCaCnt,iBaCnt,iGJCnt,iSalCnt,iPurCnt,iDebCnt,iCreCnt

Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallTy=Request("sTempValues")
Temparr = Split(sCallTy,"?")
sCallTy = Temparr(0)
sSelVal = Temparr(1)
sUnit = Temparr(2)
iAccHead = Temparr(3)
'Response.Write sSelVal
Temparr = Split(sSelVal,",")
'sBookCode=Request("BookCode")
iCaCnt = 0
iBaCnt = 0
iGJCnt = 0
iSalCnt = 0
iPurCnt = 0
iDebCnt = 0
iCreCnt = 0


if trim(iAccHead)<>"" then
	sQuery = "Select Count(1),BookCode From Acc_T_GLTransactions Where AccountHead = "&iAccHead&" and   "&_
			 "OUDefinitionID = '"&sUnit&"' Group By BookCode Order By BookCode "
'	Response.Write sQuery
	objRs.Open sQuery,con
	Do While Not objRs.EOF
		Select Case objRs(1)
			Case "01"
				iCaCnt = objRs(0)
			Case "02"
				iBaCnt = objRs(0)
			Case "04"
				iPurCnt = objRs(0)
			Case "05" 
				iSalCnt = objRs(0)
			Case "06"
				iCreCnt = objRs(0)
			Case "07"
				iDebCnt = objRs(0)
			Case "08"
				iGJCnt = objRs(0)
		End Select
		objRs.MoveNext
	Loop
	objRs.Close
end if 'if trim(iAccHead)<>"" then






%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Summary Entry Voucher Type Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "voucherTypeSelection" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="CheckVal('<%=sSelVal%>')">
<form method="POST" name="formname" action="">
<Input Type="hidden" Name="hSelBooks" Value="">

<Input Type="hidden" Name="hCaAnt" Value="<%=iCaCnt%>">
<Input Type="hidden" Name="hBaAnt" Value="<%=iBaCnt%>">
<Input Type="hidden" Name="hPurAnt" Value="<%=iPurCnt%>">
<Input Type="hidden" Name="hSalAnt" Value="<%=iSalCnt%>">
<Input Type="hidden" Name="hCreAnt" Value="<%=iCreCnt%>">
<Input Type="hidden" Name="hDebAnt" Value="<%=iDebCnt%>">
<Input Type="hidden" Name="hGJAnt" Value="<%=iGJCnt%>">

<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popuptable">
	<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Select Voucher Type</p>
			</td>
		</tr>
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
                                    <div align="center">
                                      <center>
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell">
							<!--input type="text" name="txtSearch" size="35"  ONKEYUP="selectTheItem(this,'selAccountHead')"  class="FormElem"-->
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell">
	
							<Input type="checkbox" name="chkSelVal" value="01">Cash Book <br>
							<Input type="checkbox" name="chkSelVal" value="02">Bank Book <br>
							<Input type="checkbox" name="chkSelVal" value="04">Purchase Book <br>
							<Input type="checkbox" name="chkSelVal" value="05">Sales Book <br>
							<Input type="checkbox" name="chkSelVal" value="06">Debit Note Book <br>
							<Input type="checkbox" name="chkSelVal" value="07">Credit Note Book <br>
							<Input type="checkbox" name="chkSelVal" value="08">General Journal Book <br>
							
                            </td>
							<input type="hidden" name="hRowCount" value="7">
							 </tr>
							 <tr>
							 <td class="ExcelDisplayCell">* Note Only Entries From Selected Books will come as Summary all others are Individual Entry
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
