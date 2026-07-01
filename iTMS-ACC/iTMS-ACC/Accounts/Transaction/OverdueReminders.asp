<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	OverdueReminders.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 07, 2010
	'Modified By                :
	'Modified On                :
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sUnitID,sFrmDate,sTDate,iCnt,sSql,sFromDate,sToDate
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo
	Dim iTotalPages,iTotalRecords,iPrevPage,iNextPage
	Dim sSentBy,sSentToVoucher,iSno,sPartyName,sPartySubType

	Dim Objrs,Objrs1

	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

	sUnitID = Session("organizationcode")

	Const iPageSize=16
	iPageNo = trim(Request("hPage"))
	if trim(iPageNo) = "" then iPageNo = 1

	iCurrentPage=CInt(Request.Form("hPageSelection"))

	sFromDate ="01/"&Mid(GetFromFinYear,1,2)&"/"&Mid(GetFromFinYear,3,4)
	sToDate ="31/03/2011"

	sFrmDate = Request("hFromDate")
	sTDate =  Request("hToDate")

	If sFrmDate = "" Then
		sFrmDate = sFromDate
		sTDate = sToDate
	End IF

	sSentBy = trim(Request("hsentBy"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><PartyType/></XML>
<xml id="PartyData"><Party/></xml>
<XML id="AccHeadData"><account/></XML>
<xml id="OutStandingData"><Root></Root></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<script language="javascript" src="../../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../../scripts/BankVoucher.js"></script>
<script language="javascript" src="../../scripts/ReportReminderCompat.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
	<form method="POST" name="formname" action="OverdueReminders.asp" >

	<input type=hidden name="hUnitNo" value="<% =sUnitID%>">
	<input type=hidden name="hUnitName" value="<% =session("orgShortName")%>">
	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sTDate%>">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hsentBy" value="<%=sSentBy%>">
	<Input type="hidden" name="hPartyCode" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Reminders For Overdue
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>


<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
	</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
	<%
		Response.Write ("<Input type=checkbox name=voutype value=TS checked >To Send&nbsp;")
		Response.Write ("<Input type=checkbox name=voutype value=S >Send&nbsp;")
	%>
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<tr>
	<td class="FieldCellSub">Date From </td>

    <td class="FieldCellSub" valign="middle">
		<input type="text" id="ctlVouFromDate" name="ctlVouFromDate" class="FormElem itms-date-picker" data-itms-datepicker="1" size="10">
	</td>

	<td class="FieldCellSub">To</td>
    <td class="FieldCellSub" valign="middle">
		<input type="text" id="ctlVouToDate" name="ctlVouToDate" class="FormElem itms-date-picker" data-itms-datepicker="1" size="10">
	</td>
</tr>

<tr>
	<td class="FieldCellSub">Party</td>
	<td class="FieldcellSub">
		<span id="PartyName" class="Dataonly"></span>
		<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" onclick="SelParty()"></a>
	</td>
</tr>

<tr>
	<td class="FieldCellSub">Sent By</td>
	<td class="FieldCellSub">
		<input type="Radio" Name="RadStatus" Value="C" class="FormElem">&nbsp;Courier
		<input type="Radio" Name="RadStatus" Value="E" class="FormElem">&nbsp;E-Mail
	</td>
</tr>
<tr>
<td class="FieldCell"></td>
<td class="FieldCell" colspan="2">
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table cellspacing="1px" class="ExcelTable" width="100%" >
<tr>
	<td class="ExcelHeaderCell" width="10px">S.No.</td>
	<td class="ExcelHeaderCell">
	    <a>
		<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15px" height="15px" onclick="DelSubmit()">
		</a>
	</td>
	<td class="ExcelHeaderCell">Sent To
	</td>
	<td class="ExcelHeaderCell">Sent On
	</td>
	<td class="ExcelHeaderCell">Reason
	</td>
	<td class="ExcelHeaderCell">Actions
	</td>
</tr>
<%
Response.Write "<font color=red>"
    iSno = 0
sSql = "Select R.ReminderNo,ReminderToPartyCode,Convert(varchar,ReminderDate,103),ReminderReason,ActionTaken,PartyInvoiceNo "&_
       " from APP_R_ApplicationReminders R,ACC_T_OverDueReminderDet T where R.ReminderNo = T.ReminderNo and R.PartyType = 'DR' "&_
       " Order by ReminderDate Desc"
       'Response.Write "<p>"&sSql
       Objrs.Open sSql,con
       if not Objrs.EOF then
            do while not Objrs.EOF
                iSno = iSno + 1
                sSql = "Select PartyName from APP_M_PartyMaster where PartyCode = "& Objrs(1)
                Objrs1.Open sSql,con
                if not Objrs1.EOF then
                    sPartyName = Objrs1(0)
                end if
                objrs1.Close
                sSql = "Select PartySubType from APP_R_OrgParty where PartyCode = "& Objrs(1) &"  and PartyType = 'DR'"
                objrs1.Open sSql,con
                if not objrs1.EOF then
                    sPartySubType = Objrs1(0)
                end if
                objrs1.Close
            %>
                <tr>
	                <td class="ExcelHeaderCell"><%=iSno%></td>
	                <td class="ExcelDisplayCell">
	                    <input type=checkbox name="chkReminderZ<%=iSNo%>" value="<%=objrs(0)%>">
	                    <input type=hidden name="hPartyDetZ<%=iSNo%>" value="<%=objrs(1)%>:DR:<%=sPartySubType%>:<%=objrs(5)%>">
	                </td>
	                <td class="ExcelDisplayCell" align="Left">
	                    <a href="#" class="ExcelDisplayLink" onclick=""><%=sPartyName%></a>
	                </td>
	                <td class="ExcelDisplayCell"><%=objrs(2)%>
	                </td>
	                <td class="ExcelDisplayCell"><%=objrs(3)%>
	                </td>
	                <td class="ExcelDisplayCell"><%=objrs(4)%>
	                </td>
                </tr>
              <%
                Objrs.MoveNext
            loop
       end if
       Objrs.Close

%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5px">
</td>
</tr>
<input type=hidden name="hCnt" value=<%=iSno%>>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button4 name=button4>
<input type="button" value=" << " class="ActionButtonX" id=button5 name=button5>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button6 name=button6>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button7 name=button7>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td>
	<tr>

		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td class="ActionCell">
                         <input type="button" value="Create Reminder" class="ActionButtonX"  id=button1 name=button1 OnClick="CreateRem()" >
                         <input type="button" value="View Reminder"  class="ActionButtonX"  id="button9" name=button2 OnClick="ViewRem()">
					</td>
				</tr>
			</table>
		</td>

    </tr>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel">
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
</td>
</tr>

</table>
</form>
</body>
</html>
