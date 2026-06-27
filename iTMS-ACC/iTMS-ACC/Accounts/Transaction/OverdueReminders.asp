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
<script language="javascript" src="../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../scripts/BankVoucher.js"></script>
<script language="javascript" src="../scripts/ReportReminderCompat.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<Script Language=vbscript>
'**************************************************
Function DelSubmit()
    iSelCount = 0
    nRow = document.formname.hCnt.value
    For iCnt = 1 to nRow
       set sObj = eval("document.formname.chkReminderZ"&iCnt)
       if sObj.checked = true then
        iSelCount = iSelCount + 1
            sValue = sObj.value
       end if
    Next
    if iSelCount > 1 or iSelCount = 0 then
        alert("Select Any One Record to Delete")
        exit function
    end if
    if confirm("Do you want to Delete the Reminder Permanently") then
        set objhttp = CreateObject("Microsoft.XMLHTTP")
        objhttp.open "GET","PayRecReminderDelete.asp?RemNo="& sValue,false
        objhttp.send
        if Trim(objhttp.responseText)<>"" then
        else
            alert("Reminder Deleted Successfully")
            document.formname.submit
        end if
    end if
End Function
'**********************************
'***************************************************
Function ViewRem()
     iSelCount = 0
    nRow = document.formname.hCnt.value
    For iCnt = 1 to nRow
       set sObj = eval("document.formname.chkReminderZ"&iCnt)
       if sObj.checked = true then
        iSelCount = iSelCount + 1
            sValue = sObj.value
            sPassVal = eval("document.formname.hPartyDetZ"&iCnt).value
       end if

    Next
    if iSelCount > 1 or iSelCount = 0 then
        alert("Select Any One Record to Print")
        exit function
    end if
    Set OutDataValue = showModalDialog("PartyOutstandingPrevReminder.asp?PassValue="&sPassVal&"&RemNo="&sValue&"&CallFrom=List",OutStandingData,"dialogHeight:450px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No")
End Function
'****************************************************
Function SetDate()
	Dim sFDate,sTDate
	sFDate=document.formname.hFromDate.value
	sTDate=document.formname.hToDate.value

	if Trim(sFDate)<>"" and Trim(sTDate)<>"" then
		document.formname.ctlVouFromDate.setDate=sFDate
		document.formname.ctlVouToDate.setDate=sTDate
	end if

End Function

Function SelParty()
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
	sUnitID = document.formname.hUnitNo.value

	document.formname.hPartyCode.value=""
    PartyName.innerHTML = ""

    Set nodParty = PartyData.documentElement

   		sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)



'	set OutValue = showModalDialog("../Reports/PartySelectPopup.asp?orgId="+sUnitID&"&hSelectMode=M",PartyData,"dialogHeight:550px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'
'
 '   sQuery = OutValue.getAttribute("PassQuery")
  '  if OutValue.getAttribute("Action")="CLOSE" then exit function
   '
    'while OutValue.getAttribute("Action")<>"Done"
     '   set OutValue = showModalDialog("../Reports/PartySelectPopup.asp?"&sQuery,PartyData,"dialogHeight:550px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
    '    sQuery = OutValue.getAttribute("PassQuery")
    '    if OutValue.getAttribute("Action")="CLOSE" then exit function
    'wend

  '  if OutValue.hasChildNodes() then
  '      For each ndChild in OutValue.childNodes
  '          sPartyCode = sPartyCode & "," & ndChild.getAttribute("ParCode")
  '          sPartyName = sPartyName & "," & ndChild.getAttribute("ParName")
  '      Next
  '  end if 'if OutValue.hasChildNodes() then

  Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sUnitID&"&hSelectMode=M",PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

        if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sParCode = sParCode &","& ndEntry.getAttribute("RetField1")
		            sPartyName = sPartyName &","& ndEntry.getAttribute("RetField0")
		        end if
            next
        end if



	if trim(sPartyName)<>"" then
        sPartyCode = mid(sPartyCode,2)
        sPartyName = mid(sPartyName,2)
    end if

    document.formname.hPartyCode.value=sPartyCode
    PartyName.innerHTML = sPartyName
End Function

Function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.submit()
End Function

Function Validate()
	Dim sStatus

	document.formname.hFromDate.value = document.formname.ctlVouFromDate.GetDate
	document.formname.hToDate.value	 = document.formname.ctlVouToDate.GetDate

	If document.formname.RadStatus(0).checked  Then
		document.formname.hsentBy.value  = document.formname.RadStatus(0).value
	Elseif document.formname.RadStatus(1).checked  Then
		document.formname.hsentBy.value  = document.formname.RadStatus(1).value
	End IF

	document.formname.submit
End Function

Function CreateRem()
	document.formname.action = "PartyOutstanding.asp?PartyType=DR"
	document.formname.submit
End Function
</script>
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
		<object id="ctlVouFromDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
	</td>

	<td class="FieldCellSub">To</td>
    <td class="FieldCellSub" valign="middle">
		<object id="ctlVouToDate"  classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
			<param name="_ExtentX" value="2355">
			<param name="_ExtentY" value="529">
		</object>
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
