<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	TdsGroups.asp
	'Module Name				:	ACCOUNTS (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 05, 2010
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
	Dim sSentBy,sSentToVoucher,iSno
	
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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<Script Language=vbscript>
Function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.submit()
End Function

Function Validate()
	document.formname.submit 
End Function

Function Submit(sType)
	Dim nCnt,nNoOfSelRec,sTempVal
	
	nCnt = document.formname.hCnt.value
	
	If sType = "E" or sType = "D" Then
		nNoOfSelRec = 0
		If nCnt <> "1" Then
			For k = 0 to nCnt-1
				If document.formname.chkBox(k).checked Then
					If sType = "E" Then
						sTempVal = document.formname.chkBox(k).value
					Else
						sTempVal = sTempVal & "," & document.formname.chkBox(k).value
					End IF
					nNoOfSelRec = nNoOfSelRec + 1
				End IF
			Next
			If sType = "D" Then
				If sTempVal <> "" Then sTempVal = Mid(sTempVal,2)
			End If
		Else 
			sTempVal = document.formname.chkBox.value
			If document.formname.chkBox.checked Then
				nNoOfSelRec = nNoOfSelRec + 1
			End IF
		End If
		
		If nNoOfSelRec  > 1 and sType = "E" Then 
			alert("Select any one record for edit")
			Exit Function
		End IF
		If nNoOfSelRec = "0" Then
			alert("Select any one record for edit")
			Exit Function
		End IF
		
	End IF
	'alert(sTempVal)
	
	If sType = "C" or sType = "E" Then
		document.formname.action = "TDSGroupingSetup.asp?CallType="&sType&":"&sTempVal
	Else
		document.formname.action = "TDSGroupingDelete.asp?sGrpID="&sTempVal
	End IF
	document.formname.submit 
End Function
</script>
<script language="javascript">
window.__itmsPopupCompat = { type: "tdsGroupsList" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="">
	<form method="POST" name="formname" >
	
	<input type=hidden name="hUnitNo" value="<% =sUnitID%>">
	<input type=hidden name="hUnitName" value="<% =session("orgShortName")%>">
	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sTDate%>">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hsentBy" value="<%=sSentBy%>">
	<Input type="hidden" name="hPartyCode" value="">
	<input type="hidden" name="hDelFrom" value="M">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle">
				<p align="center">TDS Groups
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
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
	</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
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


<!--<tr>
	<td class="FieldCellsub">Party</td>
	<td class="FieldcellSub"> 
		<span id="PartyName" class="Dataonly"></span>
		<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" onclick="SelParty()"></a>
	</td>
</tr>-->

<tr>
<td class="FieldCell"></td>
<td class="FieldCell" colspan=2 align=center>
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
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table cellspacing="1" class="ExcelTable" width="100%" >
	<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
	<td class="ExcelHeaderCell" align="center">
		<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="Submit('D')">
		</a>
	</td>
	<td class="ExcelHeaderCell" align="center">TDS Group Name</td>
	<td class="ExcelHeaderCell" align="center">Description</td>
	<td class="ExcelHeaderCell" align="center">Status</td>
</tr>
<%
sSql  = "Select Distinct GroupID,GroupName From ACC_M_TDSGroup where isNull(Useable,'Y') = 'Y'"

With Objrs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sSql 
	.Open 
End With
iSno = 0
Do While Not Objrs.EOF 
	iSno = iSno + 1
	%>
		<tr>
			<td class="ExcelSerial" align="center"><%=iSno%></td>
			<td class="ExcelDisplaycell" align="center">
				<Input type="checkbox" name="chkBox" value="<%=objrs(0)%>">
			</td>
			<td class="ExcelDisplaycell" align="Left"><%=Objrs(1)%></td>
			<td class="ExcelDisplaycell" align="Left">&nbsp</td>
			<td class="ExcelDisplaycell" align="Left">&nbsp</td>
		</tr>
	<%
	
	Objrs.MoveNext 
Loop
Objrs.Close 
%>
</table>
<Input type="hidden" name="hCnt" value="<%=iSno%>">
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
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
<td align="center" class="ClearPixel" width="5">
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
	<tr>
		
		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="middle" class="ActionCell">
                        <p align="center">
                         <input type="button" value="Create" class="ActionButton"  id=button1 name=button1 OnClick="Submit('C')">
                         <input type="button" value="Edit" class="ActionButton"  id=button2 name=button2 OnClick="Submit('E')">
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
