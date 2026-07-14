<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	AssetItmListEntry.asp
	'Module Name				:	Inventory (Master)
	'Author Name				:	S.UMAMAHESWARI
	'Created On					:	20 December 2010
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Purpopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="UnitData"><root></root></script>
<%
Dim sQuery,sUnitID,tempFinYear,dtFromDate,dtToDate,sGCode,sGName,sPGroup,nPageSize
Dim iPageSize,nPageCtr,iTotalRecords,iPageNo,iTotalPages,iPrevPage,iNextPage,iRecCtr
Dim sAssetCategoryCode,sFromDate,sToDate,nSlNo,iStartRec,iEndRec,nCatCode
Dim nAssetCode

Dim objRS
set objRS = Server.CreateObject("ADODB.RecordSet")

sUnitID   = Request("selUnitId")
nCatCode = Request("hCategory")
nAssetCode = Request("hAssetCode")

iPageSize=20

if trim(sUnitID) = "" then	sUnitID = Session("organizationcode")

tempFinYear = split(Trim(Session("FinPeriod")),":")
dtFromDate = "01/04/" & trim(tempFinYear(0))
dtToDate = trim(day(formatDate(date()))) + "/" + trim(month(formatDate(date()))) + "/" + trim(year(formatDate(date())))


iPageNo=trim(Request("hPage"))
if iPageNo="" then iPageNo=1

%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/assetItemList.js"></SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setDate()">

<form method="POST" name="formname" action="" >
	<input type="hidden" name="hOrgId" value="<%=sUnitID%>">
	<input type="hidden" name="hFromDate" value="<%=dtFromDate%>">
	<input type="hidden" name="hToDate" value="<%=formatDate(dtToDate)%>">
	<input type="hidden" name="hCreatedFromDate" value="<%=sFromDate%>">
	<input type="hidden" name="hCreatedToDate" value="<%=formatDate(sToDate)%>">
	<input type="hidden" name="hCategory" value="<%=nCatCode%>">
	<input type="hidden" name="hAssetCode" value="<%=nAssetCode%>">
	<input type=hidden name="hPage" value="">
	<input type="hidden" name="hTempItemname" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Asset To Item
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
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
<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="center" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="7">
</td>
</tr>

<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Unit Name</td>
	<td class="FieldCellSub" colspan="2">
	<select size="1" name="selUnitId" class="FormElem"   >
		<%populateUnit(sUnitID)	%>

	</select>
	</td>
</tr>

<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Asset Category</td>
	<td class="FieldCellSub" colspan="2"><select size="1" name="selAssetType" class="FormElem" onChange="popAsset()">
				<option value="0">Select Asset</option>
				<option value="C">All Assets Specific Category</option>
				<option value="A">Specific Asset Specific Category</option>
		</select>
	</td>
</tr>
<tr>
	<td class="FieldCellSub"></td>
   <td class="FieldCellSub"></td>
   <td class="FieldCellSub" width="0" colspan="2"><span class="dataonly" id="spCatName"></span>
   </td>
</tr>
<!--<tr>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub">Asset Category</td>
	<td class="FieldCellSub" colspan="2">
	<select size="1" name="selAssetCategory" class="FormElem"   >
		<option value="">Select</option>
		<%
		sQuery="SELECT AssetCategoryCode,AssetCategoryName,CategoryParentGroup FROM Far_M_AssetCategory ORDER BY AssetCategoryCode"

		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		set objRs.ActiveConnection = nothing

		set sGCode = objRs(0)
		set sGName = objRs(1)
		set sPGroup = objRs(2)

		If not objRs.EOF then
			Do while not objRs.EOF%>
					<option value="<%=sGCode%>"><%=sGName%></option>
				<%objRs.MoveNext
			Loop
		End IF
		objRs.close
		%>
	</select>
	</td>
</tr>-->

<tr>
	<td class="FieldCellsub"></td>
	<td class="FieldCellSub"></td>
	<td class="FieldCellSub" colspan="2">
        <p align="Right"><input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
        </p>
    </td>
	<td class="FieldCellSub" ></td>

	<td class="FieldCell">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ResetData()" >
	</td>
</tr>

</table>
</div>
</td>
</tr>
<tr>
<td align="center" class="MiddlePack">
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
<table border="0" cellspacing="1" class="ExcelTable" width="100%">

<tr>
	<td class="ExcelHeaderCell" align="center" width="10" >S.No.
	</td>
	<td class="ExcelHeaderCell" align="center" width="10"><a style="width: 1em; height: 1em;" title href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
		<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Expands this section for more search criteria." width="15" height="15">
		</a>
	</td>
	<td class="ExcelHeaderCell" align="center" width="100" >Asset Group
	</td>
	<td class="ExcelHeaderCell" align="center" width="100">Asset Name
	</td>

	<!--<td class="ExcelHeaderCell" align="center" width="50" >Quantity
	</td>
	<td class="ExcelHeaderCell" align="center" width="70"> Effective Date
	</td>
	<td class="ExcelHeaderCell" align="center" width="70">Effective Value
	</td>
	<td class="ExcelHeaderCell" align="center" width="70" >Depriciation Method
	</td>-->

</tr>
 <%
	'sQuery =" Select M.AssetCategoryName,convert(char,D.EffectiveDate,103),D.AssetDescription,"&_
	'		" D.NumberOfUnits,D.DepreciationMethod,DM.DeprtnShortName,D.AssetNumber,D.EffectiveValue,"&_
	'		" D.AssetDescID,D.SentToInventory,D.InPhysicalInventory from Far_T_AssetDetails D,Far_M_AssetCategory M,"&_
	'       " Far_M_DepreciationMethods DM Where D.AssetCategoryCode = M.AssetCategoryCode "&_
	'       " and DM.DepreciationMethod = D.DepreciationMethod "

	sQuery = " SELECT DISTINCT V.ASSETDESCID,V.ASSETDESCRIPTION,M.AssetCategoryName FROM "&_
			 " VWASSETSITEM V,Far_M_AssetCategory M  WHERE V.AssetCategoryCode = M.AssetCategoryCode "

	If Trim(nCatCode) <> "" then
		sQuery = sQuery & " AND M.AssetCategoryCode IN (select Distinct AssetcategoryCode From Far_M_AssetCategory Where categoryParentGroup = " & nCatCode &" or AssetcategoryCode = " & nCatCode &" )"
	End IF
	If Trim(nAssetCode) <> "" then
		sQuery = sQuery & "AND V.AssetNumber in ("& nAssetCode &") "
	End IF

	sQuery = sQuery &"Order By M.AssetCategoryName"

	'Response.Write "<p>sQuery = "& sQuery

	With objRS
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.PageSize=iPageSize
		.Source = sQuery
		.Open
	End With
	Set objRS.ActiveConnection = Nothing
	nSlNo = 1
	iRecCtr = 1
	IF not objRS.EOF then
		iTotalPages = objRS.PageCount
		iTotalRecords = objRS.RecordCount
		Objrs.AbsolutePage = iPageNo
	Else
		iTotalPages = 0
		iTotalRecords = 0
		iStartRec = 0
		iEndRec = 0
	End If

	if trim(iPageNo) = 1 then
		iPrevPage = 0
	else
		iPrevPage = iPageNo - 1
	end if

	if iTotalPages >= iPageNo + 1 then
		iNextPage = iPageNo + 1
	else
		iNextPage = 0
	end if

	Do while not Objrs.EOF  and nSlNo <= Objrs.PageSize
	%>
		<tr>
			<td class="ExcelSerial" align="center" ><%=nSlNo%></td>
			<td class="ExcelDisplayCell" align="center" width="10">
				<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=objRS(0)%>">
			</td>
			<td class="ExcelDisplayCell" align="Left"><%=Objrs(2)%></td>
			<td class="ExcelDisplayCell" align="left"><%=Objrs(1)%></td>
			<!--<td class="ExcelDisplayCell" align="Right"><%'=Objrs(3)%></td>
			<td class="ExcelDisplayCell" align="left"><%'=Objrs(1)%></td>
			<td class="ExcelDisplayCell" align="Right"><%'=Objrs(7)%></td>
			<td class="ExcelDisplayCell" align="Left"><%'=Objrs(5)%> </td>
			<input type="hidden" name="ChkBoxValue<%'=nSlNo%>" value="<%'=Objrs(8)%>:<%'=Objrs(2)%>:<%'=Objrs(9)%>:<%'=objRS(10)%>">	-->
		</tr>

	<%	nSlNo = nSlNo + 1
		iRecCtr = iRecCtr + 1
		Objrs.MoveNext
	loop
	Objrs.Close

	%>
<input type=hidden name="hCnt" value="<%=iRecCtr-1%>">
</table>
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
<td valign="top" align="right">

<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

<%if trim(iPrevPage) = "0" then  %>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
<%else%>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
<%end if %>


<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

<%for nPageCtr= 1 to iTotalPages %>
	<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
<%next%>

</SELECT>
<%if trim(iNextPage) = "0" then  %>
	<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
<%else%>
	<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
<%end if%>

<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

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
<!--<select size="1" name="Choice" class="FormElem">
<option Value="SEL"> Select </option>
<option Value="CRN"> Create New </option>
<option Value="EDT"> Edit Selected </option>
<option Value="ASC"> Asset Registor </option>
<option Value="ATS"> Add To Stock </option>
<option Value="DEP"> Depreciation </option>
</select>
<Input type="button" value="Proceed" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >-->
<Input type="button" value="Create Item" name="ButOpt" class="ActionButtonX" tabindex="3" onclick="CheckSubmit()" >
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
</td>
</tr>

</table>
</form>
</body>
</html>

