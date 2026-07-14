<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	STORELOCATIONS.ASP
	'Module Name				:	 
	'Author Name				:	S.Maheshwari
	'Created On					:	July 3,2008
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 07,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include virtual="/include/Purpopulate.asp"-->
<!--#include virtual="/include/purItemCommon.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/PurChkItemSpecPack.asp"-->


<!DOCTYPE HTML PUBLIC "-//w3c//DTD HTML 4.0 Transitional//EN">
<html><head><title>iTMS</title>
<meta http-equiv="Content-Type" Content="text/html; charset=ISO-8859-1">
<meta Content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"> </SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"> </SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"> </SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/VouTransactions.js"> </SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/storeLocations.js"> </SCRIPT>
</head>
<%
'Declaring variables
dim sSql,rs
dim iPrevPage,iTotalPages,nPageCtr,iPageNo,iNextPage,nPageSize,iTotalRecords,iStartRec,iEndRec
dim sUnit
Set rs = Server.CreateObject("ADODB.Recordset")	
nPageSize = 20
iPageNo=trim(Request("hPage"))
if trim(iPageNo)="" then iPageNo=1

'sUnit		= Request.Form("hUnit") 
sUnit		= Session("organizationcode")
sLocName	= Request.Form("hLocName") 
sAppFor		= Request.Form("hAppfor") 

IF trim(sUnit) = "" then sUnit = "010101"
IF trim(sAppFor) = "" then sAppFor = "S"
'Response.Write "sUnit="&sUnit &"**"&sLocName &"**"& sAppFor
%>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="FnInit()">

	<form method="POST" name="formname">
	<input type="hidden" name="hPage" value="">
	<input type="hidden" name="hUnit" value="<%=sUnit%>">
	<input type="hidden" name="hLocName" value="<%=sLocName%>">
	<input type="hidden" name="hAppfor" value="<%=sAppFor%>">
	<input type="hidden" name="hPara" value="">
	

	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" height="20">
			    <table>
		            <tr>
		                <td class="PageTitle" >
		                    <p align="center">Storage Locations
		                </td>
		                <td class="PageTitle" >
		                    <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
		                </td>
		            </tr>
			    </table>
				
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
                                 <td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
										<tr><a href="OrgControlDefn.asp">
											<td width="100%" align="center">Org Defn
											</td>
											</a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
								   <tr>
									  <td width="100%" align="center">Storage Location</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MASUOMENTRY.asp">
											<td width="100%" align="center">
                                                UOM
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
							  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
							   <tr><a href="PackingTypes.asp">
								  <td width="100%" align="center">Packing Type</td></a>
								</tr>
							  </table>
							</td>
					        <td class="TabCell" valign="bottom" align="center" width="105">
						          <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
						           <tr><a href="../../NoSeries/InventoryNoSeriesEntry.asp">
							          <td width="100%" align="center">Number Series</td></a>
							        </tr>
						          </table>
						        </td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
                            </tr>
						</table>
					</td>
				    </tr>
				    <tr>
		            <td>
		                <table border="0" cellpadding="0" cellspacing="0" width="100%">
		                    <tr>
						        <td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgPurchaseControlEntry.asp">
											<td width="100%" align="center">
                                                Purchase
											</td></a>
										</tr>
									</table>
								</td>
							
								<td class="TabCell" valign="bottom" align="center" width="82">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgInventoryControlEntry.asp">
											<td width="100%" align="center">
                                                Inventory
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="85">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgInspectionControlEntry.asp">
											<td width="100%" align="center">
                                                Inspection
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="OrgSalesControlEntry.asp">
											<td width="100%" align="center">
                                                Sales
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="OrgManufacturingControlEntry.asp">
									  <td width="100%" align="center">Manufacturing</td></a>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="150">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr><a href="MasSubCOntProcess.asp">
									  <td width="100%" align="center">Sub-Contract Process</td></a>
									</tr>
								  </table>
								</td>
					        <td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
					        </td>
					        </tr>
		                </table>
		            </td>
		        </tr>
					<tr>
						<td class="TabBody">
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
								<td valign="center">				
						<a style="width: 1em; height: 1em;" title="" onclick="Div_OnClick(idUnprocessed,'')" >
						
								<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
								</a>
								</td>
								<td valign="center" class="SubTitle">&nbsp;&nbsp;
								
																
								</td>
							</tr>
				</table>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td width="100%">
									<div id="idUnprocessed" style="display: none">
										<table cellpadding="0" cellspacing="0" class=BodyTable width="100%">
										<tr>
											<td class="MiddlePack">
											</td>
											<td class="MiddlePack" colspan="5">
											</td>
										</tr>

									<!--	<tr>
											<td class="FieldCellSub">&nbsp;&nbsp;</td>
											<td class="FieldCellSub">Organization Unit</td>
											<td class="FieldCellSub" colspan="4">
											<select size="1" name="cmbUnit" class="FormElem" >
											
												<%'To populate unit
												 populateShortUnit("")
												%>
												
											</select>
											</td> 
										</tr>-->

										<tr>
											<td class="FieldCellSub"></td>
											<td class="FieldCellSub">Location Name</td>
											<td class="FieldCellSub" colspan="4" >
											 <input type="text" name="txtLocName" class="FormElem">
											</td>
										  	<td class="FieldCellSub">&nbsp;&nbsp;</td>
											<td class="FieldCellSub">Applicable For</td>
											<td class="FieldCellSub" colspan="4">
											<select size="1" name="cmbAppFor" class="FormElem" onchange="ChngAppFor()" >
											<option value="S">Select</option>
											<%
											sSql = "Select Distinct ApplicableFor from Inv_M_Storage"
											rs.Open sSql,con
											Do while not rs.EOF 
												'
												%>
												<option Value="<%=rs(0)%>"><%=AppForName(rs(0))%></option> 
												<%
												rs.MoveNext 
											loop
											rs.Close
											
											%>
												 
											</select>
											</td> 
										</tr>
										<tr>
											<td class="FieldCell"></td>
											 <td class="FieldCell"></td>
											<td class="FieldCell"></td>
											<td class="FieldCell" >
												<input type="button" value="Go" name="ButGo" class="ActionButton" onclick="Validate()">		
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

<table border="0" cellspacing="1" class="ExcelTable" width="100%">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="10" >
		<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DeleteItems()"></a></td>
		<td class="ExcelHeaderCell" align="center" >Location Name</td>
		<td class="ExcelHeaderCell" align="center" >Applicable For</td>
		<td class="ExcelHeaderCell" align="center" >Type</td>
		
	</tr>
				
<% 
Dim iLocNo,iLocCode,sLocName,sAppFor,iStoFree,iStoBin,iNoOfBins
Dim sAppForName,sTypeName,nSlNo,iRecCtr


sSql = "Select LocationNumber,LocationCode,LocationName,ApplicableFor,StorageTypeFree,"&_
	   "StorageTypeBins,NumberOfBins from Inv_M_Storage where OUDefinitionID = '"&sUnit&"'"

If trim(sLocName) <> "" then 
	sSql = sSql &" and LocationName = '"&sLocName&"' " 	
End If
If trim(sAppFor) <> "S" then 
	sSql = sSql &" and ApplicableFor = '"&sAppFor&"' " 	
End If
'Response.Write sSql
with rs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.PageSize = nPageSize
	.Open 
end with
set rs.ActiveConnection = nothing
nSlNo=1
iRecCtr=1
''''Pagination''
If not rs.EOF Then
	iTotalPages = rs.PageCount
	iTotalRecords = rs.RecordCount
	rs.AbsolutePage = iPageNo
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
			

do while Not rs.EOF and nSlNo <=rs.PageSize

	 

	iLocNo    = rs(0)
	iLocCode  = rs(1)
	sLocName  = rs(2)
	sAppFor   = rs(3)
	iStoFree  = rs(4)
	iStoBin	  = rs(5)
	iNoOfBins = rs(6)
		 
	sAppForName = AppForName(sAppFor)
	IF cint(iStoFree) = 1 then sTypeName = "Free / Open Area"
	IF cint(iStoFree) = 0 then sTypeName = "Specific Location / Bins"
	
%> 
	<tr>
		<td class="ExcelSerial" align="center" ><%=nSlNo%></td>
		<td class="ExcelDisplayCell" align="center" width="10" >
			<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=sUnit%>:<%=iLocNo%>" >
		</td>
		<td class="ExcelDisplayCell" align="left" ><%=sLocName%> </td>		 
		<td class=ExcelDisplayCell align=left ><%=sAppForName%></td> 
		<td class=ExcelDisplayCell align=left ><%=sTypeName%></td>
	</tr>
	
<%
nSlNo = nSlNo + 1
iRecCtr = iRecCtr + 1
rs.MoveNext
loop
 
rs.Close 	
%>
<input type="hidden" name="hCnt" value="<%=iRecCtr%>">
</table>
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

	<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onclick="AssignPage('1')">
	<% if trim(iPrevPage)="0" then %>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev>
	<% else %>
	<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onclick="AssignPage('<%=iPrevPage %>')">
	<%end if %>
	                

	<SELECT class="FormElem" onChange="AssignPAge(this.value)" id="cmbpage" name="cmbpage">
		
		<% for nPageCtr=1 to iTotalPages %>
		
<option value="<%=nPageCtr%>" <% if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages%> </option>
		<% next %>
		
	</SELECT>
	<% if trim(iNextPage)= "0" then %>
	<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext alt="No next Page">
	<% else %>
	<input type="button" value=" >> " class="ActionButtonX"  onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext>
	<% end if %>
	
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
<input type="button" value="Create" name="ButCr" class="ActionButtonX" onclick="FinalSubmit('CR')">
<input type="button" value="Amend" name="ButAm" class="ActionButtonX" onclick="FinalSubmit('AM')">


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
<%
Function AppForName(Para)
Dim sPara,sName
sPara = Para
	Select Case Cstr(sPara)
		Case "PU"	
			sName = "Purchase"
		Case "OI"	
			sName = "Inspection-Outorder"
		Case "POI"	
			sName = "Inspection-Preorder"	
		Case "SA"	
			sName = "Sales"
		Case "IN"	
			sName = "Inventory"
		Case "PI"	
			sName = "Inspection Process"
		Case "PSI"	
			sName = "Post-Sale"
		Case "MA"	
			sName = "Manufacturing" 
	End Select	
	AppForName = sName
End Function
%>
