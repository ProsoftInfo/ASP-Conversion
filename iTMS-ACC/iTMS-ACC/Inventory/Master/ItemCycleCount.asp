<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ItemCycleCount.asp
	'Module Name				:	INVENTORY (Item Cycle Count)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:	Aug 28,2013
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
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!--#include File="../../include/getCurrentDate.asp"-->
<%
	'Declaring Variables
	Dim dcrs,dcrs1,oDOM
	Dim sOrgCode,sOrgName,sTempMonYr,sMonYr,arrFin,sFinFrom,sFinTo,sRequest
	Dim sTempVal,sCCDate,sMode,sQuery,dtCycleCount
	Dim iCycleCountEntryNo,iCnt,iSelItemCode,iCCQty
	Dim ndRoot,ndItem
	
'	Response.write "<font color=red>"
'	Response.write Request.Form
	'Response.write "<p>Request(Temp) = "& Request("sTemp")

	set dcrs=server.CreateObject("ADODB.Recordset")
	set dcrs1=server.CreateObject("ADODB.Recordset")
	set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	set ndRoot = oDOM.createElement("Root")
	oDOM.appendChild ndRoot
	iSelItemCode = Request("hSelectItem")
	
	sRequest = replace(Request("sTemp"),"/","-")
	if trim(sRequest)="" or IsNull(sRequest) then sRequest=Request("hTemp")
	if trim(sRequest)<>"" then
	    sTempVal = split(sRequest,":")
	    iCycleCountEntryNo = sTempVal(0)
	    sCCDate = sTempVal(1)
	    sOrgCode = sTempVal(2)
	    sMode = sTempVal(3)
	end if 
	'Response.write "<p>Mode="&sMode
	sOrgCode = Session("organizationcode")
	sOrgName = trim(GetUnitDesc(sOrgCode))
	
	sTempMonYr = mid(FormatDate(date()),4,2)
	sMonYr = sTempMonYr&Year(FormatDate(date()))

	arrFin = split(GetFinancialYear(sMonYr),":")
	sFinFrom = arrFin(0)
	sFinTo = arrFin(1)
	
	if trim(iCycleCountEntryNo)<>"" then
	    sQuery = "Select Convert(varchar,CycleCountDate,103) from INV_T_ItemCycleCount where CycleCountEntryNO = "& iCycleCountEntryNo
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        dtCycleCount = dcrs(0)
	    end if
	    dcrs.close
	else
	    dtCycleCount = getCurrentDate()
	end if


%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Item Grid</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemDetails" data-src="<%="../temp/master/ItemCycleCount"&session.sessionId&".xml"%>"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="CategoryData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempItem"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="CycleCount"><Root></Root></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemCycleCount.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  onload="Init();" >
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
		<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>"/>
		<input type="hidden" name="hCycleCountEntryNo" value="<%=iCycleCountEntryNo%>" />
		<input type="hidden" name="hTemp" value="<%=sRequest%>">
		<input type="hidden" name="hSelectItem" value="<%=iSelItemCode%>">
		<input type="hidden" name="hCycleCountDate" value="<%=dtCycleCount%>">
		
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" height="20">
				     <table>
			            <tr>
			                <td class="PageTitle" >
			                    <p align="center">
			                    Cycle Counting
			                    </p>
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
							<td class="TabBodyWithTopLine">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="center" colspan="3" class="MiddlePack" height="7">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel"></td>
										<td valign="top">
										    <table width="30%"> 
										        <tr>
										            <td class="FieldCellSub">
										                Cycle Count Date
										            </td>
										            <td class="FieldCellSub">
										                <input type="text" id="ctlCCDate" name="ctlCCDate" class="formelem itms-date-picker" data-itms-datepicker="1" size="10">
										            </td>
										        </tr>
										    </table>
											
											
										</td>
										<td align="center" width="5" class="ClearPixel"></td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel"></td>
										<td valign="top">
											<div style="height:400px;">
											<table id="tblItem" border="0" cellspacing="1" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													    <img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onclick="DeleteItem()"/>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													Item Code
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													Description
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
													Classification Name
													</td>
													<td class="ExcelHeaderCell" align="center" colspan="2">Current Stock</td>
													<td class="ExcelHeaderCell" align="center" colspan="2">Cycle Count</td>
												</tr>
												<tr>
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
													<td class="ExcelHeaderCell" align="center" >Value</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center">View</td>
												</tr>
												<%
												Response.write "<font color=red>"
												Dim iCount 
												iCount = 0
												    sQuery = "SELECT DISTINCT VI.COMPANYITEMCODE,VI.ITEMDESCRIPTION,(Select GroupName from INV_M_Classification where GroupCode in "&_
												             " (VI.CLASSIFICATIONCODE)) Classification,VS.YEARCLOSINGSTOCK,VI.STORESUOM,VI.ITEMCODE,VI.CLASSIFICATIONCODE, "&_
												             " isNull(VS.YEARCLOSINGVALUE,0) FROM VW_INV_Items VI, "&_
												             " VwYearlyStock VS WHERE VI.ORGANISATIONCODE = '"& sOrgCode &"' AND VI.ITEMACTIVE = 'Y' AND VI.ITEMONHOLD = 0 AND VI.ITEMCODE = VS.ITEMCODE"&_
												             " AND VI.CLASSIFICATIONCODE = VS.CLASSIFICATIONCODE AND VI.ORGANISATIONCODE = VS.ORGANISATIONCODE AND CONVERT(Datetime,VS.FinancialYearFrom,103) ="&_
												             " CONVERT(Datetime,'"& sFinFrom &"',103) AND CONVERT(Datetime,VS.FinancialYearTo,103) = CONVERT(Datetime,'"& sFinTo &"',103) "
												             
											        if trim(iSelItemCode)<>"" then 
											            sQuery = sQuery & " and VI.ItemCode in ("& replace(iSelItemCode,"?",",") &")"
											        else
											            if trim(iCycleCountEntryNo)<>"" then
											                 sQuery = sQuery & " and VI.ItemCode in (Select ItemCode from INV_T_ItemCycleCountHistory where CycleCountEntryNo in ("& iCycleCountEntryNo &"))"
											            end if 'if trim(iCycleCountEntryNo)<>"" then         
											        end if
											       ' Response.write "<p>"& sQuery
											        
											        dcrs.open sQuery,con
											        if not dcrs.eof then
											            do while not dcrs.eof
											            iCount = iCount + 1
											                set ndItem = oDOM.createElement("Item")
			                                                    ndItem.SetAttribute "EntryNo", iCount
			                                                    ndItem.SetAttribute "CompanyItemCode",dcrs(0)
			                                                    ndItem.SetAttribute "ItemCode",dcrs(5)
			                                                    ndItem.SetAttribute "ClassCode",dcrs(6)
			                                                    ndItem.SetAttribute "ItemName",replace(dcrs(1),"-"," ")
			                                                    ndItem.SetAttribute "ClassName",dcrs(2)
			                                                    ndItem.SetAttribute "StoresUoM",dcrs(4)
			                                                    ndItem.SetAttribute "ItemStock",dcrs(3)
			                                                    ndItem.setAttribute "ItemValue",dcrs(7)
			                                                    ndRoot.appendChild ndItem
			                                                    
			                                                    
			                                                    if trim(iCycleCountEntryNo)<>"" then
												                    sQuery = "Select CycleCountStock from INV_T_ItemCycleCountHistory  where CycleCountEntryNo = "& iCycleCountEntryNo &" and ItemCode = "& dcrs(5)
												                    'Response.write "<p>"& sQuery
												                    dcrs1.open sQuery,con
												                    if not dcrs1.eof then
												                        iCCQty = dcrs1(0)
												                    else
												                        iCCQty = "0"
												                    end if
												                    dcrs1.close
												                    ndItem.setAttribute "CCStock",iCCQty
												                else
												                    ndItem.setAttribute "CCStock","0"
												                end if 'if trim(iCycleCountEntryNo)<>"" then
			                                                    
			                                                    %>
												                    <tr>
												                        <td class="ExcelSerial" align="center"><%=iCount%></td>
												                        <td class="ExcelDisplayCell" align="center">
												                            <input type="checkbox" name="Chkbox<%=iCount%>" value="<%=dcrs(5)%>:<%=dcrs(6)%>:<%=sOrgCode%>:<%=dcrs(0)%>:<%=replace(replace(replace(dcrs(1)," ","-"),chr(34),"~~"),chr(39),"~")%>:Y:0:<%=dcrs(3)%>:<%=dcrs(7)%>" />
												                        </td>
												                        <td class="ExcelDisplayCell"><%=dcrs(0)%></td>
												                        <td class="ExcelDisplayCell"><%=dcrs(1)%></td>
												                        <td class="ExcelDisplayCell" align="center"><%=dcrs(2)%></td>
												                        <td class="ExcelDisplayCell" align="right"><%=dcrs(3)%></td>
												                        <td class="ExcelDisplayCell" align="right"><%=dcrs(7)%></td>
												                        <td class="ExcelDisplayCell" align="right">
												                            <input type="text" name="txtCQty<%=iCount%>" value="<%=iCCQty%>" class="FormElem" style="text-align:right;" size="12" />
												                        </td>
												                        <td class="ExcelDisplayCell" align="center">
												                            <img border="0" src="../../assets/images/iTMS%20Icons/DetailsIcon.gif" width="10" height="10" onclick="ViewCycleCount('<%=iCount%>')"/>
												                        </td>
												                    </tr>    
												                <%
											                dcrs.movenext
											            loop
											        end if
											        dcrs.close
											        oDOM.save(Server.Mappath("../temp/master/ItemCycleCount"&session.sessionId&".xml"))
												
												    
												%>
											</table>
											<input type="button" name="btnAddItem" value="Add Item" onclick="AddItem()" class="AddButtonX" />
											</div>
										</td>
										<td align="center" class="ClearPixel" width="5"></td>
									</tr>
									<tr>
										<td align="center" class="MiddlePack" colspan="3">
										<Input Type=Hidden name="hCtr" Value="<%=iCount%>" >
										</td>
									</tr>
									<tr>
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
													    <input type="button" name="btnSave" value="Save" class="ActionButtonX" onclick="CheckSubmit()" />
													</p>
													
												</td>
											</tr>
										</table>
									</td>
									<td align="center" class="ClearPixel" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td align="center" class="BottomPack" colspan="3"></td>
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



 %>

