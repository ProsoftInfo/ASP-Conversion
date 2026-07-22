<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	MaterialConsumption.asp
	'Module Name				:	INVENTORY
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Jun 17,2011
	'Modified By				:
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
    Dim oDOM,Root,Node,GNode
	dim dcrs,dcrs1,iCtr,sAmendedBy,sCheck,sStatus,sOrgID,sOrgName,sAction,iCnt,sSql,sAttId
	dim dFrmDate,dToDate,sType,sUsage,sCreatedBy,sItemType,iItemCode,iClass,sItemName,sClassName
	Dim iTotalPages,iTotalRecords,iStartRec,iEndRec,iPrevPage,iNextPage
	Dim iQtyReceipt,iQtyIssued,iQtyConsumed,iQtyReturned,iQtyBalance,iQtyTotalReceipt
	Dim iQtyTotalIssued,iQtyTotalConsumed,iQtyTotalReturned,iQtyTotalBalance
	Dim sAutoConsumption,sIMType,sCategory,sPartyCode,sPartyNames,sQuery
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	set oDOM = server.CreateObject("Microsoft.XMLDOM")
	dim sFinPeriod,Arr

    sOrgID = Session("organizationcode")
    sOrgName = Session("OrgShortName")

	sCreatedBy = Session("userid")

	Response.Write "<font color=#000000>"
	sItemType = trim(Request.QueryString("ItemType"))
	iItemCode = Request.QueryString("ItemCode")
	iClass = Request.QueryString("ClassCode")
	sAttId = Request.QueryString("AttID")
	sItemType = Request.QueryString("ItemType")
	sIMType = Request.QueryString("IMType")
	sCategory = Request.QueryString("Category")
	sPartyCode = Request.QueryString("PartyCode")

	if trim(sIMType)="" or IsNull(sIMType) then sIMType="I"

	'Response.Write "dToDate = "& dToDate
	if trim(Request.QueryString)="" then
	    sItemType = Request.Form("hItemType")
	    iItemCode = Request.Form("hItemCode")
	    iClass = Request.Form("hClassCode")
	    sItemType = Request.Form("hItemType")
	    sAttId = Request.Form("hAttID")
	    dFrmDate = Request.Form("hFrmDate")
	    dToDate = Request.Form("hToDate")
	end if
	'Response.Write " dToDate = "& dToDate
	if dFrmDate = "" and dToDate="" then
	    sFinPeriod = session("Finperiod")
	    Arr = split(sFinPeriod,":")
	    dFrmDate = "01/04/"& Arr(0)
	    dToDate = "31/03/"& Arr(1)
	end if
'	Response.Write "dToDate = "& dToDate
	if sOrgID = "" then sOrgID = "010101"

	'if sItemType = "" then sItemType = "STO"

	if sType = "select" then sType = ""
	if sCreatedBy = "select" then sCreatedBy = ""

	''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
    Const iPageSize=23	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = Request.Form("hPageSelection")
    if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    'iCtr = (Cint(iPageSize) * (iCurrentPage - 1))

    con.CursorLocation = 3

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if trim(iItemCode)<>"" then
	    sSql = "Select ItemDescription from VWITEM where ItemCode In ("&iItemCode&")"
	    dcrs.open sSql,con
	    if not dcrs.eof then
	        do while not dcrs.eof
	            sItemName =  sItemName &","& trim(dcrs(0))
	            dcrs.movenext
	        loop
	    end if
	    dcrs.close
	end if 'if trim(iItemCode)<>"" then
	if trim(sItemName)<>"" then
	    sItemName = mid(sItemName,2)
	end if

	if trim(iClass)<>"" then
	    sSql = "Select GroupName from INV_M_Classification where GroupCode in ("& iClass &")"
	    dcrs.open sSql,con
	    if not dcrs.eof then
	        do while not dcrs.eof
	            sClassName = sClassName &","& trim(dcrs(0))
	            dcrs.movenext
	        loop
	    end if
	    dcrs.close
	end if
	if trim(sClassName)<>"" then
	    sClassName = mid(sClassName,2)
	end if

	set Root = oDOM.createElement("ROOT")
	oDOM.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CODETYPE,CODETYPENAME,ITEMTYPEID FROM APP_M_CODETYPES ORDER BY DISPLAYORDER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM.createElement("CODE")
		Node.setAttribute "CODEID",trim(dcrs(0))
		Node.setAttribute "CODENAME",trim(dcrs(1))
		Node.setAttribute "ITEMTYPEID",trim(dcrs(2))

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT CODE,DESCRIPTION FROM APP_M_CODEMASTER WHERE CODETYPE = " & trim(dcrs(0)) & " AND ITEMTYPEID = " & Pack(trim(dcrs(2))) & " ORDER BY 2"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			do while Not dcrs1.EOF
				set GNode = oDOM.createElement("GROUP")
				GNode.setAttribute "GROUPCODE",trim(dcrs1(0))
				GNode.setAttribute "GROUPNAME",trim(dcrs1(1))
				GNode.setAttribute "CODEID",trim(dcrs(0))
				Node.appendChild GNode
				dcrs1.movenext
			loop
			dcrs1.Close

		Root.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ITEMTYPEID, ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.Source = "SELECT ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM.createElement("ATTRIBUTES")
		Node.setAttribute "ITEMTYPEID",""'trim(dcrs(0))
		Node.setAttribute "ATTRID",trim(dcrs(0))
		Node.setAttribute "ATTRNAME",trim(dcrs(1))

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OPTIONVALUE,OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE ITEMTYPEATTRIBUTEID = " & trim(dcrs(0)) & " ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		do while Not dcrs1.EOF
			set GNode = oDOM.createElement("GROUP")
			GNode.setAttribute "ATTRID",trim(dcrs(1))
			GNode.setAttribute "OPTIONVALUE",trim(dcrs1(0))
			GNode.setAttribute "OPTIONNAME",trim(dcrs1(1))

			Node.appendChild GNode
			dcrs1.movenext
		loop
		dcrs1.Close

		Root.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close

	if trim(sPartyCode)<>"" then
	    sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode in ("& sPartyCode &")"
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        do while not dcrs.eof
	            sPartyNames = sPartyNames &","& trim(dcrs(0))
	            dcrs.movenext
	        loop
	    end if
	    dcrs.close
	end if 'if trim(sPartyCode)<>"" then

	if trim(sPartyNames)<>"" then sPartyNames = mid(sPartyNames,2)
	if trim(sPartyNames)="" or IsNull(sPartyNames) then sPartyNames ="All Party"


	oDOM.save server.MapPath("../Temp/reports/ItemStockData"&Session.SessionID&".Xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Stock Replenishment</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="Data" data-src="<%="../Temp/reports/ItemStockData"&Session.SessionID&".Xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/materialConsumptionsModern.js"></SCRIPT>
</head>
<%


%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"  onLoad="Init('<%=dFrmDate%>','<%=dToDate%>')">
	<form method="POST" name="formname" action="">
	    <input type=hidden name="hOrgID" value="<%=sOrgID%>">
		<input type=hidden name="hOrgName" value="<%=sOrgName%>">
		<input type=hidden name="hUsage" value="">
		<input type=hidden name="hCheck" value="<%=sCheck%>">
		<input type=hidden name="hClassName" value="">
		<input type=hidden name="hClassCode" value="<%=iClass%>">
		<input type=hidden name="hItemCode" value="<%=iItemCode%>">
		<input type=hidden name="hItemType" value="<%=sItemType%>">
		<input type=hidden name="hAttID" value="<%=sAttId%>">
		<input type=hidden name="hFrmDate" value="<%=dFrmDate%>">
		<input type=hidden name="hToDate" value="<%=dToDate%>">
		<input type="hidden" name="hCategory" value="<%=sCategory%>" />
		<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>" />
		<input type="hidden" name="hTemp" value="<%=Request.QueryString%>" />

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td height="1px"></td></tr>
			<tr>
				<td class="PageTitle">
				    Material Consumption
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
																	<td valign="center"><a style="width: 1em; height: 1em;" title="" onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
																		<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
																		</a>
																	</td>
																	<td valign="center" class="SubTitle">&nbsp;&nbsp;
																	    <input type="radio" name="radType" value="I" class="FormElem" onclick="ListChange()" <%if trim(sIMType)="I" then response.write "checked"%> />Item Wise
																	    <input type="radio" name="radType" value="M" class="FormElem" onclick="ListChange()" <%if trim(sIMType)="M" then response.write "checked"%> />Issue Wise
																	</td>
																</tr>
															</table>
															<table border="0" cellpadding="0" cellspacing="0" class="BodyTable" width=100%>
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="width: 100%; display: none">
																		<table cellpadding="0" cellspacing="0">
	                                                                    <tr>
                                                                                <td class="FieldCellSub">Range From</td>
                                                                                <td class="FieldCellSub" colspan="3">
	                                                                                <object id="ctlFromDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"      codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
													                                    <param name="_ExtentX" value="2355">
													                                    <param name="_ExtentY" value="529">
												                                    </object>
											                                    &nbsp;&nbsp;To&nbsp;
											                                    <object id="ctlToDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"       codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem"  viewastext>
												                                    <param name="_ExtentX" value="2355">
												                                    <param name="_ExtentY" value="529">
											                                    </object>
											                                    </td>
                                                                            </tr>

                                                                            <tr>
                                                                               <!-- <td class="FieldCellSub">Item Type</td>
                                                                                <td class="FieldCellSub">
												                                    <select size="7" name="selIType" class="FormElem" onBlur="GetData()">
													                                    <%	'Calling the Function which populates the Item Type list
														                                   ' populateItemType
													                                    %>
												                                    </select>
                                                                                </td>-->
                                                                                <td class="FieldCellSub">Category</td>
											                                    <td class="FieldCellSub">
												                                    <select size="7" name="selCategory" class="FormElem">
													                                    <%	'Calling the Function which populates the Category list
														                                    populateCategory

													                                    %>
												                                    </select>
											                                    </td>
                                                                            </tr>
                                                                            <tr>
											                                    <td class="FieldCellSub">Classification</td>
											                                    <td class="FieldCellSub" colspan="3">
												                                    <span id="spanClassification" class="DataOnly">
												                                    <%
												                                        if trim(sClassName)<>"" then
												                                            Response.Write sClassName
												                                        else
												                                            Response.Write "All Classifcations&nbsp;"
												                                        end if
												                                    %>
												                                    </span><img src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Item Classifications" onclick="popClass()">
											                                    </td>
                                                                            </tr>
										                                    <!--<tr>
											                                    <td class="FieldCellSub">Attributes</td>
											                                    <td class="FieldCellSub" colspan="3">
										                                              <table border="0" cellspacing="1" class="ExcelTable" id="tblLot">
													                                    <tr>
													                                      <td class="ExcelDisplayCell" align="left" colspan="5">

												                                    <%
													                                    with dcrs
														                                    .CursorLocation = 3
														                                    .CursorType = 3
														                                    '.Source = "SELECT ITEMTYPEATTRIBUTEID,ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES WHERE ITEMTYPEID IN('GAR','FAB','FIB') ORDER BY 1"
														                                    .Source = "SELECT ITEMTYPEATTRIBUTEID,ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 1"
														                                    .ActiveConnection = con
														                                    .Open
													                                    end with
													                                    set dcrs.ActiveConnection = nothing
												                                    '	 Response.Write dcrs.source

													                                    if not dcrs.EOF then
														                                    Do While Not dcrs.EOF

												                                    %>
															                                    <select size="1" name="selAttrZ<%=trim(dcrs(0))%>" class="FormElem" >
															                                    </select>
												                                    <%

														                                    dcrs.MoveNext
														                                    loop
													                                    end if
													                                    dcrs.Close

												                                    %>
													                                      </td>
													                                    </tr>
												                                    </table>
											                                    </td>
										                                    </tr>
										                                    <tr>-->
										                                    <td class="FieldCellSub">Items
										                                    </td>
                                                                            <td class="FieldCellSub" colspan="3">
											                                    <span class="DataOnly" id="idItemName">
											                                    <%if trim(sItemName)<>"" then
											                                        Response.Write sItemName
											                                      else
											                                        Response.Write "All Items&nbsp;"
											                                      end if
											                                    %>
											                                    </span><img src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Items" ONCLICK="Search()">
										                                    </td>
										                                    </tr>
										                                    <tr>
										                                        <td class="FieldCellSub">Party</td>
										                                        <td class="FieldCellSub"><span id="spanParty" class="DataOnly"><%=sPartyNames%>&nbsp;</span>&nbsp;<img src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" ONCLICK="SelectParty()"></td>
										                                    </tr>
										                                    <tr>
										                                        <td class="FieldCellSub">
										                                        </td>
										                                        <td class="FieldCellSub">
										                                            <input type=button name=btnName value="GO" onClick="CheckSubmit('<%=FormatDate(date()) %>')" class="ActionButton" >
										                                        </td>
										                                    </tr>
										                                    <tr>
										                                        <td colspan=2 align=center class="MiddlePack">
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
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
						            <tr>
								        <td align="center">
								        </td>
								        <td valign="top" width="100%" align="left">
                                            <table border="0" cellpadding="0" cellspacing="0" width=100%>
                                                <tr>
											        <td>
												        <DIV class="frmbody" style="width: 100%; height:100%;">
													        <table border="0" cellspacing="1" class="ExcelTable" width="100%">

														        <tr>
												                    <td class="ExcelHeaderCell" align="center" rowspan=2 width="10">S.No.</td>
												                    <%if trim(sIMType)="I" then %>
												                        <td class="ExcelHeaderCell" align="center" rowspan=2>Item Description</td>
												                        <td class="ExcelHeaderCell" align="center" rowspan=2>UoM</td>
												                    <%elseif trim(sIMType)="M" then %>
												                        <td class="ExcelHeaderCell" align="center" rowspan="2" width="10"></td>
												                        <td class="ExcelHeaderCell" align="center" rowspan="2">Issue No - Date</td>
												                    <%end if %>

												                    <td class="ExcelHeaderCell" align="center" Colspan=4>Quantity</td>
											                    </tr>
											                    <tr>
											                        <td class="ExcelHeaderCell" align="center">Issued</td>
												                    <td class="ExcelHeaderCell" align="center">Consumed</td>
												                    <td class="ExcelHeaderCell" align="center">Returned</td>
												                    <td class="ExcelHeaderCell" align="center">Balance</td>
												                </tr>

												         <%
												            if trim(sIMType)="I" then

										                        sSql = " Select D.ItemCode,D.Classificationcode,D.OrganisationCode,SUM(D.QuantityIssued),"&_
										                             " SUM(D.QuantityConsumed),SUM(D.QuantityReturned),D.QuantityUOM from INV_T_MaterialIssueHeader H,INV_T_MaterialIssueDetails D "&_
										                             " where H.IssueEntryNo = D.IssueEntryNo and (H.IssueDate >= Convert(datetime,'"& dFrmDate &"',103))"&_
										                             " and (H.IssueDate <=Convert(datetime,'"& dToDate &"',103))"

										                        If Trim(iItemCode) <> "" and Trim(iItemCode) <> "0" Then
										                            sSql = sSql & " and D.ItemCode IN (" & iItemCode & ") "
									                            End IF

								                    	        If iClass <> "" and iClass <> "0" Then
								                    		        sSql = sSql & " and D.ClassificationCode IN(" & iClass & ")  "
								                    	        End IF

								                    	        if trim(sItemType)<>"" then
								                    	            sSql = sSql & " and H.ItemTypeID='"& sItemType &"'"
								                    	        end if

								                    	        if trim(sCategory)<>"" then
								                    	            sSql = sSql & " and D.ClassificationCode in (Select GroupCode from INV_M_Classification where ParentGroup in (Select GroupCode from INV_M_Classification where GroupCategory in ("& sCategory &")))"
								                    	        end if

								                    	        if trim(sPartyCode)<>"" then
								                    	            sSql = sSql &" and D.ItemCode in (Select ItemCode from INV_T_MaterialIssueDetails where IssueEntryNo in (Select IssueEntryNo from INV_T_MaterialIssueHeader where IssuedToType='Party' and IssuedToCode="& sPartyCode &"))"
								                    	        end if

									                            sSql = sSql & "  Group by D.ItemCode,D.ClassificationCode,D.OrganisationCode,D.QuantityUOM  Order by D.ItemCode"
									                            'Response.Write sSql
									                            dcrs.Open sSql,con
											                    if not dcrs.eof then
											                        dcrs.pagesize = iPageSize
											                        iTotalPages = dcrs.PageCount
			                                                        iTotalRecords = dcrs.RecordCount
			                                                        dcrs.AbsolutePage = iCurrentPage
		                                                        Else
			                                                        iTotalPages = 0
			                                                        iTotalRecords = 0

			                                                        iStartRec = 0
			                                                        iEndRec = 0
		                                                        End If

		                                                        if trim(iCurrentPage) = 1 then
			                                                        iPrevPage = 0
		                                                        else
			                                                        iPrevPage = iCurrentPage - 1
		                                                        end if


		                                                        if iTotalPages >= iCurrentPage + 1 then
			                                                        iNextPage = iCurrentPage + 1
		                                                        else
			                                                        iNextPage = 0
		                                                        end if


											                    If not dcrs.EOF then
   												                    Do while Not dcrs.EOF and iCtr < dcrs.PageSize
   													                    iCtr = iCtr + 1

   													                    iQtyIssued = Trim(dcrs(3))
												                        iQtyConsumed = Trim(dcrs(4))
												                        iQtyReturned = Trim(dcrs(5))

   													                    sSql = "Select ItemDescription from VWITEM where ItemCode = "& dcrs(0) &" and ClassificationCode = "& dcrs(1)
   													                    dcrs1.Open sSql,con
   													                    if not dcrs1.EOF then
   													                        sItemName = dcrs1(0)
   													                    end if
   													                    dcrs1.Close

   													                    iQtyBalance = iQtyIssued - iQtyConsumed - iQtyReturned

   													                    %>
   													                    <tr>
   														                    <td class="ExcelSerial" align="center" ><%=iCtr%></td>
   														                    <td class="ExcelDisplayCell" align="Left"><a href="#" class="ExcelDisplayLink" onclick="MatConsumption('<%=dcrs(0)%>','<%=dcrs(1)%>','<%=dFrmDate%>','<%=dToDate%>')"><%=sItemName%></a></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=trim(dcrs(6))%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyIssued),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyConsumed),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyReturned),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyBalance),2,,,0)%></td>
   													                    </tr>
   													                    <%
   													                    dcrs.MoveNext
												                    Loop
											                    Else
											                    %>
												                    <tr>
													                    <td colspan=8 class="ExcelDisplayCell" align="center" valign="middle"><B> No records To Be Found</B></td>
												                    </tr>
											                    <%End If 'If not dcrs.EOF then
											                    dcrs.Close
											                elseif trim(sIMType)="M" then
											                    sSql = "Select H.IssueEntryNo,H.IssueEntryCode,Convert(varchar,H.IssueDate,103),H.OrganisationCode,"
											                    sSql = sSql & " SUM(D.QuantityIssued),SUM(D.QuantityConsumed),SUM(D.QuantityReturned) from "
											                    sSql = sSql & " INV_T_MaterialIssueHeader H,INV_T_MaterialIssueDetails D where H.IssueEntryNo = "
											                    sSql = sSql & " D.IssueEntryNo and (H.IssueDate >= Convert(datetime,'"& dFrmDate &"',103)) "
											                    sSql = sSql & " and (H.IssueDate <=Convert(datetime,'"& dToDate &"',103)) "


											                    If Trim(iItemCode) <> "" and Trim(iItemCode) <> "0" Then
										                            sSql = sSql & " and D.ItemCode IN (" & iItemCode & ") "
									                            End IF

								                    	        If iClass <> "" and iClass <> "0" Then
								                    		        sSql = sSql & " and D.ClassificationCode IN(" & iClass & ")  "
								                    	        End IF

								                    	        if trim(sItemType)<>"" then
								                    	            sSql = sSql & " and H.ItemTypeID='"& sItemType &"'"
								                    	        end if


											                    if trim(sCategory)<>"" then
								                    	            sSql = sSql & " and D.ClassificationCode in (Select GroupCode from INV_M_Classification where ParentGroup in (Select GroupCode from INV_M_Classification where GroupCategory in ("& sCategory &")))"
								                    	        end if

								                    	        if trim(sPartyCode)<>"" then
								                    	            sSql = sSql &" and D.ItemCode in (Select ItemCode from INV_T_MaterialIssueDetails where IssueEntryNo in (Select IssueEntryNo from INV_T_MaterialIssueHeader where IssuedToType='Party' and IssuedToCode="& sPartyCode &"))"
								                    	        end if


											                    sSql = sSql & " Group By H.IssueEntryNo,H.IssueEntryCode,H.IssueDate,H.OrganisationCode Order By H.IssueEntryNo Desc "

									                            'Response.Write sSql
									                            dcrs.Open sSql,con
											                    if not dcrs.eof then
											                        dcrs.pagesize = iPageSize
											                        iTotalPages = dcrs.PageCount
			                                                        iTotalRecords = dcrs.RecordCount
			                                                        dcrs.AbsolutePage = iCurrentPage
		                                                        Else
			                                                        iTotalPages = 0
			                                                        iTotalRecords = 0

			                                                        iStartRec = 0
			                                                        iEndRec = 0
		                                                        End If

		                                                        if trim(iCurrentPage) = 1 then
			                                                        iPrevPage = 0
		                                                        else
			                                                        iPrevPage = iCurrentPage - 1
		                                                        end if


		                                                        if iTotalPages >= iCurrentPage + 1 then
			                                                        iNextPage = iCurrentPage + 1
		                                                        else
			                                                        iNextPage = 0
		                                                        end if


											                    If not dcrs.EOF then
   												                    Do while Not dcrs.EOF and iCtr < dcrs.PageSize
   													                    iCtr = iCtr + 1

   													                    iQtyIssued = Trim(dcrs(4))
												                        iQtyConsumed = Trim(dcrs(5))
												                        iQtyReturned = Trim(dcrs(6))

   													                    iQtyBalance = iQtyIssued - iQtyConsumed - iQtyReturned

   													                    %>
   													                    <tr>
   														                    <td class="ExcelSerial" align="center" ><%=iCtr%></td>
   														                    <td class="ExcelDisplayCell" align="center">
   														                        <input type="checkbox" name="ChkZ<%=iCtr%>" value="<%=dcrs(0)%>" />
   														                    </td>
   														                    <td class="ExcelDisplayCell" align="Left"><a href="#" class="ExcelDisplayLink" onclick="MatConsumptionIss('<%=dcrs(0)%>','<%=dFrmDate%>','<%=dToDate%>')"><%=dcrs(1)%>-<%=dcrs(2)%></a></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyIssued),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyConsumed),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyReturned),2,,,0)%></td>
   														                    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(Trim(iQtyBalance),2,,,0)%></td>
   													                    </tr>
   													                    <%
   													                    dcrs.MoveNext
												                    Loop
											                    Else
											                    %>
												                    <tr>
													                    <td colspan=8 class="ExcelDisplayCell" align="center" valign="middle"><B> No records To Be Found</B></td>
												                    </tr>
											                    <%End If 'If not dcrs.EOF then
											                    dcrs.Close
											                end if 'if trim(sIMType)="I" then
											                %>
												            </table>
												        </div>
									                </td>
									            </tr>
									        </table>
								        </td>
								        <td align="center"></td>
                                    </tr>
                            		<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
									<tr>
										<td align="right" colspan=3>
											<p align="right">
							                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
							                        <Input Type=Hidden name="hCtr" Value="<%=iCnt%>" >
							                        <Input Type=Hidden name="hPageSelection" Value="" >
													<%	If iTotalPages >= 2 Then
															if iCurrentPage = 1 then
													%>
													<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
													<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
													<%		else	%>
													<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
													<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
													<%		end if	%>
													<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
													<%
														For lnPage = 1 To iTotalPages
															If cdbl(lnPage) = cdbl(iCurrentPage) Then
													%>
														<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPages%></OPTION>
													<%		else	%>
														<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
													<%		end if
														next
													%>
													</SELECT>
													<%
															if iCurrentPage = iTotalPages then
													%>
													<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
													<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

													<%		else	%>
													<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
													<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPages%>')" id=button8 name=button8>
													<%		end if
														End If
													%>

											</td>
									</tr>
									<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
									<!--<tr>
										<td align="center" width="5" class="ClearPixel">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
										<td valign="top">
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
											 <tr>
											    <td>
                                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
									                <tr>
									                <td valign="middle" class="ActionCell">
									                <p align="center">
										                <input type="button" value="Update" name="btnUpdate" class="ActionButton" onClick="">
									                </td>
									                </tr>
									                </table>
									            </td>
                                             </tr>
                                             <tr>
									            <td align="center" class="MiddlePack"></td>
								            </tr>
										</table>
									</td>
								</tr>-->
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
