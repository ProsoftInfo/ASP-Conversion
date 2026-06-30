<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TransferClosingEntry.asp
	'Module Name				:	Transfer Closing Values
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 12, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	TransferClosingDetailsEntry.asp
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
<!--#include file="../../include/populate.asp"-->
<%
    Dim sQuery,rsObj,rsTemp,rsObjUnit,rsObjItmType,dcrs,rsObjClassCat,rsObjClass,rsObjClassParent,rsTemp1
    Dim sActive,sClosed,sTransfer,sTransOn,sTransBy,sCloseOn,sCloseBy
    Dim iRowCtr,iUnitCnt,sItmCount,sNSClosed
    Dim dCurFinSDt,dCurFinEDt,iCurPeriodFrom,iCurPeriodTo,sUnitCode,dArrSDt,dArrEDt
    Dim iConsYearEndCloseClassCate,iConsYearEndCloseClass,iConsYearEndCloseClassParent
    Dim sGroupNameParent,sGroupCodeParent,sCategoryCode,sCategoryName,sGroupName,sGroupCode

    Set dcrs = Server.CreateObject("ADODB.Recordset")
    Set rsObj = Server.CreateObject("ADODB.Recordset")
    Set rsTemp = Server.CreateObject("ADODB.Recordset")
    Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
    Set rsObjUnit = Server.CreateObject("ADODB.Recordset")
    Set rsObjItmType = Server.CreateObject("ADODB.Recordset")
    Set rsObjClassCat = Server.CreateObject("ADODB.Recordset")
    Set rsObjClass = Server.CreateObject("ADODB.Recordset")
    Set rsObjClassParent =  Server.CreateObject("ADODB.Recordset")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Transfer Closing</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/AdminTransferClosingCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
ITMSAdminTransferClosingCompat.installCloseEntry();
</SCRIPT>
</HEAD>
<%
	dim sWho,sFor
	'sWho = Request.QueryString("sWho")
	sFor = Request.QueryString("Frm")
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name="selPFinStartDate" value="">
<input type=hidden name="hItemType" value="">
<input type=hidden name="hUnitCode" value="">
<input type=hidden name="hOrgName" value="">
<input type=hidden name="hCFinStartDate" value="">
<input type=hidden name="hCFinEndDate" value="">
<input type=hidden name="hWho" value="<%=sWho%>">
<input type=hidden name="hFor" value="<%=sFor%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20">
			<p align="center">
			<%
			    if trim(sFor)="GL" then
			        Response.Write "Accounts Closing"
			    elseif trim(sFor)="AG" then
			        Response.Write "Audit Closing [Accounts]"
			    elseif trim(sFor)="IS" then
			        Response.Write "Stock Closing"
			    elseif trim(sFor)="NS" then
			        Response.Write "Number Series Transfer"
			    else
			        Response.Write "Transfer Closing"
			    end if
			%>

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">

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
								<td valign="top" width="100%" align="left">
                                    <table BORDER="0" CELLSPACING="1" CELLPADDING="0">
                                    <%if trim(sFor)<>"IS" and trim(sFor)<>"AG" and trim(sFor)<>"GL" and trim(sFor)<>"NS" then %>
                                        <tr>
											<td class="FieldCell" valign="top">For Unit</td>
											<td class="FieldCellSub" valign="top">
												<select size="1" name="selUnit" class="FormElem">
													<option value="select">Select</option>

														<%	'Calling the Function which populates the Organization Units list
														  IF CStr(sFor) = "GL" or CStr(sFor) = "PA" Then %>
															<option value="0" Selected>All Units</option>
														<%Else
															populateUnit
														  End IF
														%>
												</select>
											</td>
                                        </tr>
                                        <%if trim(sFor)<>"GL" and trim(sFor)<>"AG" and trim(sFor)<>"NS" then%>
                                        <tr>
                                        	<td class="FieldCell" valign="top">
                                        		For Item Type
                                        	</td>
                                        	<td class="FieldCellSub" valign="top">
                                        		<select size="1" name="selItemType" class="FormElem">
													<option value="select">Select</option>
													<% populateItemType %>
												</select>
                                        	</td>
                                        </tr>
                                        <%end if 'if trim(sFor)<>"GL" then %>

                                   <%end if 'if trim(sFor)<>"IS" then %>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<%if trim(sFor)="GL" then%>
							    <tr>
							        <td></td>
							        <td>
							            <table cellpadding=0 cellspacing=1 class="ExcelTable" width=90%>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center ></td>
							                    <td class="ExcelHeaderCell" align=left colspan=10>Unit</td>

							                </tr>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center rowspan=2 width=30>S.No.</td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2 width=30></td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2>Financial Period</td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2>Active</td>
							                    <td class="ExcelHeaderCell" align=center colspan=3>Transferred</td>
							                    <td class="ExcelHeaderCell" align=center colspan=3>Last Closed</td>
							                </tr>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center>Status</td>
							                    <td class="ExcelHeaderCell" align=center>By</td>
							                    <td class="ExcelHeaderCell" align=center>On</td>
							                    <td class="ExcelHeaderCell" align=center>Closed</td>
							                    <td class="ExcelHeaderCell" align=center>By</td>
							                    <td class="ExcelHeaderCell" align=center>On</td>
							                </tr>

							                <tr>
							                    <td class="ExcelHeaderCell" align=center></td>
							                    <td class="ExcelDisplayCell" align=left colspan=10><b>All Units</B></td>
							                </tr>

							                <%
							                    sQuery = "Select Convert(varchar,FromPeriod,103),Convert(varchar,ToPeriod,103),isNull(Active,'N'),isNull(Closed,'N'),isNull(Transferred,'N'),TransferredBy,Convert(varchar,TransferredOn,103),FinYearClosedBy,Convert(varchar,FinYearClosedOn,103) from MS_FinancialPeriod order by FromPeriod desc"
							                    rsObj.Open sQuery,con
							                    if not rsObj.EOF then
							                        do while not rsObj.EOF
							                            iRowCtr = iRowCtr + 1
							                            sActive = rsObj(2)
							                            sClosed = rsObj(3)
							                            sTransfer = rsObj(4)
							                            sTransBy = rsObj(5)
							                            sTransOn = rsObj(6)
							                            sCloseBy = rsObj(7)
							                            sCloseOn = rsObj(8)
							                            %>
							                                <tr>
							                                    <td class="ExcelHeaderCell" align=center><%=iRowCtr%></td>
							                                    <td class="ExcelDisplayCell" align=center>

							                                    <%
							                                        if trim(sFor)="GL" then
							                                            if trim(sActive)="N" and trim(sTransfer)="N" then %>
							                                                <input type=radio name="radFinPeriod" value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=0%>#<%="All Units"%>" >
							                                            <%else %>
    							                                            <input type=radio name="radFinPeriod" value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=0%>#<%="All Units"%>" disabled>
	    						                                        <%end if 'if trim(rsObj(2))="N" and trim(rsObj(3))="N" then
	    						                                    end if
	    						                                %>
							                                    </td>
							                                    <td class="ExcelDisplayCell" align=center><%=rsObj(0)%> - <%=rsObj(1)%></td>
							                                    <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if trim(sActive)="Y" then
							                                                Response.Write "Yes"
							                                            else
							                                                Response.Write "No"
							                                            end if
							                                        %>
							                                    </td>
							                                     <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if trim(sTransfer)="Y" then
							                                                Response.Write "Yes"
							                                            else
							                                                Response.Write "No"
							                                            end if
							                                        %>
							                                    </td>

							                                     <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if Trim(sTransfer)="Y" then
							                                                sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sTransBy
							                                                rsTemp.Open sQuery,con
							                                                if not rsTemp.EOF then
							                                                    Response.Write trim(rsTemp(0))
							                                                end if
							                                                rsTemp.Close
							                                            else
							                                                Response.Write "-"
							                                            end if

							                                        %>
							                                    </td>
							                                     <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if Trim(sTransfer)="Y" then
							                                                Response.Write sTransOn
							                                            else
							                                                Response.Write "-"
							                                            end if
							                                        %>
							                                    </td>
							                                      <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            sQuery = "Select FinYearAuditClosedBy,Convert(varchar,FinYearAuditClosedOn,103) from Ms_AuditorClosing where FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103) Order by FinYearAuditClosedOn desc"
						                                                        'Response.Write sQuery
				                                                        rsTemp.Open sQuery,con
				                                                        if not rsTemp.EOF then
				                                                            sClosed = "Y"
				                                                            sCloseBy = rsTemp(0)
				                                                            sCloseOn = rsTemp(1)
				                                                        else
				                                                            sClosed = "N"
				                                                        end if
				                                                        rsTemp.Close


							                                            if trim(sClosed)="Y" then
							                                                Response.Write "Yes"
							                                            else
							                                                Response.Write "No"
							                                            end if
							                                        %>
							                                    </td>
							                                     <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if trim(sClosed)="Y" then
							                                                sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sCloseBy
							                                                rsTemp.Open sQuery,con
							                                                if not rsTemp.EOF then
							                                                    Response.Write trim(rsTemp(0))
							                                                end if
							                                                rsTemp.Close
							                                            else
							                                                Response.Write "-"
							                                            end if
							                                        %>
							                                    </td>
							                                    <td class="ExcelDisplayCell" align=center>
							                                        <%
							                                            if trim(sClosed)="Y" then
							                                                Response.Write sCloseOn
							                                            else
							                                                Response.Write "-"
							                                            end if
							                                        %>
							                                    </td>
							                                </tr>
							                            <%
							                            rsObj.MoveNext
							                        loop
							                    end if
							                    rsObj.Close
							                %>
							            </table>
							            <input type=hidden name=hRowCtr value="<%=iRowCtr%>">
							        </td>
							        <td></td>
							    </tr>
							<%elseif trim(sFor)="AG" then%>
							    <tr>
							        <td></td>
							        <td>
							            <table cellpadding=0 cellspacing=1 class="ExcelTable" width=90%>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center ></td>
							                    <td class="ExcelHeaderCell" align=left colspan=10>Unit</td>

							                </tr>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center rowspan=2 width=30>S.No.</td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2 width=30></td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2>Financial Period</td>
							                    <td class="ExcelHeaderCell" align=center rowspan=2>Active</td>
							                    <td class="ExcelHeaderCell" align=center colspan=3>Transferred</td>
							                    <td class="ExcelHeaderCell" align=center colspan=3>Last Closed</td>
							                </tr>
							                <tr>
							                    <td class="ExcelHeaderCell" align=center>Status</td>
							                    <td class="ExcelHeaderCell" align=center>By</td>
							                    <td class="ExcelHeaderCell" align=center>On</td>
							                    <td class="ExcelHeaderCell" align=center>Closed</td>
							                    <td class="ExcelHeaderCell" align=center>By</td>
							                    <td class="ExcelHeaderCell" align=center>On</td>
							                </tr>
							                <%
							                    sQuery = "Select OrgUnitDescription,OUDefinitionID from DCS_OrganizationUnitDefinitions where len(OUDefinitionID)>4"
							                        rsObjUnit.Open sQuery,con
							                        if not rsObjUnit.EOF then
							                            do while not rsObjUnit.EOF
							                                %>
							                                    <tr>
							                                        <td class="ExcelSerial" align=left ></td>
							                                        <td class="ExcelDisplayCell" align=left colspan=10><b><%=trim(rsObjUnit(0))%></b></td>
							                                    </tr>
							                                <%

							                                sQuery = "Select Convert(varchar,FromPeriod,103),Convert(varchar,ToPeriod,103),isNull(Active,'N'),isNull(Closed,'N'),isNull(Transferred,'N'),TransferredBy,Convert(varchar,TransferredOn,103),FinYearClosedBy,Convert(varchar,FinYearClosedOn,103) from MS_FinancialPeriod order by FromPeriod desc"
							                                rsObj.Open sQuery,con
							                                if not rsObj.EOF then
							                                    do while not rsObj.EOF
							                                        iRowCtr = iRowCtr + 1
							                                        sActive = rsObj(2)
							                                        sClosed = rsObj(3)
							                                        sTransfer = rsObj(4)
							                                        sTransBy = rsObj(5)
							                                        sTransOn = rsObj(6)
							                                        sCloseBy = rsObj(7)
							                                        sCloseOn = rsObj(8)
							                                        %>
							                                            <tr>
							                                                <td class="ExcelHeaderCell" align=center><%=iRowCtr%></td>
							                                                <td class="ExcelDisplayCell" align=center>

							                                                <%
							                                                        if trim(sActive)="N" and trim(sTransfer)="Y" then %>
							                                                            <input type=radio name="radFinPeriod" value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=rsobjUnit(1)%>#<%=rsObjUnit(0)%>" >
							                                                        <%else %>
    							                                                        <input type=radio name="radFinPeriod" value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=rsobjUnit(1)%>#<%=rsObjUnit(0)%>" disabled>
	    						                                                    <%end if 'if trim(rsObj(2))="N" and trim(rsObj(3))="N" then
	    						                                            %>
							                                                </td>
							                                                <td class="ExcelDisplayCell" align=center><%=rsObj(0)%> - <%=rsObj(1)%></td>
							                                                <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if trim(sActive)="Y" then
							                                                            Response.Write "Yes"
							                                                        else
							                                                            Response.Write "No"
							                                                        end if
							                                                    %>
							                                                </td>
							                                                 <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if trim(sTransfer)="Y" then
							                                                            Response.Write "Yes"
							                                                        else
							                                                            Response.Write "No"
							                                                        end if
							                                                    %>
							                                                </td>
							                                                 <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if Trim(sTransfer)="Y" then
							                                                            sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sTransBy
							                                                            rsTemp.Open sQuery,con
							                                                            if not rsTemp.EOF then
							                                                                Response.Write trim(rsTemp(0))
							                                                            end if
							                                                            rsTemp.Close
							                                                        else
							                                                            Response.Write "-"
							                                                        end if

							                                                    %>
							                                                </td>
							                                                 <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if Trim(sTransfer)="Y" then
							                                                            Response.Write sTransOn
							                                                        else
							                                                            Response.Write "-"
							                                                        end if
							                                                    %>
							                                                </td>
							                                                 <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        sQuery = "Select FinYearAuditClosedBy,Convert(varchar,FinYearAuditClosedOn,103) from Ms_AuditorClosing where OUDefinitionID = "& rsObjUnit(1) &" and FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103) Order by FinYearAuditClosedOn desc"
							                                                        'Response.Write sQuery
							                                                        rsTemp.Open sQuery,con
							                                                        if not rsTemp.EOF then
							                                                            sClosed = "Y"
							                                                            sCloseBy = rsTemp(0)
							                                                            sCloseOn = rsTemp(1)
							                                                        else
							                                                            sClosed = "N"
							                                                        end if
							                                                        rsTemp.Close

							                                                        if trim(sClosed)="Y" then
							                                                            Response.Write "Yes"
							                                                        else
							                                                            Response.Write "No"
							                                                        end if
							                                                    %>
							                                                </td>
							                                                 <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if trim(sClosed)="Y" then
							                                                            sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sCloseBy
							                                                            rsTemp.Open sQuery,con
							                                                            if not rsTemp.EOF then
							                                                                Response.Write trim(rsTemp(0))
							                                                            end if
							                                                            rsTemp.Close
							                                                        else
							                                                            Response.Write "-"
							                                                        end if
							                                                    %>
							                                                </td>
							                                                <td class="ExcelDisplayCell" align=center>
							                                                    <%
							                                                        if trim(sClosed)="Y" then
							                                                            Response.Write sCloseOn
							                                                        else
							                                                            Response.Write "-"
							                                                        end if
							                                                    %>
							                                                </td>
							                                            </tr>
							                                        <%
							                                        rsObj.MoveNext
							                                    loop
							                                end if
							                                rsObj.Close
							                                rsObjUnit.MoveNext
							                            loop
							                          end if '
							                          rsObjUnit.Close
							                %>
							            </table>
							            <input type=hidden name=hRowCtr value="<%=iRowCtr%>">
							        </td>
							        <td></td>
							    </tr>
							 <%elseif trim(sFor)="NS" then %>
							        <tr>
							            <td></td>
							            <td>
							                <table cellpadding=0 cellspacing=1 class="ExcelTable" width=90%>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=left width=10></td>
							                        <td class="ExcelHeaderCell" align=left colspan=6>Unit</td>
							                    </tr>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=center rowspan=2>S.No.</td>
							                        <td class="ExcelHeaderCell" align=center colspan=2>Financial Period</td>
							                        <td class="ExcelHeaderCell" align=center rowspan=2>Actvie</td>
							                        <td class="ExcelHeaderCell" align=center colspan=3>Transferred</td>
							                    </tr>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=center></td>
							                        <td class="ExcelHeaderCell" align=center>Item Type</td>
							                        <td class="ExcelHeaderCell" align=center>Status</td>
							                        <td class="ExcelHeaderCell" align=center>By</td>
							                        <td class="ExcelHeaderCell" align=center>On</td>
							                    </tr>
							                    <%
							                        sQuery = "Select OrgUnitDescription,OUDefinitionID from DCS_OrganizationUnitDefinitions where len(OUDefinitionID)>4"
							                        rsObjUnit.Open sQuery,con
							                        if not rsObjUnit.EOF then
							                            do while not rsObjUnit.EOF
							                                %>
							                                    <tr>
							                                        <td class="ExcelSerial" align=left ></td>
							                                        <td class="ExcelDisplayCell" align=left colspan=6><b><%=trim(rsObjUnit(0))%></b></td>
							                                    </tr>
							                                <%
							                                        sQuery = "Select Convert(varchar,FromPeriod,103),Convert(varchar,ToPeriod,103),isNull(Active,'N') from MS_FinancialPeriod order by FromPeriod desc"
							                                        'Response.write sQuery
							                                        rsObj.Open sQuery,con
							                                        if not rsObj.EOF then
							                                            do while not rsObj.EOF
							                                                sActive = rsObj(2)
							                                                iUnitCnt = iUnitCnt  + 1
							                                                dArrSDt = split(rsObj(0),"/")
							                                                dArrEDt = split(rsObj(1),"/")
							                                                dCurFinSDt = dArrSDt(0)&"/"&dArrSDt(1)&"/"& dArrSDt(2) + 1
							                                                dCurFinEDt = dArrEDt(0) &"/"& dArrEDt(1)&"/"& dArrEDt(2) + 1
																			'Response.write dCurFinSDt & "--" & dCurFinEDt
							                                                %>
							                                                    <tr>
							                                                        <td class="ExcelSerial" align=center><%=iUnitCnt%></td>
							                                                        <td class="ExcelDisplayCell" align=center colspan=2><%=rsObj(0)%> - <%=rsObj(1)%></td>
        							                                                <td class="ExcelDisplayCell" align=center>
							                                                            <%
							                                                                if trim(sActive)="Y" then
							                                                                    Response.Write "Yes"
							                                                                else
							                                                                    Response.Write "No"
							                                                                end if
							                                                            %>
							                                                        </td>
							                                                        <td class="ExcelDisplayCell" align=center colspan=3></td>
							                                                    </tr>
							                                                <%
							                                                    if trim(sActive)="N" then
							                                                        sNSClosed = CheckNoSeriesTransfer(1,rsObjUnit(1),dCurFinSDt,dCurFinEDt)
							                                                        'Response.Write " sNSClosed ="& sNSClosed
							                                                        %>
							                                                            <tr>
							                                                                <td class="ExcelSerial" align=center></td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                    <%if sNSClosed>1 then %>
							                                                                        <input type=radio name=radModule value="AC#1#1#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>" disabled>
							                                                                    <%else %>
							                                                                        <input type=radio name=radModule value="AC#1#1#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>">
							                                                                    <%end if  %>
							                                                                </td>
        							                                                        <td class="ExcelDisplayCell" align=left>
        							                                                            ACCOUNTS
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed)>1 then
							                                                                        Response.Write "Yes"
							                                                                    else
							                                                                        Response.Write "No"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                            </tr>
							                                                            <%sNSClosed = CheckNoSeriesTransfer(4,rsObjUnit(1),dCurFinSDt,dCurFinEDt)%>
							                                                             <tr>
							                                                                <td class="ExcelSerial" align=center></td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                    <%if sNSClosed>1 then %>
							                                                                        <input type=radio name=radModule value="IN#1#4#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>" disabled>
							                                                                    <%else %>
							                                                                        <input type=radio name=radModule value="IN#1#4#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>">
							                                                                    <%end if  %>
							                                                                </td>
        							                                                        <td class="ExcelDisplayCell" align=left>
        							                                                            STORES (ALL ITEM TYPES)
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed)>1 then
							                                                                        Response.Write "Yes"
							                                                                    else
							                                                                        Response.Write "No"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                            </tr>
							                                                            <%sNSClosed = CheckNoSeriesTransfer(2,rsObjUnit(1),dCurFinSDt,dCurFinEDt)%>
							                                                             <tr>
							                                                                <td class="ExcelSerial" align=center></td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                    <%if sNSClosed>1 then %>
							                                                                        <input type=radio name=radModule value="PU#1#2#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>" disabled>
							                                                                    <%else %>
							                                                                        <input type=radio name=radModule value="PU#1#2#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>">
							                                                                    <%end if  %>
							                                                                </td>
        							                                                        <td class="ExcelDisplayCell" align=left>
        							                                                            PURCHASE (ALL ITEM TYPES)
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed)>1 then
							                                                                        Response.Write "Yes"
							                                                                    else
							                                                                        Response.Write "No"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                            </tr>
							                                                            <%sNSClosed = CheckNoSeriesTransfer(3,rsObjUnit(1),dCurFinSDt,dCurFinEDt)%>
							                                                             <tr>
							                                                                <td class="ExcelSerial" align=center></td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                    <%if sNSClosed>1 then %>
							                                                                        <input type=radio name=radModule value="SA#1#3#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>" disabled>
							                                                                    <%else %>
							                                                                        <input type=radio name=radModule value="SA#1#3#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>">
							                                                                    <%end if  %>
							                                                                </td>
        							                                                        <td class="ExcelDisplayCell" align=left>
        							                                                            SALES (ALL ITEM TYPES)
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed)>1 then
							                                                                        Response.Write "Yes"
							                                                                    else
							                                                                        Response.Write "No"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                            </tr>
							                                                            <%sNSClosed = CheckNoSeriesTransfer(6,rsObjUnit(1),dCurFinSDt,dCurFinEDt)%>
							                                                             <tr>
							                                                                <td class="ExcelSerial" align=center></td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                    <%if sNSClosed>1 then %>
							                                                                        <input type=radio name=radModule value="PD#1#6#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>" disabled>
							                                                                    <%else %>
							                                                                        <input type=radio name=radModule value="PD#1#6#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>#<%=rsobj(0)%>#<%=rsobj(1)%>#<%=dCurFinSDt%>#<%=dCurFinEDt%>">
							                                                                    <%end if  %>
							                                                                </td>
        							                                                        <td class="ExcelDisplayCell" align=left>
        							                                                           PRODUCTION (ALL ITEM TYPES)
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed)>1 then
							                                                                        Response.Write "Yes"
							                                                                    else
							                                                                        Response.Write "No"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                                <td class="ExcelDisplayCell" align=center>
							                                                                <%
							                                                                    if trim(sNSClosed) > 1 then
							                                                                    else
							                                                                        Response.Write "-"
							                                                                    end if
							                                                                %>
							                                                                </td>
							                                                            </tr>

							                                                        <%
							                                                        iRowCtr = iRowCtr + 5
							                                                    end if  'if trim(sActive)="N" then
							                                                rsObj.MoveNext
							                                            loop
							                                        end if
							                                        rsObj.Close

							                                rsObjUnit.MoveNext
							                            loop
							                        end if
							                        rsObjUnit.Close


							                    %>
							                     </table>
							                <input type=hidden name=hRowCtr value="<%=iRowCtr%>">
							            </td>
							            <td></td>
							        </tr>

							<%else %>
							        <tr>
							            <td></td>
							            <td>
							                <table cellpadding=0 cellspacing=1 class="ExcelTable" width=90%>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=left width=10></td>
							                        <td class="ExcelHeaderCell" align=left colspan=6>Unit</td>
							                    </tr>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=center rowspan=2>S.No.</td>
							                        <td class="ExcelHeaderCell" align=center colspan=2>Financial Period</td>
							                        <td class="ExcelHeaderCell" align=center rowspan=2>Actvie</td>
							                        <td class="ExcelHeaderCell" align=center colspan=3>Transferred</td>
							                    </tr>
							                    <tr>
							                        <td class="ExcelHeaderCell" align=center></td>
							                        <td class="ExcelHeaderCell" align=center>Item Type</td>
							                        <td class="ExcelHeaderCell" align=center>Status</td>
							                        <td class="ExcelHeaderCell" align=center>By</td>
							                        <td class="ExcelHeaderCell" align=center>On</td>
							                    </tr>
							                    <%
							                        sQuery = "Select OrgUnitDescription,OUDefinitionID from DCS_OrganizationUnitDefinitions where len(OUDefinitionID)>4"
							                        rsObjUnit.Open sQuery,con
							                        if not rsObjUnit.EOF then
							                            do while not rsObjUnit.EOF
							                                %>
							                                    <tr>
							                                        <td class="ExcelSerial" align=left ></td>
							                                        <td class="ExcelDisplayCell" align=left colspan=6><b><%=trim(rsObjUnit(0))%></b></td>
							                                    </tr>
							                                <%
							                                        sQuery = "Select Convert(varchar,FromPeriod,103),Convert(varchar,ToPeriod,103),isNull(Active,'N') from MS_FinancialPeriod order by FromPeriod desc"
							                                        rsObj.Open sQuery,con
							                                        if not rsObj.EOF then
							                                            do while not rsObj.EOF
							                                                sActive = rsObj(2)
							                                                iUnitCnt = iUnitCnt  + 1
							                                                %>
							                                                    <tr>
							                                                        <td class="ExcelSerial" align=center><%=iUnitCnt%></td>
							                                                        <td class="ExcelDisplayCell" align=center colspan=2><%=rsObj(0)%> - <%=rsObj(1)%></td>
        							                                                <td class="ExcelDisplayCell" align=center>
							                                                            <%
							                                                                if trim(sActive)="Y" then
							                                                                    Response.Write "Yes"
							                                                                else
							                                                                    Response.Write "No"
							                                                                end if
							                                                            %>
							                                                        </td>
							                                                        <td class="ExcelDisplayCell" align=center colspan=3></td>
							                                                    </tr>
							                                                    <%
							                                                    if trim(sActive)="N" then
							                                                    Response.Write "<font color=red>"

							                                                        sQuery = "Select ConsiderForYearEndClosing,CategoryName,CategoryCode from Inv_M_ClassificationCategory"
							                                                        rsObjClassCat.Open sQuery,con
							                                                        if not rsObjClassCat.EOF then
							                                                            do while not rsObjClassCat.EOF
							                                                                iConsYearEndCloseClassCate = rsObjClassCat(0)
							                                                                sCategoryName = rsObjClassCat(1)
							                                                                sCategoryCode = rsObjClassCat(2)
							                                                                sTransfer = ""
							                                                                sTransBy = ""
							                                                                sTransOn =""
							                                                                if iConsYearEndCloseClassCate=1 then
							                                                                iRowCtr = iRowCtr + 1


							                                                                    ' sQuery = "Select Count(*) from VWItemClassCatForYearEndClosing C,Inv_T_ItemYearlyStock IY where"&_
							                                                                     '                " C.GroupCode = IY.ClassificationCode and "&_
							                                                                      '               " IY.FinancialYearFrom >= Convert(datetime,'"&rsObj(0)&"',103) and IY.FinancialYearTo <="&_
							                                                                       '              " Convert(datetime,'"&rsObj(1)&"',103) and C.ItemCode=IY.ItemCode and C.GroupCategory = '"& sCategoryCode  &"'"


							                                                                        sQuery = " Select Count(*) from Inv_T_ItemYearlyStock IY where IY.FinancialYearFrom >= Convert(datetime,'"& rsObj(0) &"',103) "&_
							                                                                                 " and IY.FinancialYearTo <= Convert(datetime,'"& rsObj(1) &"',103) and IY.ClassificationCode in (Select ClassificationCode "&_
							                                                                                 " from VWItemClassCatForYearEndClosing where ParentGroup in (Select GroupCode from VWItemClassCatForYearEndClosing where GroupCategory = '"& sCategoryCode &"'))"&_
							                                                                                 " and IY.ItemCode in (Select ItemCode from VWItemClassCatForYearEndClosing where ParentGroup in (Select GroupCode from VWItemClassCatForYearEndClosing where GroupCategory = '"& sCategoryCode &"'))"
							                                                                                 'Response.Write "<textarea> "& sQuery &"</textarea>"
							                                                                                     rsTemp.Open sQuery,con
							                                                                                     if not rsTemp.EOF then
							                                                                                            sItmCount = rsTemp(0)
							                                                                                     end if
							                                                                                     rsTemp.Close

							                                                                                     sQuery = "Select Transferred,TransferredBy,Convert(varchar,TransferredOn,103) from MS_StockClosing where CategoryCode = '"& sCategoryCode &"' "&_
        							                                                                                      " and FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103)"
        							                                                                             rsTemp.Open sQuery,con
							                                                                                     if not rsTemp.EOF then
							                                                                                        sTransfer =  rsTemp(0)
							                                                                                        sTransBy = rsTemp(1)
							                                                                                        sTransOn = rsTemp(2)
							                                                                                     end if
							                                                                                     rsTemp.Close
							                                                                %>

							                                                                    <tr>
	                                                                                                <td class="ExcelSerial" align=center></td>
	                                                                                                <td class="ExcelDisplayCell" align=center>
	                                                                                                    <%if sItmCount = 0 OR sTransfer="Y" then%>
	                                                                                                    <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>::#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>" disabled>
	                                                                                                    <%else %>
	                                                                                                    <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>::#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>">
	                                                                                                    <%end if%>
	                                                                                                </td>
	                                                                                                <td class="ExcelDisplayCell" align=Left><%=sCategoryName%></td>

	                                                                                                <%if  sItmCount = 0 then%>
						                                                                                <td class="ExcelDisplayCell" align=Center colspan=4>No Items to Transfer</td>
						                                                                            <%else%>
						                                                                                <td class="ExcelDisplayCell" align=Center></td>
						                                                                                <td class="ExcelDisplayCell" align=Center>
						                                                                                <%
						                                                                                    if trim(sTransfer)="Y" then
						                                                                                        Response.Write "Yes"
						                                                                                    else
						                                                                                        Response.Write "No"
						                                                                                    end if
						                                                                                %>
						                                                                                </td>
						                                                                                <td class="ExcelDisplayCell" align=center>
                                                                                                            <%
                                                                                                                if Trim(sTransfer)="Y" then
                                                                                                                    sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sTransBy
                                                                                                                    rsTemp.Open sQuery,con
                                                                                                                    if not rsTemp.EOF then
                                                                                                                        Response.Write trim(rsTemp(0))
                                                                                                                    end if
                                                                                                                    rsTemp.Close
                                                                                                                else
                                                                                                                    Response.Write "-"
                                                                                                                end if
                                                                                                            %>
                                                                                                        </td>
                                                                                                        <td class="ExcelDisplayCell" align=center>
                                                                                                        <%
                                                                                                            if trim(sTransfer)="Y" then
                                                                                                                Response.Write trim(sTransOn)
                                                                                                            else
                                                                                                                Response.Write "-"
                                                                                                            end if
                                                                                                        %>
                                                                                                        </td>
						                                                                            <%end if 'if  sItmCount = 0 then%>


	                                                                                            </tr>
				                                                                    <%
						                                                                    else
						                                                                        sQuery = "Select ConsiderForYearEndClosing,GroupName,GroupCode from Inv_M_Classification where ParentGroup in ("&_
						                                                                                 " Select GroupCode from Inv_M_Classification where GroupCategory = '"& sCategoryCode &"' ) and GroupCode = ParentGroup "
																								'Response.Write sQuery & "<br>"
						                                                                        rsObjClassParent.Open sQuery,con
						                                                                        if not rsObjClassParent.EOF then
						                                                                            do while not rsObjClassParent.EOF
						                                                                                iConsYearEndCloseClassParent=rsObjClassParent(0)
						                                                                                sGroupNameParent =  rsObjClassParent(1)
						                                                                                sGroupCodeParent =  rsObjClassParent(2)
						                                                                                sTransfer = ""
							                                                                            sTransBy = ""
							                                                                            sTransOn =""
						                                                                                if iConsYearEndCloseClassParent = 1 then
						                                                                                    iRowCtr = iRowCtr + 1


						                                                                                        sQuery = "Select Count(*) from VWItemClassCatForYearEndClosing C,Inv_T_ItemYearlyStock IY where"&_
							                                                                                     " C.GroupCode = IY.ClassificationCode and "&_
							                                                                                     " IY.FinancialYearFrom >= Convert(datetime,'"&rsObj(0)&"',103) and IY.FinancialYearTo <="&_
							                                                                                     " Convert(datetime,'"&rsObj(1)&"',103) and C.ItemCode=IY.ItemCode and C.ParentGroup = "& sGroupCodeParent  &" "
							                                                                                     'Response.Write "<textarea> "& sQuery &"</textarea>"
							                                                                                     rsTemp.Open sQuery,con
							                                                                                     if not rsTemp.EOF then
							                                                                                            sItmCount = rsTemp(0)
							                                                                                     end if
							                                                                                     rsTemp.Close

							                                                                                     sQuery = "Select Transferred,TransferredBy,Convert(varchar,TransferredOn,103) from MS_StockClosing where CategoryCode = '"& sCategoryCode &"' and SubCategory = " & sGroupCodeParent & " "&_
        							                                                                                      " and FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103)"
        							                                                                             rsTemp.Open sQuery,con
							                                                                                     if not rsTemp.EOF then
							                                                                                        sTransfer =  rsTemp(0)
							                                                                                        sTransBy = rsTemp(1)
							                                                                                        sTransOn = rsTemp(2)
							                                                                                     end if
							                                                                                     rsTemp.Close


			                                                                                                %>
			                                                                                                    <tr>
                                                                                                                    <td class="ExcelSerial" align=center></td>
                                                                                                                    <td class="ExcelDisplayCell" align=center>
                                                                                                                        <%if sItmCount = 0 OR sTransfer="Y" then%>
	                                                                                                                    <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>:<%=sGroupCodeParent%>:#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>" disabled>
	                                                                                                                    <%else %>
	                                                                                                                    <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>:<%=sGroupCodeParent%>:#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>">
	                                                                                                                    <%end if%>
                                                                                                                    </td>
                                                                                                                    <td class="ExcelDisplayCell" align=Left><%=sGroupNameParent%></td>

                                                                                                                    <%if  sItmCount = 0 then%>
			                                                                                                            <td class="ExcelDisplayCell" align=Center colspan=4>No Items to Transfer</td>
			                                                                                                        <%else%>
			                                                                                                            <td class="ExcelDisplayCell" align=Center></td>
			                                                                                                            <td class="ExcelDisplayCell" align=Center>
			                                                                                                            <%
			                                                                                                                if trim(sTransfer)="Y" then
			                                                                                                                    Response.Write "Yes"
			                                                                                                                else
			                                                                                                                    Response.Write "No"
			                                                                                                                end if
			                                                                                                            %>
			                                                                                                            </td>
			                                                                                                            <td class="ExcelDisplayCell" align=center>
                                                                                                                            <%
                                                                                                                                if Trim(sTransfer)="Y" then
                                                                                                                                    sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sTransBy
                                                                                                                                    rsTemp.Open sQuery,con
                                                                                                                                    if not rsTemp.EOF then
                                                                                                                                        Response.Write trim(rsTemp(0))
                                                                                                                                    end if
                                                                                                                                    rsTemp.Close
                                                                                                                                else
                                                                                                                                    Response.Write "-"
                                                                                                                                end if
                                                                                                                            %>
                                                                                                                        </td>
                                                                                                                        <td class="ExcelDisplayCell" align=center>
                                                                                                                        <%
                                                                                                                            if trim(sTransfer)="Y" then
                                                                                                                                Response.Write trim(sTransOn)
                                                                                                                            else
                                                                                                                                Response.Write "-"
                                                                                                                            end if
                                                                                                                        %>
                                                                                                                        </td>
			                                                                                                        <%end if 'if  sItmCount = 0 then%>
                						                                                                        </tr>
			                                                                                                <%
						                                                                                else

						                                                                                    sQuery = "Select ConsiderForYearEndClosing,GroupName,GroupCode from Inv_M_Classification where GroupCode <> ParentGroup and ParentGroup = "& sGroupCodeParent
						                                                                                    rsTemp.Open sQuery,con
						                                                                                    if not rsTemp.EOF then
						                                                                                        do while not rsTemp.EOF
						                                                                                            iConsYearEndCloseClass = rsTemp(0)
						                                                                                            sGroupName = rsTemp(1)
						                                                                                            sGroupCode = rsTemp(2)
						                                                                                            sTransfer = ""
							                                                                                        sTransBy = ""
							                                                                                        sTransOn =""
						                                                                                            if iConsYearEndCloseClass = 1 then
						                                                                                                iRowCtr = iRowCtr + 1

						                                                                                                sQuery = "Select Count(*) from VWItemClassCatForYearEndClosing C,Inv_T_ItemYearlyStock IY where"&_
							                                                                                             " C.GroupCode = IY.ClassificationCode and "&_
							                                                                                             " IY.FinancialYearFrom >= Convert(datetime,'"&rsObj(0)&"',103) and IY.FinancialYearTo <="&_
							                                                                                             " Convert(datetime,'"&rsObj(1)&"',103) and C.ItemCode=IY.ItemCode and C.GroupCode = "& sGroupCode  &" "
							                                                                                             'Response.Write "<textarea> "& sQuery &"</textarea>"
							                                                                                             rsTemp1.Open sQuery,con
							                                                                                             if not rsTemp1.EOF then
							                                                                                                    sItmCount = rsTemp1(0)
							                                                                                             end if
							                                                                                             rsTemp1.Close

							                                                                                             sQuery = "Select Transferred,TransferredBy,Convert(varchar,TransferredOn,103) from MS_StockClosing where CategoryCode = '"& sCategoryCode &"' and SubCategory = " & sGroupCodeParent & " and Classification = "& sGroupCode &" "&_
        							                                                                                      " and FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103)"
        							                                                                                     rsTemp1.Open sQuery,con
							                                                                                             if not rsTemp1.EOF then
							                                                                                                sTransfer =  rsTemp1(0)
							                                                                                                sTransBy = rsTemp1(1)
							                                                                                                sTransOn = rsTemp1(2)
							                                                                                             end if
							                                                                                             rsTemp1.Close

						                                                                                            %>
						                                                                                                <tr>
	                                                                                                                        <td class="ExcelSerial" align=center></td>
	                                                                                                                        <td class="ExcelDisplayCell" align=center>
	                                                                                                                        <%if sItmCount = 0 OR sTransfer="Y" then%>
	                                                                                                                        <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>:<%=sGroupCodeParent%>:<%=sGroupCode%>#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>" disabled>
	                                                                                                                        <%else %>
	                                                                                                                        <input type=radio name=radIType value="<%=rsobj(0)%>:<%=rsobj(1)%>#<%=sCategoryCode%>:<%=sGroupCodeParent%>:<%=sGroupCode%>#<%=rsObjUnit(0)%>#<%=rsObjUnit(1)%>">
	                                                                                                                        <%end if%>
	                                                                                                                        </td>

	                                                                                                                        <td class="ExcelDisplayCell" align=Left><%=sGroupName%></td>

	                                                                                                                        <%if  sItmCount = 0 then%>
						                                                                                                        <td class="ExcelDisplayCell" align=Center colspan=4>No Items to Transfer</td>
						                                                                                                    <%else%>
						                                                                                                        <td class="ExcelDisplayCell" align=Center></td>
						                                                                                                        <td class="ExcelDisplayCell" align=Center>
						                                                                                                        <%
						                                                                                                            if trim(sTransfer)="Y" then
						                                                                                                                Response.Write "Yes"
						                                                                                                            else
						                                                                                                                Response.Write "No"
						                                                                                                            end if
						                                                                                                        %>
						                                                                                                        </td>
						                                                                                                        <td class="ExcelDisplayCell" align=center>
                                                                                                                                    <%
                                                                                                                                        if Trim(sTransfer)="Y" then
                                                                                                                                            sQuery = "Select EmployeeName from MS_EmployeeMaster Where EmployeeNumber = "& sTransBy
                                                                                                                                            rsTemp1.Open sQuery,con
                                                                                                                                            if not rsTemp1.EOF then
                                                                                                                                                Response.Write trim(rsTemp1(0))
                                                                                                                                            end if
                                                                                                                                            rsTemp1.Close
                                                                                                                                        else
                                                                                                                                            Response.Write "-"
                                                                                                                                        end if
                                                                                                                                    %>
                                                                                                                                </td>
                                                                                                                                <td class="ExcelDisplayCell" align=center>
                                                                                                                                <%
                                                                                                                                    if trim(sTransfer)="Y" then
                                                                                                                                        Response.Write trim(sTransOn)
                                                                                                                                    else
                                                                                                                                        Response.Write "-"
                                                                                                                                    end if
                                                                                                                                %>
                                                                                                                                </td>
						                                                                                                    <%end if 'if  sItmCount = 0 then%>
                            						                                                                    </tr>
						                                                                                            <%
						                                                                                            end if 'if iConsYearEndCloseClass = 1 then
						                                                                                            rsTemp.MoveNext
						                                                                                        loop
						                                                                                    end if
						                                                                                    rsTemp.Close

						                                                                                end if' if iConsYearEndCloseClassParent = 1 then
						                                                                                rsObjClassParent.MoveNext
						                                                                            loop
						                                                                        end if
						                                                                        rsObjClassParent.Close
							                                                                end if 'if iConsYearEndCloseClassCate=1 then
							                                                                rsObjClassCat.MoveNext
							                                                            loop
							                                                        end if
							                                                        rsObjClassCat.Close

							                                                      '  sQuery= "Select ItemTypeID,ItemTypeName from Inv_M_ItemType Order by ItemTypeNo"
							                                                       ' rsObjItmType.Open sQuery,con
							                                                       ' if not rsObjItmType.EOF then
							                                                        '    do while not rsObjItmType.EOF
							                                                         '       sTransfer =  ""
                                                                                      '      sTransBy = ""
                                                                                    '        sTransOn = ""
							                                                         '
							                                                          '                  sQuery = "Select Count(*) from Inv_M_ItemType IT,Inv_M_Classification C,Inv_T_ItemYearlyStock IY where"&_
							                                                           '                          " IT.ItemTypeID = C.ItemTypeID and C.GroupCode = IY.ClassificationCode and "&_
							                                                            '                         " IY.FinancialYearFrom >= Convert(datetime,'"&rsObj(0)&"',103) and IY.FinancialYearTo <="&_
							                                                            '                         " Convert(datetime,'"&rsObj(1)&"',103) and IT.ItemTypeID = '"& rsObjItmType(0) &"'"
							                                                            '                         rsTemp.Open sQuery,con
							                                                             '                        if not rsTemp.EOF then
							                                                              '                              sItmCount = rsTemp(0)
							                                                               '                      end if
							                                                                '                     rsTemp.Close
							                                                                 '
						                                                                      '                   sQuery = "Select Transferred,TransferredBy,Convert(varchar,TransferredOn,103) from MS_StockClosing where ItemTypeID = '"& rsObjItmType(0) &"' "&_
        							                                                           '                           " and FromPeriod = Convert(datetime,'"& rsObj(0) &"',103) and ToPeriod = Convert(datetime,'"& rsObj(1) &"',103)"
        							                                                            '                    'Response.Write sQuery
							                                                                     '                rsTemp.Open sQuery,con
							                                                                      '               if not rsTemp.EOF then
							                                                                       '                 sTransfer =  rsTemp(0)
							                                                                        '                sTransBy = rsTemp(1)
							                                                                         '               sTransOn = rsTemp(2)
							                                                                          '           end if
							                                                                           '          rsTemp.Close
							                                                            '
							                                                             '
							                                                              '                      iRowCtr = iRowCtr + 1
							                                                  '              rsObjItmType.MoveNext
							                                                   '         loop
							                                                    '    end if
							                                                     '   rsObjItmType.Close

							                                                    end if  'if trim(sActive)="N" then
							                                                rsObj.MoveNext
							                                            loop
							                                        end if
							                                        rsObj.Close

							                                rsObjUnit.MoveNext
							                            loop
							                        end if
							                        rsObjUnit.Close


							                    %>

							                </table>
							                <input type=hidden name=hRowCtr value="<%=iRowCtr%>">
							            </td>
							            <td></td>
							        </tr>

							    <%end if'if trim(sFor)<>"IS" then%>
							<!--<tr>
								<td align="center"></td>
								<td valign="top">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class='GroupTitleLeft' width="10">&nbsp;</td>
														<td class='GroupTitle' width="147"><p align="center">Transfer Closing Values</td>
														<td class='GroupTitleRight'><p align="left">&nbsp;</td>
													</tr>
												</table>
                                            </td>
										</tr>
										<tr>
											<td class=GroupTable>
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class=MiddlePack colspan="7"> </td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top"> Previous Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<select size="1" name="selPFinStartDate" class="FormElem" onChange="SetDates(this)">
																<option value="select">Select</option>
															<%	'Calling the Function which populates the Previous Financial Year Start Date list
																populateFinDate
															%>
															</select>
                                                        </td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idPFinEndDate">&nbsp;</span>
                                                        </td>
														<td class='FieldCellSub'>
                                                        </td>
													</tr>
													<tr>
														<td class=MiddlePack valign="top" colspan="7">
														</td>
													</tr>
													<tr>
														<td class=FieldCellSub valign="top">Current Financial Year</td>
														<td class="FieldCellSub" valign="top">Start Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idCFinStartDate">&nbsp;</span>
                                                        </td>
														<td class='FieldCellSub'></td>
														<td class='FieldCellSub'>End Date</td>
														<td class='FieldCellSub'>
															<span class="Dataonly" id="idCFinEndDate">&nbsp;</span>
														</td>
														<td class='FieldCellSub'></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>-->

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                                <%if trim(sFor)="GL" then %>
                                                    <input type="button" value="Transfer Accounts" name="btnCloseAcc" class="ActionButtonX" onclick="AccSubmit()">
                                                <%elseif trim(sFor)="AG" then %>
                                                    <input type="button" value="Final Closing" name="btnFinClose" class="ActionButtonX">
                                                    <input type="button" value="Audit Closing" name="btnCloseAcc" class="ActionButtonX" onclick="AccSubmit()">
                                                <%elseif trim(sFor)="IS" then %>
                                                    <input type="button" value="Close Stock" name="btnCloseStk" class="ActionButtonX" onclick="StockSubmit()">
                                                    <input type="Button" value="Setup" name="btnSetup" class="ActionButtonX" onclick="ItemStockSetup()">
                                                <%elseif trim(sFor)="NS" then %>
                                                    <input type="button" value="No Series Transfer" name="btnCloseStk" class="ActionButtonX" onclick="NoSeriesSubmit()">
                                                <%end if %>
												<!--<input type="button" value="Proceed" name="Proceed" class="ActionButton" onclick="CheckSubmit()">-->
                                                <input type="reset" value="Reset" name="Reset" class="ActionButton">
                                                <input type="button" value="Cancel" name="cancel" class="ActionButton" OnClick="window.location.href='../../CreateFinYear.asp'">
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

<%
	' Function which populates the Previous Financial Year Start Date list
	Function populateFinDate()
		' Declaration of variables
		Dim dcrs,sStartDate,sEndDate
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CONVERT(CHAR,FROMPERIOD,103),CONVERT(CHAR,TOPERIOD,103) FROM MS_FINANCIALPERIOD WHERE ACTIVE IN ('N','Y') ORDER BY 1 DESC"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sStartDate = dcrs(0)
		set sEndDate = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sStartDate)&":"&trim(sEndDate)&""">"&trim(sStartDate)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

<%
Function CheckNoSeriesTransfer(iApplCode,sUnitCode,dCurFinStartDate,dCurFinEndDate)
	iCurPeriodFrom = right(dCurFinStartDate,4)&mid(dCurFinStartDate,4,2)
	iCurPeriodTo = right(dCurFinEndDate,4)&mid(dCurFinEndDate,4,2)
	sQuery = "SELECT Count(1) FROM APP_R_NOSERIESMODULEENTRY WHERE (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) IN (SELECT (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) FROM APP_R_NOSERIESMODULES WHERE OUDEFINITIONID = " & Pack(sUnitCode) & " AND APPLICATIONCODE = "&iApplCode&" AND (PERIOD >= " & iCurPeriodFrom & " AND PERIOD <= " & iCurPeriodTo & "))"
	'Response.Write "<p>"&sQuery
	'Response.End
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		CheckNoSeriesTransfer = dcrs(0)
	Else
		CheckNoSeriesTransfer = 0
	end if
	dcrs.Close
End Function
%>
