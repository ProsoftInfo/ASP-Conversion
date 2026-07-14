<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	itmCodeDisplay.asp
	'Module Name				:	Inventory (Item Code Display)
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 29, 2004
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Code Display</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
</HEAD>
<%
	dim dcrs,iRecCount,arrCode,sItmType,iStart,iCount
	dim sItemCode,sItemDesc,iItemCode
	dim sWeave,iWidth,iReedCount,sReedSpace,iWeight,sVariety,iNODent,iNOTotal,iNOEnds
	dim iNOPicks,sAverageWC,iTapeLength,sWarpYarn,sWeftYarn

	set dcrs = server.CreateObject("ADODB.RecordSet")

	sItmType = Request.QueryString("ItemType")
	sItemCode = Request.QueryString("ItemCode")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMDESCRIPTION,ITEMCODE FROM INV_M_ITEMMASTER WHERE COMPANYITEMCODE = " & Pack(sItemCode)
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.eof then
		sItemDesc = trim(dcrs(0))
		iItemCode = trim(dcrs(1))
	end if
	dcrs.close

	if sItmType = "FAB" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(WEAVE,'-'),ISNULL(WIDTH,0),ISNULL(REEDCOUNT,0),ISNULL(REEDSPACE,'-'),ISNULL(WEIGHT,0),ISNULL(VARIETY,'-'),ISNULL(ENDSDENT,0),ISNULL(ENDSTOTAL,0),ISNULL(ENDSINCH,0),ISNULL(PICKSINCH,0),ISNULL(WRAPCOUNT,'-'),ISNULL(TAPELENGTH,0),ISNULL(WARPYARN,0),ISNULL(WEFTYARN,0) FROM INV_M_ITEMCONSTRUCTION WHERE ITEMCODE = " & iItemCode
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		do while not dcrs.eof
			sWeave = trim(dcrs(0))
			iWidth = trim(dcrs(1))
			iReedCount = trim(dcrs(2))
			sReedSpace = trim(dcrs(3))
			iWeight = trim(dcrs(4))
			sVariety = trim(dcrs(5))
			iNODent = trim(dcrs(6))
			iNOTotal = trim(dcrs(7))
			iNOEnds = trim(dcrs(8))
			iNOPicks = trim(dcrs(9))
			sAverageWC = trim(dcrs(10))
			iTapeLength = trim(dcrs(11))
			sWarpYarn = sWarpYarn & "," & trim(dcrs(12))
			sWeftYarn = sWeftYarn & "," & trim(dcrs(13))
			dcrs.movenext
		loop
		dcrs.close
		sWarpYarn = mid(sWarpYarn,2)
		sWeftYarn = mid(sWeftYarn,2)

		if sWeave = "" then sWeave = "-"
		if iWidth = "" or iWidth = 0 then iWidth = "-"
		if iReedCount = "" or iReedCount = 0 then iReedCount = "-"
		if sReedSpace = "" then sReedSpace = "-"
		if iWeight = "" or iWeight = 0 then iWeight = "-"
		if sVariety = "" then sVariety = "-"
		if iNODent = "" or iNODent = 0 then iNODent = "-"
		if iNOTotal = "" or iNOTotal = 0 then iNOTotal = "-"
		if iNOEnds = "" or iNOEnds = 0 then iNOEnds = "-"
		if iNOPicks = "" or iNOPicks = 0 then iNOPicks = "-"
		if sAverageWC = "" then sAverageWC = "-"
		if iTapeLength = "" or iTapeLength = 0 then iTapeLength = "-"
	end if

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CODETYPE,CODETYPENAME,CODELENGTH FROM APP_M_CODETYPES WHERE ITEMTYPEID = " & Pack(sItmType) & " ORDER BY DISPLAYORDER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	iRecCount = cint(dcrs.recordCount) - 1

	if not dcrs.EOF then
		arrCode = dcrs.getRows()
	end if
	dcrs.Close

%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname">
<input type="hidden" name="hCount" value="<%=iRecCount%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Code Display
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
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
											<td class="FieldCell" width="90">Item Code</td>
											<td class="FieldCell" colspan="3">
												<Span Class="DataOnly"><%=sItemCode%>&nbsp;</Span>
											</td>
										</tr>
                                        <tr>
											<td class="FieldCell" width="90">Description</td>
											<td class="FieldCell" colspan="3">
												<Span Class="DataOnly"><%=sItemDesc%>&nbsp;</Span>
											</td>
										</tr>
 									</table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan=3 class="MiddlePack">
							        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						<%
							if sItmType = "FAB" then
						%>
							<tr>
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                   <table border="0" cellpadding="0" cellspacing="1" class="TableOutlineOnly">
                                        <tr>
                                            <td class="FieldCellSub">Weave</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=sWeave%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends / dent</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iNODent%>&nbsp;</Span>
											</td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Width</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iWidth%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends - Total</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iNOTotal%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Reed Count</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iReedCount%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends / inch(avg)</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iNOEnds%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Reed Space</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=sReedSpace%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of Picks / inch(avg)</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iNOPicks%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Weight (lbs/yrds)</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iWeight%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">Average Wrap Count</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=sAverageWC%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Variety</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=sVariety%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">Tape Length (%)</td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=iTapeLength%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Warp Yarn</td>
                                            <td class="FieldCell" colspan=4 width="175">
												<Span Class="DataOnly"><%=DisplayYarn(sWarpYarn)%>&nbsp;</Span>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Weft Yarn</td>
                                            <td class="FieldCell" colspan=4 width="175">
												<Span Class="DataOnly" ><%=DisplayYarn(sWeftYarn)%>&nbsp;</Span>
											</td>
                                            <td class="FieldCellSub"></td>
                                        </tr>
									</table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan=3 class="MiddlePack">
							        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						<%
							end if
						%>

							<tr>
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
									<%
										iStart = 1
										for iCount = 0 to iRecCount
									%>
                                        <tr>
                                            <td class="FieldCell"><%=trim(arrCode(1,iCount))%></td>
                                            <td class="FieldCell">
												<Span Class="DataOnly"><%=DisplayCode(arrCode(0,iCount),mid(sItemCode,iStart,arrcode(2,icount)))%>&nbsp;</Span>
                                            <%
												iStart = iStart + cint(arrcode(2,icount))
												iCount = iCount + 1
                                            %>
                                            </td>
											<%
												if 	iCount > iRecCount then exit for
											%>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell"><%=trim(arrCode(1,iCount))%></td>
                                            <td class="FieldCell">
											<%	if sItmType = "FAB" and iCount = 1 then	%>
													<Span Class="DataOnly"><%=mid(sItemCode,iStart,arrcode(2,icount))%>&nbsp;</Span>
											<%	else	%>
													<Span Class="DataOnly"><%=DisplayCode(arrCode(0,iCount),mid(sItemCode,iStart,arrcode(2,icount)))%>&nbsp;</Span>
											<%
												end if
												iStart = iStart + cint(arrcode(2,icount))
											%>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>
									<%
										next
									%>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3"></td>
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
                                                    <input type="button" value="Close" name="B6" class="ActionButton" onClick=window.close();>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
	' Function to populate Codes List
	Function DisplayCode(iID, iValue)
		Dim dcrs

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODENAME FROM APP_M_CODEMASTER WHERE CODETYPE = " & Pack(iID) & " AND CODE = " & Pack(iValue)
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if Not dcrs.EOF then
			DisplayCode = trim(dcrs(0))
		else
			DisplayCode = ""
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate Yarn List
	Function DisplayYarn(iCode)
		Dim dcrs,sYarnName

		if iCode = "" then
			DisplayYarn = "-"
			exit function
		end if

		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMDESCRIPTION FROM INV_M_ITEMMASTER WHERE ITEMCODE IN (SELECT ITEMCODE FROM INV_M_ITEMORGMASTER WHERE CLASSIFICATIONCODE IN (SELECT CLASSIFICATIONCODE FROM INV_M_CLASSIFICATION WHERE ITEMTYPEID = 'YRN')) AND ITEMCODE IN (" & iCode & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		Do While Not dcrs.EOF
'			Response.Write("<TR><SPAN CLASS=""DataOnly"">"&trim(dcrs(0))&"</Span></TR>" &vbcrlf)

			sYarnName = sYarnName & ", " & trim(dcrs(0))
			dcrs.MoveNext
		Loop
		dcrs.Close
		DisplayYarn = mid(sYarnName,3)
	End Function
%>
