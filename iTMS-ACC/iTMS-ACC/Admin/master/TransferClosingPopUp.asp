<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	TransferClosingPopUp.asp
	'Module Name				:	Transfer Closing Values
	'Author Name				:	S.Maheswari
	'Created On					:	Mar 19, 2009
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Transfer Closing Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Root/></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT>
function transferClosingField(name) {
	var frm = document.formname;
	return frm && (frm.elements[name] || frm[name]) || null;
}

function transferClosingXmlObject() {
	var element;
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.upgradeXmlIslands(document);
	}
	element = document.getElementById("OutData");
	return window.OutData || element && element._itmsXmlIsland || element || null;
}

function transferClosingPost(url, body) {
	var request = new XMLHttpRequest();
	request.open("POST", url, false);
	request.send(body || null);
	return request.responseText || "";
}

function CreateXML() {
	var data = transferClosingXmlObject();
	var doc = data && (data.XMLDocument || data._doc || data);
	var root = data && data.documentElement || doc && doc.documentElement;
	var count = parseInt(transferClosingField("hCnt").value, 10) || 0;
	var responseText;
	var i;
	var element;
	if (!doc || !root) {
		alert("No. Series data is not available.");
		return false;
	}
	while (root.firstChild) {
		root.removeChild(root.firstChild);
	}
	for (i = 1; i <= count; i += 1) {
		element = doc.createElement("NoSeries");
		element.setAttribute("ExistingSuffix", transferClosingField("hExsSuff" + i).value);
		element.setAttribute("ChangeSuffix", transferClosingField("txtChgSuff" + i).value);
		root.appendChild(element);
	}
	transferClosingPost("XMLSave.asp?Name=NoSeries", doc);
	responseText = transferClosingPost("NoSeriesDetailsEntry.asp?Temp=" + transferClosingField("hTemp").value);
	alert(responseText);
	window.close();
	return true;
}
</SCRIPT>
</HEAD>
<%
Dim sAccName,sFor,sQuery
dim dcrs,sUnit,sUnitName,dPreFinStartDate,dPreFinEndDate,dCurFinStartDate,dCurFinEndDate
dim iApplication, iPrePeriodFrom,iPrePeriodTo
dim Passpara,Arr
Passpara = Request("Para")
'sUnit = trim(Request("Unit"))
	'sUnitName = trim(Request("OrgName"))
	'dCurFinStartDate = trim(Request("CFinStartDate"))
	'dCurFinEndDate = trim(Request("CFinEndDate"))
	'dPreFinStartDate = trim(Request("PFinStartDate"))
	'dPreFinEndDate = trim(Request("PFinEndDate"))
	'iApplication = trim(Request("Application"))
IF trim(Passpara) <> "" then
	Arr = Split(Passpara,"||")
	sUnit = trim(Arr(0))
	sUnitName = trim(Arr(1))
	dCurFinStartDate = trim(Arr(2))
	dCurFinEndDate = trim(Arr(3))
	dPreFinStartDate = trim(Arr(4))
	dPreFinEndDate = trim(Arr(5))
	iApplication = trim(Arr(6))


End IF

iPrePeriodFrom = right(dPreFinStartDate,4)&mid(dPreFinStartDate,4,2)
iPrePeriodTo = right(dPreFinEndDate,4)&mid(dPreFinEndDate,4,2)

' Create our DOM Document Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")

%>


<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type=hidden name="hTemp" value="<%=Passpara%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20">
			<p align="center">No. Series Update
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
								<td align="center">
								</td>

								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>

							<tr>
								<td align="center"></td>
								<td valign="top">
									<table border="0" cellspacing="1" class="ExcelTable">
										<tr>
											<td class="ExcelHeaderCell" align="center" >Existing Suffix</td>
											<td class="ExcelHeaderCell" align="center" >Change Suffix</td>
										</tr>
										<% dim  iCtr
										sQuery = "SELECT DISTINCT ISNULL(SUFFIX,'-') AS SUFFIX FROM APP_R_NOSERIESMODULEENTRY WHERE (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) IN (SELECT (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) FROM APP_R_NOSERIESMODULES WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND APPLICATIONCODE = " & iApplication & " AND (PERIOD >= " & iPrePeriodFrom & " AND PERIOD <= " & iPrePeriodTo & ")) ORDER BY 1"
										'Response.Write sQuery
										dcrs.Open sQuery,con
										iCtr = 0
										do while not dcrs.EOF

											'Response.Write dcrs(0) &"<BR><BR>"
											If trim(dcrs(0)) <> "" then
											iCtr = iCtr + 1
										%>
										<tr>
											<td class="ExcelHeaderCell"  align="center" valign="middle" colspan="5"><%=sAccName%></td>
										</tr>

										<tr>
											<td class="FieldCellSub" align="Left"><%=dcrs(0)%></td>
											<td class="FieldCell" >
												<input type="text" name="txtChgSuff<%=iCtr%>" value="<%=dcrs(0)%>" class="FormElem" size="20">
											</td>
											<input type=hidden name="hExsSuff<%=iCtr%>" value="<%=dcrs(0)%>">

										</tr>




									<%		End IF
									dcrs.MoveNext
										loop
										dcrs.Close
										'Response.end  %>
										<input type=hidden name="hCnt" value="<%=iCtr%>">
									</table>
								</td>
								<td align="center"></td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
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
												<input type="button" value="Transfer" name="ButTrans" class="ActionButton" OnClick="CreateXML()">
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
