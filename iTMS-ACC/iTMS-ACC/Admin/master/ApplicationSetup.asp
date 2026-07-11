
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ApplicationSetup.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 30,2012
	'Modified On				:
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

<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/populate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID = "AppData" src="<%="../temp/ApplicationSetup_"& session.sessionID&".xml"%>"></XML>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script>
function applicationSetupTrim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function applicationSetupXml() {
	var element;
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.upgradeXmlIslands(document);
	}
	element = document.getElementById("AppData");
	return window.AppData || document.AppData || element && element._itmsXmlIsland || element || null;
}

function applicationSetupDocument() {
	var data = applicationSetupXml();
	return data && (data.XMLDocument || data._doc || data) || null;
}

function applicationSetupRoot() {
	var data = applicationSetupXml();
	var doc = applicationSetupDocument();
	return data && data.documentElement || doc && doc.documentElement || null;
}

function applicationSetupChildElements(node, name) {
	var result = [];
	var wanted = name ? String(name).toLowerCase() : "";
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
			result.push(node.childNodes[i]);
		}
	}
	return result;
}

function applicationSetupRadioValue(name) {
	var item = document.formname.elements[name];
	if (!item) {
		return "N";
	}
	if (item.length && !item.tagName) {
		return item[0] && item[0].checked ? item[0].value : "N";
	}
	return item.checked ? item.value : "N";
}

function Validate() {
	var processCode = applicationSetupTrim(document.formname.selProcess.value);
	document.formname.hProcessName.value = processCode;
	document.formname.submit();
}

function ChkReset() {
	document.formname.hProcessName.value = "";
	document.formname.submit();
}

function Paginate(nPage) {
	document.formname.hPageSelection.value = nPage;
	document.formname.submit();
}

function PopulatePractice() {}

function FinalSubmit() {
	var root = applicationSetupRoot();
	var rows = applicationSetupChildElements(root, "Row");
	var request;
	rows.forEach(function (row) {
		var entryNo = row.getAttribute("EntryNo") || "";
		var appCode = row.getAttribute("AppCode") || "";
		var refCode = row.getAttribute("RefCode") || "";
		var suffix = entryNo + "Z" + appCode + "Z" + refCode;
		applicationSetupChildElements(row, "Det").forEach(function (detail) {
			detail.setAttribute("MREntry", applicationSetupRadioValue("radMRZ" + suffix));
			detail.setAttribute("IssEntry", applicationSetupRadioValue("radIssZ" + suffix));
			detail.setAttribute("RcptAcc", applicationSetupRadioValue("radRecAccZ" + suffix));
			detail.setAttribute("GRN", applicationSetupRadioValue("radGRNZ" + suffix));
			detail.setAttribute("ConEntry", applicationSetupRadioValue("radCONZ" + suffix));
			detail.setAttribute("ManPOS", applicationSetupRadioValue("radPOSZ" + suffix));
			detail.setAttribute("ComRcptBill", applicationSetupRadioValue("radCRBZ" + suffix));
			detail.setAttribute("AutoAcc", applicationSetupRadioValue("radAAccZ" + suffix));
		});
	});
	request = new XMLHttpRequest();
	request.open("POST", "XMLSave.asp?Name=ApplicationSetup", false);
	request.send(applicationSetupDocument());
	document.formname.action = "ApplicationSetupInsert.asp";
	document.formname.submit();
}
</Script>
<% 
    Dim objRS,rsTemp,objDOM
    Dim ndRoot,ndRow,ndDet
    Dim iCurrentPage,iSetupEntryNo,iCtr
    Dim iTotalPage
    Dim sApplicationCode,sRefCodeNo,sAutoMREntry,sAutoIssEntry,sAutoRcptAcc,sAutoGRN,sAutoConEntry,sManPOS,sComRcptBill,sAutoAcc
    Dim sUnitID,sProcessCode,sQuery,sAppName
    
    set objRS = server.CreateObject("ADODB.Recordset")
    set rsTemp =Server.CreateObject("ADODB.Recordset")
    set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	Const iPageSize = 20
	
	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	if trim(sUnitID) = "" then	sUnitID = Session("organizationcode")

	sProcessCode  = Request("selProcess")
	
	Set ndRoot = objDOM.createElement("Root")
	objDOM.appendChild ndRoot

%>
</Head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="">
<form method="POST" name="formname">

	<Input type="hidden" name="hProcessName" value="<%=sProcessCode%>">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0" >
		<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Application Setup
		</td>
		</tr>

		<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >

				<tr>
				<TD class=TabBody>
					<table border="0" cellpadding="0" cellspacing="0" >
						<tr>
						<td align="center" colspan="3" class="MiddlePack" height="7" >
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						</tr>

						<tr>
						<td align="center" width="5" class="ClearPixel">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						<td valign="top" width="100%">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
						<tr>
						<td>
						<div>
						<table class="CollapseBand" cellspacing="0" cellpadding="0" >
						<tr>
						<td valign="center">
						<a style="width: 1em; height: 1em;" title="" onclick="return Div_OnClick(idUnprocessed,event);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
						</a>
						</td>
						<td valign="right" class="SubTitle">
						</td>
						</tr>
						</table>

						<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
						<td width="100%">
						<div id="idUnprocessed" style="display: none">
						<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
						<tr>
							<td class="FieldCellSub">Application Name</td>
							<td class="FieldCellSub">
								<select size="5" name="selProcess" class="FormElem" onChange="PopulatePractice(this.value)">
									<option value="0" selected>Select</option>
								<% 
									with objRs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
									.ActiveConnection = con
									.Open
								end with
								set objRs.ActiveConnection = nothing

								if not objRs.EOF then

									Do While Not objRs.EOF
											%><option value=<%=trim(objRs(0))%>><%=trim(trim(objRs(1)))%></option><%
									objRs.MoveNext
									Loop
								end if
								objRs.Close

								%>
								</select>
							</td>
						</tr>
						<tr>
							<td class="FieldCellSub"></td>
							<td class="FieldCellSub" >
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
						<td>
						</td>
						<td valign="top" width="100%">
							<table border="0" cellspacing="1" class="ExcelTable" width="100%">
							    <tr>
							        <td class="ExcelHeaderCell" align="center" colspan="2" >Application Name</td>
							    </tr>
								<tr>
									<td class="ExcelHeaderCell" align="center" >Activity</td>
									<td class="ExcelHeaderCell" align="center" >Options</td>
								</tr>
								<%
								    sQuery = "Select SetupEntryNo,ApplicationCode,ReferenceCodeNo,isNull(AutomaticMREntry,'N'),isNull(AutomaticIssueEntry,'N'),"
                                    sQuery = sQuery & "isNull(AutomaticRcptAccounting,'N'),isNull(AutomaticGatepassEntry,'N'),isNull(AutomaticConsumptionEntry,'N'),"
                                    sQuery = sQuery & "isNull(MandatoryPOS,'N'),isNull(CommonRcptBillEntry,'N'),isNull(AutomaticAccounting ,'N') from APP_M_ApplicationSetup"
                                    objRS.Open sQuery,con
                                    if not objRS.EOF then
                                        do while not objRS.EOF
                                            iSetupEntryNo = objRS(0)
                                            sApplicationCode = objRS(1)
                                            sRefCodeNo = objRS(2)
                                            sAutoMREntry = objRS(3)
                                            sAutoIssEntry = objRS(4)
                                            sAutoRcptAcc = objRS(5)
                                            sAutoGRN = objRS(6)
                                            sAutoConEntry = objRS(7)
                                            sManPOS = objRS(8)
                                            sComRcptBill =  objRS(9)
                                            sAutoAcc = objRS(10)
                                            
                                            set ndRow = objDOM.createElement("Row")
                                                ndRow.setAttribute "EntryNo",iSetupEntryNo
                                                ndRow.setAttribute "AppCode",sApplicationCode 
                                                ndRow.setAttribute "RefCode",sRefCodeNo
                                            ndRoot.appendChild ndRow
                                            
                                            set ndDet = objDOM.createElement("Det")
                                                ndDet.setAttribute "MREntry",sAutoMREntry 
                                                ndDet.setAttribute "IssEntry",sAutoIssEntry 
                                                ndDet.setAttribute "RcptAcc",sAutoRcptAcc 
                                                ndDet.setAttribute "GRN",sAutoGRN 
                                                ndDet.setAttribute "ConEntry",sAutoConEntry 
                                                ndDet.setAttribute "ManPOS",sManPOS
                                                ndDet.setAttribute "ComRcptBill",sComRcptBill
                                                ndDet.setAttribute "AutoAcc",sAutoAcc 
                                            ndRow.appendChild ndDet 
                                            
                                            sQuery = "Select ApplicationName from Ms_Applications where ApplicationCode = "& sApplicationCode 
                                            rsTemp.Open sQuery,con
                                            if not rsTemp.EOF then
                                                sAppName = ucase(Trim(rsTemp(0)))
                                            end if
                                            rsTemp.Close 
                                                %>
                                                    <tr>
                                                        <td class="ExcelDisplayCell" colspan="3" align="center"><b><%=sAppName%></b></td>
                                                    </tr>
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic MR Entry</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoMREntry)="Y" then %>
                                                                <input type="radio" name="radMRZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radMRZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radMRZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radMRZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic Issue Entry</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoIssEntry)="Y" then %>
                                                                <input type="radio" name="radIssZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radIssZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radIssZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radIssZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic Receipt Accounting</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoRcptAcc)="Y" then %>
                                                                <input type="radio" name="radRecAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radRecAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radRecAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radRecAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic Gate Receipt Entry</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoGRN)="Y" then %>
                                                                <input type="radio" name="radGRNZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radGRNZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radGRNZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radGRNZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic Consumption Entry</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoConEntry)="Y" then %>
                                                                <input type="radio" name="radCONZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radCONZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radCONZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radCONZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                    
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Mandatory POS</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sManPOS)="Y" then %>
                                                                <input type="radio" name="radPOSZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radPOSZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radPOSZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radPOSZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Common Receipt Bill Entry</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sComRcptBill)="Y" then %>
                                                                <input type="radio" name="radCRBZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radCRBZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radCRBZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radCRBZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                    <tr>
                                                        <td class="ExcelDisplayCell">Automatic Accounting</td>
                                                        <td class="ExcelDisplayCell">
                                                            <% if Trim(sAutoAcc)="Y" then %>
                                                                <input type="radio" name="radAAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y" checked>Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radAAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N">No&nbsp;&nbsp;
                                                            <%else%>
                                                                <input type="radio" name="radAAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="Y">Yes&nbsp;&nbsp;
                                                                <input type="radio" name="radAAccZ<%=iSetupEntryNo%>Z<%=sApplicationCode%>Z<%=sRefCodeNo%>" value="N" checked>No&nbsp;&nbsp;
                                                            <%end if' if Trim(sAutoMREntry)="Y" then %>
                                                        </td>
                                                    </tr>
                                                    
                                                <%
                                            objRS.MoveNext 
                                        loop
                                    end if
                                    objRS.Close 
                                    objDOM.save Server.MapPath("../temp/ApplicationSetup_"& Session.SessionID &".xml")
								%>
							<input type="hidden" name="hCnt" value="<%=iCtr-1%>">

							</table>
							</td>
                            </tr>

							<tr>
								<td colspan="3" class="MiddlePack">
								</td>
							</tr>

							<tr>
							<td align="center" width="5" class="ClearPixel">
							</td>
							<td valign="top" align="right">
							<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
							<input type=hidden name="hPageSelection" value="0">

							<%	If iTotalPage >= 2 Then
							if iCurrentPage = 1 then
							%>
							<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
							<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
							<%		else%>
							<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
							<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
							<%		end if	%>
							<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
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
							<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
							<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPage%>')" id=button8 name=button8>
							<%		end if
							End If
							%>
							</td>
							<td align="center" class="ClearPixel" width="5">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							</tr>

							<tr>
							<td>
								<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top">

							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="middle" class="ActionCell">
									<p align="center">
										<Input type="button" value="Update" name="btnUpdate" class="ActionButton" tabindex="3" onclick="FinalSubmit()">
										<!--<Input type="button" value="Edit" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction('EDT')" onclick="GotoAction('CRN')">-->
									</td>
								</tr>
							</table>

							</td>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack">
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
<%
Function populateUnits()
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM VWITEMLIST) ORDER BY ORGANIZATIONUNITID"
		.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close

End Function

' Function to populate Applications / Process
Function PopulateProcess()
		' Declaration of variables
		Dim dcrs
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			Do While Not dcrs.EOF

				Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

			dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
