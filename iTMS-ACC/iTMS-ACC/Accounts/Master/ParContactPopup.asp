<%@ Language=VBScript %>
<%	option explicit	
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ParContactPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%
'XML DOM Variables
Dim oDOM,nodHeader,Root,sParName,sParCode,Objrs,sQuery,iPartyCode,sRecCount
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set Objrs = Server.CreateObject("ADODB.RecordSet")

iPartyCode = Request.QueryString("PartyCode")

	sQuery = "Select PartyCode,OrgnPartyCode,PartyName from APP_M_PartyMaster where PartyCode = "& iPartyCode
'	Response.Write sQuery
	Objrs.Open sQuery,con
	if not Objrs.EOF then
		sParCode = Objrs(1)
		sParName = Objrs(2)
	end if
	Objrs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self">
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<!-- XML Data Island -->
<XML id=OutData>
<Contact>
<%
	if trim(iPartyCode)<>"" then
		sQuery = "Select isNull(ContactNo,0),isNull(ContactPersonName,''),isNull(Designation,''),isNull(ContactPersonFor,''),isNull(ContactMailID,'') "&_
				 "From APP_M_PartyContactPersons Where PartyCode = "&iPartyCode&" "
		With Objrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		sRecCount = Objrs.RecordCount
		Set Objrs.ActiveConnection = Nothing
		Do While Not Objrs.EOF

	%>
	<Entry No="<%=Objrs(0)%>" Name="<%=Objrs(1)%>" Desig="<%=Objrs(2)%>" PersonFor="<%=Objrs(3)%>" Maillid="<%=Objrs(4)%>" />
	<%
		Objrs.MoveNext
		loop
		Objrs.Close
	end if 'if trim(iPartyCode)<>"" then
%>
</Contact>
</xml>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script>
function validate()
{
	if (trim(document.formname.txtName.value)=="")
	{
		alert("Enter Name");
		document.formname.txtName.select();
		return false;
	}

	if (trim(document.formname.txtDesignation.value)=="")
	{
		alert("Enter Designation");
		document.formname.txtDesignation.select();
		return false;
	}
	if (trim(document.formname.txtContact.value)=="")
	{
		alert("Enter Contact For");
		document.formname.txtContact.select();
		return false;
	}
	if (trim(document.formname.txtMailId.value)!="")
	{
		if (checkmailid(document.formname.txtMailId.value)==false)
		{
			document.formname.txtMailId.select();
			return false;
		}
	}
	//else
	//{
	//	if (checkmailid(document.formname.txtMailId.value)==false)
	//	{
	//		document.formname.txtMailId.select();
	//		return false;
	//	}
	//}

	return true;
}
</script>
<script>
window.__itmsPopupCompat = { type: "partyContactPopup" };
</script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="PopulateContact()">

<form method="POST" name="formname">
<input type="hidden" name="hRecCount" value="<%=sRecCount+1%>">
<input type="hidden" name="hEntNo" value="">
<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Contacts</p>
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
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
																										<tr>
															<td class="FieldCell" width="72" valign="top">Party Name</td>
															<td class="FieldsubCell" ><span class="DataOnly"><%=sParName%></td>
													</tr>
													<tr>
															<td class="FieldCell" width="72" valign="top">Party Code</td>
															<td class="FieldsubCell"><span class="DataOnly"><%=sParCode%></td>
													</tr>

														<!--tr>
															<td class=FieldCell width="72"> Select Name</td>
															<td class='FieldCell'>
															<Select name="selContact" class="FormElem" onChange="DispCon()">
															<Option value="0">Select Contact </option>

															</td>
														</tr-->

														<tr>
															<td class=FieldCell width="72"> Name</td>
															<td class='FieldCell'><input type="text" name="txtName" size="65" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="72"> Designation</td>
															<td class='FieldCell'><input type="text" name="txtDesignation" size="35" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="72"> Contact For</td>
															<td class='FieldCell'><input type="text" name="txtContact" size="50" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCell width="72"> E-mail ID</td>
															<td class='FieldCell'><input type="text" name="txtMailId" size="50" class="Formelem"></td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
                                                                <input type="button" value="Add Next" name="btnAdd" class="ActionButton" onClick="addEntry('A')"   >
                                                                <input type="button" value="Update" name="btnUpdate" class="ActionButton" onClick="updateEntry()" disabled  >
                                                                <input type="button" value="Delete Contact" name="btnDel" class="ActionButtonX" onClick="DelEntry()" disabled   >
                                                                <input type="button" value="Save" name="btnNext" onClick="addEntry('S')" class="ActionButton"  >
                                                                <input type="button" value="Close" name="btnCancel" onClick="window.close()"  class="ActionButton" >
                                                                <!--<input type="Button" value="Reset" name="B1" onClick="clearXML()" class="ActionButton" >-->
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
							<tr>
								<td align="center">
								</td>
								<td valign="top">
												<DIV class=frmBody id=frm1 style="width: 585; height:140;">
                                                <table border="0" id="tblBin"cellspacing="1" class="ExcelTable" width="100%">
												    <tr>
												<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10"><p align="center">&nbsp;</td>
												<td class="ExcelHeaderCell" align="center">Name</td>
												<td class="ExcelHeaderCell" align="center"> Designation</td>
												<td class="ExcelHeaderCell" align="center">Contact For</td>
												<td class="ExcelHeaderCell" align="center">Email ID</td>
												    </tr>

                                                </table>
												</div>
								</td>
								<td align="center">
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
