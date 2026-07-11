<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParTypeCreatEntryPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 21,2011
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
<%
    Dim rsObj
    Dim sQuery,iSNo,sAction,sParType,sParSubType,sSubTypeName,sSubTypeShortName,iCount
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sAction = Request.QueryString("Action")
    if trim(sAction)="" then sAction =  "C"
    if trim(sAction)="E" then
        sParType = Request.QueryString("ParType")
        sParSubType = Request.QueryString("ParSubType")
    end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<base target="_self"></base>
<TITLE>Party Sub Type</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
function finaldone()
{
	if ((document.formname.radParType(0).checked==true)||(document.formname.radParType(1).checked==true))
	{
		if (trim(document.formname.txtSubTypeName.value)=="")
		{
			alert("Enter Sub-Type Name");
			document.formname.txtSubTypeName.select();
			return false;
		}
		if (trim(document.formname.txtSubTypeShortName.value) =="")
		{
			alert("Enter Sub-Type Short Name");
			document.formname.txtSubTypeShortName.select();
			return false;
		}

	}
	else
	{
			alert("Select Party Type");
			document.formname.radParType(0).select();
			return false;
	}
return true;
}
</script>
<script>
window.__itmsPopupCompat = { type: "partySubtypeEntry" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
<script src="../../scripts/itms-modern-compat.js"></script>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Sub Type</p>
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
								<td align="center" rowspan="3" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
													    <%
													        sQuery = "Select PartyType,PartySubType,SubTypeName,SubTypeShortName from APP_M_PartyTypes where PartyType = '"& sParType &"' and PartySubType = '"& sParSubType &"'"
													        rsObj.Open sQuery,con
													        if not rsObj.EOF then
													            sSubTypeName = trim(rsObj(2))
													            sSubTypeShortName = trim(rsObj(3))
													        end if
													        rsObj.Close
													    %>

														<tr>
															<td class=FieldCell width="155"> Party Type</td>
															<td class='FieldCell'>
															<%if trim(sParType)="CR" then %>
                                                                <input type="radio" value="CR" name="radParType" class="FormElem" checked>  Creditor &nbsp;
                                                                <input type="radio" value="DR" name="radParType" class="FormElem">  Debtor
                                                            <%elseif trim(sParType)="DR" then %>
                                                                <input type="radio" value="CR" name="radParType" class="FormElem">  Creditor &nbsp;
                                                                <input type="radio" value="DR" name="radParType" class="FormElem"  checked>  Debtor
                                                            <%else %>
                                                                <input type="radio" value="CR" name="radParType" class="FormElem">  Creditor &nbsp;
                                                                <input type="radio" value="DR" name="radParType" class="FormElem">  Debtor
                                                            <%end if %>
                                                            </td>
														</tr>

														<tr>
															<td class=FieldCell width="155" valign="top"> Sub-Type Name</td>
															<td class='FieldCell'>
                                                <input type="text" name="txtSubTypeName" size="52" maxlength="50" class="FormElem" value="<%=sSubTypeName%>"></td>
														</tr>
														<tr>
															<td class=FieldCell width="155" valign="top"> Sub-Type Short Name</td>
															<td class='FieldCell'><input type="text" name="txtSubTypeShortName" size="12" maxlength="10" class="FormElem" value="<%=sSubTypeShortName%>" ></td>
														</tr>
													</table>
												</center>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center" rowspan="3" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															<%if trim(sAction)="C" then%>
                                                                <input type="button" value="Create" name="B2" class="ActionButton" tabindex="3" onClick="CheckSubmit('C','','')">
                                                            <%elseif trim(sAction)="E" then%>
                                                                <input type="button" value="Update" name="B2" class="ActionButton" tabindex="3" onClick="CheckSubmit('U','<%=sParType%>','<%=sParSubType%>')">
                                                            <%end if 'if trim(sAction)="C" then%>
																<input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >

														</td>
													</tr>
												</table>
								</td>
							</tr>
							<tr>
								<td valign="top" class="BottomPack">
								</td>
							</tr>
							<tr>
							    <td>
								<td colspan=2>
								    <div style="height:300px;width:400px">
								        <table border=0 cellspacing=1 cellpadding=0 class="ExcelTable" width=100%>
    								        <tr>
    								            <td class="ExcelHeaderCell" align=center>S.No</td>
    								            <td class="ExcelHeaderCell" align=center>
    								                <img src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width=15 height=15 onclick="CheckSubmit('D','','')">
    								            </td>
    								            <td class="ExcelHeaderCell" align=center>Party Sub Type Name</td>
    								        </tr>
    								        <%
    								            sQuery = "Select PartyType,PartySubType,SubTypeName,SubTypeShortName,SystemDefault from APP_M_PartyTypes"
    								            rsObj.open sQuery,con
    								            if not rsObj.eof then
    								                iSNo = 0
    								                iCount = 0
    								                do while not rsObj.eof
    								                    iSNo = iSNo +1
    								                    %>
    								                        <tr>
        								                    <td class="ExcelSerial" align=center><%=iSNo%></td>
        								                    <td class="ExcelDisplayCell" align=center>
        								                    <%if trim(rsObj(4))<>"Y" then
        								                        iCount = iCount + 1
        								                    %>
        								                    <input type=checkbox name="chkParType<%=iCount%>" value="<%=rsobj(0)&":"&rsobj(1)%>">
        								                    <%end if 'if trim(rsObj(4))<>"Y" then%>
        								                    </td>
                                                            <td class="ExcelDisplayCell"><a href="#" onclick="CheckSubmit('E','<%=rsObj(0)%>','<%=rsobj(1)%>'); return false;" class="ExcelDisplayLink" ><%=rsObj(2)%></a></td>
        								                    </tr>
    								                    <%
    								                    rsObj.movenext
    								                loop
    								            end if
    								            rsObj.close
    								        %>
    								        <input type=hidden name="hRowCtr" value="<%=iCount%>">
								        </table>
								    </div>
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
