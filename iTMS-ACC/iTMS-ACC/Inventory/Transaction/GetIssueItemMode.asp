<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetIssueItemMode.asp
	'Module Name				:	Inventory (Issue Additional Details)
	'Author Name				:	Ragavendran
	'Created On					:	Oct 10,2013
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<%
Dim sItemName,sOrgCode,sOrgName,sTempValues,sArrTempVal,sItemMode
Dim iItemCode,iClassCode
sTempValues= Request("Temp")

sArrTempVal = split(sTempValues,"|")

iClassCode = sArrTempVal(0)
iItemCode = sArrTempVal(1)
sOrgCode = sArrTempVal(2)
sItemMode = sArrTempVal(3)

sItemName = GetItemName(iItemCode,iClassCode)
sOrgName = GetUnitDesc(sOrgCode)

if trim(sItemMode)="" or IsNull(sItemMode) then sItemMode = "P"



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Additional Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function CheckSubmit()
    if document.formname.radType(0).checked then
        sItemMode =document.formname.radType(0).value
    elseif document.formname.radType(1).checked then
        sItemMode =document.formname.radType(1).value
    elseif document.formname.radType(2).checked then
        sItemMode =document.formname.radType(2).value
    end if
    
    window.returnvalue = sItemMode
    window.close()
End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Mode
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Type</td>
                                            <td class="FieldCellSub">
                                                <input type="radio" name="radType" value="P" <%if trim(sItemMode)="P" then response.write "Checked"%>>Primary
                                                <input type="radio" name="radType" value="C" <%if trim(sItemMode)="C" then response.write "Checked"%>>Consumable
                                                <input type="radio" name="radType" value="A" <%if trim(sItemMode)="A" then response.write "Checked"%>>Accessories
											</td>
                                        </tr>
                                    </table>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="window.close()">
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
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Work Center list
	Function populateWorkCenters()
		' Declaration of variables
		Dim dcrs,sWCID,sWCName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT WORKCENTERCODE,WORKCENTERNAME FROM VWWORKMACHINECENTER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sWCID = dcrs(0)
		set sWCName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sWCID)&""">"&trim(sWCName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Work Groups list
	Function populateWorkGroups()
		' Declaration of variables
		Dim dcrs,sWCID,sWCName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT WORKGROUPCODE,WORKGROUPNAME FROM PRD_M_WORKGROUP ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sWCID = dcrs(0)
		set sWCName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sWCID)&""">"&trim(sWCName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
