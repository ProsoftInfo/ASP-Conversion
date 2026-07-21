<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CommonConfirmPop.asp
	'Module Name				:	Inventory (Common)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	MARCH 23,2010
	'Modified On				:	Sep 09,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<%
	Dim sHeading,sCallFrom,sArrTemp,sTempValues,sIssType,sPickType,sLotOrPickFlag
	sHeading = Request.QueryString("sHead")
	sCallFrom = Request.QueryString("CallFrom")
	sTempValues = Request.QueryString("Issue")
	if trim(sTempValues)<>"" and (not isNull(sTempValues)) then
	    sArrTemp = split(sTempValues,":")
	    sIssType = sArrTemp(0)
	    sPickType = sArrTemp(1)
	    sLotOrPickFlag = sArrTemp(2)
	end if
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root Confirm="N"/></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
'Dim objTemp
'set objTemp = window.dialogArguments
'***********************************************************
Function FinalSubmit()
Dim sInvValue,sProValue,sIssType,sPickType,sLotOrPackFlag
	set Root = RefData.documentElement
	
	sIssType = document.formname.hIssType.value
	sPickType = document.formname.hPickType.value
	sLotOrPackFlag = document.formname.hLotOrPickFlag.value
	
	if trim(document.formname.hCallFrom.value)="DIS" then
	    if trim(sIssType)="M" and trim(sPickType)="N" and trim(sLotOrPackFlag)="P" then
	        if document.formname.radConfirm(0).checked = true then
		        Root.setAttribute "Confirm","Y"
	        else
		        Root.setAttribute "Confirm","N"
	        end if
	    elseif trim(sIssType)="F" then
	        if document.formname.radConfirm(0).checked = true then
		        Root.setAttribute "Confirm","Y"
	        else
		        Root.setAttribute "Confirm","N"
	        end if
	    else
	        Root.setAttribute "Confirm","N"
	    end if
	else
	    if document.formname.radConfirm(0).checked = true then
		    Root.setAttribute "Confirm","Y"
	    else
		    Root.setAttribute "Confirm","N"
	    end if
	end if
	
	
	''Blocked on Sep 09,2010
'	if trim(document.formname.hCallFrom.value)="DIS" then
'		if document.formname.radInv(0).checked = true then
'			sInvValue = "A"
'		elseif document.formname.radInv(1).checked = true then
'			sInvValue = "P"
'		end if
'		Root.setAttribute "Invoice",sInvValue
'	end if

	if trim(document.formname.hCallFrom.value)="DIS" then
		Root.setAttribute "Invoice","A"
	end if


'''blocked by ragav on apr 28,2011

'	if trim(document.formname.hCallFrom.value)="SUB" then
'		if document.formname.radConfirm(0).checked = true then
'			if confirm("Do you want to create proforma invoice ?") then
'				sProValue="Y"
'			else
'				sProValue="N"
'			end if
'		else
'			sProValue="N"
'		end if
'		Root.setAttribute "ProInv",sProValue
'	end if
	
	
'''added by ragav on apr 28,2011
	if trim(document.formname.hCallFrom.value)="SUB" then
		Root.setAttribute "Confirm","Y"
		if document.formname.radConfirm(0).checked = true then
			sProValue="Y"
		else
			sProValue="N"
		end if
		Root.setAttribute "ProInv",sProValue
	end if 'if trim(document.formname.hCallFrom.value)="SUB" then	
window.close
End Function
'******************************************
Function window_onunload()
	set window.returnvalue = RefData.documentElement
End Function
'==================================================================
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hIssType" value="<%=sIssType%>">
<input type="hidden" name="hPickType" value="<%=sPickType%>">
<input type="hidden" name="hLotOrPickFlag" value="<%=sLotOrPickFlag%>">
<OBJECT id="penDet" type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version"></PARAM>
</OBJECT>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sHeading%>
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
					<td width="10px"></td>
					<TD class=TabBodywithtopline>

						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td width="10">
								</td>
								<td>
									<table border=0 cellspacing=0 cellpadding=0 width="100%">

										<%if trim(sCallFrom)="DIS" then%>
										<tr>
											<td>
												<table border=0 class=ExcelTable cellspacing=1 cellpadding=0 width="80%">
												<!-- Blocked on Sep 09,2010-->
													<!--<tr>
														<td class="FieldCellSub">
															<input type=radio name="radInv" value="A" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Actual Invoice
														</td>
													</tr>-->
													<!--<tr>
														<td class="FieldCellSub">
															<input type=radio name="radInv" value="P" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Proforma Invoice
														</td>
													</tr>-->
												</table>
											<td>
										</tr>
										<%end if 'if trim(sCallFrom)="DIS" then %>
										
										
										<%if trim(sCallFrom)="DIS" then%>
										    <%if (trim(sIssType)="M" and trim(sPickType)="N" and trim(sLotOrPickFlag)="P") or trim(sIssType)="F" then %>
    									    <tr>
											    <td class="FieldCellSub">
												    <input type=radio name="radConfirm" value="Y" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Now
											    </td>
										    </tr>
										    <%end if %>
										    <tr>
											    <td class="FieldCellSub">
												    <input type=radio name="radConfirm" value="N" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Later
											    </td>
										    </tr>
										<%else%>
										    <tr>
											    <td class="FieldCellSub">
												    <input type=radio name="radConfirm" value="Y" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Now
											    </td>
										    </tr>
										    <tr>
											    <td class="FieldCellSub">
												    <input type=radio name="radConfirm" value="N">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Later
											    </td>
										    </tr>
										<%end if%>

									</table>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="ActionCell">
									<input type=button name="btnProceed" value="Proceed" class="ActionButtonX" onClick="FinalSubmit()">
									<!--<input type=button name="btnReset" value="Cancel" class="ActionButtonX">-->
								</td>
							</tr>


                        </table>
					</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
	' Function to populate Usage
	Function populateUsage()
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTDESCRIPTION,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sUnit) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function

	Function Issue()
		MsgBox "ok"
	End Function
%>
