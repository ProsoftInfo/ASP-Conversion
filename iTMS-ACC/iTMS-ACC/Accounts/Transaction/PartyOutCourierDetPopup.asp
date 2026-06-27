<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartyOutstandingPrevReminder.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	UmaMaheswari S
	'Created On					:	09 April 2011
	'Modified By				:	
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
<!--#include file="../../include/Accpopulate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Reminder Preview</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<XML id="GenReminder"><Root/></XML>
<XML id="OutData"><Root Done=""/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script Language="Vbs">
Dim objTemp,Root
set objTemp = window.dialogArguments
set Root = objTemp.documentElement

Function CloseWindow()
	window.close()
End Function

Function window_onunload()
	set window.returnvalue = OutData.documentElement
	window.close 
End Function

Function Submit()
	Dim sSendBy,sCourCompName,sCourTransID,sCourComAddress
	
	If document.formname.radSendBy(0).checked Then
		sSendBy = document.formname.radSendBy(0).value 
	Elseif document.formname.radSendBy(1).checked Then
		sSendBy = document.formname.radSendBy(1).value 
	End If
		
	sCourCompName = Trim(document.formname.txtCouComName.value)
	sCourTransID = Trim(document.formname.txtCouTransID.value)
	sCourComAddress = Trim(document.formname.txtCouComAddress.value)
	
	sExp = "/Reminder"
	set TempNode = objTemp.selectNodes(sExp)
	
	If TempNode.length > 0 Then
		TempNode.Item(0).Attributes.getNamedItem("SENDBY").Value = sSendBy
		TempNode.Item(0).Attributes.getNamedItem("NAME").Value = sCourCompName 
		TempNode.Item(0).Attributes.getNamedItem("ID").Value = sCourTransID 
		TempNode.Item(0).Attributes.getNamedItem("ADDRESS").Value = sCourComAddress 
	End IF
	
	set objHttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","GenReminderInsert.asp",False
	objHttp.send objTemp.XMLDocument
		
	If objHttp.responseText <> "" Then
		alert(objHttp.responseText)
	Else
		alert("Remonder Generated")
		set sRoot = OutData.documentElement
		sRoot.setAttribute("Done") = "Y"
	End IF
	window.close 
End FUnction
</Script>
<script language="javascript" src="../scripts/ModalReturnCompat.js"></script>
<script language="javascript">
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("OutData");
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">	
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Courier Details
		
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
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
								</td>
								<td valign="top" width="100%">
                                <!--<div class="frmBody" id="frm2" style="width: 640; height:310;">-->
									<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 WIDTH=100% class="ExcelTable">
										<tr>
											<td class="FieldCellSub">To Send By</td>
											<td class="FieldCell">
												<Input type="Radio" name="radSendBy" Value="C">Courier
												<Input type="Radio" name="radSendBy" Value="E">E-Mail
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Courier Company Name</td>
											<td class="FieldCell">
												<Input type="Text" name="txtCouComName" value="" class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Courier TransactionID</td>
											<td class="FieldCell">
												<Input type="Text" name="txtCouTransID" value="" class="FormElem">
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Address</td>
											<td class="FieldCell" colspan="2">
												<Textarea type="Text" name="txtCouComAddress" value="" class="FormElem" cols="40"></Textarea>
											</td>
										</tr>
									</TABLE>
                                 <!--</div>-->
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
													<input type="button" value="Done" class="ActionButton" onclick="Submit()">
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
