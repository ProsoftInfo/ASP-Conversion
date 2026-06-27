<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartySubTypeMultipleSel.asp 
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MAHESHWARI S
	'Created On					:	Feb 26,2007
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>TDS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/TempItem.js"></SCRIPT>
<!-- XML Data Island -->
 



<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>

<SCRIPT language="vbscript">
Dim objhttp,objTemp
Set objTemp = window.dialogArguments
'*******************************************************************************************
Function init()
	dim objRoot
	Set objRoot = objTemp.documentElement
'	set objRoot = AccHeadData.documentElement
	iCnt = document.formname.hCnt.value
	'alert objRoot.xml
	IF Not  objRoot.haschildnodes then exit function
	If objRoot.haschildnodes then
		
		For each node in objRoot.childnodes
			if trim(node.nodename) = "Details" then
				sChkVal = node.getAttribute("PartyType")&"?"&node.getAttribute("PartySubType")&"?"&node.getAttribute("PartySubTypeName")
					For i = 1 to iCnt 
						If trim(eval("document.formname.ChkType"&i).value)  = trim(sChkVal) then
							eval("document.formname.ChkType"&i).checked = true
						End if
					Next
			Else
				exit for
			End IF		
		 
		Next
	End If
 
	 
End Function 
'*******************************************************************************************
Function CheckSubmit()
Dim i,iCnt,sVal ,iVal,Root,Elem
	Set Root = objTemp.documentElement
	'Set Root = AccHeadData.createElement("AccHead")
	'AccHeadData.appendChild Root
	'alert "b4="&Root.xml
	If Root.haschildnodes then 
		For each node in Root.childnodes 
			if trim(node.NodeName) = "Details"  then
				set remnode = node
			end if
			Root.removechild remnode
		Next	
		
	End IF
	'alert "A4="&Root.xml
	iCnt = document.formname.hCnt.value
	For i = 1 to iCnt 
		If eval("document.formname.ChkType"&i).Checked  = True then
			sVal = eval("document.formname.ChkType"&i).value
			sTemp = split(sVal,"?")
			Set Elem = objTemp.createElement("Details")
			Elem.setattribute "PartyType",sTemp(0)
			Elem.setattribute "PartySubType",sTemp(1)
			Elem.setattribute "PartySubTypeName",sTemp(2)
			Root.Appendchild Elem
			'iVal = iVal &":"& eval("document.formname.ChkType"&i).value
		end if
	Next
	'alert Root.xml
	Win_UnLoad()
End Function

'*******************************************************************************

Function Win_UnLoad()	 
	window.ReturnValue = objTemp.documentElement
	window.close()
End Function
'*******************************************************************************
</script>
</HEAD>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="Init()">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Account Head </td>
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5"  >
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" width="100%">
									 
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
                        <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" >
								</td>
                        </tr>
                        <tr>
							<td align="center" class="ClearPixel">
							</td>
							<td valign="top" class="FieldCell" width="100%" align="center">
                            <DIV class=frmBody id=frm3 style="width: 360; height:380;">
			
								<table border="0" cellspacing="1" class="BodyTable" width="100%" id=TabAccHead>
								<% dim objRs,objRs1,sQuery,iCnt
								dim sOrgId,sCallTy,sParType,iParSubType,sParTypeName

								set objRs = Server.CreateObject("ADODB.Recordset")
								set objRs1 = Server.CreateObject("ADODB.Recordset")
								sOrgId = Request("Unit")
								 
								'Response.Write sUnit
								sCallTy = Request("sCallTy")
								

								IF CStr(sCallTy) <> "P" Then
									sQuery="select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"'"
								Else
									sQuery = "select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"' "&_
											 "and ((PartyType = 'CR' and PartySubType >= 2) or PartyType = 'DR') "
								End IF

									With objRs
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQuery
										.ActiveConnection = con
										.Open
									End with

									Set objRs.Activeconnection = nothing

									if not objRs.EOF then
										 
										 iCnt = 0
										Do while not objRs.EOF
											
											sParType=objRs(0)
											iParSubType = objRs(1)
											sParTypeName = Replace(objRs(2)	,"&"," and ")
											sQuery="select count(1) from APP_R_OrgParty where PartyType='"&sParType&"' and PartySubType="&iParSubType&" and OUDefinitionID='"&sOrgId&"'"
											with objRs1
												.CursorLocation =3
												.CursorType =3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with
											set objRs1.ActiveConnection=nothing
											if CDbl(objRs1(0)) then
											iCnt = iCnt + 1
													%>
													<tr>
													<td class="FieldCellSub" valign="top"></td>
													<td class="FieldCell" valign="top">
													<input type="checkbox" name="ChkType<%=iCnt%>" value="<%=sParType%>?<%=iParSubType%>?<%=sParTypeName%>" class="formelem" >  
													</td>
													<td class="FieldCellSub" valign="top"></td>
													<td class="FieldCell" valign="top"><%=sParTypeName%></td>
													</tr>		
																				
													<%
													
													end if			
													objRs1.Close
												  	objRs.MoveNext		  	
												Loop				
												 
											end if
											objRs.Close
											%>
								<input type=hidden name="hCnt" value="<%=iCnt%>">
								</table>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
									<td valign="middle" class="ActionCell">
                                        <p align="center">
                                        <input type="button" value="Done" name="B2" class="ActionButton" tabindex="3" onclick = "CheckSubmit()">
                                        
									</td>
									</tr>
								</table>
							</div>
							</td>
						</tr>
							
							
					</table>
					</td>
				</tr>
				

					
							
													
						
</table>
</form>
</BODY>
</HTML>
