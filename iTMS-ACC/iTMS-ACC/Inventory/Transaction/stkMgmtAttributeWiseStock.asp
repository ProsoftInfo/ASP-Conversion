<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtAttributeWiseStock.asp
	'Module Name				:	Inventory (Stock Management Attribute Wise Stock)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 08, 2011
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management - Attribute Wise Stock</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Output/></script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"><ROOT/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin

Function FnInit(sItemCode)
	GetXML()
End Function

Function GetXML()
	ClearTable
	clearXML

	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","AttributesDetXML.asp", false
	objhttp.send
	'alert(objhttp.responseText)
				
	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		DisplayDetails()
	else
		clearXML
	end if

End Function

Function DisplayDetails()
	Dim sRecNumStatus,sItemName,sItemTypeID
	ClearTable
	j = 0
	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement
	
	i = 1
	
	sItemName = ""
	sItemTypeID = ""
	
	For Each ndItem In Root.childNodes
		'sItemName = ndItem.getAttribute("IName")
		iItem = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sorgID = ndItem.getAttribute("Unit")
		
		If sItemTypeID = "" Then
			sItemTypeID = ndItem.getAttribute("ItemTypeID")
		End IF
		
		If ndItem.getAttribute("IName") <> "" Then
			sItemName = sItemName & "," & ndItem.getAttribute("IName")
		End IF
		
		For Each HeaderNode in ndItem.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
				idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
			end if
		Next
			
		If ndItem.NodeName = "AttributeDet" Then
			For Each Node in ndItem.childNodes
				
				nOptValue = Node.getAttribute("OptValue")
				sOptName = Node.getAttribute("OptName")
				
				j = j + 1
				
				set oRow = document.all.tblData.insertRow(j)

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=j
				headerCell.className="ExcelSerial"
				headerCell.align="center"
				headerCell.width = "50"

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=sOptName
				headerCell.className="ExcelDisplayCell"
				headerCell.align="Left"
				
							
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<a href=""#"">" )
				set oText1 = document.createElement("<img name=""btn"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" value="""& nOptValue &""" onClick=""LotDetails(this,'"&sItemTypeID&"')"">")
				oText.appendChild(oText1)
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
				headerCell.width = "100"
			
			Next
		End IF 'If HeaderNode.NodeName = "AttributeDet" Then
			
	next
	
	If sItemName <> "" Then
		idItem.innerHTML = mid(sItemName,2)
	End IF
end Function

Function LotDetails(sOptValue,sItemTypeID)
	set Outvalue = showModalDialog("AttributeWiseLotDetailSelection.asp?Data="&trim(sOptValue.value)&"&sType="&sItemTypeID,OutData,"dialogHeight:310px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	set Root  = OutData.documentElement
	'alert(Root.xml)
End Function


Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			set a=Root.removeChild(HeaderNode)
		next
	end if
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
end Function

Function CheckSubmit()
	
	Set Root = OutData.documentElement
	
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","stkMgmtAttWiseLotDetInsert.asp", false
	objhttp.send OutData.XMLDocument
	
	document.formname.B7.disabled = true
	
	if objhttp.responseText = "" then
		Msgbox ("Attribute Wise Lot Details Allocated")
		window.location.href = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M"
	else
		alert(objhttp.responseText)
	end if

end Function

Function Back()
	window.location.href = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M"
End Function

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp


	set rsTemp = server.CreateObject("ADODB.Recordset")

	sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	'sOrgID = trim(Request.Form("selUnit"))
	sOrgID =session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	if sTempName  = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select ItemDescription from Inv_M_ItemMaster where ItemCode = " & mid(sTemp,1,len(sTemp)-1)
			.Open
		end with

		if not rsTemp.EOF then
			sTempName = trim(rsTemp(0)) & "|"
		end if
		rsTemp.Close
	end if 'if sTempName  = "" then

	if trim(sClassName) = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select GroupName from Inv_M_Classification where GroupCode = " & iClass
			.Open
		end with

		if not rsTemp.EOF then
			sClassName = trim(rsTemp(0))
		end if
		rsTemp.Close
	end if 'if trim(sClassName) = "" then
	
	'Response.Write "<p><font color=red>sTempName="&sTempName & "-"& sTemp 
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")

%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Attribute Wise Lot Selection
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Classification</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idClass"><%=sClassName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly" id="idUoM"></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idItem"></span>
                                            </td>
                                        </tr>
                                        
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 580; height:250;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center">Attribute Name</td>
                                                <td class="ExcelHeaderCell" align="center"></td>
                                            </tr>
                                            
                                        </table>
                                    </div>
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
                                                <input type="button" value="Done" name="B7" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="button" value="Back" name="B8" class="ActionButton" onClick="Back()">
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