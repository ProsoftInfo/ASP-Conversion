<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedBrkSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 23 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	SchBreakupSetUp.asp
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
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera
	Dim ShSubID,ShSubSubID,sInsDate
	Set objrs1 = Server.createObject("ADODB.Recordset")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	sInsDate = Request("InsDate")
	
	IF CStr(sSchedNo) = "" Then
		sSchedNo = Request("sschedno")
	End IF
	If ShSubID = "" Then ShSubID = 0
	If ShSubSubID ="" Then ShSubSubID = 0
	
	sql = "select scheduleheading from dbo.acc_m_schdsetupheads where scheduleID ='"&sSchedNo&"'"
	with objrs1
		.CursorLocation = 3
		.CursorType =3
		.ActiveConnection = con
		.Source = sql
		.Open 
	end with
	if not Objrs1.EOF then sHead = objrs1(0)
	Objrs1.Close
		
%>
<XML ID="XmlData">
<Root>
	<Details OrgID="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="">
	</Details>
</Root>
</XML>
<XML id="OutData"><Root/></xml>

<XML ID="TempData">
	<Root/>
</XML>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Function SetLevelFun()
	Dim ArrSplit,objhttp,sShID,sOrgID,sFinyr,id
	Dim sShSubID,sShSubSubID,Root,HeadNode
	If document.formname.sel1.value = "0" Then
		MsgBox("Select Head Name And Proceed...")
	End If  
	If document.formname.sel1.value = "A" or document.formname.sel1.selectedIndex = 0 Then  
		document.formname.txtLev1.disabled = False
	Else
		document.formname.txtLev1.disabled = True 
	End If
	'msgbox("Ok")
	if document.formname.optLevel(0).checked <> True Then 
	document.formname.sel2.length = 0
	
	document.formname.sel2.length = document.formname.sel2.length + 1
	document.formname.sel2.options(document.formname.sel2.length-1).text = "Add New"
	document.formname.sel2.options(document.formname.sel2.length-1).Value = "A"
	document.formname.txtLev2.disabled = False  
	If document.formname.sel1.value <> "A" and document.formname.sel1.selectedIndex <> 0 Then 
	'If document.formname.sel2.disabled = False Then  
		ArrSplit = Split(document.formname.sel1.value,"-") 
		'msgbox document.formname.sel1.value   
		document.formname.ShSubID.value = ArrSplit(0)
		document.formname.ShSubSubID.value = ArrSplit(1)    
	'End If
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	sShID = document.formname.sschedno.Value  
	sOrgID= document.formname.sUnit.Value 
	sFinyr = document.formname.sfinyr.value 
	sShSubID = ArrSplit(0)
	sShSubSubID = ArrSplit(1)
	id = "0"
	objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sShID="&sShID&"&sFinyr="&sFinyr&"&sShSubID="&sShSubID&"&sShSubSubID="&sShSubSubID&"&id="&id, false
	objhttp.send
	
	'alert objhttp.responseText 
	IF objhttp.responsexml.XML <> "" Then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentELement
	End IF
		
	Set Root = OutData.documentElement
	
	Set HeadNode = Root.selectNodes("//Root")
	For Each HeadNode in Root.ChildNodes
		document.formname.sel2.length = document.formname.sel2.length + 1
		document.formname.sel2.options(document.formname.sel2.length-1).text = HeadNode.Attributes.Item(0).nodeValue
		document.formname.sel2.options(document.formname.sel2.length-1).Value = HeadNode.Attributes.Item(1).nodeValue
	Next
	End If
	End If
End Function

Function setlevelfun1()
	Dim objhttp,sOrgID,sFinyr,id
	Dim sBreakID,Root,HeadNode,recno,sExp,Splt
	Dim Att1,Att2,Att3,Att4,Att,Att5,Att6,Att7,Att8
	If document.formname.sel2.value = "A" Then  
		document.formname.txtLev2.disabled = False  
	Else
		document.formname.txtLev2.disabled = False
		document.formname.txtLev2.value = document.formname.sel2.options(document.formname.sel2.selectedIndex).text
		Splt = Split(document.formname.sel2.value,",")     
	End If
		'msgbox document.formname.sel2.value   
		document.formname.txtLev3.disabled= False     
		document.formname.sel3.length = 0
		document.formname.sel3.length = document.formname.sel3.length + 1
		document.formname.sel3.options(document.formname.sel3.length-1).text = "Add New"
		document.formname.sel3.options(document.formname.sel3.length-1).Value = "A"
		
	If document.formname.sel3.disabled = False And document.formname.sel2.value <> "A" Then 
		sBreakID = Splt(0) 
		document.formname.txtHierarchy.value = Splt(1)   
		'MsgBox sBreakID 
		sOrgID= document.formname.sUnit.Value 
		sFinyr = document.formname.sfinyr.value  
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		id = "1"
		'MsgBox sBreakID
		objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&sBreakID="&sBreakID&"&id="&id, false
		objhttp.send
		'msgbox objhttp.responsexml.XML
		IF objhttp.responsexml.XML <> "Root/" Then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentELement
		End IF

		Set Root = OutData.documentElement
		'alert Root.xml
		
		sExp = "//Root"
		Set HeadNode = Root.selectNodes(sExp)
		For Each HeadNode in Root.ChildNodes
			document.formname.sel3.length = document.formname.sel3.length + 1
			document.formname.sel3.options(document.formname.sel3.length-1).text = HeadNode.Attributes.Item(0).nodeValue
			Att2 = HeadNode.Attributes.Item(2).nodeValue
			Att1 = HeadNode.Attributes.Item(1).nodeValue
			Att3 = HeadNode.Attributes.Item(3).nodeValue
			Att4 = HeadNode.Attributes.Item(4).nodeValue
			Att5 = HeadNode.Attributes.Item(5).nodeValue
			Att6 = HeadNode.Attributes.Item(6).nodeValue
			Att7 = HeadNode.Attributes.Item(7).nodeValue 
			Att8 = HeadNode.Attributes.Item(8).nodeValue 
			Att = Att1&"-"&Att2&"-"&Att3&"-"&Att4&"-"&Att5&"-"&Att6&"-"&Att7&"-"&Att8  
			document.formname.sel3.options(document.formname.sel3.length-1).Value = Att 
		'	msgbox att
		Next
	End If	
		
End Function	
	
	Function populateSubID()
		MsgBox "1"
	End Function
	
'================================================================================================
	'Function SetLevelFun()
	'			If document.formname.sel1.value = "A" Then  
	'				document.formname.txtLev1.disabled = false
	'			Else
	'				document.formname.txtLev1.disabled = True 
	'			End if
	'End Function

	Function LevelFun()
		If document.formname.optLevel(0).checked = True Then  
			document.formname.txtLev2.value = ""  
			document.formname.txtAcHead.value = ""
			document.formname.sel2.disabled = True
			document.formname.FinYear.disabled = True
			document.formname.txtLev3.value = ""
			document.formname.sel3.disabled = True  
			document.formname.FinYear.disabled = True   
		ElseIf document.formname.optLevel(1).checked = True  Then  
			document.formname.sel2.disabled = False
			document.formname.FinYear.disabled = True
			document.formname.txtLev3.value= ""
			document.formname.sel3.disabled = True  
			document.formname.txtLev2.value = ""  
		Else
			document.formname.txtLev2.value = ""  
			document.formname.txtLev3.value = ""  
			document.formname.sel2.disabled = False
			document.formname.FinYear.disabled = True
			document.formname.sel3.disabled = False
		End If
	End Function
	
	Function ModeFun()
			If document.formname.optMode(1).checked = True Then
				document.formname.ButAcHead.disabled = False
				document.formname.FinYear.disabled = False  
			Else
				document.formname.ButAcHead.disabled = True  
				document.formname.txtAcHead.value = ""    
				document.formname.FinYear.disabled = True 
			End If
	End Function
	
	Function AccHeadClck()
			'document.formname.txtAcHead.disabled = False 
	End Function
	
	Function popAccList()
		Dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue,sAccHeadName,sNewAcc,sOrgId
		sOrgId = document.formname.sUnit.value 
		OutValue= showModalDialog("ChgAccHeadName.asp?orgId="+sOrgId,"","dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
		While UBound(arrTemp) = 0 
			OutValue = showModalDialog("ChgAccHeadName.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		Wend
		sRetVal = OutValue		
		If UBound(arrTemp) <= 1 Then Exit Function
		document.formname.hAccHead.value = arrTemp(0)
		sAccHeadName = arrTemp(1)
		document.formname.txtAcHead.value = sAccHeadName
		'MsgBox document.formname.hAccHead.value 
	End Function 
	
	Function Setlevelfun2()
		Dim BreakUpID,BreakUpSubID,BreakUpSubSubID,isplit
				document.formname.ButAcHead.disabled= True
				document.formname.hAccHead.value = ""
				document.formname.txtAcHead.value = ""
		If document.formname.sel3.value = "A" Then
			document.formname.txtLev3.disabled= False  
		Else
			document.formname.txtLev3.disabled = False
			document.formname.txtLev3.value = document.formname.sel3.options(document.formname.sel3.selectedIndex).text         			  
			'MsgBox document.formname.sel3.value   
			isplit = Split(document.formname.sel3.value,"-")
			'msgbox document.formname.sel3.value   
			If isplit(2) = "Y" Then
				document.formname.optMode(1).checked = True
				document.formname.ButAcHead.disabled= False
				document.formname.hAccHead.value = isplit(4)
				document.formname.txtAcHead.value = isplit(5)  
				document.formname.FinYear.disabled = False        
			ElseIf isplit(3) <> "N" Then 
				 document.formname.optMode(0).checked = True  
			Else
				document.formname.optMode(2).checked = True  
			End If
			If isplit(6) = "+" or isplit(6)="++" Then
			If isplit(6) = "+" Then 
				document.formname.optCompMode(0).checked = True  
			Else
				document.formname.optCompMode(1).checked = True  
			End If
			End If
			document.formname.txtHierarchy.value = isplit(7)    
		End If
	End Function

Function CheckSubmit()
Dim Root,objhttp,sTemp,Node1
Dim sOrgId,sFinYr,sShID,sShSubID,sShSubSubID,sBreakupid,sBreakupSubID,ID
Dim sBreakSubSubID,sMode,sComputeMode,sAccHeadID,sAccHeadName
Dim sHeadName,sBreakHeadName,sBreakSubHeadName
Dim sSubIDMerged,sBreakIDMerged,sSubIDSplit,sBreakIDSplit
	sShID =0:sShSubID = 0:sShSubSubID = 0:sBreakupid = 0:sBreakupSubID = 0
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement
	
	sOrgId = document.formname.sUnit.value
	sFinYr = document.formname.sfinyr.value
	sShID = document.formname.sschedno.value
	
	If document.formname.optLevel(0).checked = True Then
		ID="1"   
	ElseIf document.formname.optLevel(1).checked = True Then
		ID="2"
	Else
		ID="3"
	End IF   
	If ID = "3" and document.formname.sel3.value="0" Then
		Msgbox "Select Level3 and Proceed..!"
		Exit Function   
	End If

	If ID = "2" And document.formname.sel2.value <>"A" Then  
		sHeadName = document.formname.txtLev1.value   
		sBreakHeadName = document.formname.txtLev2.value 
		sSubIDMerged = document.formname.sel1.value 
		sSubIDSplit = Split(sSubIDMerged,"-") 			  			
		sShSubID = sSubIDSplit(0)
		sShSubSubID= sSubIDSplit(1) 

	ElseIf ID = "3" And document.formname.sel3.value <>"A" Then   
		sBreakupid  = document.formname.sel2.value   
		sSubIDMerged = document.formname.sel3.value 
		sSubIDSplit = Split(sSubIDMerged,"-") 			  			
		sBreakupSubID = sSubIDSplit(0)
		sBreakSubSubID = sSubIDSplit(1) 
		sBreakHeadName = document.formname.txtLev2.value   		
		sBreakSubHeadName = document.formname.txtLev3.value 
		sSubIDMerged = document.formname.sel1.value 
		sSubIDSplit = Split(sSubIDMerged,"-") 			  			
		sShSubID = sSubIDSplit(0)
		sShSubSubID= sSubIDSplit(1) 
	Else 
		If ID="1" Then
			sHeadName = document.formname.txtLev1.value   
		End If
		If ID="2" Then
			IF document.formname.sel1.value = "A" Then
				sHeadName = document.formname.txtLev1.value   
				sBreakHeadName = document.formname.txtLev2.value 
			Else
				sBreakHeadName = document.formname.txtLev2.value 
				sSubIDMerged = document.formname.sel1.value 
				sSubIDSplit = Split(sSubIDMerged,"-") 			  			
				sShSubID = sSubIDSplit(0)
				sShSubSubID= sSubIDSplit(1) 
			End If
		End If
		If ID="3" Then
			If  document.formname.sel2.value = "A" Then
				sBreakHeadName = document.formname.txtLev2.value   		
				sBreakSubHeadName = document.formname.txtLev3.value 
				sSubIDMerged = document.formname.sel1.value 
				sSubIDSplit = Split(sSubIDMerged,"-") 			  			
				sShSubID = sSubIDSplit(0)
				sShSubSubID= sSubIDSplit(1) 
				msgbox(sBreakSubHeadName) 
			Else
				sBreakupid = document.formname.sel2.value  
				sBreakSubHeadName = document.formname.txtLev3.value
			End If	
		End If	
	End If


	If document.formname.optMode(1).checked = True Then  
		sAccHeadID = document.formname.hAccHead.value 
		sMode = document.formname.optMode(1).value 
	Else
		sMode ="D"
	End If
	If document.formname.optCompMode(0).checked = True Then
	  sComputeMode = "+"
	Else  
	  sComputeMode = "++"
	End If
	set Node1 = TempData.CreateElement("Schedule")
	
 	Node1.setAttribute "ID",ID 
 	Node1.setAttribute "Level2ID",document.formname.sel2.value 
 	Node1.setAttribute "Level3ID",document.formname.sel3.value   
 	Node1.setAttribute "OrgID",sOrgId 
	Node1.setAttribute "ScheduleID",sShID 
	Node1.setAttribute "ScheduleSubID",sShSubID 
	Node1.setAttribute "ScheduleSubSubID",sShSubSubID 
	Node1.setAttribute "HeadName",sHeadName 
	Node1.setAttribute "BreakUpHeadName",sBreakHeadName 
	Node1.setAttribute "BreakUpSubHead",sBreakSubHeadName 
	Node1.setAttribute "BreakupId",sBreakupid 
	Node1.setAttribute "BreakupSubId",sBreakupSubID  
	Node1.setAttribute "BreakupSubSubId",sBreakSubSubID   
	Node1.setAttribute "Mode",sMode 
	Node1.setAttribute "FinYear",sFinYr
	Node1.setAttribute "ComputeMode",sComputeMode  
	Node1.setAttribute "AccountHeadID",sAccHeadID 
	Node1.setAttribute "Hierarchy",document.formname.txtHierarchy.value   
	Node1.setAttribute "InsDate",document.formname.hInsDate.value
	Root.Appendchild Node1
			
	'alert(Root.xml)
	objhttp.Open "Post","XMLSchBrkHeadSave.asp?Name=SchedBrkSubHeads&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	
		If objhttp.responseText <> "" then
			alert(objhttp.responseText)
		Else
			window.returnvalue = "Y"
			window.close()
		End if
		
End Function


Function Del()
Dim objhttp,Node1
Dim sShID,iBreakID,iBreakSubID,iBreakSubSubID,AcCode
Dim iArr1,iArr2,id,sOrgID,sFinyr,LevelID

	sOrgID = document.formname.sUnit.value
	sFinyr = document.formname.sfinyr.value
	sShID = document.formname.sschedno.value
	
If document.formname.optLevel(1).checked = True Then  
	LevelID = 1
ElseIf document.formname.optLevel(2).checked = True Then
	LevelID = 2
End If  
iBreakID = 0
iBreakSubID = 0
iBreakSubSubID = 0
If LevelID = 1 Then
	If document.formname.sel2.value <>"A" Then
	  iBreakID = document.formname.sel2.value   
	End If
Else
	If document.formname.sel3.value <>"A" Then 
		iArr1 = split(document.formname.sel3.value,"-") 
		iBreakID = document.formname.sel2.value 
		iBreakSubID = iArr1(0)
		iBreakSubSubID = iArr1(1)   
	End If
End If
	AcCode = document.formname.hAccHead.value    
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement
	id = "2"
	set Node1 = TempData.CreateElement("Schedule")
		Node1.setAttribute "LevelID",LevelID 
	 	Node1.setAttribute "iBreakID",iBreakID 
	 	Node1.setAttribute "iBreakSubID",iBreakSubID 
	 	Node1.setAttribute "iBreakSubSubID",iBreakSubSubID 
		Node1.setAttribute "sOrgID",sOrgID
		Node1.setAttribute "sFinyr",sFinyr
		Node1.setAttribute "sShID",sShID
		Node1.setAttribute "AcCode",AcCode		
		Node1.setAttribute "id",id
	Root.Appendchild Node1
	'MsgBox TempData.XML
	
	objhttp.Open "Post","XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", false
	objhttp.send TempData.XML
	If objhttp.responseText <> "" then
		alert(objhttp.responseText)
	Else
		window.returnvalue = "Y"
		window.close()
	End IF	

End Function

</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "scheduleBreakupSubHeadsPopup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" >
<input type=hidden name="sUnit" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sFinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hAccHead" value="0">
<input type=hidden name="ShSubID" value="0">
<input type=hidden name="ShSubSubID" value="0">
<input type=hidden name="hInsDate" value="<%=sInsDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Schedule SubHeads
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="685" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr> <td class="FieldCellSub" width="120">Schedule Name<space><space><space></td>
							<td class="FieldCellSub" colspan="2"><%=sHead%></td>
							</tr>
                            	<td class="FieldCellSub" width="120">Add/Modify Level</td>
										<td class="FieldCell" colspan="2">
											<Input type="Hidden" name="optLevel" value="Lev1" class="FormElem" onclick="LevelFun()">
											<Input type="radio" name="optLevel" value="Lev2" class="FormElem" Checked onload="LevelFun()" onclick="LevelFun()">Level1
											<Input type="radio" name="optLevel" value="Lev3" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="120">Select Level 1</td>
										<td class="FieldCell">
											<select size="1" name="sel1" class="FormElem" onChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select SubHeadingName,ScheduleSubID,ScheduleSubSubID From Vw_Acc_SchSetup "&_
												  "Where EntryType = 'S' and OrganisationCode = '"&sOrgId&"' and FinYear = '"&sFinyr&"' and ScheduleID = "&sSchedNo &_ 
												  "Order By SubHeadingName"
										'Response.Write sql
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF
												
											%>
											<option value="<%=Objrs1(1)&"-"&Objrs1(2)%>"><%Response.Write (Left(Trim(Objrs1("SubHeadingName")),50)) %></b></option>
											<%Objrs1.MoveNext 
											loop
											Objrs1.Close
											%>
											</select>
										</td>
										<td class="FieldCell">
                                            <input type="Hidden" name=txtLev1 size=25 class="Formelem" align="Right" disabled>
										</td>
									</tr>
								  <tr>
								<td class="FieldCellSub" width="120">Select Level 2</td>
										<td class="FieldCell">
										<select size="1" name="sel2" class="FormElem" onchange="setlevelfun1()" >
										<OPTION Value="0">Select</option>
						
									</select>
								</td>
									<input type=hidden name=hSubHeadName value="">
										<td class="FieldCell">
                                        <input type=text name=txtLev2 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Select Level 3</td>
										<td class="FieldCell">
										<select size="1" name="sel3" class="FormElem" onchange="Setlevelfun2()" disabled>
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>
									
									</td>
									<input type=hidden name=hSubHeadName3 value="">
										<td class="FieldCell">
                                        <input type=text name=txtLev3 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								
								
								
								<tr>
								<td class="FieldCellSub" width="120">Mode</td>
								<td class="FieldCell" colspan="2">
									<Input type=radio name=optMode value="D" class="FormElem" checked OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A2" class="FormElem" OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Select A/c Head</td>
								<td class="FieldCell">
								<Input type="Button" name="ButAcHead" value="A/cHead" class="ActionButton" disabled OnClick="popAccList()" >
                                &nbsp;</td>
								<td class="FieldCell">
								<Input type="text" name="txtAcHead" size="25" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">FinYear</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="FinYear" class="FormElem" disabled>
											<Option>Select</Option>
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option> 
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Compute Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>
								<td class="FieldCell">Hierarchy &nbsp;&nbsp; 
								<Input type="text" name="txtHierarchy" size="5" class="formElem" value="" ></td>
								</tr>
							
                            <td colspan="3">
								
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> 
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "CheckSubmit()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">&nbsp;
                                                <Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick ="Del()">
											</td>
										</tr>
									</table>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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

