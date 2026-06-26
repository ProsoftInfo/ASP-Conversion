<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedBSSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 29 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	BSSetUp.asp
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
<HTML><HEAD><title>Add B/S Heads</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	Dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName,Objrs,sInsDate
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera
	Set objrs1 = Server.createObject("ADODB.Recordset")
	Set Objrs = Server.createObject("ADODB.Recordset")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	sInsDate = Request("InsDate")
	
%>
<XML ID="XmlData">
<Root>
	<Details OrgID="" SchName="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="" Hierachy="" InsDate="<%=sInsDate%>">
	</Details>
</Root>
</XML>
<XML id="OutData"><Root/></xml>
<XML ID="TempData"><Root/></XML>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Function SelHead()
	Dim sOrgID,EntryType,sFinyr,sschedno,id
	Dim objhttp,HeadNode
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		sOrgID = document.formname.hOrgId.value 
		sFinyr = document.formname.sfinyr.value
		sschedno = document.formname.selSch.value 
	If document.formname.selSch.value = "A" Then
		document.formname.txtLev.disabled = False
		document.formname.txtLev.value = ""  
		document.formname.selLevel1.length=1
		document.formname.selLevel1.options(0).Value = "A"
		document.formname.selLevel1.options(0).text ="AddNew"   
		document.formname.txtLev1.disabled = False  
	Else
		document.formname.selLevel1.length = 0
		document.formname.txtLev.value = document.formname.selSch.options(document.formname.selSch.selectedIndex).text
	'	MsgBox document.formname.txtLev.value   
		document.formname.selLevel1.length = document.formname.selLevel1.length + 1
		document.formname.selLevel1.options(document.formname.selLevel1.length-1).text = "Add New"
		document.formname.selLevel1.options(document.formname.selLevel1.length-1).Value = "A"
		id = "8"
	'	MSGBOX sOrgID&sschedno
		objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&id="&id, false
		objhttp.send
	'	msgbox objhttp.responseTEXT
		IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				recno = HeadNode.length-1
				'msgbox recno 
				For Each HeadNode in Root.ChildNodes
					document.formname.selLevel1.length = document.formname.selLevel1.length + 1
					document.formname.selLevel1.options(document.formname.selLevel1.length-1).text = HeadNode.Attributes.Item(1).nodeValue
					document.formname.selLevel1.options(document.formname.selLevel1.length-1).Value = HeadNode.Attributes.Item(0).nodeValue
				Next
		End IF
	'	document.formname.txtLev.disabled = True
	End If 
End Function

Function SetLevelFun()
		Dim sOrgID,EntryType,sFinyr,sschedno,id,splt1,splt2
		Dim objhttp,objhttp2,HeadNode,SubSplt,SubID,Value
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		set objhttp2 = CreateObject("MSXML2.XMLHTTP")
		sOrgID = document.formname.hOrgId.value 
		sFinyr = document.formname.sfinyr.value
		sschedno = document.formname.selSch.value 		
		SubSplt = Split(document.formname.selLevel1.value,",")
		SubID = SubSplt(0)
		document.formname.txtLev1.disabled = False  
		document.formname.txtLev2.disabled = False  
		'MsgBox document.formname.selLevel1.value    
	If document.formname.selLevel1.value = "A" Then  
		document.formname.selLevel2.length=1
		document.formname.selLevel2.options(0).Value = "A"
		document.formname.selLevel2.options(0).text ="AddNew"   
	Else
		Value = Split(document.formname.selLevel1.value,",")
		document.formname.txtHierarchy.value = Value(5)  
		If Value(2) = "A" Then
			document.formname.optMode(1).checked = True  
			document.formname.ButAcHead.disabled = False  
			document.formname.hAccHead.value = Value(4)
			id = "7"
			'MsgBox document.formname.selLevel1.value
			'msgbox sOrgID&sFinyr&Value(4)&id  
			objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&AcCode="&Value(4)&"&id="&id, false
			objhttp.send
			IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				recno = HeadNode.length-1
				'msgbox recno 
				For Each HeadNode in Root.ChildNodes
					document.formname.hAccHead.value = HeadNode.Attributes.Item(0).nodeValue
					document.formname.txtAcHead.value = HeadNode.Attributes.Item(1).nodeValue
				Next
			End IF
			
		ElseIf Value(2) = "S" Then
			document.formname.optMode(2).checked = True  
			document.formname.ButAcHead.disabled = False  
			document.formname.hAccHead.value = Value(4)
			id = "9"
			'msgbox document.formname.selLevel1.value   
			splt2 = Split(document.formname.selLevel1.value,",")    
	'		splt1 = Split(document.formname.selLevel2.value,",")   
'			msgbox sOrgID&sFinyr&sschedno&"'"&splt2(0)&"'"&splt1(1)&"'"&id   
			objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&splt2(0)&"&SubSubID="&splt2(1)&"&id="&id, false
			objhttp.send
			'msgbox objhttp.responsexml.XML
			
			IF objhttp.responsexml.XML <> "" Then
			objhttp2.Open "Post","XMLSave.asp?Name=SchedBSBrkSubHeads&Mod=Acc", false
			'MsgBox objhttp.responsexml.XML
			objhttp2.send objhttp.responsexml.XML
			End IF
		ElseIf Value(2) = "D" Then
			document.formname.optMode(0).checked = True  
		Else
			document.formname.optMode(3).checked = True
		End If   
		If Value(3) = "+" then
			document.formname.optCompMode(0).checked = true
		ElseIf Value(3) = "-" Then
			document.formname.optCompMode(1).checked = True
		End If
		
		document.formname.selLevel2.length = 0
		document.formname.txtLev1.value = document.formname.selLevel1.options(document.formname.selLevel1.selectedIndex).text        
		document.formname.selLevel2.length = document.formname.selLevel2.length + 1
		document.formname.selLevel2.options(document.formname.selLevel2.length-1).text = "Add New"
		document.formname.selLevel2.options(document.formname.selLevel2.length-1).Value = "A"		
		id = "10"
		objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&SubID&"&id="&id, false
		objhttp.send
		'msgbox objhttp.responseTEXT
		IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				recno = HeadNode.length-1
				For Each HeadNode in Root.ChildNodes
					document.formname.selLevel2.length = document.formname.selLevel2.length + 1
					document.formname.selLevel2.options(document.formname.selLevel2.length-1).text = HeadNode.Attributes.Item(1).nodeValue
					document.formname.selLevel2.options(document.formname.selLevel2.length-1).Value = HeadNode.Attributes.Item(0).nodeValue
					'MsgBox HeadNode.Attributes.Item(0).nodeValue
				Next
		End IF
		'document.formname.selLevel2.disabled = True
	End If
End Function
	
Function LevelFun()
	If document.formname.optLevel(0).checked = True then  
		document.formname.txtLev2.value = ""  
		document.formname.txtLev2.disabled = True  
		document.formname.txtAcHead.value = ""
		document.formname.txtAcHead.disabled = True   
		document.formname.selLevel2.disabled = true
		document.formname.FinYear.disabled = true
	Else
		document.formname.txtLev2.disabled = False
		document.formname.txtAcHead.disabled = False   
		document.formname.selLevel2.disabled = False
		document.formname.FinYear.disabled = False
	End If
End Function
	
Function setlevelfun1()
	Dim Value,id,sOrgID,sFinyr,sschedno,SubID,SubSubID,splt1,splt2,objhttp2,AcCode
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	set objhttp2 = CreateObject("MSXML2.XMLHTTP")
	sOrgID = document.formname.hOrgId.value 
	sFinyr = document.formname.sfinyr.value
	sschedno = document.formname.selSch.value 
	document.formname.txtLev2.disabled = false  
	If document.formname.selLevel2.value<>"A" Then   
		document.formname.txtLev2.value = document.formname.selLevel2.options(document.formname.selLevel2.selectedIndex).text        
		Value = Split(document.formname.selLevel2.value,",")
	'	MsgBox document.formname.selLevel2.value    
		document.formname.txtHierarchy.value = Value(5)
		If Value(2) = "A" Then
			document.formname.optMode(1).checked = True  
			document.formname.ButAcHead.disabled = False  
			document.formname.hAccHead.value = Value(4)
			id = "7"
			objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&AcCode="&Value(4)&"&id="&id, false
			objhttp.send
			IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				recno = HeadNode.length-1
				'msgbox recno 
				For Each HeadNode in Root.ChildNodes
					document.formname.hAccHead.value = HeadNode.Attributes.Item(0).nodeValue
					document.formname.txtAcHead.value = HeadNode.Attributes.Item(1).nodeValue					
				Next
			End IF
			
		ElseIf Value(2) = "S" Then
			document.formname.optMode(2).checked = True  
			document.formname.ButAcHead.disabled = False  
			document.formname.hAccHead.value = Value(4)
			id = "9"
			splt2 = Split(document.formname.selLevel1.value,",")    
			splt1 = Split(document.formname.selLevel2.value,",")   
'			msgbox sOrgID&sFinyr&sschedno&"'"&splt2(0)&"'"&splt1(1)&"'"&id   
			objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&splt1(0)&"&SubSubID="&splt1(1)&"&id="&id, false
			objhttp.send
			'msgbox objhttp.responsexml.XML
			
			IF objhttp.responsexml.XML <> "" Then
			objhttp2.Open "Post","XMLSave.asp?Name=SchedBSBrkSubHeads&Mod=Acc", false
			'MsgBox objhttp.responsexml.XML
			objhttp2.send objhttp.responsexml.XML
			End IF
			 
		ElseIf Value(2) = "D" Then
			document.formname.optMode(0).checked = True  
		Else
			document.formname.optMode(3).checked = True
		End If   
		If Value(3) = "+" then
			document.formname.optCompMode(0).checked = true
		ElseIf Value(3) = "-" Then
			document.formname.optCompMode(1).checked = True
		End If
	End If
'	MsgBox document.formname.selLevel2.value   
End Function
	
Function ModeFun()
	If document.formname.optMode(1).checked = true or document.formname.optMode(2).checked = True Then
		document.formname.ButAcHead.disabled = False
	Else
		document.formname.ButAcHead.disabled = True  
		document.formname.txtAcHead.value = ""    
	End If
End Function

Function AccHeadClck()
	document.formname.txtAcHead.disabled = false 
End Function
	
Function popAccList()
	Dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue,sAccHeadName,sNewAcc,sOrgId
	Dim iBSHead,iBSSubHd,iBSSubSubHd
	iBSHead = document.formname.selSch.value
	iBSSubHd = document.formname.selLevel1.value
	IF document.formname.optLevel(1).checked = True Then
		iBSSubSubHd = document.formname.selLevel2.value
	Else
		iBSSubSubHd = 0
	End IF
	sOrgId = document.formname.hOrgId.value 
	
	If document.formname.optMode(1).checked = True Then
		OutValue= showModalDialog("ChgAccHeadName.asp?orgId="+sOrgId,"","dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
				
		while UBound(arrTemp) = 0 
			OutValue = showModalDialog("ChgAccHeadName.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		wend
		sRetVal = OutValue
		'MsgBox nEntNo
		'msgbox OutValue		
		IF UBound(arrTemp) <= 1 then exit function
		document.formname.hAccHead.value = arrTemp(0)
		sAccHeadName = arrTemp(1)
		document.formname.txtAcHead.value = sAccHeadName
		'MsgBox document.formname.hAccHead.value 
	ElseIf	document.formname.optMode(2).checked = True Then  
		'MsgBox "1"
		OutValue= showModalDialog("SelAccBSHeadName.asp?orgId="+sOrgId&"&BSHead="&iBSHead&"&BSSubHd="&iBSSubHd&"&BSSubSubHd="&iBSSubSubHd,"A","dialogHeight:460px;dialogWidth:550px;center:Yes;help:No;resizable:Yes;status:No")
		'alert(OutValue)
	End If	
End Function 

Function SaveXML()
	Dim sExp,TempNode,recno
	If document.formname.selSch.value = "A" Then
		If document.formname.txtLev.value = "" or document.formname.txtLev1.value ="" Then   
			Msgbox "Enter All The Information And Try To Save...!"
			Exit Function
		End If
	End IF     
	
	IF CheckVal() Then
		Set Root = XmlData.documentElement
		Set objhttp = CreateObject("Microsoft.XMLHTTP")
	
		sExp = "//Details"
		Set TempNode = Root.selectNodes(sExp)
		recno = TempNode.length-1
		If document.formname.selSch.value = "A" Then
			If Trim(document.formname.txtLev.value) = "" Then MsgBox "Enter Schedule And Continue...!" : Exit Function
		End If
		TempNode.Item(recno).Attributes.getNamedItem("SchName").value = document.formname.txtLev.value 
		TempNode.Item(recno).Attributes.getNamedItem("OrgID").value = document.formname.hOrgId.Value
		TempNode.Item(recno).Attributes.getNamedItem("SchID").value = document.formname.selSch.value
	
		if document.formname.optLevel(0).checked = true then  
			TempNode.Item(recno).Attributes.getNamedItem("LevelID").value = 0 
		else
			TempNode.Item(recno).Attributes.getNamedItem("LevelID").value = 1
		end if
		
		TempNode.Item(recno).Attributes.getNamedItem("Level1ID").value = document.formname.selLevel1.value 
		TempNode.Item(recno).Attributes.getNamedItem("Level2ID").value = document.formname.selLevel2.value    
		TempNode.Item(recno).Attributes.getNamedItem("Level1Name").Value = document.formname.txtLev1.value 
		TempNode.Item(recno).Attributes.getNamedItem("Level2Name").value = document.formname.txtLev2.value 
	
		IF document.formname.optMode(0).checked = true then
			TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "D"
		elseif document.formname.optMode(1).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "A" 
		elseif document.formname.optMode(2).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "S" 
		elseif document.formname.optMode(3).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "N" 
		end if
	
		TempNode.Item(recno).Attributes.getNamedItem("AccHead").value = document.formname.hAccHead.value       
	
		IF TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "A" then
			TempNode.Item(recno).Attributes.getNamedItem("AccHeadName").value = document.formname.txtAcHead.value       
		end if
			TempNode.Item(recno).Attributes.getNamedItem("FinYear").value = document.formname.FinYear.value 
		IF document.formname.optCompMode(0).checked = true then
			   TempNode.Item(recno).Attributes.getNamedItem("ComputeMode").value = "+"
		else
			   TempNode.Item(recno).Attributes.getNamedItem("ComputeMode").value = "-"
		end If
			'	MsgBox document.formname.txtHierarchy.value   
			   TempNode.Item(recno).Attributes.getNamedItem("Hierachy").value = document.formname.txtHierarchy.value   
		'MsgBox XMLData.XML
		objhttp.Open "Post","XMLSchBSHeadSave.asp?Name=SchedBSSubHeads&Mod=Acc", false
		objhttp.send XMLData.XMLDocument
		
		If objhttp.responseText <> "" then
			alert(objhttp.responseText)
		Else
			window.returnvalue = "Y"
			window.close()
		End IF
	Else
		Exit Function
	End IF
End Function

Function Del()
	Dim objhttp,Node1
	Dim sschedno,SubID,SubSubID,iArr1,iArr2,id,sOrgID,sFinyr
	SubID = 0
	SubSubID = 0
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement
	sOrgID = document.formname.hOrgId.Value
	sFinyr = document.formname.sfinyr.value  
	
	If document.formname.selSch.value = "A" Then  
		Msgbox "Select Any Value And Then Delete...!"
	End If
	
	sschedno = document.formname.selSch.value   
	If  document.formname.optLevel(0).checked = True and document.formname.selLevel1.value <>"A" Then  
		iArr1 = Split(document.formname.selLevel1.value,",")
		SubID = iArr1(0)    
	End If
	If document.formname.optLevel(1).checked = True and document.formname.selLevel2.value <>"A" Then    
		iArr1 = Split(document.formname.selLevel1.value,",")   
		SubID = iArr1(0)
		iArr2 = Split(document.formname.selLevel2.value,",")
		SubSubID = iArr2(1)
		'msgbox SubSubID  
	End If
	id = "6"
	set Node1 = TempData.CreateElement("Schedule")
		Node1.setAttribute "id",id
	 	Node1.setAttribute "sschedno",sschedno
	 	Node1.setAttribute "SubID",SubID 
	 	Node1.setAttribute "SubSubID",SubSubID 
		Node1.setAttribute "sOrgID",sOrgID
		Node1.setAttribute "sFinyr",sFinyr 
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

Function CheckVal()
	IF document.formname.selSch.selectedIndex = 0 Then
		MsgBox "Select Schedule "
		document.formname.selSch.focus()
		CheckVal = False
		Exit Function
	End IF
	
	IF document.formname.selLevel1.selectedIndex = 0 and document.formname.selLevel1.value <>"A" Then
		MsgBox "Select Level 1 "
		document.formname.selLevel1.focus()
		CheckVal = False
		Exit Function
	End IF
	
	IF document.formname.optLevel(1).checked = True and document.formname.txtLev2.value = "" Then
		MsgBox "Select Level 2 "
		document.formname.selLevel2.focus()
		CheckVal = False
		Exit Function
	End IF
	
	IF document.formname.optMode(1).checked = True and document.formname.hAccHead.value = 0 Then
		MsgBox "Select Account Head "
		CheckVal = False
		Exit Function
	End IF
	CheckVal = True	
End Function
Function DispName()
	document.formname.txtLev.value = document.formname.selSch.options(document.formname.selSch.selectedIndex).text        
End Function

	
</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "plBsScheduleSubHeadsPopup", kind: "BS" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="DispName()">
<form method="POST" name="formname">
<input type=hidden name="hOrgId" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sfinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hAccHead" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">B/S Setup
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="650" >
				<TR>
					<TD class=TabBodyWithTopLine width="786">
						<table border="0" cellpadding="0" cellspacing="0" width="723">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" width="702">
								</td>
                            </tr>
                             <tr>
								<td class="FieldCellSub" width="187">Select Schedule</td>
                                <td class="FieldCell" width="268">
									<select size="1" name="selSch" class="formElem" onchange="SelHead()">
									<option value="0">Select</option>
									<%
									sql = "SELECT BSHeadID, BSHeading, FinYear FROM dbo.Acc_M_BSSetupHeads"
									with Objrs1
										.CursorLocation = 3
										.CursorType = 3
										.ActiveConnection = con
										.Source = sql
										.Open
									End With
									Set Objrs1.ActiveConnection = Nothing 
									while not Objrs1.EOF 
										IF CStr(sSchedNo) = CStr(Objrs1(0)) Then
									%>
														<Option Value="<%=Objrs1(0)%>" Selected><%=Objrs1(1)%></Option>
										<% Else %>
														<Option Value="<%=Objrs1(0)%>"><%=Objrs1(1)%> </Option>
											<%
										End IF
									%>					
									</option>
									<%
									Objrs1.MoveNext 
									wend
									Objrs1.Close
									%>
									<Option Value="A"><%="AddNew"%> </Option>
									</select>
									</td>
                                <td class="FieldCell" width="331">
                                    <input type=text name=txtLev size=56 class="Formelem" align="Right" value="" maxlength="200">
								</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="187">Add/Modify Level</td>
										<td class="FieldCell" width="268">
											<Input type="radio" name="optLevel" value="Lev1" class="FormElem"  onclick="LevelFun()" Checked>Level1
											<Input type="radio" name="optLevel" value="Lev2" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
										<td class="FieldCell" width="331">
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="187">Select Level 1</td>
										<td class="FieldCell" width="359">
											<select size="1" name="selLevel1" class="FormElem" OnChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select BSSubID,BSSubSubID,EntryType,Computemode,isNull(BSSubHeadingName,'') as"&_
													" BSSubHeadingName,Hierachy From ACC_M_BSSetupsubheads Where "&_
													" FinYear = '"&sFinyr&"' and BSHeadID = "&sSchedNo&" and BSsubsubid = 0 Order By Hierachy "
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF
													Dim AcVal 
													AcVal = 0
													If Objrs1(2)="S" Then
													sql = "select BSSubHeadValue from Acc_T_BSAcDetail where BSHeadID='"&sSchedNo&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
													With objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sql
														.ActiveConnection = con
														.Open
													End with
													If Not Objrs.EOF= True Then	AcVal = Objrs(0) Else AcVal = 0
													Objrs.Close 
													End If
													If objrs1(2)="A" Then
														sql = "select ApplicableACHeadCode from Acc_T_BSAcDetail where BSHeadID='"&sSchedNo&"' and BSSubID='"&objrs1(0)&"' and BSSubSubID='"&objrs1(1)&"'"    
													With Objrs 
														.CursorLocation = 3
														.CursorType = 3
														.Source = sql
														.ActiveConnection = con
														.Open
													End with
													If Not Objrs.EOF = True Then AcVal = Objrs(0) Else AcVal = 0
													Objrs.close
													End If
												
											%>
											<option value="<%=Objrs1(0)&","&Objrs1(1)&","&objrs1(2)&","&objrs1(3)&","&AcVal&","&objrs1(5) %>"><%Response.Write (Objrs1("BSSubHeadingName")) %></b></option>
											<%Objrs1.MoveNext 
											loop
											Objrs1.Close
											%>
											<option Value="A">Add New</option>
											</select>
										</td>
										<td class="FieldCell" width="233">
                                            <input type=text name=txtLev1 size=56 class="Formelem" align="Right" maxlength="200">
										</td>
									</tr>
								  <tr>
								<td class="FieldCellSub" width="187">Select Level 2</td>
										<td class="FieldCell" width="359">
										<select size="1" name="selLevel2" class="FormElem" Disabled onchange="setlevelfun1()">
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>
									
									</td>
									<input type=hidden name=hSubHeadName value="">
										<td class="FieldCell" width="233">
                                        <input type=text name=txtLev2 size=56 class="Formelem" align="Right" maxlength="200">
									</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Mode</td>
								<td class="FieldCell" width="599" colspan="2">
									<Input type=radio name=optMode value="D" class="FormElem" checked OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A1" class="FormElem" OnClick="ModeFun()">
                                    Schedule
									<Input type=radio name=optMode value="A2" class="FormElem" OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Select</td>
								<td class="FieldCell" width="359">
								<Input type="Button" name="ButAcHead" value="Select" class="ActionButton"  disabled OnClick="popAccList()" >
                                &nbsp;</td>
								<td class="FieldCell" width="233">
								<Input type="text" name="txtAcHead" size="56" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">FinYear</td>
										<td class="FieldCell" colspan="2" width="599">
											<select size="1" name="FinYear" class="FormElem">
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option> 
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Compute Mode</td>
								<td class="FieldCell" width="359">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>
								
								<td class="FieldCell">Hierarchy &nbsp;&nbsp; 
								<Input type="text" name="txtHierarchy" size="5" class="formElem" value="" ></td>
								</tr>
                            <td colspan="3">
		
								<table border="0" cellpadding="0" cellspacing="0" width="776">
										<tr>
											<td valign="middle" class="ActionCell" width="770">
                                                <p align="center"> 
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "SaveXML()" >&nbsp;
												<Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick = "Del()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">
											</td>
										</tr>
									</table>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							<tr>
								<td align="center" colspan="3" class="BottomPack" width="702">
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

