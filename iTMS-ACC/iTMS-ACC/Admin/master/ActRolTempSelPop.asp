<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ActRolTempSelPop.asp
	'Module Name				:	Admin (Role)
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Activity Role Mapping Popup</title>
<base target="_self"/>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML id="ActivityData"><Root></Root></XML>
<XML id="RoleData"><Root></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
'******************************************
Function Init()
    sAppCode = document.formname.hAppCode.value
    sProcessCode = document.formname.hProcessCode.value
    sActCode = document.formname.hActCode.value
    
    set objhttp = createObject("Microsoft.XMLHTTP")
		objhttp.open "GET","XMLGetActTemp.asp?AppCode="&sAppCode&"&ProcessCode="&sProcessCode&"&ActCode="&sActCode,false
		objhttp.send 
		'alert(objhttp.responseText)
		if trim(objhttp.responseXML.xml)<>"" then
		    ActivityData.loadXML(objhttp.responseXML.xml)
		end if
		
		ClearTable
		DisplayTable
End Function
'******************************************
Function DisplayTable() 
    iSer = 0
    iTempCnt = 0 
    set ndRoot = ActivityData.documentElement
    if ndRoot.hasChildNodes() then
        for each ndActivity in ndRoot.childNodes
            if ndActivity.nodeName="Activity" then
                    sActCode = ndActivity.getAttribute("ActCode")
                    sActName = ndActivity.getAttribute("ActName")
                    iTempCnt = ndActivity.getAttribute("TempCnt")
                    if iTempCnt>1 then
                        iSer=iSer + 1
                        set oRow = document.all.tblTempAct.insertRow(document.all.tblTempAct.rows.length)
                            set headerCell=oRow.insertCell()
                            headerCell.innerText = iSer 
                            headerCell.className="ExcelSerial"
                            headerCell.align="center"
                            
                            set headerCell=oRow.insertCell()
                            headerCell.innerText = ""
                            headerCell.className="ExcelDisplayCell"
                            headerCell.align="center"
                            
                            set headerCell=oRow.insertCell()
                            headerCell.innerText = sActName
                            headerCell.className="ExcelDisplayCell"
                            
                            for each ndTemp in ndActivity.childNodes
                                if ndTemp.nodeName="Template" then
                                    TempNo = ndTemp.getAttribute("TempNo")
                                    TempName = ndTemp.getAttribute("TempName")
                                    ProgPath = ndTemp.getAttribute("Path")
                                    sSelect = ndTemp.getAttribute("Select")
                                    
                                    set oRow = document.all.tblTempAct.insertRow(document.all.tblTempAct.rows.length)
	                                    set headerCell=oRow.insertCell()
	                                    headerCell.innerText = ""
	                                    headerCell.className="ExcelSerial"
	                                    headerCell.align="center"
            	                    
	                                    set headerCell=oRow.insertCell()
	                                    if sSelect="Y" then
	                                        set oText = document.createElement("<input type=Checkbox name=ChkTempZ"&sActCode&"Z"&TempNo&" checked>")
	                                    else
	                                        set oText = document.createElement("<input type=Checkbox name=ChkTempZ"&sActCode&"Z"&TempNo&">")
	                                    end if
	                                    headerCell.appendChild(oText)
	                                    headerCell.className="ExcelDisplayCell"
	                                    headerCell.align="center"
            	                        
	                                    set headerCell=oRow.insertCell()
	                                    headerCell.innerHtml =  "&nbsp;&nbsp;&nbsp;"& TempName 
	                                    headerCell.className="ExcelDisplayCell"
            	                    
                                end if
                            next
                    end if 'if iTempCnt>1 then
            end if
        next
    end if
End Function
'**************************************
Function ClearTable()
    K = document.all.tblTempAct.rows.length - 2
	for	i = 1 to  K
		document.all.tblTempAct.deleteRow(2)
	next
End Function
'*****************************************************
Function CheckSubmit()
    
    set ndActRoot = ActivityData.documentElement
    set ndRoleRoot = RoleData.documentElement
    
    if ndActRoot.hasChildNodes() then
        for each ndActivity in ndActRoot.childNodes
            if ndActivity.nodeName="Activity" then
                sActCode = ndActivity.getAttribute("ActCode")
                sActName = ndActivity.getAttribute("ActName")
                iTempCnt = ndActivity.getAttribute("TempCnt")
                if iTempCnt>1 then
                        for each ndTemp in ndActivity.childNodes
                            if ndTemp.nodeName="Template" then
                                
                                TempNo = ndTemp.getAttribute("TempNo")
                                TempName = ndTemp.getAttribute("TempName")
                                ProgPath = ndTemp.getAttribute("Path")
                                sSelect = ndTemp.getAttribute("Select")
                                set Obj = eval("document.formname.ChkTempZ"&sActCode&"Z"&TempNo)
                                
                                if Obj.checked and trim(sSelect)="N" then
                                
                                    set ndTempRole = RoleData.createElement("Role")
                                        ndTempRole.setAttribute "ActCode",sActCode
                                        ndTempRole.setAttribute "ActName",sActName
                                        ndTempRole.setAttribute "TempNo",TempNo
                                        ndTempRole.setAttribute "TempName",TempName
                                    ndRoleRoot.appendChild ndTempRole
                                end if
                            end if
                        next
                else
                    set ndTempRole = RoleData.createElement("Role")
                        ndTempRole.setAttribute "ActCode",sActCode
                        ndTempRole.setAttribute "ActName",sActName
                        ndTempRole.setAttribute "TempNo","1"
                        ndTempRole.setAttribute "TempName",sActName
                    ndRoleRoot.appendChild ndTempRole
                end if 'if iTempCnt>1 then
            end if
        next
    end if
    
    window.close
End Function
'*****************************************************
Function window_onunload()
    set sRoot = RoleData.documentElement
		sRoot.setAttribute "Done","Y"
	set window.returnvalue = RoleData.documentElement
End Function

</SCRIPT>

<%
    Dim nAppCode,nProcessCode,nActCode
    
    nAppCode = Request.QueryString("AppCode")
	nProcessCode = Request.QueryString("ProcessCode")
	nActCode = Request.QueryString("ActCode")
%>
</head>
<body leftmargin="5" topmargin="0" marginheight="0" marginwidth="0" onload="Init()" >

	<form method="POST" name="formname" action="">
	<Input type="hidden" name="hItemRows" value="">
	<Input type="hidden" name="hLastSelectedPractice" value="">
	<Input type="hidden" name="hAppCode" value="<%=nAppCode%>">
	<Input type="hidden" name="hProcessCode" value="<%=nProcessCode%>">
	<Input type="hidden" name="hActCode" value="<%=nActCode%>">
		
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Activity Mapping
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				    <tr>
				        <td class="toppack"></td>
				    </tr>
				    <tr>
						<td class="TabBody">
						    <table border="0" cellpadding="0" cellspacing="1" width="100%">
						        <tr>
						            <td>    
						                <div style="width:470px;height:420px">
							                <table border="0" id="tblTempAct" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
					                            <tr>
					                                <td class="ExcelHeaderCell" align="center" rowspan="2" style="width:20px">S.No.</td>
					                                <td class="ExcelHeaderCell" align="center" rowspan="2">&nbsp;</td>
					                                <td class="ExcelHeaderCell" align="center">Activity Name</td>
					                            </tr>
					                            <tr>
					                                <td class="ExcelHeaderCell" align="center">Template Name</td>
					                            </tr>
							                </table>
							            </div>
						            </td>
						        </tr>
						    </table>
						</td>
					</tr>
					 <tr>
				        <td class="BottomPack"></td>
				    </tr>
					<tr>
					    <td>
					        <table width="100%">
					            <tr>
					                <td class="ActionCell" align="center">
					                    <input type="button" name="btnDone" value="Done" class="ActionButton" onclick="CheckSubmit()">
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
</body>
</html>
