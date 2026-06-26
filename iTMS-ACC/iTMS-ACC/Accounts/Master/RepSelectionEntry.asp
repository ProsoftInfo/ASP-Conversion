<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	RepSelectionEntry.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 04,2012
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/populate.asp" -->
<%
    Dim rsTemp
    Dim sAgentName,sAgentShortName,sAgentAddress1,sAgentAddress2,sPartyCode,sQuery,sCity,sPinCode
    Dim sAgentAddress4,sAgentPhone,sAgentFax,sAgentMailID,sExternalOrInternal,sAgentType,sArrTemp
    Dim iAgentEntryID,iRepAreaCode,sOrgID
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    sPartyCode = Request("PartyCode")
    sOrgID = Session("organizationcode")
    
    Response.Write "<font colo=red>"
    
    iAgentEntryID = 0
    sQuery = "Select RepAreaCode,RepAgentEntryID from APP_R_OrgParty where PartyCode = "& sPartyCode &" and OUDefinitionID="& pack(sOrgID)
    'Response.Write sQuery 
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        iRepAreaCode = rsTemp(0)
        iAgentEntryID = rsTemp(1)
    end if
    rsTemp.Close 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self"></base>
<xml id="RepAreaData"><Root></Root></xml>
<xml id="AgentData"><Root></Root></xml>
<xml id="ConPerForArea"><Root></Root></xml>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>

<Script language="VBScript">
'**************************************************
Function PopulateContPerson(obj)
    if obj.value="0" then
		document.formname.SelContPerson.length = 1
        document.formname.SelContPerson(document.formname.SelContPerson.length-1).value = "0"
        document.formname.SelContPerson(document.formname.SelContPerson.length-1).text = "Name"
		exit function
    End IF
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","/Common/XMLGetConPersonForRepArea.asp?AreaID="&obj.value,false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        ConPerForArea.loadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if
    set ndAreaRoot = ConPerForArea.documentElement
    document.formname.SelContPerson.length = 0
    if ndAreaRoot.hasChildNodes() then
        For Each ndChild in ndAreaRoot.childNodes
            document.formname.SelContPerson.length = document.formname.SelContPerson.length + 1
            document.formname.SelContPerson(document.formname.SelContPerson.length-1).value = ndChild.getAttribute("ID")&"|"&ndChild.getAttribute("LCode")
            document.formname.SelContPerson(document.formname.SelContPerson.length-1).text = ndChild.getAttribute("CPName")
        Next
    else
        document.formname.SelContPerson.length = document.formname.SelContPerson.length + 1
        document.formname.SelContPerson(document.formname.SelContPerson.length-1).value = "0|0"
        document.formname.SelContPerson(document.formname.SelContPerson.length-1).text = "Not Available"
    end if
End Function

'************************************************
Function CreateRep()
    sValue = showModalDialog("RepCreationEntry.asp","","dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
    PopulateArea()
End Function
'***************************************************
function checkSubmit()
        Set ndRoot = AgentData.documentElement
        if document.formname.SelRepArea(document.formname.SelRepArea.selectedIndex).value ="0" then
            alert("Select Area Name")
            exit function
        elseif document.formname.SelContPerson(document.formname.SelContPerson.selectedIndex).value ="0" then
            alert("Select Representative Name")
            exit function
        end if
         
        sAgentEntryID = split(trim(document.formname.SelContPerson(document.formname.SelContPerson.selectedIndex).value),"|")(0)
        sAreaCode = document.formname.SelRepArea(document.formname.SelRepArea.selectedIndex).value
        
        Set newElem= AgentData.createElement("AGENT")
        newElem.SetAttribute "AgentEntryID",sAgentEntryID
        newElem.setAttribute "AreaCode",sAreaCode
        ndRoot.AppendChild newElem

        set objhttp = CreateObject("Microsoft.XMLHTTP")
        objhttp.open "POST","XMLSaveParty.asp?Name=RepAllocation&Mod=Party",false
        objhttp.send AgentData.XMLDocument
        
        set objhttp = CreateObject("Microsoft.XMLHTTP")
        objhttp.open "POST","RepSelectionInsert.asp?PartyCode=" &document.formname.hPartyCode.value,false
        objhttp.send
        if Trim(objhttp.responseText)<>"" then
            alert(objhttp.responseText)
        else
            window.returnvalue = "Done"
            window.close 
        end if
end function
'*************************************
Function PopulateArea()
set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","/Common/XMLGetRepresentingArea.asp",false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        RepAreaData.loadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if
    set ndAreaRoot = RepAreaData.documentElement
    if ndAreaRoot.hasChildNodes() then
        document.formname.SelRepArea.length = 1
        For Each ndChild in ndAreaRoot.childNodes
            document.formname.SelRepArea.length = document.formname.SelRepArea.length + 1
            document.formname.SelRepArea(document.formname.SelRepArea.length-1).value = ndChild.getAttribute("ACode")
            document.formname.SelRepArea(document.formname.SelRepArea.length-1).text = ndChild.getAttribute("AName")
        Next
    end if
    
    sAreaCode =document.formname.hAreaCode.value 
    For iCnt = 0 to document.formname.SelRepArea.length -1 
        if trim(sAreaCode) = Trim(document.formname.SelRepArea(iCnt).value) then
            document.formname.SelRepArea.selectedIndex = iCnt
        end if
    Next
    
End Function
'********************************
Function Init()
sAreaCode =  document.formname.hAreaCode.value 
sAgent =  document.formname.hAgentEntryID.value  
    PopulateArea()
    for iCnt = 0 to document.formname.SelRepArea.length - 1
        if trim(sAreaCode) = Trim(document.formname.SelRepArea(iCnt).value) then
            document.formname.SelRepArea.selectedIndex = iCnt
            exit for
        end if
    next
    PopulateContPerson(eval("document.formname.SelRepArea"))
    for iCnt = 0 to document.formname.SelContPerson.length - 1
        if trim(sAgent) = split(Trim(document.formname.SelContPerson(iCnt).value),"|")(0) then
            document.formname.SelContPerson.selectedIndex = iCnt
            exit for
        end if
    next
End Function
'************************************
</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "representativeSelection" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="post" name="formname" action="">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hAreaCode" value="<%=iRepAreaCode%>">
<input type="hidden" name="hAgentEntryID" value="<%=iAgentEntryID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="middle" class=PageTitle height="20"><p align="center">Representative Selection</p>
		</td>
    </tr>
	<tr>
		<td align="middle" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="middle" colspan="2" class="MiddlePack">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
                                <td class="FieldCellSub">Area&nbsp;</td>
					            <td class="FieldCell">
					            <select name=SelRepArea  class="FormElem" onchange="PopulateContPerson(this)">
					                <option value="0">Area</option>
					            </select>&nbsp;&nbsp;
					        </td>
                           </tr>
                           <tr>
                                <td class="FieldCellSub">Representative&nbsp;</td>
					            <td class="FieldCell">
					            <select name="SelContPerson" class="FormElem">
					                <option value="0">Name</option>
					                    <%
					                    sQuery = "Select AgentEntryID,LocationCode,isNull(ContactPersonName,'') from APP_M_AgentLocations"
					                    rsTemp.Open sQuery,con
					                    Do while Not rsTemp.EOF
									        If rsTemp(2) <> "" Then
									        %>
										        <option value="<%=rstemp(0)%>|<%=rstemp(1)%>"><%=rstemp(2)%></option>
									        <%
									        End If
									        rsTemp.MoveNext
					                    Loop
					                    rsTemp.Close
					                    %>
					            </select>
					        </td>
					        </tr>
					        <tr>
								<td align="middle" colspan="2" class="MiddlePack">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=2>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td valign="middle" class="ActionCell">
											<p align="center">
											    <input type="button" value="Create Representative" onClick="CreateRep()" name="btnCreate" class="ActionButtonX" tabindex="3" > 
                                                <input type="button" value="Save" onClick="checkSubmit()" name="B2" class="ActionButton" tabindex="3" > 
                                                <input type="button" value="Close" name="B3" class="ActionButton" tabindex="3" onclick="window.close()"> 
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td align="middle" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</TD>
				</TR>
			</TABLE>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
