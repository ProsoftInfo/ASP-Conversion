<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasCategoryUpdate.asp	
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 16, 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to Amend another Category")) 
				window.location.href = "MasCategoryAmendEntry.asp"
			else
				window.location.href = "MasUOMAmendEntry.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
'XML DOM Variables
Dim oDOM,newElem,Root,objfs,tempNode,bFlag,sChkFlag

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql
dim sCatCode,sCatName,sCatShName,arrTemp
arrTemp = split(trim(Request.Form("selCategory")),"|")

sCatCode = trim(arrTemp(0))
sCatName = trim(Request.Form("txtCatName"))
sCatShName = trim(Request.Form("txtCatShName"))
sChkFlag = Request("hFlag")
'Response.Write "sChkFlag ="& sChkFlag 

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

If trim(sChkFlag) = "N" then 'New Category
	sCatCode = Request.Form("txtCatCode")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CATEGORYCODE FROM INV_M_CLASSIFICATIONCATEGORY WHERE LOWER(CATEGORYCODE) = " & Pack(lcase(sCatCode)) & " OR LOWER(CATEGORYNAME) = " & Pack(lcase(sCatName)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if dcrs.EOF then
		sSql = "INSERT INTO INV_M_CLASSIFICATIONCATEGORY (CATEGORYCODE,CATEGORYNAME,CATEGORYSHORTNAME) VALUES" &_
			"(" & Pack(ucase(sCatCode)) & "," & Pack(ucase(sCatName)) & "," & Pack(ucase(sCatShName)) & ")"
		'Response.Write sSql & "<BR>"
		con.Execute sSql 
		
		if objfs.FileExists(Server.MapPath("../xmldata/Category.xml")) then
			oDOM.Load server.MapPath("../xmldata/Category.xml")
			Set Root = oDOM.documentElement
		else	
			Set Root = oDOM.createElement("Root")
			oDOM.appendChild Root
		end if

		Set newElem = oDOM.createElement("Category")

		newElem.setAttribute "CATEGORYCODE", ucase(sCatCode)
		newElem.setAttribute "CATEGORYNAME", ucase(sCatName)
		newElem.setAttribute "CATEGORYSHORTNAME", ucase(sCatShName)
		Root.appendChild newElem
		
		oDOM.Save server.MapPath("../xmldata/Category.xml")
		
	%>
		<BODY BGCOLOR="#336699" onLoad = "msgbox('Category <%=replace(sCatName,"'","\'")%> has been Created Successfully','Y')">
	<%
	else
	%>
		<BODY BGCOLOR="#336699" onLoad = "msgbox('Category Code <%=sCatCode%> or Category Name <%=replace(sCatName,"'","\'")%> already created','N')">
	<%
	End If
Else 'If trim(sChkFlag) = "A" then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3	
		.Source = "SELECT CATEGORYCODE FROM INV_M_CLASSIFICATIONCATEGORY WHERE LOWER(CATEGORYNAME) = " & Pack(lcase(sCatName)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if dcrs.EOF then
		sSql = "UPDATE INV_M_CLASSIFICATIONCATEGORY SET CATEGORYNAME = " & Pack(ucase(sCatName)) & ", " &_
			"CATEGORYSHORTNAME = " & Pack(ucase(sCatShName)) & " WHERE CATEGORYCODE = " & Pack(sCatCode) & ""
		'Response.Write sSql & "<BR>"
		con.Execute sSql
		
		bFlag = true
		if objfs.FileExists(Server.MapPath("../xmldata/Category.xml")) then
			oDOM.Load server.MapPath("../xmldata/Category.xml")
			Set Root = oDOM.documentElement
			for each tempNode in Root.childNodes
				if ucase(tempNode.attributes.getNamedItem("CATEGORYCODE").value) = ucase(sCatCode) then
					tempNode.attributes.getNamedItem("CATEGORYNAME").value = ucase(sCatName)
					tempNode.attributes.getNamedItem("CATEGORYSHORTNAME").value = ucase(sCatShName)
					bFlag = false
					exit for
				end if
			next
		else	
			Set Root = oDOM.createElement("Root")
			oDOM.appendChild Root
			bFlag = true
		end if
		if bFlag then
			Set newElem = oDOM.createElement("Category")

			newElem.setAttribute "CATEGORYCODE", ucase(sCatCode)
			newElem.setAttribute "CATEGORYNAME", ucase(sCatName)
			newElem.setAttribute "CATEGORYSHORTNAME", ucase(sCatShName)
			Root.appendChild newElem
		end if	
		oDOM.Save server.MapPath("../xmldata/Category.xml")
		'Response.End 
	%>
		<BODY onLoad = "msgbox('Category <%=replace(sCatName,"'","\'")%> has been Updated Successfully','Y')">
	<%
	else
	%>
		<BODY onLoad = "msgbox('Category Name <%=replace(sCatName,"'","\'")%> already created','N')">
	<%
	End If
End If 'If trim(sChkFlag) = "N" then
'Response.End 
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
