<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasClassificationAttributeInsert.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
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
<!--#include file="../../include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			/*if (confirm("Do You want to define another Classification Attribute"))
				window.location.href = "MasClassificationAttributeEntry.asp"
			else*/
				//document.form1.target = "bodyFrame"
				window.location.href = "MasClassificationAttributeEntry.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<FORM id="form1" target="bodyFrame">
<%
'XML DOM Variables
Dim oDOM,newElem,Root,objfs

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql,iCount,dcrs1,bflag,iCounter,iAMaxCode
bflag = false
dim sAttrName,sAttrType,iDataLen,iDecimal,spGroup,spPath,arrspPath,arrspGroup
dim sGCode,sGName
sAttrName = trim(Request.Form("txtAttrName"))
sAttrType = trim(Request.Form("selDataType"))
iDataLen = trim(Request.Form("txtDataLen"))
iDecimal = trim(Request.Form("txtDecimal"))
spPath = trim(Request.Form("hgPath"))

arrspGroup = split(trim(Request.Form("hpGroup")),":")
spGroup = trim(arrspGroup(1))

arrspPath = Split(spPath, "\")

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

For iCounter = 2 To UBound(arrspPath)
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME FROM INV_M_CLASSIFICATION WHERE GROUPNAME = " & Pack(arrspPath(iCounter)) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing	
	set sGCode = dcrs(0)
	set sGName = dcrs(1)
	do while not dcrs.EOF
		sGName = sGName
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ATTRIBUTENAME FROM INV_M_CLASSIFICATIONATTRIBUTES WHERE LOWER(ATTRIBUTENAME) = " & Pack(lcase(sAttrName)) & " AND GROUPCODE = " & trim(sGCode) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing	
		If not dcrs1.EOF Then
			bflag = true
		end if
		dcrs1.Close
	dcrs.MoveNext
	loop
	dcrs.Close
next
if not bflag then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(GROUPATTRIBUTECODE),0) + 1 FROM INV_M_CLASSIFICATIONATTRIBUTES"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing	
	set iAMaxCode = dcrs(0)
	iAMaxCode = iAMaxCode
	dcrs.Close

	If Not iDecimal = "" Then
		sSql = "INSERT INTO INV_M_CLASSIFICATIONATTRIBUTES (GROUPATTRIBUTECODE,GROUPCODE,ATTRIBUTENAME," &_
			"ATTRIBUTETYPE,ATTRIBUTEDATALENGTH,ATTRIBUTEDECIMALS) VALUES " &_
			"(" & iAMaxCode & "," & spGroup & "," & Pack(sAttrName) & ", " &_
			"" & Pack(sAttrType) & "," & iDataLen & "," & iDecimal & ")"
	Else
		sSql = "INSERT INTO INV_M_CLASSIFICATIONATTRIBUTES (GROUPATTRIBUTECODE,GROUPCODE,ATTRIBUTENAME," &_
			"ATTRIBUTETYPE,ATTRIBUTEDATALENGTH,ATTRIBUTEDECIMALS) VALUES " &_
			"(" & iAMaxCode & "," & spGroup & "," & Pack(sAttrName) & ", " &_
			"" & Pack(sAttrType) & "," & iDataLen & ",0)"
	End If
	'Response.Write sSql
	con.Execute sSql
	sSql = "UPDATE INV_M_CLASSIFICATION SET ATTRIBUTECOUNT = ISNULL(ATTRIBUTECOUNT,0) + 1 WHERE GROUPCODE = " & spGroup
	con.Execute sSql
%>
	<BODY onLoad = "msgbox('Attribute <%=replace(sAttrName,"'","\'")%> has been Created Successfully','Y')">
<%
else
%>
	<BODY onLoad = "msgbox('Attribute <%=replace(sAttrName,"'","\'")%> already Defined Under Classification <%=sGName%>','N')">
<%
End If

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		Response.Write con.Errors(iErrCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
</FORM>