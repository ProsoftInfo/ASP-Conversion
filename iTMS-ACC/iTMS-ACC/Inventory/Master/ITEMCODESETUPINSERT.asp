<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ITEMCODESETUPINSERT.asp
	'Module Name				:	Inventory (Item Code Setup)
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 04,2013
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
/*			if (confirm("Do You want to Create another Item Display"))
				window.location.href = "itmDisplayEntry.asp"
			else*/
				window.location.href = "../welcome_Inventory.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
	dim dcrs
	dim iItemCodeOrder,iItemNameOrder,iDrawingOrder,iCatalogueOrder,str

	iItemCodeOrder  = trim(Request.Form("txtItemCode"))
	iItemNameOrder  = trim(Request.Form("txtItemName"))
	iDrawingOrder  = trim(Request.Form("txtDrawing"))
	iCatalogueOrder  = trim(Request.Form("txtCatalog"))

	set dcrs=server.CreateObject("ADODB.Recordset")
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_M_ITEMDISPLAY"
		.ActiveConnection = con
		.Open 
	end with
	set dcrs.ActiveConnection = Nothing
	
	con.beginTrans
	
	if not dcrs.EOF then
		con.execute "DELETE FROM INV_M_ITEMDISPLAY"
		str="Item Display has been Amended Successfully"
	else
		str="Item Display has been Created Successfully"
	end if
	
	if iItemCodeOrder <> "" then
		con.execute "INSERT INTO INV_M_ITEMDISPLAY( DISPLAYNAME, DISPLAYORDER) " _
		& "VALUES (" & pack("INO") & "," & iItemCodeOrder & ")"
	end if
		
	if iItemNameOrder <> "" then
		con.execute "INSERT INTO INV_M_ITEMDISPLAY( DISPLAYNAME, DISPLAYORDER) " _
		& "VALUES (" & pack("INA") & "," & iItemNameOrder & ")"
	end if

	if iDrawingOrder <> "" then
		con.execute "INSERT INTO INV_M_ITEMDISPLAY( DISPLAYNAME, DISPLAYORDER) " _
		& "VALUES (" & pack("DNO") & "," & iDrawingOrder & ")"
	end if

	if iCatalogueOrder <> "" then
		con.execute "INSERT INTO INV_M_ITEMDISPLAY( DISPLAYNAME, DISPLAYORDER) " _
		& "VALUES (" & pack("CNO") & "," & iCatalogueOrder & ")"
	end if

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
<BODY onLoad = "msgbox('<%=str%>','Y')">
