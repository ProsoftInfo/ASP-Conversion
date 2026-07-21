<%@ Language=VBScript %>
<%	option explicit	%>
<%	Response.Expires = -10 %>
<%
	'Program Name				:	InternalReceiptDelete.asp
	'Module Name				:	Inventory (Receipt Updation)
	'Author Name				:	Ragavendran R
	'Created On					:	Mar 25,2014
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptInternalEntry.asp
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
<body>
<form name="formname" action ="MATERIALRECEIPTS.ASP?RCPT=A">
<%
    Dim rsTemp
	Dim dtCurr,nIntRecNo
	Dim sQuery,sInvRecNo
	
	set rsTemp = Server.CreateObject("ADODB.Recordset")
	
	nIntRecNo = Request("RcptNo")
	
	sQuery = "Select IsNull(InvRecNo,0) from APP_T_InternalReceiptHeader where InternalReceiptNo="& nIntRecNo 
	rsTemp.Open sQuery,con
	if not rsTemp.EOF then
	    sInvRecNo = rsTemp(0)
	end if 
	rsTemp.Close 
	if trim(sInvRecNo)<>"0" then
	    %>
	            <SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	                alert("This Receipt is already accounted not able to delete.")
	                document.formname.submit
	            </script>
	    <%
	else
	    sQuery = "Delete from APP_T_InternalReceiptDetails where InternalReceiptNo="& nIntRecNo 
	    con.execute sQuery
	    
	    sQuery = "Delete from APP_T_InternalReceiptHeader where InternalReceiptNo="& nIntRecNo 
	    con.execute sQuery
	    
	    %>
	            <SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	                alert("Records deleted successfully.")
	                document.formname.submit
	            </script>
	    <%
	end if 
	con.close
%>
</form>
</body>