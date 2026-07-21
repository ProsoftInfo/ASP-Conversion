<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtAttWiseLotDetInsert.asp
	'Module Name				:	Inventory (Stock Management)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 08, 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
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
<%
Dim objDom,Root,Node,LotNode,SerialNode,nSerialNo
Dim nItemCode,nClassCode,nLotNo,nOptValue,sSelection,sSql,sIType

set objDom = Server.CreateObject("Microsoft.XMLDOM")
objDom.load(Request)
objDom.async = false

set Root = objDom.documentElement

con.beginTrans

If Root.haschildNodes Then
	For Each Node in Root.childNodes
		If Node.nodeName = "Item" Then
			nItemCode = Node.getAttribute("ICode")
			nClassCode = Node.getAttribute("CCode")
			sIType = Node.getAttribute("ItemTypeID")
			
			For Each LotNode in Node.childNodes
				If LotNode.nodeName = "Lot" Then
					
					nLotNo = LotNode.getAttribute("No")
					nOptValue = LotNode.getAttribute("OptValue")
					sSelection =  LotNode.getAttribute("Selection")
					If nOptValue <> "" and sSelection <> "N" Then
						sSql = "Update Inv_T_LocationLot set AttributeList = '"& nOptValue &"' where Itemcode = "& nItemCode &" and classificationcode ="& nClassCode &" and LotNumber = "& nLotNo &" "
						'Response.Write "<p>Sql="&sSql
						con.Execute sSql
					End IF
				
				Elseif LotNode.nodename = "BaseItem" Then
					
					nOptValue = LotNode.getAttribute("OptValue")
					sSelection =  LotNode.getAttribute("Selection")
					
						For Each SerialNode in LotNode.childNodes
							If SerialNode.nodeName = "Serial" Then
								
								nSerialNo = SerialNode.getAttribute("No")
								If nOptValue <> "" and sSelection <> "N" Then
									sSql = "Update Inv_T_LocationLot set AttributeList = '"& nOptValue &"' where Itemcode = "& nItemCode &" and classificationcode ="& nClassCode &" and SerialNumber = "& nSerialNo &" "
									'Response.Write "<p>Sql="&sSql
									con.Execute sSql
								End IF
								
							End IF
						Next
				
				End IF
			Next
			
		End IF
	Next
End If

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	'response.end
	con.CommitTrans
end if

%>