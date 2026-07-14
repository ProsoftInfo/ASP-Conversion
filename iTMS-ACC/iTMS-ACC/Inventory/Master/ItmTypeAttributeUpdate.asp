<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmTypeAttributeUpdate.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	March 08,2010
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
<%
dim newxml,Root,PageNode,OptNode

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)

Set Root = newxml.documentElement

dim dcrs,dcrs1,sSql
dim sTypeID,iHeader,sAttrName,sAttDataType,iDataLen,iDecLen,sVal,iAttrID,iValue,sClassCode

sTypeID = trim(Root.Attributes.Item(0).nodeValue)
iHeader = trim(Root.Attributes.Item(1).nodeValue)
sAttrName = trim(Root.Attributes.Item(2).nodeValue)
sAttDataType = trim(Root.Attributes.Item(3).nodeValue)
iDataLen = trim(Root.Attributes.Item(4).nodeValue)
iDecLen = trim(Root.Attributes.Item(5).nodeValue)
iAttrID = trim(Root.Attributes.Item(6).nodeValue)
sClassCode = trim(Root.getAttribute("CLASSCODE"))


'Response.Write "iAttrID="& iAttrID

if IsNull(iDataLen) or IsEmpty(iDataLen) or iDataLen = "" then iDataLen = "0"
if IsNull(iDecLen) or IsEmpty(iDecLen) or iDecLen = "" then iDecLen = "0"

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

'sSql ="UPDATE INV_M_ITEMTYPEATTRIBUTES SET ITEMTYPEID = "& Pack(sTypeID) &",HEADERID = "&iHeader &", "&_
'	  "ITEMTYPEATTRIBUTENAME = "& Pack(sAttrName) &",ITEMTYPEATTRIBUTETYPE="&Pack(sAttDataType)&", "&_
'	  "ITEMTYPEATTRIBUTEDATALENGTH="& iDataLen &",ITEMTYPEATTRIBUTEDECIMAL=" & iDecLen &",CLASSIFICATIONCODE ="& sClassCode &" "&_
'	  " WHERE ITEMTYPEATTRIBUTEID = "& iAttrID 
sSql ="UPDATE INV_M_ITEMTYPEATTRIBUTES SET HEADERID = "&iHeader &", "&_
	  "ITEMTYPEATTRIBUTENAME = "& Pack(sAttrName) &",ITEMTYPEATTRIBUTETYPE="&Pack(sAttDataType)&", "&_
	  "ITEMTYPEATTRIBUTEDATALENGTH="& iDataLen &",ITEMTYPEATTRIBUTEDECIMAL=" & iDecLen &",CLASSIFICATIONCODE ="& sClassCode &" "&_
	  " WHERE ITEMTYPEATTRIBUTEID = "& iAttrID 
'	Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql
	
	sSql  = "Delete from INV_M_ITEMTYPEOPTIONS where ItemTypeAttributeID = "& iAttrID
'	Response.Write sSql
	con.execute sSql

	For Each PageNode In Root.childNodes
		for each OptNode in PageNode.childNodes
			if StrComp(OptNode.nodeName,"Option") = 0 then
				sVal = OptNode.getAttribute("Name")

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(MAX(OPTIONVALUE)+1,1) FROM INV_M_ITEMTYPEOPTIONS"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
										
				if not dcrs1.EOF then
					iValue = trim(dcrs1(0))
				end if
				dcrs1.Close

				sSql = "INSERT INTO INV_M_ITEMTYPEOPTIONS (ITEMTYPEATTRIBUTEID,OPTIONVALUE," &_
					"OPTIONNAME) VALUES " &_
					"(" & iAttrID & "," & iValue & "," & Pack(sVal) & ")"
			'	Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
		next
	next
	
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
Response.Write "Y<BR>"
'	con.RollbackTrans
'	Response.End 
	con.CommitTrans
end if

con.close
set con = nothing
%>
