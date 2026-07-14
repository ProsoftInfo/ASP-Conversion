<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmTypeAttributeInsert.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	January 29, 2003
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
dim newxml,Root,PageNode

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
sClassCode = trim(Root.getAttribute("CLASSCODE"))

if IsNull(iDataLen) or IsEmpty(iDataLen) or iDataLen = "" then iDataLen = "0"
if IsNull(iDecLen) or IsEmpty(iDecLen) or iDecLen = "" then iDecLen = "0"

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	'.Source = "SELECT ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES WHERE ITEMTYPEID = " & Pack(sTypeID) & " AND HEADERID = " & iHeader & " AND LOWER(ITEMTYPEATTRIBUTENAME) = " & Pack(lcase(sAttrName)) & " and ClassificationCode ="& sClassCode 
	.Source = "SELECT ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = " & iHeader & " AND LOWER(ITEMTYPEATTRIBUTENAME) = " & Pack(lcase(sAttrName)) & " and ClassificationCode ="& sClassCode 
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if dcrs.EOF then
	Response.Write "Y<BR>"
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ITEMTYPEATTRIBUTEID)+1,1) FROM INV_M_ITEMTYPEATTRIBUTES"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing
							
	if not dcrs1.EOF then
		iAttrID = trim(dcrs1(0))
	end if
	dcrs1.Close

'	sSql = "INSERT INTO INV_M_ITEMTYPEATTRIBUTES (ITEMTYPEATTRIBUTEID,ITEMTYPEID,HEADERID," &_
'		"ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE,ITEMTYPEATTRIBUTEDATALENGTH,ITEMTYPEATTRIBUTEDECIMAL,CLASSIFICATIONCODE) VALUES " &_
'		"(" & iAttrID & "," & Pack(sTypeID) & "," & iHeader & "," &_
'		"" & Pack(sAttrName) & "," & Pack(sAttDataType) & "," & iDataLen & "," & iDecLen & ","& sClassCode &")"
    
    sSql = "INSERT INTO INV_M_ITEMTYPEATTRIBUTES (ITEMTYPEATTRIBUTEID,HEADERID," &_
		"ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE,ITEMTYPEATTRIBUTEDATALENGTH,ITEMTYPEATTRIBUTEDECIMAL,CLASSIFICATIONCODE) VALUES " &_
		"(" & iAttrID & "," & iHeader & "," &_
		"" & Pack(sAttrName) & "," & Pack(sAttDataType) & "," & iDataLen & "," & iDecLen & ","& sClassCode &")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	For Each PageNode In Root.childNodes
		if StrComp(PageNode.nodeName,"OptionEntry") = 0 then
			sVal = PageNode.Attributes.Item(0).nodeValue

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
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
	next
else
	Response.Write "N<BR>"
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
