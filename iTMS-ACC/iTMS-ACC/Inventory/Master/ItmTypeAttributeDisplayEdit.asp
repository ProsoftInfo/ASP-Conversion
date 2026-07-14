<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmTypeAttributeDisplayEdit.asp
	'Module Name				:	Inventory (Item Type Attribute)
	'Author Name				:	Ragavendran
	'Created On					:	March 6,2010
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
dim newxml,Root,PageNode,OptionNode,AttributeNode
dim dcrs,dcrs1,sSql,sClassCode
dim sTypeID

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

Set Root = newxml.createElement("Root")

sTypeID = Request.QueryString("ItemType")
sClassCode = Request.QueryString("ClassCode")
	'Response.Write sTypeID & vbCrLf
	'sSql ="Select ItemTypeAttributeID,ItemTypeID,HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal,ClassificationCode from INV_M_ITEMTYPEATTRIBUTES where ItemTypeID = '"& sTypeID &"' and ClassificationCode = "& sClassCode
	sSql ="Select ItemTypeAttributeID,HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal,ClassificationCode from INV_M_ITEMTYPEATTRIBUTES where ClassificationCode = "& sClassCode
	'Response.Write sSql & vbCrLf 
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open 
	end with
	if not dcrs.EOF then
		do while not dcrs.EOF 
		
				
				set AttributeNode = newxml.createElement("Attribute")
					AttributeNode.setAttribute "ID",dcrs(0)
					AttributeNode.setAttribute "ItemType",""'dcrs(1)
					AttributeNode.setAttribute "Header",dcrs(1)
					AttributeNode.setAttribute "Name",dcrs(2)
					AttributeNode.setAttribute "Type",dcrs(3)
					AttributeNode.setAttribute "Length",dcrs(4)
					AttributeNode.setAttribute "Decimal",dcrs(5)
					AttributeNode.setAttribute "ClassCode",dcrs(6)
					
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "Select ItemTypeHeaderName from Inv_M_ItemTypeHeader where HeaderID = "&dcrs(1)
						'Response.Write dcrs1.Source 
						.ActiveConnection = con
						.Open
					end with
					if not dcrs1.EOF then
						AttributeNode.setAttribute "HName",dcrs1(0)
					end if	
					dcrs1.Close 
					
					Root.appendChild AttributeNode
					
						sSql = "Select OptionValue,OptionName from INV_M_ITEMTYPEOPTIONS where ItemTypeAttributeID = "&dcrs(0)
					'Response.Write sSql & vbCrLf 
						with dcrs1
							.CursorLocation = 3
							.CursorType = 3
							.Source = sSql
							.ActiveConnection = con
							.Open
						end with
						if not dcrs1.EOF then
							do while not dcrs1.EOF 
								set OptionNode= newxml.createElement("Option")
									OptionNode.setAttribute "Value",dcrs1(0)
									OptionNode.setAttribute "Name",dcrs1(1)
								AttributeNode.appendChild OptionNode
							dcrs1.MoveNext 
							loop
						end if
						dcrs1.Close
			dcrs.MoveNext 
		loop
	end if
	dcrs.Close 
	Response.ContentType = "text/XML"
	Response.Write Root.xml
%>

