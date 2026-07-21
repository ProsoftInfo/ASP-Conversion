
<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetAttributeName.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	Maheswari
	'Created On					:	March 26, 2008
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	JULY 24,2010
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

<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<%
	Dim sTemp,i,sOptName,iOptVal,dcrs1,sArrAttribList,sAttributeID,sAttribList,dcrs2
	sArrAttribList = split(trim(Request("Para")),"#")
	if UBound(sArrAttribList)>0 then
		sAttributeID = sArrAttribList(0)
		sAttribList = sArrAttribList(1)
	else
		sAttributeID = sArrAttribList(0)
	end if
	'Dim sTemp,i,sOptName,iOptVal,dcrs1
	Set dcrs1= Server.CreateObject("ADODB.RecordSet")
	Set dcrs2= Server.CreateObject("ADODB.RecordSet")
	Response.Write  "sAttribList = "& sAttribList 
	Response.Write "sAttributeID = "& sAttributeID 
	
	if UBound(sArrAttribList)>1 then
		If trim(sAttribList) <> "0" and trim(sAttribList)<>"" then
			sOptName = ""
			iOptVal = ""
			
			sTemp = split(sAttribList,",")
			For i = 0 to UBOUND(sTemp) 
				iOptVal = sTemp(i)
				
				if iOptVal <> "" then 
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal&" "
						.ActiveConnection = con
						.Open
					end with
					If not dcrs1.EOF then
						sOptName = sOptName &","& dcrs1(0)
					End If
					dcrs1.Close 
				else
					sOptName = ""
				end if	
			Next
		End If
	else
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue in ("& sAttributeID &")"
			'Response.Write dcrs1.Source 
			.ActiveConnection = con
			.Open
		end with
		If not dcrs1.EOF then
			sOptName = sOptName &","& dcrs1(0)
		End If
		dcrs1.Close 
		
	end if
	
	
	IF sOptName <> "" then 
		sOptName = " [" & mid(sOptName,2) &"] "
	End IF
	
	Response.ContentType="text/xml"
	Response.Write  sOptName

%>
