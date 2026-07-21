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
	'Program Name				:	XMLGetStockDetails.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	Maheswari
	'Created On					:	March 22, 2008
	'Modified By				:	Ragavendran R
	'Modified On				:	July 24,2010
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
	dim dcrs,dcrs1,sSql,OutData,Root,newElem,newElem1,sOrgID
	dim sTemp,arrTemp,iItemCode,iClass,iEntNo,sFinFrom,sFinTo
	Dim sQuery,sAttList,sAttID,sArrAttribute,sArrAttTemp,sArrAttSub1
	Dim iCnt
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")

	sTemp = trim(Request("Para"))
	'Response.Write sTemp
	Response.Write Request.QueryString 
	Response.Write vbCrLf

	arrTemp = split(sTemp,":")
	iItemCode = arrTemp(0)
	iEntNo =arrTemp(1)
	iClass = arrTemp(2)
	sOrgID = arrTemp(3)
	sFinFrom = arrTemp(4)
	sFinTo = arrTemp(5)
	'Response.Write "arrTemp(6)="& arrTemp(6)
	sArrAttTemp = Split(arrTemp(6),",")
	
	     For iCnt = 0 to UBound(sArrAttTemp)
	        sArrAttSub1 = Split(sArrAttTemp(iCnt),"@")
	        Response.Write sArrAttSub1(0)
	        if UBound(sArrAttSub1)=0 then
	            sAttList = sAttList &","& sArrAttSub1(0)
	        else
	            sArrAttribute = split(sArrAttSub1(0),"$")
	            if Trim(sArrAttribute(1))<>"0" then
	                sAttID = sAttID &","& sArrAttribute(0)
	                sAttList = sAttList &","& sArrAttribute(1)
	            end if
	        end if 
	    Next
	    
	  	if trim(sAttList)<>"" then
		 '   sAttID = Mid(sAttID,2)
	        sAttList = Mid(sAttList,2)
	    else
		   ' sAttID = ""
		    sAttList = ""
	    end if ' if trim(arrTemp(6))<>"" then

	Response.Write "sAttID = "& sAttID & " sAttList = "& sAttList
	Response.Clear 
	
	Set Root = OutData.createElement("ROOT")
	OutData.appendChild Root
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0) FROM VWItemStockStatus WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'IN' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			sQuery = " SELECT isNull(SUM(AVAILABLENETSTOCK),0) FROM VW_ITEMLOCATIONLOT_STOCK WHERE ITEMCODE = " & iItemCode & " AND ClassificationCode = "& iClass
			
				if trim(sAttList)<>"0" and trim(sAttList)<>"" then
				  	sQuery  = sQuery & " and AttributeList = '"& sAttList &"'"
			'	elseif trim(sAttID)<>"" and trim(sAttID)<>"0" then
			'		sQuery  = sQuery & " and AttributeList in ("& sAttID &")"
				end if
				'Response.Write sQuery
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.Source 
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("ITEMDETAILS")
				newElem.setAttribute "ITEMCODE", iItemCode
				newElem.setAttribute "CLASSCODE", iClass
				newElem.setAttribute "UNIT", sOrgID
				newElem.setAttribute "STOCK", dcrs(0)
				Root.appendChild newElem
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	 

	Response.ContentType="text/xml"
	Response.Write OutData.xml

%>
