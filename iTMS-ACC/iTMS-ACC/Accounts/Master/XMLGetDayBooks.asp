<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLGetDayBooks.asp
	'Module Name				:	Accounts(master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 19,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<%
Dim oDOM,objRs,objRs1,sUnitID,sUnitLName,sUnitSName
Dim sbookid,sBookName,sQuery,Root,newElem,newElem1
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")
set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set Root = oDOM.createElement("Root")
	    oDOM.appendChild Root

with objRs	
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"
	.ActiveConnection = con
	.Open
end with
if not objrs.EOF then
	 sQuery = "select BookCode,BookName from Acc_M_DayBooks"
	 objrs1.Open sQuery,con
	 if not objRs1.EOF then
	    dim bFlag	
	    Do While Not objRs.EOF
	        sUnitID = objRs(0)
	        sUnitLName = objRs(1)
	        sUnitSName = objRs(3)
		    Set newElem = oDOM.createElement("Unit")
			    newElem.setAttribute "ID", sUnitID
			    newElem.setAttribute "SName", sUnitSName
			    newElem.setAttribute "LName", sUnitLName
		    Root.appendChild newElem
		    objRs1.MoveFirst
		    do while not objRs1.EOF
		        sBookID = objRs1(0)
	            sBookName = objRs1(1)
			    Set newElem1 = oDOM.createElement("Book")
				    newElem1.setAttribute "ID" ,sBookID
				    newElem1.setAttribute "Count", 0
				    newElem1.Text= sBookName
			    newElem.appendChild newElem1
			    objRs1.MoveNext
		    loop
		    objRs.MoveNext
    	Loop
    end if
end if
objRs1.Close
objRs.Close

Response.ContentType = "text/xml"
Response.Write oDOM.xml


%>