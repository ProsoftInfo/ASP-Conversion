<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetItemStoreInfo.asp
	'Module Name				:	Include
	'Author Name				:	Ragavendran
	
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/populate.asp" -->
<%
	dim newxml,ndRoot,ndStorage
	dim rsTemp,sItemCode,sQuery,sOrgID,sStoreName,iCountStore,iStoreNo,iBinNo

	Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")
	Set ndRoot = newxml.createElement("Root")
	newxml.appendChild ndRoot
	sItemCode = Request.QueryString("ItemCode")
	sOrgID = session("organizationcode")
    
    sStoreName=""
	
	sQuery = "SELECT count(IM.LOCATIONNUMBER) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & sItemCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1"
	'Response.write sQuery
	rsTemp.Open sQuery,con
	if not rsTemp.EOF then
		iCountStore =  rsTemp(0)
	end if
	rsTemp.Close 
	ndRoot.setAttribute "StoreCount",iCountStore
	if not iCountStore > 1 then
	    sQuery = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0) FROM INV_M_STORAGE IC,INV_M_ITEMSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & sItemCode & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1 ORDER BY 1,2"
	    rsTemp.Open sQuery,con
	    if not rsTemp.EOF then
		    Do While Not rsTemp.EOF
    			
			    iStoreNo = rsTemp(0)
			    iBinNo = rsTemp(1)
			    sStoreName = DisplayStore(trim(rsTemp(0)),trim(rsTemp(1)))
			    set ndStorage = newxml.createElement("Storage")
			        ndStorage.setAttribute "LocNo",iStoreNo
			        ndStorage.setAttribute "BinNo",iBinNo
			        ndStorage.setAttribute "StoreName",sStoreName
			    ndRoot.appendChild ndStorage
		        rsTemp.MoveNext
		    loop
	    end if
	    rsTemp.Close
	end if 'if not iCountStore > 1 then

    Response.ContentType = "text/xml"
    Response.Write newxml.xml	
    '***********************************************
    Function DisplayStore(sLoc,sBin)
		' Declaration of variables
		Dim dcrs,sBinName,sLocName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT LOCATIONNAME,LOCATIONCODE FROM INV_M_STORAGE WHERE LOCATIONNUMBER = " & sLoc & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sLocName = trim(dcrs(0))
		else
			sLocName = "-"
		end if
		dcrs.close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) ORDER BY BINNUMBER"
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			DisplayStore = trim(sLocName)&" -- "&trim(dcrs(0))
		else
			DisplayStore = trim(sLocName)
		end if
		dcrs.Close

	End Function
%>