<%
Dim sSql,rs,rs1
Dim oDOM2,newElem,newElem1,RtElem,newElem2
dim StoNode,DelNode

Set oDOM2 = Server.CreateObject("Microsoft.XMLDOM")

Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs1 = Server.CreateObject("ADODB.RecordSet")
Set Root = oDOM2.documentElement

Set RtElem = oDOM2.CreateElement("Root")
oDOM2.appendchild RtElem

sSql = "Select OUDefinitionID,LocationNumber,LocationCode,LocationName,ApplicableFor,StorageTypeFree,"&_
	   "StorageTypeBins,isNull(UsableFreeArea,0),isNull(NumberOfBins,0) from Inv_M_Storage "
	rs.Open sSql,con
	if not rs.eof then
		Set newElem1 = oDOM2.createElement("Organization")
			newElem1.setAttribute "OUDEFINITIONID",rs(0)
		RtElem.appendChild newElem1
	end if
	do while not rs.EOF

	   Set newElem = oDOM2.CreateElement("Storage")
		newElem.setAttribute "LOCATIONNUMBER", rs(1)
		newElem.setAttribute "LOCATIONCODE", ucase(rs(2))
		newElem.setAttribute "LOCATIONNAME", ucase(rs(3))
		newElem.setAttribute "APPLICABLEFOR", rs(4)
		newElem.setAttribute "STORAGETYPEFREE", rs(5)
		newElem.setAttribute "STORAGETYPEBINS", rs(6)
		newElem.setAttribute "USABLEFREEAREA", rs(7)
		newElem.setAttribute "NUMBEROFBINS",rs(8)
		newElem1.appendChild newElem


		sSql = "Select BinNumber,BinCode,BinName,BinArea from Inv_M_StoreBinDetails where OUDefinitionID = '"& rs(0) &"' and "&_
				"LocationNumber = "& rs(1) &" "
		rs1.Open sSql,con
		do while not rs1.EOF
			Set newElem2 = oDOM2.createElement("Bin")
			newElem2.setAttribute "BINNUMBER", rs1(0)
			newElem2.setAttribute "BINCODE", ucase(rs1(1))
			newElem2.setAttribute "BINNAME", ucase(rs1(2))
			newElem2.setAttribute "BINAREA", rs1(3)
			newElem.appendChild newElem2

			rs1.MoveNext
		loop
		rs1.Close
		rs.MoveNext
	loop
	rs.Close
	oDOM2.save  server.MapPath("../xmlData/Storage.xml")
%>