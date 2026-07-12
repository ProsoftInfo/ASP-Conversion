<%Option Explicit%>
<%
	'Program Name				:	GetBoMItemDet.asp
	'Module Name				:	
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:
%>
<!-- #include file="../include/DatabaseConnection.asp"-->
<!--#include file="../include/populate.asp"-->
<%
Dim rsObj,oDOM,dcrs,dcrs1
Dim ndBOMRoot,ndBOMItemDet
Dim iItemCode,iSNo,iBomApplicability
Dim sSql
set rsObj = Server.CreateObject("ADODB.Recordset")
set dcrs = Server.CreateObject("ADODB.Recordset")
set dcrs1 = Server.CreateObject("ADODB.Recordset")
set oDOM = Server.CreateObject("Microsoft.XMLDOM")
iItemCode = Request.QueryString("ItemCode")
	sSql ="Select isNull(BOMApplicability,0) from INV_M_ItemMaster where ItemCode = "& iItemCode
	dcrs.open sSql,con
	if not dcrs.eof then
		iBomApplicability = dcrs(0)
	end if
	dcrs.close
	
	set ndBOMRoot = oDOM.createElement("BOM")
		ndBOMRoot.setAttribute "Applicable",iBomApplicability
		oDOM.appendChild ndBOMRoot 
								
	if Trim(iBomApplicability)="1" then
							
		iSNo = 0
		sSql = "Select BOMItemCode,BOMClassificationCode,ItemDescription,Quantity,UOM,Type,Consumable from INV_M_ItemMasterBOM A,VWItem B where A.BOMItemCode=B.ItemCode and A.ItemCode = "& iItemCode
		dcrs1.Open sSql,con
		if not dcrs1.EOF then
			do while not dcrs1.EOF 
				iSNo = iSNo + 1
				
					set ndBOMItemDet = oDOM.createElement("BOMItem")
						ndBOMItemDet.setAttribute "SNo",iSNo
						ndBOMItemDet.setAttribute "ItemCode",dcrs1(0)
						ndBOMItemDet.setAttribute "ClassCode",dcrs1(1)
						ndBOMItemDet.setAttribute "ItemName",dcrs1(2)
						ndBOMItemDet.setAttribute "Quantity",dcrs1(3)
						ndBOMItemDet.setAttribute "UOM",dcrs1(4)
						ndBOMItemDet.setAttribute "Type",dcrs1(5)
						ndBOMItemDet.setAttribute "Consumable",dcrs1(6)
						ndBOMRoot.appendChild ndBOMItemDet 
						
				dcrs1.MoveNext 
			loop
		end if 'if not dcrs1.EOF then
		dcrs1.Close 
	end if 'if Trim(iBomApplicability)="1" then
	
	Response.ContentType = "text/xml"
	Response.Write oDOM.xml
%>

