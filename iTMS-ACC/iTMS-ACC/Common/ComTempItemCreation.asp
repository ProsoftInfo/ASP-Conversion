
<%
	'Program Name				:	ComTempItemCreation.asp
	'Module Name				:	Inventory (Temporary Item Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	October 06,2011
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

<%
	Function TemporaryItemInsert(sItemCode)
		dim oDOM,RootNode,SessionNode,DetailsNode,oNode,sExp,sExp1,ItemNode
		dim dcrs,dcrs1,sSql,sSessionID
		dim sitmShDesc,sitmDesc,sitmAddDesc,sitmType
		dim iAppCode,iModCode,sCreationStage,iCreatedBy,iTempItemCode

		' Create our DOM Document Objects
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

		if oDOM.Load(server.MapPath("../../inventory/xmldata/TEMPORARYITEM.xml")) then

			Set dcrs = Server.CreateObject("ADODB.RecordSet")
			Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

			iCreatedBy = getUserid

			Set RootNode = oDOM.documentElement

			sExp1 ="//ITEMCODE [ @ITMCODE = '"&sItemCode&"']"
			Set ItemNode = RootNode.Selectnodes(sExp1)

			sExp ="//ITEMCODE [ @ITMCODE = '"&sItemCode&"']/SESSIONDETAILS [ @SESSIONID = '"&Session.SessionID&"']"
			Set SessionNode = RootNode.Selectnodes(sExp)

			if SessionNode.Length > 0 then
				set DetailsNode = SessionNode.Item(0).childNodes(0)
				sitmType = trim(DetailsNode.Attributes.getNamedItem("ITMTYPE").Value)
				iAppCode = trim(DetailsNode.Attributes.getNamedItem("APPCODE").Value)
				iModCode = trim(DetailsNode.Attributes.getNamedItem("MODCODE").Value)
				sCreationStage = trim(DetailsNode.Attributes.getNamedItem("CRESTAGE").Value)
				sitmDesc = trim(DetailsNode.Attributes.getNamedItem("ITMDESC").Value)
				sitmShDesc = trim(DetailsNode.Attributes.getNamedItem("ITMSHDESC").Value)
				sitmAddDesc = trim(DetailsNode.Attributes.getNamedItem("ITMADDDESC").Value)

				if sitmShDesc = "" then
					sitmShDesc = "NULL"
				else
					sitmShDesc = Pack(sitmShDesc)
				end if

				if sitmAddDesc = "" then
					sitmAddDesc = "NULL"
				else
					sitmAddDesc = Pack(sitmAddDesc)
				end if

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT TEMPITEMCODE,ITEMDESCRIPTION,ITEMTYPEID,SHORTDESCRIPTION,ADDITIONALDESCRIPTION FROM MS_TEMPORARYITEMMASTER WHERE UPPER(GENITEMCODE) = " & Pack(ucase(sItemCode)) & " ORDER BY 1"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				if dcrs1.EOF then
					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(MAX(TEMPITEMCODE)+1,1) FROM MS_TEMPORARYITEMMASTER"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing
					if not dcrs.EOF then
						iTempItemCode = trim(dcrs(0))
						sSql = "INSERT INTO MS_TEMPORARYITEMMASTER (TEMPITEMCODE,GENITEMCODE,SHORTDESCRIPTION," &_
							"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,ITEMTYPEID,CREATEDBY,CREATEDON, " &_
							"APPLICATIONCODE,MODULECODE,CREATIONSTAGE,FINALSTATUS) VALUES " &_
							"(" & iTempItemCode & "," & Pack(ucase(sItemCode)) & "," & sitmShDesc & "," &_
							"" & Pack(sitmDesc) & "," & sitmAddDesc & "," & Pack(sitmType) & "," &_
							"" & iCreatedBy & ",CONVERT(DATETIME,GETDATE(),103)," & iAppCode & "," &_
							"" & iModCode & "," & Pack(sCreationStage) & ",'N')"
						'Response.Write sSql
						con.Execute sSql

						TemporaryItemInsert = iTempItemCode&"**"&sitmDesc

						Set oNode = ItemNode.Item(0).RemoveChild(SessionNode.Item(0))
						oDOM.Save server.MapPath("../../inventory/xmldata/TEMPORARYITEM.xml")
						exit function
					end if
					dcrs.close
				else
					iTempItemCode = trim(dcrs1(0))
					sitmDesc = trim(dcrs1(1))
					sitmType = trim(dcrs1(2))
					sitmShDesc = trim(dcrs1(3))
					sitmAddDesc = trim(dcrs1(4))

					if sitmShDesc = "" or isNull(sitmShDesc) then
						sitmShDesc = "NULL"
					else
						sitmShDesc = Pack(sitmShDesc)
					end if

					if sitmAddDesc = "" or isNull(sitmAddDesc) then
						sitmAddDesc = "NULL"
					else
						sitmAddDesc = Pack(sitmAddDesc)
					end if

					sSql = "INSERT INTO MS_TEMPORARYITEMMASTER (TEMPITEMCODE,GENITEMCODE,SHORTDESCRIPTION," &_
						"ITEMDESCRIPTION,ADDITIONALDESCRIPTION,ITEMTYPEID,CREATEDBY,CREATEDON, " &_
						"APPLICATIONCODE,MODULECODE,CREATIONSTAGE,FINALSTATUS) VALUES " &_
						"(" & iTempItemCode & "," & Pack(ucase(sItemCode)) & "," & sitmShDesc & "," &_
						"" & Pack(sitmDesc) & "," & sitmAddDesc & "," & Pack(sitmType) & "," &_
						"" & iCreatedBy & ",CONVERT(DATETIME,GETDATE(),103)," & iAppCode & "," &_
						"" & iModCode & "," & Pack(sCreationStage) & ",'N')"
					'Response.Write sSql
					con.Execute sSql

					TemporaryItemInsert = iTempItemCode&"**"&sitmDesc

					Set oNode = ItemNode.Item(0).RemoveChild(SessionNode.Item(0))
					oDOM.Save server.MapPath("../../inventory/xmldata/TEMPORARYITEM.xml")
					exit function
				end if
				dcrs1.close
			else
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT TEMPITEMCODE,ITEMDESCRIPTION,ITEMTYPEID,SHORTDESCRIPTION,ADDITIONALDESCRIPTION FROM MS_TEMPORARYITEMMASTER WHERE UPPER(GENITEMCODE) = " & Pack(ucase(sItemCode)) & " ORDER BY 1"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				if not dcrs1.EOF then
					iTempItemCode = trim(dcrs1(0))
					sitmDesc = trim(dcrs1(1))
					sitmType = trim(dcrs1(2))
					sitmShDesc = trim(dcrs1(3))
					sitmAddDesc = trim(dcrs1(4))

					TemporaryItemInsert = iTempItemCode&"**"&sitmDesc

					exit function
				end if
				dcrs1.close

			end if
		else
			TemporaryItemInsert = "Err**"
			exit Function
		end if
	end Function
%>