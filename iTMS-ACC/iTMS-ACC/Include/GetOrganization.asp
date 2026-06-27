
<%
	' Function to Get the Organization Name
	Function GetOrganization()
		' Declaration of variables
		Dim oDom,fs,Root,PGNode,dcrs
		dim sOrgName,sUnitName

		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGANIZATIONNAME FROM DCS_ORGANIZATION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		If not dcrs.EOF then
			sOrgName = trim(dcrs(0))
		Else
			sOrgName = "None"
		end if
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = '" & Session("organizationcode") & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		If not dcrs.EOF then
			sUnitName = trim(dcrs(0))
		end if
		dcrs.Close

		'GetOrganization = sOrgName & " (" & sUnitName & ")"
		GetOrganization = sOrgName
	End Function
%>

<%
	Private Function InsertMenu(AppCode)
%>
	<!--OBJECT id="<%=AppCode%>" classid=clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11 width=84 height=20 type=application/x-oleobject VIEWASTEXT>
		<PARAM NAME="Width" VALUE="1455">
		<PARAM NAME="Height" VALUE="556">
		<PARAM name="Font" value=", 8,,#000000, BOLD">
		<PARAM NAME="Command" VALUE="Related Topics, menu">
		<PARAM NAME="text" VALUE="text:Modules"-->
		<select size="1" name="cmbModules" class="FormElem" OnChange="GoToModuleHome()">
<%
	dim dcrs,iCounter
	iCounter = 0
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		if lcase(trim(Session("loginid"))) = "admin" then
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS ORDER BY APPLICATIONCODE"
		else
			.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS WHERE APPLICATIONCODE IN (SELECT DISTINCT APPLICATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & trim(Session("userid")) & ") ORDER BY APPLICATIONCODE"
		end if
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		Do While Not dcrs.EOF
			iCounter = cdbl(iCounter) + 1
			If trim(dcrs(0)) = AppCode Then
%>
		    <!--PARAM name="Item<%=iCounter%>" value="<%=trim(dcrs(1))%>;<%=trim(dcrs(2))%>"-->
		    <option value="<%=trim(dcrs(0))%>~<%=trim(dcrs(2))%>" Selected><%=trim(dcrs(1))%></option>
<%
			Else
%>
			<option value="<%=trim(dcrs(0))%>~<%=trim(dcrs(2))%>"><%=trim(dcrs(1))%></option>

<%
			End If
		dcrs.MoveNext
		Loop
	end if
	dcrs.Close
%>
	<!--/OBJECT-->
	</Select>
<%
	End Function
%>

<%
	Private Function InsertSplitButtonMenu(sName)
%>
	<table>
			<tr>
				<td>
					<fieldset id="splitbuttonsfrommarkup">
						<input type="submit" id="splitbutton1" name="splitbutton1_button" value="Menu">
						<select id="splitbutton1select" name="splitbutton1select" multiple>
						<%
							dim dcrs,iCounter
								iCounter = 0
								'Declaration of Objects
								Set dcrs = Server.CreateObject("ADODB.RecordSet")
								with dcrs
									.CursorLocation = 3
									.CursorType = 3
									if lcase(trim(Session("loginid"))) = "admin" then
										.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS ORDER BY APPLICATIONCODE"
									else
										.Source = "SELECT DISTINCT APPLICATIONCODE,APPLICATIONNAME,APPLICATIONPATH FROM MS_APPLICATIONS WHERE APPLICATIONCODE IN (SELECT DISTINCT APPLICATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & trim(Session("userid")) & ") ORDER BY APPLICATIONCODE"
									end if
									.ActiveConnection = con
									.Open
								end with
								set dcrs.ActiveConnection = nothing

								if not dcrs.EOF then
									Do While Not dcrs.EOF
										iCounter = cdbl(iCounter) + 1
						%>
											<option value="<%=trim(dcrs(1))%>;<%=trim(dcrs(2))%>"><%=trim(dcrs(1))%></option>
						<%
										dcrs.MoveNext
									Loop
								end if
								dcrs.Close
						%>
						</select>
	        		</fieldset>
				</td>
			</tr>
	</table>
<%
	End Function
%>
