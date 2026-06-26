<%
	'Program Name				:	IncludeDatePicker.asp
	'Module Name				:	Common
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 07, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	Name
	'							:
	'Connects To				:	Calling Program
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

<script type="text/javascript" src="../../scripts/itms-modern-compat.js"></script>

<%
	Private Function InsertDatePicker(sName)
%>
	<input type="date" id="<%=sName%>" name="<%=sName%>" class="FormElem itms-date-picker" data-itms-datepicker="1">
	<script type="text/javascript">
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.decorateDateInput(document.getElementById("<%=sName%>"));
		}
	</script>
<%
	End Function
%>

<%
	Private Function ValidateDatePicker(sName)
%>
	<input type="date" id="<%=sName%>" name="<%=sName%>" class="FormElem itms-date-picker" data-itms-datepicker="1">
	<script type="text/javascript">
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.decorateDateInput(document.getElementById("<%=sName%>"));
		}
	</script>
<%
	End Function
%>
