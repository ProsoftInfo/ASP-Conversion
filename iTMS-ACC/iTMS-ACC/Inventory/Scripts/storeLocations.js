(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function checkbox(name) {
		var item = field(name);
		return item && item.type === "checkbox" ? item : null;
	}

	function Help() {
		window.open("../HelpFiles/StorageLocation.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
		return false;
	}

	function selectedLocationValues() {
		var count = Number(valueOf("hCnt")) || 0;
		var values = [];
		var box;
		var parts;
		for (var i = 1; i < count; i += 1) {
			box = checkbox("chkbox" + i) || checkbox("Chkbox" + i);
			if (box && box.checked) {
				parts = String(box.value || "").split(":");
				values.push(parts[1] || "");
			}
		}
		return values;
	}

	function FinalSubmit(action) {
		var values;
		if (trim(action) === "CR") {
			form().action = "ORGSTORAGEDEFINITIONENTRY.ASP";
		} else if (trim(action) === "AM") {
			values = selectedLocationValues();
			if (!values.length) {
				alert("Select Location Name to Amend");
				return false;
			}
			if (values.length > 1) {
				alert("Select any One Location Name to Amend");
				return false;
			}
			setValue("hPara", values[0]);
			form().action = "ORGSTORAGEDEFINITIONENTRY.ASP";
		}
		form().submit();
		return false;
	}

	function Validate() {
		setValue("hLocName", valueOf("txtLocName"));
		setValue("hAppfor", valueOf("cmbAppFor"));
		form().submit();
		return false;
	}

	function ChngAppFor() {
		setValue("hAppfor", valueOf("cmbAppFor"));
		return false;
	}

	function FnInit() {
		setValue("txtLocName", valueOf("hLocName"));
		setValue("cmbAppFor", valueOf("hAppfor"));
	}

	function DeleteItems() {
		var values = selectedLocationValues();
		if (!values.length) {
			alert("Select Location Name to Delete");
			return false;
		}
		if (confirm("Do U want to Delete the selected Storage Locations?")) {
			setValue("hPara", values.join("^"));
			form().action = "OrgStorageDeletionInsert.asp";
			form().submit();
		}
		return false;
	}

	function AssignPage(pageNo) {
		setValue("hPage", pageNo);
		form().submit();
		return false;
	}

	window.Help = Help;
	window.FinalSubmit = FinalSubmit;
	window.Validate = Validate;
	window.ChngAppFor = ChngAppFor;
	window.FnInit = FnInit;
	window.DeleteItems = DeleteItems;
	window.AssignPage = AssignPage;
}(window, document));
