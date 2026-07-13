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
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
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

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function selectedPackCodes() {
		var count = Number(valueOf("hCnt")) || 0;
		var selected = [];
		var box;
		for (var i = 1; i <= count; i += 1) {
			box = field("chkBox" + i);
			if (box && box.checked) {
				selected.push(box.value);
			}
		}
		return selected;
	}

	function postText(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(null);
		return xhr.responseText || "";
	}

	function Help() {
		window.open("../HelpFiles/PackingType.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
		return false;
	}

	function CheckSubmit() {
		form().action = "PackingTypes.asp?PackName=" + encodeURIComponent(valueOf("txtPackName"));
		form().submit();
		return false;
	}

	function fnInit() {
		setValue("txtPackName", valueOf("hPackName"));
	}

	function ShowPackingType(type) {
		var selected = selectedPackCodes();
		var packCode = "";
		if (type === "E") {
			if (selected.length !== 1) {
				alert("Select any one Pack for Edit");
				return false;
			}
			packCode = selected[0];
		}
		openDialog("PackingType.asp?Type=" + encodeURIComponent(type + ":" + packCode), "", "dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (value) {
			if (trim(value) === "Done") {
				form().submit();
			}
		});
		return false;
	}

	function DelItem() {
		var selected = selectedPackCodes();
		var response;
		if (!selected.length) {
			alert("Select any one Pack for Edit");
			return false;
		}
		response = trim(postText("PackingTypeDelete.asp?Code=" + encodeURIComponent(selected.join(","))));
		if (response) {
			alert(response);
		} else {
			alert("Selected Pack Is Deleted Successfully");
			form().submit();
		}
		return false;
	}

	function ShowPackingForEntry() {
		window.open("PackingForEntry.asp", "", "height=500,width=600,status=no,top=0,left=0");
		return false;
	}

	function ShowPackingForEntryStatic() {
		openDialog("PackingForEntryStatic.asp", "", "dialogHeight:330px;dialogWidth:570px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	}

	window.Help = Help;
	window.CheckSubmit = CheckSubmit;
	window.fnInit = fnInit;
	window.ShowPackingType = ShowPackingType;
	window.DelItem = DelItem;
	window.ShowPackingForEntry = ShowPackingForEntry;
	window.ShowPackingForEntryStatic = ShowPackingForEntryStatic;
}(window, document));
