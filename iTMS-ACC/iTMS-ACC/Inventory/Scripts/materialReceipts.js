(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function selectedValue(name) {
		var frm = form();
		var control = frm && frm.elements ? frm.elements[name] : null;
		if (!control) {
			return "";
		}
		if (control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].value;
		}
		return control.value || "";
	}

	window.Submit = function (obj) {
		form().action = "MaterialReceipts.asp?OptType=" + encodeURIComponent(obj && obj.value || "");
		form().submit();
		return false;
	};

	window.CreateReceipt = function (orgId) {
		window.location.href = "receiptinternalEntry.asp?OrgID=" + encodeURIComponent(orgId || "");
		return false;
	};

	window.SubmitMe = function () {
		var frm = form();
		frm.action = "MaterialReceipts.asp?OptType=" + encodeURIComponent(frm.hOptType ? frm.hOptType.value : "") + "&RCPTTYPE=" + encodeURIComponent(selectedValue("cmbRcptType"));
		frm.submit();
		return false;
	};

	window.CreateInvoice = function () {
		form().action = "InventoryReceipt.asp";
		form().submit();
		return false;
	};
}(window, document));
