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
		return frm && frm.elements[name];
	}

	function tempRoot() {
		return window.TempData && window.TempData.documentElement;
	}

	function setReturnValue() {
		var root = tempRoot();
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.Done_Clk = function () {
		var rateUomField = field("mCmbRateUOM");
		var rateUom = rateUomField && rateUomField.value || "";
		var classCode = field("hClassCode") && field("hClassCode").value || "";
		var quantityUom = field("hUOM") && field("hUOM").value || "";
		var rateField = field("mTxtNewRate");
		var rateValue = trim(rateField && rateField.value);
		var root;

		if (rateUom === "0") {
			alert("Select UOM");
			if (rateUomField && rateUomField.focus) {
				rateUomField.focus();
			}
			return;
		}

		if (trim(rateUom)) {
			rateUom = rateUom.split(":")[0];
		}

		if ((classCode === "0" || classCode === "TEMP") && trim(rateUom) !== trim(quantityUom)) {
			alert("Quantity UoM and Rate UoM should be same for Temporary Items");
			if (rateUomField && rateUomField.focus) {
				rateUomField.focus();
			}
			return;
		}

		if (!rateValue) {
			alert("Enter Rate");
			if (rateField && rateField.focus) {
				rateField.focus();
			}
			return;
		}

		if (isNaN(Number(rateValue))) {
			alert("Enter Number");
			if (rateField && rateField.focus) {
				rateField.focus();
			}
			return;
		}

		if (Number(rateValue) <= 0) {
			alert("Rate should be > 0 ");
			if (rateField && rateField.focus) {
				rateField.focus();
			}
			return;
		}

		root = tempRoot();
		root.setAttribute("RateUOM", rateUom);
		root.setAttribute("RatePerQtyUoM", rateValue);
		setReturnValue();
		window.close();
	};

	window.addEventListener("beforeunload", setReturnValue);
})(window, document);
