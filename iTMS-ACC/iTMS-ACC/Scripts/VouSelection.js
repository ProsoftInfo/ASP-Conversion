(function (window, document) {
	var compatPath = "/Scripts/itms-modern-compat.js";

	function ensureCompat() {
		if (window.ITMSModernCompat) {
			return;
		}
		var script = document.createElement("script");
		script.src = compatPath;
		document.head.appendChild(script);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || null;
	}

	function selectedValue(control) {
		if (!control) {
			return "";
		}
		if (typeof control.selectedIndex === "number" && control.options) {
			var option = control.options[control.selectedIndex];
			return option ? option.value : "";
		}
		return control.value || "";
	}

	function selectedIndex(name) {
		var control = field(name);
		return control && typeof control.selectedIndex === "number" ? control.selectedIndex : -1;
	}

	function validateSelection(trans) {
		if (String(trans) === "GJR") {
			if (selectedIndex("selUnitId") === 0) {
				alert("Select Unit ");
				field("selUnitId").focus();
				return false;
			}
			if (selectedIndex("selBook") === 0) {
				alert("Select Book ");
				field("selBook").focus();
				return false;
			}
			return true;
		}
		if (selectedIndex("selUnitId") === 0) {
			alert("Select Unit ");
			field("selUnitId").focus();
			return false;
		}
		if (selectedIndex("selBook") === 0) {
			alert("Select Book ");
			field("selBook").focus();
			return false;
		}
		if (selectedIndex("selVouType") === 0) {
			alert("Select Voucher Type ");
			field("selVouType").focus();
			return false;
		}
		return true;
	}

	function transType(baseTrans, voucherType) {
		if (voucherType === "C") {
			return baseTrans + "P";
		}
		if (voucherType === "0") {
			return baseTrans;
		}
		return baseTrans + "R";
	}

	function normalizeBookNo(bookNo) {
		var parts = String(bookNo || "").split("-");
		return parts.length === 2 ? parts[0] : bookNo;
	}

	function openVoucherSelection(callType, bookId, trans, orgId, voucherType, bookNo, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		var url = "VoucherSelection.asp?flag=" + encodeURIComponent(callType) +
			"&orgId=" + encodeURIComponent(orgId || "") +
			"&BookCode=" + encodeURIComponent(bookId || "") +
			"&BookNo=" + encodeURIComponent(bookNo || "") +
			"&TransType=" + encodeURIComponent(trans || "");
		window.ITMSModernCompat.openModalDialog(url, "", "dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", callback);
	}

	function applyVoucherSelection(callType, result, displayOnly) {
		var temp = trim(result);
		if (temp === "0" || temp === "") {
			if (field("hTransNo")) {
				field("hTransNo").value = "";
			}
			if (field("txtVouNo")) {
				field("txtVouNo").value = "";
			}
			return;
		}

		var parts = temp.split("~");
		var voucherText = parts[1] || "";
		if (field("hTransNo")) {
			field("hTransNo").value = parts[0] || "";
		}
		if (field("txtVouNo")) {
			field("txtVouNo").value = callType === "A" ? voucherText.split(">>")[0] : voucherText;
		}
		if (!displayOnly && field("hAmendTy")) {
			field("hAmendTy").value = callType === "A" ? "A" : "N";
		}
		if (!displayOnly && field("selUnitId")) {
			field("selUnitId").disabled = callType !== "A";
		}
		if (!displayOnly && typeof window.MakeDispVou === "function") {
			window.MakeDispVou(callType);
		}
	}

	window.popVoucherNo = function (callType, bookId, trans, orgId, voucherType, bookNo) {
		if (trim(voucherType) === "" && field("selVouType")) {
			voucherType = selectedValue(field("selVouType"));
		}
		if (!validateSelection(trans)) {
			return;
		}
		openVoucherSelection(
			callType,
			bookId,
			transType(trans, voucherType),
			orgId,
			voucherType,
			normalizeBookNo(bookNo),
			function (value) {
				applyVoucherSelection(callType, value, false);
			}
		);
	};

	window.popVoucherNoDispOnly = function (callType, bookId, trans, orgId, voucherType, bookNo) {
		if (!validateSelection(trans)) {
			return;
		}
		openVoucherSelection(
			callType,
			bookId,
			transType(trans, voucherType),
			orgId,
			voucherType,
			bookNo,
			function (value) {
				applyVoucherSelection(callType, value, true);
			}
		);
	};

	ensureCompat();
})(window, document);
