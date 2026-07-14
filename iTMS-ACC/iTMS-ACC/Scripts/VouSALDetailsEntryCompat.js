(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function xmlObject(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || byId(name);
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		if (!object) {
			return null;
		}
		return object.XMLDocument || object._doc || (object.nodeType === 9 ? object : null);
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function serializeXml(name) {
		var doc = xmlDocument(name);
		var root = xmlRoot(name);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body || "");
		return xhr;
	}

	function setButton(name, disabled) {
		var button = field(name);
		if (button) {
			button.disabled = !!disabled;
		}
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	window.InitVouSALDetailsEntry = function () {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		window.SalesVoucherEntryMode = "legacy-details";
		window.VouRoot = xmlRoot("DetData");
		window.EntryRoot = xmlRoot("EntryData");
		window.iEntryNo = 1;
		window.bVouFlag = false;
		window.bSavFlag = false;
		setButton("btnAdd", false);
		setButton("btnNext", false);
		setButton("btnUpdate", true);
		setButton("btnDel", true);
	};

	window.validate = function () {
		return true;
	};

	window.VouCreate = function () {
		var root = xmlRoot("DetData");
		if (root && field("hInvDate")) {
			setAttr(root, "VouDate", field("hInvDate").value);
		}
		return true;
	};

	window.SaveXML = function () {
		var xhr = syncPost("XMLUpdate.asp?Mod=SAL&Name=Voucher Entry", serializeXml("DetData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		if (field("txtDescription") && trim(field("txtDescription").value) === "" && Number(window.iEntryNo) < 1) {
			alert("Select Item ");
			field("txtDescription").focus();
			return false;
		}
		setButton("btnNext", true);
		setButton("btnAdd", true);
		form().submit();
		return true;
	};
}(window, document));
