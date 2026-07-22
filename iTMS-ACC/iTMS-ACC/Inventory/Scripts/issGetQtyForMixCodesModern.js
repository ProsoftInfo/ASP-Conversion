(function (window, document) {
	"use strict";

	var mixDoc = null;
	var mixRoot = null;

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function xmlIslandDocument(id, rootName) {
		var island;
		ensureCompat();
		island = document.getElementById(id);
		if (island && island.XMLDocument) {
			return island.XMLDocument;
		}
		if (island && island._doc) {
			return island._doc;
		}
		if (island && island.documentElement) {
			return island;
		}
		return document.implementation.createDocument("", rootName, null);
	}

	function loadMixXml() {
		mixDoc = mixDoc || xmlIslandDocument("MixData", "MixData");
		mixRoot = mixRoot || mixDoc.documentElement || mixDoc.appendChild(mixDoc.createElement("MixData"));
		return mixRoot;
	}

	function clearMixRows(root) {
		var rows = root ? root.getElementsByTagName("Mix") : [];
		for (var i = rows.length - 1; i >= 0; i -= 1) {
			rows[i].parentNode.removeChild(rows[i]);
		}
	}

	function mixCodes() {
		var source = field("hRefNo");
		var parts = String(source ? source.value : "").split(",");
		var codes = [];
		for (var i = 0; i < parts.length; i += 1) {
			if (trim(parts[i]) !== "") {
				codes.push(trim(parts[i]));
			}
		}
		return codes;
	}

	function setReturnValue(root) {
		if (!root) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.CheckSubmit = function () {
		var codes = mixCodes();
		var totalMixQty = 0;
		var totalQty = toNumber(field("hTotQty") && field("hTotQty").value);
		var root = loadMixXml();
		var qtyField;
		var nameField;
		var node;
		var i;

		for (i = 0; i < codes.length; i += 1) {
			qtyField = field("txtQtyZ" + codes[i]);
			nameField = field("hMixNameZ" + codes[i]);
			if (!qtyField || trim(qtyField.value) === "" || toNumber(qtyField.value) === 0) {
				alert("Enter the Quantity for " + trim(nameField && nameField.value));
				if (qtyField) {
					qtyField.focus();
				}
				return false;
			}
		}

		for (i = 0; i < codes.length; i += 1) {
			qtyField = field("txtQtyZ" + codes[i]);
			totalMixQty += toNumber(qtyField && qtyField.value);
		}

		if (Math.abs(totalQty - totalMixQty) > 0.000001) {
			alert("Mix Quantity Should be equal to Total Quantity");
			return false;
		}

		clearMixRows(root);
		for (i = 0; i < codes.length; i += 1) {
			qtyField = field("txtQtyZ" + codes[i]);
			nameField = field("hMixNameZ" + codes[i]);
			node = mixDoc.createElement("Mix");
			node.setAttribute("Code", codes[i]);
			node.setAttribute("Name", trim(nameField && nameField.value));
			node.setAttribute("Qty", trim(qtyField && qtyField.value));
			root.appendChild(node);
		}

		root.setAttribute("Action", "Done");
		setReturnValue(root);
		window.close();
		return false;
	};

	window.window_onunload = function () {
		setReturnValue(loadMixXml());
	};
	window.addEventListener("beforeunload", window.window_onunload);
}(window, document));
