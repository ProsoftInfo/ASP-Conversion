(function (window, document) {
	"use strict";

	function loadCompat() {
		var currentScript;
		var src;
		var loader;
		if (window.ITMSModernCompat || document.querySelector('script[src*="itms-modern-compat.js"]')) {
			return;
		}
		currentScript = document.currentScript;
		src = currentScript ? currentScript.getAttribute("src") || "" : "";
		loader = document.createElement("script");
		loader.type = "text/javascript";
		loader.src = src ? src.replace(/TaxEntryCompat\.js(?:\?.*)?$/i, "itms-modern-compat.js") : "../../scripts/itms-modern-compat.js";
		(document.head || document.documentElement).appendChild(loader);
	}

	loadCompat();

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || null;
	}

	function xmlObject(name) {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlDocument(name) {
		var obj = xmlObject(name);
		return obj && (obj.XMLDocument || obj._doc || obj) || null;
	}

	function xmlRoot(name) {
		var doc = xmlDocument(name);
		return doc && doc.documentElement || null;
	}

	function attr(node, name) {
		var value = node && node.attributes && node.attributes.getNamedItem(name);
		return value ? value.value : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return context.selectNodes(expression);
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (var i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		return nodes;
	}

	function setApproval(invoiceNodeName) {
		var root = xmlRoot("TaxData");
		var approverRequired = field("hApprover") && field("hApprover").value === "Y";
		var approver = field("selUserId");
		var invoiceNodes;
		if (approverRequired && (!approver || approver.selectedIndex <= 0)) {
			alert("Select Approver");
			if (approver) {
				approver.focus();
			}
			return false;
		}
		invoiceNodes = selectNodes(root, "//" + invoiceNodeName);
		for (var i = 0; i < invoiceNodes.length; i += 1) {
			setAttr(invoiceNodes[i], "Approval", field("hApprover") ? field("hApprover").value : "");
			setAttr(invoiceNodes[i], "Approver", approverRequired && approver ? approver.value : "0");
		}
		return true;
	}

	function updateTaxAccHeads(options) {
		var root = xmlRoot("TaxData");
		var taxNodes = selectNodes(root, "//Tax");
		for (var i = 0; i < taxNodes.length; i += 1) {
			var taxCode = attr(taxNodes[i], "TaxCode");
			var catCode = attr(taxNodes[i], "CatCode");
			var accountSelect;
			if (options.skipZeroTax && (taxCode === "0" || catCode === "0")) {
				continue;
			}
			accountSelect = field("SelAccHead" + catCode + taxCode);
			if (!accountSelect) {
				continue;
			}
			if (options.enableTaxSelects) {
				accountSelect.disabled = false;
			}
			setAttr(taxNodes[i], "AccHead", accountSelect.value);
		}
	}

	function postTaxXml(saveUrl) {
		var request = new XMLHttpRequest();
		request.open("POST", saveUrl, false);
		request.send(xmlDocument("TaxData"));
		return request.responseText || "";
	}

	function checkSubmit(options) {
		var frm = form();
		var responseText;
		options = options || {};
		if (!setApproval(options.invoiceNodeName || "PurInvoice")) {
			return;
		}
		updateTaxAccHeads(options);
		responseText = postTaxXml(options.saveUrl || "XMLSave.asp");
		if (trim(responseText) !== "") {
			alert(responseText);
			return;
		}
		if (typeof options.beforeSubmit === "function") {
			options.beforeSubmit(frm);
		}
		if (options.disableButton && frm && frm.elements[options.disableButton]) {
			frm.elements[options.disableButton].disabled = true;
		}
		if (frm) {
			frm.submit();
		}
	}

	window.ITMSTaxEntryCompat = {
		checkSubmit: checkSubmit,
		trim: trim
	};
}(window, document));
